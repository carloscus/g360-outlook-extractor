import logging
import os
import re
import sys
import threading
import time
from collections.abc import Callable
from datetime import datetime, time
from pathlib import Path
from typing import Any

import pythoncom
import win32com.client

# ---------------------------------------------------------------------------
# Rutas base robustas (funciona en dev y en exe congelado con PyInstaller)
# ---------------------------------------------------------------------------

FOLDER_SEPARATOR = " > "
MAX_BASENAME_LEN = 80


def get_app_base_dir() -> Path:
    """Directorio base para logs y datos (exe dir si frozen, si no raíz proyecto)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # .../src/g360_outlook_extractor/engine.py -> raíz = parents[2]
    return Path(__file__).resolve().parents[2]


def get_audit_log_path() -> Path:
    return get_app_base_dir() / "audit.log"


logging.basicConfig(
    filename=str(get_audit_log_path()),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    encoding="utf-8",
)
logger = logging.getLogger("G360.OutlookEngine")


def get_desktop_dir() -> Path:
    """Resuelve el Escritorio tolerando OneDrive, español y fallbacks."""
    candidates = [
        Path.home() / "Desktop",
        Path.home() / "OneDrive" / "Desktop",
        Path.home() / "OneDrive" / "Escritorio",
        Path.home() / "Escritorio",
    ]
    # SHGetFolderPath si está disponible (más fiable en Windows)
    try:
        from win32com.shell import shell, shellcon

        desktop = shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
        if desktop:
            candidates.insert(0, Path(desktop))
    except Exception:
        pass
    for cand in candidates:
        if cand.is_dir():
            return cand
    fallback = Path.home() / "Documents"
    fallback.mkdir(parents=True, exist_ok=True)
    return fallback


class OutlookEngine:
    """
    Motor de extracción MAPI (Outlook) optimizado para el ecosistema G360.
    Cada método COM inicializa su propio contexto STA y no comparte
    objetos COM entre hilos.
    """

    # Extensiones por categoría para filtrado de adjuntos
    CATEGORIES = {
        "Excel": [".xls", ".xlsx", ".xlsm", ".csv"],
        "PDF": [".pdf"],
        "Docs": [".doc", ".docx", ".txt"],
    }

    def __init__(self):
        self._is_cancelled = threading.Event()

    def stop_extraction(self):
        """Señala al hilo de extracción que debe detenerse."""
        self._is_cancelled.set()
        logger.warning("Solicitud de cancelación de extracción recibida.")

    # -- COM helpers (por hilo, sin estado compartido) ---------------------

    def _initialize_com(self) -> tuple[Any, Any]:
        """Inicializa COM en el hilo actual y retorna (outlook, namespace)."""
        pythoncom.CoInitialize()
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            namespace = outlook.GetNamespace("MAPI")
            logger.info("Conexión MAPI establecida exitosamente.")
            return outlook, namespace
        except Exception as e:
            # Emparejar CoInitialize si Dispatch falló
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            logger.error(f"Fallo en conexión COM: {e}")
            raise Exception(f"No se pudo conectar con Outlook: {e}") from e

    @staticmethod
    def _uninitialize_com():
        """Libera el apartamento COM del hilo actual (idempotente)."""
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass
        logger.info("Recursos COM liberados.")

    # -- Utilidades ----------------------------------------------------------

    @staticmethod
    def _extract_oc(filename: str) -> str | None:
        """Extrae el código OC (Orden de Compra) del nombre del archivo.

        Busca patrones como:
          - "OC 15153500" / "OC15153500"
          - "OC-15153500"
          - "N°OC 15153500"
        Retorna el código sin el prefijo, o None si no se encuentra.
        """
        patterns = [
            r"[Oo][Cc][- ]?(\d{7,12})",  # OC-1234567 o OC 1234567
            r"[Nn]°[ ]*[Oo][Cc][- ]?(\d{7,12})",  # N°OC-1234567
        ]
        for pattern in patterns:
            m = re.search(pattern, filename)
            if m:
                return m.group(1).upper()
        return None

    @staticmethod
    def _sanitize_filename(filename: str, max_len: int = MAX_BASENAME_LEN) -> str:
        """Limpia caracteres inválidos de Windows y trunca para evitar MAX_PATH.

        La extensión se separa ANTES de truncar: si no, un nombre largo la
        perdería y el archivo se guardaría sin extensión utilizable.
        """
        cleaned = re.sub(r'[\\/:*?"<>|]', "_", filename).strip().rstrip(". ")
        if not cleaned:
            return "archivo"
        if len(cleaned) <= max_len:
            return cleaned

        stem, dot, ext = cleaned.rpartition(".")
        if dot and 0 < len(ext) <= 10:
            stem = stem[: max_len - len(ext) - 1].rstrip("_ .") or "archivo"
            return f"{stem}.{ext}"
        return cleaned[:max_len].rstrip("_ .") or "archivo"

    @staticmethod
    def _split_folder_path(folder_path: str) -> list[str]:
        return [p.strip() for p in folder_path.split(FOLDER_SEPARATOR) if p.strip()]

    def _resolve_folder(self, namespace, account_name: str, folder_path: str):
        """Navega cuenta > carpeta > subcarpeta... sin asumir Item(path).

        Si `account_name` no coincide exactamente con una cuenta de Outlook,
        se intenta buscar por nombre aproximado entre todas las cuentas
        disponibles (evita fallos cuando PickFolder() devuelve un nombre
        de cuenta distinto al esperado).
        """
        account_name = (account_name or "").strip()
        parts = self._split_folder_path(folder_path)
        if not account_name or not parts:
            raise Exception("Cuenta o carpeta vacía. Seleccione una carpeta válida.")

        account = None
        try:
            account = namespace.Folders.Item(account_name)
        except Exception:
            pass
        if account is None:
            # Fallback: buscar la cuenta por nombre aproximado
            try:
                root = namespace.Folders
                count = getattr(root, "Count", 0)
                for i in range(1, count + 1):
                    try:
                        candidate = root.Item(i)
                        cname = candidate.Name or ""
                    except Exception:
                        continue
                    if cname.lower() == account_name.lower():
                        account = candidate
                        account_name = cname
                        break
                    if account_name.lower() in cname.lower() or cname.lower() in account_name.lower():
                        account = candidate
                        account_name = cname
                        break
            except Exception as ex:
                raise Exception(f"No se encontró la cuenta '{account_name}' en Outlook.") from ex
        if account is None:
            raise Exception(f"No se encontró la cuenta '{account_name}' en Outlook.")

        current = account
        traversed: list[str] = []
        for part in parts:
            try:
                nxt = current.Folders.Item(part)
            except Exception as ex:
                raise Exception(
                    f"No se encontró la carpeta '{part}' en "
                    f"'{account_name}{' > ' + ' > '.join(traversed) if traversed else ''}'."
                ) from ex
            if nxt is None:
                raise Exception(f"No se encontró la carpeta '{part}'.")
            current = nxt
            traversed.append(part)
        return current

    @staticmethod
    def _as_date(value) -> Any | None:
        """Normaliza a datetime.date (acepta date o datetime)."""
        if value is None:
            return None
        if hasattr(value, "date") and callable(value.date):
            return value.date()
        return value

    @staticmethod
    def _build_date_restrict(start_date, end_date) -> str:
        """Construye un filtro Restrict de Outlook SOLO para [ReceivedTime].

        Usa formato YYYY/MM/DD HH:MM (año/mes/día) para evitar ambigüedades
        entre formato estadounidense (MM/DD/YYYY) e internacional (DD/MM/YYYY).
        La UI muestra dd/mm/yyyy; internamente se usa YYYY/MM/DD.
        """
        logger.debug(
            f"[_build_date_restrict] Entrada | start_date={start_date} "
            f"(tipo={type(start_date).__name__}) | "
            f"end_date={end_date} (tipo={type(end_date).__name__})"
        )
        clauses: list[str] = []
        if start_date:
            sd = OutlookEngine._as_date(start_date)
            start_dt = datetime.combine(sd, time.min)
            clause_start = f"[ReceivedTime] >= '{start_dt.strftime('%Y/%m/%d %H:%M')}'"
            clauses.append(clause_start)
            logger.debug(f"[_build_date_restrict] start_date parseado: sd={sd} | clausula='{clause_start}'")
        else:
            logger.debug("[_build_date_restrict] start_date es None/vacío")
        if end_date:
            ed = OutlookEngine._as_date(end_date)
            end_dt = datetime.combine(ed, time(23, 59, 59))
            clause_end = f"[ReceivedTime] <= '{end_dt.strftime('%Y/%m/%d %H:%M:%S')}'"
            clauses.append(clause_end)
            logger.debug(f"[_build_date_restrict] end_date parseado: ed={ed} | clausula='{clause_end}'")
        else:
            logger.debug("[_build_date_restrict] end_date es None/vacío")
        resultado = " AND ".join(clauses)
        logger.debug(f"[_build_date_restrict] Resultado final: '{resultado}'")
        return resultado

    @staticmethod
    def _match_keyword(item, keyword: str) -> bool:
        """Devuelve True si el correo contiene la keyword en Asunto o Cuerpo."""
        if not keyword:
            return True
        kw = keyword.strip().lower()
        if not kw:
            return True
        try:
            subject = getattr(item, "Subject", "") or ""
        except Exception:
            subject = ""
        try:
            body = getattr(item, "Body", "") or ""
        except Exception:
            body = ""
        return kw in subject.lower() or kw in body.lower()

    @staticmethod
    def _match_sender(item, sender_filter: str) -> bool:
        """Devuelve True si el remitente coincide con el filtro (email o dominio @dominio)."""
        if not sender_filter:
            return True
        sf = sender_filter.strip().lower()
        if not sf:
            return True
        try:
            sender = (getattr(item, "SenderEmailAddress", "") or "").lower()
        except Exception:
            sender = ""
        try:
            sender_name = (getattr(item, "SenderName", "") or "").lower()
        except Exception:
            sender_name = ""
        # Coincidencia exacta de email
        if sf == sender:
            return True
        # Filtro por dominio (@dominio.com)
        if sf.startswith("@") and sender.endswith(sf):
            return True
        # Substring en email o nombre de remitente
        return sf in sender or sf in sender_name

    @staticmethod
    def _attachment_allowed(filename: str, allowed_extensions: list[str] | None) -> bool:
        """Devuelve True si el archivo debe descargarse según las extensiones permitidas.

        - Si `allowed_extensions` es None o vacío, se permite todo (comportamiento retrocompatible).
        - Si hay lista, se filtra por extensión (case-insensitive).
        """
        if not allowed_extensions:
            return True
        _, ext = os.path.splitext(filename)
        return ext.lower() in [e.lower() for e in allowed_extensions]

    @staticmethod
    def _match_date_range(item, start_date, end_date) -> bool:
        """Filtra en Python por ReceivedTime cuando MAPI Restrict falla.

        Si start_date/end_date son None, el rango correspondiente queda abierto.
        """
        sd = OutlookEngine._as_date(start_date)
        ed = OutlookEngine._as_date(end_date)
        received_raw = getattr(item, "ReceivedTime", None)
        if received_raw is None:
            return False
        # Normalizar a datetime naive (Comparable sin tzinfo)
        if hasattr(received_raw, "replace") and hasattr(received_raw, "tzinfo"):
            received_dt = received_raw.replace(tzinfo=None)
        elif isinstance(received_raw, str):
            try:
                received_dt = datetime.strptime(received_raw, "%d/%m/%Y %H:%M")
            except ValueError:
                try:
                    received_dt = datetime.strptime(received_raw, "%m/%d/%Y %H:%M")
                except ValueError:
                    return False
        else:
            return False

        if sd and received_dt.date() < sd:
            return False
        return not (ed and received_dt.date() > ed)

    # -- API pública ----------------------------------------------------------

    def get_available_folders(self) -> list[dict[str, str]]:
        """Retorna una lista de carpetas disponibles en la cuenta principal."""
        logger.info("Consultando estructura de carpetas de Outlook...")
        _, namespace = self._initialize_com()
        folders_list: list[dict[str, str]] = []
        try:
            root = namespace.Folders
            for account in root:
                for folder in account.Folders:
                    try:
                        count = folder.Items.Count
                    except Exception:
                        count = 0
                    folders_list.append(
                        {
                            "account": account.Name,
                            "folder_name": folder.Name,
                            "count": count,
                        }
                    )
                    for subfolder in folder.Folders:
                        try:
                            sub_count = subfolder.Items.Count
                        except Exception:
                            sub_count = 0
                        folders_list.append(
                            {
                                "account": account.Name,
                                "folder_name": (f"{folder.Name}{FOLDER_SEPARATOR}{subfolder.Name}"),
                                "count": sub_count,
                            }
                        )
            return folders_list
        finally:
            self._uninitialize_com()

    def select_folder_dialog(self) -> dict[str, str] | None:
        """Abre el selector nativo de carpetas de Outlook (UI)."""
        _, namespace = self._initialize_com()
        try:
            folder = namespace.PickFolder()
            if folder is None:
                return None

            # Construir el path completo INCLUYENDO la cuenta como primer elemento.
            # El primer elemento del path es el nombre real de la cuenta en Outlook,
            # lo que evita fallos por comparaciones de objetos COM poco fiables.
            full_path = self._get_full_folder_path(namespace, folder)
            parts = self._split_folder_path(full_path)
            if not parts:
                return None
            account_name = parts[0]
            folder_path = FOLDER_SEPARATOR.join(parts[1:]) if len(parts) > 1 else parts[0]

            return {"account": account_name, "folder_name": folder_path}
        finally:
            self._uninitialize_com()

    def _get_full_folder_path(self, namespace, folder) -> str:
        """Construye el path completo de una carpeta INCLUYENDO la cuenta raíz.

        No compara objetos COM (poco fiable). En su lugar, obtiene los nombres
        de cuenta conocidos de ``namespace.Folders`` y se detiene al navegar
        hacia arriba cuando encuentra el primero que coincida.
        """
        account_names: set = set()
        try:
            for acc in namespace.Folders:
                try:
                    account_names.add(acc.Name)
                except Exception:
                    pass
        except Exception:
            pass

        path_parts: list[str] = []
        current = folder
        guard = 0
        while current is not None and guard < 30:
            guard += 1
            try:
                if hasattr(current, "Name"):
                    name = current.Name
                    path_parts.insert(0, name)
                    # Si este nombre coincide con una cuenta conocida,
                    # hemos llegado a la raíz de la cuenta -> detener.
                    if name in account_names:
                        break
                parent = current.Parent
                if parent is None:
                    break
                current = parent
            except Exception:
                break
        return FOLDER_SEPARATOR.join(path_parts) if path_parts else folder.Name

    def _get_folder_path(self, namespace, folder) -> str:
        """Construye el path completo de una carpeta (sin incluir la cuenta)."""
        full = self._get_full_folder_path(namespace, folder)
        parts = self._split_folder_path(full)
        return FOLDER_SEPARATOR.join(parts[1:]) if len(parts) > 1 else parts[0] if parts else folder.Name

    def start_extraction(
        self,
        account_name: str,
        folder_name: str,
        filters: dict[str, Any],
        download_attachments: bool,
        on_progress: Callable[[float, str], None],
        on_item_processed: Callable[[dict[str, Any]], None],
        on_complete: Callable[..., None],
    ):
        """Inicia el proceso de extracción en un hilo separado."""
        worker = threading.Thread(
            target=self._extraction_worker,
            args=(
                account_name,
                folder_name,
                filters,
                download_attachments,
                on_progress,
                on_item_processed,
                on_complete,
            ),
            daemon=True,
        )
        worker.start()

    def _unique_path(self, directory: str, base: str, suffix_ext: str) -> str:
        """Retorna un path que no existe agregando _1, _2... si hace falta."""
        candidate = os.path.join(directory, f"{base}{suffix_ext}")
        counter = 1
        while os.path.exists(candidate):
            candidate = os.path.join(directory, f"{base}_{counter}{suffix_ext}")
            counter += 1
        return candidate

    def _versioned_path(self, directory: str, base: str, ext: str, seen: dict, key: str) -> str:
        """Retorna un path con sufijo de versión (_v2, _v3...) si el key ya existe.

        El primer archivo conserva su nombre original. Los siguientes reciben _v2, _v3...
        para identificarlos como actualizaciones del mismo adjunto.
        """
        if key in seen:
            seen[key] += 1
            versioned_base = f"{base}_v{seen[key]}"
        else:
            seen[key] = 1
            versioned_base = base
        return os.path.join(directory, f"{versioned_base}{ext}")

    def _get_all_subfolders(self, folder) -> list:
        """Retorna lista de todas las subcarpetas recursivamente."""
        result = [folder]
        try:
            count = getattr(folder.Folders, "Count", 0)
            for i in range(1, count + 1):
                try:
                    sub = folder.Folders.Item(i)
                    result.extend(self._get_all_subfolders(sub))
                except Exception:
                    pass
        except Exception:
            pass
        return result

    def _extraction_worker(
        self,
        account_name,
        folder_name,
        filters,
        download_attachments,
        on_progress,
        on_item_processed,
        on_complete,
    ):
        """Lógica interna de extracción ejecutada en segundo plano."""
        _, namespace = self._initialize_com()
        self._is_cancelled.clear()
        extracted_data: list[dict[str, Any]] = []
        save_path = ""
        cancelled = False
        # Para identificar duplicados de adjuntos (_v2, _v3...)
        seen_attachments: dict[str, int] = {}
        attachments_saved = 0
        try:
            if namespace is None:
                raise Exception("No hay conexión con Outlook. Verifique que Outlook esté abierto.")

            logger.info("Resolviendo carpeta: cuenta='%s', carpeta='%s'", account_name, folder_name)
            target_folder = self._resolve_folder(namespace, account_name, folder_name)
            logger.info("Carpeta resuelta correctamente")

            # Si incluye subcarpetas, construir lista de todas las carpetas a procesar
            include_subfolders = filters.get("include_subfolders", False)
            folders_to_process = [target_folder]
            folder_labels = [folder_name]
            if include_subfolders:
                all_folders = self._get_all_subfolders(target_folder)
                folders_to_process = all_folders
                folder_labels = [f.Name for f in all_folders]
                logger.info(f"[SUBFOLDERS] Procesando {len(folders_to_process)} carpetas (incluye subcarpetas).")

            # Log detallado de filtros recibidos
            start_filtro = filters.get("start_date")
            end_filtro = filters.get("end_date")
            kw_filtro = filters.get("keyword")
            logger.info(
                f"[FILTROS_FECHA] start_date={start_filtro} "
                f"(tipo={type(start_filtro).__name__}) | "
                f"end_date={end_filtro} (tipo={type(end_filtro).__name__}) | "
                f"keyword='{kw_filtro}'"
            )

            # Preparar filtros compartidos
            keyword = (filters.get("keyword") or "").strip()
            sender_filter = (filters.get("sender") or "").strip()
            extensiones_permitidas = filters.get("extensiones") or []
            attachment_name_filter = (filters.get("attachment_name") or "").strip().lower()
            date_restrict = self._build_date_restrict(filters.get("start_date"), filters.get("end_date"))
            use_python_date_filter = bool(filters.get("start_date") or filters.get("end_date"))

            if download_attachments:
                base_dir = get_desktop_dir() / "G360_Extraidos"
                sd = OutlookEngine._as_date(filters.get("start_date"))
                ed = OutlookEngine._as_date(filters.get("end_date"))
                if sd and ed:
                    range_suffix = f"_{sd.strftime('%Y%m%d')}_hasta_{ed.strftime('%Y%m%d')}"
                elif sd:
                    range_suffix = f"_{sd.strftime('%Y%m%d')}_EnAdelante"
                elif ed:
                    range_suffix = f"_hasta_{ed.strftime('%Y%m%d')}"
                else:
                    range_suffix = ""
                folder_name_out = datetime.now().strftime("%Y-%m-%d") + range_suffix
                save_path = str(base_dir / folder_name_out)
                os.makedirs(save_path, exist_ok=True)

            # Procesar cada carpeta (principal + subcarpetas si aplica)
            total_folders = len(folders_to_process)
            global_total = 0
            global_processed = 0

            for fi, tf in enumerate(folders_to_process):
                folder_label = folder_labels[fi]
                logger.info(f"[SUBFOLDER {fi + 1}/{total_folders}] Procesando: {folder_label}")

                try:
                    items = tf.Items
                except Exception as ex:
                    logger.warning(f"[SUBFOLDER] No se pudo acceder a '{folder_label}': {ex}")
                    continue

                # Filtro MAPI Restrict por fecha
                use_py_filter = use_python_date_filter
                if date_restrict:
                    try:
                        restricted = items.Restrict(date_restrict)
                        if restricted is not None:
                            items = restricted
                            logger.info(f"[SUBFOLDER] Restrict aplicado en '{folder_label}'.")
                        else:
                            use_py_filter = True
                    except Exception as ex:
                        logger.warning(f"[SUBFOLDER] Restrict falló en '{folder_label}': {ex}")
                        use_py_filter = True

                try:
                    items.Sort("[ReceivedTime]", True)
                except Exception:
                    pass

                total_items = len(items)
                global_total += total_items
                if total_items == 0:
                    logger.info(f"[SUBFOLDER] '{folder_label}' vacía, saltando.")
                    continue

                # Progreso: porcentaje por carpeta
                base_pct = (fi / total_folders) * 0.9 if total_folders > 1 else 0.1
                on_progress(base_pct, f"Carpeta {fi + 1}/{total_folders}: {folder_label} ({total_items} items)")

                for i, item in enumerate(items):
                    if self._is_cancelled.is_set():
                        logger.info("Extracción interrumpida por el usuario.")
                        cancelled = True
                        break

                    if getattr(item, "Class", None) != 43:
                        continue

                    if keyword and not self._match_keyword(item, keyword):
                        continue
                    if sender_filter and not self._match_sender(item, sender_filter):
                        continue
                    if use_py_filter and not self._match_date_range(
                        item, filters.get("start_date"), filters.get("end_date")
                    ):
                        continue

                    try:
                        received_raw = getattr(item, "ReceivedTime", None)
                        if received_raw is None:
                            received_time = datetime.now()
                        elif hasattr(received_raw, "replace") and hasattr(received_raw, "date"):
                            received_time = received_raw.replace(tzinfo=None)
                        else:
                            received_time = datetime.now()
                        data = {
                            "subject": getattr(item, "Subject", None) or "Sin Asunto",
                            "sender": getattr(item, "SenderEmailAddress", None) or "Desconocido",
                            "received": received_time.strftime("%d/%m/%Y %H:%M"),
                            "size": getattr(item, "Size", 0),
                            "has_attachments": item.Attachments.Count > 0,
                            "folder": folder_label,
                        }
                        extracted_data.append(data)
                        on_item_processed(data)
                        logger.info(f"Procesando item: {data['subject']} de {data['sender']} [{folder_label}]")

                        if download_attachments:
                            date_prefix = received_time.strftime("%Y%m%d_%H%M")
                            base_name = self._sanitize_filename(f"{date_prefix}_{data['subject'][:30]}")

                            if filters.get("download_body", True):
                                body_full = self._unique_path(save_path, f"{base_name}_CUERPO", ".txt")
                                with open(body_full, "w", encoding="utf-8") as f:
                                    f.write(
                                        f"REMITENTE: {data['sender']}\n"
                                        f"FECHA: {data['received']}\n"
                                        f"ASUNTO: {data['subject']}\n"
                                        f"CARPETA: {folder_label}\n"
                                    )
                                    f.write("-" * 30 + "\n")
                                    f.write(getattr(item, "Body", "") or "")
                                logger.debug(f"Contexto guardado: {body_full}")

                            for attachment in item.Attachments:
                                try:
                                    original_name = attachment.FileName
                                except Exception:
                                    continue
                                if not OutlookEngine._attachment_allowed(original_name, extensiones_permitidas):
                                    continue
                                if attachment_name_filter and attachment_name_filter not in original_name.lower():
                                    continue
                                safe_name = self._sanitize_filename(original_name, max_len=100)
                                oc_code = self._extract_oc(original_name)
                                if oc_code:
                                    final_base = self._sanitize_filename(f"OC_{oc_code}_{safe_name}", max_len=100)
                                else:
                                    final_base = self._sanitize_filename(f"{date_prefix}_{safe_name}", max_len=100)
                                file_full_path = self._versioned_path(
                                    save_path, final_base, "", seen_attachments, safe_name
                                )
                                try:
                                    attachment.SaveAsFile(file_full_path)
                                    attachments_saved += 1
                                    logger.info(f"Adjunto guardado: {os.path.basename(file_full_path)}")
                                except Exception as ex:
                                    if "16" in str(ex) or "too many" in str(ex).lower():
                                        logger.warning(f"Esperando por muchos archivos: {original_name}")
                                        time.sleep(0.5)
                                        try:
                                            attachment.SaveAsFile(file_full_path)
                                            attachments_saved += 1
                                        except Exception as ex2:
                                            logger.warning(f"No se pudo guardar {original_name}: {ex2}")
                                    else:
                                        logger.warning(f"No se pudo guardar adjunto {original_name}: {ex}")
                    except Exception as ex:
                        logger.warning(f"Error procesando correo en '{folder_label}': {ex}")
                        continue

                    global_processed += 1
                    if global_total > 0:
                        pct = base_pct + (global_processed / global_total) * 0.85
                        on_progress(min(pct, 0.99), f"{folder_label}: {global_processed}/{global_total}")
                    if (fi * 1000 + i) % 50 == 0:
                        time.sleep(0.02)

            if global_total == 0:
                on_progress(1.0, "No se encontraron elementos con los filtros aplicados.")
                on_complete([], "", False)
                return

            cancelled = self._is_cancelled.is_set()
            if cancelled:
                on_progress(0.95, "Extracción cancelada por el usuario.")
            else:
                on_progress(1.0, f"Extracción completada. {len(extracted_data)} correos, {attachments_saved} adjuntos.")

            try:
                on_complete(extracted_data, save_path, cancelled, attachments_saved)
            except TypeError:
                # Compatibilidad con callbacks antiguos de 2 args
                on_complete(extracted_data, save_path)
        except Exception as e:
            logger.error(f"Error crítico en extracción: {e}", exc_info=True)
            error_msg = str(e).strip()
            # Traducir códigos COM comunes a mensajes legibles
            com_messages = {
                "15": "Error de acceso al correo (ERROR_GEN_FAILURE). Verifica que Outlook esté cerrado o que la carpeta tenga permisos.",
                "16": "Error de recursos agotados (ERROR_TOO_MANY_OPEN_FILES). Cierra Outlook y reintenta.",
                "9": "Error de handle inválido (ERROR_INVALID_HANDLE). La carpeta tiene muchos elementos; reduce el rango de fecha o cierra Outlook.",
                "-2147467259": "Error de automatización COM. Intenta reiniciar Outlook.",
                "-2147023675": "Error de acceso: verifica que Outlook esté abierto.",
            }
            if error_msg in com_messages:
                error_msg = com_messages[error_msg]
            elif error_msg.isdigit():
                error_msg = f"Error COM #{error_msg}. Revisa el log audit.log para más detalles."
            on_progress(0, f"Error crítico: {error_msg}")
            try:
                on_complete([], "", False, str(e), 0)
            except TypeError:
                try:
                    on_complete([], "", False)
                except TypeError:
                    on_complete([], "")
        finally:
            self._uninitialize_com()
