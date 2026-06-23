import win32com.client
import pythoncom
import threading
import logging
from datetime import datetime
import os
from typing import List, Callable, Dict, Any, Optional
import re

# Configuración del Sistema de Auditoría
logging.basicConfig(
    filename='audit.log',
    level=logging.INFO,
    format='%(asctime)s | %(levelname)s | %(message)s',
    encoding='utf-8'
)
logger = logging.getLogger("G360.OutlookEngine")

class OutlookEngine:
    """
    Motor de extracción MAPI (Outlook) optimizado para el ecosistema G360.
    Maneja la inicialización de hilos COM para evitar bloqueos en la interfaz Flet.
    """
    
    # Extensiones por categoría (basado en el prototipo exitoso)
    CATEGORIES = {
        "Excel": [".xls", ".xlsx", ".xlsm", ".csv"],
        "PDF": [".pdf"],
        "Docs": [".doc", ".docx", ".txt"]
    }

    def __init__(self):
        self._outlook = None
        self._namespace = None
        self._is_cancelled = threading.Event()

    def stop_extraction(self):
        """Señala al hilo de extracción que debe detenerse."""
        self._is_cancelled.set()
        logger.warning("Solicitud de cancelación de extracción recibida.")

    def _initialize_com(self):
        """Inicializa el entorno COM para el hilo actual."""
        # Crítico: Outlook COM requiere CoInitialize en cada hilo nuevo
        pythoncom.CoInitialize()
        try:
            self._outlook = win32com.client.Dispatch("Outlook.Application")
            self._namespace = self._outlook.GetNamespace("MAPI")
            logger.info("Conexión MAPI establecida exitosamente.")
        except Exception as e:
            logger.error(f"Fallo en conexión COM: {str(e)}")
            raise Exception(f"No se pudo conectar con Outlook: {str(e)}")

    def _uninitialize_com(self):
        """Libera los recursos COM del hilo."""
        self._outlook = None
        self._namespace = None
        pythoncom.CoUninitialize()
        logger.info("Recursos COM liberados.")

    def _sanitize_filename(self, filename: str) -> str:
        """Limpia caracteres inválidos para Windows (Mejorado de samples)"""
        return re.sub(r'[\\/:*?"<>|]', "_", filename)

    def get_available_folders(self) -> List[Dict[str, str]]:
        """Retorna una lista de carpetas disponibles en la cuenta principal."""
        logger.info("Consultando estructura de carpetas de Outlook...")
        self._initialize_com()
        folders_list = []
        try:
            # Obtenemos la carpeta raíz (generalmente la dirección de correo)
            root = self._namespace.Folders
            for account in root:
                for folder in account.Folders:
                    # Añadir carpeta principal (ej: Bandeja de entrada)
                    folders_list.append({
                        "account": account.Name,
                        "folder_name": folder.Name,
                        "count": folder.Items.Count
                    })
                    # Añadir subcarpetas (ej: Inbox/Facturas) para mayor visibilidad
                    for subfolder in folder.Folders:
                        folders_list.append({
                            "account": account.Name,
                            "folder_name": f"{folder.Name} > {subfolder.Name}",
                            "count": subfolder.Items.Count
                        })
            return folders_list
        finally:
            self._uninitialize_com()

    def select_folder_dialog(self) -> Optional[Dict[str, str]]:
        """Abre el selector nativo de carpetas de Outlook (UI)."""
        self._initialize_com()
        try:
            folder = self._namespace.PickFolder()
            if folder:
                # Buscar la cuenta raíz navegando hacia arriba en la jerarquía
                account_name = "Outlook"
                try:
                    current = folder
                    while current.Parent:
                        if hasattr(current.Parent, 'Class') and current.Parent.Class == 0:
                            # OlPublicFolders o similar - llegar a la cuenta
                            account_name = current.Name
                            break
                        if hasattr(current.Parent, 'Name'):
                            # Si el padre es la raíz del namespace, este es el nivel de cuenta
                            if current.Parent == self._namespace.Folders:
                                account_name = current.Name
                                break
                        current = current.Parent
                except:
                    account_name = "Outlook"
                
                # Construir el path completo de la carpeta
                folder_path = self._get_folder_path(folder)
                
                return {"account": account_name, "folder_name": folder_path}
            return None
        finally:
            self._uninitialize_com()

    def _get_folder_path(self, folder) -> str:
        """Construye el path completo de una carpeta."""
        path_parts = []
        current = folder
        while current:
            try:
                if hasattr(current, 'Name'):
                    path_parts.insert(0, current.Name)
                if current.Parent == self._namespace.Folders:
                    break
                current = current.Parent
            except:
                break
        return " > ".join(path_parts[1:]) if len(path_parts) > 1 else path_parts[0] if path_parts else folder.Name

    def start_extraction(self, 
                         account_name: str, 
                         folder_name: str, 
                         filters: Dict[str, Any],
                         download_attachments: bool,
                         on_progress: Callable[[float, str], None],
                         on_item_processed: Callable[[Dict[str, Any]], None],
                         on_complete: Callable[[List[Dict[str, Any]], str], None]):
        """Inicia el proceso de extracción en un hilo separado."""
        worker = threading.Thread(
            target=self._extraction_worker,
            args=(account_name, folder_name, filters, download_attachments, on_progress, on_item_processed, on_complete),
            daemon=True
        )
        worker.start()

    def _extraction_worker(self, account_name, folder_name, filters, download_attachments, on_progress, on_item_processed, on_complete):
        """Lógica interna de extracción ejecutada en segundo plano."""
        self._initialize_com()
        self._is_cancelled.clear()
        extracted_data = []
        save_path = ""
        try:
            # Verificar que la conexión MAPI esté activa
            if not self._namespace:
                raise Exception("No hay conexión con Outlook. Verifique que Outlook esté abierto.")
            
            # Localizar carpeta
            try:
                account = self._namespace.Folders.Item(account_name)
                if not account:
                    raise Exception(f"No se encontró la cuenta '{account_name}' en Outlook.")
                target_folder = account.Folders.Item(folder_name)
                if not target_folder:
                    raise Exception(f"No se encontró la carpeta '{folder_name}' en la cuenta '{account_name}'.")
            except Exception as ex:
                if "No se encontró" in str(ex):
                    raise ex
                raise Exception(f"Error al acceder a la carpeta: {str(ex)}. Verifique que Outlook esté abierto y la carpeta exista.")
            
            # Aplicar filtro de fecha y palabra clave (Optimización MAPI)
            filter_query = ""
            if filters.get("start_date") and filters.get("end_date"):
                start_str = filters["start_date"].strftime("%m/%d/%Y %I:%M %p")
                end_str = filters["end_date"].strftime("%m/%d/%Y %I:%M %p")
                filter_query = f"[ReceivedTime] >= '{start_str}' AND [ReceivedTime] <= '{end_str}'"
            
            if filters.get("keyword"):
                kw = filters["keyword"].replace("'", "''")
                kw_query = f"@SQL=\"urn:schemas:httpmail:subject\" LIKE '%{kw}%' OR \"urn:schemas:httpmail:textdescription\" LIKE '%{kw}%'"
                filter_query = f"{filter_query} AND ({kw_query})" if filter_query else kw_query

            items = target_folder.Items
            if filter_query:
                items = items.Restrict(filter_query)

            # Ordenar por fecha de recepción (descendente)
            items.Sort("[ReceivedTime]", True)
            
            total_items = len(items)
            if total_items == 0:
                on_progress(1.0, "No se encontraron elementos con los filtros aplicados.")
                on_complete([], "") # Aseguramos siempre pasar dos argumentos
                return

            if download_attachments:
                desktop = os.path.join(os.environ['USERPROFILE'], 'Desktop')
                # Carpeta raíz única y subcarpeta por fecha de procesamiento
                base_dir = os.path.join(desktop, "G360_Extraidos")
                processing_date = datetime.now().strftime('%Y-%m-%d')
                save_path = os.path.join(base_dir, processing_date)
                os.makedirs(save_path, exist_ok=True)

            on_progress(0.1, f"Iniciando extracción de {total_items} elementos...")

            for i, item in enumerate(items):
                # Verificar cancelación en cada iteración
                if self._is_cancelled.is_set():
                    logger.info("Extracción interrumpida por el usuario.")
                    break

                # Verificar si es un objeto de correo (olMail = 43)
                if item.Class == 43:
                    try:
                        received_time = item.ReceivedTime.replace(tzinfo=None)
                        data = {
                            "subject": getattr(item, "Subject", "Sin Asunto"),
                            "sender": getattr(item, "SenderEmailAddress", "Desconocido"),
                            "received": received_time.strftime("%d/%m/%Y %H:%M"),
                            "size": getattr(item, "Size", 0),
                            "has_attachments": item.Attachments.Count > 0
                        }
                        extracted_data.append(data)
                        on_item_processed(data)
                        logger.info(f"Procesando item: {data['subject']} de {data['sender']}")

                        if download_attachments:
                            # Nombre base para archivos de este correo
                            date_prefix = received_time.strftime("%Y%m%d_%H%M")
                            base_name = self._sanitize_filename(f"{date_prefix}_{data['subject'][:30]}")
                            
                            # 1. Guardar cuerpo del correo como TXT (Explicación para almacén)
                            body_path = os.path.join(save_path, f"{base_name}_CUERPO.txt")
                            with open(body_path, "w", encoding="utf-8") as f:
                                f.write(f"REMITENTE: {data['sender']}\nFECHA: {data['received']}\nASUNTO: {data['subject']}\n")
                                f.write("-" * 30 + "\n")
                                f.write(getattr(item, "Body", ""))
                            logger.debug(f"Contexto guardado: {body_path}")

                            # 2. Guardar Adjuntos (OC, Reportes, etc.)
                            for attachment in item.Attachments:
                                original_name = attachment.FileName
                                name, ext = os.path.splitext(original_name)
                                final_filename = self._sanitize_filename(f"{date_prefix}_{original_name}")
                                file_full_path = os.path.join(save_path, final_filename)
                                
                                # Manejo de duplicados idéntico al prototipo
                                counter = 1
                                while os.path.exists(file_full_path):
                                    file_full_path = os.path.join(save_path, f"{date_prefix}_{name}_{counter}{ext}")
                                    counter += 1
                                
                                attachment.SaveAsFile(file_full_path)
                                logger.info(f"Adjunto guardado: {final_filename}")

                    except Exception as ex:
                        logger.warning(f"Error procesando correo: {str(ex)}")
                        continue # Ignorar correos corruptos o inaccesibles

                # Reportar progreso cada 5% o al finalizar
                if i % max(1, (total_items // 20)) == 0:
                    on_progress(i / total_items, f"Procesando: {i}/{total_items}")

            if self._is_cancelled.is_set():
                on_progress(i / total_items if total_items > 0 else 1.0, "Extracción cancelada por el usuario.")
            else:
                on_progress(1.0, "Extracción completada con éxito.")
            
            on_complete(extracted_data, save_path)
        except Exception as e:
            logger.error(f"Error crítico en extracción: {str(e)}")
            on_progress(0, f"Error crítico: {str(e)}")
            on_complete([], "")  # Asegurar reset de UI incluso en error
        finally:
            self._uninitialize_com()