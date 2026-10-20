import logging
import os
from datetime import datetime

import flet as ft

# Importación explícita para que PyInstaller incluya flet_desktop en el build
try:
    import flet_desktop  # noqa: F401
except ImportError:
    pass

try:
    from .engine import OutlookEngine, get_audit_log_path
except ImportError:
    from engine import OutlookEngine, get_audit_log_path  # type: ignore

# Logger dedicado para depuración de fechas (solo en archivo, no consola)
logging.basicConfig(
    filename=str(get_audit_log_path()),
    level=logging.INFO,
    format="%(asctime)s | %(levelname)s | %(name)s | %(message)s",
    encoding="utf-8",
)
DATE_LOGGER = logging.getLogger("G360.OutlookExtractor.Dates")

try:
    from g360_flet.g360_signature import G360Signature
except ImportError:
    G360Signature = None

G360_THEME = {
    # TEMA OSCURO (flet-desktop tokens G360 oficiales + WCAG AA)
    "bg": "#0b1220",
    "surface": "#1a2332",
    "accent": "#00d084",
    "text": "#f0f4f8",
    "text_dim": "#94a3b8",
    "text_placeholder": "#7c8db0",
    "border": "#5a6a80",
    "border_focus": "#00d084",
    "danger": "#b91c1c",
    "danger_hover": "#f87171",
    "warning": "#f59e0b",
    "success": "#22c55e",
    "info": "#60a5fa",
    # TEMA CLARO (flet-desktop-corporativo tokens G360 oficiales + WCAG AA)
    "bg_light": "#ffffff",
    "surface_light": "#f8fafc",
    "accent_light": "#00796B",
    "text_light": "#1e293b",
    "text_dim_light": "#64748b",
    "text_placeholder_light": "#374151",
    "border_light": "#64748b",
    "danger_light": "#b91c1c",
    "warning_light": "#f59e0b",
    "success_light": "#10b981",
    "info_light": "#00796B",
}

# Estado global del tema (default: oscuro)
_current_theme_mode = "dark"


def _set_theme_mode(mode: str):
    """Actualiza el modo de tema actual (dark/light)."""
    global _current_theme_mode
    _current_theme_mode = mode


def _get_theme_color(key: str) -> str:
    """Retorna el color del tema activo."""
    suffix = f"_{_current_theme_mode}" if _current_theme_mode != "dark" else ""
    return G360_THEME.get(f"{key}{suffix}", G360_THEME.get(key, ""))


class G360OutlookApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.engine = OutlookEngine()
        self.all_folders = []
        self._extraction_running = False
        self.setup_page()
        self.init_ui()

    def setup_page(self):
        self.page.title = "G360 | Outlook Extractor"
        self.page.bgcolor = G360OutlookApp._t("bg")
        self.page.padding = 0
        self.page.theme_mode = ft.ThemeMode.DARK
        # Flet >= 0.23: el tamaño de ventana vive en Page.window
        self.page.window.width = 600
        self.page.window.height = 700
        self.page.window.min_width = 520
        self.page.window.min_height = 550
        self.page.window.resizable = True
        self.page.on_close = self.on_close

        self.audit_rows = []
        self.audit_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Fecha", weight="bold")),
                ft.DataColumn(ft.Text("Remitente", weight="bold")),
                ft.DataColumn(ft.Text("Asunto", weight="bold")),
            ],
            rows=self.audit_rows,
            heading_row_color=G360OutlookApp._t("surface"),
        )
        self.global_status = ft.Text("Listo", color=G360OutlookApp._t("text_dim"))

    def init_ui(self):
        # Sidebar Lateral
        # El toggle de tema se mueve al leading para dejar el trailing limpio
        self.theme_toggle = ft.Switch(
            value=False,
            label="Tema",
            on_change=self._toggle_theme,
            thumb_color=ft.Colors.WHITE,
            active_color=G360OutlookApp._t("accent"),
        )
        self.nav_rail = ft.NavigationRail(
            selected_index=0,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=90,
            min_extended_width=200,
            bgcolor=G360OutlookApp._t("surface"),
            leading=ft.Column(
                controls=[
                    ft.Container(height=8),
                    self.theme_toggle,
                    ft.Container(height=6),
                    ft.Icon(ft.Icons.MAIL_OUTLINE, color=G360OutlookApp._t("accent"), size=24),
                ],
                alignment=ft.alignment.Alignment.CENTER,
            ),
            trailing=ft.Column(
                controls=[
                    ft.Container(expand=True),
                    G360Signature(mode="logo")
                    if G360Signature
                    else ft.Text("•> G360", color=G360OutlookApp._t("text_dim"), size=8),
                    ft.Container(height=6),
                ],
                alignment=ft.alignment.Alignment.CENTER,
            ),
            destinations=[
                ft.NavigationRailDestination(icon=ft.Icons.HOME_OUTLINED, label="Inicio"),
                ft.NavigationRailDestination(icon=ft.Icons.STORAGE_OUTLINED, label="Extracción"),
                ft.NavigationRailDestination(icon=ft.Icons.CHECK_CIRCLE_OUTLINE, label="Auditoría"),
            ],
            on_change=self.on_nav_change,
        )

        self.main_content = ft.Container(
            expand=True,
            padding=30,
            bgcolor=G360OutlookApp._t("bg"),
        )

        self.page.add(
            ft.Row(
                [
                    self.nav_rail,
                    ft.VerticalDivider(width=1, thickness=1, color=G360OutlookApp._t("border")),
                    self.main_content,
                ],
                expand=True,
                spacing=0,
            )
        )

        # Cargamos la vista de inicio por defecto
        self.show_home_view()
        self.page.update()

    def on_nav_change(self, e):
        index = e.control.selected_index
        if index == 0:
            self.show_home_view()
        elif index == 1:
            self.show_extraction_view()
        elif index == 2:
            self.show_audit_log_view()
        self.page.update()

    def show_home_view(self):
        self.main_content.content = ft.Column(
            expand=True,
            controls=[
                ft.Row(
                    [
                        ft.Text(
                            "Outlook Extractor", size=22, weight=ft.FontWeight.BOLD, color=G360OutlookApp._t("text")
                        ),
                        ft.Container(
                            content=ft.Text("v1.0.0", size=10, color=G360OutlookApp._t("accent")),
                            border=ft.Border.all(1, G360OutlookApp._t("accent")),
                            border_radius=5,
                            padding=ft.Padding(6, 2, 6, 2),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
                ft.Text(
                    "Herramienta de Auditoría y Extracción MAPI",
                    color=G360OutlookApp._t("text_dim"),
                    size=12,
                ),
                ft.Divider(height=40, color=G360OutlookApp._t("border")),
                ft.Container(
                    content=ft.Column(
                        [
                            ft.Text("Resumen del Sistema", weight="bold", size=20, color=G360OutlookApp._t("accent")),
                            ft.Container(
                                content=ft.Column(
                                    [
                                        ft.Row(
                                            [
                                                ft.Icon(
                                                    ft.Icons.MAIL_OUTLINE, color=G360OutlookApp._t("info"), size=24
                                                ),
                                                ft.Text(
                                                    "Correos Seleccionados:", size=14, color=G360OutlookApp._t("text")
                                                ),
                                                ft.Text(
                                                    "0",
                                                    size=14,
                                                    weight=ft.FontWeight.BOLD,
                                                    color=G360OutlookApp._t("accent"),
                                                ),
                                            ],
                                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                        ),
                                        ft.Container(height=10),
                                        ft.Row(
                                            [
                                                ft.Icon(ft.Icons.DOWNLOAD, color=G360OutlookApp._t("success"), size=24),
                                                ft.Text(
                                                    "Archivos Descargados:", size=14, color=G360OutlookApp._t("text")
                                                ),
                                                ft.Text(
                                                    "0",
                                                    size=14,
                                                    weight=ft.FontWeight.BOLD,
                                                    color=G360OutlookApp._t("success"),
                                                ),
                                            ],
                                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                        ),
                                        ft.Container(height=10),
                                        ft.Row(
                                            [
                                                ft.Icon(
                                                    ft.Icons.FOLDER_SPECIAL, color=G360OutlookApp._t("warning"), size=24
                                                ),
                                                ft.Text(
                                                    "Carpetas Procesadas:", size=14, color=G360OutlookApp._t("text")
                                                ),
                                                ft.Text(
                                                    "0",
                                                    size=14,
                                                    weight=ft.FontWeight.BOLD,
                                                    color=G360OutlookApp._t("warning"),
                                                ),
                                            ],
                                            alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                                        ),
                                    ],
                                    spacing=10,
                                ),
                                padding=15,
                                bgcolor=G360OutlookApp._t("surface"),
                                border_radius=8,
                                border=ft.Border.all(1, G360OutlookApp._t("border")),
                            ),
                            ft.Container(height=20),
                            ft.Text("Actividad Reciente", weight="bold", size=18, color=G360OutlookApp._t("accent")),
                            ft.Container(
                                content=ft.Column(
                                    [
                                        ft.Text(
                                            "No hay actividad reciente.",
                                            size=12,
                                            italic=True,
                                            color=G360OutlookApp._t("text_dim"),
                                        ),
                                    ],
                                    alignment=ft.CrossAxisAlignment.CENTER,
                                ),
                                height=100,
                                bgcolor=G360OutlookApp._t("surface"),
                                border_radius=8,
                                border=ft.Border.all(1, G360OutlookApp._t("border")),
                                alignment=ft.Alignment(0, 0),
                            ),
                        ],
                        spacing=20,
                    ),
                    padding=20,
                    expand=True,
                    bgcolor=G360OutlookApp._t("bg"),
                ),
                self.global_status,
            ],
        )
        self.page.update()

    def _safe_update(self):
        """Update tolerante a hilos/cierre de ventana.

        Se usa Page.update() (no Control.update()) porque Flet >= 0.70
        lanza RuntimeError si el control todavía no está montado.
        """
        try:
            self.page.update()
        except Exception:
            pass

    def show_audit_log_view(self):
        audit_path = str(get_audit_log_path())
        log_content = "No se encontró el archivo de auditoría."
        if os.path.exists(audit_path):
            try:
                with open(audit_path, encoding="utf-8") as f:
                    lines = f.readlines()
                    log_content = "".join(lines[-50:])
            except Exception as e:
                log_content = f"Error al leer logs: {str(e)}"

        self.main_content.content = ft.Column(
            expand=True,
            scroll=ft.ScrollMode.ALWAYS,
            controls=[
                ft.Text("Auditoría de Sistema", size=22, weight=ft.FontWeight.BOLD, color=G360OutlookApp._t("text")),
                ft.Text(
                    "Historial técnico de operaciones y errores (audit.log)",
                    color=G360OutlookApp._t("text_dim"),
                    size=11,
                ),
                ft.Divider(height=20, color=G360OutlookApp._t("border")),
                ft.Container(
                    content=ft.Text(log_content, font_family="Consolas", size=11, color=G360OutlookApp._t("accent")),
                    bgcolor=G360OutlookApp._t("surface"),
                    padding=12,
                    border_radius=8,
                    border=ft.Border.all(1, G360OutlookApp._t("border")),
                ),
                ft.ElevatedButton(
                    "Actualizar Logs", icon=ft.Icons.REFRESH, on_click=lambda _: self.show_audit_log_view()
                ),
            ],
        )
        self.page.update()

    def show_extraction_view(self):
        self.folder_dropdown = ft.Dropdown(
            label="Seleccionar Carpeta de Outlook",
            hint_text="Cargue carpetas o use el explorador",
            border_color=G360OutlookApp._t("border"),
            color=G360OutlookApp._t("text"),
            expand=True,
        )
        self.btn_connect = ft.ElevatedButton(
            content=ft.Text("Listar Carpetas", color=G360OutlookApp._t("text")),
            icon=ft.Icons.FOLDER_OPEN,
            style=ft.ButtonStyle(color=G360OutlookApp._t("text"), bgcolor=G360OutlookApp._t("surface")),
            on_click=self.on_load_folders,
        )
        self.btn_pick_folder = ft.ElevatedButton(
            content=ft.Text("Explorar Carpetas (Outlook)", color=G360OutlookApp._t("bg")),
            icon=ft.Icons.FOLDER_SPECIAL,
            style=ft.ButtonStyle(color=G360OutlookApp._t("bg"), bgcolor=G360OutlookApp._t("accent")),
            on_click=self.on_pick_outlook_folder,
        )

        self.btn_extract = ft.ElevatedButton(
            content=ft.Text("Iniciar Extracción", color=G360OutlookApp._t("bg")),
            icon=ft.Icons.PLAY_ARROW_ROUNDED,
            style=ft.ButtonStyle(color=G360OutlookApp._t("bg"), bgcolor=G360OutlookApp._t("accent")),
            on_click=self.on_start_extraction,
            disabled=True,
        )
        self.btn_cancel = ft.ElevatedButton(
            content=ft.Text("Cancelar", color=ft.Colors.WHITE),
            icon=ft.Icons.CANCEL,
            style=ft.ButtonStyle(bgcolor=G360OutlookApp._t("danger")),
            on_click=self.on_cancel_extraction,
            disabled=True,
        )
        self.btn_reset = ft.ElevatedButton(
            content=ft.Text("Reiniciar", color=G360OutlookApp._t("text")),
            icon=ft.Icons.REFRESH,
            style=ft.ButtonStyle(color=G360OutlookApp._t("text"), bgcolor=G360OutlookApp._t("surface")),
            on_click=self.on_reset,
            disabled=False,
        )
        self.progress_bar = ft.ProgressBar(
            width=None,
            height=10,
            color=G360OutlookApp._t("accent"),
            bgcolor=G360OutlookApp._t("border"),
            value=0,
            visible=False,
        )
        self.status_text = ft.Text("", color=G360OutlookApp._t("text_dim"))
        self.btn_start_date = ft.TextField(
            label="Fecha Inicio (dd/mm/yyyy)",
            hint_text="Ej: 01/01/2024",
            width=150,
            value="",
        )
        self.btn_end_date = ft.TextField(
            label="Fecha Fin (dd/mm/yyyy)",
            hint_text="Ej: 31/12/2024",
            width=150,
            value="",
        )
        self._start_date_picker = ft.DatePicker(
            first_date=datetime(2020, 1, 1),
            last_date=datetime(2030, 12, 31),
            on_change=self._on_start_date_pick,
            confirm_text="Confirmar",
            cancel_text="Cancelar",
        )
        self._end_date_picker = ft.DatePicker(
            first_date=datetime(2020, 1, 1),
            last_date=datetime(2030, 12, 31),
            on_change=self._on_end_date_pick,
            confirm_text="Confirmar",
            cancel_text="Cancelar",
        )
        self.btn_start_date_picker = ft.IconButton(
            icon=ft.Icons.CALENDAR_MONTH,
            tooltip="Seleccionar fecha de inicio",
            icon_color=G360OutlookApp._t("accent"),
            on_click=lambda _: self.page.show_dialog(self._start_date_picker),
        )
        self.btn_end_date_picker = ft.IconButton(
            icon=ft.Icons.CALENDAR_MONTH,
            tooltip="Seleccionar fecha fin",
            icon_color=G360OutlookApp._t("accent"),
            on_click=lambda _: self.page.show_dialog(self._end_date_picker),
        )
        self.check_download = ft.Checkbox(
            label="Descargar Adjuntos y Cuerpo (.txt) al Escritorio",
            value=True,
            fill_color=G360OutlookApp._t("accent"),
        )
        self.check_attachments_only = ft.Checkbox(
            label="Solo adjuntos (sin cuerpo)",
            value=False,
            fill_color=G360OutlookApp._t("accent"),
        )
        self.check_open_folder = ft.Checkbox(
            label="Abrir carpeta al finalizar",
            value=True,
            fill_color=G360OutlookApp._t("accent"),
        )
        self.check_subfolders = ft.Checkbox(
            label="Incluir subcarpetas",
            value=False,
            fill_color=G360OutlookApp._t("accent"),
        )
        self.check_xlsx = ft.Checkbox(
            label="Excel (.xls, .xlsx, .xlsm, .csv)",
            value=True,
            fill_color=G360OutlookApp._t("accent"),
        )
        self.check_pdf = ft.Checkbox(
            label="PDF (.pdf)",
            value=False,
            fill_color=G360OutlookApp._t("accent"),
        )
        self.check_docs = ft.Checkbox(
            label="Docs (.doc, .docx, .txt)",
            value=False,
            fill_color=G360OutlookApp._t("accent"),
        )
        self.search_folder = ft.TextField(
            label="Filtrar carpetas...",
            prefix_icon=ft.Icons.SEARCH,
            on_change=self.filter_dropdown_options,
            border_color=G360OutlookApp._t("border"),
            height=45,
            expand=True,
        )
        self.keyword_filter = ft.TextField(
            label="Palabra clave (Asunto/Cuerpo)",
            border_color=G360OutlookApp._t("border"),
            expand=True,
        )
        self.sender_filter = ft.TextField(
            label="Filtrar por Remitente (email o dominio)",
            hint_text="Ej: proveedor@empresa.com o @dominio.com",
            border_color=G360OutlookApp._t("border"),
            expand=True,
        )
        self.attachment_name_filter = ft.TextField(
            label="Filtrar adjuntos por nombre",
            hint_text="Ej: lista de precio, factura, OC...",
            border_color=G360OutlookApp._t("border"),
        )

        self.main_content.content = ft.Container(
            content=ft.Column(
                scroll=ft.ScrollMode.AUTO,
                expand=True,
                controls=[
                    ft.Text(
                        "Módulo de Extracción", size=20, weight=ft.FontWeight.BOLD, color=G360OutlookApp._t("text")
                    ),
                    ft.Text(
                        "Seleccione una cuenta y carpeta para iniciar el proceso de auditoría.",
                        color=G360OutlookApp._t("text_dim"),
                    ),
                    ft.Divider(height=20, color=G360OutlookApp._t("border")),
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Text(
                                    "Origen de Datos (MAPI)",
                                    weight="bold",
                                    size=14,
                                    color=G360OutlookApp._t("accent"),
                                ),
                                ft.Row([self.btn_pick_folder, self.btn_connect], spacing=8),
                                self.search_folder,
                                self.folder_dropdown,
                            ],
                            spacing=10,
                        ),
                        padding=14,
                        bgcolor=G360OutlookApp._t("surface"),
                        border_radius=8,
                        border=ft.Border.all(1, G360OutlookApp._t("border")),
                    ),
                    ft.Container(height=5),
                    ft.Container(
                        content=ft.Column(
                            [
                                ft.Text(
                                    "Parámetros de Auditoría",
                                    weight="bold",
                                    size=14,
                                    color=G360OutlookApp._t("accent"),
                                ),
                                ft.Row(
                                    [
                                        ft.Row([self.btn_start_date, self.btn_start_date_picker], spacing=4),
                                        ft.Row([self.btn_end_date, self.btn_end_date_picker], spacing=4),
                                    ],
                                    spacing=10,
                                ),
                                self.sender_filter,
                                self.keyword_filter,
                                self.attachment_name_filter,
                                ft.Row([self.check_download, self.check_open_folder], wrap=True),
                                self.check_attachments_only,
                                self.check_subfolders,
                                ft.Row([self.check_xlsx, self.check_pdf, self.check_docs], wrap=True),
                                ft.Row([self.btn_extract, self.btn_cancel, self.btn_reset], spacing=10),
                            ],
                            spacing=10,
                        ),
                        padding=14,
                        bgcolor=G360OutlookApp._t("surface"),
                        border_radius=8,
                        border=ft.Border.all(1, G360OutlookApp._t("border")),
                    ),
                    ft.Container(height=12),
                    self.status_text,
                    self.progress_bar,
                    ft.Text(
                        "Auditoría de Procesamiento (Tiempo Real)",
                        size=12,
                        weight="bold",
                        color=G360OutlookApp._t("accent"),
                    ),
                    ft.Container(
                        content=self.audit_table,
                        height=280,
                        border=ft.Border.all(1, G360OutlookApp._t("border")),
                        border_radius=8,
                        padding=8,
                        bgcolor=G360OutlookApp._t("surface"),
                    ),
                ],
            ),
        )
        self.page.update()

    def on_pick_outlook_folder(self, e):
        self.status_text.value = "Abriendo selector de Outlook..."
        self.page.update()

        def pick():
            try:
                folder = self.engine.select_folder_dialog()
                if folder:
                    self.folder_dropdown.options = [
                        ft.dropdown.Option(
                            key=f"{folder['account']}|{folder['folder_name']}",
                            text=f"{folder['account']} > {folder['folder_name']}",
                        )
                    ]
                    self.folder_dropdown.value = f"{folder['account']}|{folder['folder_name']}"
                    self.btn_extract.disabled = False
                    self.status_text.value = f"Seleccionado: {folder['folder_name']}"
                else:
                    self.status_text.value = "Selección cancelada."
            except Exception as ex:
                self.status_text.value = f"Error: {str(ex)}"
            self._safe_update()

        # Flet >= 0.23: page.run_thread gestiona el contexto de la sesión
        self.page.run_thread(pick)

    def on_load_folders(self, e):
        self.btn_connect.disabled = True
        self.status_text.value = "Cargando carpetas MAPI..."
        self.page.update()

        def load():
            try:
                self.all_folders = self.engine.get_available_folders()
                self.folder_dropdown.options = [
                    ft.dropdown.Option(
                        key=f"{f['account']}|{f['folder_name']}",
                        text=f"{f['account']} > {f['folder_name']} ({f['count']} elementos)",
                    )
                    for f in self.all_folders
                ]
                self.folder_dropdown.hint_text = "Seleccione una carpeta"
                if self.all_folders:
                    self.btn_extract.disabled = False
                self.status_text.value = f"Listo: {len(self.all_folders)} carpetas."
            except Exception as ex:
                self.status_text.value = f"Error: {str(ex)}"
            self.btn_connect.disabled = False
            self._safe_update()

        self.page.run_thread(load)

    def filter_dropdown_options(self, e):
        search_term = (self.search_folder.value or "").strip().lower()
        filtered = [
            f
            for f in self.all_folders
            if search_term in f["folder_name"].lower() or search_term in f["account"].lower()
        ]
        self.folder_dropdown.options = [
            ft.dropdown.Option(
                key=f"{f['account']}|{f['folder_name']}",
                text=f"{f['account']} > {f['folder_name']} ({f['count']} elementos)",
            )
            for f in filtered
        ]
        self.page.update()

    def on_start_extraction(self, e):
        if not self.engine:
            self.status_text.value = "Primero conecte Outlook"
            self.page.update()
            return

        if not self.folder_dropdown.value:
            self.status_text.value = "Seleccione una carpeta primero"
            self.page.update()
            return

        try:
            acc, fol = self.folder_dropdown.value.split("|", 1)
            acc, fol = acc.strip(), fol.strip()
            if not acc or not fol:
                raise ValueError("vacío")
        except (ValueError, AttributeError):
            self.status_text.value = "Error en la selección de carpeta"
            self.page.update()
            return

        self.btn_extract.disabled = True
        self.btn_cancel.disabled = False
        self.progress_bar.visible = True
        self.audit_table.rows.clear()
        self._extraction_running = True
        # Leer fechas de TextField (formato dd/mm/yyyy)
        start_text = (self.btn_start_date.value or "").strip()
        end_text = (self.btn_end_date.value or "").strip()
        start_fecha = self._parse_fecha(start_text)
        end_fecha = self._parse_fecha(end_text)

        # Validar fechas: si el usuario escribió algo pero falló el parseo, avisar.
        date_error = ""
        if (start_text and start_fecha is None) or (end_text and end_fecha is None):
            date_error = "Formato de fecha inválido. Use dd/mm/yyyy (ej: 01/01/2024)."
        elif start_fecha and end_fecha and start_fecha > end_fecha:
            date_error = "La fecha de inicio es posterior a la fecha fin."

        if date_error:
            self.btn_start_date.error_text = date_error if start_text and start_fecha is None else None
            self.btn_end_date.error_text = date_error if end_text and end_fecha is None else None
            self.status_text.value = date_error
            self.btn_extract.disabled = False
            self._safe_update()
            return

        # Filtro de extensiones de archivo seleccionadas
        extensiones = []
        if self.check_xlsx.value:
            extensiones.extend([".xls", ".xlsx", ".xlsm", ".csv"])
        if self.check_pdf.value:
            extensiones.append(".pdf")
        if self.check_docs.value:
            extensiones.extend([".doc", ".docx", ".txt"])

        DATE_LOGGER.debug(
            f"[on_start_extraction] Fechas | inicio={start_fecha} | fin={end_fecha} | extensiones={extensiones}"
        )

        filters = {
            "keyword": (self.keyword_filter.value or "").strip(),
            "sender": (self.sender_filter.value or "").strip(),
            "attachment_name": (self.attachment_name_filter.value or "").strip(),
            "start_date": start_fecha,
            "end_date": end_fecha,
            "extensiones": extensiones,
            "download_body": self.check_download.value and not self.check_attachments_only.value,
            "include_subfolders": self.check_subfolders.value,
        }
        self.engine.start_extraction(
            acc,
            fol,
            filters,
            self.check_download.value or self.check_attachments_only.value,
            self.handle_progress,
            self.on_item_processed,
            self.on_complete,
        )

    def on_cancel_extraction(self, e):
        if self.engine:
            self.btn_cancel.disabled = True
            self.status_text.value = "Cancelando proceso de extracción..."
            self.page.update()
            self.engine.stop_extraction()

    def on_reset(self, e):
        """Reinicia la interfaz para una nueva extracción."""
        # Cancelar extracción en curso
        if self.engine:
            self.engine.stop_extraction()
        # Limpiar campos
        self.folder_dropdown.value = None
        self.folder_dropdown.options = []
        self.btn_start_date.value = ""
        self.btn_end_date.value = ""
        self.keyword_filter.value = ""
        self.sender_filter.value = ""
        self.attachment_name_filter.value = ""
        self.check_download.value = True
        self.check_attachments_only.value = False
        self.check_open_folder.value = True
        self.check_xlsx.value = True
        self.check_pdf.value = False
        self.check_docs.value = False
        self.check_subfolders.value = False
        # Restablecer botones y estado
        self.btn_extract.disabled = True
        self.btn_cancel.disabled = True
        self.progress_bar.value = 0
        self.progress_bar.visible = False
        self.audit_table.rows.clear()
        self.status_text.value = "Listo"
        self._extraction_running = False
        self.page.update()

    def on_close(self, e):
        """Cierra limpio: cancela extracción en vuelo para no dejar hilos COM."""
        try:
            self.engine.stop_extraction()
        except Exception:
            pass

    def _toggle_theme(self, e):
        """Cambia entre modo oscuro y claro."""
        is_light = self.theme_toggle.value
        if is_light:
            self.page.theme_mode = ft.ThemeMode.LIGHT
            self.theme_toggle.label = "Modo Oscuro"
            _set_theme_mode("light")
            self.page.bgcolor = _get_theme_color("bg_light")
            self.main_content.bgcolor = _get_theme_color("bg_light")
        else:
            self.page.theme_mode = ft.ThemeMode.DARK
            self.theme_toggle.label = "Modo Claro"
            _set_theme_mode("dark")
            self.page.bgcolor = _get_theme_color("bg")
            self.main_content.bgcolor = _get_theme_color("bg")
        # Actualizar colores del nav_rail (sidebar) directamente
        self._refresh_nav_rail_colors()
        # Reconstruir vista principal para colores de contenedores
        self._rebuild_active_view()
        self._safe_update()

    def _refresh_nav_rail_colors(self):
        """Actualiza colores del nav_rail."""
        t = G360OutlookApp._t
        self.nav_rail.bgcolor = t("surface")
        for ctrl in self.nav_rail.leading.controls or []:
            if isinstance(ctrl, ft.Icon):
                ctrl.color = t("accent")
        # Mantener label correcto del toggle según estado actual
        if hasattr(self, "theme_toggle"):
            self.theme_toggle.label = "Modo Oscuro" if self.theme_toggle.value else "Modo Claro"

    @staticmethod
    def _t(key: str) -> str:
        """Retorna el color del tema activo (dark por defecto, light si se solicitó)."""
        return _get_theme_color(key)

    def _rebuild_active_view(self):
        """Reconstruye la vista actual aplicando el theme activo.

        Preserva los valores de los controles (TextField, Dropdown, Checkbox)
        recreándolos con sus valores actuales.
        """
        idx = self.nav_rail.selected_index
        if idx == 1:
            # Guardar estado de la vista de extracción antes de reconstruir
            saved = {
                "folder_value": getattr(self, "folder_dropdown", None) and self.folder_dropdown.value,
                "start_date": getattr(self, "btn_start_date", None) and self.btn_start_date.value,
                "end_date": getattr(self, "btn_end_date", None) and self.btn_end_date.value,
                "keyword": getattr(self, "keyword_filter", None) and self.keyword_filter.value,
                "sender": getattr(self, "sender_filter", None) and self.sender_filter.value,
                "check_download": getattr(self, "check_download", None) and self.check_download.value,
                "check_attachments_only": getattr(self, "check_attachments_only", None)
                and self.check_attachments_only.value,
                "check_open_folder": getattr(self, "check_open_folder", None) and self.check_open_folder.value,
                "check_xlsx": getattr(self, "check_xlsx", None) and self.check_xlsx.value,
                "check_pdf": getattr(self, "check_pdf", None) and self.check_pdf.value,
                "check_docs": getattr(self, "check_docs", None) and self.check_docs.value,
                "check_subfolders": getattr(self, "check_subfolders", None) and self.check_subfolders.value,
                "attachment_name": getattr(self, "attachment_name_filter", None) and self.attachment_name_filter.value,
                "btn_extract_disabled": self.btn_extract.disabled,
                "btn_cancel_disabled": self.btn_cancel.disabled,
                "status": self.status_text.value,
            }
            self.show_extraction_view()
            # Restaurar valores
            if saved["folder_value"]:
                self.folder_dropdown.value = saved["folder_value"]
            if saved["start_date"]:
                self.btn_start_date.value = saved["start_date"]
            if saved["end_date"]:
                self.btn_end_date.value = saved["end_date"]
            if saved["keyword"]:
                self.keyword_filter.value = saved["keyword"]
            if saved["sender"]:
                self.sender_filter.value = saved["sender"]
            if not saved.get("check_download", True):
                self.check_download.value = False
            if saved.get("check_attachments_only"):
                self.check_attachments_only.value = True
            if not saved.get("check_open_folder", True):
                self.check_open_folder.value = False
            if not saved.get("check_xlsx", True):
                self.check_xlsx.value = False
            if saved.get("check_pdf"):
                self.check_pdf.value = True
            if saved.get("check_docs"):
                self.check_docs.value = True
            if saved.get("check_subfolders"):
                self.check_subfolders.value = True
            if saved.get("attachment_name"):
                self.attachment_name_filter.value = saved["attachment_name"]
            if saved.get("btn_extract_disabled") is not None:
                self.btn_extract.disabled = saved["btn_extract_disabled"]
            if saved.get("btn_cancel_disabled") is not None:
                self.btn_cancel.disabled = saved["btn_cancel_disabled"]
            self.status_text.value = saved["status"]
        elif idx == 0:
            self.show_home_view()
        elif idx == 2:
            self.show_audit_log_view()

    def _on_start_date_pick(self, e):
        """Actualiza el campo de fecha inicio cuando se selecciona en el DatePicker."""
        selected = e.control.value
        if selected:
            self.btn_start_date.value = selected.strftime("%d/%m/%Y")
            self.btn_start_date.error_text = None
            self._safe_update()

    def _on_end_date_pick(self, e):
        """Actualiza el campo de fecha fin cuando se selecciona en el DatePicker."""
        selected = e.control.value
        if selected:
            self.btn_end_date.value = selected.strftime("%d/%m/%Y")
            self.btn_end_date.error_text = None
            self._safe_update()

    def _parse_fecha(self, texto: str):
        """Parsea una fecha en formato dd/mm/yyyy a objeto date."""
        if not texto or not texto.strip():
            return None
        try:
            return datetime.strptime(texto.strip(), "%d/%m/%Y").date()
        except ValueError:
            return None

    def handle_progress(self, progress, status):
        self.progress_bar.value = progress
        self.status_text.value = status
        self._safe_update()

    def on_item_processed(self, item):
        self.audit_table.rows.insert(
            0,
            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(str(item.get("received", ""))[:10], size=11)),
                    ft.DataCell(ft.Text(str(item.get("sender", ""))[:25], size=11, overflow=ft.TextOverflow.ELLIPSIS)),
                    ft.DataCell(ft.Text(str(item.get("subject", ""))[:35], size=11, overflow=ft.TextOverflow.ELLIPSIS)),
                ]
            ),
        )
        if len(self.audit_table.rows) > 30:
            self.audit_table.rows.pop()
        self._safe_update()

    def on_complete(self, data, save_path, cancelled=False, error=None, attachments_saved=0):
        self._extraction_running = False
        if error:
            self.status_text.value = f"Error: {error}"
        elif cancelled:
            self.status_text.value = f"Extracción cancelada. {len(data)} correos procesados."
        else:
            if attachments_saved > 0:
                total = len(data) + attachments_saved
                self.status_text.value = (
                    f"Finalizado: {len(data)} correos + {attachments_saved} adjuntos ({total} archivos en total)."
                )
            else:
                self.status_text.value = f"Finalizado: {len(data)} correos procesados correctamente."
        self.btn_extract.disabled = False
        self.btn_cancel.disabled = True
        self.progress_bar.value = 1.0 if not error else 0
        if not cancelled and not error and self.check_open_folder.value and save_path and os.path.exists(save_path):
            try:
                os.startfile(save_path)
            except Exception as ex:
                self.status_text.value = f"Error al abrir carpeta: {str(ex)}"
        self._safe_update()


def main(page: ft.Page):
    G360OutlookApp(page)


if __name__ == "__main__":
    ft.run(main)
