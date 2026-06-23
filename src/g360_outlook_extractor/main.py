import flet as ft
import os
from datetime import datetime
import threading
try:
    from .engine import OutlookEngine
except ImportError:
    from engine import OutlookEngine

try:
    from g360_flet.g360_signature import G360Signature
except ImportError:
    G360Signature = None

G360_THEME = {
    "bg": "#0b1220",
    "surface": "#151e2e",
    "accent": "#00d084",
    "text": "#e1e1e1",
    "text_dim": "#94a3b8",
    "border": "#1e293b",
    "danger": "#f43f5e",
    "warning": "#fbbf24"
}
class G360OutlookApp:
    def __init__(self, page: ft.Page):
        self.page = page
        self.engine = OutlookEngine()
        self.start_date = None
        self.end_date = None
        self.all_folders = []
        self.global_status = ft.Text("Listo", color=G360_THEME["text_dim"])
        self.setup_page()
        self.init_ui()

    def setup_page(self):
        self.page.title = "G360 | Outlook Extractor"
        self.page.bgcolor = G360_THEME["bg"]
        self.page.padding = 0
        self.page.window_width = 1200
        self.page.window_height = 800
        self.page.theme_mode = ft.ThemeMode.DARK
        
        # Date pickers
        self.start_date_picker = ft.DatePicker(
            on_change=self.on_start_date_change,
            first_date=datetime(2020, 1, 1),
            last_date=datetime(2030, 12, 31),
        )
        self.end_date_picker = ft.DatePicker(
            on_change=self.on_end_date_change,
            first_date=datetime(2020, 1, 1),
            last_date=datetime(2030, 12, 31),
        )
        self.page.overlay.extend([self.start_date_picker, self.end_date_picker])

        self.audit_rows = []
        self.audit_table = ft.DataTable(
            columns=[
                ft.DataColumn(ft.Text("Fecha", weight="bold")),
                ft.DataColumn(ft.Text("Remitente", weight="bold")),
                ft.DataColumn(ft.Text("Asunto", weight="bold")),
            ],
            rows=self.audit_rows,
            heading_row_color=G360_THEME["surface"],
        )
        self.global_status = ft.Text("Listo", color=G360_THEME["text_dim"])

    def init_ui(self):
        # Sidebar Lateral - Corrección de visibilidad y sintaxis
        self.nav_rail = ft.NavigationRail(
            selected_index=1,
            label_type=ft.NavigationRailLabelType.ALL,
            min_width=100,
            min_extended_width=200,
            bgcolor=G360_THEME["surface"],
            leading=ft.Container(
                content=ft.Icon(ft.Icons.MAIL_OUTLINE, color=G360_THEME["accent"], size=40),
                padding=ft.padding.symmetric(vertical=30),
            ),
            trailing=G360Signature(mode="powered", version="2.0") if G360Signature
            else ft.Text("Powered by G360", color=G360_THEME["text_dim"], size=10),
            destinations=[
                ft.NavigationRailDestination(icon=ft.Icons.HOME_OUTLINED, label="Inicio"),
                ft.NavigationRailDestination(icon=ft.Icons.STORAGE_OUTLINED, label="Extracción"),
                ft.NavigationRailDestination(icon=ft.Icons.CHECK_CIRCLE_OUTLINE, label="Auditoría"),
                ft.NavigationRailDestination(icon=ft.Icons.SETTINGS_OUTLINED, label="Ajustes"),
            ],
            on_change=self.on_nav_change,
        )

        self.main_content = ft.Container(
            expand=True,
            padding=30,
            bgcolor=G360_THEME["bg"],
        )

        self.page.add(
            ft.Row(
                [
                    self.nav_rail,
                    ft.VerticalDivider(width=1, thickness=1, color=G360_THEME["border"]),
                    self.main_content,
                ],
                expand=True,
                spacing=0,
            )
        )

        # Cargamos la vista de extracción por defecto
        self.show_extraction_view()
        self.page.update()

    def on_nav_change(self, e):
        index = e.control.selected_index
        if index == 0:
            self.show_home_view()
        elif index == 1:
            self.show_extraction_view()
        elif index == 2:
            self.show_audit_log_view()
        elif index == 3:
            self.show_settings_view()
        self.page.update()

    def show_home_view(self):
        self.main_content.content = ft.Column(
            expand=True,
            controls=[
                ft.Row(
                    [
                        ft.Text("Outlook Extractor", size=32, weight=ft.FontWeight.BOLD, color=G360_THEME["text"]),
                        ft.Container(
                            content=ft.Text("v1.0.0", size=12, color=G360_THEME["accent"]),
                            border=ft.Border.all(1, G360_THEME["accent"]),
                            border_radius=5,
                            padding=ft.Padding(8, 2, 8, 2),
                        ),
                    ],
                    alignment=ft.MainAxisAlignment.SPACE_BETWEEN,
                ),
                ft.Text("Herramienta de Auditoría y Extracción de Datos MAPI para el ecosistema G360", color=G360_THEME["text_dim"]),
                ft.Divider(height=40, color=G360_THEME["border"]),
                ft.Container(
                    expand=True,
                    content=ft.Text("Bienvenido. Seleccione una opción del panel lateral para gestionar la auditoría de correos.", italic=True, color=G360_THEME["text_dim"]),
                    alignment=ft.Alignment(0, 0),
                ),
                self.global_status,
            ],
        )
        self.main_content.update()

    def show_audit_log_view(self):
        log_content = "No se encontró el archivo de auditoría."
        if os.path.exists("audit.log"):
            try:
                with open("audit.log", "r", encoding="utf-8") as f:
                    lines = f.readlines()
                    log_content = "".join(lines[-50:])
            except Exception as e:
                log_content = f"Error al leer logs: {str(e)}"

        self.main_content.content = ft.Column(
            expand=True,
            scroll=ft.ScrollMode.ALWAYS,
            controls=[
                ft.Text("Auditoría de Sistema", size=32, weight=ft.FontWeight.BOLD, color=G360_THEME["text"]),
                ft.Text("Historial técnico de operaciones y errores (audit.log)", color=G360_THEME["text_dim"]),
                ft.Divider(height=40, color=G360_THEME["border"]),
                ft.Container(
                    content=ft.Text(log_content, font_family="Consolas", size=12, color=G360_THEME["accent"]),
                    bgcolor=G360_THEME["bg"],
                    padding=20,
                    border_radius=10,
                    border=ft.Border.all(1, G360_THEME["border"]),
                ),
                ft.ElevatedButton("Actualizar Logs", icon=ft.Icons.REFRESH, on_click=lambda _: self.show_audit_log_view()),
            ]
        )
        self.main_content.update()

    def show_extraction_view(self):
        self.folder_dropdown = ft.Dropdown(
            label="Seleccionar Carpeta de Outlook",
            hint_text="Cargue carpetas o use el explorador",
            border_color=G360_THEME["border"],
            color=G360_THEME["text"],
            expand=True,
        )
        self.btn_connect = ft.ElevatedButton(
            content=ft.Text("Listar Carpetas", color=G360_THEME["bg"]),
            icon=ft.Icons.FOLDER_OPEN,
            style=ft.ButtonStyle(color=G360_THEME["bg"], bgcolor=G360_THEME["text_dim"]),
            on_click=self.on_load_folders,
        )
        self.btn_pick_folder = ft.ElevatedButton(
            content=ft.Text("Explorar Carpetas (Outlook)", color=G360_THEME["bg"]),
            icon=ft.Icons.FOLDER_SPECIAL,
            style=ft.ButtonStyle(color=G360_THEME["bg"], bgcolor=G360_THEME["accent"]),
            on_click=self.on_pick_outlook_folder,
        )

        self.btn_extract = ft.ElevatedButton(
            content=ft.Text("Iniciar Extracción", color=G360_THEME["bg"]),
            icon=ft.Icons.PLAY_ARROW_ROUNDED,
            style=ft.ButtonStyle(color=G360_THEME["bg"], bgcolor=G360_THEME["accent"]),
            on_click=self.on_start_extraction,
            disabled=True,
        )
        self.btn_cancel = ft.ElevatedButton(
            content=ft.Text("Cancelar", color=ft.Colors.WHITE),
            icon=ft.Icons.CANCEL,
            style=ft.ButtonStyle(
                bgcolor=G360_THEME["danger"]
            ),
            on_click=self.on_cancel_extraction,
            disabled=True,
        )
        self.progress_bar = ft.ProgressBar(
            width=None,
            height=10,
            color=G360_THEME["accent"],
            bgcolor=G360_THEME["border"],
            value=0,
            visible=False,
        )
        self.status_text = ft.Text("", color=G360_THEME["text_dim"])
        self.btn_start_date = ft.TextButton(
            "Fecha Inicio: No seleccionada",
            icon="calendar_month_rounded",
            on_click=lambda _: self._open_date_picker(self.start_date_picker),
        )
        self.btn_end_date = ft.TextButton(
            "Fecha Fin: No seleccionada",
            icon="calendar_month_rounded",
            on_click=lambda _: self._open_date_picker(self.end_date_picker),
        )
        self.check_download = ft.Checkbox(
            label="Descargar Adjuntos y Cuerpo (.txt) al Escritorio",
            value=True,
            fill_color=G360_THEME["accent"],
        )
        self.check_open_folder = ft.Checkbox(
            label="Abrir carpeta al finalizar",
            value=True,
            fill_color=G360_THEME["accent"],
        )
        self.search_folder = ft.TextField(
            label="Filtrar carpetas...",
            prefix_icon=ft.Icons.SEARCH,
            on_change=self.filter_dropdown_options,
            border_color=G360_THEME["border"],
            height=45,
            expand=True,
        )
        self.keyword_filter = ft.TextField(
            label="Palabra clave (Asunto/Cuerpo)",
            border_color=G360_THEME["border"],
            expand=True,
        )

        # Construcción de la interfaz - Módulo de Extracción
        self.main_content.content = ft.Column(
            expand=True,
            scroll=ft.ScrollMode.AUTO,
            controls=[
                ft.Text("Módulo de Extracción", size=32, weight=ft.FontWeight.BOLD, color=G360_THEME["text"]),
                ft.Text("Seleccione una cuenta y carpeta para iniciar el proceso de auditoría.", color=G360_THEME["text_dim"]),
                ft.Divider(height=40, color=G360_THEME["border"]),
                
                ft.Container(
                    content=ft.Column([
                        ft.Text("Origen de Datos (MAPI)", weight="bold", size=16, color=G360_THEME["accent"]),
                        ft.Row([self.btn_pick_folder, self.btn_connect], spacing=10),
                        self.search_folder,
                        self.folder_dropdown,
                    ], spacing=15),
                    padding=20,
                    bgcolor=G360_THEME["surface"],
                    border_radius=10,
                    border=ft.Border.all(1, G360_THEME["border"]),
                ),
                ft.Container(height=5),
                ft.Container(
                    content=ft.Column([
                        ft.Text("Parámetros de Auditoría", weight="bold", size=16, color=G360_THEME["accent"]),
                        ft.Row([self.btn_start_date, self.btn_end_date], spacing=20),
                        self.keyword_filter,
                        ft.Row([self.check_download, self.check_open_folder]),
                        ft.Row([self.btn_extract, self.btn_cancel], spacing=20),
                    ], spacing=15),
                    padding=20,
                    bgcolor=G360_THEME["surface"],
                    border_radius=10,
                    border=ft.Border.all(1, G360_THEME["border"]),
                ),
                
                ft.Container(height=20),
                self.status_text,
                self.progress_bar,
                
                ft.Text("Auditoría de Procesamiento (Tiempo Real)", size=14, weight="bold", color=G360_THEME["accent"]),
                ft.Container(
                    content=self.audit_table,
                    height=380, # Altura fija para que el scroll de la columna funcione bien
                    border=ft.Border.all(1, G360_THEME["border"]),
                    border_radius=10,
                    padding=10,
                    bgcolor=G360_THEME["surface"],
                ),
            ]
        )
        self.main_content.update()
        self.page.update()

    def show_settings_view(self):
        self.main_content.content = ft.Column(
            expand=True,
            controls=[
                ft.Text("Ajustes", size=32, weight=ft.FontWeight.BOLD, color=G360_THEME["text"]),
                ft.Text("Configuración de la aplicación", color=G360_THEME["text_dim"]),
                ft.Divider(height=40, color=G360_THEME["border"]),
                ft.Text("Próximamente...", color=G360_THEME["accent"]),
            ]
        )
        self.main_content.update()

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
                            text=f"{folder['account']} > {folder['folder_name']}"
                        )
                    ]
                    self.folder_dropdown.value = f"{folder['account']}|{folder['folder_name']}"
                    self.btn_extract.disabled = False
                    self.status_text.value = f"Seleccionado: {folder['folder_name']}"
                else:
                    self.status_text.value = "Selección cancelada."
            except Exception as ex:
                self.status_text.value = f"Error: {str(ex)}"
            self.page.update()
        
        threading.Thread(target=pick, daemon=True).start()

    def on_load_folders(self, e):
        self.btn_connect.disabled = True
        self.status_text.value = "Cargando carpetas MAPI..."
        self.page.update()
        
        def load():
            try:
                self.all_folders = self.engine.get_available_folders()
                self.folder_dropdown.options = [
                    ft.dropdown.Option(key=f"{f['account']}|{f['folder_name']}", text=f"{f['account']} > {f['folder_name']} ({f['count']} elementos)")
                    for f in self.all_folders
                ]
                self.folder_dropdown.hint_text = "Seleccione una carpeta"
                if self.all_folders:
                    self.btn_extract.disabled = False
                self.status_text.value = f"Listo: {len(self.all_folders)} carpetas."
            except Exception as ex:
                self.status_text.value = f"Error: {str(ex)}"
            self.btn_connect.disabled = False
            self.page.update()
        
        threading.Thread(target=load, daemon=True).start()

    def filter_dropdown_options(self, e):
        search_term = self.search_folder.value.lower()
        filtered = [
            f for f in self.all_folders
            if search_term in f['folder_name'].lower() or search_term in f['account'].lower()
        ]
        self.folder_dropdown.options = [
            ft.dropdown.Option(key=f"{f['account']}|{f['folder_name']}", text=f"{f['account']} > {f['folder_name']} ({f['count']} elementos)")
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
            acc, fol = self.folder_dropdown.value.split("|")
        except (ValueError, AttributeError):
            self.status_text.value = "Error en la selección de carpeta"
            self.page.update()
            return
        
        self.btn_extract.disabled = True
        self.btn_cancel.disabled = False
        self.progress_bar.visible = True
        self.audit_rows.clear()
        self.audit_table.rows.clear()
        filters = {
            "keyword": self.keyword_filter.value,
            "start_date": self.start_date,
            "end_date": self.end_date,
        }
        self.engine.start_extraction(
            acc,
            fol,
            filters,
            self.check_download.value,
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

    def _open_date_picker(self, picker):
        """Abre un DatePicker de forma compatible con la versión de Flet."""
        picker.open = True
        self.page.update()

    def on_start_date_change(self, e):
        self.start_date = self.start_date_picker.value
        if self.start_date:
            self.btn_start_date.text = f"Inicio: {self.start_date.strftime('%d/%m/%Y')}"
        self.page.update()

    def on_end_date_change(self, e):
        self.end_date = self.end_date_picker.value
        if self.end_date:
            self.btn_end_date.text = f"Fin: {self.end_date.strftime('%d/%m/%Y')}"
        self.page.update()

    def handle_progress(self, progress, status):
        self.progress_bar.value = progress
        self.status_text.value = status
        self.page.update()

    def on_item_processed(self, item):
        self.audit_table.rows.insert(
            0,
            ft.DataRow(
                cells=[
                    ft.DataCell(ft.Text(item["received"][:10], size=11)),
                    ft.DataCell(ft.Text(item["sender"][:25], size=11, overflow=ft.TextOverflow.ELLIPSIS)),
                    ft.DataCell(ft.Text(item["subject"][:35], size=11, overflow=ft.TextOverflow.ELLIPSIS)),
                ]
            ),
        )
        if len(self.audit_table.rows) > 30:
            self.audit_table.rows.pop()
        self.page.update()

    def on_complete(self, data, save_path):
        if "cancelada" not in self.status_text.value and "Error" not in self.status_text.value:
            self.status_text.value = f"Finalizado: {len(data)} elementos procesados correctamente."
        self.btn_extract.disabled = False
        self.btn_cancel.disabled = True
        if self.check_open_folder.value and save_path and os.path.exists(save_path):
            try:
                os.startfile(save_path)
            except Exception as ex:
                self.status_text.value = f"Error al abrir carpeta: {str(ex)}"
        self.page.update()


def main(page: ft.Page):
    G360OutlookApp(page)


if __name__ == "__main__":
    ft.app(target=main)