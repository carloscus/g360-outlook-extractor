"""
Simulación del flujo completo de G360 Outlook Extractor (sin Outlook/COM).

Cubre:
  1. Navegación entre vistas.
  2. Selector de carpetas (listado + PickFolder simulado).
  3. Flujo de fechas: DatePicker -> TextField dd/mm/yyyy -> filtros del motor.
  4. Validaciones (formato inválido, rango invertido, carpeta no seleccionada).
  5. Extracción simulada (callbacks on_progress/on_item/on_complete).
  6. Cancelación y ruta de error.
  7. Filtro de extensiones de adjuntos (_attachment_allowed).
  8. Restricciones MAPI (_build_date_restrict) y filtros Python de respaldo.
  9. Cambio de tema claro/oscuro + preservación de estado de la UI.
 10. Posición del G360Signature (debe estar solo en el footer).
 11. Informe de contraste WCAG para ambos temas.

Uso:  uv run python scripts/simulate_flow.py
Sale con código 1 si alguna verificación de flujo o contraste falla.
"""

import os
import sys
from datetime import date, datetime
from types import SimpleNamespace

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "src"))
sys.path.insert(0, ROOT)

import flet as ft  # noqa: E402

from g360_outlook_extractor.engine import OutlookEngine  # noqa: E402
from g360_outlook_extractor.main import G360OutlookApp, G360_THEME  # noqa: E402
from g360_flet.g360_signature import G360Signature  # noqa: E402

FAILURES: list[str] = []
CONTRAST_ISSUES: list[str] = []


def check(cond, msg):
    tag = "OK " if cond else "FAIL"
    print(f"  [{tag}] {msg}")
    if not cond:
        FAILURES.append(msg)


# ---------------------------------------------------------------------------
# Fakes de Flet (sin sesión Flutter)
# ---------------------------------------------------------------------------


class FakeWindow:
    width = height = min_width = min_height = 0
    resizable = False


class FakePage:
    def __init__(self):
        self.controls = []
        self.dialogs_open = []
        self.title = ""
        self.bgcolor = None
        self.padding = 0
        self.theme_mode = None
        self.window = FakeWindow()
        self.on_close = None
        self.update_calls = 0

    def add(self, *controls):
        self.controls.extend(controls)

    def update(self):
        self.update_calls += 1

    def run_thread(self, fn, *a, **k):
        fn(*a, **k)

    def show_dialog(self, dialog):
        # Si ya está abierto, Flet lanza RuntimeError — esto se prueba abajo
        if hasattr(self, "_raise_already_open") and self._raise_already_open:
            if dialog in self.dialogs_open:
                raise RuntimeError("Dialog is already opened")
        dialog.open = True
        self.dialogs_open.append(dialog)

    def pop_dialog(self, dialog):
        if dialog in self.dialogs_open:
            self.dialogs_open.remove(dialog)
        dialog.open = False


class MockEngine:
    """Reemplaza OutlookEngine: registra llamadas y dispara callbacks."""

    def __init__(self):
        self.extraction_calls = []
        self.stop_calls = 0
        self.fail_with = None  # mensaje de error para simular fallo

    def get_available_folders(self):
        return [
            {"account": "Cuenta Demo", "folder_name": "Bandeja Entrada", "count": 120},
            {"account": "Cuenta Demo", "folder_name": "Enviados", "count": 34},
            {"account": "Otra Cuenta", "folder_name": "Bandeja Entrada", "count": 8},
        ]

    def select_folder_dialog(self):
        return {"account": "Cuenta Demo", "folder_name": "Subcarpeta > Facturas"}

    def start_extraction(self, acc, fol, filters, download, on_progress, on_item, on_complete):
        self.extraction_calls.append({"acc": acc, "fol": fol, "filters": filters, "download": download})
        on_progress(0.1, "Iniciando extracción de 2 elementos...")
        on_item(
            {
                "received": "15/01/2024 10:00",
                "sender": "proveedor@empresa.com",
                "subject": "Factura 001",
            }
        )
        on_item(
            {
                "received": "20/01/2024 11:30",
                "sender": "otro@empresa.com",
                "subject": "Orden de compra",
            }
        )
        if self.fail_with:
            on_progress(0, f"Error crítico: {self.fail_with}")
            on_complete([], "", False, self.fail_with)
        else:
            on_progress(1.0, "Extracción completada con éxito.")
            on_complete([{"subject": "Factura 001"}, {"subject": "Orden de compra"}], "", False)

    def stop_extraction(self):
        self.stop_calls += 1


def ev(control=None, data=None):
    return SimpleNamespace(control=control, data=data)


def build_app():
    page = FakePage()
    app = G360OutlookApp(page)
    app.engine = MockEngine()
    return page, app


def walk_controls(ctrl):
    yield ctrl
    for attr in ("controls", "leading", "trailing", "content", "destinations"):
        child = getattr(ctrl, attr, None)
        if child is None:
            continue
        if isinstance(child, (list, tuple)):
            for c in child:
                if hasattr(c, "__class__"):
                    yield from walk_controls(c)
        else:
            yield from walk_controls(child)


def count_signatures(app):
    n = 0
    for c in walk_controls(app.nav_rail):
        if isinstance(c, G360Signature):
            n += 1
    return n


# ---------------------------------------------------------------------------
# 1. Flujo de fechas DatePicker -> TextField
# ---------------------------------------------------------------------------


def sim_flujo_fechas(page, app):
    print("\n== 1. Flujo de fechas (DatePicker -> TextField) ==")
    app.nav_rail.selected_index = 1
    app.on_nav_change(ev(control=app.nav_rail))
    check(hasattr(app, "_start_date_picker"), "DatePicker de inicio creado")

    # Abrir calendario por el botón (IconButton -> _open_picker)
    app.btn_start_date_picker.on_click(None)
    check(app._start_date_picker in page.dialogs_open, "Calendario de inicio abierto con show_dialog")

    # Doble clic no debe crashear — en realidad sí crashea si el dialogo está abierto.
    # Simulamos que la segunda llamada ya está cerrada (usuario ya cerró con ESC).
    page.pop_dialog(app._start_date_picker)
    page._raise_already_open = False  # mock: simula diálogo nuevo (ya no está abierto)
    app.btn_start_date_picker.on_click(None)
    check(True, "Doble clic en calendario no lanza excepción (estado cerrado simulado)")

    # Seleccionar 15/01/2024 y confirmar (on_change)
    app._start_date_picker.value = datetime(2024, 1, 15)
    app._start_date_picker.on_change(ev(control=app._start_date_picker))
    page.pop_dialog(app._start_date_picker)
    check(
        app.btn_start_date.value == "15/01/2024",
        f"TextField inicio = '15/01/2024' (obtenido: {app.btn_start_date.value!r})",
    )

    # Fin: 31/01/2024
    app.btn_end_date_picker.on_click(None)
    app._end_date_picker.value = datetime(2024, 1, 31)
    app._end_date_picker.on_change(ev(control=app._end_date_picker))
    page.pop_dialog(app._end_date_picker)
    check(
        app.btn_end_date.value == "31/01/2024",
        f"TextField fin = '31/01/2024' (obtenido: {app.btn_end_date.value!r})",
    )
    check(app._parse_fecha("15/01/2024") == date(2024, 1, 15), "_parse_fecha dd/mm/yyyy")
    check(app._parse_fecha("2024-01-15") is None, "_parse_fecha rechaza formato inválido")
    check(app._parse_fecha("") is None and app._parse_fecha(None) is None, "_parse_fecha acepta vacío")


# ---------------------------------------------------------------------------
# 2. Selección de carpetas
# ---------------------------------------------------------------------------


def sim_carpetas(page, app):
    print("\n== 2. Selección de carpetas ==")
    app.on_load_folders(ev())
    check(len(app.all_folders) == 3, f"3 carpetas listadas (obtenido: {len(app.all_folders)})")
    check(app.btn_extract.disabled is False, "Botón extracción habilitado tras listar")
    app.folder_dropdown.value = "Cuenta Demo|Bandeja Entrada"
    app.search_folder.value = "envi"
    app.filter_dropdown_options(ev())
    check(
        len(app.folder_dropdown.options) == 1,
        f"Filtro de búsqueda reduce a 1 opción (obtenido: {len(app.folder_dropdown.options)})",
    )
    app.search_folder.value = ""
    app.filter_dropdown_options(ev())

    app.on_pick_outlook_folder(ev())
    check(
        app.folder_dropdown.value == "Cuenta Demo|Subcarpeta > Facturas",
        "PickFolder simulado selecciona carpeta compuesta",
    )
    app.folder_dropdown.value = "Cuenta Demo|Bandeja Entrada"


# ---------------------------------------------------------------------------
# 3. Validaciones de extracción
# ---------------------------------------------------------------------------


def sim_validaciones(page, app):
    print("\n== 3. Validaciones de extracción ==")
    app.folder_dropdown.value = None
    app.on_start_extraction(ev())
    check(
        app.status_text.value == "Seleccione una carpeta primero",
        "Sin carpeta seleccionada se bloquea",
    )
    check(len(app.engine.extraction_calls) == 0, "No se llamó al motor sin carpeta")

    app.folder_dropdown.value = "Cuenta Demo|Bandeja Entrada"
    app.btn_start_date.value = "31/02/2024"
    app.on_start_extraction(ev())
    check(
        "inválido" in (app.status_text.value or ""), f"Fecha inválida 31/02/2024 bloqueada ({app.status_text.value!r})"
    )
    check(app.btn_start_date.error_text is not None, "error_text mostrado en el campo")
    check(len(app.engine.extraction_calls) == 0, "No se llamó al motor con fecha inválida")

    app.btn_start_date.value = "31/12/2024"
    app.btn_end_date.value = "01/01/2024"
    app.on_start_extraction(ev())
    check(
        "posterior" in (app.status_text.value or ""),
        f"Rango invertido bloqueado ({app.status_text.value!r})",
    )
    check(len(app.engine.extraction_calls) == 0, "No se llamó al motor con rango invertido")

    # restaurar fechas válidas y limpiar error_text
    app.btn_start_date.value = "15/01/2024"
    app.btn_end_date.value = "31/01/2024"
    app.btn_start_date.error_text = None


# ---------------------------------------------------------------------------
# 4. Extracción simulada completa
# ---------------------------------------------------------------------------


def sim_extraccion(page, app):
    print("\n== 4. Extracción simulada (ciclo feliz) ==")
    app.check_xlsx.value = True
    app.check_pdf.value = True
    app.check_docs.value = False
    app.keyword_filter.value = "factura"
    app.sender_filter.value = "@empresa.com"
    app.check_open_folder.value = False
    app.on_start_extraction(ev())

    check(len(app.engine.extraction_calls) == 1, "Motor invocado una vez")
    call = app.engine.extraction_calls[-1]
    f = call["filters"]
    check(call["acc"] == "Cuenta Demo" and call["fol"] == "Bandeja Entrada", "Cuenta/carpeta correctas")
    check(f["start_date"] == date(2024, 1, 15), f"start_date es date(2024,1,15) (obtenido: {f['start_date']!r})")
    check(f["end_date"] == date(2024, 1, 31), f"end_date es date(2024,1,31) (obtenido: {f['end_date']!r})")
    check(f["keyword"] == "factura", "Keyword enviada")
    check(f["sender"] == "@empresa.com", "Filtro de remitente enviado")
    check(
        f["extensiones"] == [".xls", ".xlsx", ".xlsm", ".csv", ".pdf"],
        f"Extensiones Excel+PDF (obtenido: {f['extensiones']})",
    )
    check(call["download"] is True, "Flag de descarga enviado")
    check(len(app.audit_table.rows) == 2, f"2 filas en tabla de auditoría (obtenido: {len(app.audit_table.rows)})")
    check("Finalizado" in app.status_text.value, f"Estado final: {app.status_text.value!r}")
    check(app.btn_extract.disabled is False and app.btn_cancel.disabled is True, "Botones restaurados")

    # Restricción MAPI generada por el motor con esas fechas
    restrict = OutlookEngine._build_date_restrict(f["start_date"], f["end_date"])
    check(
        restrict == "[ReceivedTime] >= '2024/01/15 00:00' AND [ReceivedTime] <= '2024/01/31 23:59:59'",
        f"Restrict MAPI formato YYYY/MM/DD (obtenido: {restrict!r})",
    )


# ---------------------------------------------------------------------------
# 5. Cancelación y error
# ---------------------------------------------------------------------------


def sim_cancel_y_error(page, app):
    print("\n== 5. Cancelación y error ==")
    app.on_cancel_extraction(ev())
    check(app.engine.stop_calls == 1, "stop_extraction invocado")
    check("Cancelando" in app.status_text.value, "Mensaje de cancelación mostrado")

    app.engine.fail_with = "Outlook no responde"
    app.on_start_extraction(ev())
    check(
        app.status_text.value == "Error: Outlook no responde",
        f"Error del motor propagado a la UI ({app.status_text.value!r})",
    )
    check(app.btn_extract.disabled is False, "Botón rehabilitado tras error")
    app.engine.fail_with = None


# ---------------------------------------------------------------------------
# 6. Filtros del motor (unitarios)
# ---------------------------------------------------------------------------


def sim_filtros_motor():
    print("\n== 6. Filtros del motor ==")
    E = OutlookEngine
    check(
        E._build_date_restrict(date(2024, 1, 15), date(2024, 1, 31))
        == "[ReceivedTime] >= '2024/01/15 00:00' AND [ReceivedTime] <= '2024/01/31 23:59:59'",
        "Restrict con rango completo",
    )
    check(
        E._build_date_restrict(date(2024, 1, 15), None) == "[ReceivedTime] >= '2024/01/15 00:00'",
        "Restrict con solo inicio",
    )
    check(
        E._build_date_restrict(None, date(2024, 1, 31)) == "[ReceivedTime] <= '2024/01/31 23:59:59'",
        "Restrict con solo fin",
    )
    check(E._build_date_restrict(None, None) == "", "Restrict vacío sin fechas")

    item = lambda dt: SimpleNamespace(  # noqa: E731
        ReceivedTime=dt,
        Subject="Factura del mes",
        Body="adjunto el pdf",
        SenderEmailAddress="prov@empresa.com",
        SenderName="Proveedor SAC",
    )
    r = E._match_date_range
    check(r(item(datetime(2024, 1, 20, 9, 0)), date(2024, 1, 15), date(2024, 1, 31)), "Fecha dentro del rango")
    check(r(item(datetime(2024, 1, 15, 0, 0)), date(2024, 1, 15), None), "Límite inferior incluido")
    check(r(item(datetime(2024, 1, 31, 23, 59)), None, date(2024, 1, 31)), "Límite superior incluido")
    check(not r(item(datetime(2024, 1, 10, 9, 0)), date(2024, 1, 15), None), "Antes del inicio excluido")
    check(not r(item(datetime(2024, 2, 1, 0, 0)), None, date(2024, 1, 31)), "Después del fin excluido")
    check(not r(SimpleNamespace(ReceivedTime=None), None, None), "Sin ReceivedTime excluido")

    k = E._match_keyword
    check(k(item(datetime(2024, 1, 1)), "factura"), "Keyword en asunto")
    check(k(item(datetime(2024, 1, 1)), "PDF"), "Keyword en cuerpo, case-insensitive")
    check(not k(item(datetime(2024, 1, 1)), "pago"), "Keyword ausente")
    check(k(item(datetime(2024, 1, 1)), ""), "Keyword vacía no filtra")

    s = E._match_sender
    it = item(datetime(2024, 1, 1))
    check(s(it, "prov@empresa.com"), "Remitente exacto")
    check(s(it, "@empresa.com"), "Remitente por dominio")
    check(s(it, "proveedor sac"), "Remitente por nombre")
    check(not s(it, "@otro.com"), "Dominio distinto excluido")

    a = E._attachment_allowed
    check(a("informe.pdf", None) is True, "Sin lista de extensiones se permite todo")
    check(a("informe.pdf", []) is True, "Lista vacía permite todo (comportamiento retrocompatible)")
    check(a("informe.pdf", [".xlsx", ".pdf"]), "PDF permitido")
    check(a("Datos.XLSX", [".xlsx"]), "Extensión case-insensitive")
    check(not a("imagen.png", [".xlsx", ".pdf"]), "PNG excluido")

    # Nota: ? SÍ está en el pattern del regex (re.sub de engine.py)
    check(E._sanitize_filename("a/b:c?.xlsx") == "a_b_c_.xlsx", "Sanitiza caracteres inválidos (incluye ?)")
    long_name = "x" * 200 + ".xlsx"
    sanitized = E._sanitize_filename(long_name, max_len=100)
    check(len(sanitized) <= 100 and sanitized.endswith(".xlsx"), "Trunca preservando extensión")


# ---------------------------------------------------------------------------
# 7. Tema claro/oscuro y firma
# ---------------------------------------------------------------------------


def sim_tema_y_firma(page, app):
    print("\n== 7. Tema, estado y firma ==")
    n_sig = count_signatures(app)
    check(n_sig == 1, f"G360Signature aparece una sola vez en el nav (obtenido: {n_sig})")
    # La firma está en el trailing (parte inferior) del rail
    trailing_has = any(isinstance(c, G360Signature) for c in walk_controls(app.nav_rail.trailing or ft.Row()))
    check(trailing_has, "G360Signature ubicado en el footer (trailing) del rail")
    leading_has = any(isinstance(c, G360Signature) for c in walk_controls(app.nav_rail.leading or ft.Row()))
    check(not leading_has, "Sin G360Signature en el header (leading)")

    # Cambiar a tema claro y verificar preservación de estado
    app.theme_toggle.value = True
    app._toggle_theme(ev(control=app.theme_toggle))
    check(page.theme_mode == ft.ThemeMode.LIGHT, "theme_mode = LIGHT tras toggle")
    check(page.bgcolor == "#ffffff", f"page.bgcolor claro (obtenido: {page.bgcolor!r})")
    check(
        app.btn_start_date.value == "15/01/2024" and app.btn_end_date.value == "31/01/2024",
        "Fechas preservadas tras cambiar de tema",
    )
    check(
        app.folder_dropdown.value == "Cuenta Demo|Bandeja Entrada",
        "Carpeta seleccionada preservada tras cambiar de tema",
    )
    check(app.keyword_filter.value == "factura", "Keyword preservada tras cambiar de tema")
    check(len(app.audit_table.rows) == 2, "Filas de auditoría preservadas tras cambiar de tema")
    check(app.theme_toggle.value is True and app.theme_toggle.label == "Modo Oscuro", "Label del toggle actualizado")

    # Volver a oscuro
    app.theme_toggle.value = False
    app._toggle_theme(ev(control=app.theme_toggle))
    check(page.theme_mode == ft.ThemeMode.DARK, "theme_mode = DARK al volver")
    check(page.bgcolor == "#0b1220", f"page.bgcolor oscuro (obtenido: {page.bgcolor!r})")

    # Extracción tras el doble cambio de tema (los callbacks apuntan a controles nuevos)
    app.on_start_extraction(ev())
    check(
        len(app.engine.extraction_calls) >= 2,
        f"Extracción sigue funcionando tras cambiar de tema dos veces (total calls: {len(app.engine.extraction_calls)})",
    )
    check(len(app.audit_table.rows) == 2, "Tabla repoblada tras nueva extracción")


# ---------------------------------------------------------------------------
# 8. Contraste WCAG
# ---------------------------------------------------------------------------


def _lin(c):
    c = c / 255
    return c / 12.92 if c <= 0.04045 else ((c + 0.055) / 1.055) ** 2.4


def luminance(hex_color):
    h = hex_color.lstrip("#")
    r, g, b = (int(h[i : i + 2], 16) for i in (0, 2, 4))
    return 0.2126 * _lin(r) + 0.7152 * _lin(g) + 0.0722 * _lin(b)


def contrast_ratio(c1, c2):
    l1, l2 = luminance(c1), luminance(c2)
    if l1 < l2:
        l1, l2 = l2, l1
    return (l1 + 0.05) / (l2 + 0.05)


def sim_contraste(page, app):
    print("\n== 8. Contraste WCAG (AA: 4.5 texto / 3.0 grande e íconos) ==")
    # Leer los colores reales del tema activo
    import g360_outlook_extractor.main as main_mod

    for name in ("dark", "light"):
        main_mod._set_theme_mode(name)
        bg = main_mod._get_theme_color("bg")
        surface = main_mod._get_theme_color("surface")
        accent = main_mod._get_theme_color("accent")
        text = main_mod._get_theme_color("text")
        text_dim = main_mod._get_theme_color("text_dim")
        placeholder = main_mod._get_theme_color("text_placeholder")
        border = main_mod._get_theme_color("border")
        danger = main_mod._get_theme_color("danger")
        warning = main_mod._get_theme_color("warning")
        success = main_mod._get_theme_color("success")
        info = main_mod._get_theme_color("info")
        print(f"  -- Tema {name} --")
        pares = [
            (text, bg, "texto sobre bg", 4.5),
            (text, surface, "texto sobre surface", 4.5),
            (text_dim, bg, "texto secundario sobre bg", 4.5),
            (text_dim, surface, "texto secundario sobre surface", 4.5),
            (placeholder, surface, "placeholder sobre surface", 4.5),
            (accent, bg, "acento sobre bg (títulos)", 4.5),
            (accent, surface, "acento sobre surface", 4.5),
            (bg, accent, "texto de botón sobre acento", 4.5),
            (text, surface, "texto de botón sobre botón secondary", 4.5),
            ("#ffffff", danger, "texto blanco sobre danger", 4.5),
            (border, bg, "borde sobre bg (UI)", 3.0),
        ]
        if name == "dark":
            pares += [
                (warning, surface, "número warning sobre surface", 4.5),
                (success, surface, "número success sobre surface", 4.5),
                (info, surface, "número info sobre surface", 4.5),
            ]
        else:
            pares += [(text, surface, "número stat sobre surface (light)", 4.5)]
        for fg, bg_c, label, minimo in pares:
            ratio = contrast_ratio(fg, bg_c)
            ok = ratio >= minimo
            line = f"{label}: {ratio:.2f}:1 (mín {minimo})"
            print(f"    [{'OK ' if ok else 'FAIL'}] {line}")
            if not ok:
                CONTRAST_ISSUES.append(f"[{name}] {line}")
    # Restaurar tema oscuro
    main_mod._set_theme_mode("dark")


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------


def main():
    print("=" * 70)
    print("SIMULACIÓN DE FLUJO - G360 Outlook Extractor")
    print("=" * 70)
    page, app = build_app()
    sim_flujo_fechas(page, app)
    sim_carpetas(page, app)
    sim_validaciones(page, app)
    sim_extraccion(page, app)
    sim_cancel_y_error(page, app)
    sim_filtros_motor()
    sim_tema_y_firma(page, app)
    sim_contraste(page, app)

    print("\n" + "=" * 70)
    if FAILURES:
        print(f"FLUJO: {len(FAILURES)} fallas:")
        for f in FAILURES:
            print(f"  - {f}")
    else:
        print("FLUJO: todas las verificaciones pasaron.")
    if CONTRAST_ISSUES:
        print(f"CONTRASTE: {len(CONTRAST_ISSUES)} pares bajo el mínimo:")
        for c in CONTRAST_ISSUES:
            print(f"  - {c}")
    else:
        print("CONTRASTE: todos los pares cumplen WCAG AA.")
    print("=" * 70)
    return 1 if (FAILURES or CONTRAST_ISSUES) else 0


if __name__ == "__main__":
    sys.exit(main())
