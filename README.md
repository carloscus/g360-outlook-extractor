# G360 | Outlook Extractor

> Herramienta profesional de auditoría y extracción de datos MAPI (Outlook). Forma parte de la familia de microherramientas G360 para apoyo CRM y gestión de datos en escritorio.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## Tabla de Contenidos

- [Descripción](#descripción)
- [Características](#características-principales)
- [Requisitos](#requisitos-técnico)
- [Instalación](#instalación-y-desarrollo)
- [Estructura de Salida](#estructura-de-salida)
- [Notas Técnicas](#notas-de-mantenimiento)
- [Contribución](#contribución)
- [Licencia](#licencia)
- [Familia G360](#familia-g360)

---

## 🚀 Características Principales

- **Extracción Inteligente**: Filtrado por rango de fechas (MAPI Restrict + fallback Python), palabras clave, remitente y nombre de adjunto.
- **Subcarpetas**: Opción para procesar recursivamente todas las subcarpetas de una cuenta.
- **Gestión de Adjuntos**: Descarga automática con renombrado inteligente (prefix temporal + código OC si aplica).
- **Filtro por Extensión**: Selecciona qué tipos de archivos descargar (Excel, PDF, Docs).
- **Modo Solo Adjuntos**: Opción para descargar solo archivos adjuntos sin el cuerpo del correo.
- **Preservación de Contexto**: Generación de archivos `.txt` vinculados con metadatos del correo.
- **Tema Claro/Oscuro**: UI adaptable con paleta G360 oficial (flet-desktop / flet-desktop-corporativo).
- **Arquitectura Anti-Bloqueo**: Ejecución en hilos separados (STA) para mantener la interfaz Flet fluida.
- **Sistema de Auditoría**: Registro detallado de operaciones en `audit.log` y visualización en tiempo real en la UI.
- **Error Handling Traducido**: Códigos COM traducidos a mensajes legibles (error 9, 15, 16).

## 🎨 Paleta de Colores (G360 oficial)

| Token | Oscuro (flet-desktop) | Claro (flet-desktop-corporativo) |
|---|---|---|
| bg | `#0b1220` | `#ffffff` |
| surface | `#1a2332` | `#f8fafc` |
| accent | `#00d084` | `#00796B` |
| text | `#f0f4f8` | `#1e293b` |
| muted | `#94a3b8` | `#64748b` |
| border | `#5a6a80` | `#64748b` |
| danger | `#b91c1c` | `#b91c1c` |
| warning | `#f59e0b` | `#f59e0b` |
| success | `#22c55e` | `#10b981` |
| info | `#60a5fa` | `#00796B` |

## 🛠️ Requisitos Técnico

- **SO**: Windows (Requerido para soporte COM/MAPI).
- **Outlook**: Debe estar instalado y con una sesión activa.
- **Python**: >= 3.11.
- **Gestor de Paquetes**: uv.

## 📦 Instalación y Desarrollo

1. **Configurar el entorno**:
   ```bash
   uv sync
   ```

2. **Ejecutar la aplicación**:
   ```bash
   uv run main.py
   ```

3. **Construir ejecutable portátil**:
   ```bash
   build.bat
   ```

4. **Instalar dependencias de desarrollo** (opcional):
   ```bash
   uv sync --extra dev
   ```

## 📂 Estructura de Salida

La herramienta crea una estructura organizada en el Escritorio:
```text
Desktop/G360_Extraidos/
└── YYYY-MM-DD[_hasta_YYYYMMDD]/
    ├── 20241019_1430_OC_12345.pdf
    ├── 20241019_1430_OC_12345_CUERPO.txt  <-- Contexto del correo + carpeta origen
    └── ...
```

## 🛡️ Notas de Mantenimiento

- **Threading**: Outlook usa *Single Threaded Apartment* (STA). Cada hilo que interactúa con Outlook debe invocar `pythoncom.CoInitialize()`.
- **Zombies COM**: El motor está diseñado para liberar referencias y llamar a `CoUninitialize()` para evitar que procesos `OUTLOOK.EXE` queden huérfanos.
- **Logs**: Los logs detallados van a `audit.log` en la raíz del proyecto (nivel INFO).
- **Cancelación**: La extracción puede cancelarse en cualquier momento. El motor verifica la señal de cancelación en cada iteración.
- **Subcarpetas**: Con "Incluir subcarpetas" marcado, se procesan recursivamente todas las subcarpetas de la seleccionada.
- **Filtro de fecha**: Se intenta MAPI Restrict primero; si falla, se aplica en Python automáticamente.
- **Error 9/15/16**: Manejados con mensajes traducidos. Si aparece, cierra Outlook y reintenta con un rango de fecha más corto.

## Contribución

1. Fork el repositorio
2. Crea una rama
3. Commit tus cambios
4. Push a la rama
5. Abre un Pull Request

## Licencia

MIT License - ver [LICENSE](LICENSE) para más detalles.

---

## Familia G360

Este proyecto forma parte de la familia de microherramientas **G360** para apoyo CRM y gestión de datos en escritorio, enfocadas en áreas como ventas, finanzas y logística.

### Herramientas Relacionadas

- **[g360-cli](https://github.com/carloscus/g360-cli)**: Bootstrap de proyectos G360
- **[g360-signature](https://github.com/carloscus/g360-signature)**: Web component de branding G360
- **[g360-nc-sustentor](https://github.com/carloscus/g360-nc-sustentor)**: Generador de cuadros de sustento para Notas de Crédito

---

## 🔧 Cambios Recientes (v1.1.0)

- **Agregado**: Checkbox "Incluir subcarpetas" para procesamiento recursivo
- **Agregado**: Campo "Filtrar adjuntos por nombre" (búsqueda parcial case-insensitive)
- **Agregado**: Checkbox "Solo adjuntos (sin cuerpo)"
- **Agregado**: Extractor de código OC de nombres de archivos para deduplicación
- **Agregado**: Mensajes de error COM traducidos (9, 15, 16)
- **Mejorado**: Filtro de fecha ahora usa MAPI Restrict primero, luego Python
- **Mejorado**: Paleta de colores actualizada a tokens oficiales G360 (flet-desktop)
- **Mejorado**: Tema claro/oscuro funcional con preservación de estado
- **Corregido**: Sidebar ahora actualiza colores al cambiar tema
- **Corregido**: Firma G360 posicionada correctamente en el footer del nav rail
- **Corregido**: Ventana ajustada a 600×700px con sidebar de 90px
- **Corregido**: Botón "Reiniciar" para limpiar estado entre extracciones

---
**Marca**: G360
**Isotipo**: 3 puntos verticales paralelos (gris-verde-gris) + chevron `>`
**Autor**: Carlos Cusi
**Desarrollo**: Con asistencia de herramientas de código IA (Vibe Code)
**Powered by**: [g360-signature](https://github.com/carloscus/g360-signature)
