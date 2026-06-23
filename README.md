# G360 | Outlook Extractor

> Herramienta profesional de auditoría y extracción de datos MAPI (Outlook). Forma parte de la familia de microherramientas G360 para apoyo CRM y gestión de datos en escritorio.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

---

## Tabla de Contenidos

- [Descripción](#descripción)
- [Características](#características-principales)
- [Requisitos](#requisitos-técnico)
- [Instalación](#instalación-y-desarrollo)
- [Estructura](#estructura-de-salida)
- [Notas Técnicas](#notas-de-mantenimiento)
- [Contribución](#contribución)
- [Licencia](#licencia)
- [Familia G360](#familia-g360)

---

## 🚀 Características Principales

- **Extracción Inteligente**: Filtrado por rango de fechas y palabras clave directamente en el servidor MAPI (vía `@SQL`).
- **Gestión de Adjuntos**: Descarga automática de archivos con renombrado cronológico (`YYYYMMDD_Nombre.ext`).
- **Preservación de Contexto**: Generación de archivos `.txt` vinculados que contienen el cuerpo explicativo del correo.
- **Arquitectura Anti-Bloqueo**: Ejecución en hilos separados (STA) para mantener la interfaz Flet fluida.
- **Sistema de Auditoría**: Registro detallado de operaciones en `audit.log` y visualización en tiempo real en la UI.

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
└── YYYY-MM-DD/ (Fecha de proceso)
    ├── 20240520_1430_OC_12345.pdf
    ├── 20240520_1430_OC_12345_CUERPO.txt  <-- Contexto del correo
    └── ...
```

## 🛡️ Notas de Mantenimiento

- **Threading**: Outlook usa *Single Threaded Apartment* (STA). Cada hilo que interactúa con Outlook debe invocar `pythoncom.CoInitialize()`.
- **Zombies COM**: El motor está diseñado para liberar referencias y llamar a `CoUninitialize()` para evitar que procesos `OUTLOOK.EXE` queden huérfanos.
- **Logs**: Cualquier error crítico de conexión se registra en `audit.log` en la raíz del proyecto.
- **Cancelación**: La extracción puede cancelarse en cualquier momento. El motor verifica la señal de cancelación en cada iteración.
- **Validación**: Se valida la selección de carpeta antes de iniciar la extracción para evitar errores.

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

## 🔧 Cambios Recientes (v1.0.1)

- **Corregido**: Bug de cancelación de extracción
- **Corregido**: Validación de selección de carpeta
- **Mejorado**: Logging de errores detallado
- **Mejorado**: Manejo de excepciones
- **Agregado**: Archivo `__init__.py` para el paquete
- **Agregado**: Archivo `build.bat` para construcción
- **Actualizado**: `.gitignore` con reglas completas
- **Actualizado**: `pyproject.toml` con herramientas de desarrollo

---
**Marca**: G360
**Isotipo**: 3 puntos verticales paralelos (gris-verde-gris) + chevron `>`
**Autor**: Carlos Cusi
**Desarrollo**: Con asistencia de herramientas de código IA (Vibe Code)
**Powered by**: [g360-signature](https://github.com/carloscus/g360-signature)
