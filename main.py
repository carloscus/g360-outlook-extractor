"""
Punto de entrada principal para G360 Outlook Extractor.
Este archivo redirige al paquete src/g360_outlook_extractor.
"""
import sys
import os

# Agregar src al path para importaciones
sys.path.insert(0, os.path.join(os.path.dirname(__file__), 'src'))

from g360_outlook_extractor.main import main
import flet as ft

if __name__ == "__main__":
    ft.run(main)
