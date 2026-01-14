# scripts/motor_automatizacion.py
from __future__ import annotations

from pathlib import Path
import json
import sys
from io import BytesIO

# Dependencias de procesamiento
import pandas as pd
from openpyxl.workbook import Workbook
from openpyxl import load_workbook

# --- 1. ÚNICA FUENTE DE VERDAD PARA RUTAS ---
# Se definen rutas relativas desde la raíz del proyecto. Streamlit Cloud
# ejecuta la app desde la raíz del repo, por lo que estas rutas son portables.
CONFIG_PATH = Path("config/rangos_hojas.json")
FORMATOS_PATH = Path("config/formatos_hojas.json")
PLANTILLA_PATH = Path("plantilla/plantilla_base_final.docx")


# --- 2. GESTIÓN DE IMPORTS INTERNOS ---
# Se elimina la manipulación de sys.path. Al ejecutar `streamlit run app.py`
# desde la raíz, Python maneja los módulos en `scripts/` correctamente.
from scripts.core_secciones import (
    cargar_rangos,
    cargar_formatos,
    extraer_seccion_desde_hoja,
    generar_docx_final_en_memoria,
)
from scripts.extractor_inteligente import extraer_bloques_desde_hoja


# --- 3. LÓGICA DE NEGOCIO CENTRALIZADA ---
# Aquí moveremos las funciones comunes.

def load_project_ranges() -> dict:
    """Carga los rangos de hojas desde el archivo de configuración."""
    return cargar_rangos(CONFIG_PATH)

def load_project_formats() -> dict | None:
    """Carga los formatos de hojas desde el archivo de configuración."""
    try:
        return cargar_formatos(FORMATOS_PATH)
    except FileNotFoundError:
        return None

def discover_and_load_blocks(wb: Workbook, rangos_manuales: dict, formatos_config: dict | None) -> dict:
    """
    Analiza todas las hojas de un libro de Excel y carga los bloques de contenido
    usando el método híbrido: automático primero, manual como fallback.
    Esta función es agnóstica a la UI (no contiene st.info ni messagebox).
    """
    rangos_descubiertos = {}
    for sheet_name in wb.sheetnames:
        sheet_object = wb[sheet_name]
        
        # Asegurar que formatos_config no es None para el extractor
        if not formatos_config:
            formatos_config = {}

        # 1. Intentar extracción automática
        bloques_automaticos = extraer_bloques_desde_hoja(sheet_object, formatos_config)
        
        if bloques_automaticos:
            # Usar los bloques descubiertos automáticamente
            rangos_descubiertos[sheet_name] = bloques_automaticos
        elif sheet_name in rangos_manuales:
            # 2. Fallback: usar la configuración manual si existe para esa hoja
            rangos_descubiertos[sheet_name] = rangos_manuales[sheet_name]

    return rangos_descubiertos

# Aquí se añadirán más funciones como `generar_dictamen_final`
