from pathlib import Path
import sys

import streamlit as st
from openpyxl import load_workbook

# Configurar rutas base
BASE_DIR = Path(__file__).resolve().parent
CONFIG_PATH = BASE_DIR / "config" / "rangos_hojas.json"
FORMATOS_PATH = BASE_DIR / "config" / "formatos_hojas.json"
PLANTILLA_PATH = BASE_DIR / "plantilla" / "plantilla_base_final.docx"

# Asegurar import del núcleo de secciones
sys.path.append(str(BASE_DIR / "scripts"))
from core_secciones import (  # type: ignore  # noqa: E402
    cargar_rangos,
    cargar_formatos,
    extraer_seccion_desde_hoja,
    generar_docx_final_en_memoria,
)
import pandas as pd
from scripts.extractor_inteligente import extraer_bloques_desde_hoja # type: ignore

# Cargar config de rangos y formatos una vez
RANGOS = cargar_rangos(CONFIG_PATH)
try:
    FORMATOS = cargar_formatos(FORMATOS_PATH)
except FileNotFoundError:
    FORMATOS = None

# Orden definido (alineado con unir_documentos.py)
ORDER = [
    "Portada",
    "contenido",
    "Dictamen 1",
    "Dictamen 2",
    "BG",
    "ER",
    "ECC",
    "CF",
    "Nota 1 y 2",
    "Nota 1 tablas",
    "Nota 3",
    "N4 Efectivo",
    "Nota 5 Txt",
    "N5 Inventarios Inm(Tablas)",
    "N6 Proveedores",
    "N7 Dep��sitos en garant��a",
    "N8 PrǸstamos",
    "N9 Otras Aportaciones Fid",
    "Nota 10 Impuestos",
    "N11 Patrimonio",
    "N12 Vencimientos",
    "N13 Partes relacionadas",
    "Nota 14",
    "Nota 15",
    "Nota 16",
]

def discover_and_load_blocks(wb: Workbook, rangos_manuales: dict, formatos_config: dict) -> dict:
    """
    Analiza todas las hojas de un libro de Excel y carga los bloques de contenido
    usando el método híbrido: automático primero, manual como fallback.
    """
    rangos_descubiertos = {}
    for sheet_name in wb.sheetnames:
        sheet_object = wb[sheet_name]
        
        # 1. Intentar extracción automática
        if not formatos_config:
            # Asegurar que formatos no es None para el extractor
            formatos_config = {}

        bloques_automaticos = extraer_bloques_desde_hoja(sheet_object, formatos_config)
        
        if bloques_automaticos:
            # Usar los bloques descubiertos automáticamente
            rangos_descubiertos[sheet_name] = bloques_automaticos
            st.info(f"Hoja '{sheet_name}': bloques detectados automáticamente.")
        elif sheet_name in rangos_manuales:
            # Fallback: usar la configuración manual si existe para esa hoja
            rangos_descubiertos[sheet_name] = rangos_manuales[sheet_name]
            st.info(f"Hoja '{sheet_name}': usando configuración manual (rangos.json).")

    return rangos_descubiertos


# ===================== INTERFAZ STREAMLIT ===================== #

st.title("Generador de Dictamen UNC")

uploaded_file = st.file_uploader("Sube tu archivo UNC en Excel", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    # Cargar workbook en memoria una sola vez y cachearlo en la sesión
    if 'workbook' not in st.session_state or st.session_state.get('file_name') != uploaded_file.name:
        with st.spinner("Analizando archivo Excel..."):
            st.session_state.workbook = load_workbook(uploaded_file, data_only=True)
            st.session_state.file_name = uploaded_file.name
            
            # Renombrar RANGOS a RANGOS_ESTATICOS para claridad
            RANGOS_ESTATICOS = RANGOS
            
            # Descubrir y cargar los bloques en la sesión
            st.session_state.rangos_dinamicos = discover_and_load_blocks(
                st.session_state.workbook, RANGOS_ESTATICOS, FORMATOS
            )
        st.success(f"Archivo '{uploaded_file.name}' cargado y analizado.")

    # Mostrar hojas disponibles según los rangos dinámicos encontrados
    if 'rangos_dinamicos' in st.session_state and st.session_state.rangos_dinamicos:
        hojas_disponibles = [s for s in ORDER if s in st.session_state.rangos_dinamicos]

        hoja_sel = st.selectbox("Selecciona una sección/hoja para previsualizar", hojas_disponibles)

        if hoja_sel:
            bloques = st.session_state.rangos_dinamicos.get(hoja_sel, [])
            if not bloques:
                 st.warning("No se encontraron bloques de contenido para esta sección.")
            else:
                # La previsualización ahora muestra los bloques detectados
                st.subheader(f"Previsualización de Bloques: {hoja_sel}")
                for i, bloque in enumerate(bloques, 1):
                    st.markdown(f"**Bloque {i}:** Tipo=`{bloque['tipo']}`")
                    if isinstance(bloque.get('contenido'), pd.DataFrame):
                        st.table(bloque['contenido'])
                    else:
                        st.text_area(f"Contenido {i}", value=str(bloque.get('contenido', '')), height=100, disabled=True)


        st.markdown("---")
        st.subheader("Generar dictamen completo")

        if st.button("Generar DOCX final"):
            with st.spinner("Generando documento final... Por favor espera."):
                buf_final = generar_docx_final_en_memoria(
                    st.session_state.workbook, 
                    st.session_state.rangos_dinamicos, 
                    PLANTILLA_PATH, 
                    orden=ORDER, 
                    formatos=FORMATOS
                )
                st.download_button(
                    label="Descargar DICTAMEN_FINAL.docx",
                    data=buf_final,
                    file_name="DICTAMEN_FINAL.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
                st.success("¡Documento generado con éxito!")
    else:
        st.warning("No se encontraron secciones o bloques de contenido en el archivo.")

