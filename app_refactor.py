from pathlib import Path
import tempfile
import streamlit as st
from openpyxl import load_workbook
import pandas as pd
import gc

# El motor de automatización es ahora la única fuente de verdad para la lógica de negocio
from scripts.motor_automatizacion import (
    load_project_ranges,
    load_project_formats,
    discover_and_load_blocks,
    ejecutar_generacion_completa, # <- Nueva función endurecida
    ORDER, # <- Constante de orden
)

# Cargar config de rangos y formatos una vez
RANGOS_ESTATICOS = load_project_ranges()
FORMATOS = load_project_formats()

# ===================== GESTIÓN DE ESTADO ===================== #

# Inicializar variables en el estado de la sesión si no existen
if 'buf_final' not in st.session_state:
    st.session_state.buf_final = None
if 'file_name' not in st.session_state:
    st.session_state.file_name = None
# ATENCIÓN: Se reemplaza 'workbook' por 'temp_file_path' para evitar guardar objetos grandes en sesión.
if 'temp_file_path' not in st.session_state:
    st.session_state.temp_file_path = None
if 'rangos_dinamicos' not in st.session_state:
    st.session_state.rangos_dinamicos = None

# ===================== LÓGICA REACTIVA CENTRAL (INPUT Y PROCESAMIENTO) ===================== #

# 1. Input principal: Carga de archivo
uploaded_file = st.file_uploader(
    "Sube tu archivo UNC en Excel",
    type=["xlsx", "xlsm"],
    help="Sube el archivo Excel para procesar. Se analizarán las hojas y sus contenidos para generar el dictamen.",
    key="file_uploader"
)

# 2. Procesamiento del archivo (Fase de Análisis para Previsualización)
# Se activa si se sube un nuevo archivo.
if uploaded_file is not None and st.session_state.temp_file_path is None:
    with st.spinner("Analizando archivo Excel..."):
        # Guardar en archivo temporal y almacenar su RUTA en la sesión
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            tmp.write(uploaded_file.getvalue())
            st.session_state.temp_file_path = tmp.name
        
        st.session_state.file_name = uploaded_file.name
        
        # Comentario: Gestión de memoria para el análisis.
        # El workbook se carga en una variable local y se libera explícitamente
        # después de extraer los datos necesarios para la previsualización.
        wb = None
        try:
            wb = load_workbook(st.session_state.temp_file_path, data_only=True)
            st.session_state.rangos_dinamicos = discover_and_load_blocks(
                wb, RANGOS_ESTATICOS, FORMATOS
            )
        finally:
            # Paso clave: Liberar el objeto pesado de la memoria
            if wb:
                del wb
                gc.collect()

        st.session_state.buf_final = None  # Limpiar buffer en cada nueva carga
        st.success(f"Archivo '{st.session_state.file_name}' cargado y analizado.")
        st.rerun()

# ===================== INTERFAZ STREAMLIT (RENDERIZADO) ===================== #

st.title("Generador de Dictamen UNC (Refactorizado)")

# --- Layout de dos columnas ---
col1, col2 = st.columns([1, 2])

# --- COLUMNA IZQUIERDA: CONTROLES ---
with col1:
    # Controles de acción se muestran si el análisis inicial ya se realizó
    if st.session_state.temp_file_path:
        st.header("2. Generar y Descargar")

        if st.button("Generar DOCX final", type="primary", help="Crea el documento Word con el dictamen completo.", key="generate_button"):
            with st.spinner("Generando documento final... Por favor espera."):
                # Se llama a la nueva función encapsulada que gestiona la memoria internamente
                st.session_state.buf_final = ejecutar_generacion_completa(
                    workbook_path=st.session_state.temp_file_path,
                    rangos_dinamicos=st.session_state.rangos_dinamicos,
                    formatos=FORMATOS,
                )
            st.success("¡Documento generado con éxito!")

        if st.session_state.buf_final:
            st.download_button(
                label="Descargar DICTAMEN_FINAL.docx",
                data=st.session_state.buf_final,
                file_name="DICTAMEN_FINAL.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                help="Haz clic para descargar el documento Word generado.",
                key="download_button"
            )
        else:
            st.info("Genera el dictamen para habilitar la descarga.")

        st.markdown("---")
        if st.button("Empezar de Nuevo", help="Limpia la aplicación para procesar un nuevo archivo.", key="reset_button"):
            # Limpiar todo el estado de la sesión para un reinicio limpio
            for key in list(st.session_state.keys()):
                del st.session_state[key]
            # Podríamos también borrar el archivo temporal si existe, pero el SO lo suele gestionar.
            st.rerun()
    else:
        st.info("Sube un archivo para activar los controles.")

# --- COLUMNA DERECHA: PREVISUALIZACIÓN Y RESULTADOS ---
with col2:
    st.header("Detalle y Previsualización")

    if not st.session_state.temp_file_path:
        st.info("Sube un archivo Excel para comenzar el análisis y la previsualización.")

    elif st.session_state.rangos_dinamicos:
        hojas_disponibles = [s for s in ORDER if s in st.session_state.rangos_dinamicos]
        
        with st.expander("Ver Previsualización de Secciones", expanded=True):
            hoja_sel = st.selectbox(
                "Selecciona una sección/hoja para previsualizar:",
                hojas_disponibles,
                help="Permite ver el contenido detectado en cada sección."
            )

            if hoja_sel:
                bloques = st.session_state.rangos_dinamicos.get(hoja_sel, [])
                if not bloques:
                    st.warning("No se encontraron bloques de contenido para esta sección.")
                else:
                    st.subheader(f"Contenido de '{hoja_sel}'")
                    for i, bloque in enumerate(bloques, 1):
                        st.markdown(f"**Bloque {i}:** Tipo=`{bloque['tipo']}`")
                        if isinstance(bloque.get('contenido'), pd.DataFrame):
                            st.table(bloque['contenido'])
                        else:
                            st.text_area(
                                f"Contenido del Bloque {i}",
                                value=str(bloque.get('contenido', '')),
                                height=100,
                                disabled=True,
                            )
            else:
                st.info("Selecciona una hoja para ver su previsualización.")
                
    elif st.session_state.temp_file_path and not st.session_state.rangos_dinamicos:
        st.warning("El archivo fue cargado, pero no se encontraron secciones o bloques de contenido con los criterios actuales.")
