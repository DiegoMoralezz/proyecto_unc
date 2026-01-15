from pathlib import Path
import tempfile
import streamlit as st
from openpyxl import load_workbook, Workbook

# El motor de automatización es ahora la única fuente de verdad para la lógica de negocio
from scripts.motor_automatizacion import (
    load_project_ranges,
    load_project_formats,
    discover_and_load_blocks,
    PLANTILLA_PATH,  # Importar la ruta de la plantilla desde el motor
)
# La lógica de generación final todavía se importa directamente, se moverá en un paso posterior
from scripts.core_secciones import generar_docx_final_en_memoria
import pandas as pd


# Cargar config de rangos y formatos una vez usando el motor
RANGOS_ESTATICOS = load_project_ranges()
FORMATOS = load_project_formats()

# Orden definido (alineado con unir_documentos.py)
ORDER = [
    "Portada",
    "Contenido",
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




# ===================== GESTIÓN DE ESTADO ===================== #




# Inicializar variables en el estado de la sesión si no existen




if 'buf_final' not in st.session_state:




    st.session_state.buf_final = None




if 'file_name' not in st.session_state:




    st.session_state.file_name = None




if 'workbook' not in st.session_state:




    st.session_state.workbook = None




if 'rangos_dinamicos' not in st.session_state:
























    st.session_state.rangos_dinamicos = None



















# ===================== LÓGICA REACTIVA CENTRAL (INPUT Y PROCESAMIENTO) =====================









# 1. Input principal: Carga de archivo




# Se mueve al nivel raíz para asegurar que se ejecute antes del layout.




uploaded_file = st.file_uploader(




    "Sube tu archivo UNC en Excel",




    type=["xlsx", "xlsm"],




    help="Sube el archivo Excel para procesar. Se analizarán las hojas y sus contenidos para generar el dictamen.",




    key="file_uploader"




)









# 2. Procesamiento del archivo




# Esta lógica ahora se ejecuta a nivel raíz, garantizando que el estado se




# actualice de forma atómica antes de renderizar la UI.




if uploaded_file is not None and st.session_state.workbook is None:




    with st.spinner("Analizando archivo Excel..."):




        # Uso de archivo temporal para compatibilidad con openpyxl




        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:




            tmp.write(uploaded_file.getvalue())




            temp_path = tmp.name




        




        # Carga del workbook y descubrimiento de bloques




        st.session_state.workbook = load_workbook(temp_path, data_only=True)




        st.session_state.file_name = uploaded_file.name




        st.session_state.rangos_dinamicos = discover_and_load_blocks(




            st.session_state.workbook, RANGOS_ESTATICOS, FORMATOS




        )




        st.session_state.buf_final = None  # Limpiar buffer en cada nueva carga




        st.success(f"Archivo '{st.session_state.file_name}' cargado y analizado.")




        # Se recomienda un rerun para asegurar que la UI se actualice con el nuevo estado




        st.rerun()














# ===================== INTERFAZ STREAMLIT (RENDERIZADO) =====================









st.title("Generador de Dictamen UNC")









# --- Layout de dos columnas ---




col1, col2 = st.columns([1, 2])









# --- COLUMNA IZQUIERDA: CONTROLES ---









with col1:









    # El header de carga ya no es necesario aquí, el uploader es global.









    # Esta columna ahora solo muestra los controles de acción post-análisis.









    if st.session_state.workbook:









        st.header("2. Generar y Descargar")









        if st.button("Generar DOCX final", type="primary", help="Crea el documento Word con el dictamen completo.", key="generate_button"):




            with st.spinner("Generando documento final... Por favor espera."):




                st.session_state.buf_final = generar_docx_final_en_memoria(




                    st.session_state.workbook,




                    st.session_state.rangos_dinamicos,




                    PLANTILLA_PATH,




                    orden=ORDER,




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




            st.rerun()




    else:




        st.info("Sube un archivo para activar los controles.")





























# --- COLUMNA DERECHA: PREVISUALIZACIÓN Y RESULTADOS ---














with col2:














    st.header("Detalle y Previsualización")





























    if st.session_state.workbook is None:














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














                














    elif st.session_state.workbook and not st.session_state.rangos_dinamicos:














        st.warning("El archivo fue cargado, pero no se encontraron secciones o bloques de contenido con los criterios actuales.")
