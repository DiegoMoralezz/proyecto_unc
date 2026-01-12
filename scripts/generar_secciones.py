# -*- coding: utf-8 -*-
# CLI que usa core_secciones para generar
# un DOCX por cada hoja/rango definido.

from pathlib import Path

from core_secciones import (
    cargar_rangos,
    cargar_workbook,
    cargar_formatos,
    generar_docx_seccion_a_archivo,
)


# --- RUTAS BASE ---
BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / "config" / "rangos_hojas.json"
FORMATOS_PATH = BASE_DIR / "config" / "formatos_hojas.json"
# Ajusta el nombre del archivo de Excel según tu caso real
EXCEL_PATH = BASE_DIR / "excel" / "UNC Lomas Verdes v01.xlsx"
PLANTILLA_PATH = BASE_DIR / "plantilla" / "plantilla_base_final.docx"
OUTPUT_SECCIONES = BASE_DIR / "output" / "secciones"


def asegurar_directorios() -> None:
    OUTPUT_SECCIONES.mkdir(parents=True, exist_ok=True)


def main() -> None:
    asegurar_directorios()

    wb = cargar_workbook(EXCEL_PATH)
    rangos = cargar_rangos(CONFIG_PATH)
    try:
        formatos = cargar_formatos(FORMATOS_PATH)
    except FileNotFoundError:
        formatos = None

    for hoja, bloques in rangos.items():
        if hoja not in wb.sheetnames:
            print(f"[ADVERTENCIA] La hoja '{hoja}' no existe en el Excel, se omite.")
            continue

        print(f"[INFO] Procesando hoja '{hoja}' con {len(bloques)} bloque(s)...")

        destino = OUTPUT_SECCIONES / f"{hoja}.docx"
        generar_docx_seccion_a_archivo(
            wb=wb,
            sheet_name=hoja,
            bloques=bloques,
            plantilla_path=PLANTILLA_PATH,
            destino=destino,
            formatos=formatos,
        )
        print(f"[OK] Generado: {destino}")

    print("\n[FIN] Todas las secciones han sido generadas correctamente.")


if __name__ == "__main__":
    main()
