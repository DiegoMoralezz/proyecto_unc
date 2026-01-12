# -*- coding: utf-8 -*-
# CLI que genera el dictamen final
# a partir del Excel y los rangos,
# reutilizando la lógica de core_secciones.

from pathlib import Path

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from core_secciones import (
    cargar_rangos,
    cargar_workbook,
    cargar_formatos,
    generar_docx_final_a_archivo,
)


BASE_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = BASE_DIR / "config" / "rangos_hojas.json"
FORMATOS_PATH = BASE_DIR / "config" / "formatos_hojas.json"
# Ajusta el nombre del archivo de Excel según tu caso real
EXCEL_PATH = BASE_DIR / "excel" / "UNC Lomas Verdes v01.xlsx"
PLANTILLA_PATH = BASE_DIR / "plantilla" / "plantilla_base_final.docx"

FINAL_DIR = BASE_DIR / "output" / "FINAL"
FINAL_DOC = FINAL_DIR / "DICTAMEN_FINAL.docx"

# Orden estricto (hojas, no nombres de archivos)
ORDER_SHEETS = [
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

FIRST_NUMBERED_SHEET = "Dictamen 1"


def ajustar_numeracion_paginas(orden_efectivo: list[str]) -> None:
    """
    Intenta reiniciar la numeración de páginas a partir de la sección
    correspondiente a FIRST_NUMBERED_SHEET.
    """
    try:
        doc = Document(FINAL_DOC)
    except Exception as e:  # pragma: no cover - solo logging sencillo
        print(f"[ADVERTENCIA] No se pudo abrir el DOCX final para ajustar numeración: {e}")
        return

    try:
        sec_index = orden_efectivo.index(FIRST_NUMBERED_SHEET)
    except ValueError:
        print(
            f"[ADVERTENCIA] FIRST_NUMBERED_SHEET ('{FIRST_NUMBERED_SHEET}') "
            "no está en el orden efectivo; no se ajusta numeración."
        )
        return

    sections = doc.sections
    if sec_index >= len(sections):
        print(
            f"[ADVERTENCIA] No hay suficientes secciones en el documento para ajustar "
            f"numeración (secciones={len(sections)}, idx={sec_index})."
        )
        return

    target_section = sections[sec_index]
    sectPr = target_section._sectPr  # type: ignore[attr-defined]

    pgNumType = sectPr.find(qn("w:pgNumType"))
    if pgNumType is None:
        pgNumType = OxmlElement("w:pgNumType")
        sectPr.append(pgNumType)

    pgNumType.set(qn("w:start"), "1")
    doc.save(FINAL_DOC)

    print(f"[OK] Numeración reiniciada correctamente a partir de la hoja: {FIRST_NUMBERED_SHEET}")


def main() -> None:
    FINAL_DIR.mkdir(parents=True, exist_ok=True)

    wb = cargar_workbook(EXCEL_PATH)
    rangos = cargar_rangos(CONFIG_PATH)
    try:
        formatos = cargar_formatos(FORMATOS_PATH)
    except FileNotFoundError:
        formatos = None

    # Orden efectivo de las hojas que realmente existen y tienen rango
    orden_efectivo = [h for h in ORDER_SHEETS if h in rangos and h in wb.sheetnames]

    print("[INFO] Hojas a incluir en el dictamen final (en orden):")
    for name in orden_efectivo:
        print(f" - {name}")

    if not orden_efectivo:
        raise RuntimeError("No hay hojas válidas para generar el dictamen final.")

    generar_docx_final_a_archivo(
        wb=wb,
        rangos=rangos,
        plantilla_path=PLANTILLA_PATH,
        destino=FINAL_DOC,
        orden=orden_efectivo,
        formatos=formatos,
    )

    print(f"[OK] Documento combinado guardado en: {FINAL_DOC}")

    # Ajuste de numeración de páginas (opcional, mejor esfuerzo)
    ajustar_numeracion_paginas(orden_efectivo)

    print("[FIN] Documento completo generado.")


if __name__ == "__main__":
    main()
