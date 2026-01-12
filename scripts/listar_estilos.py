from __future__ import annotations

from pathlib import Path
from docx import Document

# Ruta a la plantilla
PLANTILLA_PATH = Path(__file__).resolve().parent.parent / "plantilla" / "plantilla_base_final.docx"

def listar_estilos(plantilla_path: Path):
    """
    Abre un documento de Word y lista los nombres de todos los estilos disponibles.
    """
    if not plantilla_path.exists():
        print(f"Error: No se encontró la plantilla en la ruta: {plantilla_path}")
        return

    try:
        document = Document(plantilla_path)
        styles = document.styles
        
        print("="*30)
        print(f"Estilos encontrados en '{plantilla_path.name}':")
        print("="*30)
        
        # Filtrar y mostrar solo estilos de párrafo y de tabla, que son los más comunes
        
        print("\n--- Estilos de Párrafo ---")
        for s in sorted([style.name for style in styles if style.type == 1]): # WD_STYLE_TYPE.PARAGRAPH = 1
            print(f"- '{s}'")
            
        print("\n--- Estilos de Tabla ---")
        for s in sorted([style.name for style in styles if style.type == 3]): # WD_STYLE_TYPE.TABLE = 3
            print(f"- '{s}'")
            
        print("\n--- Otros Estilos (Carácter, etc.) ---")
        for s in sorted([style.name for style in styles if style.type not in [1, 3]]):
            print(f"- '{s}'")

        print("\n" + "="*30)
        print("\nInstrucción: Copia el nombre exacto del estilo (incluyendo mayúsculas y espacios)")
        print("y pégalo en tu archivo 'config/formatos_hojas.json' en el campo 'style' o 'table_style'.")

    except Exception as e:
        print(f"Ocurrió un error al procesar el documento: {e}")

if __name__ == "__main__":
    listar_estilos(PLANTILLA_PATH)
