from __future__ import annotations

from pathlib import Path
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- IMPORTS DE UI ---
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

# --- IMPORTS DEL MOTOR DE AUTOMATIZACIÓN ---
# El motor es ahora la fuente de verdad para la lógica y las rutas.
from scripts.motor_automatizacion import (
    load_project_ranges,
    load_project_formats,
    discover_and_load_blocks,
    PLANTILLA_PATH,
)


class DictamenDesktopApp:
    """
    App de escritorio para generar dictámenes.
    La carga de dependencias pesadas y archivos se hace de forma diferida
    para garantizar que la UI se inicie siempre.
    """

    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title("Generador de Dictamen UNC - Escritorio")
        self.root.geometry("1200x700")

        # --- Dependencias que se cargarán bajo demanda ---
        self.pd = None
        self.Workbook = None
        self.core_secciones = None  # Se mantiene por ahora para funciones de generación
        self.extractor_inteligente = None # Será eliminado eventualmente

        try:
            style = ttk.Style()
            if "clam" in style.theme_names():
                style.theme_use("clam")
        except Exception:
            pass

        # --- Variables de la UI ---
        self.excel_path_var = tk.StringVar()
        self.bloque_rango_var = tk.StringVar()
        self.bloque_tipo_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Listo")
        self.hoja_actual_var = tk.StringVar(value="(ninguna)")

        # --- Estado de la Aplicación ---
        self.rangos_estaticos = {}
        self.rangos = {}
        self.formatos = None
        self.tipos_disponibles = []
        self.workbook = None
        self.hojas_disponibles: list[str] = []
        self.orden_hojas: list[str] = []

        # --- Carga de configuraciones iniciales (usando el motor) ---
        self._load_initial_configs()

        # --- Construir la UI ---
        self._build_ui()

    def _load_initial_configs(self):
        """Carga las configuraciones JSON de forma segura usando el motor."""
        try:
            self.rangos_estaticos = load_project_ranges()
            self.formatos = load_project_formats()
        except Exception as e:
            self.formatos = None
            messagebox.showerror("Error de Configuración", f"No se pudieron cargar los archivos de configuración: {e}")


        # Calcular los tipos disponibles para la UI
        default_tipos = {
            "texto_normal", "texto_sangria", "titulo_estado", "titulo_nota",
            "Portada_Titulo", "Portada_Subtitulo", "viñetas_circulo", "viñetas_linea",
            "tabla_bg", "tabla_er", "tabla_ecc", "tabla_cf", "tabla_nota",
        }
        tipos_from_cfg = set((self.formatos.get("tipos", {}) or {}).keys()) if self.formatos else set()
        tipos_from_rangos_estaticos = {
            str(b.get("tipo", "")).strip()
            for bloques in self.rangos_estaticos.values() if isinstance(bloques, list)
            for b in bloques if str(b.get("tipo", "")).strip()
        }
        self.tipos_disponibles = sorted(default_tipos | tipos_from_cfg | tipos_from_rangos_estaticos)


    def _discover_and_load_blocks(self) -> dict:
        """
        Llama al motor central para analizar las hojas del libro de Excel y cargar
        los bloques de contenido.
        """
        if not self.workbook:
            return {}
        
        # Llama a la función centralizada desde el motor
        return discover_and_load_blocks(
            self.workbook, self.rangos_estaticos, self.formatos
        )
        
    def _build_ui(self) -> None:
        # ... (El código de construcción de UI se mantiene, referenciando self.métodos) ...
        self._build_menu()
        top_frame = ttk.Frame(self.root, padding=10)
        top_frame.pack(side=tk.TOP, fill=tk.X)
        
        ttk.Label(top_frame, text="Archivo Excel:").pack(side=tk.LEFT)
        entry = ttk.Entry(top_frame, textvariable=self.excel_path_var, width=70)
        entry.pack(side=tk.LEFT, padx=5)
        self.browse_button = ttk.Button(top_frame, text="Buscar...", command=self._browse_excel)
        self.browse_button.pack(side=tk.LEFT)
        self.load_button = ttk.Button(top_frame, text="Cargar Excel", command=self._load_excel)
        self.load_button.pack(side=tk.LEFT, padx=5)

        middle_frame = ttk.Frame(self.root, padding=10)
        middle_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        left_frame = ttk.LabelFrame(middle_frame, text="Hojas y rangos")
        left_frame.pack(side=tk.LEFT, fill=tk.Y)
        columns = ("hoja", "resumen")
        self.tree_hojas = ttk.Treeview(left_frame, columns=columns, show="headings", height=16)
        self.tree_hojas.heading("hoja", text="Hoja")
        self.tree_hojas.heading("resumen", text="Resumen")
        self.tree_hojas.column("hoja", width=150, anchor="w")
        self.tree_hojas.column("resumen", width=120, anchor="center")
        tree_scroll = ttk.Scrollbar(left_frame, orient="vertical", command=self.tree_hojas.yview)
        self.tree_hojas.configure(yscrollcommand=tree_scroll.set)
        self.tree_hojas.pack(side=tk.LEFT, fill=tk.Y)
        tree_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.tree_hojas.bind("<<TreeviewSelect>>", self._on_hoja_select)
        
        order_buttons = ttk.Frame(left_frame)
        order_buttons.pack(fill=tk.X, pady=5)
        ttk.Button(order_buttons, text="Subir", command=self._move_sheet_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(order_buttons, text="Bajar", command=self._move_sheet_down).pack(side=tk.LEFT, padx=2)
        ttk.Button(order_buttons, text="Quitar de dictamen", command=self._remove_sheet_from_order).pack(side=tk.LEFT, padx=2)

        right_frame = ttk.Frame(middle_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10)
        
        info_frame = ttk.Frame(right_frame)
        info_frame.pack(fill=tk.X)
        ttk.Label(info_frame, text="Hoja seleccionada:").pack(side=tk.LEFT)
        ttk.Label(info_frame, textvariable=self.hoja_actual_var, foreground="blue").pack(side=tk.LEFT, padx=5)

        rango_frame = ttk.Frame(right_frame)
        rango_frame.pack(fill=tk.X, pady=5)
        ttk.Label(rango_frame, text="Rango del bloque:").pack(side=tk.LEFT)
        rango_entry = ttk.Entry(rango_frame, textvariable=self.bloque_rango_var, width=25)
        rango_entry.pack(side=tk.LEFT, padx=5)
        ttk.Label(rango_frame, text="Tipo:").pack(side=tk.LEFT)
        self.combo_tipo = ttk.Combobox(rango_frame, textvariable=self.bloque_tipo_var, values=self.tipos_disponibles, width=20, state="readonly")
        self.combo_tipo.pack(side=tk.LEFT, padx=5)

        bloques_frame = ttk.LabelFrame(right_frame, text="Bloques de la hoja")
        bloques_frame.pack(fill=tk.BOTH, expand=True, pady=5)
        cols_bloques = ("rango", "tipo")
        self.tree_bloques = ttk.Treeview(bloques_frame, columns=cols_bloques, show="headings", height=8)
        self.tree_bloques.heading("rango", text="Rango/Info")
        self.tree_bloques.heading("tipo", text="Tipo")
        self.tree_bloques.column("rango", width=150, anchor="w")
        self.tree_bloques.column("tipo", width=150, anchor="w")
        bloques_scroll = ttk.Scrollbar(bloques_frame, orient="vertical", command=self.tree_bloques.yview)
        self.tree_bloques.configure(yscrollcommand=bloques_scroll.set)
        self.tree_bloques.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        bloques_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self.tree_bloques.bind("<<TreeviewSelect>>", self._on_bloque_select)
        
        bloques_buttons = ttk.Frame(right_frame)
        bloques_buttons.pack(fill=tk.X, pady=5)
        ttk.Button(bloques_buttons, text="Añadir bloque", command=self._add_block).pack(side=tk.LEFT, padx=2)
        ttk.Button(bloques_buttons, text="Actualizar bloque", command=self._update_block).pack(side=tk.LEFT, padx=2)
        ttk.Button(bloques_buttons, text="Eliminar bloque", command=self._delete_block).pack(side=tk.LEFT, padx=2)
        ttk.Button(bloques_buttons, text="Subir bloque", command=self._move_block_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(bloques_buttons, text="Bajar bloque", command=self._move_block_down).pack(side=tk.LEFT, padx=2)

        actions_frame = ttk.Frame(right_frame)
        actions_frame.pack(fill=tk.X, pady=5)
        ttk.Button(actions_frame, text="Generar DOCX de esta sección", command=self._generate_section_docx).pack(side=tk.LEFT, padx=5)
        ttk.Button(actions_frame, text="Generar dictamen completo", command=self._generate_full_docx).pack(side=tk.LEFT, padx=5)
        ttk.Button(actions_frame, text="Guardar rangos en JSON", command=self._save_rangos_to_file).pack(side=tk.LEFT, padx=5)

        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor="w", padding=5)
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)
        self._set_status("Listo. Carga un Excel para comenzar.")

    def _build_menu(self) -> None:
        menubar = tk.Menu(self.root)
        menu_archivo = tk.Menu(menubar, tearoff=0)
        menu_archivo.add_command(label="Abrir Excel...", command=self._browse_excel)
        menu_archivo.add_command(label="Cargar Excel", command=self._load_excel)
        menu_archivo.add_separator()
        menu_archivo.add_command(label="Guardar rangos", command=self._save_rangos_to_file)
        menu_archivo.add_separator()
        menu_archivo.add_command(label="Salir", command=self.root.quit)
        menu_ayuda = tk.Menu(menubar, tearoff=0)
        menu_ayuda.add_command(label="Acerca de", command=self._show_about)
        menubar.add_cascade(label="Archivo", menu=menu_archivo)
        menubar.add_cascade(label="Ayuda", menu=menu_ayuda)
        self.root.config(menu=menubar)

    def _set_status(self, text: str) -> None:
        self.status_var.set(text)

    def _browse_excel(self) -> None:
        path = filedialog.askopenfilename(title="Selecciona el archivo Excel", filetypes=[("Archivos Excel", "*.xlsx *.xlsm"), ("Todos los archivos", "*.*")])
        if path:
            self.excel_path_var.set(path)
            self._set_status(f"Excel seleccionado: {path}")

    def _load_excel(self) -> None:
        path_str = self.excel_path_var.get().strip()
        if not path_str:
            messagebox.showwarning("Atención", "Selecciona primero un archivo de Excel.")
            return
        
        try:
            # --- Importación diferida ---
            from core_secciones import cargar_workbook
            self.workbook = cargar_workbook(Path(path_str))
            
            self.rangos = self._discover_and_load_blocks()
            
            self.hojas_disponibles = list(self.workbook.sheetnames)
            self.orden_hojas = [s for s in self.hojas_disponibles if s in self.rangos]
            
            self._populate_tree()
            if self.orden_hojas:
                self._select_sheet_in_tree(self.orden_hojas[0])
            self._set_status("Excel cargado y analizado correctamente.")
            messagebox.showinfo("Éxito", "Excel cargado y analizado correctamente.")
        except Exception as e:
            messagebox.showerror("Error al cargar Excel", str(e))
            self._set_status("Error al cargar Excel.")

    def _move_sheet_up(self) -> None:
        selection = self.tree_hojas.selection()
        if not selection: return
        hoja = selection[0]
        if hoja not in self.orden_hojas: return
        idx = self.orden_hojas.index(hoja)
        if idx > 0:
            self.orden_hojas[idx], self.orden_hojas[idx - 1] = self.orden_hojas[idx - 1], self.orden_hojas[idx]
            self._populate_tree()
            self._select_sheet_in_tree(hoja)

    def _move_sheet_down(self) -> None:
        selection = self.tree_hojas.selection()
        if not selection: return
        hoja = selection[0]
        if hoja not in self.orden_hojas: return
        idx = self.orden_hojas.index(hoja)
        if idx < len(self.orden_hojas) - 1:
            self.orden_hojas[idx], self.orden_hojas[idx + 1] = self.orden_hojas[idx + 1], self.orden_hojas[idx]
            self._populate_tree()
            self._select_sheet_in_tree(hoja)

    def _remove_sheet_from_order(self) -> None:
        selection = self.tree_hojas.selection()
        if not selection: return
        hoja = selection[0]
        if hoja in self.orden_hojas:
            self.orden_hojas.remove(hoja)
            self._populate_tree()
            self.hoja_actual_var.set("(ninguna)")
            self.bloque_rango_var.set("")
    
    def _populate_tree(self) -> None:
        self.tree_hojas.delete(*self.tree_hojas.get_children())
        for hoja in self.orden_hojas:
            bloques = self.rangos.get(hoja, [])
            resumen = f"{len(bloques)} bloques detectados"
            self.tree_hojas.insert("", tk.END, iid=hoja, values=(hoja, resumen))

    def _select_sheet_in_tree(self, hoja: str) -> None:
        if hoja in self.tree_hojas.get_children():
            self.tree_hojas.selection_set(hoja)
            self.tree_hojas.see(hoja)

    def _on_hoja_select(self, event=None) -> None:
        selection = self.tree_hojas.selection()
        if not selection: return
        hoja = selection[0]
        self.hoja_actual_var.set(hoja)
        self._populate_blocks_for_sheet(hoja)

    def _populate_blocks_for_sheet(self, hoja: str) -> None:
        if self.pd is None:
            import pandas as pd
            self.pd = pd

        self.tree_bloques.delete(*self.tree_bloques.get_children())
        bloques = self.rangos.get(hoja, [])
        if not isinstance(bloques, list): return
        for idx, bloque in enumerate(bloques):
            rango_display = bloque.get("rango")
            if rango_display is None and isinstance(bloque.get("contenido"), self.pd.DataFrame):
                df = bloque.get("contenido")
                rango_display = f"Tabla ({len(df)} filas)"
            elif rango_display is None:
                 rango_display = "Contenido simple"
            self.tree_bloques.insert("", tk.END, iid=str(idx), values=(rango_display, bloque.get("tipo", "")))
        self._update_hoja_summary(hoja)

    def _on_bloque_select(self, event=None) -> None:
        selection = self.tree_bloques.selection()
        if not selection: return
        idx = int(selection[0])
        hoja = self.hoja_actual_var.get()
        bloques = self.rangos.get(hoja, [])
        if not isinstance(bloques, list) or idx >= len(bloques): return
        
        rango_val = bloques[idx].get("rango", "")
        self.bloque_rango_var.set(rango_val)
        self.bloque_tipo_var.set(bloques[idx].get("tipo", ""))

    def _add_block(self) -> None:
        hoja = self.hoja_actual_var.get()
        if not hoja or hoja == "(ninguna)":
            messagebox.showwarning("Atención", "Selecciona primero una hoja.")
            return
        rango = self.bloque_rango_var.get().strip()
        if not rango:
            messagebox.showwarning("Atención", "Especifica un rango para el bloque.")
            return
        tipo = self.bloque_tipo_var.get().strip() or "texto_normal"
        bloques = self.rangos.get(hoja, [])
        if not isinstance(bloques, list): bloques = []
        bloques.append({"rango": rango, "tipo": tipo})
        self.rangos[hoja] = bloques
        self._populate_blocks_for_sheet(hoja)

    def _update_block(self) -> None:
        hoja = self.hoja_actual_var.get()
        if not hoja or hoja not in self.rangos:
            messagebox.showwarning("Atención", "Selecciona primero una hoja.")
            return
        selection = self.tree_bloques.selection()
        if not selection:
            messagebox.showwarning("Atención", "Selecciona primero un bloque.")
            return
        idx = int(selection[0])
        bloques = self.rangos.get(hoja, [])
        if not isinstance(bloques, list) or idx >= len(bloques): return
        
        if 'rango' not in bloques[idx]:
            messagebox.showinfo("Info", "Los bloques detectados automáticamente no se pueden editar de esta forma.")
            return

        rango = self.bloque_rango_var.get().strip()
        if not rango:
            messagebox.showwarning("Atención", "Especifica un rango para el bloque.")
            return
        tipo = self.bloque_tipo_var.get().strip() or "texto_normal"
        bloques[idx] = {"rango": rango, "tipo": tipo}
        self.rangos[hoja] = bloques
        self._populate_blocks_for_sheet(hoja)

    def _delete_block(self) -> None:
        hoja = self.hoja_actual_var.get()
        if not hoja or hoja not in self.rangos:
            messagebox.showwarning("Atención", "Selecciona primero una hoja.")
            return
        selection = self.tree_bloques.selection()
        if not selection:
            messagebox.showwarning("Atención", "Selecciona primero un bloque.")
            return
        idx = int(selection[0])
        bloques = self.rangos.get(hoja, [])
        if not isinstance(bloques, list) or idx >= len(bloques): return
        del bloques[idx]
        self.rangos[hoja] = bloques
        self.bloque_rango_var.set("")
        self.bloque_tipo_var.set("")
        self._populate_blocks_for_sheet(hoja)

    def _move_block_up(self) -> None:
        selection = self.tree_bloques.selection()
        if not selection: return
        hoja = self.hoja_actual_var.get()
        if not hoja or hoja not in self.rangos: return
        idx = int(selection[0])
        bloques = self.rangos.get(hoja, [])
        if not isinstance(bloques, list) or idx <= 0 or idx >= len(bloques): return
        bloques[idx - 1], bloques[idx] = bloques[idx], bloques[idx - 1]
        self.rangos[hoja] = bloques
        self._populate_blocks_for_sheet(hoja)
        self.tree_bloques.selection_set(str(idx - 1))

    def _move_block_down(self) -> None:
        selection = self.tree_bloques.selection()
        if not selection: return
        hoja = self.hoja_actual_var.get()
        if not hoja or hoja not in self.rangos: return
        idx = int(selection[0])
        bloques = self.rangos.get(hoja, [])
        if not isinstance(bloques, list) or idx < 0 or idx >= len(bloques) - 1: return
        bloques[idx + 1], bloques[idx] = bloques[idx], bloques[idx + 1]
        self.rangos[hoja] = bloques
        self._populate_blocks_for_sheet(hoja)
        self.tree_bloques.selection_set(str(idx + 1))

    def _update_hoja_summary(self, hoja: str) -> None:
        bloques = self.rangos.get(hoja, [])
        resumen = f"{len(bloques)} bloques detectados"
        if hoja in self.tree_hojas.get_children():
            self.tree_hojas.item(hoja, values=(hoja, resumen))

    def _generate_section_docx(self) -> None:
        if self.workbook is None:
            messagebox.showwarning("Atención", "Carga primero un archivo de Excel.")
            return
        selection = self.tree_hojas.selection()
        if not selection:
            messagebox.showwarning("Atención", "Selecciona primero una hoja.")
            return
        
        # --- Importación diferida ---
        from core_secciones import generar_docx_seccion_a_archivo
        
        hoja = selection[0]
        bloques = self.rangos.get(hoja, [])
        save_path = filedialog.asksaveasfilename(title="Guardar sección como DOCX",defaultextension=".docx", initialfile=f"{hoja}.docx", filetypes=[("Documento Word", "*.docx")])
        if not save_path: return
        try:
            generar_docx_seccion_a_archivo(wb=self.workbook, sheet_name=hoja, bloques=bloques, plantilla_path=PLANTILLA_PATH, destino=Path(save_path), formatos=self.formatos)
            messagebox.showinfo("Éxito", f"Sección guardada en:\n{save_path}")
            self._set_status(f"Sección '{hoja}' generada correctamente.")
        except Exception as e:
            messagebox.showerror("Error al generar sección", str(e))
            self._set_status("Error al generar sección DOCX.")

    def _generate_full_docx(self) -> None:
        if self.workbook is None:
            messagebox.showwarning("Atención", "Carga primero un archivo de Excel.")
            return

        # --- Importación diferida ---
        from core_secciones import generar_docx_final_a_archivo

        save_path = filedialog.asksaveasfilename(title="Guardar dictamen completo como DOCX",defaultextension=".docx", initialfile="DICTAMEN_FINAL.docx", filetypes=[("Documento Word", "*.docx")])
        if not save_path: return
        orden_efectivo = [h for h in self.orden_hojas if h in self.rangos]
        try:
            generar_docx_final_a_archivo(wb=self.workbook, rangos=self.rangos, plantilla_path=PLANTILLA_PATH, destino=Path(save_path), orden=orden_efectivo, formatos=self.formatos)
            messagebox.showinfo("Éxito", f"Dictamen completo guardado en:\n{save_path}")
            self._set_status("Dictamen completo generado correctamente.")
        except Exception as e:
            messagebox.showerror("Error al generar dictamen completo", str(e))
            self._set_status("Error al generar dictamen completo.")

    def _save_rangos_to_file(self) -> None:
        # Filtrar solo bloques manuales para guardar
        rangos_a_guardar = {}
        for hoja, bloques in self.rangos.items():
            bloques_manuales = [b for b in bloques if 'rango' in b]
            if bloques_manuales:
                rangos_a_guardar[hoja] = bloques_manuales

        try:
            with CONFIG_PATH.open("w", encoding="utf-8") as f:
                json.dump(rangos_a_guardar, f, ensure_ascii=False, indent=2)
            messagebox.showinfo("Éxito", f"Rangos manuales guardados en:\n{CONFIG_PATH}")
            self._set_status("Rangos manuales guardados en el archivo de configuración.")
        except Exception as e:
            messagebox.showerror("Error al guardar rangos", str(e))
            self._set_status("Error al guardar rangos en JSON.")

    def _show_about(self) -> None:
        messagebox.showinfo("Acerca de", "Generador de Dictamen UNC\n\nApp de escritorio para generar dictámenes en Word\na partir de un archivo Excel y una plantilla.")

def main() -> None:
    print("DEBUG: Entrando a la función main()")
    root = tk.Tk()
    print("DEBUG: Objeto tk.Tk() creado")
    app = DictamenDesktopApp(root)
    print("DEBUG: Clase DictamenDesktopApp instanciada")
    root.mainloop()
    print("DEBUG: root.mainloop() finalizado.")

if __name__ == "__main__":
    print("DEBUG: El script se está ejecutando como principal (__name__ == '__main__')")
    main()