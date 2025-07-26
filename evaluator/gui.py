# evaluator/gui.py

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import openpyxl
import re
import unicodedata
import json
import os
import shutil
from datetime import datetime
import logging

from . import clients
from . import mapping
from . import processor


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        logging.info("=====================================")
        logging.info("Iniciando la aplicación (versión simplificada)...")
        self.title("Evaluador Canvas")
        self.geometry("1000x500")

        self.source_type = tk.StringVar(value="excel")
        self.excel_file_path = None
        self.spreadsheet_id = None

        self.cursos_canvas_dict = {}
        self.tareas_canvas_dict = {}
        self.df_alumnos_del_curso = None
        self.trimester_data_map = []

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="20");
        main_frame.pack(fill="both", expand=True)
        top_frame = ttk.Frame(main_frame);
        top_frame.pack(fill="x", pady=5)

        canvas_frame = ttk.LabelFrame(top_frame, text="Paso 1: Origen de Notas (Canvas)", padding="10");
        canvas_frame.pack(side="left", fill="x", expand=True, padx=5)
        self.btn_conectar_canvas = ttk.Button(canvas_frame, text="Conectar y Cargar Cursos",
                                              command=self._load_canvas_courses);
        self.btn_conectar_canvas.pack(fill="x", pady=5)
        ttk.Label(canvas_frame, text="Curso de Canvas:").pack(anchor="w")
        self.combo_canvas_cursos = ttk.Combobox(canvas_frame, state="disabled", exportselection=False);
        self.combo_canvas_cursos.pack(fill="x")
        self.combo_canvas_cursos.bind("<<ComboboxSelected>>", self._on_course_selected)

        ttk.Label(canvas_frame, text="Tarea de Canvas (para obtener notas):").pack(anchor="w", pady=(10, 0))
        self.combo_canvas_tareas = ttk.Combobox(canvas_frame, state="disabled", exportselection=False);
        self.combo_canvas_tareas.pack(fill="x")
        self.combo_canvas_tareas.bind("<<ComboboxSelected>>", self._on_canvas_task_selected)

        dest_frame = ttk.LabelFrame(top_frame, text="Paso 2: Destino de Notas", padding="10");
        dest_frame.pack(side="left", fill="x", expand=True, padx=5)
        source_chooser_frame = ttk.Frame(dest_frame);
        source_chooser_frame.pack(fill="x", pady=5)
        ttk.Radiobutton(source_chooser_frame, text="Excel Local", variable=self.source_type, value="excel",
                        command=self._on_source_change).pack(side="left", padx=5)
        ttk.Radiobutton(source_chooser_frame, text="Google Sheets", variable=self.source_type, value="sheets",
                        command=self._on_source_change).pack(side="left", padx=5)

        self.excel_controls_frame = ttk.Frame(dest_frame)
        self.btn_load_excel = ttk.Button(self.excel_controls_frame, text="Seleccionar Plantilla Excel",
                                         command=self._select_excel_file);
        self.btn_load_excel.pack(fill="x", pady=5)

        self.sheets_controls_frame = ttk.Frame(dest_frame)
        ttk.Label(self.sheets_controls_frame, text="URL de la Hoja de Google:").pack(anchor="w")
        self.entry_gsheet_url = ttk.Entry(self.sheets_controls_frame);
        self.entry_gsheet_url.pack(fill="x", pady=2)
        self.btn_load_gsheet = ttk.Button(self.sheets_controls_frame, text="Cargar Hoja de Google",
                                          command=self._load_google_sheet);
        self.btn_load_gsheet.pack(fill="x", pady=5)

        action_frame = ttk.LabelFrame(main_frame, text="Paso 3: Ejecutar", padding="10");
        action_frame.pack(fill="x", pady=20)
        ttk.Label(action_frame, text="Trimestre de Destino:").pack(anchor="w")
        self.combo_trimestre = ttk.Combobox(action_frame, state="disabled", exportselection=False);
        self.combo_trimestre.pack(fill="x", pady=2)
        ttk.Label(action_frame, text="Tarea de Destino:").pack(anchor="w")
        self.combo_excel_tareas = ttk.Combobox(action_frame, state="disabled", exportselection=False);
        self.combo_excel_tareas.pack(fill="x", pady=2)

        self.btn_escribir = ttk.Button(action_frame, text="Cotejar y Escribir Todas las Notas",
                                       command=self._execute_full_write, state="disabled");
        self.btn_escribir.pack(fill="x", ipady=10, pady=10)

        self._on_source_change()

    def _on_source_change(self):
        source = self.source_type.get()
        if source == "excel":
            self.sheets_controls_frame.pack_forget()
            self.excel_controls_frame.pack(fill="x")
        else:
            self.excel_controls_frame.pack_forget()
            self.sheets_controls_frame.pack(fill="x")

    def _load_canvas_courses(self):
        try:
            self.cursos_canvas_dict = clients.obtener_cursos()
            self.combo_canvas_cursos['values'] = list(self.cursos_canvas_dict.keys())
            self.combo_canvas_cursos.config(state="readonly")
            messagebox.showinfo("Éxito", f"Se cargaron {len(self.cursos_canvas_dict)} cursos de Canvas.")
        except Exception as e:
            messagebox.showerror("Error de Conexión", f"No se pudo conectar a Canvas: {e}")

    def _on_course_selected(self, event=None):
        nombre_curso = self.combo_canvas_cursos.get()
        if not nombre_curso: return
        curso_id = self.cursos_canvas_dict[nombre_curso]
        try:
            self.df_alumnos_del_curso = clients.obtener_alumnos(curso_id)
            self.tareas_canvas_dict = clients.obtener_tareas(curso_id)
            self.combo_canvas_tareas['values'] = list(self.tareas_canvas_dict.keys())
            self.combo_canvas_tareas.config(state="readonly")
            messagebox.showinfo("Curso Seleccionado",
                                f"Se han cargado {len(self.df_alumnos_del_curso)} alumnos y {len(self.tareas_canvas_dict)} tareas. Selecciona una tarea.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los datos del curso: {e}")

    def _on_canvas_task_selected(self, event=None):
        nombre_curso = self.combo_canvas_cursos.get()
        nombre_tarea = self.combo_canvas_tareas.get()
        if not all([nombre_curso, nombre_tarea, self.df_alumnos_del_curso is not None]): return

        curso_id = self.cursos_canvas_dict[nombre_curso]
        tarea_id = self.tareas_canvas_dict[nombre_tarea]

        try:
            df_calificaciones = clients.obtener_calificaciones(curso_id, tarea_id)
            df_alumnos_renamed = self.df_alumnos_del_curso.rename(columns={'id': 'user_id'})
            df_final = pd.merge(df_calificaciones, df_alumnos_renamed, on='user_id', how='right')
            df_final.dropna(subset=['name'], inplace=True)

            def round_score(score):
                try:
                    return round(float(score), 1)
                except (ValueError, TypeError):
                    return score

            df_final['score'] = df_final['score'].apply(round_score)

            df_final.to_json("canvas_grades_to_write.json", orient='records', indent=4, force_ascii=False)
            logging.info(f"Guardadas {len(df_final)} notas en 'canvas_grades_to_write.json'.")
            messagebox.showinfo("Notas de Canvas Guardadas",
                                f"Se han extraído y guardado {len(df_final)} notas para la tarea seleccionada.")
            self._check_if_ready_to_write()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron obtener las calificaciones: {e}")

    def _select_excel_file(self):
        path = filedialog.askopenfilename(title="Selecciona la plantilla Excel",
                                          filetypes=[("Archivos de Excel", "*.xlsx")])
        if path:
            self._refresh_excel_data(path)

    def _refresh_excel_data(self, path):
        self.excel_file_path = path
        try:
            workbook = openpyxl.load_workbook(path, data_only=True)
            self.trimester_data_map = mapping.build_map_from_excel(workbook)
            self._update_dest_combos()
            messagebox.showinfo("Excel Cargado", "Archivo Excel cargado y mapeado correctamente.")
        except Exception as e:
            messagebox.showerror("Error al leer Excel", f"No se pudo procesar el archivo Excel:\n{e}")

    def _load_google_sheet(self):
        url = self.entry_gsheet_url.get().strip()
        if not url: messagebox.showwarning("URL Vacía", "Por favor, pega la URL."); return
        spreadsheet_id = self._get_spreadsheet_id_from_url(url)
        if not spreadsheet_id: messagebox.showerror("URL Inválida", "La URL no parece ser válida."); return

        self.spreadsheet_id = spreadsheet_id
        try:
            sheet_data = clients.get_gsheet_values(self.spreadsheet_id, "EVALUACIÓN!A1:Z50")
            self.trimester_data_map = mapping.build_map_from_gsheet_data(sheet_data)
            self._update_dest_combos()
            messagebox.showinfo("Google Sheet Cargado", "Hoja de Google cargada y mapeada correctamente.")
        except Exception as e:
            messagebox.showerror("Error al Cargar", f"No se pudo cargar o procesar la hoja de Google:\n{e}")

    def _get_spreadsheet_id_from_url(self, url):
        match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
        return match.group(1) if match else None

    def _update_dest_combos(self):
        trimestre_names = [t['trimestre_name'] for t in self.trimester_data_map]
        self.combo_trimestre['values'] = trimestre_names
        if trimestre_names:
            self.combo_trimestre.config(state="readonly");
            self.combo_trimestre.set(trimestre_names[0])
        self._on_trimestre_selected()
        self._check_if_ready_to_write()

    def _on_trimestre_selected(self, event=None):
        trimestre_str = self.combo_trimestre.get()
        if not trimestre_str: return
        trimestre_info = next((t for t in self.trimester_data_map if t['trimestre_name'] == trimestre_str), None)
        if trimestre_info and trimestre_info['tasks']:
            task_names = sorted(list(trimestre_info['tasks'].keys()))
            self.combo_excel_tareas['values'] = task_names
            self.combo_excel_tareas.set(task_names[0])
            self.combo_excel_tareas.config(state="readonly")
        else:
            self.combo_excel_tareas.set('')
            self.combo_excel_tareas.config(state="disabled")

    def _execute_full_write(self):
        dest_config = {
            'type': self.source_type.get(),
            'trimestre': self.combo_trimestre.get(),
            'tarea': self.combo_excel_tareas.get(),
            'path': self.excel_file_path,
            'id': self.spreadsheet_id
        }
        try:
            result = processor.run_grade_processing(dest_config)
            summary_message = (
                f"Proceso completado.\n\n"
                f"Alumnos de Canvas procesados: {result['processed']}\n"
                f"Notas escritas con éxito: {result['written']}\n"
                f"Alumnos no encontrados: {result['not_found']}"
            )
            if result.get('backup_path'):
                summary_message += f"\n\nCopia de seguridad creada en:\n{os.path.basename(result['backup_path'])}"
            if result['not_found'] > 0:
                summary_message += "\n\nConsulta 'app.log' para ver los nombres de los alumnos no encontrados."
            messagebox.showinfo("Resumen de la Operación", summary_message)
            if dest_config['type'] == 'excel':
                self._refresh_excel_data(self.excel_file_path)
        except Exception as e:
            logging.error(f"Fallo en el proceso principal de escritura: {e}", exc_info=True)
            messagebox.showerror("Error en el Proceso", f"Ha ocurrido un error:\n{e}")

    def _check_if_ready_to_write(self):
        canvas_ready = os.path.exists("canvas_grades_to_write.json")
        dest_excel_ready = self.source_type.get() == 'excel' and self.excel_file_path
        dest_gsheet_ready = self.source_type.get() == 'sheets' and self.spreadsheet_id
        if canvas_ready and (dest_excel_ready or dest_gsheet_ready):
            self.btn_escribir.config(state="normal")
        else:
            self.btn_escribir.config(state="disabled")

    def _on_task_selected(self, event=None):
        # Esta función ahora solo sirve para chequear si el botón debe activarse
        self._check_if_ready_to_write()