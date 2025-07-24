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
from . import matcher


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        logging.info("=====================================")
        logging.info("Iniciando la aplicación (versión simplificada)...")
        self.title("Evaluador Canvas")
        self.geometry("1000x450")

        self.source_type = tk.StringVar(value="excel")
        self.excel_file_path = None
        self.spreadsheet_id = None

        self.cursos_canvas_dict = {}
        self.trimester_data_map = []

        self._create_widgets()

    def _create_widgets(self):
        main_frame = ttk.Frame(self, padding="20");
        main_frame.pack(fill="both", expand=True)
        top_frame = ttk.Frame(main_frame);
        top_frame.pack(fill="x", pady=5)

        canvas_frame = ttk.LabelFrame(top_frame, text="Paso 1: Origen de Alumnos (Canvas)", padding="10");
        canvas_frame.pack(side="left", fill="x", expand=True, padx=5)
        self.btn_conectar_canvas = ttk.Button(canvas_frame, text="Conectar y Cargar Cursos",
                                              command=self._load_canvas_courses);
        self.btn_conectar_canvas.pack(fill="x", pady=5)
        ttk.Label(canvas_frame, text="Curso de Canvas:").pack(anchor="w")
        self.combo_canvas_cursos = ttk.Combobox(canvas_frame, state="disabled", exportselection=False);
        self.combo_canvas_cursos.pack(fill="x")
        self.combo_canvas_cursos.bind("<<ComboboxSelected>>", self._on_course_selected)

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

        action_frame = ttk.LabelFrame(main_frame, text="Paso 3: Ejecutar Prueba", padding="10");
        action_frame.pack(fill="x", pady=20)
        ttk.Label(action_frame, text="Trimestre de Destino:").pack(anchor="w")
        self.combo_trimestre = ttk.Combobox(action_frame, state="disabled", exportselection=False);
        self.combo_trimestre.pack(fill="x", pady=2)
        ttk.Label(action_frame, text="Tarea de Destino:").pack(anchor="w")
        self.combo_excel_tareas = ttk.Combobox(action_frame, state="disabled", exportselection=False);
        self.combo_excel_tareas.pack(fill="x", pady=2)

        self.btn_escribir = ttk.Button(action_frame, text="Procesar y Escribir Nota de Prueba al Primer Alumno",
                                       command=self._execute_test_write, state="disabled");
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
            df_alumnos = clients.obtener_alumnos(curso_id)
            df_alumnos.dropna(subset=['name'], inplace=True)
            df_alumnos.to_json("canvas_students.json", orient='records', indent=4, force_ascii=False)
            logging.info(f"Guardados {len(df_alumnos)} alumnos de Canvas en 'canvas_students.json'.")
            messagebox.showinfo("Alumnos de Canvas Guardados",
                                f"Se han extraído y guardado {len(df_alumnos)} alumnos en el archivo 'canvas_students.json'.")
            self._check_if_ready_to_write()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar los alumnos del curso: {e}")

    def _select_excel_file(self):
        path = filedialog.askopenfilename(title="Selecciona la plantilla Excel",
                                          filetypes=[("Archivos de Excel", "*.xlsx")])
        if path:
            self.excel_file_path = path
            try:
                workbook = openpyxl.load_workbook(path, data_only=True)
                sheet = workbook['EVALUACIÓN']
                alumnos = [sheet[f'C{i}'].value for i in range(10, 45) if sheet[f'C{i}'].value]

                with open('dest_students.json', 'w', encoding='utf-8') as f:
                    json.dump(alumnos, f, indent=4, ensure_ascii=False)

                self.trimester_data_map = mapping.build_map_from_excel(workbook)
                self._update_dest_combos()
                messagebox.showinfo("Excel Cargado",
                                    f"Se han extraído {len(alumnos)} alumnos a 'dest_students.json' y se ha mapeado el archivo.")
            except Exception as e:
                messagebox.showerror("Error al leer Excel", f"No se pudo procesar el archivo Excel:\n{e}")

    def _load_google_sheet(self):
        url = self.entry_gsheet_url.get().strip()
        if not url: messagebox.showwarning("URL Vacía", "Por favor, pega la URL de tu Google Sheet."); return
        spreadsheet_id = self._get_spreadsheet_id_from_url(url)
        if not spreadsheet_id: messagebox.showerror("URL Inválida", "La URL no parece ser válida."); return

        self.spreadsheet_id = spreadsheet_id
        try:
            sheet_data = clients.get_gsheet_values(self.spreadsheet_id, "EVALUACIÓN!A1:Z50")
            alumnos = [row[2] for row in sheet_data[9:44] if len(row) > 2 and row[2]]

            with open('dest_students.json', 'w', encoding='utf-8') as f:
                json.dump(alumnos, f, indent=4, ensure_ascii=False)

            self.trimester_data_map = mapping.build_map_from_gsheet_data(sheet_data)
            self._update_dest_combos()
            messagebox.showinfo("Google Sheet Cargado",
                                f"Se han extraído {len(alumnos)} alumnos a 'dest_students.json'.")
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

    def _execute_test_write(self):
        try:
            with open('canvas_students.json', 'r', encoding='utf-8') as f:
                canvas_students = json.load(f)
            if not canvas_students:
                messagebox.showerror("Error", "'canvas_students.json' está vacío.");
                return

            first_student = canvas_students[0]
            student_name_to_find = first_student['name']
            nota_de_prueba = 9.9

            trimestre_str = self.combo_trimestre.get()
            target_task = self.combo_excel_tareas.get()
            trimestre_info = next((t for t in self.trimester_data_map if t['trimestre_name'] == trimestre_str), None)
            target_col = trimestre_info['tasks'].get(target_task) if trimestre_info else None
            if not target_col:
                messagebox.showerror("Error", "No se ha podido determinar la columna de destino.");
                return

            if self.source_type.get() == "excel":
                self._write_test_to_excel(student_name_to_find, target_col, nota_de_prueba)
            else:
                self._write_test_to_gsheet(student_name_to_find, target_col, nota_de_prueba)

        except FileNotFoundError as e:
            messagebox.showerror("Archivo no Encontrado", f"Asegúrate de que el archivo '{e.filename}' existe.")
        except Exception as e:
            logging.error(f"Error en la escritura de prueba: {e}", exc_info=True)
            messagebox.showerror("Error en el Proceso", f"Ha ocurrido un error inesperado:\n{e}")

    # --- CORRECCIÓN CLAVE ---
    def _write_test_to_excel(self, student_name, column, grade):
        # 1. Abrir una versión de SOLO LECTURA para buscar
        workbook_read = openpyxl.load_workbook(self.excel_file_path, data_only=True)
        sheet_read = workbook_read['EVALUACIÓN']

        row_to_write = matcher.find_match_in_excel(sheet_read, student_name)
        workbook_read.close()  # Cerrar el libro de lectura

        if not row_to_write:
            messagebox.showwarning("No Encontrado", f"No se encontró al alumno '{student_name}' en el archivo Excel.");
            return

        # 2. Abrir una versión de ESCRITURA para modificar y guardar
        workbook_write = openpyxl.load_workbook(self.excel_file_path)
        sheet_write = workbook_write['EVALUACIÓN']

        sheet_write[f"{column}{row_to_write}"].value = grade
        workbook_write.save(self.excel_file_path)
        workbook_write.close()

        messagebox.showinfo("Éxito",
                            f"Se ha escrito la nota {grade} para '{student_name}' en la celda {column}{row_to_write} del archivo Excel.")

    def _write_test_to_gsheet(self, student_name, column, grade):
        sheet_data = clients.get_gsheet_values(self.spreadsheet_id, "EVALUACIÓN!A1:Z50")
        row_to_write = matcher.find_match_in_gsheet(sheet_data, student_name)

        if not row_to_write:
            messagebox.showwarning("No Encontrado", f"No se encontró al alumno '{student_name}' en la hoja de Google.");
            return

        range_to_update = f"EVALUACIÓN!{column}{row_to_write}"
        clients.update_gsheet_values(self.spreadsheet_id, range_to_update, [[grade]])
        messagebox.showinfo("Éxito",
                            f"Se ha escrito la nota {grade} para '{student_name}' en la celda {column}{row_to_write}.")

    def _check_if_ready_to_write(self):
        canvas_ready = os.path.exists("canvas_students.json")
        dest_ready = os.path.exists("dest_students.json")
        if canvas_ready and dest_ready:
            self.btn_escribir.config(state="normal")
        else:
            self.btn_escribir.config(state="disabled")

    def _show_log_window(self):
        logging.info("Abriendo ventana de registro de actividad.")
        log_window = tk.Toplevel(self)
        log_window.title("Registro de Actividad")
        log_window.geometry("800x600")
        text_area = tk.Text(log_window, wrap="word", state="normal", font=("Courier New", 9))
        vsb = ttk.Scrollbar(log_window, orient="vertical", command=text_area.yview)
        text_area.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y");
        text_area.pack(fill="both", expand=True)
        try:
            with open("app.log", 'r', encoding='utf-8') as f:
                log_content = f.read()
            text_area.insert("1.0", log_content)
        except Exception as e:
            text_area.insert("1.0", f"No se pudo leer el archivo de registro: {e}")
        text_area.config(state="disabled");
        text_area.see(tk.END)