# canvas_gui.py

import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import canvas_client  # Asegúrate de que este import funcione


class CanvasSelector(tk.Tk):
    """
    Interfaz gráfica para seleccionar cursos y tareas de Canvas,
    combinar los datos y descargar los archivos finales.
    """

    def __init__(self):
        super().__init__()
        self.title("Selector de Cursos y Tareas de Canvas")
        self.geometry("450x300")

        # --- Almacenamiento de datos de Canvas ---
        self.cursos_dict = {}
        self.tareas_dict = {}

        # --- Creación de Widgets ---
        self.main_frame = ttk.Frame(self, padding="15")
        self.main_frame.pack(fill="both", expand=True)

        # 1. Botón para iniciar la conexión
        self.btn_conectar = ttk.Button(self.main_frame, text="Conectar y Cargar Cursos", command=self.cargar_cursos)
        self.btn_conectar.pack(pady=5, fill="x")

        # 2. Menú desplegable para los Cursos
        ttk.Label(self.main_frame, text="Paso 1: Selecciona un Curso").pack(anchor="w", pady=(10, 0))
        self.combo_cursos = ttk.Combobox(self.main_frame, state="disabled", exportselection=False)
        self.combo_cursos.pack(pady=5, fill="x")
        self.combo_cursos.bind("<<ComboboxSelected>>", self.cargar_tareas_del_curso)

        # 3. Menú desplegable para las Tareas
        ttk.Label(self.main_frame, text="Paso 2: Selecciona una Tarea").pack(anchor="w", pady=(10, 0))
        self.combo_tareas = ttk.Combobox(self.main_frame, state="disabled", exportselection=False)
        self.combo_tareas.pack(pady=5, fill="x")
        self.combo_tareas.bind("<<ComboboxSelected>>", lambda e: self.btn_descargar.config(state="normal"))

        # 4. Botón final para descargar los datos
        self.btn_descargar = ttk.Button(self.main_frame, text="Descargar y Combinar Datos",
                                        command=self.descargar_y_combinar_datos, state="disabled")
        self.btn_descargar.pack(pady=20, fill="x")

        # 5. Etiqueta de estado en la parte inferior
        self.status_label = ttk.Label(self, text="Bienvenido. Haz clic en 'Conectar' para empezar.", padding="5",
                                      relief="sunken")
        self.status_label.pack(side="bottom", fill="x")

    def cargar_cursos(self):
        # ... (Esta función no cambia)
        self.status_label.config(text="Conectando con Canvas...")
        self.update_idletasks()
        try:
            self.cursos_dict = canvas_client.obtener_cursos()
            if not self.cursos_dict:
                messagebox.showerror("Error", "No se encontraron cursos o hubo un error de conexión.")
                self.status_label.config(text="Error al cargar cursos.")
                return

            self.combo_cursos['values'] = list(self.cursos_dict.keys())
            self.combo_cursos.config(state="readonly")
            self.status_label.config(text="Cursos cargados. Por favor, selecciona uno.")
        except Exception as e:
            messagebox.showerror("Error de Conexión", f"No se pudo conectar a Canvas: {e}")
            self.status_label.config(text="Error de conexión.")

    def cargar_tareas_del_curso(self, event=None):
        # ... (Esta función no cambia)
        nombre_curso = self.combo_cursos.get()
        if not nombre_curso: return
        curso_id = self.cursos_dict[nombre_curso]
        self.status_label.config(text=f"Cargando tareas para '{nombre_curso}'...")
        self.update_idletasks()
        self.combo_tareas.set('')
        self.combo_tareas.config(state="disabled")
        self.btn_descargar.config(state="disabled")
        try:
            self.tareas_dict = canvas_client.obtener_tareas(curso_id)
            if not self.tareas_dict:
                messagebox.showwarning("Sin Tareas", "No se encontraron tareas para este curso.")
                self.status_label.config(text="Curso sin tareas. Elige otro curso.")
                return
            self.combo_tareas['values'] = list(self.tareas_dict.keys())
            self.combo_tareas.config(state="readonly")
            self.status_label.config(text="Tareas cargadas. Por favor, selecciona una.")
        except Exception as e:
            messagebox.showerror("Error", f"No se pudieron cargar las tareas: {e}")
            self.status_label.config(text="Error al cargar tareas.")

    def descargar_y_combinar_datos(self):
        """
        Usa los IDs seleccionados para descargar datos, los combina
        y guarda los archivos CSV finales.
        """
        nombre_curso = self.combo_cursos.get()
        nombre_tarea = self.combo_tareas.get()

        if not nombre_curso or not nombre_tarea:
            messagebox.showwarning("Selección Incompleta", "Por favor, asegúrate de seleccionar un curso y una tarea.")
            return

        curso_id = self.cursos_dict[nombre_curso]
        tarea_id = self.tareas_dict[nombre_tarea]

        self.status_label.config(text="Descargando y combinando datos... Por favor, espera.")
        self.update_idletasks()

        try:
            # 1. Obtener los DataFrames de alumnos y calificaciones
            df_alumnos = canvas_client.obtener_alumnos(curso_id)
            df_calificaciones = canvas_client.obtener_calificaciones(curso_id, tarea_id)

            if df_alumnos.empty or df_calificaciones.empty:
                messagebox.showerror("Error", "No se pudieron obtener los datos de alumnos o calificaciones.")
                self.status_label.config(text="Error en la obtención de datos.")
                return

            # 2. Renombrar la columna 'id' en el df de alumnos para que coincida con 'user_id'
            df_alumnos.rename(columns={'id': 'user_id'}, inplace=True)

            # 3. Unir los DataFrames usando 'user_id'
            df_final = pd.merge(df_calificaciones, df_alumnos[['user_id', 'name']], on='user_id', how='left')

            # 4. Reordenar las columnas para una mejor legibilidad
            if 'name' in df_final.columns:
                df_final = df_final[['user_id', 'name', 'score']]

            # 5. Guardar los archivos finales
            df_alumnos.to_csv('alumnos.csv', index=False)
            df_final.to_csv('calificaciones.csv', index=False)  # Sobreescribimos el archivo con la versión mejorada

            messagebox.showinfo(
                "Descarga Completa",
                "Archivos guardados con éxito:\n\n1. `alumnos.csv`\n2. `calificaciones.csv` (¡Ahora con los nombres de los alumnos!)"
            )
            self.status_label.config(text="¡Datos descargados y combinados con éxito!")

        except Exception as e:
            messagebox.showerror("Error de Procesamiento", f"No se pudieron procesar los datos: {e}")
            self.status_label.config(text="Error durante el procesamiento de datos.")


if __name__ == "__main__":
    app = CanvasSelector()
    app.mainloop()