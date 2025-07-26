# test_excel_reader.py

import openpyxl
import os

# --- CONFIGURACIÓN DE LA PRUEBA ---
# Construimos la ruta al archivo basándonos en la información que proporcionaste.
# Esto asume que el script se ejecuta desde la carpeta raíz del proyecto.
FOLDER_PATH = ""
# --- CONFIGURACIÓN DE LA PRUEBA (Versión modificada para ejecutar desde 'evaluator') ---

# El nombre del archivo sigue siendo el mismo
FILE_NAME = "SISTEMAS OPERATIVOS MONOPUESTO _ 1ºA SMR _ 24-25 (1).xlsx"

# Como el script está en la misma carpeta que el Excel, la ruta es solo el nombre del archivo.
FULL_FILE_PATH = FILE_NAME

print(f"--- INICIANDO PRUEBA DE LECTURA AISLADA ---")
print(f"Intentando leer el archivo: {FULL_FILE_PATH}\n")

# Comprobar si el archivo existe antes de intentar abrirlo
if not os.path.exists(FULL_FILE_PATH):
    print(f"ERROR: El archivo no se encuentra en la ruta especificada.")
    print(
        "Asegúrate de que el nombre del archivo es correcto y que estás ejecutando este script desde la carpeta raíz del proyecto.")
else:
    try:
        # Abrimos el libro exactamente como lo hace la aplicación, pidiendo solo los valores.
        workbook = openpyxl.load_workbook(FULL_FILE_PATH, data_only=True)
        sheet = workbook['EVALUACIÓN']
        print("Archivo Excel y hoja 'EVALUACIÓN' cargados con éxito.\n")
        print("--- Contenido de la columna de Alumnos (C10 a C20) ---")

        found_names = 0
        for i in range(10, 21):  # Leer las primeras 11 filas de alumnos
            cell_ref = f'C{i}'
            cell_value = sheet[cell_ref].value
            if cell_value:
                found_names += 1
            print(f"Celda {cell_ref}: {cell_value} (Tipo: {type(cell_value)})")

        workbook.close()

        print("\n--- DIAGNÓSTICO ---")
        if found_names > 0:
            print("RESULTADO: ¡Se han leído nombres de alumnos correctamente!")
            print(
                "Esto sugiere que el problema NO está en la librería 'openpyxl', sino en cómo se pasa el objeto 'workbook' dentro de la aplicación.")
        else:
            print("RESULTADO: NO se ha leído ningún nombre de alumno (todos son 'None').")
            print(
                "Esto CONFIRMA que el problema está en cómo 'openpyxl' lee este archivo específico cuando se usa 'data_only=True'.")
            print(
                "La librería no puede resolver los valores de las fórmulas ('=INSTITUTO!C10') sin que Excel los haya pre-calculado y guardado.")

        print("\nPrueba finalizada.")

    except Exception as e:
        print(f"\nHa ocurrido un error durante la prueba: {e}")