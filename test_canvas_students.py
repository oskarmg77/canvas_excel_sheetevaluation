# test_canvas_students.py

import json
import logging
from canvasapi import Canvas
from config.settings import API_URL, API_KEY

# --- Configuración de la Prueba ---
# IMPORTANTE: Reemplaza este valor con el ID numérico del curso que quieres probar.
CURSO_ID_PARA_PROBAR = 78207  # Usando el ID del log de error. ¡Cámbialo si es otro!

# Nombre del archivo donde se guardará la salida JSON
OUTPUT_FILE = "canvas_students_response.json"

# Configuración básica de logging para ver mensajes en la consola
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')


def test_fetch_and_save_students():
    """
    Se conecta a Canvas, obtiene la lista de alumnos de un curso específico
    y guarda los datos relevantes en un archivo JSON.
    """
    logging.info("Iniciando prueba de obtención de alumnos de Canvas...")

    try:
        # 1. Conexión con la API de Canvas
        logging.info(f"Conectando a Canvas en la URL: {API_URL}")
        canvas = Canvas(API_URL, API_KEY)

        # 2. Obtener el objeto del curso
        logging.info(f"Obteniendo el curso con ID: {CURSO_ID_PARA_PROBAR}")
        curso = canvas.get_course(CURSO_ID_PARA_PROBAR)
        logging.info(f"Curso encontrado: '{curso.name}'")

        # 3. Obtener la lista de usuarios/alumnos
        logging.info("Pidiendo la lista de alumnos, ordenada por 'sortable_name'...")
        alumnos = curso.get_users(
            enrollment_type=['student'],
            sort="sortable_name",
            order="asc"
        )

        # 4. --- CORRECCIÓN ---
        # Convertir los objetos de Canvas a un diccionario limpio, extrayendo solo los atributos de datos.
        lista_alumnos_dict = []
        for alumno in alumnos:
            # Creamos un diccionario explícito con los campos que nos interesan.
            # Usamos getattr para obtener los valores de forma segura.
            alumno_data = {
                'id': getattr(alumno, 'id', None),
                'name': getattr(alumno, 'name', None),
                'sortable_name': getattr(alumno, 'sortable_name', None),
                'short_name': getattr(alumno, 'short_name', None),
                'sis_user_id': getattr(alumno, 'sis_user_id', None),
                'login_id': getattr(alumno, 'login_id', None),
                'email': getattr(alumno, 'email', None)
            }
            lista_alumnos_dict.append(alumno_data)

        logging.info(f"Se han procesado {len(lista_alumnos_dict)} alumnos.")

        # 5. Guardar la lista en un archivo JSON
        with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
            json.dump(lista_alumnos_dict, f, ensure_ascii=False, indent=4)

        logging.info(f"¡Éxito! La respuesta de la API se ha guardado en el archivo: '{OUTPUT_FILE}'")

    except Exception as e:
        logging.error(f"Ocurrió un error durante la prueba: {e}", exc_info=True)
        logging.error("Asegúrate de que tus credenciales (API_URL, API_KEY) y el CURSO_ID_PARA_PROBAR son correctos.")


if __name__ == "__main__":
    test_fetch_and_save_students()