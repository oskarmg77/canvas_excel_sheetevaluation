# canvas_client.py

import logging
import pandas as pd
from canvasapi import Canvas
from config.settings import API_URL, API_KEY  # Asumo que tu URL y KEY están en un archivo de configuración

# --- Configuración ---
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

# --- Conexión con Canvas ---
try:
    canvas = Canvas(API_URL, API_KEY)
    logging.info("Conexión con la API de Canvas establecida correctamente.")
except Exception as e:
    logging.error(f"No se pudo establecer la conexión inicial con Canvas: {e}")
    canvas = None


def obtener_cursos():
    """
    Obtiene todos los cursos activos del usuario.
    Retorna un diccionario con {nombre_curso: id_curso}.
    """
    if not canvas:
        logging.error("La instancia de Canvas no está disponible.")
        return {}
    try:
        cursos = canvas.get_courses(enrollment_state='active')
        dict_cursos = {curso.name: curso.id for curso in cursos if hasattr(curso, 'name')}
        logging.info(f"Se encontraron {len(dict_cursos)} cursos activos.")
        return dict_cursos
    except Exception as e:
        logging.error(f"Error al obtener los cursos: {e}")
        return {}


def obtener_tareas(curso_id):
    """
    Obtiene todas las tareas (assignments) de un curso específico.
    Retorna un diccionario con {nombre_tarea: id_tarea}.
    """
    if not canvas:
        logging.error("La instancia de Canvas no está disponible.")
        return {}
    try:
        curso = canvas.get_course(curso_id)
        tareas = curso.get_assignments()
        dict_tareas = {tarea.name: tarea.id for tarea in tareas}
        logging.info(f"Se encontraron {len(dict_tareas)} tareas para el curso {curso_id}.")
        return dict_tareas
    except Exception as e:
        logging.error(f"Error al obtener las tareas para el curso {curso_id}: {e}")
        return {}


# En canvas_client.py

def obtener_alumnos(curso_id):
    """
    Obtiene los alumnos de un curso, ordenados por apellido, y retorna un DataFrame de Pandas.
    """
    if not canvas:
        logging.error("La instancia de Canvas no está disponible.")
        return pd.DataFrame()
    try:
        curso = canvas.get_course(curso_id)
        # --- CAMBIO CLAVE: Pedir a la API que ordene los alumnos ---
        alumnos = curso.get_users(
            enrollment_type=['student'],
            sort="sortable_name", # Ordenar por el campo 'Apellido, Nombre'
            order="asc"          # En orden ascendente
        )

        lista_alumnos = [
            {
                'id': alumno.id,
                # Usamos sortable_name para la visualización y ordenación fiable
                'name': alumno.sortable_name,
                'email': getattr(alumno, 'email', 'N/A')
            }
            for alumno in alumnos if hasattr(alumno, 'sortable_name')
        ]

        df_alumnos = pd.DataFrame(lista_alumnos)
        logging.info(f"Se han obtenido {len(df_alumnos)} alumnos ordenados por apellido.")
        return df_alumnos
    except Exception as e:
        logging.error(f"Error al obtener los alumnos: {e}")
        return pd.DataFrame()


def obtener_calificaciones(curso_id, tarea_id):
    """
    Obtiene las calificaciones de una tarea específica y retorna un DataFrame de Pandas.
    """
    if not canvas:
        logging.error("La instancia de Canvas no está disponible.")
        return pd.DataFrame()
    try:
        curso = canvas.get_course(curso_id)
        tarea = curso.get_assignment(tarea_id)
        submissions = tarea.get_submissions()

        lista_calificaciones = [
            {
                'user_id': submission.user_id,
                'score': submission.score,
            }
            for submission in submissions
        ]

        df_calificaciones = pd.DataFrame(lista_calificaciones)
        logging.info(f"Se han obtenido {len(df_calificaciones)} calificaciones.")
        return df_calificaciones
    except Exception as e:
        logging.error(f"Error al obtener las calificaciones: {e}")
        return pd.DataFrame()