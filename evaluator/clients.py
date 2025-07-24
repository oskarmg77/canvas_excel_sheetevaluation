# evaluator/clients.py

import logging
import pandas as pd
from canvasapi import Canvas
from config.settings import API_URL, API_KEY
import os
from google.oauth2 import service_account
from googleapiclient.discovery import build

# --- CANVAS CLIENT ---
try:
    canvas = Canvas(API_URL, API_KEY)
    logging.info("Conexión con la API de Canvas establecida correctamente.")
except Exception as e:
    logging.error(f"No se pudo establecer la conexión inicial con Canvas: {e}")
    canvas = None

def obtener_cursos():
    if not canvas: raise ConnectionError("La instancia de Canvas no está disponible.")
    cursos = canvas.get_courses(enrollment_state='active')
    return {curso.name: curso.id for curso in cursos if hasattr(curso, 'name')}

def obtener_tareas(curso_id):
    if not canvas: raise ConnectionError("La instancia de Canvas no está disponible.")
    curso = canvas.get_course(curso_id)
    return {tarea.name: tarea.id for tarea in curso.get_assignments()}

def obtener_alumnos(curso_id):
    if not canvas: raise ConnectionError("La instancia de Canvas no está disponible.")
    curso = canvas.get_course(curso_id)
    alumnos = curso.get_users(enrollment_type=['student'], sort="sortable_name", order="asc")
    lista_alumnos = [{'id': al.id, 'name': al.sortable_name} for al in alumnos if hasattr(al, 'sortable_name')]
    return pd.DataFrame(lista_alumnos)

def obtener_calificaciones(curso_id, tarea_id):
    if not canvas: raise ConnectionError("La instancia de Canvas no está disponible.")
    curso = canvas.get_course(curso_id)
    tarea = curso.get_assignment(tarea_id)
    submissions = tarea.get_submissions()
    return pd.DataFrame([{'user_id': s.user_id, 'score': s.score} for s in submissions])

# --- GOOGLE SHEETS CLIENT ---
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SERVICE_ACCOUNT_FILE = os.path.join(BASE_DIR, "config", "service_account.json")

def get_sheets_service():
    try:
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        return build('sheets', 'v4', credentials=creds)
    except FileNotFoundError:
        logging.error(f"Archivo de credenciales de Google no encontrado en: {SERVICE_ACCOUNT_FILE}")
        raise
    except Exception as e:
        logging.error("Error al inicializar el servicio de Google Sheets: %s", e)
        raise

def get_gsheet_values(spreadsheet_id, range_name):
    service = get_sheets_service()
    result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    return result.get('values', [])

def update_gsheet_values(spreadsheet_id, range_name, values):
    service = get_sheets_service()
    body = {'values': values}
    result = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_name, valueInputOption="USER_ENTERED", body=body).execute()
    logging.info(f"Se han actualizado {result.get('updatedCells')} celdas en Google Sheets.")