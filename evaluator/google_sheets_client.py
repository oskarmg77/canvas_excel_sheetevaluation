# google_sheets_client.py

import os
import logging
import re
import unicodedata
from google.oauth2 import service_account
from googleapiclient.discovery import build
import openpyxl.utils

logging.basicConfig(level=logging.INFO, format="%(asctime)s [%(levelname)s] %(message)s")

SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(BASE_DIR), "config", "service_account.json")


def get_sheets_service():
    """Inicializa y devuelve el servicio de la API de Google Sheets."""
    try:
        creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
        service = build('sheets', 'v4', credentials=creds)
        return service
    except FileNotFoundError:
        logging.error(f"Archivo de credenciales no encontrado en: {SERVICE_ACCOUNT_FILE}")
        raise
    except Exception as e:
        logging.error("Error al inicializar el servicio de Google Sheets: %s", e)
        raise


def get_values(spreadsheet_id: str, range_name: str) -> list:
    """Función para LEER datos de un rango específico de una hoja."""
    try:
        service = get_sheets_service()
        result = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
        values = result.get('values', [])
        logging.info(f"Se han leído {len(values)} filas del rango '{range_name}'.")
        return values
    except Exception as e:
        logging.error("Error al leer valores de Google Sheets: %s", e)
        raise


def update_values(spreadsheet_id: str, range_name: str, values: list) -> None:
    """Función principal para escribir valores en un rango específico de una hoja."""
    try:
        service = get_sheets_service()
        body = {'values': values}
        result = service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=range_name,
                                                        valueInputOption="USER_ENTERED", body=body).execute()
        logging.info("Se han actualizado %s celdas en el rango '%s'.", result.get('updatedCells'), range_name)
    except Exception as e:
        logging.error("Error al actualizar valores en Google Sheets: %s", e)
        raise


def build_activity_map_from_api_data(data: list, header_text="RESULTADO APRENDIZAJE", activity_row=9):
    """Construye el mapa de trimestres y tareas a partir de los datos leídos de la API."""
    activity_row_index = activity_row - 1
    if len(data) <= activity_row_index:
        raise ValueError("Los datos de la hoja no tienen suficientes filas.")

    header_starts = []
    for row_to_check_idx in range(activity_row_index - 4, activity_row_index):
        if row_to_check_idx >= 0 and row_to_check_idx < len(data):
            for i, cell in enumerate(data[row_to_check_idx]):
                if header_text.lower() in str(cell).lower() and i not in header_starts:
                    header_starts.append(i)
    header_starts.sort()

    if not header_starts:
        raise ValueError(f"No se encontró el texto del encabezado '{header_text}'.")

    header_ranges = []
    for i, start_col in enumerate(header_starts):
        end_col = header_starts[i + 1] - 1 if i + 1 < len(header_starts) else len(data[activity_row_index]) - 1
        header_ranges.append({'min_col': start_col, 'max_col': end_col})

    pattern = re.compile(r"(TAREA|ACTIVIDAD)\s*(\d+)", re.IGNORECASE)
    trimester_map = []
    trimestre_labels = ["1er Trimestre", "2do Trimestre", "3er Trimestre"]

    for i, header_range in enumerate(header_ranges):
        current_trimester_tasks = {}
        activity_row_data = data[activity_row_index]
        for col_idx in range(header_range['min_col'], header_range['max_col'] + 1):
            if col_idx < len(activity_row_data):
                cell_value = activity_row_data[col_idx]
                match = pattern.search(str(cell_value))
                if match:
                    task_name = f"TAREA {match.group(2)}"
                    col_letter = openpyxl.utils.get_column_letter(col_idx + 1)
                    current_trimester_tasks[task_name] = col_letter

        if current_trimester_tasks:
            trimestre_name = trimestre_labels[i] if i < len(trimestre_labels) else f"{i + 1}to Trimestre"
            trimester_map.append({
                'trimestre_name': trimestre_name,
                'tasks': current_trimester_tasks
            })

    return trimester_map