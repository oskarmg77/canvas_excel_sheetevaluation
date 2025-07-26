# evaluator/processor.py

import logging
import json
import openpyxl
import os
import shutil
from datetime import datetime
import pandas as pd

from . import clients
from . import mapping
from . import matcher


def _write_grades_to_excel(file_path: str, workbook_read_only, grades_to_write: list) -> dict:
    logging.info(f"Iniciando proceso de escritura para Excel: {file_path}")

    try:
        root, ext = os.path.splitext(file_path)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{root}_backup_{timestamp}{ext}"
        shutil.copy2(file_path, backup_path)
        logging.info(f"Copia de seguridad creada en: {backup_path}")
    except Exception as e:
        raise IOError(f"No se pudo crear la copia de seguridad: {e}")

    sheet_read = workbook_read_only['EVALUACIÓN']

    written_count = 0
    not_found_students = []
    updates_to_perform = {}

    for record in grades_to_write:
        student_name = record.get('name')
        score = record.get('score')

        if not student_name or pd.isna(score):
            continue

        row_to_write = matcher.find_match_in_excel(sheet_read, student_name)

        if row_to_write:
            try:
                score_value = float(score)
                column = record['target_col']
                cell = f"{column}{row_to_write}"
                updates_to_perform[cell] = score_value
            except (ValueError, TypeError):
                logging.warning(f"Nota no válida para '{student_name}': {score}. Se omite.")
                continue
        else:
            not_found_students.append(student_name)
            logging.warning(f"No se encontró coincidencia para el alumno de Canvas: '{student_name}'")

    if updates_to_perform:
        logging.info(f"Escribiendo {len(updates_to_perform)} notas en el archivo Excel...")
        workbook_write = openpyxl.load_workbook(file_path)
        sheet_write = workbook_write['EVALUACIÓN']
        for cell, grade in updates_to_perform.items():
            sheet_write[cell].value = grade
        workbook_write.save(file_path)
        workbook_write.close()
        written_count = len(updates_to_perform)

    return {
        "processed": len(grades_to_write),
        "written": written_count,
        "not_found": len(not_found_students),
        "not_found_names": not_found_students,
        "backup_path": backup_path
    }


def _write_grades_to_gsheet(spreadsheet_id: str, grades_to_write: list) -> dict:
    logging.info(f"Iniciando proceso de escritura para Google Sheet ID: {spreadsheet_id}")
    sheet_data = clients.get_gsheet_values(spreadsheet_id, "EVALUACIÓN!A1:AZ100")
    written_count = 0
    not_found_students = []

    for record in grades_to_write:
        student_name = record.get('name')
        score = record.get('score')
        if not student_name or pd.isna(score):
            continue
        row_to_write = matcher.find_match_in_gsheet(sheet_data, student_name)
        if row_to_write:
            try:
                score_value = float(score)
                column = record['target_col']
                range_to_update = f"EVALUACIÓN!{column}{row_to_write}"
                clients.update_gsheet_values(spreadsheet_id, range_to_update, [[score_value]])
                written_count += 1
            except (ValueError, TypeError):
                logging.warning(f"Nota no válida para '{student_name}': {score}. Se omite.")
                continue
        else:
            not_found_students.append(student_name)
            logging.warning(f"No se encontró coincidencia para el alumno de Canvas: '{student_name}'")

    return {
        "processed": len(grades_to_write),
        "written": written_count,
        "not_found": len(not_found_students),
        "not_found_names": not_found_students,
        "backup_path": None
    }


def run_grade_processing(dest_config: dict) -> dict:
    logging.info(f"Procesador iniciado. Configuración de destino: {dest_config}")
    try:
        with open('canvas_grades_to_write.json', 'r', encoding='utf-8') as f:
            grades_to_write = json.load(f)
        if not grades_to_write:
            raise ValueError("'canvas_grades_to_write.json' está vacío.")
    except FileNotFoundError:
        raise FileNotFoundError("El archivo 'canvas_grades_to_write.json' no existe.")

    if dest_config['type'] == 'excel':
        workbook = openpyxl.load_workbook(dest_config['path'], data_only=True)
        trimester_map = mapping.build_map_from_excel(workbook)
    else:
        sheet_data = clients.get_gsheet_values(dest_config['id'], "EVALUACIÓN")
        trimester_map = mapping.build_map_from_gsheet_data(sheet_data)

    trimestre_info = next((t for t in trimester_map if t['trimestre_name'] == dest_config['trimestre']), None)
    target_column = trimestre_info['tasks'].get(dest_config['tarea']) if trimestre_info else None
    if not target_column:
        raise ValueError(f"No se pudo encontrar la columna para la tarea '{dest_config['tarea']}'.")

    for record in grades_to_write:
        record['target_col'] = target_column

    if dest_config['type'] == 'excel':
        result = _write_grades_to_excel(dest_config['path'], workbook, grades_to_write)
        workbook.close()
        return result
    else:
        return _write_grades_to_gsheet(dest_config['id'], grades_to_write)