# evaluator/matcher.py

import unicodedata


def normalize_name(text: str) -> str:
    """
    Limpia un nombre para la comparación: quita acentos, comas, convierte a
    minúsculas y elimina palabras vacías comunes en español.
    """
    if not isinstance(text, str):
        return ""

    stop_words = {'de', 'la', 'del', 'los', 'las', 'y', 'e', 'maria'}

    nfkd_form = unicodedata.normalize('NFD', text)
    text_no_accents = "".join([c for c in nfkd_form if not unicodedata.combining(c)])

    text_lower = text_no_accents.lower().replace(',', '')

    words = text_lower.split()
    filtered_words = [word for word in words if word not in stop_words]

    return " ".join(filtered_words)


def find_match_in_excel(sheet, name_canvas: str, col: str = 'C', start_row: int = 10, end_row: int = 44) -> int | None:
    """
    Busca la mejor coincidencia para un nombre de Canvas en una hoja de Excel.
    Devuelve el número de fila si se encuentra una coincidencia, de lo contrario None.
    """
    canvas_parts = set(normalize_name(name_canvas).split())
    if not canvas_parts:
        return None

    best_match = {'row': None, 'score': -1, 'name': ''}

    for row_idx in range(start_row, end_row + 1):
        cell_value = sheet[f"{col}{row_idx}"].value
        if not cell_value:
            continue

        dest_parts = set(normalize_name(str(cell_value)).split())
        common_words = len(canvas_parts.intersection(dest_parts))

        if common_words > best_match['score']:
            best_match['score'] = common_words
            best_match['row'] = row_idx
            best_match['name'] = str(cell_value)

    if best_match['row'] is None:
        return None

    is_short_name_match = (
                len(canvas_parts) <= 2 and len(normalize_name(best_match['name']).split()) <= 2 and best_match[
            'score'] >= 1)

    if best_match['score'] >= 2 or is_short_name_match:
        return best_match['row']

    return None


def find_match_in_gsheet(sheet_data: list, name_canvas: str, col_idx: int = 2, start_row: int = 10,
                         end_row: int = 44) -> int | None:
    """
    Busca la mejor coincidencia para un nombre de Canvas en los datos de una hoja de Google.
    Devuelve el número de fila si se encuentra una coincidencia, de lo contrario None.
    """
    canvas_parts = set(normalize_name(name_canvas).split())
    if not canvas_parts:
        return None

    best_match = {'row': None, 'score': -1, 'name': ''}

    for i, row_data in enumerate(sheet_data[start_row - 1: end_row]):
        current_row_idx = i + start_row

        if len(row_data) > col_idx and row_data[col_idx]:
            cell_value = row_data[col_idx]
            dest_parts = set(normalize_name(cell_value).split())
            common_words = len(canvas_parts.intersection(dest_parts))

            if common_words > best_match['score']:
                best_match['score'] = common_words
                best_match['row'] = current_row_idx
                best_match['name'] = cell_value

    if best_match['row'] is None:
        return None

    is_short_name_match = (
                len(canvas_parts) <= 2 and len(normalize_name(best_match['name']).split()) <= 2 and best_match[
            'score'] >= 1)

    if best_match['score'] >= 2 or is_short_name_match:
        return best_match['row']

    return None