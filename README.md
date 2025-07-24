# Automatización de Notas desde Canvas LMS a Google Sheets

## Objetivo

El objetivo de este proyecto es desarrollar un sistema que extraiga las calificaciones desde Canvas LMS y las registre automáticamente en un documento de Google Sheets. Con esta solución se busca garantizar que cada nota se inserte en la celda correspondiente, evitando errores en la asignación de calificaciones.

## Fases Completadas

1. **Identificación de la estructura del archivo Excel**  
   - **Hoja relevante:** `EVALUACIÓN`.
   - **Alumnos:** Listados en una columna específica (por defecto la **C**) a partir de la fila **10**.
   - **Actividades/Tareas:** Se detectan en la fila de actividades (por defecto la **9**) agrupadas mediante celdas combinadas que contienen el encabezado  
     **"RESULTADO APRENDIZAJE-CRITERIO DE EVALUACIÓN PRÁCTICOS"**.

2. **Desarrollo del Mapeo de Celdas**  
   - Se implementó una función que recorre los bloques de celdas combinadas y, utilizando expresiones regulares, detecta automáticamente las actividades.  
   - Se reconoce tanto el formato "TAREA X" como "ACTIVIDAD X" y se normaliza a **"TAREA X"** para homogeneizar la nomenclatura.
   - Se construye un diccionario de mapeo que asocia cada tarea a una lista de referencias de celda (por ejemplo, `"TAREA 1": ["D9", "L9", "S9"]`).
   - Se ordenan las referencias de cada tarea en orden ascendente según el índice de columna para que la celda más a la izquierda (primer trimestre) se utilice primero.

3. **Verificación y Lectura de Valores**  
   - Se implementó una interfaz gráfica con Tkinter que permite:
     - Seleccionar el archivo Excel.
     - Ingresar el nombre del alumno.
     - Seleccionar la tarea mediante un combobox, que se rellena automáticamente con los nombres detectados.
     - Elegir el trimestre (1T, 2T o 3T) mediante otro combobox.
   - Al buscar la celda correspondiente, se obtiene la referencia de la celda correcta (basada en la posición del alumno y la tarea) y se lee el valor (la nota) usando openpyxl en modo `data_only`.


## Instalación y Requisitos

- **Python 3.6+**
- **Librerías necesarias:**
  - `pandas`
  - `openpyxl`
  - `tkinter` (incluido en la mayoría de las instalaciones de Python)
  - *(Para futuras integraciones: `google-auth` y `google-api-python-client` para conectar con Google Sheets)*

Para instalar las dependencias, ejecuta:

```bash
pip install -r requirements.txt
```
## Mapeo de Celdas

La aplicación utiliza el contenido textual de los encabezados para delimitar bloques de actividades. Los archivos Excel cuentan con celdas combinadas (generalmente en las filas **7** y **8**) que contienen el texto:

> "RESULTADO APRENDIZAJE-CRITERIO DE EVALUACIÓN PRÁCTICOS"

Dentro de cada bloque, en la fila 9 se encuentran los nombres de las tareas. La función de mapeo:

- Recorre cada bloque detectado..
- Busca celdas que contengan expresiones como "TAREA X" o "ACTIVIDAD X", ignorando diferencias en mayúsculas o pequeños cambios de formato..
- Normaliza el nombre a formato "TAREA X".
- Almacena las referencias de celda correspondientes, ordenándolas de manera ascendente por el índice de columna (de izquierda a derecha).
Dentro de cada bloque, en la fila **9** se encuentran los nombres específicos de las actividades. La función de mapeo recorre cada rango combinado que contenga el encabezado indicado, extrae de la fila **9** el nombre de cada actividad y genera un diccionario en el que se asocia cada actividad con la(s) referencia(s) de celda correspondiente(s).

### Ejemplo del Diccionario de Mapeo

```python
{
  "TAREA 1": ["D9", "L9", "S9"],
  "TAREA 2": ["E9", "M9", "T9"],
  "TAREA 3": ["F9", "N9", "U9"],
  "TAREA 4": ["G9", "V9"],
  "TAREA 5": ["W9"],
  "TAREA 6": ["X9"],
  "TAREA 7": ["Y9"],
  "TAREA 8": ["Z9"]
}

```
## VVerificación del Mapeo y Lectura de Notas
La interfaz gráfica permite verificar la correspondencia entre alumno, tarea y nota:

- Búsqueda del alumno:
Se recorre la columna de alumnos (por defecto la C a partir de la fila 10) para localizar la fila donde se encuentra el nombre ingresado.

- Selección de Tarea y Trimestre:

   - La tarea se selecciona de un combobox con los nombres detectados en el Excel, evitando la necesidad de escribirla manualmente. 
   - El trimestre se selecciona mediante otro combobox (con opciones "1T", "2T" y "3T"), y se utiliza para determinar cuál de las múltiples referencias de celda se debe usar.
- Lectura de la Nota:
Una vez determinada la celda, se vuelve a cargar el workbook en modo data_only para obtener el valor evaluado (la nota) y se muestra en la interfaz.

## Próximos Pasos

Las fases siguientes del proyecto incluyen:

### Integración con la API de Canvas LMS
Se desarrollará un módulo para extraer las calificaciones directamente desde Canvas, utilizando su API.

### Automatización en Google Sheets
Se implementará la conexión con la API de Google Sheets para escribir las notas en las celdas correspondientes, usando las credenciales configuradas en `credentials.json`.

### Orquestación del flujo completo
Se coordinará la extracción de datos, el mapeo de actividades y la actualización en Google Sheets para lograr una solución integral que automatice el proceso de calificación.

## Estructura del Proyecto
```plaintext
canvas_to_sheets/
├── config/
│   ├── settings.py           # Variables globales y configuración
│   ├── credentials.json      # Credenciales de Google Sheets
├── evaluator/
│   ├── __init__.py           # Permite que sea un módulo
│   ├── cell_locator.py       # Lógica para identificar la celda correcta
│   ├── test_locator.py       # Pruebas del localizador de celdas y mapeo
├── main.py                   # Punto de entrada de la aplicación
├── requirements.txt          # Dependencias necesarias
└── README.md                 # Documentación del proyecto (este archivo)
```
# Canvas Grade Exporter

Esta es una herramienta de escritorio desarrollada en Python que proporciona una interfaz gráfica (GUI) para que los educadores puedan conectarse a la API de Canvas, seleccionar cursos y tareas específicas, y descargar las listas de alumnos y sus calificaciones en archivos CSV limpios y listos para usar.

## ✨ Características Principales

- **Interfaz Gráfica Sencilla**: Utiliza una ventana fácil de usar para guiar al usuario a través del proceso.
- **Conexión Segura a Canvas**: Se conecta a la API de Canvas para obtener datos en tiempo real.
- **Selector de Cursos y Tareas**: Carga dinámicamente los cursos y tareas del usuario, permitiendo una selección precisa.
- **Exportación de Datos**: Genera dos archivos CSV principales:
    1.  `alumnos.csv`: Una lista de todos los estudiantes del curso seleccionado.
    2.  `calificaciones.csv`: Un informe de notas de la tarea seleccionada que incluye el ID del alumno, su nombre completo y su puntuación numérica (`score`).
- **Limpieza de Datos**: El script procesa los datos para eliminar información redundante (como la columna `grade`) y combina la información para que sea más útil.

## 🚀 Instalación y Configuración

Sigue estos pasos para poner en marcha la aplicación.

**1. Clona el repositorio**
```bash
git clone <URL_de_tu_repositorio>
cd <nombre_de_la_carpeta>