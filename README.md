Este archivo explica qué hace el proyecto, cómo configurarlo y su estado actual.

**Instrucciones:**
1.  En la raíz de tu proyecto, crea un archivo llamado `README.md`.
2.  Copia y pega el siguiente texto en ese archivo.

```markdown
# Evaluador Canvas-Excel/Sheets

Este proyecto es una aplicación de escritorio con interfaz gráfica (GUI) desarrollada en Python y Tkinter, diseñada para automatizar el proceso de transferencia de calificaciones desde la plataforma de e-learning **Canvas LMS** a hojas de cálculo de evaluación, ya sean archivos de **Excel locales** o documentos de **Google Sheets** en la nube.

La aplicación permite a los docentes seleccionar un curso y una tarea de Canvas, y un archivo y una tarea de destino, para luego cotejar las listas de alumnos y escribir las notas correspondientes en las celdas correctas.

## Características Principales

- **Conexión con Canvas LMS:** Se conecta a la API de Canvas para obtener la lista de cursos, tareas y calificaciones de los alumnos de forma dinámica.
- **Doble Destino de Datos:** Soporta tanto archivos de Microsoft Excel (`.xlsx`) locales como hojas de cálculo de Google Sheets como destino para las notas.
- **Mapeo Inteligente de Tareas:** Analiza la estructura de la hoja de cálculo de destino para identificar automáticamente la ubicación de las diferentes tareas y trimestres.
- **Cotejo de Alumnos Flexible:** Utiliza un algoritmo de normalización de nombres para encontrar coincidencias entre las listas de alumnos de Canvas y del archivo de destino, incluso si los formatos no son idénticos.
- **Procesamiento Aislado:** La lógica de la aplicación está separada en módulos claros:
    - `gui.py`: Interfaz gráfica.
    - `clients.py`: Comunicación con las APIs externas.
    - `mapping.py`: Análisis de la estructura de las hojas de cálculo.
    - `matcher.py`: Algoritmos de comparación de nombres.
    - `processor.py`: Orquestación del proceso de escritura.
- **Generación de Archivos Intermedios:** Guarda las listas de alumnos y notas extraídas en archivos `.json` para facilitar la depuración y la verificación del flujo de datos.
- **Logging de Actividad:** Registra todas las operaciones importantes en un archivo `app.log`.

## Estado del Proyecto

Actualmente, el proyecto se encuentra en una fase funcional con las siguientes características operativas:
- **Integración con Canvas:** Totalmente funcional.
- **Integración con Google Sheets:** Totalmente funcional, incluyendo lectura, mapeo y escritura de notas.
- **Integración con Excel Local:** La carga y el mapeo de la plantilla funcionan correctamente. La escritura de notas está implementada.

### Problema Conocido (Aparcado)
- **Lectura de Fórmulas en Excel Local:** Existe un problema conocido con la librería `openpyxl`. Cuando se intenta leer un archivo Excel local donde los nombres de los alumnos están en celdas que contienen fórmulas (ej: `=OTRA_HOJA!A1`), `openpyxl` no puede resolver el valor de la fórmula si el archivo no fue guardado previamente por Microsoft Excel. Esto causa que el cotejo de alumnos falle en el modo "Excel Local" con este tipo de plantillas. La funcionalidad con Google Sheets no se ve afectada.

## Configuración y Puesta en Marcha

### Prerrequisitos
- Python 3.10 o superior.
- Una cuenta de Canvas LMS con una clave de API generada.
- Una cuenta de Google y un proyecto en Google Cloud con la API de Google Sheets habilitada y credenciales de cuenta de servicio.

### Pasos de Instalación

1.  **Clonar el repositorio:**
    ```bash
    git clone https://github.com/oskarmg77/canvas_excel_sheetevaluation.git
    cd canvas_excel_sheetevaluation
    ```

2.  **Crear y activar un entorno virtual:**
    ```bash
    python -m venv .venv
    # En Windows
    .\.venv\Scripts\activate
    # En macOS/Linux
    source .venv/bin/activate
    ```

3.  **Instalar las dependencias:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Configurar las credenciales (¡IMPORTANTE!):**
    - En la carpeta `config`, renombra `settings.py.example` a `settings.py`.
    - Edita `settings.py` y añade tu `API_URL` y `API_KEY` de Canvas.
    - Descarga el archivo de credenciales JSON de tu cuenta de servicio de Google Cloud.
    - Renombra el archivo a `service_account.json` y colócalo dentro de la carpeta `config`.
    
    **Nota:** Los archivos de credenciales están correctamente listados en `.gitignore` para evitar que se suban al repositorio.

### Cómo Ejecutar la Aplicación

Una vez configurado, ejecuta el siguiente comando desde la carpeta raíz del proyecto:

```bash
python main.py