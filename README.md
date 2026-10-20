# Laboratorio 4 — Ingestión de CVs Word (.docx) y Homologación Excel (.xlsx)

Desarrollado para la Cátedra *Aplicaciones de la Inteligencia Artificial*
(Licenciatura en Análisis de Sistemas, UNIBE — Semestre II 2026).

Incluye **dos interfaces equivalentes**:

- `app.py` — App **web** con [Streamlit](https://streamlit.io/) (también dockerizada).
- `gui.py` — App de **escritorio** con `tkinter` (librería estándar de Python, sin dependencias de UI extra).

## ¿Qué hace?

1. **Procesador de CVs en Word**
   - Sube un currículum en `.docx` y lo procesa **en memoria** (`io.BytesIO`, sin guardar archivos temporales en disco) usando `python-docx`.
   - Extrae párrafos y tablas estructuradas del documento.
   - Detecta automáticamente correo electrónico y teléfono con expresiones regulares.
   - Convierte las tablas del CV a `DataFrame`s interactivos de pandas.
   - Muestra una ficha formateada del postulante y permite editar/descargar el texto extraído como `.txt`.

2. **Homologador de Excel** (desafío de extensión del laboratorio)
   - Sube una planilla `.xlsx` con columnas desordenadas o con nombres inconsistentes.
   - Sugiere automáticamente el mapeo hacia un esquema estándar (nombre, email, teléfono, cargo, años de experiencia).
   - Permite ajustar manualmente el mapeo y descargar la planilla ya homologada en `.xlsx`.

## Requisitos

- Python 3.10+

## Opción A — App web (Streamlit)

```bash
pip install -r requirements.txt
streamlit run app.py
```

La app se abrirá en `http://localhost:8501`.

### Con Docker

```bash
docker compose up --build
```

o manualmente:

```bash
docker build -t lab04-cv-excel .
docker run -p 8501:8501 lab04-cv-excel
```

Luego abrí `http://localhost:8501` en tu navegador.

## Opción B — App de escritorio (Tkinter)

Interfaz gráfica nativa equivalente, sin servidor ni navegador:

```bash
pip install -r requirements.txt
python gui.py
```

> La GUI de escritorio usa `tkinter` (incluido en la instalación estándar de Python) y no corre dentro del contenedor Docker, ya que ese requiere un entorno gráfico local.

## Stack

- [Streamlit](https://streamlit.io/) (app web)
- [tkinter](https://docs.python.org/3/library/tkinter.html) (app de escritorio)
- [python-docx](https://python-docx.readthedocs.io/)
- [pandas](https://pandas.pydata.org/) + [openpyxl](https://openpyxl.readthedocs.io/)
- [Docker](https://www.docker.com/) (para la app web)
