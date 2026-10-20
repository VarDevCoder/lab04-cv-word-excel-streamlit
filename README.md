# Laboratorio 4 — Ingestión de CVs Word (.docx) y Homologación Excel (.xlsx)

App web en **Streamlit** desarrollada para la Cátedra *Aplicaciones de la Inteligencia Artificial*
(Licenciatura en Análisis de Sistemas, UNIBE — Semestre II 2026).

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

## Instalación y ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

La app se abrirá en `http://localhost:8501`.

## Stack

- [Streamlit](https://streamlit.io/)
- [python-docx](https://python-docx.readthedocs.io/)
- [pandas](https://pandas.pydata.org/) + [openpyxl](https://openpyxl.readthedocs.io/)
