"""
Laboratorio 4: App Web en Python (Streamlit) para Ingestión de CVs Word (.docx)
y Homologación de Excel (.xlsx)

Cátedra: Aplicaciones de la Inteligencia Artificial
Licenciatura en Análisis de Sistemas - UNIBE
"""

import io
import re

import pandas as pd
import streamlit as st
from docx import Document

st.set_page_config(
    page_title="Lector de CVs Word & Homologador Excel",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Ingestión y Procesamiento de CVs en Word (.docx)")
st.caption(
    "Laboratorio 4 · Unidad II · Cátedra de Aplicaciones de la Inteligencia Artificial"
)

tab_cv, tab_excel = st.tabs(["📄 Procesador de CVs (Word)", "📊 Homologador de Excel"])

# ---------------------------------------------------------------------------
# TAB 1 — Ingestión y procesamiento de CVs en Word (.docx)
# ---------------------------------------------------------------------------
with tab_cv:
    archivo_subido = st.file_uploader(
        "Selecciona un CV en formato Word (.docx)", type=["docx"]
    )

    if archivo_subido:
        # Paso 1: Ingestión directa en memoria (segura, sin archivos temporales en disco)
        doc_bytes = io.BytesIO(archivo_subido.read())
        doc = Document(doc_bytes)
        parrafos = [p.text.strip() for p in doc.paragraphs if p.text.strip()]

        st.success(
            f"✅ Documento procesado: {len(parrafos)} párrafos y "
            f"{len(doc.tables)} tablas encontradas."
        )

        # Paso 2: Extracción heurística de entidades (regex) y tablas estructuradas
        texto_completo = "\n".join(parrafos)

        email_match = re.search(
            r"[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+", texto_completo
        )
        telefono_match = re.search(
            r"(\+?\d{1,4}[\s-]?)?(\(?\d{2,4}\)?[\s-]?)?\d{3,4}[\s-]?\d{3,4}",
            texto_completo,
        )

        col1, col2 = st.columns(2)
        with col1:
            st.metric(
                "📧 Correo Detectado",
                email_match.group(0) if email_match else "No detectado",
            )
        with col2:
            st.metric(
                "📱 Teléfono Detectado",
                telefono_match.group(0) if telefono_match else "No detectado",
            )

        if doc.tables:
            st.subheader("📊 Tablas Encontradas en el Documento")
            for idx, tabla in enumerate(doc.tables):
                filas = [
                    [celda.text.strip() for celda in fila.cells] for fila in tabla.rows
                ]
                if len(filas) > 1:
                    df_tabla = pd.DataFrame(filas[1:], columns=filas[0])
                    st.markdown(f"**Tabla {idx + 1}**")
                    st.dataframe(df_tabla, use_container_width=True)

        # Paso 3: Ficha de candidato, texto editable y exportación
        tab1, tab2 = st.tabs(["📋 Ficha Formateada", "✏️ Texto Editable & Descarga"])

        with tab1:
            st.markdown("### Perfil del Postulante")
            for p in parrafos[:6]:
                if p.isupper():
                    st.markdown(f"#### 📌 {p}")
                else:
                    st.write(f"• {p}")

        with tab2:
            texto_editado = st.text_area(
                "Texto extraído (editable para notas de entrevista):",
                value=texto_completo,
                height=250,
            )
            st.download_button(
                label="💾 Descargar CV Procesado (.txt)",
                data=texto_editado,
                file_name="cv_procesado.txt",
                mime="text/plain",
            )
    else:
        st.info("⬆️ Sube un archivo .docx para comenzar el procesamiento.")

# ---------------------------------------------------------------------------
# TAB 2 — Desafío de extensión: Homologador de columnas Excel
# ---------------------------------------------------------------------------
with tab_excel:
    st.subheader("📊 Mapeador y Homologador de Columnas Excel (.xlsx)")
    st.write(
        "Sube una planilla con nombres de columna desordenados o inconsistentes "
        "y homológala a un esquema estándar antes de descargarla."
    )

    ESQUEMA_ESTANDAR = {
        "nombre": ["nombre", "nombre completo", "apellido y nombre", "candidato", "name"],
        "email": ["email", "correo", "correo electronico", "e-mail", "mail"],
        "telefono": ["telefono", "teléfono", "celular", "phone", "contacto"],
        "cargo": ["cargo", "puesto", "posicion", "posición", "role", "position"],
        "experiencia_anios": [
            "experiencia",
            "años experiencia",
            "anios experiencia",
            "experience",
            "años",
        ],
    }

    def normalizar(texto: str) -> str:
        return re.sub(r"[^a-z0-9]", "", str(texto).strip().lower())

    def sugerir_columna(col_original: str) -> str | None:
        col_norm = normalizar(col_original)
        for campo_estandar, alias in ESQUEMA_ESTANDAR.items():
            for a in alias:
                if normalizar(a) == col_norm:
                    return campo_estandar
        return None

    archivo_excel = st.file_uploader(
        "Selecciona una planilla Excel (.xlsx)", type=["xlsx"], key="excel_uploader"
    )

    if archivo_excel:
        excel_bytes = io.BytesIO(archivo_excel.read())
        df_original = pd.read_excel(excel_bytes, engine="openpyxl")

        st.markdown("**Vista previa de los datos originales:**")
        st.dataframe(df_original, use_container_width=True)

        st.markdown("**Mapeo de columnas:**")
        st.caption(
            "Revisa las sugerencias automáticas y ajusta manualmente las columnas "
            "que no coincidan con el esquema estándar."
        )

        opciones_destino = ["(Ignorar columna)"] + list(ESQUEMA_ESTANDAR.keys())
        mapeo = {}

        for columna in df_original.columns:
            sugerencia = sugerir_columna(columna)
            indice_default = (
                opciones_destino.index(sugerencia) if sugerencia in opciones_destino else 0
            )
            seleccion = st.selectbox(
                f"Columna original: «{columna}»",
                options=opciones_destino,
                index=indice_default,
                key=f"map_{columna}",
            )
            if seleccion != "(Ignorar columna)":
                mapeo[columna] = seleccion

        if st.button("🔄 Homologar y Generar Planilla Estándar"):
            if mapeo:
                df_homologado = df_original[list(mapeo.keys())].rename(columns=mapeo)
                st.success("✅ Planilla homologada correctamente.")
                st.dataframe(df_homologado, use_container_width=True)

                buffer_salida = io.BytesIO()
                df_homologado.to_excel(buffer_salida, index=False, engine="openpyxl")
                buffer_salida.seek(0)

                st.download_button(
                    label="💾 Descargar Excel Homologado (.xlsx)",
                    data=buffer_salida,
                    file_name="planilla_homologada.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            else:
                st.warning("⚠️ Selecciona al menos una columna para homologar.")
    else:
        st.info("⬆️ Sube un archivo .xlsx para comenzar la homologación.")
