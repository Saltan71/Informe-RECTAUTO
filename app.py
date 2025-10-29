import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import io
import zipfile
from fpdf import FPDF
import matplotlib.pyplot as plt
import os

FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"
ESTADOS_PENDIENTES = ["Abierto"]

st.set_page_config(page_title="Informe Rectauto", layout="wide")
st.title("📊 Seguimiento Equipo Regional RECTAUTO")

class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 8)
        self.cell(0, 10, 'Informe de Expedientes Pendientes', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 6)
        self.cell(0, 10, f'Página {self.page_no()}', 0, 0, 'C')

def dataframe_to_pdf_bytes(df, title):
    pdf = PDF('L', 'mm', 'A4')
    pdf.add_page()
    pdf.set_font("Arial", "B", 8)
    pdf.cell(0, 10, title, 0, 1, 'C')
    pdf.ln(5)

    col_widths = [43, 14, 14, 8, 24, 14, 14, 24, 14, 40, 24, 14, 26]
    df_mostrar_pdf = df.iloc[:, :len(col_widths)]

    ALTURA_ENCABEZADO = 11

    def imprimir_encabezados():
        pdf.set_font("Arial", "B", 5)
        pdf.set_fill_color(200, 220, 255)
        y_inicio = pdf.get_y()
        for i, header in enumerate(df_mostrar_pdf.columns):
            x = pdf.get_x()
            y = pdf.get_y()
            pdf.cell(col_widths[i], ALTURA_ENCABEZADO, "", 1, 0, 'C', True)
            pdf.set_xy(x, y)
            texto = str(header)
            ancho_texto = pdf.get_string_width(texto)
            if ancho_texto <= col_widths[i] - 2:
                altura_texto = 3
                y_pos = y + (ALTURA_ENCABEZADO - altura_texto) / 2
                pdf.set_xy(x, y_pos)
                pdf.cell(col_widths[i], altura_texto, texto, 0, 0, 'C')
            else:
                pdf.set_xy(x, y + 1)
                pdf.multi_cell(col_widths[i], 2.5, texto, 0, 'C')
            pdf.set_xy(x + col_widths[i], y)
        pdf.set_xy(pdf.l_margin, y_inicio + ALTURA_ENCABEZADO)

    imprimir_encabezados()

    pdf.set_font("Arial", "", 7)
    for _, row in df_mostrar_pdf.iterrows():
        if pdf.get_y() + 6 > 190:
            pdf.add_page()
            imprimir_encabezados()
        for i, col_data in enumerate(row):
            text = str(col_data).replace("\n", " ")
            pdf.cell(col_widths[i], 6, text, 1, 0, 'L')
        pdf.ln()

    pdf_output = pdf.output(dest='B')
    return pdf_output

archivo = st.file_uploader("📁 Sube el archivo Excel (rectauto*.xlsx)", type=["xlsx", "xls"])

if archivo:
    df = pd.read_excel(archivo, sheet_name=HOJA, header=0, engine="openpyxl" if archivo.name.endswith("xlsx") else "xlrd")
    df.columns = [col.upper() for col in df.columns]
    columnas = [0, 1, 2, 3, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]
    columna_fecha = df.columns[10]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Sin fecha"
    st.subheader(f"📅 Semana {num_semana} a {fecha_max_str}")

    # === NUEVAS 12 COLUMNAS ===
    st.markdown("### 🧮 Generando columnas adicionales...")
    df_enriquecido = df.copy()
    df_enriquecido["DIAS_DESDE_REFERENCIA"] = (df_enriquecido[columna_fecha] - FECHA_REFERENCIA).dt.days
    df_enriquecido["DIAS_HASTA_MAX"] = (fecha_max - df_enriquecido[columna_fecha]).dt.days
    df_enriquecido["SEMANA_EXPEDIENTE"] = ((df_enriquecido[columna_fecha] - FECHA_REFERENCIA).dt.days // 7 + 1)
    df_enriquecido["ANTIGUEDAD_MESES"] = df_enriquecido["DIAS_DESDE_REFERENCIA"] / 30.4
    df_enriquecido["ES_RECIENTE"] = df_enriquecido["DIAS_HASTA_MAX"] < 7
    df_enriquecido["PENDIENTE"] = df_enriquecido["ESTADO"].isin(ESTADOS_PENDIENTES)
    df_enriquecido["EQUIPO_USUARIO"] = df_enriquecido["EQUIPO"] + " - " + df_enriquecido["USUARIO"]
    df_enriquecido["FECHA_REFERENCIA_STR"] = FECHA_REFERENCIA.strftime("%d/%m/%Y")
    df_enriquecido["FECHA_MAX_STR"] = fecha_max_str
    df_enriquecido["SEMANA_ACTUAL"] = num_semana
    df_enriquecido["DIFERENCIA_DIAS_REF_MAX"] = (fecha_max - FECHA_REFERENCIA).days
    df_enriquecido["RATIO_TIEMPO"] = df_enriquecido["DIAS_DESDE_REFERENCIA"] / df_enriquecido["DIFERENCIA_DIAS_REF_MAX"]

    df = df_enriquecido
