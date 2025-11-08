import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from datetime import datetime, timedelta
import io
import zipfile
from fpdf import FPDF
import matplotlib.pyplot as plt
import os
import hashlib
import tempfile
import shutil
import uuid
import getpass
from PIL import Image


# === NUEVA CLASE PARA ENTORNO DE USUARIO ===
class UserEnvironment:
    def __init__(self):
        try:
            self.username = getpass.getuser()
        except:
            self.username = "unknown_user"
        self.session_id = f"{self.username}_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}"
        self.working_dir = os.path.join(tempfile.gettempdir(), f"rectauto_{self.session_id}")
        os.makedirs(self.working_dir, exist_ok=True)
        print(f"User environment created: {self.working_dir}")  # Para debugging
        
    def get_cache_key(self, base_key):
        return f"{self.session_id}_{base_key}"
    
    def get_temp_path(self, filename):
        return os.path.join(self.working_dir, filename)
    
    def cleanup(self):
        """Limpia los archivos temporales (opcional)"""
        try:
            shutil.rmtree(self.working_dir)
        except:
            pass

# Inicializar el entorno de usuario
user_env = UserEnvironment()

# Constantes
FECHA_REFERENCIA = datetime(2022, 11, 1)
HOJA = "Sheet1"
ESTADOS_PENDIENTES = ["Abierto"]
CACHE_TTL = 7200  # 2 horas en segundos

# Test file en directorio único por usuario
test_file = user_env.get_temp_path("test_write_access.tmp")
with open(test_file, 'w') as f:
    f.write("test")

st.set_page_config(page_title="Informe Rectauto", layout="wide", page_icon=Image.open("icono.ico"))
st.title("📊 Seguimiento Equipo Regional RECTAUTO")

# Inicializar variables de sesión para documentación
if 'mostrar_descarga' not in st.session_state:
    st.session_state.mostrar_descarga = False
if 'documentos_actualizados' not in st.session_state:
    st.session_state.documentos_actualizados = None
if 'cambios_documentacion' not in st.session_state:
    st.session_state.cambios_documentacion = {}

# Clase PDF (mejorada con ordenamiento)
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 8)
        self.cell(0, 5, 'Informe de Expedientes Pendientes', 0, 1, 'C')
        self.ln(5)

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 6)
        self.cell(0, 5, f'Página {self.page_no()}', 0, 0, 'C')
    
    def aplicar_formato_condicional_pdf(self, df_original, idx, col_name, col_width, altura_fila, x, y):
        """Aplica formato condicional a celdas específicas en el PDF con NUEVAS condiciones"""
        try:
            if idx >= len(df_original):
                return
                
            fila = df_original.iloc[idx]
            
            # Condición 1: USUARIO vs USUARIO-CSV
            if col_name == 'USUARIO-CSV':
                usuario_principal = fila.get('USUARIO', '')
                usuario_csv = fila.get('USUARIO-CSV', '')
                
                # Comparar los valores (manejar NaN y tipos diferentes)
                if pd.notna(usuario_principal) and pd.notna(usuario_csv):
                    if str(usuario_principal).strip() != str(usuario_csv).strip():
                        # Fondo rojo cuando USUARIO es distinto de USUARIO-CSV
                        self.set_fill_color(255, 0, 0)
                        self.rect(x, y, col_width, altura_fila, 'F')
                        self.set_fill_color(255, 255, 255)
            
            # Condición 2: RUE con MÚLTIPLES condiciones específicas
            elif col_name == 'RUE':
                etiq_penultimo = fila.get('ETIQ. PENÚLTIMO TRAM.', '')
                fecha_notif = fila.get('FECHA NOTIFICACIÓN', None)
                docum_incorp = fila.get('DOCUM.INCORP.', '')
                
                es_amarillo = False
                
                # CONDICIÓN 2.1: "80 PROPRES" con fecha límite superada
                if (str(etiq_penultimo).strip() == "80 PROPRES" and 
                    pd.notna(fecha_notif) and 
                    isinstance(fecha_notif, (pd.Timestamp, datetime))):
                    
                    fecha_limite = fecha_notif + timedelta(days=23)
                    if datetime.now() > fecha_limite:
                        es_amarillo = True
                
                # NUEVA CONDICIÓN 2.2: "50 REQUERIR" con fecha límite superada
                elif (str(etiq_penultimo).strip() == "50 REQUERIR" and 
                    pd.notna(fecha_notif) and 
                    isinstance(fecha_notif, (pd.Timestamp, datetime))):
                    
                    fecha_limite = fecha_notif + timedelta(days=23)
                    if datetime.now() > fecha_limite:
                        es_amarillo = True
                
                # NUEVA CONDICIÓN 2.3: "70 ALEGACI" o "60 CONTESTA"
                elif str(etiq_penultimo).strip() in ["70 ALEGACI", "60 CONTESTA"]:
                    es_amarillo = True
                
                # NUEVA CONDICIÓN 2.4: DOCUM.INCORP. no vacío Y distinto de "SOLICITUD" o "REITERA SOLICITUD"
                elif (pd.notna(docum_incorp) and 
                    str(docum_incorp).strip() != '' and
                    str(docum_incorp).strip().upper() not in ["SOLICITUD", "REITERA SOLICITUD"]):
                    es_amarillo = True
                
                if es_amarillo:
                    # Fondo amarillo cuando se cumple alguna condición
                    self.set_fill_color(255, 255, 0)
                    self.rect(x, y, col_width, altura_fila, 'F')
                    self.set_fill_color(255, 255, 255)
            
            # Condición 3: Resaltar DOCUM.INCORP. cuando tiene valor
            elif col_name == 'DOCUM.INCORP.':
                docum_incorp = fila.get('DOCUM.INCORP.', '')
                if pd.notna(docum_incorp) and str(docum_incorp).strip() != '':
                    # Fondo azul claro cuando hay documentación incorporada
                    self.set_fill_color(173, 216, 230)
                    self.rect(x, y, col_width, altura_fila, 'F')
                    self.set_fill_color(255, 255, 255)
                        
        except Exception as e:
            # Silenciar errores para no interrumpir la generación del PDF
            pass

# === CLASE PDFResumenKPI CORREGIDA (CON FUENTE QUE SOPORTA CARACTERES ESPECIALES) ===
class PDFResumenKPI(FPDF):
    def __init__(self):
        super().__init__()
        # Agregar soporte para caracteres latinos
        self.add_page()
        
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Resumen de KPIs Semanales', 0, 1, 'C')
        self.ln(5)
    
    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Pagina {self.page_no()}', 0, 0, 'C')
    
    def add_section_title(self, title):
        self.set_font('Arial', 'B', 10)
        # Reemplazar caracteres problemáticos
        title_safe = title.replace('Í', 'I').replace('É', 'E').replace('Á', 'A').replace('Ó', 'O').replace('Ú', 'U')
        self.cell(0, 8, title_safe, 0, 1, 'L')
        self.ln(2)
    
    def add_metric(self, label, value, explanation=""):
        # Limpiar caracteres especiales de las etiquetas
        label_safe = label.replace('Í', 'I').replace('É', 'E').replace('Á', 'A').replace('Ó', 'O').replace('Ú', 'U')
        
        self.set_font('Arial', 'B', 9)
        self.cell(60, 6, label_safe, 0, 0)
        self.set_font('Arial', '', 9)
        
        # Limpiar también el valor si es texto
        if isinstance(value, str):
            value_safe = value.replace('Í', 'I').replace('É', 'E').replace('Á', 'A').replace('Ó', 'O').replace('Ú', 'U')
        else:
            value_safe = value
            
        self.cell(40, 6, str(value_safe), 0, 1)
        if explanation:
            self.set_font('Arial', 'I', 8)
            explanation_safe = explanation.replace('Í', 'I').replace('É', 'E').replace('Á', 'A').replace('Ó', 'O').replace('Ú', 'U')
            self.cell(0, 4, explanation_safe, 0, 1)
            self.ln(1)

# Funciones optimizadas con cache
@st.cache_data(ttl=CACHE_TTL, show_spinner="Procesando archivo Excel...")
def cargar_y_procesar_rectauto(archivo, _user_key=user_env.session_id):
    """Carga y procesa el archivo RECTAUTO con cache de 2 horas"""
    df = pd.read_excel(
        archivo, 
        sheet_name=HOJA, 
        header=0, 
        thousands='.', 
        decimal=',', 
        engine="openpyxl" if archivo.name.endswith("xlsx") else "xlrd"
    )
    df.columns = [col.upper().strip() for col in df.columns]
    columnas = [0, 1, 2, 3, 4, 5, 6, 12, 14, 15, 16, 17, 18, 20, 21, 23, 26, 27]
    df = df.iloc[:, columnas]
    return df

@st.cache_data(ttl=CACHE_TTL)
def cargar_y_procesar_notifica(archivo, _user_key=user_env.session_id):
    """Carga y procesa el archivo NOTIFICA"""
    try:
        df = pd.read_excel(archivo, sheet_name=HOJA)
        df.columns = [col.upper().strip() for col in df.columns]
        # Ordenar por RUE ORIGEN (ascendente) y FECHA APERTURA (descendente)
        if 'RUE ORIGEN' in df.columns and 'FECHA APERTURA' in df.columns:
            df['FECHA APERTURA'] = pd.to_datetime(df['FECHA APERTURA'], errors='coerce')
            df = df.sort_values(['RUE ORIGEN', 'FECHA APERTURA'], ascending=[True, False])
        
        # Mantener solo columnas relevantes
        columnas_a_mantener = ['RUE ORIGEN', 'FECHA NOTIFICACIÓN']
        columnas_existentes = [col for col in columnas_a_mantener if col in df.columns]
        df = df[columnas_existentes]
        
        return df
    except Exception as e:
        st.error(f"Error procesando NOTIFICA: {e}")
        return None

@st.cache_data(ttl=CACHE_TTL)
def cargar_y_procesar_triaje(archivo, _user_key=user_env.session_id):
    """Carga y procesa el archivo TRIAJE"""
    try:
        df = pd.read_excel(archivo, sheet_name='Triaje')
        df.columns = [col.upper().strip() for col in df.columns]
        
        # Crear RUE a partir de las primeras 4 columnas
        if df.shape[1] >= 4:
            # Formatear la cuarta columna a 6 dígitos
            df['RUE_TEMP'] = df.iloc[:, 3].astype(str).str.zfill(6)
            
            # Concatenar las cuatro primeras columnas
            df['RUE'] = (
                df.iloc[:, 0].astype(str) + 
                df.iloc[:, 1].astype(str) + 
                df.iloc[:, 2].astype(str) + 
                df['RUE_TEMP']
            )
            
            # Mantener solo columnas relevantes
            columnas_a_mantener = ['USUARIO-CSV', 'CALIFICACIÓN', 'OBSERVACIONES', 'FECHA ASIG']
            # Normalizar también los nombres de las columnas a mantener
            columnas_a_mantener = [col.upper().strip() for col in columnas_a_mantener]
            columnas_existentes = [col for col in columnas_a_mantener if col in df.columns]
            df = df[['RUE'] + columnas_existentes]
            
            return df
        else:
            st.warning("TRIAJE no tiene al menos 4 columnas")
            return None
    except Exception as e:
        st.error(f"Error procesando TRIAJE: {e}")
        return None

@st.cache_data(ttl=CACHE_TTL)
def cargar_y_procesar_usuarios(archivo, _user_key=user_env.session_id):
    """Carga y procesa el archivo USUARIOS"""
    try:
        df = pd.read_excel(archivo, sheet_name=HOJA)
        df.columns = [col.upper().strip() for col in df.columns]
        return df
    except Exception as e:
        st.error(f"Error procesando USUARIOS: {e}")
        return None

@st.cache_data(ttl=CACHE_TTL)
def cargar_y_procesar_documentos(archivo, _user_key=user_env.session_id):
    """Carga y procesa el archivo DOCUMENTOS"""
    try:
        # Cargar hoja DOCU para los valores del desplegable
        df_docu = pd.read_excel(archivo, sheet_name='DOCU')
        opciones_docu = df_docu.iloc[:, 0].dropna().tolist()
        
        # Cargar hoja DOCUMENTOS para los valores guardados
        df_documentos = pd.read_excel(archivo, sheet_name='DOCUMENTOS')
        df_documentos.columns = [col.upper().strip() for col in df_documentos.columns]
        
        return {
            'opciones': opciones_docu,
            'documentos': df_documentos,
            'archivo': archivo
        }
    except Exception as e:
        st.error(f"Error procesando DOCUMENTOS: {e}")
        return None

def guardar_documentos_actualizados(archivo_original, df_documentos_actualizado):
    """Guarda los datos actualizados en el archivo DOCUMENTOS.xlsx"""
    try:
        # Usar directorio temporal único por usuario
        tmp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=user_env.working_dir)
        tmp_path = tmp_file.name
        tmp_file.close()  # 🔹 Muy importante: libera el archivo en Windows

        # Escribir las dos hojas en el archivo temporal
        with pd.ExcelWriter(tmp_path, engine="openpyxl") as writer:
            # Hoja DOCUMENTOS actualizada - SOLO los registros actuales
            df_documentos_actualizado.to_excel(writer, sheet_name="DOCUMENTOS", index=False)

            # Hoja DOCU - MANTENER las opciones del desplegable del original
            archivo_original.seek(0)
            df_docu_original = pd.read_excel(archivo_original, sheet_name="DOCU")
            df_docu_original.to_excel(writer, sheet_name="DOCU", index=False)

        # Leer contenido ya guardado
        with open(tmp_path, "rb") as f:
            contenido = f.read()

        # Eliminar el archivo temporal (libre ya de bloqueos)
        os.remove(tmp_path)

        return contenido

    except Exception as e:
        st.error(f"Error guardando DOCUMENTOS: {e}")
        return None

@st.cache_data(ttl=CACHE_TTL)
def combinar_archivos(rectauto_df, notifica_df=None, triaje_df=None, usuarios_df=None, documentos_data=None, _user_key=user_env.session_id):
    """Combina los archivos en un único DataFrame incluyendo DOCUM.INCORP."""
    df_combinado = rectauto_df.copy()
    
    # Combinar con NOTIFICA
    if notifica_df is not None and 'RUE ORIGEN' in notifica_df.columns:
        # Tomar solo la última notificación por RUE ORIGEN (debido al ordenamiento previo)
        notifica_ultima = notifica_df.drop_duplicates(subset=['RUE ORIGEN'], keep='first')
        df_combinado = pd.merge(
            df_combinado, 
            notifica_ultima, 
            left_on='RUE', 
            right_on='RUE ORIGEN', 
            how='left'
        )
        # Eliminar la columna RUE ORIGEN ya que ya tenemos RUE
        if 'RUE ORIGEN' in df_combinado.columns:
            df_combinado.drop('RUE ORIGEN', axis=1, inplace=True)
        st.sidebar.info(f"✅ NOTIFICA combinado: {len(notifica_ultima)} registros")
    
    # Combinar con TRIAJE
    if triaje_df is not None and 'RUE' in triaje_df.columns:
        df_combinado = pd.merge(
            df_combinado, 
            triaje_df, 
            on='RUE', 
            how='left'
        )
        st.sidebar.info(f"✅ TRIAJE combinado: {len(triaje_df)} registros")
    
    # Añadir columna DOCUM.INCORP. después de FECHA NOTIFICACIÓN
    if 'FECHA NOTIFICACIÓN' in df_combinado.columns:
        # Buscar la posición de FECHA NOTIFICACIÓN
        columnas = df_combinado.columns.tolist()
        pos_fecha_notif = columnas.index('FECHA NOTIFICACIÓN')
        
        # Insertar DOCUM.INCORP. después de FECHA NOTIFICACIÓN
        columnas.insert(pos_fecha_notif + 1, 'DOCUM.INCORP.')
        df_combinado = df_combinado.reindex(columns=columnas)
        
        # Si tenemos datos de documentación, rellenar los valores
        if documentos_data is not None and not documentos_data['documentos'].empty:
            df_documentos = documentos_data['documentos']
            # Crear un mapeo RUE -> DOCUMENTACION
            mapeo_documentos = df_documentos.set_index('RUE')['DOCUM.INCORP.'].to_dict()
            # Aplicar el mapeo a la nueva columna
            df_combinado['DOCUM.INCORP.'] = df_combinado['RUE'].map(mapeo_documentos)
    
    return df_combinado

@st.cache_data(ttl=CACHE_TTL)
def convertir_fechas(df, _user_key=user_env.session_id):
    """Convierte columnas con 'FECHA' en el nombre a datetime"""
    for col in df.columns:
        if 'FECHA' in col:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

@st.cache_data(ttl=CACHE_TTL)
def dataframe_to_pdf_bytes(df_mostrar, title, df_original):
    """Genera un PDF desde un DataFrame con formato condicional (compatible con fpdf2 y Windows)."""
    try:
        # Crear el PDF usando tu clase personalizada (hereda de FPDF)
        pdf = PDF('L', 'mm', 'A4')
        pdf.add_page()
        pdf.set_font("Arial", "B", 8)
        pdf.cell(0, 5, title, 0, 1, 'C')
        pdf.ln(5)

        # --- Definición de columnas REDUCIDAS ---
        # Anchos reducidos para ETIQ. PENÚLTIMO TRAM. y ETIQ. ÚLTIMO TRAM.
        # Eliminada FECHA DE ACTUALIZACIÓN DATOS
        col_widths = [28, 11, 11, 8, 16, 11, 11, 16, 11, 20, 20, 9, 18, 11, 14, 9, 24, 20, 11]  # 19 columnas en lugar de 20
        
        # Ajustar si hay menos columnas
        if len(df_mostrar.columns) < len(col_widths):
            col_widths = col_widths[:len(df_mostrar.columns)]
        elif len(df_mostrar.columns) > len(col_widths):
            col_widths.extend([18] * (len(df_mostrar.columns) - len(col_widths)))

        ALTURA_ENCABEZADO = 13
        ALTURA_LINEA = 3
        ALTURA_BASE_FILA = 2

        # --- Encabezados de tabla ---
        def imprimir_encabezados():
            pdf.set_font("Arial", "", 5)
            pdf.set_fill_color(200, 220, 255)
            y_inicio = pdf.get_y()

            for i, header in enumerate(df_mostrar.columns):
                # EXCLUIR columna "FECHA DE ACTUALIZACIÓN DATOS"
                if "FECHA DE ACTUALIZACIÓN DATOS" in header.upper():
                    continue
                    
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
        pdf.set_font("Arial", "", 5)

        # --- Filas de datos ---
        for idx, (_, row) in enumerate(df_mostrar.iterrows()):
            max_lineas = 1
            for col_idx, (col_name, col_data) in enumerate(zip(df_mostrar.columns, row)):
                # EXCLUIR columna "FECHA DE ACTUALIZACIÓN DATOS"
                if "FECHA DE ACTUALIZACIÓN DATOS" in col_name.upper():
                    continue
                    
                texto = str(col_data).replace("\n", " ")
                # REEMPLAZAR "nan" por vacío
                if texto.lower() == "nan" or texto.strip() == "":
                    texto = ""
                    
                if not texto.strip():
                    continue
                ancho_disponible = min(col_widths) - 2
                ancho_texto = pdf.get_string_width(texto)
                if ancho_texto > ancho_disponible:
                    lineas_necesarias = max(1, int(ancho_texto / ancho_disponible) + 1)
                    max_lineas = max(max_lineas, lineas_necesarias)

            altura_fila = ALTURA_BASE_FILA + ((max_lineas - 1) * ALTURA_LINEA) / 2

            # Saltar de página si es necesario
            if pdf.get_y() + altura_fila > 190:
                pdf.add_page()
                imprimir_encabezados()

            x_inicio = pdf.get_x()
            y_inicio = pdf.get_y()

            # Bordes de las celdas (excluyendo FECHA DE ACTUALIZACIÓN DATOS)
            ancho_total = 0
            for i, col_name in enumerate(df_mostrar.columns):
                if "FECHA DE ACTUALIZACIÓN DATOS" in col_name.upper():
                    continue
                ancho_total += col_widths[i]
                pdf.rect(x_inicio + sum(col_widths[:i]), y_inicio, col_widths[i], altura_fila)

            # Contenido con formato condicional (excluyendo FECHA DE ACTUALIZACIÓN DATOS)
            col_idx_visible = 0
            for i, (col_name, col_data) in enumerate(zip(df_mostrar.columns, row)):
                # EXCLUIR columna "FECHA DE ACTUALIZACIÓN DATOS"
                if "FECHA DE ACTUALIZACIÓN DATOS" in col_name.upper():
                    continue
                    
                texto = str(col_data).replace("\n", " ")
                # REEMPLAZAR "nan" por vacío
                if texto.lower() == "nan" or texto.strip() == "":
                    texto = ""
                    
                x_celda = x_inicio + sum(col_widths[:col_idx_visible])
                y_celda = y_inicio

                pdf.aplicar_formato_condicional_pdf(
                    df_original, idx, col_name, col_widths[col_idx_visible], altura_fila, x_celda, y_celda
                )

                pdf.set_xy(x_celda, y_celda)
                pdf.multi_cell(col_widths[col_idx_visible], ALTURA_LINEA, texto, 0, 'L')
                
                col_idx_visible += 1

            pdf.set_xy(pdf.l_margin, y_inicio + altura_fila)

        # --- Exportar a bytes (compatible con todas las versiones de fpdf2) ---
        pdf_output = pdf.output(dest='S')

        # Normalizar salida: puede ser str, bytes o bytearray según versión de fpdf2
        if isinstance(pdf_output, str):
            pdf_bytes = pdf_output.encode('latin1')
        elif isinstance(pdf_output, (bytes, bytearray)):
            pdf_bytes = bytes(pdf_output)
        else:
            raise TypeError(f"Tipo inesperado devuelto por fpdf.output(): {type(pdf_output)}")

        return io.BytesIO(pdf_bytes).getvalue()

    except Exception as e:
        st.error(f"Error generando PDF: {e}")
        return None

# NUEVA FUNCIÓN: Generar PDF por EQUIPO solo con expedientes prioritarios
def generar_pdf_equipo_prioritarios(equipo, df_pendientes, num_semana, fecha_max_str):
    """Genera el PDF para un equipo específico solo con expedientes prioritarios"""
    # Filtrar por equipo
    df_equipo = df_pendientes[df_pendientes["EQUIPO"] == equipo].copy()
    
    if df_equipo.empty:
        return None
    
    # Filtrar solo expedientes prioritarios (RUE amarillos)
    df_prioritarios = identificar_filas_prioritarias(df_equipo)
    df_prioritarios = df_prioritarios[df_prioritarios['_prioridad'] == 1].copy()
    
    if df_prioritarios.empty:
        return None
    
    # Eliminar columna temporal de prioridad
    df_prioritarios = df_prioritarios.drop('_prioridad', axis=1)
    
    # 🔥 CORRECCIÓN: ORDENAR POR ANTIGÜEDAD DESCENDENTE
    # Buscar columna de antigüedad
    columnas_antiguedad = [col for col in df_prioritarios.columns if 'ANTIGÜEDAD' in col.upper() or 'DÍAS' in col.upper()]
    if columnas_antiguedad:
        columna_antiguedad = columnas_antiguedad[0]
        df_prioritarios[columna_antiguedad] = pd.to_numeric(df_prioritarios[columna_antiguedad], errors='coerce').fillna(0).astype(int)
        df_prioritarios = df_prioritarios.sort_values(columna_antiguedad, ascending=False)
    
    # Procesar datos para PDF
    indices_a_incluir = list(range(df_prioritarios.shape[1]))
    indices_a_excluir = {1, 4, 5, 6, 13}
    
    # EXCLUIR también la columna "FECHA DE ACTUALIZACIÓN DATOS" si existe
    for idx, col_name in enumerate(df_prioritarios.columns):
        if "FECHA DE ACTUALIZACIÓN DATOS" in col_name.upper():
            indices_a_excluir.add(idx)
    
    indices_finales = [i for i in indices_a_incluir if i not in indices_a_excluir]
    NOMBRES_COLUMNAS_PDF = df_prioritarios.columns[indices_finales].tolist()

    # Crear DataFrame para mostrar (con fechas formateadas y "nan" reemplazados)
    df_pdf_mostrar = df_prioritarios[NOMBRES_COLUMNAS_PDF].copy()
    
    # Formatear fechas y reemplazar "nan" - Y ASEGURAR ANTIGÜEDAD COMO ENTERO
    for col in df_pdf_mostrar.columns:
        if df_pdf_mostrar[col].dtype == 'object':
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: "" if pd.isna(x) or str(x).lower() == "nan" else x
            )
        elif 'fecha' in col.lower():
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""
            )
        # 🔥 CORRECCIÓN: FORZAR COLUMNAS NUMÉRICAS A ENTEROS SIN DECIMALES
        elif df_pdf_mostrar[col].dtype in ['float64', 'float32', 'int64', 'int32']:
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: f"{int(x):,}".replace(",", ".") if pd.notna(x) else "0"
            )

    num_expedientes = len(df_pdf_mostrar)
    
    # Nombre único con timestamp
    timestamp = datetime.now().strftime("%H%M%S")
    titulo_pdf = f"{equipo} - Semana {num_semana} a {fecha_max_str} - Expedientes Prioritarios ({num_expedientes})"
    
    return dataframe_to_pdf_bytes(df_pdf_mostrar, titulo_pdf, df_original=df_prioritarios)

def generar_pdf_resumen_kpi(df_kpis_semanales, num_semana, fecha_max_str, df_combinado, semanas_disponibles, FECHA_REFERENCIA, fecha_max):
    """Genera un PDF con el resumen de KPIs principales que incluye gráficos"""
    try:
        # Filtrar datos de la semana actual
        kpis_semana = df_kpis_semanales[df_kpis_semanales['semana_numero'] == num_semana].iloc[0]
        
        pdf = PDFResumenKPI()
        
        # Título principal
        pdf.set_font('Arial', 'B', 14)
        pdf.cell(0, 10, f'RESUMEN DE KPIs - SEMANA {num_semana}', 0, 1, 'C')
        pdf.cell(0, 5, f'Periodo: {fecha_max_str}', 0, 1, 'C')
        pdf.ln(10)
        
        # SECCIÓN 1: KPIs PRINCIPALES
        pdf.add_section_title("KPIs PRINCIPALES")
        
        # Semanales
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, "Semanales:", 0, 1)
        pdf.add_metric("Nuevos Expedientes", f"{int(kpis_semana['nuevos_expedientes']):,}".replace(",", "."))
        pdf.add_metric("Expedientes Despachados", f"{int(kpis_semana['despachados_semana']):,}".replace(",", "."))
        pdf.add_metric("Coef. Absorcion (Desp/Nuevos)", f"{kpis_semana['c_abs_despachados_sem']:.2f}%".replace(".", ","))
        pdf.add_metric("Expedientes Cerrados", f"{int(kpis_semana['expedientes_cerrados']):,}".replace(",", "."))
        pdf.add_metric("Coef. Absorcion (Cer/Asig)", f"{kpis_semana['c_abs_cerrados_sem']:.2f}%".replace(".", ","))
        pdf.add_metric("Expedientes Abiertos", f"{int(kpis_semana['total_abiertos']):,}".replace(",", "."))
        
        pdf.ln(5)
        
        # Totales
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, "Totales (desde 01/11/2022):", 0, 1)
        pdf.add_metric("Nuevos Expedientes", f"{int(kpis_semana['nuevos_expedientes_totales']):,}".replace(",", "."))
        pdf.add_metric("Expedientes Despachados", f"{int(kpis_semana['despachados_totales']):,}".replace(",", "."))
        pdf.add_metric("Coef. Absorcion (Desp/Nuevos)", f"{kpis_semana['c_abs_despachados_tot']:.2f}%".replace(".", ","))
        pdf.add_metric("Expedientes Cerrados", f"{int(kpis_semana['expedientes_cerrados_totales']):,}".replace(",", "."))
        pdf.add_metric("Coef. Absorcion (Cer/Asig)", f"{kpis_semana['c_abs_cerrados_tot']:.2f}%".replace(".", ","))
        
        pdf.ln(10)
        
        # SECCIÓN 2: EXPEDIENTES ESPECIALES
        pdf.add_section_title("EXPEDIENTES CON 029, 033, PRE, RSL, PENDIENTE DE FIRMA, DECISION O COMPLETAR TRAMITE")
        pdf.add_metric("Expedientes Especiales", f"{int(kpis_semana['expedientes_especiales']):,}".replace(",", "."))
        pdf.add_metric("Porcentaje sobre Abiertos", f"{kpis_semana['porcentaje_especiales']:.2f}%".replace(".", ","))
        
        pdf.ln(10)
        
        # SECCIÓN 3: TIEMPOS DE TRAMITACION
        pdf.add_section_title("TIEMPOS DE TRAMITACION (en dias)")
        
        # Expedientes Despachados
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 7, "Expedientes Despachados:", 0, 1)
        pdf.add_metric("Tiempo Medio", f"{kpis_semana['tiempo_medio_despachados']:.0f} dias")
        pdf.add_metric("Percentil 90", f"{kpis_semana['percentil_90_despachados']:.0f} dias")
        pdf.add_metric("<=180 dias", f"{kpis_semana['percentil_180_despachados']:.2f}%".replace(".", ","))
        pdf.add_metric("<=120 dias", f"{kpis_semana['percentil_120_despachados']:.2f}%".replace(".", ","))
        
        pdf.ln(3)
        
        # Expedientes Cerrados
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 7, "Expedientes Cerrados:", 0, 1)
        pdf.add_metric("Tiempo Medio", f"{kpis_semana['tiempo_medio_cerrados']:.0f} dias")
        pdf.add_metric("Percentil 90", f"{kpis_semana['percentil_90_cerrados']:.0f} dias")
        pdf.add_metric("<=180 dias", f"{kpis_semana['percentil_180_cerrados']:.2f}%".replace(".", ","))
        pdf.add_metric("<=120 dias", f"{kpis_semana['percentil_120_cerrados']:.2f}%".replace(".", ","))
        
        pdf.ln(3)
        
        # Expedientes Abiertos
        pdf.set_font('Arial', 'B', 9)
        pdf.cell(0, 7, "Expedientes Abiertos:", 0, 1)
        pdf.add_metric("Percentil 90", f"{kpis_semana['percentil_90_abiertos']:.0f} dias")
        pdf.add_metric("<=180 dias", f"{kpis_semana['percentil_180_abiertos']:.2f}%".replace(".", ","))
        pdf.add_metric("<=120 dias", f"{kpis_semana['percentil_120_abiertos']:.2f}%".replace(".", ","))
        
        # Información del período
        pdf.ln(10)
        pdf.set_font('Arial', 'I', 8)
        if kpis_semana['es_semana_actual']:
            periodo_texto = f"Periodo de la semana (ACTUAL): {kpis_semana['inicio_semana'].strftime('%d/%m/%Y')} a {kpis_semana['fin_semana'].strftime('%d/%m/%Y')} - {kpis_semana['dias_semana']} dias"
        else:
            periodo_texto = f"Periodo de la semana: {kpis_semana['inicio_semana'].strftime('%d/%m/%Y')} a {kpis_semana['fin_semana'].strftime('%d/%m/%Y')} - {kpis_semana['dias_semana']} dias"
        
        pdf.cell(0, 5, periodo_texto, 0, 1)
        
        # ===== NUEVA SECCIÓN: GRÁFICOS =====
        pdf.add_page()
        pdf.add_section_title("GRAFICOS DE EVOLUCION - SEMANA " + str(num_semana))

        try:
            # Verificar si kaleido está disponible
            import plotly.io as pio
            if not hasattr(pio, 'kaleido') or pio.kaleido.scope is None:
                raise ImportError("Kaleido no disponible")
            
            # Generar gráficos temporales
            datos_grafico = df_kpis_semanales.copy()
            
            # GRÁFICO 1: Evolución de expedientes principales
            fig1 = px.line(
                datos_grafico,
                x='semana_numero',
                y=['nuevos_expedientes', 'despachados_semana', 'expedientes_cerrados'],
                title=f'Evolucion de Expedientes - Semana {num_semana}',
                labels={'semana_numero': 'Semana', 'value': 'Cantidad', 'variable': 'KPI'},
                color_discrete_map={
                    'nuevos_expedientes': '#1f77b4',
                    'despachados_semana': '#ff7f0e',
                    'expedientes_cerrados': '#2ca02c'
                }
            )
            fig1.update_layout(height=400, showlegend=True)
            fig1.add_vline(x=num_semana, line_dash="dash", line_color="red")
            
            # Guardar gráfico temporalmente
            temp_chart1 = user_env.get_temp_path(f"chart1_{num_semana}.png")
            fig1.write_image(temp_chart1)
            
            # Insertar gráfico en PDF
            pdf.ln(5)
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Evolucion de Expedientes (Nuevos, Despachados, Cerrados)", 0, 1)
            pdf.image(temp_chart1, x=10, w=190)
            
            pdf.ln(5)
            
            # GRÁFICO 2: Expedientes Abiertos
            fig2 = px.line(
                datos_grafico,
                x='semana_numero',
                y=['total_abiertos'],
                title=f'Expedientes Abiertos - Semana {num_semana}',
                labels={'semana_numero': 'Semana', 'value': 'Cantidad'},
                color_discrete_sequence=['#d62728']
            )
            fig2.update_layout(height=400, showlegend=False)
            fig2.add_vline(x=num_semana, line_dash="dash", line_color="red")
            
            temp_chart2 = user_env.get_temp_path(f"chart2_{num_semana}.png")
            fig2.write_image(temp_chart2)
            
            pdf.add_page()
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Evolucion de Expedientes Abiertos", 0, 1)
            pdf.image(temp_chart2, x=10, w=190)
            
            pdf.ln(5)
            
            # GRÁFICO 3: Coeficientes de absorción
            fig3 = px.line(
                datos_grafico,
                x='semana_numero',
                y=['c_abs_despachados_sem', 'c_abs_cerrados_sem'],
                title=f'Coeficientes de Absorcion - Semana {num_semana}',
                labels={'semana_numero': 'Semana', 'value': 'Porcentaje (%)', 'variable': 'Indicador'},
                color_discrete_map={
                    'c_abs_despachados_sem': '#9467bd',
                    'c_abs_cerrados_sem': '#8c564b'
                }
            )
            fig3.update_layout(height=400, showlegend=True)
            fig3.add_vline(x=num_semana, line_dash="dash", line_color="red")
            
            temp_chart3 = user_env.get_temp_path(f"chart3_{num_semana}.png")
            fig3.write_image(temp_chart3)
            
            pdf.add_page()
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Coeficientes de Absorcion Semanales (%)", 0, 1)
            pdf.image(temp_chart3, x=10, w=190)
            
            pdf.ln(5)
            
            # GRÁFICO 4: Tiempos de tramitación
            fig4 = px.line(
                datos_grafico,
                x='semana_numero',
                y=['tiempo_medio_despachados', 'tiempo_medio_cerrados', 'percentil_90_despachados', 'percentil_90_cerrados'],
                title=f'Tiempos de Tramitacion - Semana {num_semana}',
                labels={'semana_numero': 'Semana', 'value': 'Dias', 'variable': 'Indicador'},
                color_discrete_map={
                    'tiempo_medio_despachados': '#ff7f0e',
                    'tiempo_medio_cerrados': '#2ca02c',
                    'percentil_90_despachados': '#ffbb78',
                    'percentil_90_cerrados': '#98df8a'
                }
            )
            fig4.update_layout(height=400, showlegend=True)
            fig4.add_vline(x=num_semana, line_dash="dash", line_color="red")
            
            temp_chart4 = user_env.get_temp_path(f"chart4_{num_semana}.png")
            fig4.write_image(temp_chart4)
            
            pdf.add_page()
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Tiempos de Tramitacion (Medios y Percentil 90)", 0, 1)
            pdf.image(temp_chart4, x=10, w=190)
            
            pdf.ln(5)
            
            # GRÁFICO 5: Porcentajes 120/180 días
            fig5 = px.line(
                datos_grafico,
                x='semana_numero',
                y=['percentil_180_despachados', 'percentil_120_despachados', 'percentil_180_cerrados', 'percentil_120_cerrados'],
                title=f'Expedientes dentro de Plazos - Semana {num_semana}',
                labels={'semana_numero': 'Semana', 'value': 'Porcentaje (%)', 'variable': 'Indicador'},
                color_discrete_map={
                    'percentil_180_despachados': '#ff7f0e',
                    'percentil_120_despachados': '#ffddaa',
                    'percentil_180_cerrados': '#2ca02c',
                    'percentil_120_cerrados': '#98df8a'
                }
            )
            fig5.update_layout(height=400, showlegend=True)
            fig5.add_vline(x=num_semana, line_dash="dash", line_color="red")
            
            temp_chart5 = user_env.get_temp_path(f"chart5_{num_semana}.png")
            fig5.write_image(temp_chart5)
            
            pdf.add_page()
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Porcentaje de Expedientes dentro de Plazos (120/180 dias)", 0, 1)
            pdf.image(temp_chart5, x=10, w=190)
            
            # Limpiar archivos temporales
            try:
                os.remove(temp_chart1)
                os.remove(temp_chart2)
                os.remove(temp_chart3)
                os.remove(temp_chart4)
                os.remove(temp_chart5)
            except:
                pass
                
        except Exception as chart_error:
            pdf.ln(5)
            pdf.set_font('Arial', 'I', 8)
            pdf.cell(0, 5, f"Nota: No se pudieron incluir los graficos. Error: {str(chart_error)}", 0, 1)
            pdf.cell(0, 5, "Instale Kaleido: pip install kaleido", 0, 1)
            
            # Agregar tabla de datos como alternativa
            pdf.ln(3)
            pdf.set_font('Arial', 'B', 9)
            pdf.cell(0, 6, "Datos de evolucion (ultimas 8 semanas):", 0, 1)
            
            # Mostrar tabla con datos numéricos
            datos_tabla = datos_grafico.tail(8)[[
                'semana_numero', 'nuevos_expedientes', 'despachados_semana', 
                'expedientes_cerrados', 'total_abiertos', 'c_abs_despachados_sem'
            ]]
            pdf.set_font('Arial', '', 6)
            
            # Encabezados de tabla
            headers = ["Sem", "Nuevos", "Despach", "Cerrad", "Abiert", "CoefAbs%"]
            widths = [12, 15, 15, 15, 15, 15]
            
            for i, header in enumerate(headers):
                pdf.cell(widths[i], 5, header, 1)
            pdf.ln()
            
            # Datos
            for _, row in datos_tabla.iterrows():
                pdf.cell(widths[0], 5, str(int(row['semana_numero'])), 1)
                pdf.cell(widths[1], 5, str(int(row['nuevos_expedientes'])), 1)
                pdf.cell(widths[2], 5, str(int(row['despachados_semana'])), 1)
                pdf.cell(widths[3], 5, str(int(row['expedientes_cerrados'])), 1)
                pdf.cell(widths[4], 5, str(int(row['total_abiertos'])), 1)
                pdf.cell(widths[5], 5, f"{row['c_abs_despachados_sem']:.1f}%", 1)
                pdf.ln()
        
        # Exportar a bytes
        pdf_output = pdf.output(dest='S')
        
        if isinstance(pdf_output, str):
            pdf_bytes = pdf_output.encode('latin1')
        elif isinstance(pdf_output, (bytes, bytearray)):
            pdf_bytes = bytes(pdf_output)
        else:
            raise TypeError(f"Tipo inesperado devuelto por fpdf.output(): {type(pdf_output)}")

        return io.BytesIO(pdf_bytes).getvalue()

    except Exception as e:
        st.error(f"Error generando PDF de resumen KPI: {e}")
        import traceback
        st.error(f"Detalle del error: {traceback.format_exc()}")
        return None

# NUEVA FUNCIÓN: Calcular KPIs para una semana específica
def calcular_kpis_para_semana(_df, semana_fin, es_semana_actual=False):
    """Calcula los KPIs para una semana específica"""
    # Si es la semana actual (última), incluir el día completo de fecha_max (viernes a viernes - 8 días)
    if es_semana_actual:
        inicio_semana = semana_fin - timedelta(days=7)  # Viernes anterior
        fin_semana = semana_fin  # Viernes actual (fecha_max)
        dias_semana = 8
    else:
        # Para semanas históricas: viernes a jueves (7 días)
        inicio_semana = semana_fin - timedelta(days=7)  # Viernes anterior
        fin_semana = semana_fin - timedelta(days=1)     # Jueves
        dias_semana = 7
    
    # Fecha de inicio para totales (01/11/2022)
    fecha_inicio_totales = pd.to_datetime("2022-11-01")
    
    # ===== NUEVOS EXPEDIENTES =====
    if 'FECHA APERTURA' in _df.columns:
        nuevos_expedientes = _df[
            (_df['FECHA APERTURA'] >= inicio_semana) & 
            (_df['FECHA APERTURA'] <= fin_semana)
        ].shape[0]
        
        nuevos_expedientes_totales = _df[
            (_df['FECHA APERTURA'] >= fecha_inicio_totales) & 
            (_df['FECHA APERTURA'] <= fin_semana)
        ].shape[0]
    else:
        nuevos_expedientes = 0
        nuevos_expedientes_totales = 0

    # ===== EXPEDIENTES DESPACHADOS =====
    # FECHA RESOLUCIÓN distinta de 09/09/9999 más CERRADOS con FECHA RESOLUCIÓN 09/09/9999
    if all(col in _df.columns for col in ['FECHA RESOLUCIÓN', 'ESTADO', 'FECHA CIERRE']):
        # Convertir columnas de fecha a datetime
        _df['FECHA RESOLUCIÓN'] = pd.to_datetime(_df['FECHA RESOLUCIÓN'], errors='coerce')
        _df['FECHA CIERRE'] = pd.to_datetime(_df['FECHA CIERRE'], errors='coerce')

        fecha_9999 = pd.to_datetime('9999-09-09', errors='coerce')

        # 1️⃣ Expedientes con FECHA RESOLUCIÓN real (distinta de 9999 y no nula) dentro del rango semanal
        despachados_cond1 = _df[
            (_df['FECHA RESOLUCIÓN'].notna()) &
            (_df['FECHA RESOLUCIÓN'] != fecha_9999) &
            (_df['FECHA RESOLUCIÓN'] >= inicio_semana) &
            (_df['FECHA RESOLUCIÓN'] <= fin_semana)
        ]

        # 2️⃣ Expedientes CERRADOS con FECHA RESOLUCIÓN = 9999-09-09 o vacía
        #     y cuya FECHA CIERRE esté dentro del rango semanal
        despachados_cond2 = _df[
            (_df['ESTADO'] == 'Cerrado') &
            (
                (_df['FECHA RESOLUCIÓN'].isna()) |
                (_df['FECHA RESOLUCIÓN'] == fecha_9999)
            ) &
            (_df['FECHA CIERRE'].notna()) &
            (_df['FECHA CIERRE'] >= inicio_semana) &
            (_df['FECHA CIERRE'] <= fin_semana)
        ]

        # 🔹 Despachados semana = reales + cerrados con 9999/vacía pero cierre en rango
        despachados_semana = pd.concat([despachados_cond1, despachados_cond2]).drop_duplicates().shape[0]

        # 3️⃣ Totales: igual pero usando fecha_inicio_totales
        despachados_cond1_totales = _df[
            (_df['FECHA RESOLUCIÓN'].notna()) &
            (_df['FECHA RESOLUCIÓN'] != fecha_9999) &
            (_df['FECHA RESOLUCIÓN'] >= fecha_inicio_totales) &
            (_df['FECHA RESOLUCIÓN'] <= fin_semana)
        ]

        despachados_cond2_totales = _df[
            (_df['ESTADO'] == 'Cerrado') &
            (
                (_df['FECHA RESOLUCIÓN'].isna()) |
                (_df['FECHA RESOLUCIÓN'] == fecha_9999)
            ) &
            (_df['FECHA CIERRE'].notna()) &
            (_df['FECHA CIERRE'] >= fecha_inicio_totales) &
            (_df['FECHA CIERRE'] <= fin_semana)
        ]

        # 🔹 Despachados totales = reales + cerrados especiales (con cierre válido)
        despachados_totales = pd.concat([despachados_cond1_totales, despachados_cond2_totales]).drop_duplicates().shape[0]

    else:
        despachados_semana = 0
        despachados_totales = 0

    # ===== COEFICIENTE DE ABSORCIÓN 1 (Despachados/Nuevos) =====
    c_abs_despachados_sem = (despachados_semana / nuevos_expedientes * 100) if nuevos_expedientes > 0 else 0
    c_abs_despachados_tot = (despachados_totales / nuevos_expedientes_totales * 100) if nuevos_expedientes_totales > 0 else 0

    # ===== EXPEDIENTES CERRADOS =====
    if 'FECHA CIERRE' in _df.columns:
        expedientes_cerrados = _df[
            (_df['FECHA CIERRE'] >= inicio_semana) & 
            (_df['FECHA CIERRE'] <= fin_semana)
        ].shape[0]
        
        expedientes_cerrados_totales = _df[
            (_df['FECHA CIERRE'] >= fecha_inicio_totales) & 
            (_df['FECHA CIERRE'] <= fin_semana)
        ].shape[0]
    else:
        expedientes_cerrados = 0
        expedientes_cerrados_totales = 0

    # ===== EXPEDIENTES ABIERTOS =====
    if 'FECHA CIERRE' in _df.columns and 'FECHA APERTURA' in _df.columns:
        total_abiertos = _df[
            (_df['FECHA APERTURA'] <= fin_semana) & 
            ((_df['FECHA CIERRE'] > fin_semana) | (_df['FECHA CIERRE'].isna()))
        ].shape[0]
    else:
        total_abiertos = 0

    # ===== COEFICIENTE DE ABSORCIÓN 2 (Cerrados/Asignados) =====
    # Asumimos que "Asignados" son los Nuevos Expedientes
    c_abs_cerrados_sem = (expedientes_cerrados / nuevos_expedientes * 100) if nuevos_expedientes > 0 else 0
    c_abs_cerrados_tot = (expedientes_cerrados_totales / nuevos_expedientes_totales * 100) if nuevos_expedientes_totales > 0 else 0

    # ===== EXPEDIENTES CON 029, 033, PRE o RSL =====
    if 'ETIQ. PENÚLTIMO TRAM.' in _df.columns:
        expedientes_especiales = _df[
            (_df['ETIQ. PENÚLTIMO TRAM.'].notna()) & 
            # INDICAMOS LOS QUE NO TIENEN EL PENÚLTIMO TRÁMITE 1 APERTURA Y 10 DATEXPTE, QUE NO SON COMPETENCIA DE LOS EQUIPOS RECTAUTO
            (~_df['ETIQ. PENÚLTIMO TRAM.'].isin(['1 APERTURA', '10 DATEXPTE'])) &
            (_df['FECHA APERTURA'] <= fin_semana) & 
            ((_df['FECHA CIERRE'] > fin_semana) | (_df['FECHA CIERRE'].isna()))
        ].shape[0]
        
        porcentaje_especiales = (expedientes_especiales / total_abiertos * 100) if total_abiertos > 0 else 0
    else:
        expedientes_especiales = 0
        porcentaje_especiales = 0

    # ===== CÁLCULOS DE TIEMPOS =====
    # Tiempos para expedientes Despachados
    if all(col in _df.columns for col in ['FECHA RESOLUCIÓN', 'FECHA INICIO TRAMITACIÓN', 'ESTADO', 'FECHA CIERRE']):
        # Convertir fechas
        _df['FECHA RESOLUCIÓN'] = pd.to_datetime(_df['FECHA RESOLUCIÓN'], errors='coerce')
        _df['FECHA INICIO TRAMITACIÓN'] = pd.to_datetime(_df['FECHA INICIO TRAMITACIÓN'], errors='coerce')
        _df['FECHA CIERRE'] = pd.to_datetime(_df['FECHA CIERRE'], errors='coerce')

        fecha_9999 = pd.to_datetime('9999-09-09', errors='coerce')

        # 1️⃣ Expedientes con resolución real (fecha válida)
        despachados_real = _df[
            (_df['FECHA RESOLUCIÓN'].notna()) &
            (_df['FECHA RESOLUCIÓN'] != fecha_9999) &
            (_df['FECHA RESOLUCIÓN'] >= fecha_inicio_totales) &
            (_df['FECHA RESOLUCIÓN'] <= fin_semana)
        ].copy()

        # 2️⃣ Expedientes cerrados con resolución vacía o 9999, pero con fecha de cierre válida
        despachados_cerrados = _df[
            (_df['ESTADO'] == 'Cerrado') &
            (
                (_df['FECHA RESOLUCIÓN'].isna()) |
                (_df['FECHA RESOLUCIÓN'] == fecha_9999)
            ) &
            (_df['FECHA CIERRE'].notna()) &
            (_df['FECHA CIERRE'] >= fecha_inicio_totales) &
            (_df['FECHA CIERRE'] <= fin_semana)
        ].copy()

        # Unificar ambos conjuntos
        despachados_tiempo = pd.concat([despachados_real, despachados_cerrados]).drop_duplicates().copy()

        if not despachados_tiempo.empty:
            # Crear columna de referencia de fecha final (resolución o cierre)
            despachados_tiempo['FECHA_FINAL'] = despachados_tiempo.apply(
                lambda r: r['FECHA RESOLUCIÓN'] if pd.notna(r['FECHA RESOLUCIÓN']) and r['FECHA RESOLUCIÓN'] != fecha_9999 else r['FECHA CIERRE'],
                axis=1
            )

            # Calcular días de tramitación
            despachados_tiempo['dias_tramitacion'] = (
                despachados_tiempo['FECHA_FINAL'] - despachados_tiempo['FECHA INICIO TRAMITACIÓN']
            ).dt.days

            # KPIs
            tiempo_medio_despachados = despachados_tiempo['dias_tramitacion'].mean()
            percentil_90_despachados = despachados_tiempo['dias_tramitacion'].quantile(0.9)
            percentil_180_despachados = (despachados_tiempo['dias_tramitacion'] <= 180).mean() * 100
            percentil_120_despachados = (despachados_tiempo['dias_tramitacion'] <= 120).mean() * 100
        else:
            tiempo_medio_despachados = 0
            percentil_90_despachados = 0
            percentil_180_despachados = 0
            percentil_120_despachados = 0

    else:
        tiempo_medio_despachados = 0
        percentil_90_despachados = 0
        percentil_180_despachados = 0
        percentil_120_despachados = 0

    # Tiempos para expedientes Cerrados
    if 'FECHA CIERRE' in _df.columns and 'FECHA INICIO TRAMITACIÓN' in _df.columns:
        cerrados_tiempo = _df[
            (_df['FECHA CIERRE'].notna()) &
            (_df['FECHA CIERRE'] >= fecha_inicio_totales) & 
            (_df['FECHA CIERRE'] <= fin_semana)
        ].copy()
        
        if not cerrados_tiempo.empty:
            cerrados_tiempo['dias_tramitacion'] = (cerrados_tiempo['FECHA CIERRE'] - cerrados_tiempo['FECHA INICIO TRAMITACIÓN']).dt.days
            tiempo_medio_cerrados = cerrados_tiempo['dias_tramitacion'].mean()
            percentil_90_cerrados = cerrados_tiempo['dias_tramitacion'].quantile(0.9)
            
            # Percentiles para 180 y 120 días
            percentil_180_cerrados = (cerrados_tiempo['dias_tramitacion'] <= 180).mean() * 100
            percentil_120_cerrados = (cerrados_tiempo['dias_tramitacion'] <= 120).mean() * 100
        else:
            tiempo_medio_cerrados = 0
            percentil_90_cerrados = 0
            percentil_180_cerrados = 0
            percentil_120_cerrados = 0
    else:
        tiempo_medio_cerrados = 0
        percentil_90_cerrados = 0
        percentil_180_cerrados = 0
        percentil_120_cerrados = 0

    # Tiempos para expedientes Abiertos
    if 'FECHA INICIO TRAMITACIÓN' in _df.columns:
        abiertos_tiempo = _df[
            (_df['FECHA APERTURA'] <= fin_semana) & 
            ((_df['FECHA CIERRE'] > fin_semana) | (_df['FECHA CIERRE'].isna()))
        ].copy()
        
        if not abiertos_tiempo.empty:
            abiertos_tiempo['dias_tramitacion'] = (fin_semana - abiertos_tiempo['FECHA INICIO TRAMITACIÓN']).dt.days
            percentil_90_abiertos = abiertos_tiempo['dias_tramitacion'].quantile(0.9)
            
            # Percentiles para 180 y 120 días
            percentil_180_abiertos = (abiertos_tiempo['dias_tramitacion'] <= 180).mean() * 100
            percentil_120_abiertos = (abiertos_tiempo['dias_tramitacion'] <= 120).mean() * 100
        else:
            percentil_90_abiertos = 0
            percentil_180_abiertos = 0
            percentil_120_abiertos = 0
    else:
        percentil_90_abiertos = 0
        percentil_180_abiertos = 0
        percentil_120_abiertos = 0
    
    return {
        'nuevos_expedientes': nuevos_expedientes,
        'nuevos_expedientes_totales': nuevos_expedientes_totales,
        'despachados_semana': despachados_semana,
        'despachados_totales': despachados_totales,
        'c_abs_despachados_sem': c_abs_despachados_sem,
        'c_abs_despachados_tot': c_abs_despachados_tot,
        'expedientes_cerrados': expedientes_cerrados,
        'expedientes_cerrados_totales': expedientes_cerrados_totales,
        'total_abiertos': total_abiertos,
        'c_abs_cerrados_sem': c_abs_cerrados_sem,
        'c_abs_cerrados_tot': c_abs_cerrados_tot,
        'expedientes_especiales': expedientes_especiales,
        'porcentaje_especiales': porcentaje_especiales,
        'tiempo_medio_despachados': tiempo_medio_despachados,
        'percentil_90_despachados': percentil_90_despachados,
        'percentil_180_despachados': percentil_180_despachados,
        'percentil_120_despachados': percentil_120_despachados,
        'tiempo_medio_cerrados': tiempo_medio_cerrados,
        'percentil_90_cerrados': percentil_90_cerrados,
        'percentil_180_cerrados': percentil_180_cerrados,
        'percentil_120_cerrados': percentil_120_cerrados,
        'percentil_90_abiertos': percentil_90_abiertos,
        'percentil_180_abiertos': percentil_180_abiertos,
        'percentil_120_abiertos': percentil_120_abiertos,
        'inicio_semana': inicio_semana,
        'fin_semana': fin_semana,
        'dias_semana': dias_semana,
        'es_semana_actual': es_semana_actual
    }

# CALCULAR KPIs PARA TODAS LAS SEMANAS con cache de 2 horas - ACTUALIZADA
@st.cache_data(ttl=CACHE_TTL, show_spinner="📊 Calculando KPIs históricos...")
def calcular_kpis_todas_semanas_optimizado(_df, _semanas, _fecha_referencia, _fecha_max, _user_key=user_env.session_id):
    datos_semanales = []
    
    for i, semana in enumerate(_semanas):
        # Determinar si es la semana actual (la última)
        es_semana_actual = (i == len(_semanas) - 1)  # Última semana en la lista
        
        kpis = calcular_kpis_para_semana(_df, semana, es_semana_actual)
        num_semana = ((semana - _fecha_referencia).days) // 7 + 1
        
        datos_semanales.append({
            'semana_numero': num_semana,
            'semana_fin': semana,
            'semana_str': semana.strftime('%d/%m/%Y'),
            'nuevos_expedientes': kpis['nuevos_expedientes'],
            'nuevos_expedientes_totales': kpis['nuevos_expedientes_totales'],
            'despachados_semana': kpis['despachados_semana'],
            'despachados_totales': kpis['despachados_totales'],
            'c_abs_despachados_sem': kpis['c_abs_despachados_sem'],
            'c_abs_despachados_tot': kpis['c_abs_despachados_tot'],
            'expedientes_cerrados': kpis['expedientes_cerrados'],
            'expedientes_cerrados_totales': kpis['expedientes_cerrados_totales'],
            'total_abiertos': kpis['total_abiertos'],
            'c_abs_cerrados_sem': kpis['c_abs_cerrados_sem'],
            'c_abs_cerrados_tot': kpis['c_abs_cerrados_tot'],
            'expedientes_especiales': kpis['expedientes_especiales'],
            'porcentaje_especiales': kpis['porcentaje_especiales'],
            'tiempo_medio_despachados': kpis['tiempo_medio_despachados'],
            'percentil_90_despachados': kpis['percentil_90_despachados'],
            'percentil_180_despachados': kpis['percentil_180_despachados'],
            'percentil_120_despachados': kpis['percentil_120_despachados'],
            'tiempo_medio_cerrados': kpis['tiempo_medio_cerrados'],
            'percentil_90_cerrados': kpis['percentil_90_cerrados'],
            'percentil_180_cerrados': kpis['percentil_180_cerrados'],
            'percentil_120_cerrados': kpis['percentil_120_cerrados'],
            'percentil_90_abiertos': kpis['percentil_90_abiertos'],
            'percentil_180_abiertos': kpis['percentil_180_abiertos'],
            'percentil_120_abiertos': kpis['percentil_120_abiertos'],
            'inicio_semana': kpis['inicio_semana'],
            'fin_semana': kpis['fin_semana'],
            'dias_semana': kpis['dias_semana'],
            'es_semana_actual': kpis['es_semana_actual']
        })
    
    return pd.DataFrame(datos_semanales)

def obtener_hash_archivo(archivo):
    """Genera un hash único del archivo para detectar cambios"""
    if archivo is None:
        return None
    archivo.seek(0)
    file_hash = hashlib.md5(archivo.read()).hexdigest()
    archivo.seek(0)
    return file_hash

# CSS ACTUALIZADO CON ESTILOS PARA BOTONES Y MENÚS DESPLEGABLES
st.markdown("""
<style>
    [data-testid="stSidebar"] {
        background-color: #007933 !important;
    }
    
    .main .block-container {
        background-color: #C4DDCA !important;
        padding: 2rem;
        border-radius: 10px;
    }
    
    .stApp {
        background-color: #92C88F !important;
    }
    
    [data-testid="stSidebar"] * {
        color: white !important;
    }
    
    .main .stMarkdown, .main h1, .main h2, .main h3 {
        color: #333333 !important;
    }
    
    /* ESTILOS ESPECÍFICOS PARA BOTONES EN LA BARRA LATERAL */
    [data-testid="stSidebar"] button {
        background-color: #005a25 !important;
        color: white !important;
        border: 1px solid white !important;
        border-radius: 5px !important;
        padding: 0.5rem 1rem !important;
        font-weight: bold !important;
    }
    
    [data-testid="stSidebar"] button:hover {
        background-color: #003d1a !important;
        color: white !important;
        border: 1px solid white !important;
    }
    
    [data-testid="stSidebar"] button:focus {
        background-color: #003d1a !important;
        color: white !important;
        border: 2px solid #ffffff !important;
        box-shadow: 0 0 0 0.2rem rgba(255, 255, 255, 0.25) !important;
    }
    
    [data-testid="stSidebar"] button[kind="primary"] {
        background-color: #00802b !important;
        color: white !important;
        border: 2px solid #ffffff !important;
    }
    
    [data-testid="stSidebar"] button[kind="primary"]:hover {
        background-color: #006622 !important;
        color: white !important;
    }
    
    /* Estilos para los botones de navegación de semanas en KPI */
    [data-testid="stSidebar"] .stButton button {
        background-color: #005a25 !important;
        color: white !important;
        border: 1px solid white !important;
        border-radius: 5px !important;
        padding: 0.5rem 1rem !important;
        font-weight: bold !important;
        width: 100% !important;
    }
    
    [data-testid="stSidebar"] .stButton button:hover {
        background-color: #003d1a !important;
        color: white !important;
        border: 1px solid white !important;
    }
    
    /* ESTILOS PARA MENÚS DESPLEGABLES EN BARRA LATERAL */
    [data-testid="stSidebar"] .stSelectbox > div > div {
        background-color: #009945 !important;
        color: white !important;
        border: 1px solid white !important;
        border-radius: 5px !important;
    }
    
    [data-testid="stSidebar"] .stSelectbox > div > div:hover {
        background-color: #007933 !important;
        color: white !important;
        border: 1px solid white !important;
    }
    
    [data-testid="stSidebar"] .stSelectbox input {
        color: white !important;
    }
    
    [data-testid="stSidebar"] .stMultiSelect > div > div {
        background-color: #009945 !important;
        color: white !important;
        border: 1px solid white !important;
        border-radius: 5px !important;
    }
    
    [data-testid="stSidebar"] .stMultiSelect > div > div:hover {
        background-color: #007933 !important;
        color: white !important;
        border: 1px solid white !important;
    }
    
    [data-testid="stSidebar"] .stMultiSelect input {
        color: white !important;
    }
    
    /* Estilos para las opciones del menú desplegable */
    [data-testid="stSidebar"] .stSelectbox [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] .stMultiSelect [data-testid="stMarkdownContainer"] p {
        color: white !important;
    }
    
    /* Estilos para las opciones seleccionadas en multiselect */
    [data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"] {
        background-color: #005a25 !important;
        color: white !important;
        border: 1px solid white !important;
        border-radius: 12px !important;
    }
    
    [data-testid="stSidebar"] .stMultiSelect [data-baseweb="tag"]:hover {
        background-color: #003d1a !important;
        color: white !important;
    }
    
    /* Estilos para el menú desplegable abierto */
    [data-testid="stSidebar"] div[data-baseweb="popover"] {
        background-color: #009945 !important;
        border: 1px solid white !important;
        border-radius: 5px !important;
    }
    
    [data-testid="stSidebar"] div[data-baseweb="menu"] {
        background-color: #009945 !important;
        color: white !important;
    }
    
    [data-testid="stSidebar"] div[data-baseweb="menu"] li {
        background-color: #009945 !important;
        color: white !important;
    }
    
    [data-testid="stSidebar"] div[data-baseweb="menu"] li:hover {
        background-color: #007933 !important;
        color: white !important;
    }
    
    [data-testid="stSidebar"] div[data-baseweb="menu"] li:focus {
        background-color: #005a25 !important;
        color: white !important;
    }
    
    /* Estilos para el texto dentro de los menús desplegables */
    [data-testid="stSidebar"] div[data-baseweb="popover"] * {
        color: white !important;
    }
</style>
""", unsafe_allow_html=True)

# Logo
st.sidebar.image("Logo Atrian.png", width=260)

# Botón para limpiar cache
with st.sidebar:
    st.markdown("---")
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("🔄 Limpiar cache", help="Limpiar toda la cache y recargar"):
            st.cache_data.clear()
            # Mantener solo los datos esenciales
            keys_to_keep = ['df_combinado', 'df_usuarios', 'archivos_hash', 'filtro_estado', 'filtro_equipo', 'filtro_usuario']
            for key in list(st.session_state.keys()):
                if key not in keys_to_keep:
                    del st.session_state[key]
            st.success("Cache limpiada correctamente")
            st.rerun()
    
    with col2:
        if st.button("🧹 Limpiar temp", help="Limpiar archivos temporales"):
            user_env.cleanup()
            st.success("Archivos temporales limpiados")

# NUEVA SECCIÓN: CARGA DE CINCO ARCHIVOS (incluyendo DOCUMENTOS)
st.markdown("---")
st.subheader("📁 Carga de Archivos")

# Crear cinco columnas para los archivos
col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.markdown('<div style="text-align: center; background-color: #1f77b4; padding: 10px; border-radius: 5px; margin-bottom: 10px;">'
                '<h4 style="color: white; margin: 0;">📊 RECTAUTO</h4>'
                '</div>', unsafe_allow_html=True)
    archivo_rectauto = st.file_uploader(
        "Archivo principal de expedientes",
        type=["xlsx", "xls"],
        key="rectauto_upload",
        label_visibility="collapsed",
        help="Sube el archivo principal RECTAUTO"
    )
    if archivo_rectauto:
        st.success(f"✅ {archivo_rectauto.name}")
    else:
        st.info("⏳ Esperando archivo RECTAUTO")

with col2:
    st.markdown('<div style="text-align: center; background-color: #ff7f0e; padding: 10px; border-radius: 5px; margin-bottom: 10px;">'
                '<h4 style="color: white; margin: 0;">📨 NOTIFICA</h4>'
                '</div>', unsafe_allow_html=True)
    archivo_notifica = st.file_uploader(
        "Archivo de notificaciones",
        type=["xlsx", "xls"],
        key="notifica_upload",
        label_visibility="collapsed",
        help="Sube el archivo NOTIFICA"
    )
    if archivo_notifica:
        st.success(f"✅ {archivo_notifica.name}")
    else:
        st.info("⏳ Esperando archivo NOTIFICA")

with col3:
    st.markdown('<div style="text-align: center; background-color: #2ca02c; padding: 10px; border-radius: 5px; margin-bottom: 10px;">'
                '<h4 style="color: white; margin: 0;">⚡ TRIAJE</h4>'
                '</div>', unsafe_allow_html=True)
    archivo_triaje = st.file_uploader(
        "Archivo de triaje",
        type=["xlsx", "xls"],
        key="triaje_upload",
        label_visibility="collapsed",
        help="Sube el archivo TRIAJE"
    )
    if archivo_triaje:
        st.success(f"✅ {archivo_triaje.name}")
    else:
        st.info("⏳ Esperando archivo TRIAJE")

with col4:
    st.markdown('<div style="text-align: center; background-color: #9467bd; padding: 10px; border-radius: 5px; margin-bottom: 10px;">'
                '<h4 style="color: white; margin: 0;">👥 USUARIOS</h4>'
                '</div>', unsafe_allow_html=True)
    archivo_usuarios = st.file_uploader(
        "Archivo de configuración de usuarios",
        type=["xlsx", "xls"],
        key="usuarios_upload",
        label_visibility="collapsed",
        help="Sube el archivo USUARIOS.xlsx"
    )
    if archivo_usuarios:
        st.success(f"✅ {archivo_usuarios.name}")
    else:
        st.info("⏳ Esperando archivo USUARIOS")

with col5:
    st.markdown('<div style="text-align: center; background-color: #d62728; padding: 10px; border-radius: 5px; margin-bottom: 10px;">'
                '<h4 style="color: white; margin: 0;">📄 DOCUMENTOS</h4>'
                '</div>', unsafe_allow_html=True)
    archivo_documentos = st.file_uploader(
        "Archivo de documentación incorporada",
        type=["xlsx"],
        key="documentos_upload",
        label_visibility="collapsed",
        help="Sube el archivo DOCUMENTOS.xlsx"
    )
    if archivo_documentos:
        st.success(f"✅ {archivo_documentos.name}")
    else:
        st.info("⏳ Esperando archivo DOCUMENTOS")

# Estado de carga
st.markdown("---")
st.subheader("📋 Estado de Carga")

# Mostrar estado con métricas (6 columnas ahora)
estado_col1, estado_col2, estado_col3, estado_col4, estado_col5, estado_col6 = st.columns(6)

with estado_col1:
    rectauto_status = "✅ Cargado" if archivo_rectauto else "❌ Pendiente"
    st.metric("RECTAUTO", rectauto_status)

with estado_col2:
    notifica_status = "✅ Cargado" if archivo_notifica else "❌ Pendiente"
    st.metric("NOTIFICA", notifica_status)

with estado_col3:
    triaje_status = "✅ Cargado" if archivo_triaje else "❌ Pendiente"
    st.metric("TRIAJE", triaje_status)

with estado_col4:
    usuarios_status = "✅ Cargado" if archivo_usuarios else "❌ Pendiente"
    st.metric("USUARIOS", usuarios_status)

with estado_col5:
    documentos_status = "✅ Cargado" if archivo_documentos else "❌ Pendiente"
    st.metric("DOCUMENTOS", documentos_status)

with estado_col6:
    archivos_cargados = sum([1 for f in [archivo_rectauto, archivo_notifica, archivo_triaje, archivo_usuarios, archivo_documentos] if f])
    st.metric("Total Cargados", f"{archivos_cargados}/5")

# Procesar archivos cuando estén listos
if archivo_rectauto:
    # Verificar si los archivos han cambiado
    archivos_actuales = {
        'rectauto': obtener_hash_archivo(archivo_rectauto),
        'notifica': obtener_hash_archivo(archivo_notifica) if archivo_notifica else None,
        'triaje': obtener_hash_archivo(archivo_triaje) if archivo_triaje else None,
        'usuarios': obtener_hash_archivo(archivo_usuarios) if archivo_usuarios else None,
        'documentos': obtener_hash_archivo(archivo_documentos) if archivo_documentos else None
    }
    
    archivos_guardados = st.session_state.get("archivos_hash", {})
    
    # Si los archivos son nuevos o cambiaron, procesar
    if (archivos_actuales != archivos_guardados or 
        "df_combinado" not in st.session_state):
        
        with st.spinner("🔄 Combinando archivos por RUE..."):
            try:
                # Cargar RECTAUTO
                df_rectauto = cargar_y_procesar_rectauto(archivo_rectauto)
                
                # Cargar NOTIFICA si está disponible
                df_notifica = None
                if archivo_notifica:
                    df_notifica = cargar_y_procesar_notifica(archivo_notifica)
                
                # Cargar TRIAJE si está disponible
                df_triaje = None
                if archivo_triaje:
                    df_triaje = cargar_y_procesar_triaje(archivo_triaje)
                
                # Cargar USUARIOS si está disponible
                df_usuarios = None
                if archivo_usuarios:
                    df_usuarios = cargar_y_procesar_usuarios(archivo_usuarios)
                
                # Cargar DOCUMENTOS si está disponible
                datos_documentos = None
                if archivo_documentos:
                    datos_documentos = cargar_y_procesar_documentos(archivo_documentos)
                
                # Combinar todos los archivos incluyendo documentación
                df_combinado = combinar_archivos(df_rectauto, df_notifica, df_triaje, df_usuarios, datos_documentos)
                # Convertir columnas de fecha
                df_combinado = convertir_fechas(df_combinado)
                
                # Guardar en session_state
                st.session_state["df_combinado"] = df_combinado
                st.session_state["df_usuarios"] = df_usuarios
                st.session_state["datos_documentos"] = datos_documentos
                st.session_state["archivos_hash"] = archivos_actuales
                
                st.success(f"✅ Archivos combinados correctamente")
                st.info(f"📊 Dataset final: {len(df_combinado)} registros, {len(df_combinado.columns)} columnas")
                if df_usuarios is not None:
                    st.info(f"👥 Usuarios cargados: {len(df_usuarios)} registros")
                if datos_documentos is not None:
                    st.info(f"📄 Documentos cargados: {len(datos_documentos['documentos'])} registros")
                
            except Exception as e:
                st.error(f"❌ Error combinando archivos: {e}")
                # Fallback: usar solo RECTAUTO
                with st.spinner("🔄 Cargando solo RECTAUTO..."):
                    df_rectauto = cargar_y_procesar_rectauto(archivo_rectauto)
                    st.session_state["df_combinado"] = df_rectauto
                    st.session_state["df_usuarios"] = None
                    st.session_state["datos_documentos"] = None
                    st.session_state["archivos_hash"] = archivos_actuales
                    st.warning("⚠️ Usando solo archivo RECTAUTO debido a errores en combinación")
    
    else:
        # Usar datos cacheados
        df_combinado = st.session_state["df_combinado"]
        df_usuarios = st.session_state.get("df_usuarios", None)
        datos_documentos = st.session_state.get("datos_documentos", None)
        st.sidebar.success("✅ Usando datos combinados cacheados")

elif "df_combinado" in st.session_state:
    # Usar datos previamente cargados
    df_combinado = st.session_state["df_combinado"]
    df_usuarios = st.session_state.get("df_usuarios", None)
    datos_documentos = st.session_state.get("datos_documentos", None)
    st.sidebar.info("📊 Datos combinados cargados desde cache")
else:
    st.warning("⚠️ **Carga obligatoria:** Sube al menos el archivo RECTAUTO para continuar")
    st.info("💡 **Archivos opcionales:** NOTIFICA, TRIAJE, USUARIOS y DOCUMENTOS enriquecerán el análisis")
    st.stop()

# Función para identificar filas prioritarias (RUE amarillo)
def identificar_filas_prioritarias(df):
    """Identifica filas que deben aparecer primero (RUE amarillo) con NUEVAS condiciones"""
    try:
        # Crear una columna temporal para ordenar
        df_priorizado = df.copy()
        
        # Identificar filas que cumplen la condición de RUE amarillo
        def es_rue_amarillo(fila):
            try:
                etiq_penultimo = fila.get('ETIQ. PENÚLTIMO TRAM.', '')
                fecha_notif = fila.get('FECHA NOTIFICACIÓN', None)
                docum_incorp = fila.get('DOCUM.INCORP.', '')
                
                # CONDICIÓN 1: "80 PROPRES" con fecha límite superada
                if (str(etiq_penultimo).strip() == "80 PROPRES" and 
                    pd.notna(fecha_notif) and 
                    isinstance(fecha_notif, (pd.Timestamp, datetime))):
                    
                    fecha_limite = fecha_notif + timedelta(days=23)
                    if datetime.now() > fecha_limite:
                        return True
                
                # NUEVA CONDICIÓN 2: "50 REQUERIR" con fecha límite superada
                if (str(etiq_penultimo).strip() == "50 REQUERIR" and 
                    pd.notna(fecha_notif) and 
                    isinstance(fecha_notif, (pd.Timestamp, datetime))):
                    
                    fecha_limite = fecha_notif + timedelta(days=23)
                    if datetime.now() > fecha_limite:
                        return True
                
                # NUEVA CONDICIÓN 3: "70 ALEGACI" o "60 CONTESTA"
                if (str(etiq_penultimo).strip() in ["70 ALEGACI", "60 CONTESTA"]):
                    return True
                
                # NUEVA CONDICIÓN 4: DOCUM.INCORP. no vacío Y distinto de "SOLICITUD" o "REITERA SOLICITUD"
                if (pd.notna(docum_incorp) and 
                    str(docum_incorp).strip() != '' and
                    str(docum_incorp).strip().upper() not in ["SOLICITUD", "REITERA SOLICITUD"]):
                    return True
                    
            except:
                pass
            return False
        
        # Aplicar prioridad
        df_priorizado['_prioridad'] = df_priorizado.apply(es_rue_amarillo, axis=1).astype(int)
        
        return df_priorizado
    
    except Exception as e:
        st.error(f"Error al identificar filas prioritarias: {e}")
        return df

# Función para ordenar DataFrame (RUE amarillos primero y luego por antigüedad)
def ordenar_dataframe_por_prioridad_y_antiguedad(df):
    """Ordena el DataFrame: RUE amarillos primero, luego por antigüedad descendente"""
    try:
        # Primero identificar filas prioritarias
        df_priorizado = identificar_filas_prioritarias(df)
        
        # Buscar la columna de antigüedad por nombre (más robusto)
        columnas_antiguedad = [col for col in df_priorizado.columns if 'ANTIGÜEDAD' in col.upper() or 'DÍAS' in col.upper()]
        
        if columnas_antiguedad:
            columna_antiguedad = columnas_antiguedad[0]
            # 🔥 CORRECCIÓN: FORZAR A ENTEROS ANTES DE ORDENAR
            df_priorizado[columna_antiguedad] = pd.to_numeric(df_priorizado[columna_antiguedad], errors='coerce').fillna(0).astype(int)
        else:
            # Fallback: usar la columna 5 como en el código original
            columna_antiguedad = df_priorizado.columns[5]
            st.warning(f"Usando columna {columna_antiguedad} para antigüedad")
        
        # Ordenar por prioridad (True primero) y luego por antigüedad descendente
        df_ordenado = df_priorizado.sort_values(
            ['_prioridad', columna_antiguedad], 
            ascending=[False, False]  # Prioridad: False=1ros, Antigüedad: False=descendente
        )
        df_ordenado = df_ordenado.drop('_prioridad', axis=1)
        
        return df_ordenado
    
    except Exception as e:
        st.error(f"Error al ordenar DataFrame: {e}")
        return df

# Función para aplicar formato condicional al DataFrame mostrado
def aplicar_formato_condicional_dataframe(df):
    """Aplica formato condicional al DataFrame para Streamlit con NUEVAS condiciones"""
    styles = pd.DataFrame('', index=df.index, columns=df.columns)
    
    try:
        # Condición 1: USUARIO-CSV con fondo rojo cuando USUARIO es distinto de USUARIO-CSV
        if 'USUARIO-CSV' in df.columns and 'USUARIO' in df.columns:
            mask_usuario = df['USUARIO'] != df['USUARIO-CSV']
            styles.loc[mask_usuario, 'USUARIO-CSV'] = 'background-color: rgb(255, 0, 0)'
        
        # Condición 2: RUE con MÚLTIPLES condiciones específicas
        if 'RUE' in df.columns:
            for idx, row in df.iterrows():
                try:
                    etiq_penultimo = row.get('ETIQ. PENÚLTIMO TRAM.', '')
                    fecha_notif = row.get('FECHA NOTIFICACIÓN', None)
                    docum_incorp = row.get('DOCUM.INCORP.', '')
                    
                    es_amarillo = False
                    
                    # CONDICIÓN 2.1: "80 PROPRES" con fecha límite superada
                    if (str(etiq_penultimo) == "80 PROPRES" and 
                        pd.notna(fecha_notif)):
                        if isinstance(fecha_notif, (pd.Timestamp, datetime)):
                            fecha_limite = fecha_notif + timedelta(days=23)
                            if datetime.now() > fecha_limite:
                                es_amarillo = True
                    
                    # NUEVA CONDICIÓN 2.2: "50 REQUERIR" con fecha límite superada
                    elif (str(etiq_penultimo) == "50 REQUERIR" and 
                          pd.notna(fecha_notif)):
                        if isinstance(fecha_notif, (pd.Timestamp, datetime)):
                            fecha_limite = fecha_notif + timedelta(days=23)
                            if datetime.now() > fecha_limite:
                                es_amarillo = True
                    
                    # NUEVA CONDICIÓN 2.3: "70 ALEGACI" o "60 CONTESTA"
                    elif str(etiq_penultimo) in ["70 ALEGACI", "60 CONTESTA"]:
                        es_amarillo = True
                    
                    # NUEVA CONDICIÓN 2.4: DOCUM.INCORP. no vacío Y distinto de "SOLICITUD" o "REITERA SOLICITUD"
                    elif (pd.notna(docum_incorp) and 
                          str(docum_incorp).strip() != '' and
                          str(docum_incorp).strip().upper() not in ["SOLICITUD", "REITERA SOLICITUD"]):
                        es_amarillo = True
                    
                    if es_amarillo:
                        styles.loc[idx, 'RUE'] = 'background-color: rgb(255, 255, 0)'
                        
                except (TypeError, ValueError):
                    continue  # Ignorar errores de fecha
        
        # Condición 3: Resaltar DOCUM.INCORP. cuando tiene valor
        if 'DOCUM.INCORP.' in df.columns:
            mask_docum = df['DOCUM.INCORP.'].notna() & (df['DOCUM.INCORP.'] != '')
            styles.loc[mask_docum, 'DOCUM.INCORP.'] = 'background-color: rgb(173, 216, 230)'  # Azul claro
            
    except Exception as e:
        st.error(f"Error en formato condicional: {e}")
    
    return styles

# Mostrar información del dataset combinado
with st.expander("📊 Información del Dataset Combinado"):
    col1, col2, col3, col4 = st.columns(4)
    
    with col1:
        st.metric("Total Registros", f"{len(df_combinado):,}".replace(",", "."))
    
    with col2:
        st.metric("Total Columnas", len(df_combinado.columns))
    
    with col3:
        archivos_usados = 1
        if archivo_notifica and 'FECHA NOTIFICACIÓN' in df_combinado.columns:
            archivos_usados += 1
        if archivo_triaje and any(col in df_combinado.columns for col in ['USUARIO-CSV', 'CALIFICACIÓN', 'OBSERVACIONES', 'FECHA ASIG']):
            archivos_usados += 1
        if archivo_documentos and 'DOCUM.INCORP.' in df_combinado.columns:
            archivos_usados += 1
        st.metric("Archivos Usados", f"{archivos_usados}/4")
    
    with col4:
        documentos_status = "✅ Cargado" if datos_documentos is not None else "❌ No cargado"
        st.metric("DOCUMENTOS", documentos_status)
    
    # Mostrar primeras filas SIN formato condicional para evitar errores
    st.write("**Vista previa del dataset combinado:**")
    df_mostrar_preview = df_combinado.head(3).copy()
    for col in df_mostrar_preview.select_dtypes(include='datetime').columns:
        df_mostrar_preview[col] = df_mostrar_preview[col].dt.strftime("%d/%m/%Y")
    
    st.dataframe(df_mostrar_preview, use_container_width=True)
    
    # Mostrar columnas disponibles
    st.write("**Columnas disponibles:**")
    columnas_grupos = {}
    for col in df_combinado.columns:
        if col == 'FECHA NOTIFICACIÓN':
            grupo = 'NOTIFICA'
        elif col in ['USUARIO-CSV', 'CALIFICACIÓN', 'OBSERVACIONES', 'FECHA ASIG']:
            grupo = 'TRIAJE'
        elif col == 'DOCUM.INCORP.':
            grupo = 'DOCUMENTACIÓN'
        else:
            grupo = 'RECTAUTO'
        
        if grupo not in columnas_grupos:
            columnas_grupos[grupo] = []
        columnas_grupos[grupo].append(col)
    
    for grupo, columnas in columnas_grupos.items():
        with st.expander(f"📋 Columnas de {grupo} ({len(columnas)})"):
            st.write(columnas)

# Menú principal
menu = ["Principal", "Indicadores clave (KPI)"]
eleccion = st.sidebar.selectbox("Menú", menu)

# Función para gráficos dinámicos (SIN CACHE)
def crear_grafico_dinamico(_conteo, columna, titulo):
    """Crea gráficos dinámicos que responden a los filtros"""
    if _conteo.empty:
        return None
    
    fig = px.bar(_conteo, y=columna, x="Cantidad", title=titulo, text="Cantidad", 
                 color=columna, height=400)
    fig.update_traces(texttemplate='%{text:,}', textposition="auto")
    return fig

# Función para generar PDF de usuario (MEJORADA con ordenamiento)
def generar_pdf_usuario(usuario, df_pendientes, num_semana, fecha_max_str):
    """Genera el PDF para un usuario específico con nombre único"""
    df_user = df_pendientes[df_pendientes["USUARIO"] == usuario].copy()
    
    if df_user.empty:
        return None
    
    # ORDENAR el DataFrame: RUE amarillos primero Y luego por antigüedad
    df_user_ordenado = ordenar_dataframe_por_prioridad_y_antiguedad(df_user)
    
    # Procesar datos para PDF - mantener las columnas originales para el formato condicional
    indices_a_incluir = list(range(df_user_ordenado.shape[1]))
    indices_a_excluir = {1, 4, 5, 6, 13}
    
    # EXCLUIR también la columna "FECHA DE ACTUALIZACIÓN DATOS" si existe
    for idx, col_name in enumerate(df_user_ordenado.columns):
        if "FECHA DE ACTUALIZACIÓN DATOS" in col_name.upper():
            indices_a_excluir.add(idx)
    
    indices_finales = [i for i in indices_a_incluir if i not in indices_a_excluir]
    NOMBRES_COLUMNAS_PDF = df_user_ordenado.columns[indices_finales].tolist()

    # 🔥 CORRECCIÓN: FORZAR COLUMNA ANTIGÜEDAD A ENTEROS
    # Buscar la columna de antigüedad
    columnas_antiguedad = [col for col in df_user_ordenado.columns if 'ANTIGÜEDAD' in col.upper() or 'DÍAS' in col.upper()]
    if columnas_antiguedad:
        columna_antiguedad = columnas_antiguedad[0]
        # Convertir a numérico y luego a enteros, manejando errores
        df_user_ordenado[columna_antiguedad] = pd.to_numeric(df_user_ordenado[columna_antiguedad], errors='coerce').fillna(0).astype(int)

    # Crear DataFrame para mostrar (con fechas formateadas y "nan" reemplazados)
    df_pdf_mostrar = df_user_ordenado[NOMBRES_COLUMNAS_PDF].copy()
    
    # Formatear fechas y reemplazar "nan" - Y ASEGURAR ANTIGÜEDAD COMO ENTERO
    for col in df_pdf_mostrar.columns:
        if df_pdf_mostrar[col].dtype == 'object':
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: "" if pd.isna(x) or str(x).lower() == "nan" else x
            )
        elif 'fecha' in col.lower():
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""
            )
        # 🔥 CORRECCIÓN: FORZAR COLUMNAS NUMÉRICAS A ENTEROS SIN DECIMALES
        elif df_pdf_mostrar[col].dtype in ['float64', 'float32', 'int64', 'int32']:
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: f"{int(x):,}".replace(",", ".") if pd.notna(x) else "0"
            )

    num_expedientes = len(df_pdf_mostrar)
    
    # NOMBRE ÚNICO que incluye timestamp para evitar colisiones
    timestamp = datetime.now().strftime("%H%M%S")
    titulo_pdf = f"{usuario} - Semana {num_semana} a {fecha_max_str} - Expedientes Pendientes ({num_expedientes})"
    
    # Pasar el DataFrame ORIGINAL ORDENADO (con fechas datetime) para el formato condicional
    return dataframe_to_pdf_bytes(df_pdf_mostrar, titulo_pdf, df_original=df_user_ordenado)

if eleccion == "Principal":
    # Usar df_combinado en lugar de df
    df = df_combinado
    
    columna_fecha = df.columns[13]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Sin fecha"
    st.subheader(f"📅 Semana {num_semana} a {fecha_max_str}")

    # Sidebar para filtros
    st.sidebar.header("Filtros")

    if 'filtro_estado' not in st.session_state:
        st.session_state.filtro_estado = ['Abierto'] if 'Abierto' in df['ESTADO'].values else []

    if 'filtro_equipo' not in st.session_state:
        st.session_state.filtro_equipo = sorted(df['EQUIPO'].dropna().unique())

    if 'filtro_usuario' not in st.session_state:
        st.session_state.filtro_usuario = sorted(df['USUARIO'].dropna().unique())

    if st.sidebar.button("Mostrar todos / Resetear filtros"):
        st.session_state.filtro_estado = sorted(df['ESTADO'].dropna().unique())
        st.session_state.filtro_equipo = sorted(df['EQUIPO'].dropna().unique())
        st.session_state.filtro_usuario = sorted(df['USUARIO'].dropna().unique())
        st.rerun()

    opciones_estado = sorted(df['ESTADO'].dropna().unique())
    opciones_equipo = sorted(df['EQUIPO'].dropna().unique())
    opciones_usuario = sorted(df['USUARIO'].dropna().unique())

    estado_sel = st.sidebar.multiselect(
        "Selecciona Estado:",
        options=opciones_estado,
        default=st.session_state.filtro_estado,
        key='filtro_estado'
    )

    equipo_sel = st.sidebar.multiselect(
        "Selecciona Equipo:",
        options=opciones_equipo,
        default=st.session_state.filtro_equipo,
        key='filtro_equipo'
    )

    usuario_sel = st.sidebar.multiselect(
        "Selecciona Usuario:",
        options=opciones_usuario,
        default=st.session_state.filtro_usuario,
        key='filtro_usuario'
    )

    # Aplicar filtros al DataFrame
    df_filtrado = df.copy()

    if estado_sel:
        df_filtrado = df_filtrado[df_filtrado['ESTADO'].isin(estado_sel)]

    if equipo_sel:
        df_filtrado = df_filtrado[df_filtrado['EQUIPO'].isin(equipo_sel)]

    if usuario_sel:
        df_filtrado = df_filtrado[df_filtrado['USUARIO'].isin(usuario_sel)]

    # Mostrar filtros activos
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filtros activos")
    if estado_sel:
        st.sidebar.write(f"Estados: {', '.join(estado_sel)}")
    if equipo_sel:
        st.sidebar.write(f"Equipos: {len(equipo_sel)} seleccionados")
    if usuario_sel:
        st.sidebar.write(f"Usuarios: {len(usuario_sel)} seleccionados")

    # NUEVO: Opciones de ordenamiento y auto-filtros
    st.sidebar.markdown("---")
    st.sidebar.subheader("Opciones de Visualización")
    
    # Checkbox para ordenar por prioridad
    ordenar_prioridad = st.sidebar.checkbox("Ordenar por prioridad (RUE amarillo primero)", value=True)
    
    # Auto-filtros para mostrar solo filas con formato condicional
    st.sidebar.markdown("**Auto-filtros:**")
    mostrar_solo_amarillos = st.sidebar.checkbox("Solo RUE prioritarios", value=False)
    mostrar_solo_rojos = st.sidebar.checkbox("Solo USUARIO-CSV discrepantes", value=False)
    mostrar_solo_docum = st.sidebar.checkbox("Solo con DOCUM.INCORP.", value=False)

    # Aplicar ordenamiento si está activado
    if ordenar_prioridad:
        df_filtrado = ordenar_dataframe_por_prioridad_y_antiguedad(df_filtrado)

    # Aplicar auto-filtros
    if mostrar_solo_amarillos or mostrar_solo_rojos or mostrar_solo_docum:
        df_filtrado_temp = df_filtrado.copy()
        
        if mostrar_solo_amarillos:
            # Filtrar solo RUE amarillos - CON TODAS LAS CONDICIONES
            mask_amarillo = pd.Series(False, index=df_filtrado_temp.index)
            for idx, row in df_filtrado_temp.iterrows():
                try:
                    etiq_penultimo = row.get('ETIQ. PENÚLTIMO TRAM.', '')
                    fecha_notif = row.get('FECHA NOTIFICACIÓN', None)
                    docum_incorp = row.get('DOCUM.INCORP.', '')
                    
                    # CONDICIÓN 1: "80 PROPRES" con fecha límite superada
                    if (str(etiq_penultimo).strip() == "80 PROPRES" and 
                        pd.notna(fecha_notif) and 
                        isinstance(fecha_notif, (pd.Timestamp, datetime))):
                        
                        fecha_limite = fecha_notif + timedelta(days=23)
                        if datetime.now() > fecha_limite:
                            mask_amarillo[idx] = True
                    
                    # CONDICIÓN 2: "50 REQUERIR" con fecha límite superada
                    elif (str(etiq_penultimo).strip() == "50 REQUERIR" and 
                        pd.notna(fecha_notif) and 
                        isinstance(fecha_notif, (pd.Timestamp, datetime))):
                        
                        fecha_limite = fecha_notif + timedelta(days=23)
                        if datetime.now() > fecha_limite:
                            mask_amarillo[idx] = True
                    
                    # CONDICIÓN 3: "70 ALEGACI" o "60 CONTESTA"
                    elif str(etiq_penultimo).strip() in ["70 ALEGACI", "60 CONTESTA"]:
                        mask_amarillo[idx] = True
                    
                    # CONDICIÓN 4: DOCUM.INCORP. no vacío Y distinto de "SOLICITUD" o "REITERA SOLICITUD"
                    elif (pd.notna(docum_incorp) and 
                        str(docum_incorp).strip() != '' and
                        str(docum_incorp).strip().upper() not in ["SOLICITUD", "REITERA SOLICITUD"]):
                        mask_amarillo[idx] = True
                        
                except:
                    pass
            
            df_filtrado_temp = df_filtrado_temp[mask_amarillo]
        
        if mostrar_solo_rojos:
            # Filtrar solo USUARIO-CSV rojos
            if 'USUARIO' in df_filtrado_temp.columns and 'USUARIO-CSV' in df_filtrado_temp.columns:
                mask_rojo = df_filtrado_temp['USUARIO'] != df_filtrado_temp['USUARIO-CSV']
                df_filtrado_temp = df_filtrado_temp[mask_rojo]
        
        if mostrar_solo_docum:
            # Filtrar solo expedientes con documentación incorporada
            if 'DOCUM.INCORP.' in df_filtrado_temp.columns:
                mask_docum = df_filtrado_temp['DOCUM.INCORP.'].notna() & (df_filtrado_temp['DOCUM.INCORP.'] != '')
                df_filtrado_temp = df_filtrado_temp[mask_docum]
        
        df_filtrado = df_filtrado_temp

    # Gráficos Generales - CORREGIDOS: datos siempre frescos según filtros
    st.subheader("📈 Gráficos Generales")
    columnas_graficos = st.columns(3)
    graficos = [("EQUIPO", "Expedientes por equipo"), 
                ("USUARIO", "Expedientes por usuario"), 
                ("ESTADO", "Distribución por estado")]

    for i, (col, titulo) in enumerate(graficos):
        if col in df_filtrado.columns:
            # Calcular el conteo actual (siempre fresco según los filtros)
            conteo_actual = df_filtrado[col].value_counts().reset_index()
            conteo_actual.columns = [col, "Cantidad"]
            
            # Crear gráfico con datos actualizados (SIN CACHE)
            fig = crear_grafico_dinamico(conteo_actual, col, titulo)
            if fig:
                columnas_graficos[i].plotly_chart(fig, use_container_width=True)

    if "NOTIFICADO" in df_filtrado.columns:
        conteo_notificado = df_filtrado["NOTIFICADO"].value_counts().reset_index()
        conteo_notificado.columns = ["NOTIFICADO", "Cantidad"]
        fig = crear_grafico_dinamico(conteo_notificado, "NOTIFICADO", "Expedientes notificados")
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    # VISTA GENERAL - SOLUCIÓN STREAMLIT NATIVO OPTIMIZADA
    st.subheader("📋 Vista general de expedientes")

    # Crear copia y formatear fechas
    df_mostrar = df_filtrado.copy()

    # Formatear TODAS las columnas de fecha
    for col in df_mostrar.select_dtypes(include='datetime').columns:
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%Y")

    # Mostrar tabla principal SIN formato condicional pero CON fechas formateadas
    st.dataframe(df_mostrar, use_container_width=True, height=400)
    registros_mostrados = f"{len(df_mostrar):,}".replace(",", ".")
    registros_totales = f"{len(df):,}".replace(",", ".")
    st.write(f"Mostrando {registros_mostrados} de {registros_totales} registros")

    # Sección de RESUMEN de formatos condicionales
    st.markdown("---")
    st.subheader("🔍 Resumen de Expedientes con Formatos Condicionales")

    # Crear pestañas para cada tipo de formato condicional
    tab1, tab2, tab3 = st.tabs(["🟡 RUE Prioritarios", "🔴 USUARIO-CSV Discrepantes", "🔵 Con Documentación"])

    with tab1:
        # RUE amarillos - CON TODAS LAS CONDICIONES
        st.write("**Expedientes con RUE prioritario (amarillo):**")
        df_amarillos = df_filtrado.copy()
        mask_amarillo = pd.Series(False, index=df_amarillos.index)
        
        for idx, row in df_amarillos.iterrows():
            try:
                etiq_penultimo = row.get('ETIQ. PENÚLTIMO TRAM.', '')
                fecha_notif = row.get('FECHA NOTIFICACIÓN', None)
                docum_incorp = row.get('DOCUM.INCORP.', '')
                
                # CONDICIÓN 1: "80 PROPRES" con fecha límite superada
                if (str(etiq_penultimo).strip() == "80 PROPRES" and 
                    pd.notna(fecha_notif) and 
                    isinstance(fecha_notif, (pd.Timestamp, datetime))):
                    
                    fecha_limite = fecha_notif + timedelta(days=23)
                    if datetime.now() > fecha_limite:
                        mask_amarillo[idx] = True
                
                # CONDICIÓN 2: "50 REQUERIR" con fecha límite superada
                elif (str(etiq_penultimo).strip() == "50 REQUERIR" and 
                    pd.notna(fecha_notif) and 
                    isinstance(fecha_notif, (pd.Timestamp, datetime))):
                    
                    fecha_limite = fecha_notif + timedelta(days=23)
                    if datetime.now() > fecha_limite:
                        mask_amarillo[idx] = True
                
                # CONDICIÓN 3: "70 ALEGACI" o "60 CONTESTA"
                elif str(etiq_penultimo).strip() in ["70 ALEGACI", "60 CONTESTA"]:
                    mask_amarillo[idx] = True
                
                # CONDICIÓN 4: DOCUM.INCORP. no vacío Y distinto de "SOLICITUD" o "REITERA SOLICITUD"
                elif (pd.notna(docum_incorp) and 
                    str(docum_incorp).strip() != '' and
                    str(docum_incorp).strip().upper() not in ["SOLICITUD", "REITERA SOLICITUD"]):
                    mask_amarillo[idx] = True
                    
            except:
                pass
        
        # FILTRAR Y MOSTRAR EL DATAFRAME - CORREGIDO: APLICAR ORDENAMIENTO
        if mask_amarillo.any():
            df_temp = df_amarillos[mask_amarillo].copy()
            
            # 🔥 CORRECCIÓN: APLICAR ORDENAMIENTO POR PRIORIDAD
            df_temp = ordenar_dataframe_por_prioridad_y_antiguedad(df_temp)
            
            # Formatear fechas para mostrar
            for col in df_temp.select_dtypes(include='datetime').columns:
                df_temp[col] = df_temp[col].dt.strftime("%d/%m/%Y")
            
            st.dataframe(df_temp, use_container_width=True)
            st.warning(f"**Total: {mask_amarillo.sum()} expedientes prioritarios**")
        else:
            st.success("✅ No hay expedientes con RUE prioritario")

    with tab2:
        # USUARIO-CSV rojos
        st.write("**Expedientes con discrepancia USUARIO/USUARIO-CSV (rojo):**")
        
        if 'USUARIO' in df_filtrado.columns and 'USUARIO-CSV' in df_filtrado.columns:
            mask_rojo = df_filtrado['USUARIO'] != df_filtrado['USUARIO-CSV']
            
            if mask_rojo.any():
                df_temp = df_filtrado[mask_rojo].copy()
                # Formatear fechas para mostrar
                for col in df_temp.select_dtypes(include='datetime').columns:
                    df_temp[col] = df_temp[col].dt.strftime("%d/%m/%Y")
                
                # Mostrar solo columnas relevantes
                columnas_relevantes = ['RUE', 'USUARIO', 'USUARIO-CSV', 'EQUIPO', 'ESTADO']
                columnas_disponibles = [col for col in columnas_relevantes if col in df_temp.columns]
                
                st.dataframe(df_temp[columnas_disponibles], use_container_width=True)
                st.error(f"**Total: {mask_rojo.sum()} expedientes con discrepancia**")
                
                # Análisis de discrepancias
                st.write("**Análisis de discrepancias:**")
                discrepancia_detalle = df_temp.groupby(['USUARIO', 'USUARIO-CSV']).size().reset_index(name='Cantidad')
                st.dataframe(discrepancia_detalle, use_container_width=True)
            else:
                st.success("✅ No hay discrepancias entre USUARIO y USUARIO-CSV")
        else:
            st.info("Columnas USUARIO y/o USUARIO-CSV no disponibles")

    with tab3:
        # DOCUM.INCORP. azules
        st.write("**Expedientes con documentación incorporada (azul):**")
        
        if 'DOCUM.INCORP.' in df_filtrado.columns:
            mask_docum = df_filtrado['DOCUM.INCORP.'].notna() & (df_filtrado['DOCUM.INCORP.'] != '')
            
            if mask_docum.any():
                df_temp = df_filtrado[mask_docum].copy()
                # Formatear fechas para mostrar
                for col in df_temp.select_dtypes(include='datetime').columns:
                    df_temp[col] = df_temp[col].dt.strftime("%d/%m/%Y")
                
                # Mostrar solo columnas relevantes
                columnas_relevantes = ['RUE', 'DOCUM.INCORP.', 'USUARIO', 'EQUIPO', 'ETIQ. PENÚLTIMO TRAM.']
                columnas_disponibles = [col for col in columnas_relevantes if col in df_temp.columns]
                
                st.dataframe(df_temp[columnas_disponibles], use_container_width=True)
                st.info(f"**Total: {mask_docum.sum()} expedientes con documentación**")
                
                # Análisis por tipo de documentación
                if 'DOCUM.INCORP.' in df_temp.columns:
                    st.write("**Tipos de documentación:**")
                    conteo_docum = df_temp['DOCUM.INCORP.'].value_counts()
                    for doc_type, count in conteo_docum.head(10).items():  # Mostrar top 10
                        st.write(f"- {doc_type}: {count} expedientes")
            else:
                st.info("No hay expedientes con documentación incorporada")
        else:
            st.info("Columna DOCUM.INCORP. no disponible")

    # Estadísticas generales
    st.markdown("---")
    st.subheader("📊 Estadísticas Generales")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        registros_mostrados = f"{len(df_mostrar):,}".replace(",", ".")
        registros_totales = f"{len(df):,}".replace(",", ".")
        st.metric("Registros mostrados", f"{registros_mostrados}/{registros_totales}")

    with col2:
        # Contar RUE amarillos - CON NUEVAS CONDICIONES
        mask_amarillo = pd.Series(False, index=df_filtrado.index)
        for idx, row in df_filtrado.iterrows():
            try:
                etiq_penultimo = row.get('ETIQ. PENÚLTIMO TRAM.', '')
                fecha_notif = row.get('FECHA NOTIFICACIÓN', None)
                docum_incorp = row.get('DOCUM.INCORP.', '')
                
                # CONDICIÓN 1: "80 PROPRES" con fecha límite superada
                if (str(etiq_penultimo).strip() == "80 PROPRES" and 
                    pd.notna(fecha_notif) and 
                    isinstance(fecha_notif, (pd.Timestamp, datetime))):
                    
                    fecha_limite = fecha_notif + timedelta(days=23)
                    if datetime.now() > fecha_limite:
                        mask_amarillo[idx] = True
                
                # NUEVA CONDICIÓN 2: "50 REQUERIR" con fecha límite superada
                elif (str(etiq_penultimo).strip() == "50 REQUERIR" and 
                    pd.notna(fecha_notif) and 
                    isinstance(fecha_notif, (pd.Timestamp, datetime))):
                    
                    fecha_limite = fecha_notif + timedelta(days=23)
                    if datetime.now() > fecha_limite:
                        mask_amarillo[idx] = True
                
                # NUEVA CONDICIÓN 3: "70 ALEGACI" o "60 CONTESTA"
                elif str(etiq_penultimo).strip() in ["70 ALEGACI", "60 CONTESTA"]:
                    mask_amarillo[idx] = True
                
                # NUEVA CONDICIÓN 4: DOCUM.INCORP. no vacío Y distinto de "SOLICITUD" o "REITERA SOLICITUD"
                elif (pd.notna(docum_incorp) and 
                    str(docum_incorp).strip() != '' and
                    str(docum_incorp).strip().upper() not in ["SOLICITUD", "REITERA SOLICITUD"]):
                    mask_amarillo[idx] = True
                    
            except:
                pass
        st.metric("RUE prioritarios", f"{mask_amarillo.sum():,}".replace(",", "."))

    with col3:
        # Contar USUARIO-CSV rojos
        if 'USUARIO' in df_filtrado.columns and 'USUARIO-CSV' in df_filtrado.columns:
            mask_rojo = (df_filtrado['USUARIO'] != df_filtrado['USUARIO-CSV']).sum()
            st.metric("Discrepancias", f"{mask_rojo:,}".replace(",", "."))
        else:
            st.metric("Discrepancias", "N/A")

    with col4:
        # Contar DOCUM.INCORP. azules
        if 'DOCUM.INCORP.' in df_filtrado.columns:
            mask_docum = df_filtrado['DOCUM.INCORP.'].notna().sum()
            st.metric("Con documentación", f"{mask_docum:,}".replace(",", "."))
        else:
            st.metric("Con documentación", "N/A")

    # Mostrar estadísticas de los filtros aplicados
    st.sidebar.markdown("---")
    st.sidebar.subheader("Estadísticas")
    
    # Contar filas prioritarias
    df_priorizado = identificar_filas_prioritarias(df_filtrado)
    filas_amarillas = df_priorizado['_prioridad'].sum()
    filas_totales = len(df_filtrado)
    
    st.sidebar.write(f"Total filas: {filas_totales}")
    st.sidebar.write(f"RUE prioritarios: {filas_amarillas}")
    
    # Contar USUARIO-CSV rojos
    if 'USUARIO' in df_filtrado.columns and 'USUARIO-CSV' in df_filtrado.columns:
        filas_rojas = (df_filtrado['USUARIO'] != df_filtrado['USUARIO-CSV']).sum()
        st.sidebar.write(f"USUARIO-CSV discrepantes: {filas_rojas}")
    
    # Contar DOCUM.INCORP.
    if 'DOCUM.INCORP.' in df_filtrado.columns:
        filas_docum = df_filtrado['DOCUM.INCORP.'].notna().sum()
        st.sidebar.write(f"Con DOCUM.INCORP.: {filas_docum}")

    registros_mostrados = f"{len(df_mostrar):,}".replace(",", ".")
    registros_totales = f"{len(df):,}".replace(",", ".")
    st.write(f"Mostrando {registros_mostrados} de {registros_totales} registros")

    # NUEVA SECCIÓN: GESTIÓN DE DOCUMENTACIÓN INCORPORADA
    st.markdown("---")
    st.header("📄 Gestión de Documentación Incorporada")

    if archivo_documentos and datos_documentos:
        # Filtro para expedientes con ETIQ. PENÚLTIMO TRAM. = "90 INCDOCU"
        df_incdocu = df_filtrado[
            df_filtrado['ETIQ. PENÚLTIMO TRAM.'] == "90 INCDOCU"
        ].copy()
        
        if df_incdocu.empty:
            st.info("ℹ️ No hay expedientes con ETIQ. PENÚLTIMO TRAM. = '90 INCDOCU' en los filtros actuales")
        else:
            st.success(f"📋 Encontrados {len(df_incdocu)} expedientes con '90 INCDOCU'")
            
            # Obtener opciones del desplegable
            opciones_docu = datos_documentos['opciones']
            opciones_combo = [""] + opciones_docu  # Añadir opción vacía
            
            # Crear interfaz para editar la documentación
            st.subheader("Editar Documentación Incorporada")
            
            # Mostrar tabla editable
            for idx, row in df_incdocu.iterrows():
                rue = row['RUE']
                docum_actual = row.get('DOCUM.INCORP.', '')
                
                col1, col2, col3 = st.columns([2, 3, 1])
                
                with col1:
                    st.write(f"**RUE:** {rue}")
                
                with col2:
                    # Selectbox para elegir documentación
                    clave_docum = f"docum_{rue}"
                    nueva_docum = st.selectbox(
                        "Documentación:",
                        options=opciones_combo,
                        index=opciones_combo.index(docum_actual) if docum_actual in opciones_combo else 0,
                        key=clave_docum,
                        label_visibility="collapsed"
                    )
                    
                    # Guardar cambio en session_state
                    if nueva_docum != docum_actual:
                        st.session_state.cambios_documentacion[rue] = nueva_docum
                
                with col3:
                    if st.button("💾 Guardar", key=f"btn_{rue}"):
                        # Actualizar el DataFrame combinado
                        df_combinado.loc[df_combinado['RUE'] == rue, 'DOCUM.INCORP.'] = nueva_docum
                        st.session_state["df_combinado"] = df_combinado
                        st.success(f"✅ Documentación actualizada para RUE {rue}")
            
            # Botón para guardar todos los cambios en el archivo DOCUMENTOS.xlsx
            st.markdown("---")
            col1, col2 = st.columns([3, 1])
            
            with col2:
                # En la sección de guardar documentos, busca esta parte y actualízala:
                if st.button("💾 Guardar Todos los Cambios en DOCUMENTOS.xlsx", type="primary", key="guardar_documentos"):
                    with st.spinner("Guardando cambios..."):
                        # Crear DataFrame SOLO con los registros que tienen DOCUM.INCORP. actualmente
                        df_documentos_actualizado = df_combinado[['RUE', 'DOCUM.INCORP.']].copy()
                        df_documentos_actualizado = df_documentos_actualizado.dropna(subset=['DOCUM.INCORP.'])
                        df_documentos_actualizado = df_documentos_actualizado[df_documentos_actualizado['DOCUM.INCORP.'] != '']
                        
                        # Guardar en el archivo DOCUMENTOS.xlsx (esto reemplazará completamente el contenido anterior)
                        contenido_actualizado = guardar_documentos_actualizados(
                            archivo_documentos, 
                            df_documentos_actualizado
                        )
                        
                        if contenido_actualizado:
                            st.session_state.documentos_actualizados = contenido_actualizado
                            st.session_state.mostrar_descarga = True
                            st.success("✅ Archivo DOCUMENTOS.xlsx actualizado correctamente")
                            
                            # Limpiar cambios
                            st.session_state.cambios_documentacion = {}
                            
                            # Actualizar cache
                            st.cache_data.clear()
                            
                        else:
                            st.error("❌ Error al guardar el archivo DOCUMENTOS.xlsx")
            
            # Mostrar botón de descarga si hay archivo actualizado
            if st.session_state.get('mostrar_descarga', False) and st.session_state.get('documentos_actualizados'):
                st.markdown("---")
                st.subheader("📥 Descargar Archivo Actualizado")
                
                # Generar nombre de archivo con timestamp
                from datetime import datetime
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                nombre_archivo = f"DOCUMENTOS_actualizado_{timestamp}.xlsx"
                
                st.download_button(
                    label="⬇️ Descargar DOCUMENTOS.xlsx actualizado",
                    data=st.session_state.documentos_actualizados,
                    file_name=nombre_archivo,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="descarga_documentos"
                )
                
                # Opción para continuar sin descargar
                if st.button("Continuar sin descargar", key="continuar_sin_descargar"):
                    st.session_state.mostrar_descarga = False
                    st.rerun()
    else:
        st.warning("⚠️ Carga el archivo DOCUMENTOS.xlsx para gestionar la documentación incorporada")

    # Descarga de informes
    st.markdown("---")
    st.header("Descarga de Informes")
    st.subheader("Generar Informes PDF Pendientes por Usuario")

    df_pendientes = df[df["ESTADO"].isin(ESTADOS_PENDIENTES)].copy()
    usuarios_pendientes = df_pendientes["USUARIO"].dropna().unique()

    # NUEVO: Generar también PDFs por equipo (solo prioritarios) y resumen KPI
    equipos_pendientes = df_pendientes["EQUIPO"].dropna().unique()
    
    if st.button(f"Generar {len(usuarios_pendientes)} Informes PDF + Equipos + Resumen KPI", key="generar_pdfs_completos"):
        if usuarios_pendientes.size == 0:
            st.info("No se encontraron expedientes pendientes para generar informes.")
        else:
            with st.spinner('Generando PDFs y comprimiendo...'):
                zip_buffer = io.BytesIO()

            with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
                # 1. PDFs por usuario (todos los pendientes)
                for usuario in usuarios_pendientes:
                    pdf_data = generar_pdf_usuario(usuario, df_pendientes, num_semana, fecha_max_str)
                    if pdf_data:
                        file_name = f"{num_semana}{usuario}.pdf"
                        zip_file.writestr(file_name, pdf_data)
                
                # 2. PDFs por equipo (solo expedientes prioritarios)
                for equipo in equipos_pendientes:
                    pdf_data = generar_pdf_equipo_prioritarios(equipo, df_pendientes, num_semana, fecha_max_str)
                    if pdf_data:
                        file_name = f"{num_semana}{equipo}_PRIORITARIOS.pdf"
                        zip_file.writestr(file_name, pdf_data)
                
                # 3. PDF de resumen de KPIs
                # Calcular KPIs para la semana actual
                columna_fecha = df.columns[13]
                df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
                fecha_max = df[columna_fecha].max()
                
                # Crear rango de semanas disponibles
                fecha_inicio = pd.to_datetime("2022-11-01")
                semanas_disponibles = pd.date_range(
                    start=fecha_inicio,
                    end=fecha_max,
                    freq='W-FRI'
                ).tolist()
                
                # Calcular KPIs para todas las semanas
                df_kpis_semanales = calcular_kpis_todas_semanas_optimizado(df, semanas_disponibles, FECHA_REFERENCIA, fecha_max)
                
                # Generar PDF de resumen KPI
                pdf_resumen = generar_pdf_resumen_kpi(
                    df_kpis_semanales, 
                    num_semana, 
                    fecha_max_str, 
                    df_combinado, 
                    semanas_disponibles, 
                    FECHA_REFERENCIA, 
                    fecha_max
                )
                if pdf_resumen:
                    file_name = f"{num_semana}_RESUMEN_KPI.pdf"
                    zip_file.writestr(file_name, pdf_resumen)

            zip_buffer.seek(0)
            zip_file_name = f"Informes_Completos_Semana_{num_semana}.zip"
            st.download_button(
                label=f"⬇️ Descargar {len(usuarios_pendientes)} Informes PDF + Equipos + Resumen KPI (ZIP)",
                data=zip_buffer.read(),
                file_name=zip_file_name,
                mime="application/zip",
                help="Descarga todos los informes PDF listos.",
                key='pdf_download_button_completo'
            )

    # SECCIÓN: ENVÍO DE CORREOS INTEGRADA
    st.markdown("---")
    st.header("📧 Envío de Correos")
    
    # Verificar que estamos usando la última semana
    st.info(f"**📅 Semana activa para envío:** {num_semana} (Última semana disponible - {fecha_max_str})")

    # Definir si está disponible el envío de resumen
    envio_resumen_disponible = 'ENVÍO RESUMEN' in df_usuarios.columns
    if envio_resumen_disponible:
        st.info("✅ **Envío de resumen KPI disponible** - Los usuarios con 'ENVÍO RESUMEN = SÍ' recibirán el resumen")
    else:
        st.info("ℹ️ **Envío de resumen KPI no disponible** - La columna 'ENVÍO RESUMEN' no existe en USUARIOS.xlsx")
    
    # Verificar si el archivo USUARIOS está cargado
    if df_usuarios is None:
        st.error("❌ No se ha cargado el archivo USUARIOS. Por favor, cárgalo en la sección de arriba.")
        st.stop()
        envio_resumen_disponible = False
    
    # Verificar columnas requeridas en USUARIOS
    columnas_requeridas = ['USUARIOS', 'ENVIAR', 'EMAIL', 'ASUNTO', 'MENSAJE1']
    columnas_faltantes = [col for col in columnas_requeridas if col not in df_usuarios.columns]
    
    if columnas_faltantes:
        st.error(f"❌ Faltan columnas en el archivo USUARIOS: {', '.join(columnas_faltantes)}")
        st.stop()
    
    # NUEVO: Verificar columna para envío de resumen
    def verificar_envio_resumen(usuario_row):
        """Verifica si el usuario debe recibir resumen KPI de forma segura"""
        try:
            # Verificar si la columna existe y tiene valor
            if 'ENVÍO RESUMEN' not in usuario_row.index or pd.isna(usuario_row['ENVÍO RESUMEN']):
                return False
            
            valor = str(usuario_row['ENVÍO RESUMEN']).strip().upper()
            return valor in ['SÍ', 'SI', 'S', 'YES', 'Y', 'TRUE', '1']
        except:
            return False

    # Filtrar usuarios activos
    usuarios_activos = df_usuarios[
        (df_usuarios['ENVIAR'].str.upper() == 'SÍ') | 
        (df_usuarios['ENVIAR'].str.upper() == 'SI')
    ]
    
    if usuarios_activos.empty:
        st.warning("⚠️ No hay usuarios activos para envío (ENVIAR = 'SÍ' o 'SI')")
    else:
        # Función para generar el cuerpo del mensaje dinámicamente
        def generar_cuerpo_mensaje(mensaje_base):
            """Genera el cuerpo del mensaje con saludo según la hora"""
            from datetime import datetime
            
            hora_actual = datetime.now().hour
            saludo = "Buenos días" if hora_actual < 14 else "Buenas tardes"
            
            cuerpo_mensaje = f"{saludo},\n\n{mensaje_base}"
            return cuerpo_mensaje
        
        # Función para procesar el asunto con variables
        def procesar_asunto(asunto_template, num_semana, fecha_max_str):
            """Reemplaza variables en el asunto del correo"""
            asunto_procesado = asunto_template.replace("&num_semana&", str(num_semana))
            asunto_procesado = asunto_procesado.replace("&fecha_max&", fecha_max_str)
            return asunto_procesado
        
        # Función para enviar correos con Outlook (funciona con Outlook cerrado)
        def enviar_correo_outlook(destinatario, asunto, cuerpo_mensaje, archivos_adjuntos, cc=None, bcc=None):
            """
            Envía correo usando Outlook local con múltiples adjuntos.
            archivos_adjuntos: lista de tuplas (nombre_archivo, datos_archivo)
            """
            try:
                import win32com.client
                import os
                import tempfile
                import shutil

                outlook = win32com.client.Dispatch("Outlook.Application")
                mail = outlook.CreateItem(0)  # 0 = olMailItem

                mail.To = destinatario
                mail.Subject = asunto
                mail.Body = cuerpo_mensaje

                if cc and pd.notna(cc) and str(cc).strip():
                    mail.CC = str(cc)
                if bcc and pd.notna(bcc) and str(bcc).strip():
                    mail.BCC = str(bcc)

                # Adjuntar archivos
                rutas_temporales = []
                for nombre_archivo, datos_archivo in archivos_adjuntos:
                    # Crear archivo temporal
                    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_file:
                        temp_file.write(datos_archivo)
                        temp_path = temp_file.name
                    
                    # Crear copia con nombre real
                    carpeta_temp = tempfile.gettempdir()
                    ruta_final = os.path.join(carpeta_temp, nombre_archivo)
                    
                    # Asegurar extensión .pdf
                    if not ruta_final.lower().endswith(".pdf"):
                        ruta_final += ".pdf"
                    
                    shutil.copy(temp_path, ruta_final)
                    rutas_temporales.append((temp_path, ruta_final))
                    
                    # Adjuntar el archivo
                    mail.Attachments.Add(Source=ruta_final)

                # Enviar correo
                mail.Send()

                # Limpieza
                for temp_path, ruta_final in rutas_temporales:
                    try:
                        os.remove(temp_path)
                        os.remove(ruta_final)
                    except:
                        pass

                return True

            except ImportError:
                st.error("❌ Error: win32com.client no está disponible. Instala pywin32: pip install pywin32")
                return False
            except Exception as e:
                st.error(f"❌ Error al enviar correo a {destinatario}: {e}")
                return False
        
        # Verificar usuarios con expedientes pendientes
        usuarios_con_pendientes = df_pendientes['USUARIO'].dropna().unique()
        usuarios_para_envio = []
        
        # === FUNCIÓN AUXILIAR - COLOCAR ANTES DEL BUCLE ===
        def verificar_envio_resumen(usuario_row):
            """Verifica si el usuario debe recibir resumen KPI de forma segura"""
            try:
                # Verificar si la columna existe y tiene valor
                if 'ENVÍO RESUMEN' not in usuario_row.index or pd.isna(usuario_row['ENVÍO RESUMEN']):
                    return False
                
                valor = str(usuario_row['ENVÍO RESUMEN']).strip().upper()
                return valor in ['SÍ', 'SI', 'S', 'YES', 'Y', 'TRUE', '1']
            except:
                return False
        # DEBUG: Mostrar información de depuración
        st.sidebar.markdown("---")
        st.sidebar.subheader("🔍 Depuración Envío Correos")
        st.sidebar.write(f"Usuarios activos: {len(usuarios_activos)}")
        st.sidebar.write(f"Usuarios con pendientes: {len(usuarios_con_pendientes)}")
        st.sidebar.write(f"Ejemplo usuarios activos: {list(usuarios_activos['USUARIOS'].head(3))}")
        st.sidebar.write(f"Ejemplo usuarios con pendientes: {list(usuarios_con_pendientes[:3])}")

        # === BUCLE CORREGIDO ===
        for _, usuario_row in usuarios_activos.iterrows():
            usuario = usuario_row['USUARIOS']
            # VERIFICACIÓN MÁS ROBUSTA - compara normalizando strings
            usuario_encontrado = False
            for usuario_pendiente in usuarios_con_pendientes:
                if str(usuario).strip().upper() == str(usuario_pendiente).strip().upper():
                    usuario_encontrado = True
                    break
            
            if usuario_encontrado:
                num_expedientes = len(df_pendientes[
                    df_pendientes['USUARIO'].apply(lambda x: str(x).strip().upper() if pd.notna(x) else '') == str(usuario).strip().upper()
                ])
                
                # Procesar asunto con variables
                asunto_template = usuario_row['ASUNTO'] if pd.notna(usuario_row['ASUNTO']) else f"Situación RECTAUTO asignados en la semana {num_semana} a {fecha_max_str}"
                asunto_procesado = procesar_asunto(asunto_template, num_semana, fecha_max_str)
                
                # Generar cuerpo del mensaje
                mensaje_base = f"{usuario_row['MENSAJE1']} \n\n {usuario_row['MENSAJE2']} \n\n {usuario_row['MENSAJE3']} \n\n {usuario_row['DESPEDIDA']} \n\n __________________ \n\n Equipo RECTAUTO." if pd.notna(usuario_row['MENSAJE1']) else "Se adjunta informe de expedientes pendientes."
                cuerpo_mensaje = generar_cuerpo_mensaje(mensaje_base)
                
                # Verificar si debe recibir resumen
                # Definir envio_resumen_disponible basado en si existe la columna ENVÍO RESUMEN
                envio_resumen_disponible = 'ENVÍO RESUMEN' in df_usuarios.columns
                recibir_resumen = verificar_envio_resumen(usuario_row) if envio_resumen_disponible else False
                
                usuarios_para_envio.append({
                    'usuario': usuario,
                    'resumen': usuario_row.get('RESUMEN', ''),
                    'email': usuario_row['EMAIL'],
                    'cc': usuario_row.get('CC', ''),
                    'bcc': usuario_row.get('BCC', ''),
                    'expedientes': num_expedientes,
                    'asunto': asunto_procesado,
                    'mensaje': mensaje_base,
                    'cuerpo_mensaje': cuerpo_mensaje,
                    'recibir_resumen': recibir_resumen
                })
            else:
                st.sidebar.warning(f"Usuario {usuario} no encontrado en pendientes")

        # MOSTRAR MÁS INFORMACIÓN DE DEPURACIÓN
        if not usuarios_para_envio:
            st.error("❌ DEPURACIÓN: No se encontraron coincidencias entre usuarios activos y pendientes")
            st.write("**Usuarios activos:**", list(usuarios_activos['USUARIOS'].values))
            st.write("**Usuarios con pendientes:**", list(usuarios_con_pendientes))
        else:
            st.success(f"✅ Se encontraron {len(usuarios_para_envio)} usuarios para envío")
        
        if usuarios_para_envio:
            st.success(f"✅ {len(usuarios_para_envio)} usuarios tienen expedientes pendientes para enviar")
            
            # Mostrar resumen
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Usuarios activos", len(usuarios_activos))
            with col2:
                st.metric("Con expedientes", len(usuarios_para_envio))
            with col3:
                st.metric("Semana activa", f"Semana {num_semana}")
            
            # Mostrar tabla de usuarios para envío
            with st.expander("📋 Ver detalles de usuarios para envío"):
                df_envio = pd.DataFrame(usuarios_para_envio)
                columnas_mostrar = ['usuario', 'resumen', 'email', 'expedientes', 'asunto', 'recibir_resumen']
                st.dataframe(df_envio[columnas_mostrar], use_container_width=True)
            
            # Previsualización de correo
            st.subheader("👁️ Previsualización del Correo")
            
            if usuarios_para_envio:
                usuario_ejemplo = usuarios_para_envio[0]
                col1, col2 = st.columns([1, 2])
                
                with col1:
                    st.write("**Destinatario:**", usuario_ejemplo['email'])
                    if usuario_ejemplo['cc']:
                        st.write("**CC:**", usuario_ejemplo['cc'])
                    if usuario_ejemplo['bcc']:
                        st.write("**BCC:**", usuario_ejemplo['bcc'])
                    st.write("**Asunto:**", usuario_ejemplo['asunto'])
                    st.write("**Expedientes:**", usuario_ejemplo['expedientes'])
                    st.write("**Recibir resumen:**", "✅ Sí" if usuario_ejemplo['recibir_resumen'] else "❌ No")
                
                with col2:
                    st.text_area("**Cuerpo del Mensaje:**", usuario_ejemplo['cuerpo_mensaje'], height=200, key="preview_mensaje")
            
            # Botón de envío masivo
            st.markdown("---")
            st.subheader("🚀 Envío de Correos")
            
            st.warning("""
            **⚠️ Importante antes de enviar:**
            - Se usará la cuenta de Outlook predeterminada
            - No es necesario tener Outlook abierto
            - Los correos se enviarán inmediatamente
            - Se adjuntará el PDF individual de cada usuario
            - **Los que tengan 'ENVÍO RESUMEN = SÍ' recibirán también el resumen KPI**
            - **Verifica que los datos sean correctos**
            """)
            
            if st.button("📤 Enviar Correos a Todos los Usuarios", type="primary", key="enviar_correos"):
                progress_bar = st.progress(0)
                status_text = st.empty()
                
                correos_enviados = 0
                correos_fallidos = 0
                
                # Generar PDF de resumen KPI (una sola vez para todos)
                pdf_resumen = None
                if envio_resumen_disponible:
                    # Calcular KPIs para la semana actual
                    columna_fecha = df.columns[13]
                    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
                    fecha_max = df[columna_fecha].max()
                    
                    fecha_inicio = pd.to_datetime("2022-11-01")
                    semanas_disponibles = pd.date_range(
                        start=fecha_inicio,
                        end=fecha_max,
                        freq='W-FRI'
                    ).tolist()
                    
                    df_kpis_semanales = calcular_kpis_todas_semanas_optimizado(df, semanas_disponibles, FECHA_REFERENCIA, fecha_max)
                    pdf_resumen = generar_pdf_resumen_kpi(
                        df_kpis_semanales, 
                        num_semana, 
                        fecha_max_str, 
                        df_combinado, 
                        semanas_disponibles, 
                        FECHA_REFERENCIA, 
                        fecha_max
                    )
                
                for i, usuario_info in enumerate(usuarios_para_envio):
                    status_text.text(f"📨 Enviando a: {usuario_info['usuario']} ({usuario_info['email']})")
                    
                    # Generar PDF individual usando la función reutilizable
                    pdf_individual = generar_pdf_usuario(usuario_info['usuario'], df_pendientes, num_semana, fecha_max_str)
                    
                    if pdf_individual:
                        # Preparar archivos adjuntos
                        archivos_adjuntos = []
                        
                        # 1. PDF individual del usuario
                        nombre_individual = f"Expedientes_Pendientes_{usuario_info['usuario']}_Semana_{num_semana}.pdf"
                        archivos_adjuntos.append((nombre_individual, pdf_individual))
                        
                        # 2. PDF de resumen KPI (si corresponde)
                        if usuario_info['recibir_resumen'] and pdf_resumen:
                            nombre_resumen = f"Resumen_KPI_Semana_{num_semana}.pdf"
                            archivos_adjuntos.append((nombre_resumen, pdf_resumen))
                        
                        # Enviar correo con Outlook
                        exito = enviar_correo_outlook(
                            destinatario=usuario_info['email'],
                            asunto=usuario_info['asunto'],
                            cuerpo_mensaje=usuario_info['cuerpo_mensaje'],
                            archivos_adjuntos=archivos_adjuntos,
                            cc=usuario_info.get('cc'),
                            bcc=usuario_info.get('bcc')
                        )
                        
                        if exito:
                            correos_enviados += 1
                            st.success(f"✅ Enviado a {usuario_info['usuario']}" + 
                                      (" + Resumen KPI" if usuario_info['recibir_resumen'] and pdf_resumen else ""))
                        else:
                            correos_fallidos += 1
                            st.error(f"❌ Falló envío a {usuario_info['usuario']}")
                    else:
                        st.warning(f"⚠️ No se pudo generar PDF para {usuario_info['usuario']}")
                        correos_fallidos += 1
                    
                    progress_bar.progress((i + 1) / len(usuarios_para_envio))
                
                status_text.text("")
                
                # Mostrar resumen final
                st.markdown("---")
                st.subheader("📊 Resumen del Envío")
                
                col1, col2, col3 = st.columns(3)
                with col1:
                    st.metric("Total procesados", len(usuarios_para_envio))
                with col2:
                    st.metric("Correos enviados", correos_enviados, delta=f"+{correos_enviados}")
                with col3:
                    st.metric("Correos fallidos", correos_fallidos, delta=f"-{correos_fallidos}", delta_color="inverse")
                
                # Verificación segura con valor por defecto
                if 'envio_resumen_disponible' in locals() or 'envio_resumen_disponible' in globals():
                    if envio_resumen_disponible:
                        resumenes_enviados = sum(1 for u in usuarios_para_envio if u.get('recibir_resumen', False))
                        st.info(f"📊 Resúmenes KPI enviados: {resumenes_enviados} de {len(usuarios_para_envio)} usuarios")
                else:
                    # Si por alguna razón la variable no existe
                    resumenes_enviados = sum(1 for u in usuarios_para_envio if u.get('recibir_resumen', False))
                    if resumenes_enviados > 0:
                        st.info(f"📊 Resúmenes KPI enviados: {resumenes_enviados} de {len(usuarios_para_envio)} usuarios")
                
                if correos_enviados > 0:
                    st.balloons()
                    st.success("🎉 ¡Envío de correos completado!")
        
        else:
            st.warning("⚠️ No hay usuarios con expedientes pendientes para enviar")
    
    # Información de configuración
    st.markdown("---")
    with st.expander("⚙️ Configuración de Envío de Correos"):
        st.info("""
        **📋 Para que funcione el envío de correos:**
        
        1. **Outlook instalado** en el equipo
        2. **Cuenta de correo predeterminada** configurada en Outlook
        3. **Librería pywin32 instalada**: 
           ```bash
           pip install pywin32
           ```
        4. **Archivo USUARIOS.xlsx** con la estructura correcta
        
        **📋 Estructura del archivo USUARIOS.xlsx (Hoja Sheet1):**
        - USUARIOS: Código del usuario (debe coincidir con RECTAUTO)
        - ENVIAR: "SI" o "SÍ" (en mayúsculas)
        - ENVÍO RESUMEN: "SI" o "SÍ" para recibir resumen KPI (opcional)
        - EMAIL: Dirección de correo
        - ASUNTO: Puede usar &num_semana& y &fecha_max& como variables
        - MENSAJE1, MENSAJE2, MENSAJE3, DESPEDIDA: Texto del mensaje
        - CC, BCC: Opcionales (separar múltiples emails con ;)
        - RESUMEN: Opcional (nombre completo del usuario)
        - Otras columnas: Se pueden añadir sin afectar el funcionamiento
        """)

elif eleccion == "Indicadores clave (KPI)":
    # ... (el código existente de la sección KPI se mantiene igual)
    # Usar df_combinado en lugar de df
    df = df_combinado
    
    st.subheader("Indicadores clave (KPI)")
    
    # Obtener fecha de referencia para cálculos
    columna_fecha = df.columns[13]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    
    if pd.isna(fecha_max):
        st.error("No se pudo encontrar la fecha máxima en los datos")
        st.stop()
    
    # Crear rango de semanas disponibles
    fecha_inicio = pd.to_datetime("2022-11-01")
    semanas_disponibles = pd.date_range(
        start=fecha_inicio,
        end=fecha_max,
        freq='W-FRI'
    ).tolist()
    
    if not semanas_disponibles:
        st.error("No hay semanas disponibles para mostrar")
        st.stop()
    
    # Inicialización del estado para KPI
    if 'kpi_semana_index' not in st.session_state:
        st.session_state.kpi_semana_index = len(semanas_disponibles) - 1
    
    # Obtener la semana seleccionada actual
    semana_seleccionada = semanas_disponibles[st.session_state.kpi_semana_index]
    num_semana_seleccionada = ((semana_seleccionada - FECHA_REFERENCIA).days) // 7 + 1
    fecha_str = semana_seleccionada.strftime('%d/%m/%Y')
    
    # Selector de semana en el área principal
    st.markdown("---")
    st.header("🗓️ Selector de Semana")
    
    # Crear etiquetas formateadas para el slider
    opciones_slider = []
    for i, fecha in enumerate(semanas_disponibles):
        num_semana = ((fecha - FECHA_REFERENCIA).days) // 7 + 1
        fecha_str_opcion = fecha.strftime('%d/%m/%Y')
        opciones_slider.append(f"Semana {num_semana} ({fecha_str_opcion})")
    
    # Slider
    semana_index_slider = st.select_slider(
        "Selecciona la semana:",
        options=list(range(len(semanas_disponibles))),
        value=st.session_state.kpi_semana_index,
        format_func=lambda x: opciones_slider[x]
    )
    
    # Actualizar el índice si el slider cambió
    if semana_index_slider != st.session_state.kpi_semana_index:
        st.session_state.kpi_semana_index = semana_index_slider
        st.rerun()
    
    # RECALCULAR después de posibles cambios del slider
    semana_seleccionada = semanas_disponibles[st.session_state.kpi_semana_index]
    num_semana_seleccionada = ((semana_seleccionada - FECHA_REFERENCIA).days) // 7 + 1
    fecha_str = semana_seleccionada.strftime('%d/%m/%Y')
    
    # Mostrar información de la semana seleccionada
    st.info(f"**Semana seleccionada:** {fecha_str} (Semana {num_semana_seleccionada})")
    
    # Sidebar con botones de navegación
    with st.sidebar:
        st.header("🗓️ Navegación por Semanas")
        
        st.write(f"**Semana actual:**")
        st.write(f"{fecha_str}")
        st.write(f"(Semana {num_semana_seleccionada})")
        
        st.markdown("---")
        
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("◀️ Anterior", use_container_width=True, key="btn_anterior_kpi"):
                nuevo_indice = st.session_state.kpi_semana_index - 1
                if nuevo_indice >= 0:
                    st.session_state.kpi_semana_index = nuevo_indice
                    st.rerun()
        
        with col2:
            if st.button("Siguiente ▶️", use_container_width=True, key="btn_siguiente_kpi"):
                nuevo_indice = st.session_state.kpi_semana_index + 1
                if nuevo_indice < len(semanas_disponibles):
                    st.session_state.kpi_semana_index = nuevo_indice
                    st.rerun()
        
        st.write(f"**Posición:** {st.session_state.kpi_semana_index + 1} de {len(semanas_disponibles)}")
        
        if st.button("📅 Ir a semana actual", use_container_width=True, key="btn_actual_kpi"):
            st.session_state.kpi_semana_index = len(semanas_disponibles) - 1
            st.rerun()

    # Calcular KPIs para todas las semanas (usando cache)
    df_kpis_semanales = calcular_kpis_todas_semanas_optimizado(df, semanas_disponibles, FECHA_REFERENCIA, fecha_max)

    # Función para mostrar los nuevos KPIs principales
    def mostrar_kpis_principales(_df_kpis, _semana_seleccionada, _num_semana):
        kpis_semana = _df_kpis[_df_kpis['semana_numero'] == _num_semana].iloc[0]
        
        fecha_str = _semana_seleccionada.strftime('%d/%m/%Y')
        st.header(f"📊 KPIs de la Semana: {fecha_str} (Semana {_num_semana})")
        
        # Mostrar el rango correcto de la semana
        inicio_str = kpis_semana['inicio_semana'].strftime('%d/%m/%Y')
        fin_str = kpis_semana['fin_semana'].strftime('%d/%m/%Y')
        dias_semana = kpis_semana['dias_semana']
        es_actual = kpis_semana['es_semana_actual']
        
        if es_actual:
            st.info(f"**📅 Período de la semana (ACTUAL):** {inicio_str} (viernes) a {fin_str} (viernes) - {dias_semana} días")
            st.warning("ℹ️ **Semana actual:** Incluye todos los expedientes hasta la fecha de actualización")
        else:
            st.info(f"**📅 Período de la semana:** {inicio_str} (viernes) a {fin_str} (jueves) - {dias_semana} días")
        
        # PRIMERA FILA DE KPIs - DOS COLUMNAS (SEMANAL vs TOTALES)
        st.subheader("📈 KPIs Principales")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("### Semanal")
            st.metric(
                label="💰 Nuevos Expedientes",
                value=f"{int(kpis_semana['nuevos_expedientes']):,}".replace(",", ".")
            )
            st.metric(
                label="📤 Expedientes Despachados",
                value=f"{int(kpis_semana['despachados_semana']):,}".replace(",", ".")
            )
            st.metric(
                label="🎯 Coef. Absorción (Desp/Nuevos)",
                value=f"{kpis_semana['c_abs_despachados_sem']:.2f}%".replace(".", ",")
            )
            st.metric(
                label="🛒 Expedientes Cerrados",
                value=f"{int(kpis_semana['expedientes_cerrados']):,}".replace(",", ".")
            )
            st.metric(
                label="📊 Coef. Absorción (Cer/Asig)",
                value=f"{kpis_semana['c_abs_cerrados_sem']:.2f}%".replace(".", ",")
            )
            st.metric(
                label="👥 Expedientes Abiertos",
                value=f"{int(kpis_semana['total_abiertos']):,}".replace(",", ".")
            )

        with col2:
            st.markdown("### Totales (desde 01/11/2022)")
            st.metric(
                label="💰 Nuevos Expedientes",
                value=f"{int(kpis_semana['nuevos_expedientes_totales']):,}".replace(",", ".")
            )
            st.metric(
                label="📤 Expedientes Despachados",
                value=f"{int(kpis_semana['despachados_totales']):,}".replace(",", ".")
            )
            st.metric(
                label="🎯 Coef. Absorción (Desp/Nuevos)",
                value=f"{kpis_semana['c_abs_despachados_tot']:.2f}%".replace(".", ",")
            )
            st.metric(
                label="🛒 Expedientes Cerrados",
                value=f"{int(kpis_semana['expedientes_cerrados_totales']):,}".replace(",", ".")
            )
            # No hay totales para abiertos (es un snapshot)
            st.metric(
                label="📊 Coef. Absorción (Cer/Asig)",
                value=f"{kpis_semana['c_abs_cerrados_tot']:.2f}%".replace(".", ",")
            )
        
        st.markdown("---")
        
        # SEGUNDA FILA - EXPEDIENTES ESPECIALES
        st.subheader("📋 Expedientes con 029, 033, PRE, RSL, pendiente de firma, de decisión o de completar trámite")
        col1, col2 = st.columns(2)
        
        with col1:
            st.metric(
                label="🔍 Expedientes Especiales",
                value=f"{int(kpis_semana['expedientes_especiales']):,}".replace(",", ".")
            )
        
        with col2:
            st.metric(
                label="📈 Porcentaje sobre Abiertos",
                value=f"{kpis_semana['porcentaje_especiales']:.2f}%".replace(".", ",")
            )
        
        st.markdown("---")
        
        # TERCERA FILA - TIEMPOS DE TRAMITACIÓN (TRES COLUMNAS)
        st.subheader("⏱️ Tiempos de Tramitación (en días)")
        
        # Expedientes Despachados
        st.markdown("#### 📤 Expedientes Despachados")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                label="📊 Tiempo Medio",
                value=f"{kpis_semana['tiempo_medio_despachados']:.0f}"
            )
        
        with col2:
            st.metric(
                label="🎯 Percentil 90",
                value=f"{kpis_semana['percentil_90_despachados']:.0f}"
            )
        
        with col3:
            col3a, col3b = st.columns(2)
            with col3a:
                st.metric(
                    label="≤180 días",
                    value=f"{kpis_semana['percentil_180_despachados']:.2f}%".replace(".", ",")
                )
            with col3b:
                st.metric(
                    label="≤120 días",
                    value=f"{kpis_semana['percentil_120_despachados']:.2f}%".replace(".", ",")
                )
        
        # Expedientes Cerrados
        st.markdown("#### 📤 Expedientes Cerrados")
        col1, col2, col3 = st.columns(3)
        
        with col1:
            st.metric(
                label="📊 Tiempo Medio",
                value=f"{kpis_semana['tiempo_medio_cerrados']:.0f}"
            )
        
        with col2:
            st.metric(
                label="🎯 Percentil 90",
                value=f"{kpis_semana['percentil_90_cerrados']:.0f}"
            )
        
        with col3:
            col3a, col3b = st.columns(2)
            with col3a:
                st.metric(
                    label="≤180 días",
                    value=f"{kpis_semana['percentil_180_cerrados']:.2f}%".replace(".", ",")
                )
            with col3b:
                st.metric(
                    label="≤120 días",
                    value=f"{kpis_semana['percentil_120_cerrados']:.2f}%".replace(".", ",")
                )
        
        # Expedientes Abiertos
        st.markdown("#### 📤 Expedientes Abiertos")
        col1, col2, col3 = st.columns(3)
        
        #with col1:
        #    st.metric(
        #        label="📊 Tiempo Actual",
        #        value="-"  # No hay tiempo medio para abiertos
        #    )
        
        with col2:
            st.metric(
                label="🎯 Percentil 90",
                value=f"{kpis_semana['percentil_90_abiertos']:.0f}"
            )
        
        with col3:
            col3a, col3b = st.columns(2)
            with col3a:
                st.metric(
                    label="≤180 días",
                    value=f"{kpis_semana['percentil_180_abiertos']:.2f}%".replace(".", ",")
                )
            with col3b:
                st.metric(
                    label="≤120 días",
                    value=f"{kpis_semana['percentil_120_abiertos']:.2f}%".replace(".", ",")
                )

    # Mostrar dashboard principal
    mostrar_kpis_principales(df_kpis_semanales, semana_seleccionada, num_semana_seleccionada)

    # GRÁFICO DE EVOLUCIÓN TEMPORAL (ACTUALIZADO) - CORREGIDO
    st.markdown("---")
    st.subheader("📈 Evolución de KPIs Principales y Porcentajes")

    datos_grafico = df_kpis_semanales.copy()
    col1, col2, col3 = st.columns(3)
    num_semana_seleccionada = ((semana_seleccionada - FECHA_REFERENCIA).days) // 7 + 1

    # --- Gráfico 1: KPI principales (sin "Abiertos") ---
    with col1:
        fig1 = px.line(
            datos_grafico,
            x='semana_numero',
            y=['nuevos_expedientes', 'despachados_semana', 'expedientes_cerrados'],
            title='Evolución de Expedientes (Nuevos, Despachados, Cerrados)',
            labels={'semana_numero': 'Semana', 'value': 'Cantidad', 'variable': 'KPI'},
            color_discrete_map={
                'nuevos_expedientes': '#1f77b4',
                'despachados_semana': '#ff7f0e',
                'expedientes_cerrados': '#2ca02c'
            }
        )
        fig1.update_layout(
            height=400,
            hovermode="x unified",
            legend_title="KPI",
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5)
        )
        # Línea discontinua vertical indicando la semana seleccionada
        fig1.add_vline(
            x=num_semana_seleccionada,
            line_dash="dash",
            line_color="red",
            annotation_text=f"Semana {num_semana_seleccionada}",
            annotation_position="top left"
        )
        st.plotly_chart(fig1, use_container_width=True)

    # --- Gráfico 2: Solo "Expedientes Abiertos" ---
    with col2:
        fig2 = px.line(
            datos_grafico,
            x='semana_numero',
            y=['total_abiertos'],
            title='Evolución de Expedientes Abiertos',
            labels={'semana_numero': 'Semana', 'value': 'Cantidad', 'variable': 'KPI'},
            color_discrete_map={'total_abiertos': '#d62728'}
        )
        fig2.update_layout(
            height=400,
            hovermode="x unified",
            legend_title="KPI",
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5)
        )
        # Línea discontinua vertical indicando la semana seleccionada
        fig2.add_vline(
            x=num_semana_seleccionada,
            line_dash="dash",
            line_color="red",
            annotation_text=f"Semana {num_semana_seleccionada}",
            annotation_position="top left"
        )
        st.plotly_chart(fig2, use_container_width=True)

    # --- Gráfico 3: Porcentajes (4) ---
    with col3:
        fig3 = px.line(
            datos_grafico,
            x='semana_numero',
            y=[
                'c_abs_despachados_sem', 'c_abs_despachados_tot',
                'c_abs_cerrados_sem', 'c_abs_cerrados_tot'
            ],
            title='Evolución de Coeficientes de Absorción (%)',
            labels={'semana_numero': 'Semana', 'value': 'Porcentaje (%)', 'variable': 'Indicador'},
            color_discrete_map={
                'c_abs_despachados_sem': '#9467bd',
                'c_abs_despachados_tot': '#c5b0d5',
                'c_abs_cerrados_sem': '#8c564b',
                'c_abs_cerrados_tot': '#c49c94'
            }
        )
        fig3.update_layout(
            height=400,
            hovermode="x unified",
            legend_title="Indicador",
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5)
        )
        # Línea discontinua vertical indicando la semana seleccionada
        fig3.add_vline(
            x=num_semana_seleccionada,
            line_dash="dash",
            line_color="red",
            annotation_text=f"Semana {num_semana_seleccionada}",
            annotation_position="top left"
        )
        st.plotly_chart(fig3, use_container_width=True)

    # -------------------------------------------------------------
    # SEGUNDO BLOQUE: TIEMPOS DE TRAMITACIÓN
    # -------------------------------------------------------------
    st.markdown("---")
    st.subheader("⏱️ Tiempos de Tramitación")

    col1, col2 = st.columns(2)

    # --- Gráfico 1: Tiempos medios y percentiles 90 ---
    with col1:
        fig_tiempo = px.line(
            datos_grafico,
            x='semana_numero',
            y=[
                'tiempo_medio_despachados', 'tiempo_medio_cerrados',
                'percentil_90_despachados', 'percentil_90_cerrados'
            ],
            title='Tiempos Medios y Percentiles 90 (días)',
            labels={'semana_numero': 'Semana', 'value': 'Días', 'variable': 'Indicador'},
            color_discrete_map={
                'tiempo_medio_despachados': '#ff7f0e',
                'tiempo_medio_cerrados': '#2ca02c',
                'percentil_90_despachados': '#ffbb78',
                'percentil_90_cerrados': '#98df8a'
            }
        )
        fig_tiempo.update_layout(
            height=400,
            hovermode="x unified",
            legend_title="Indicador",
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5)
        )
        # Línea discontinua vertical indicando la semana seleccionada
        fig_tiempo.add_vline(
            x=num_semana_seleccionada,
            line_dash="dash",
            line_color="red",
            annotation_text=f"Semana {num_semana_seleccionada}",
            annotation_position="top left"
        )
        st.plotly_chart(fig_tiempo, use_container_width=True)

    # --- Gráfico 2: Porcentajes 120/180 ---
    with col2:
        fig_percentiles = px.line(
            datos_grafico,
            x='semana_numero',
            y=[
                'percentil_180_despachados', 'percentil_120_despachados',
                'percentil_180_cerrados', 'percentil_120_cerrados'
            ],
            title='Porcentaje de Expedientes ≤120 y ≤180 días (%)',
            labels={'semana_numero': 'Semana', 'value': 'Porcentaje (%)', 'variable': 'Indicador'},
            color_discrete_map={
                'percentil_180_despachados': '#ff7f0e',
                'percentil_120_despachados': '#ffddaa',
                'percentil_180_cerrados': '#2ca02c',
                'percentil_120_cerrados': '#98df8a'
            }
        )
        fig_percentiles.update_layout(
            height=400,
            hovermode="x unified",
            legend_title="Indicador",
            legend=dict(orientation='h', yanchor='bottom', y=1.02, xanchor='center', x=0.5)
        )
        # Línea discontinua vertical indicando la semana seleccionada
        fig_percentiles.add_vline(
            x=num_semana_seleccionada,
            line_dash="dash",
            line_color="red",
            annotation_text=f"Semana {num_semana_seleccionada}",
            annotation_position="top left"
        )
        st.plotly_chart(fig_percentiles, use_container_width=True)
