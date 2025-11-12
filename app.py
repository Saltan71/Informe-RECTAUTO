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
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, ColumnsAutoSizeMode

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
CACHE_TTL_STATIC = 86400  # 24 horas para datos estáticos
CACHE_TTL_DYNAMIC = 3600  # 1 hora para datos dinámicos
COL_WIDTHS_OPTIMIZED = [28, 11, 11, 8, 16, 11, 11, 16, 11, 20, 20, 9, 18, 11, 14, 9, 24, 20, 11]

# Test file en directorio único por usuario
test_file = user_env.get_temp_path("test_write_access.tmp")
with open(test_file, 'w') as f:
    f.write("test")

st.set_page_config(page_title="Informe Rectauto", layout="wide", page_icon=Image.open("icono.ico"))

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

# Menú principal reorganizado
menu = ["Carga de Archivos", "Vista de Expedientes", "Indicadores clave (KPI)", "Análisis del Rendimiento", "Informes y Correos"]
eleccion = st.sidebar.selectbox("Menú", menu)

# Función para obtener información de la semana actual
def obtener_info_semana_actual(df_combinado):
    """Obtiene la información de la semana actual para mostrar en todas las páginas"""
    if df_combinado is None or df_combinado.empty:
        return None, None, None
    
    try:
        columna_fecha = df_combinado.columns[13]
        df_combinado[columna_fecha] = pd.to_datetime(df_combinado[columna_fecha], errors='coerce')
        fecha_max = df_combinado[columna_fecha].max()
        
        if pd.isna(fecha_max):
            return None, None, None
            
        dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
        num_semana = dias_transcurridos // 7 + 1
        fecha_max_str = fecha_max.strftime("%d/%m/%Y")
        
        return num_semana, fecha_max_str, fecha_max
    except:
        return None, None, None

# Mostrar información de la semana actual en todas las páginas
if "df_combinado" in st.session_state:
    num_semana, fecha_max_str, fecha_max = obtener_info_semana_actual(st.session_state["df_combinado"])
    if num_semana and fecha_max_str:
        st.title(f"📊 Seguimiento Equipo Regional RECTAUTO - Semana {num_semana} a {fecha_max_str}")
    else:
        st.title("📊 Seguimiento Equipo Regional RECTAUTO")
else:
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

# === FUNCIONES OPTIMIZADAS ===

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

@st.cache_data(ttl=CACHE_TTL, show_spinner="Procesando archivos...")
def procesar_archivos_combinado(archivos_dict, _user_key=user_env.session_id):
    """Procesa todos los archivos en una sola función optimizada"""
    try:
        # Cargar RECTAUTO primero
        df_rectauto = cargar_y_procesar_rectauto(archivos_dict['rectauto'])
        
        # Cargar otros archivos en paralelo (si están disponibles)
        resultados = {}
        for nombre, archivo in archivos_dict.items():
            if archivo and nombre != 'rectauto':
                if nombre == 'notifica':
                    resultados['notifica'] = cargar_y_procesar_notifica(archivo)
                elif nombre == 'triaje':
                    resultados['triaje'] = cargar_y_procesar_triaje(archivo)
                elif nombre == 'usuarios':
                    resultados['usuarios'] = cargar_y_procesar_usuarios(archivo)
                elif nombre == 'documentos':
                    resultados['documentos'] = cargar_y_procesar_documentos(archivo)
        
        # Combinar todo
        df_combinado = combinar_archivos(
            df_rectauto, 
            resultados.get('notifica'),
            resultados.get('triaje'), 
            resultados.get('usuarios'),
            resultados.get('documentos')
        )
        
        return df_combinado, resultados.get('usuarios'), resultados.get('documentos')
        
    except Exception as e:
        st.error(f"Error en procesamiento combinado: {e}")
        return df_rectauto, None, None

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
        if 'FECHA' in col.upper():
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df

# === FUNCIONES DE CÁLCULO OPTIMIZADAS ===

def calcular_despachados_optimizado(_df, inicio_semana, fin_semana, fecha_inicio_totales):
    """Calcula despachados de forma optimizada usando operaciones vectorizadas"""
    if not all(col in _df.columns for col in ['FECHA RESOLUCIÓN', 'ESTADO', 'FECHA CIERRE']):
        return 0, 0

    fecha_9999 = pd.to_datetime('9999-09-09', errors='coerce')

    # Expedientes con FECHA RESOLUCIÓN real (distinta de 9999 y no nula) dentro del rango semanal
    mask_despachados_semana_reales = (
        _df['FECHA RESOLUCIÓN'].notna() & 
        (_df['FECHA RESOLUCIÓN'] != fecha_9999) &
        (_df['FECHA RESOLUCIÓN'] >= inicio_semana) &
        (_df['FECHA RESOLUCIÓN'] <= fin_semana)
    )

    # Expedientes CERRADOS con FECHA RESOLUCIÓN = 9999-09-09 o vacía
    mask_despachados_semana_cerrados = (
        (_df['ESTADO'] == 'Cerrado') &
        (_df['FECHA RESOLUCIÓN'].isna() | (_df['FECHA RESOLUCIÓN'] == fecha_9999)) &
        _df['FECHA CIERRE'].notna() &
        (_df['FECHA CIERRE'] >= inicio_semana) &
        (_df['FECHA CIERRE'] <= fin_semana)
    )

    mask_despachados_semana = mask_despachados_semana_reales | mask_despachados_semana_cerrados
    despachados_semana = _df[mask_despachados_semana].shape[0]

    # Totales: igual pero usando fecha_inicio_totales
    mask_despachados_totales_reales = (
        _df['FECHA RESOLUCIÓN'].notna() & 
        (_df['FECHA RESOLUCIÓN'] != fecha_9999) &
        (_df['FECHA RESOLUCIÓN'] >= fecha_inicio_totales) &
        (_df['FECHA RESOLUCIÓN'] <= fin_semana)
    )

    mask_despachados_totales_cerrados = (
        (_df['ESTADO'] == 'Cerrado') &
        (_df['FECHA RESOLUCIÓN'].isna() | (_df['FECHA RESOLUCIÓN'] == fecha_9999)) &
        _df['FECHA CIERRE'].notna() &
        (_df['FECHA CIERRE'] >= fecha_inicio_totales) &
        (_df['FECHA CIERRE'] <= fin_semana)
    )

    mask_despachados_totales = mask_despachados_totales_reales | mask_despachados_totales_cerrados
    despachados_totales = _df[mask_despachados_totales].shape[0]

    return despachados_semana, despachados_totales

# === FUNCIÓN AUXILIAR OPTIMIZADA PARA CÁLCULO DE TIEMPOS ===
def calcular_tiempos_optimizado(_df, fecha_inicio_totales, fin_semana):
    """Calcula los tiempos de tramitación de forma optimizada - VERSIÓN COMPLETAMENTE CORREGIDA"""
    
    # Inicializar resultados por defecto
    resultados = {
        'tiempo_medio_despachados': 0,
        'percentil_90_despachados': 0,
        'percentil_180_despachados': 0,
        'percentil_120_despachados': 0,
        'tiempo_medio_cerrados': 0,
        'percentil_90_cerrados': 0,
        'percentil_180_cerrados': 0,
        'percentil_120_cerrados': 0,
        'percentil_90_abiertos': 0,
        'percentil_180_abiertos': 0,
        'percentil_120_abiertos': 0
    }
    
    try:
        # Asegurarse de que las fechas son datetime
        fecha_inicio_totales = pd.to_datetime(fecha_inicio_totales)
        fin_semana = pd.to_datetime(fin_semana)
        
        # Crear copia para no modificar el original
        df_temp = _df.copy()
        
        # DEBUG: Verificar columnas disponibles
        print(f"Columnas disponibles: {list(df_temp.columns)}")
        
        # Convertir columnas de fecha necesarias
        columnas_fecha = ['FECHA RESOLUCIÓN', 'FECHA INICIO TRAMITACIÓN', 'FECHA CIERRE', 'FECHA APERTURA']
        for col in columnas_fecha:
            if col in df_temp.columns:
                df_temp[col] = pd.to_datetime(df_temp[col], errors='coerce')
                print(f"Columna {col} convertida, valores no nulos: {df_temp[col].notna().sum()}")
        
        fecha_9999 = pd.to_datetime('9999-09-09', errors='coerce')
        
        # ===== TIEMPOS PARA EXPEDIENTES DESPACHADOS =====
        if all(col in df_temp.columns for col in ['FECHA RESOLUCIÓN', 'FECHA INICIO TRAMITACIÓN', 'ESTADO', 'FECHA CIERRE']):
            print("Calculando tiempos para expedientes despachados...")
            
            # Crear máscaras vectorizadas para ambos tipos de despachados
            mask_despachados_reales = (
                df_temp['FECHA RESOLUCIÓN'].notna() & 
                (df_temp['FECHA RESOLUCIÓN'] != fecha_9999) &
                (df_temp['FECHA RESOLUCIÓN'] >= fecha_inicio_totales) &
                (df_temp['FECHA RESOLUCIÓN'] <= fin_semana)
            )
            
            mask_despachados_cerrados = (
                (df_temp['ESTADO'] == 'Cerrado') &
                (df_temp['FECHA RESOLUCIÓN'].isna() | (df_temp['FECHA RESOLUCIÓN'] == fecha_9999)) &
                df_temp['FECHA CIERRE'].notna() &
                (df_temp['FECHA CIERRE'] >= fecha_inicio_totales) &
                (df_temp['FECHA CIERRE'] <= fin_semana)
            )
            
            # Combinar máscaras
            mask_despachados_tiempo = mask_despachados_reales | mask_despachados_cerrados
            
            print(f"Expedientes despachados encontrados: {mask_despachados_tiempo.sum()}")
            
            if mask_despachados_tiempo.any():
                despachados_tiempo = df_temp[mask_despachados_tiempo].copy()
                
                # Calcular fecha final
                condiciones = [
                    despachados_tiempo['FECHA RESOLUCIÓN'].notna() & 
                    (despachados_tiempo['FECHA RESOLUCIÓN'] != fecha_9999)
                ]
                opciones = [despachados_tiempo['FECHA RESOLUCIÓN']]
                
                despachados_tiempo['FECHA_FINAL'] = np.select(
                    condiciones, 
                    opciones, 
                    default=despachados_tiempo['FECHA CIERRE']
                )
                
                # Calcular días de tramitación
                despachados_tiempo['dias_tramitacion'] = (
                    despachados_tiempo['FECHA_FINAL'] - despachados_tiempo['FECHA INICIO TRAMITACIÓN']
                ).dt.days
                
                # Filtrar valores válidos (mayores o iguales a 0)
                dias_validos = despachados_tiempo['dias_tramitacion'][despachados_tiempo['dias_tramitacion'] >= 0]
                
                print(f"Días válidos para despachados: {len(dias_validos)}")
                
                if not dias_validos.empty and len(dias_validos) > 0:
                    resultados['tiempo_medio_despachados'] = round(dias_validos.mean(), 1)
                    resultados['percentil_90_despachados'] = round(dias_validos.quantile(0.9), 1)
                    resultados['percentil_180_despachados'] = round((dias_validos <= 180).mean() * 100, 1)
                    resultados['percentil_120_despachados'] = round((dias_validos <= 120).mean() * 100, 1)
        
        # ===== TIEMPOS PARA EXPEDIENTES CERRADOS =====
        if all(col in df_temp.columns for col in ['FECHA CIERRE', 'FECHA INICIO TRAMITACIÓN']):
            print("Calculando tiempos para expedientes cerrados...")
            
            # Crear máscara vectorizada para expedientes cerrados
            mask_cerrados_tiempo = (
                df_temp['FECHA CIERRE'].notna() &
                (df_temp['FECHA CIERRE'] >= fecha_inicio_totales) & 
                (df_temp['FECHA CIERRE'] <= fin_semana)
            )
            
            print(f"Expedientes cerrados encontrados: {mask_cerrados_tiempo.sum()}")
            
            if mask_cerrados_tiempo.any():
                cerrados_tiempo = df_temp[mask_cerrados_tiempo].copy()
                
                # Calcular días de tramitación
                cerrados_tiempo['dias_tramitacion'] = (
                    cerrados_tiempo['FECHA CIERRE'] - cerrados_tiempo['FECHA INICIO TRAMITACIÓN']
                ).dt.days
                
                # Filtrar valores válidos
                dias_validos_cerrados = cerrados_tiempo['dias_tramitacion'][cerrados_tiempo['dias_tramitacion'] >= 0]
                
                print(f"Días válidos para cerrados: {len(dias_validos_cerrados)}")
                
                if not dias_validos_cerrados.empty and len(dias_validos_cerrados) > 0:
                    resultados['tiempo_medio_cerrados'] = round(dias_validos_cerrados.mean(), 1)
                    resultados['percentil_90_cerrados'] = round(dias_validos_cerrados.quantile(0.9), 1)
                    resultados['percentil_180_cerrados'] = round((dias_validos_cerrados <= 180).mean() * 100, 1)
                    resultados['percentil_120_cerrados'] = round((dias_validos_cerrados <= 120).mean() * 100, 1)
        
        # ===== TIEMPOS PARA EXPEDIENTES ABIERTOS =====
        if all(col in df_temp.columns for col in ['FECHA INICIO TRAMITACIÓN', 'FECHA APERTURA', 'FECHA CIERRE']):
            print("Calculando tiempos para expedientes abiertos...")
            
            # Crear máscara vectorizada para expedientes abiertos
            mask_abiertos_tiempo = (
                (df_temp['FECHA APERTURA'] <= fin_semana) & 
                ((df_temp['FECHA CIERRE'] > fin_semana) | (df_temp['FECHA CIERRE'].isna()))
            )
            
            print(f"Expedientes abiertos encontrados: {mask_abiertos_tiempo.sum()}")
            
            if mask_abiertos_tiempo.any():
                abiertos_tiempo = df_temp[mask_abiertos_tiempo].copy()
                
                # Calcular días de tramitación (hasta fin_semana)
                abiertos_tiempo['dias_tramitacion'] = (
                    fin_semana - abiertos_tiempo['FECHA INICIO TRAMITACIÓN']
                ).dt.days
                
                # Filtrar valores válidos
                dias_validos_abiertos = abiertos_tiempo['dias_tramitacion'][abiertos_tiempo['dias_tramitacion'] >= 0]
                
                print(f"Días válidos para abiertos: {len(dias_validos_abiertos)}")
                
                if not dias_validos_abiertos.empty and len(dias_validos_abiertos) > 0:
                    resultados['percentil_90_abiertos'] = round(dias_validos_abiertos.quantile(0.9), 1)
                    resultados['percentil_180_abiertos'] = round((dias_validos_abiertos <= 180).mean() * 100, 1)
                    resultados['percentil_120_abiertos'] = round((dias_validos_abiertos <= 120).mean() * 100, 1)
        
        print(f"Resultados finales: {resultados}")
        
    except Exception as e:
        print(f"ERROR en calcular_tiempos_optimizado: {e}")
        import traceback
        print(f"Detalle del error: {traceback.format_exc()}")
    
    return resultados

def calcular_kpis_para_semana(_df, semana_fin, es_semana_actual=False):
    """Versión optimizada del cálculo de KPIs con tiempos optimizados"""
    # Determinar rango de fechas
    if es_semana_actual:
        inicio_semana = semana_fin - timedelta(days=7)
        fin_semana = semana_fin
        dias_semana = 8
    else:
        inicio_semana = semana_fin - timedelta(days=7)
        fin_semana = semana_fin - timedelta(days=1)
        dias_semana = 7
    
    # 🔥 CORRECCIÓN: DEFINIR EXPLÍCITAMENTE fecha_inicio_totales
    fecha_inicio_totales = datetime(2022, 11, 1)  # Fecha de referencia explícita
    
    # Asegurar que las fechas son datetime
    inicio_semana = pd.to_datetime(inicio_semana)
    fin_semana = pd.to_datetime(fin_semana)
    
    # PRE-CALCULAR máscaras reutilizables
    if 'FECHA APERTURA' in _df.columns:
        mask_semana = (_df['FECHA APERTURA'] >= inicio_semana) & (_df['FECHA APERTURA'] <= fin_semana)
        mask_totales = (_df['FECHA APERTURA'] >= fecha_inicio_totales) & (_df['FECHA APERTURA'] <= fin_semana)
        
        nuevos_expedientes = _df.loc[mask_semana].shape[0]
        nuevos_expedientes_totales = _df.loc[mask_totales].shape[0]
    else:
        nuevos_expedientes = 0
        nuevos_expedientes_totales = 0

    # EXPEDIENTES DESPACHADOS - lógica optimizada
    despachados_semana, despachados_totales = calcular_despachados_optimizado(_df, inicio_semana, fin_semana, fecha_inicio_totales)

    # COEFICIENTE DE ABSORCIÓN 1 (Despachados/Nuevos)
    c_abs_despachados_sem = (despachados_semana / nuevos_expedientes * 100) if nuevos_expedientes > 0 else 0
    c_abs_despachados_tot = (despachados_totales / nuevos_expedientes_totales * 100) if nuevos_expedientes_totales > 0 else 0

    # EXPEDIENTES CERRADOS
    if 'FECHA CIERRE' in _df.columns:
        mask_cerrados_semana = (_df['FECHA CIERRE'] >= inicio_semana) & (_df['FECHA CIERRE'] <= fin_semana)
        mask_cerrados_totales = (_df['FECHA CIERRE'] >= fecha_inicio_totales) & (_df['FECHA CIERRE'] <= fin_semana)
        
        expedientes_cerrados = _df.loc[mask_cerrados_semana].shape[0]
        expedientes_cerrados_totales = _df.loc[mask_cerrados_totales].shape[0]
    else:
        expedientes_cerrados = 0
        expedientes_cerrados_totales = 0

    # EXPEDIENTES ABIERTOS
    if 'FECHA CIERRE' in _df.columns and 'FECHA APERTURA' in _df.columns:
        mask_abiertos = (_df['FECHA APERTURA'] <= fin_semana) & ((_df['FECHA CIERRE'] > fin_semana) | (_df['FECHA CIERRE'].isna()))
        total_abiertos = _df.loc[mask_abiertos].shape[0]
    else:
        total_abiertos = 0

    # COEFICIENTE DE ABSORCIÓN 2 (Cerrados/Asignados)
    c_abs_cerrados_sem = (expedientes_cerrados / nuevos_expedientes * 100) if nuevos_expedientes > 0 else 0
    c_abs_cerrados_tot = (expedientes_cerrados_totales / nuevos_expedientes_totales * 100) if nuevos_expedientes_totales > 0 else 0

    # EXPEDIENTES CON 029, 033, PRE o RSL
    if 'ETIQ. PENÚLTIMO TRAM.' in _df.columns:
        mask_especiales = (
            _df['ETIQ. PENÚLTIMO TRAM.'].notna() & 
            (~_df['ETIQ. PENÚLTIMO TRAM.'].isin(['1 APERTURA', '10 DATEXPTE'])) &
            (_df['FECHA APERTURA'] <= fin_semana) & 
            ((_df['FECHA CIERRE'] > fin_semana) | (_df['FECHA CIERRE'].isna()))
        )
        expedientes_especiales = _df.loc[mask_especiales].shape[0]
        porcentaje_especiales = (expedientes_especiales / total_abiertos * 100) if total_abiertos > 0 else 0
    else:
        expedientes_especiales = 0
        porcentaje_especiales = 0

    # ===== CÁLCULOS DE TIEMPOS OPTIMIZADOS =====
    tiempos = calcular_tiempos_optimizado(_df, fecha_inicio_totales, fin_semana)
    
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
        'tiempo_medio_despachados': tiempos['tiempo_medio_despachados'],
        'percentil_90_despachados': tiempos['percentil_90_despachados'],
        'percentil_180_despachados': tiempos['percentil_180_despachados'],
        'percentil_120_despachados': tiempos['percentil_120_despachados'],
        'tiempo_medio_cerrados': tiempos['tiempo_medio_cerrados'],
        'percentil_90_cerrados': tiempos['percentil_90_cerrados'],
        'percentil_180_cerrados': tiempos['percentil_180_cerrados'],
        'percentil_120_cerrados': tiempos['percentil_120_cerrados'],
        'percentil_90_abiertos': tiempos['percentil_90_abiertos'],
        'percentil_180_abiertos': tiempos['percentil_180_abiertos'],
        'percentil_120_abiertos': tiempos['percentil_120_abiertos'],
        'inicio_semana': inicio_semana,
        'fin_semana': fin_semana,
        'dias_semana': dias_semana,
        'es_semana_actual': es_semana_actual
    }

def identificar_filas_prioritarias(df):
    """Versión optimizada usando operaciones vectorizadas - MEJORADA"""
    try:
        # Crear copia para no modificar el original
        df_priorizado = df.copy(deep=True)
        
        # Inicializar todas las prioridades como 0
        df_priorizado['_prioridad'] = 0
        
        # Verificar que las columnas necesarias existen
        columnas_necesarias = ['ETIQ. PENÚLTIMO TRAM.', 'FECHA NOTIFICACIÓN', 'DOCUM.INCORP.']
        if not all(col in df_priorizado.columns for col in columnas_necesarias):
            return df_priorizado
        
        # Crear máscaras vectorizadas para cada condición
        etiq_penultimo = df_priorizado['ETIQ. PENÚLTIMO TRAM.'].astype(str).str.strip()
        fecha_notif = pd.to_datetime(df_priorizado['FECHA NOTIFICACIÓN'], errors='coerce')
        docum_incorp = df_priorizado['DOCUM.INCORP.'].astype(str).str.strip()
        
        # Inicializar máscara de prioritarios
        mask_prioritarios = pd.Series(False, index=df_priorizado.index)
        
        # CONDICIÓN 1: "80 PROPRES" con fecha límite superada
        mask_80_propres = (etiq_penultimo == "80 PROPRES")
        if mask_80_propres.any():
            fechas_limite_80 = fecha_notif[mask_80_propres] + timedelta(days=23)
            mask_80_vencido = mask_80_propres & (datetime.now() > fechas_limite_80)
            mask_prioritarios = mask_prioritarios | mask_80_vencido
        
        # CONDICIÓN 2: "50 REQUERIR" con fecha límite superada  
        mask_50_requerir = (etiq_penultimo == "50 REQUERIR")
        if mask_50_requerir.any():
            fechas_limite_50 = fecha_notif[mask_50_requerir] + timedelta(days=23)
            mask_50_vencido = mask_50_requerir & (datetime.now() > fechas_limite_50)
            mask_prioritarios = mask_prioritarios | mask_50_vencido
        
        # CONDICIÓN 3: "70 ALEGACI" o "60 CONTESTA"
        mask_alegaci_contesta = etiq_penultimo.isin(["70 ALEGACI", "60 CONTESTA"])
        mask_prioritarios = mask_prioritarios | mask_alegaci_contesta
        
        # CONDICIÓN 4: DOCUM.INCORP. válido
        mask_docum_valido = (
            docum_incorp.notna() & 
            (docum_incorp != '') &
            (docum_incorp != 'nan') &
            (~docum_incorp.str.upper().isin(["SOLICITUD", "REITERA SOLICITUD"]))
        )
        mask_prioritarios = mask_prioritarios | mask_docum_valido
        
        # Asignar prioridad 1 a los que cumplen alguna condición
        df_priorizado.loc[mask_prioritarios, '_prioridad'] = 1
        
        return df_priorizado
    
    except Exception as e:
        st.error(f"❌ Error al identificar filas prioritarias: {e}")
        # En caso de error, devolver DataFrame con prioridad 0
        df['_prioridad'] = 0
        return df

def actualizar_filtros_optimizados(df, estado_sel, equipo_sel, usuario_sel):
    """Versión optimizada del filtrado interconectado"""
    # Pre-calcular opciones disponibles
    opciones_base = {
        'estados': sorted(df['ESTADO'].dropna().unique()),
        'equipos': sorted(df['EQUIPO'].dropna().unique()),
        'usuarios': sorted(df['USUARIO'].dropna().unique())
    }
    
    # Aplicar filtros secuencialmente de forma vectorizada
    mask_actual = pd.Series(True, index=df.index)
    
    if estado_sel:
        mask_actual &= df['ESTADO'].isin(estado_sel)
    
    # Calcular equipos disponibles basado en estado
    equipos_disponibles = df.loc[mask_actual, 'EQUIPO'].dropna().unique()
    equipos_disponibles = sorted(equipos_disponibles)
    
    if equipo_sel:
        equipos_seleccionados = [eq for eq in equipo_sel if eq in equipos_disponibles]
        mask_actual &= df['EQUIPO'].isin(equipos_seleccionados) if equipos_seleccionados else mask_actual
    
    # Calcular usuarios disponibles basado en estado y equipo
    usuarios_disponibles = df.loc[mask_actual, 'USUARIO'].dropna().unique()
    usuarios_disponibles = sorted(usuarios_disponibles)
    
    if usuario_sel:
        usuarios_seleccionados = [us for us in usuario_sel if us in usuarios_disponibles]
        mask_actual &= df['USUARIO'].isin(usuarios_seleccionados) if usuarios_seleccionados else mask_actual
    
    return df[mask_actual].copy(), equipos_disponibles, usuarios_disponibles

@st.cache_data(ttl=3600)
def dataframe_to_pdf_bytes(df_mostrar, title, df_original):
    """Versión optimizada de generación de PDFs"""
    try:
        pdf = PDF('L', 'mm', 'A4')
        pdf.add_page()
        
        pdf.set_font("Arial", "B", 8)
        pdf.cell(0, 5, title, 0, 1, 'C')
        pdf.ln(5)

        # PRE-CALCULAR estructura de columnas
        columnas_a_excluir = {
            idx for idx, col_name in enumerate(df_mostrar.columns) 
            if "FECHA DE ACTUALIZACIÓN DATOS" in col_name.upper()
        }
        
        # Filtrar columnas una sola vez
        columnas_visibles = [
            (idx, col_name, col_width) 
            for idx, (col_name, col_width) in enumerate(zip(df_mostrar.columns, COL_WIDTHS_OPTIMIZED))
            if idx not in columnas_a_excluir
        ]

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
                pdf.cell(COL_WIDTHS_OPTIMIZED[i], ALTURA_ENCABEZADO, "", 1, 0, 'C', True)
                pdf.set_xy(x, y)

                texto = str(header)
                ancho_texto = pdf.get_string_width(texto)

                if ancho_texto <= COL_WIDTHS_OPTIMIZED[i] - 2:
                    altura_texto = 3
                    y_pos = y + (ALTURA_ENCABEZADO - altura_texto) / 2
                    pdf.set_xy(x, y_pos)
                    pdf.cell(COL_WIDTHS_OPTIMIZED[i], altura_texto, texto, 0, 0, 'C')
                else:
                    pdf.set_xy(x, y + 1)
                    pdf.multi_cell(COL_WIDTHS_OPTIMIZED[i], 2.5, texto, 0, 'C')

                pdf.set_xy(x + COL_WIDTHS_OPTIMIZED[i], y)

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
                ancho_disponible = min(COL_WIDTHS_OPTIMIZED) - 2
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
                ancho_total += COL_WIDTHS_OPTIMIZED[i]
                pdf.rect(x_inicio + sum(COL_WIDTHS_OPTIMIZED[:i]), y_inicio, COL_WIDTHS_OPTIMIZED[i], altura_fila)

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
                    
                x_celda = x_inicio + sum(COL_WIDTHS_OPTIMIZED[:col_idx_visible])
                y_celda = y_inicio

                pdf.aplicar_formato_condicional_pdf(
                    df_original, idx, col_name, COL_WIDTHS_OPTIMIZED[col_idx_visible], altura_fila, x_celda, y_celda
                )

                pdf.set_xy(x_celda, y_celda)
                pdf.multi_cell(COL_WIDTHS_OPTIMIZED[col_idx_visible], ALTURA_LINEA, texto, 0, 'L')
                
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

# === FUNCIONES ORIGINALES (MANTENIDAS POR COMPATIBILIDAD) ===

def ordenar_dataframe_por_prioridad_y_antiguedad(df):
    """Ordena el DataFrame: RUE amarillos primero, luego por antigüedad descendente - CORREGIDA DEFINITIVAMENTE"""
    try:
        # Crear una COPIA PROFUNDA para no modificar el original
        df_priorizado = df.copy(deep=True)
        
        # Identificar filas prioritarias en la copia
        df_priorizado = identificar_filas_prioritarias(df_priorizado)
        
        # Buscar la columna de antigüedad
        columnas_antiguedad = [col for col in df_priorizado.columns if 'ANTIGÜEDAD' in col.upper() or 'DÍAS' in col.upper()]
        
        if columnas_antiguedad:
            columna_antiguedad = columnas_antiguedad[0]
            # 🔥 CORRECCIÓN CRÍTICA: NO MODIFICAR LA COLUMNA DE ANTIGÜEDAD
            # Solo asegurarnos de que es numérica para ordenar, pero sin cambiar los valores
            if not pd.api.types.is_numeric_dtype(df_priorizado[columna_antiguedad]):
                # Si no es numérica, convertir temporalmente solo para ordenar
                df_priorizado['_antiguedad_temp'] = pd.to_numeric(
                    df_priorizado[columna_antiguedad], 
                    errors='coerce'
                ).fillna(0)
                columna_para_ordenar = '_antiguedad_temp'
            else:
                columna_para_ordenar = columna_antiguedad
        else:
            st.warning("⚠️ No se encontró columna de antigüedad, usando orden por prioridad solamente")
            columna_para_ordenar = None
        
        # Ordenar por prioridad (True primero) y luego por antigüedad descendente si existe
        if columna_para_ordenar:
            df_ordenado = df_priorizado.sort_values(
                ['_prioridad', columna_para_ordenar], 
                ascending=[False, False]
            )
            # Eliminar columna temporal si se creó
            if '_antiguedad_temp' in df_ordenado.columns:
                df_ordenado = df_ordenado.drop('_antiguedad_temp', axis=1)
        else:
            df_ordenado = df_priorizado.sort_values('_prioridad', ascending=False)
        
        # Eliminar columna temporal de prioridad
        if '_prioridad' in df_ordenado.columns:
            df_ordenado = df_ordenado.drop('_prioridad', axis=1)
        
        return df_ordenado
    
    except Exception as e:
        st.error(f"❌ Error al ordenar DataFrame: {e}")
        return df

def generar_pdf_usuario(usuario, df_pendientes, num_semana, fecha_max_str):
    """Genera el PDF para un usuario específico con nombre único - CORREGIDA PARA DECIMALES"""
    # Crear copia para no modificar el original
    df_user = df_pendientes[df_pendientes["USUARIO"] == usuario].copy(deep=True)
    
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

    # 🔥 CORRECCIÓN: Identificar columna de antigüedad
    columnas_antiguedad = [col for col in NOMBRES_COLUMNAS_PDF if 'ANTIGÜEDAD' in col.upper() or 'DÍAS' in col.upper()]
    
    # Crear DataFrame para mostrar (SOLO para visualización)
    df_pdf_mostrar = df_user_ordenado[NOMBRES_COLUMNAS_PDF].copy()
    
    # Formatear para visualización - CORREGIDO PARA DECIMALES
    for col in df_pdf_mostrar.columns:
        if df_pdf_mostrar[col].dtype == 'object':
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "nat", "none"] else str(x)
            )
        elif 'fecha' in col.lower():
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""
            )
        # 🔥 CORRECCIÓN DEFINITIVA: REDONDEAR EN LUGAR DE TRUNCAR
        elif df_pdf_mostrar[col].dtype in ['float64', 'float32']:
            # Para columnas flotantes (como antigüedad con decimales), REDONDEAR
            if col in columnas_antiguedad:
                df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                    lambda x: str(round(x)) if pd.notna(x) else "0"
                )
            else:
                df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                    lambda x: str(round(x)) if pd.notna(x) else "0"
                )
        elif df_pdf_mostrar[col].dtype in ['int64', 'int32']:
            # Para columnas enteras, mostrar normalmente
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: str(int(x)) if pd.notna(x) else "0"
            )

    num_expedientes = len(df_pdf_mostrar)
    
    titulo_pdf = f"{usuario} - Semana {num_semana} a {fecha_max_str} - Expedientes Pendientes ({num_expedientes})"
    
    return dataframe_to_pdf_bytes(df_pdf_mostrar, titulo_pdf, df_original=df_user_ordenado)

# === FUNCIONES AUXILIARES ===

def obtener_hash_archivo(archivo):
    """Genera un hash único del archivo para detectar cambios"""
    if archivo is None:
        return None
    archivo.seek(0)
    file_hash = hashlib.md5(archivo.read()).hexdigest()
    archivo.seek(0)
    return file_hash

@st.cache_data(ttl=CACHE_TTL)
def generar_pdf_equipo_prioritarios(equipo, df_pendientes, num_semana, fecha_max_str):
    """Genera el PDF para un equipo específico solo con expedientes prioritarios - CORREGIDA PARA DECIMALES"""
    # Crear copia para no modificar el original
    df_equipo = df_pendientes[df_pendientes["EQUIPO"] == equipo].copy(deep=True)
    
    if df_equipo.empty:
        return None
    
    # CORRECCIÓN: Usar la función de identificación de prioritarios
    df_prioritarios = identificar_filas_prioritarias(df_equipo)
    df_prioritarios = df_prioritarios[df_prioritarios['_prioridad'] == 1].copy()
    
    if df_prioritarios.empty:
        return None
    
    # Eliminar columna temporal de prioridad
    df_prioritarios = df_prioritarios.drop('_prioridad', axis=1)
    
    # 🔥 CORRECCIÓN: ORDENAR POR PRIORIDAD Y ANTIGÜEDAD
    df_prioritarios = ordenar_dataframe_por_prioridad_y_antiguedad(df_prioritarios)
    
    # Procesar datos para PDF
    indices_a_incluir = list(range(df_prioritarios.shape[1]))
    indices_a_excluir = {1, 4, 5, 6, 13}
    
    # EXCLUIR también la columna "FECHA DE ACTUALIZACIÓN DATOS" si existe
    for idx, col_name in enumerate(df_prioritarios.columns):
        if "FECHA DE ACTUALIZACIÓN DATOS" in col_name.upper():
            indices_a_excluir.add(idx)
    
    indices_finales = [i for i in indices_a_incluir if i not in indices_a_excluir]
    NOMBRES_COLUMNAS_PDF = df_prioritarios.columns[indices_finales].tolist()

    # 🔥 CORRECCIÓN: Identificar columna de antigüedad
    columnas_antiguedad = [col for col in NOMBRES_COLUMNAS_PDF if 'ANTIGÜEDAD' in col.upper() or 'DÍAS' in col.upper()]
    
    # Crear DataFrame para mostrar (SOLO para visualización)
    df_pdf_mostrar = df_prioritarios[NOMBRES_COLUMNAS_PDF].copy()
    
    # Formatear para visualización - CORREGIDO PARA DECIMALES
    for col in df_pdf_mostrar.columns:
        if df_pdf_mostrar[col].dtype == 'object':
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "nat", "none"] else str(x)
            )
        elif 'fecha' in col.lower():
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: x.strftime("%d/%m/%Y") if pd.notna(x) else ""
            )
        # 🔥 CORRECCIÓN DEFINITIVA: REDONDEAR EN LUGAR DE TRUNCAR
        elif df_pdf_mostrar[col].dtype in ['float64', 'float32']:
            # Para columnas flotantes (como antigüedad con decimales), REDONDEAR
            if col in columnas_antiguedad:
                df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                    lambda x: str(round(x)) if pd.notna(x) else "0"
                )
            else:
                df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                    lambda x: str(round(x)) if pd.notna(x) else "0"
                )
        elif df_pdf_mostrar[col].dtype in ['int64', 'int32']:
            # Para columnas enteras, mostrar normalmente
            df_pdf_mostrar[col] = df_pdf_mostrar[col].apply(
                lambda x: str(int(x)) if pd.notna(x) else "0"
            )

    num_expedientes = len(df_pdf_mostrar)
    
    titulo_pdf = f"{equipo} - Semana {num_semana} a {fecha_max_str} - Expedientes Prioritarios ({num_expedientes})"
    
    return dataframe_to_pdf_bytes(df_pdf_mostrar, titulo_pdf, df_original=df_prioritarios)

@st.cache_data(ttl=CACHE_TTL)
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
        pdf.ln(3)
        
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
        
        pdf.ln(3)
        
        # Totales
        pdf.set_font('Arial', 'B', 10)
        pdf.cell(0, 8, "Totales (desde 01/11/2022):", 0, 1)
        pdf.add_metric("Nuevos Expedientes", f"{int(kpis_semana['nuevos_expedientes_totales']):,}".replace(",", "."))
        pdf.add_metric("Expedientes Despachados", f"{int(kpis_semana['despachados_totales']):,}".replace(",", "."))
        pdf.add_metric("Coef. Absorcion (Desp/Nuevos)", f"{kpis_semana['c_abs_despachados_tot']:.2f}%".replace(".", ","))
        pdf.add_metric("Expedientes Cerrados", f"{int(kpis_semana['expedientes_cerrados_totales']):,}".replace(",", "."))
        pdf.add_metric("Coef. Absorcion (Cer/Asig)", f"{kpis_semana['c_abs_cerrados_tot']:.2f}%".replace(".", ","))
        
        pdf.ln(3)
        
        # SECCIÓN 2: EXPEDIENTES ESPECIALES
        pdf.add_section_title("EXPEDIENTES CON 029, 033, PRE, RSL, PENDIENTE DE FIRMA, DECISION O COMPLETAR TRAMITE")
        pdf.add_metric("Expedientes Especiales", f"{int(kpis_semana['expedientes_especiales']):,}".replace(",", "."))
        pdf.add_metric("Porcentaje sobre Abiertos", f"{kpis_semana['porcentaje_especiales']:.2f}%".replace(".", ","))
        
        pdf.ln(3)
        
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
        pdf.ln(3)
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
            pdf.ln(3)
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Evolucion de Expedientes (Nuevos, Despachados, Cerrados)", 0, 1)
            pdf.image(temp_chart1, x=10, w=190)
            
            pdf.ln(3)
            
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
            
            # pdf.add_page()
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Evolucion de Expedientes Abiertos", 0, 1)
            pdf.image(temp_chart2, x=10, w=190)
            
            pdf.ln(3)
            
            # GRÁFICO 3: Coeficientes de absorción
            fig3 = px.line(
                datos_grafico,
                x='semana_numero',
                y=['c_abs_despachados_sem', 'c_abs_despachados_tot', 'c_abs_cerrados_sem', 'c_abs_cerrados_tot'],
                title=f'Coeficientes de Absorcion - Semana {num_semana}',
                labels={'semana_numero': 'Semana', 'value': 'Porcentaje (%)', 'variable': 'Indicador'},
                color_discrete_map={
                    'c_abs_despachados_sem': '#9467bd',
                    'c_abs_cerrados_sem': '#8c564b',
                    'c_abs_despachados_tot': '#c5b0d5',
                    'c_abs_cerrados_tot': '#c49c94'
                }
            )
            fig3.update_layout(height=400, showlegend=True)
            fig3.add_vline(x=num_semana, line_dash="dash", line_color="red")
            
            temp_chart3 = user_env.get_temp_path(f"chart3_{num_semana}.png")
            fig3.write_image(temp_chart3)
            
            # df.add_page()
            pdf.set_font('Arial', 'B', 10)
            pdf.cell(0, 8, "Coeficientes de Absorcion Semanales (%)", 0, 1)
            pdf.image(temp_chart3, x=10, w=190)
            
            pdf.ln(3)
            
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
            
            pdf.ln(3)
            
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
            
            # pdf.add_page()
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
            pdf.ln(3)
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

def enviar_correo_outlook(destinatario, asunto, cuerpo_mensaje, archivos_adjuntos, cc=None, bcc=None):
    """
    Envía correo usando MAPI - versión específicamente corregida para el resumen
    """
    try:
        import win32com.client
        import pythoncom
        import os
        import tempfile
        
        # Forzar inicialización COM
        pythoncom.CoInitialize()
        
        try:
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)
            
            mail.To = destinatario
            mail.Subject = asunto
            mail.Body = cuerpo_mensaje

            if cc and pd.notna(cc) and str(cc).strip():
                mail.CC = str(cc)
            if bcc and pd.notna(bcc) and str(bcc).strip():
                mail.BCC = str(bcc)

            # Adjuntar archivos - SOLUCIÓN ESPECÍFICA PARA RESUMEN
            for nombre_archivo, datos_archivo in archivos_adjuntos:
                try:
                    # Crear archivo temporal con el nombre EXACTO que queremos
                    temp_path = os.path.join(user_env.working_dir, nombre_archivo)
                    
                    # Escribir el archivo con el nombre correcto
                    with open(temp_path, 'wb') as temp_file:
                        temp_file.write(datos_archivo)
                    
                    # Adjuntar el archivo - método simple y directo
                    mail.Attachments.Add(temp_path)
                    
                    # 🔥 SOLUCIÓN ESPECÍFICA: Forzar el cierre del archivo antes de enviar
                    del temp_file  # Liberar el handle del archivo
                    
                except Exception as attach_error:
                    st.error(f"Error adjuntando archivo {nombre_archivo}: {attach_error}")
                    # Continuar con otros archivos

            # Enviar correo
            mail.Send()
            
            # Limpiar archivos temporales después del envío
            for nombre_archivo, datos_archivo in archivos_adjuntos:
                try:
                    temp_path = os.path.join(user_env.working_dir, nombre_archivo)
                    if os.path.exists(temp_path):
                        os.unlink(temp_path)
                except:
                    pass  # Ignorar errores de limpieza
            
            return True
            
        except Exception as e:
            st.error(f"❌ Error Outlook al enviar a {destinatario}: {e}")
            return False
        finally:
            # Liberar objetos COM
            try:
                if 'mail' in locals():
                    del mail
                if 'outlook' in locals():
                    del outlook
            except:
                pass
            
    except Exception as e:
        st.error(f"❌ Error general al enviar a {destinatario}: {e}")
        return False
    finally:
        # Siempre desinicializar COM
        try:
            pythoncom.CoUninitialize()
        except:
            pass

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

# Función para gráficos dinámicos (SIN CACHE)
def crear_grafico_dinamico(_conteo, columna, titulo):
    """Crea gráficos dinámicos que responden a los filtros"""
    if _conteo.empty:
        return None
    
    fig = px.bar(_conteo, y=columna, x="Cantidad", title=titulo, text="Cantidad", 
                 color=columna, height=400)
    fig.update_traces(texttemplate='%{text:,}', textposition="auto")
    return fig

# =============================================
# PÁGINA 1: CARGA DE ARCHIVOS
# =============================================
if eleccion == "Carga de Archivos":
    st.header("📁 Carga de Archivos")
    
    # Mostrar botones de limpieza solo en esta página
    with st.sidebar:
        st.markdown("---")
        st.subheader("🛠️ Herramientas de Mantenimiento")
        col1, col2 = st.columns(2)
        
        with col1:
            if st.button("🔄 Limpiar cache", help="Limpiar toda la cache y recargar", use_container_width=True):
                st.cache_data.clear()
                # Mantener solo los datos esenciales
                keys_to_keep = ['df_combinado', 'df_usuarios', 'archivos_hash', 'filtro_estado', 'filtro_equipo', 'filtro_usuario']
                for key in list(st.session_state.keys()):
                    if key not in keys_to_keep:
                        del st.session_state[key]
                st.success("Cache limpiada correctamente")
                st.rerun()
        
        with col2:
            if st.button("🧹 Limpiar temp", help="Limpiar archivos temporales", use_container_width=True):
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

    # Procesar archivos cuando estén listos usando la función optimizada
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
            
            with st.spinner("🔄 Procesando archivos combinados..."):
                try:
                    # Usar la función optimizada de procesamiento combinado
                    archivos_dict = {
                        'rectauto': archivo_rectauto,
                        'notifica': archivo_notifica,
                        'triaje': archivo_triaje,
                        'usuarios': archivo_usuarios,
                        'documentos': archivo_documentos
                    }
                    
                    df_combinado, df_usuarios, datos_documentos = procesar_archivos_combinado(archivos_dict)
                    
                    # Convertir columnas de fecha
                    df_combinado = convertir_fechas(df_combinado)
                    
                    # Guardar en session_state
                    st.session_state["df_combinado"] = df_combinado
                    st.session_state["df_usuarios"] = df_usuarios
                    st.session_state["datos_documentos"] = datos_documentos
                    st.session_state["archivos_hash"] = archivos_actuales
                    
                    st.success(f"✅ Archivos procesados correctamente")
                    st.info(f"📊 Dataset final: {len(df_combinado)} registros, {len(df_combinado.columns)} columnas")
                    if df_usuarios is not None:
                        st.info(f"👥 Usuarios cargados: {len(df_usuarios)} registros")
                    if datos_documentos is not None:
                        st.info(f"📄 Documentos cargados: {len(datos_documentos['documentos'])} registros")
                    
                except Exception as e:
                    st.error(f"❌ Error procesando archivos: {e}")
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

# =============================================
# PÁGINA 2: VISTA DE EXPEDIENTES
# =============================================
elif eleccion == "Vista de Expedientes":
    st.header("📋 Vista de Expedientes")
    
    if "df_combinado" not in st.session_state:
        st.warning("⚠️ Primero carga los archivos en la sección 'Carga de Archivos'")
        st.stop()
    
    # Usar df_combinado en lugar de df
    df = st.session_state["df_combinado"]
    df_usuarios = st.session_state.get("df_usuarios", None)
    datos_documentos = st.session_state.get("datos_documentos", None)
    
    # =============================================
    # FILTROS DINÁMICOS INTERCONECTADOS - VERSIÓN MEJORADA CON ETIQUETAS
    # =============================================
    
    st.sidebar.header("Filtros Interconectados")

    # Inicializar variables de sesión si no existen
    if 'filtro_estado' not in st.session_state:
        st.session_state.filtro_estado = ['Abierto'] if 'Abierto' in df['ESTADO'].values else []

    if 'filtro_equipo' not in st.session_state:
        st.session_state.filtro_equipo = []

    if 'filtro_usuario' not in st.session_state:
        st.session_state.filtro_usuario = []

    if 'filtro_etiq_penultimo' not in st.session_state:
        st.session_state.filtro_etiq_penultimo = []

    if 'filtro_etiq_ultimo' not in st.session_state:
        st.session_state.filtro_etiq_ultimo = []

    # Botón para resetear filtros
    if st.sidebar.button("🔄 Mostrar todos / Resetear filtros", use_container_width=True):
        st.session_state.filtro_estado = sorted(df['ESTADO'].dropna().unique())
        st.session_state.filtro_equipo = sorted(df['EQUIPO'].dropna().unique())
        st.session_state.filtro_usuario = sorted(df['USUARIO'].dropna().unique())
        if 'ETIQ. PENÚLTIMO TRAM.' in df.columns:
            st.session_state.filtro_etiq_penultimo = sorted(df['ETIQ. PENÚLTIMO TRAM.'].dropna().unique())
        if 'ETIQ. ÚLTIMO TRAM.' in df.columns:
            st.session_state.filtro_etiq_ultimo = sorted(df['ETIQ. ÚLTIMO TRAM.'].dropna().unique())
        st.rerun()

    # 1. Primero aplicar filtros secuencialmente para calcular opciones disponibles
    df_filtrado_temp = df.copy()

    # Aplicar filtro de ESTADO primero
    if st.session_state.filtro_estado:
        df_filtrado_temp = df_filtrado_temp[df_filtrado_temp['ESTADO'].isin(st.session_state.filtro_estado)]

    # Calcular EQUIPOS disponibles basados en el filtro de ESTADO
    equipos_disponibles = sorted(df_filtrado_temp['EQUIPO'].dropna().unique())

    # Aplicar filtro de EQUIPO (si hay selección)
    equipos_seleccionados = [eq for eq in st.session_state.filtro_equipo if eq in equipos_disponibles]
    if equipos_seleccionados:
        df_filtrado_temp = df_filtrado_temp[df_filtrado_temp['EQUIPO'].isin(equipos_seleccionados)]

    # Calcular USUARIOS disponibles basados en filtros de ESTADO y EQUIPO
    usuarios_disponibles = sorted(df_filtrado_temp['USUARIO'].dropna().unique())

    # Aplicar filtro de USUARIO (si hay selección)
    usuarios_seleccionados = [us for us in st.session_state.filtro_usuario if us in usuarios_disponibles]
    if usuarios_seleccionados:
        df_filtrado_temp = df_filtrado_temp[df_filtrado_temp['USUARIO'].isin(usuarios_seleccionados)]

    # Calcular ETIQ. PENÚLTIMO TRAM. disponibles basados en filtros anteriores
    etiq_penultimo_disponibles = []
    if 'ETIQ. PENÚLTIMO TRAM.' in df_filtrado_temp.columns:
        etiq_penultimo_disponibles = sorted(df_filtrado_temp['ETIQ. PENÚLTIMO TRAM.'].dropna().unique())

    # Aplicar filtro de ETIQ. PENÚLTIMO TRAM. (si hay selección)
    etiq_penultimo_seleccionados = [etiq for etiq in st.session_state.filtro_etiq_penultimo if etiq in etiq_penultimo_disponibles]
    if etiq_penultimo_seleccionados:
        df_filtrado_temp = df_filtrado_temp[df_filtrado_temp['ETIQ. PENÚLTIMO TRAM.'].isin(etiq_penultimo_seleccionados)]

    # Calcular ETIQ. ÚLTIMO TRAM. disponibles basados en todos los filtros anteriores
    etiq_ultimo_disponibles = []
    if 'ETIQ. ÚLTIMO TRAM.' in df_filtrado_temp.columns:
        etiq_ultimo_disponibles = sorted(df_filtrado_temp['ETIQ. ÚLTIMO TRAM.'].dropna().unique())

    # Aplicar filtro de ETIQ. ÚLTIMO TRAM. (si hay selección)
    etiq_ultimo_seleccionados = [etiq for etiq in st.session_state.filtro_etiq_ultimo if etiq in etiq_ultimo_disponibles]
    if etiq_ultimo_seleccionados:
        df_filtrado_temp = df_filtrado_temp[df_filtrado_temp['ETIQ. ÚLTIMO TRAM.'].isin(etiq_ultimo_seleccionados)]

    # 2. Ahora crear los widgets de filtro con opciones actualizadas
    st.sidebar.markdown("---")
    st.sidebar.subheader("Filtros Activos")

    # FILTRO DE ESTADO (siempre muestra todas las opciones)
    opciones_estado = sorted(df['ESTADO'].dropna().unique())
    estado_sel = st.sidebar.multiselect(
        "🔘 Selecciona Estado:",
        options=opciones_estado,
        default=st.session_state.filtro_estado,
        key='filtro_estado_selector'
    )

    # FILTRO DE EQUIPO (se actualiza según estado seleccionado)
    equipo_sel = st.sidebar.multiselect(
        "👥 Selecciona Equipo:",
        options=equipos_disponibles,
        default=equipos_seleccionados,
        key='filtro_equipo_selector'
    )

    # FILTRO DE USUARIO (se actualiza según estado y equipo seleccionados)
    usuario_sel = st.sidebar.multiselect(
        "👤 Selecciona Usuario:",
        options=usuarios_disponibles,
        default=usuarios_seleccionados,
        key='filtro_usuario_selector'
    )

    # NUEVOS FILTROS: ETIQ. PENÚLTIMO TRAM. (se actualiza según filtros anteriores)
    if 'ETIQ. PENÚLTIMO TRAM.' in df.columns:
        etiq_penultimo_sel = st.sidebar.multiselect(
            "🏷️ ETIQ. PENÚLTIMO TRAM.:",
            options=etiq_penultimo_disponibles,
            default=etiq_penultimo_seleccionados,
            key='filtro_etiq_penultimo_selector'
        )
    else:
        etiq_penultimo_sel = []
        st.sidebar.info("ℹ️ Columna 'ETIQ. PENÚLTIMO TRAM.' no disponible")

    # NUEVOS FILTROS: ETIQ. ÚLTIMO TRAM. (se actualiza según todos los filtros anteriores)
    if 'ETIQ. ÚLTIMO TRAM.' in df.columns:
        etiq_ultimo_sel = st.sidebar.multiselect(
            "🏷️ ETIQ. ÚLTIMO TRAM.:",
            options=etiq_ultimo_disponibles,
            default=etiq_ultimo_seleccionados,
            key='filtro_etiq_ultimo_selector'
        )
    else:
        etiq_ultimo_sel = []
        st.sidebar.info("ℹ️ Columna 'ETIQ. ÚLTIMO TRAM.' no disponible")

    # 3. Actualizar session_state cuando cambian los filtros
    if estado_sel != st.session_state.filtro_estado:
        st.session_state.filtro_estado = estado_sel
        # Cuando cambia el estado, resetear todos los filtros dependientes
        st.session_state.filtro_equipo = []
        st.session_state.filtro_usuario = []
        st.session_state.filtro_etiq_penultimo = []
        st.session_state.filtro_etiq_ultimo = []
        st.rerun()

    if equipo_sel != st.session_state.filtro_equipo:
        st.session_state.filtro_equipo = equipo_sel
        # Cuando cambia el equipo, resetear usuario y etiquetas
        st.session_state.filtro_usuario = []
        st.session_state.filtro_etiq_penultimo = []
        st.session_state.filtro_etiq_ultimo = []
        st.rerun()

    if usuario_sel != st.session_state.filtro_usuario:
        st.session_state.filtro_usuario = usuario_sel
        # Cuando cambia el usuario, resetear etiquetas
        st.session_state.filtro_etiq_penultimo = []
        st.session_state.filtro_etiq_ultimo = []
        st.rerun()

    if 'ETIQ. PENÚLTIMO TRAM.' in df.columns and etiq_penultimo_sel != st.session_state.filtro_etiq_penultimo:
        st.session_state.filtro_etiq_penultimo = etiq_penultimo_sel
        # Cuando cambia etiq penúltimo, resetear etiq último
        st.session_state.filtro_etiq_ultimo = []
        st.rerun()

    if 'ETIQ. ÚLTIMO TRAM.' in df.columns and etiq_ultimo_sel != st.session_state.filtro_etiq_ultimo:
        st.session_state.filtro_etiq_ultimo = etiq_ultimo_sel
        st.rerun()

    # 4. Aplicar filtros finales al DataFrame principal
    df_filtrado = df.copy()

    if st.session_state.filtro_estado:
        df_filtrado = df_filtrado[df_filtrado['ESTADO'].isin(st.session_state.filtro_estado)]

    if st.session_state.filtro_equipo:
        df_filtrado = df_filtrado[df_filtrado['EQUIPO'].isin(st.session_state.filtro_equipo)]

    if st.session_state.filtro_usuario:
        df_filtrado = df_filtrado[df_filtrado['USUARIO'].isin(st.session_state.filtro_usuario)]

    if 'ETIQ. PENÚLTIMO TRAM.' in df.columns and st.session_state.filtro_etiq_penultimo:
        df_filtrado = df_filtrado[df_filtrado['ETIQ. PENÚLTIMO TRAM.'].isin(st.session_state.filtro_etiq_penultimo)]

    if 'ETIQ. ÚLTIMO TRAM.' in df.columns and st.session_state.filtro_etiq_ultimo:
        df_filtrado = df_filtrado[df_filtrado['ETIQ. ÚLTIMO TRAM.'].isin(st.session_state.filtro_etiq_ultimo)]

    # Mostrar resumen de filtros activos
    st.sidebar.markdown("---")
    st.sidebar.subheader("📊 Resumen Filtros")
    
    if st.session_state.filtro_estado:
        st.sidebar.write(f"**Estados:** {len(st.session_state.filtro_estado)} seleccionados")
    
    if st.session_state.filtro_equipo:
        st.sidebar.write(f"**Equipos:** {len(st.session_state.filtro_equipo)} seleccionados")
    
    if st.session_state.filtro_usuario:
        st.sidebar.write(f"**Usuarios:** {len(st.session_state.filtro_usuario)} seleccionados")
    
    if 'ETIQ. PENÚLTIMO TRAM.' in df.columns and st.session_state.filtro_etiq_penultimo:
        st.sidebar.write(f"**ETIQ. PENÚLTIMO:** {len(st.session_state.filtro_etiq_penultimo)} seleccionados")
    
    if 'ETIQ. ÚLTIMO TRAM.' in df.columns and st.session_state.filtro_etiq_ultimo:
        st.sidebar.write(f"**ETIQ. ÚLTIMO:** {len(st.session_state.filtro_etiq_ultimo)} seleccionados")
    
    st.sidebar.write(f"**Registros:** {len(df_filtrado):,}".replace(",", "."))

    # Mostrar detalles de filtros activos en un expander
    with st.sidebar.expander("📋 Ver detalles de filtros"):
        if st.session_state.filtro_estado:
            st.write("**Estados seleccionados:**")
            for estado in st.session_state.filtro_estado:
                st.write(f"- {estado}")
        
        if st.session_state.filtro_equipo:
            st.write("**Equipos seleccionados:**")
            for equipo in st.session_state.filtro_equipo:
                st.write(f"- {equipo}")
        
        if st.session_state.filtro_usuario:
            st.write("**Usuarios seleccionados:**")
            for usuario in st.session_state.filtro_usuario:
                st.write(f"- {usuario}")
        
        if 'ETIQ. PENÚLTIMO TRAM.' in df.columns and st.session_state.filtro_etiq_penultimo:
            st.write("**ETIQ. PENÚLTIMO TRAM. seleccionados:**")
            for etiq in st.session_state.filtro_etiq_penultimo:
                st.write(f"- {etiq}")
        
        if 'ETIQ. ÚLTIMO TRAM.' in df.columns and st.session_state.filtro_etiq_ultimo:
            st.write("**ETIQ. ÚLTIMO TRAM. seleccionados:**")
            for etiq in st.session_state.filtro_etiq_ultimo:
                st.write(f"- {etiq}")

    # NUEVO: Opciones de ordenamiento y auto-filtros
    st.sidebar.markdown("---")
    st.sidebar.subheader("Opciones de Visualización")
    
    # Checkbox para ordenar por prioridad
    ordenar_prioridad = st.sidebar.checkbox("Ordenar por prioridad (RUE amarillo primero)", value=True, key="ordenar_prioridad_checkbox")

    # Aplicar ordenamiento si está activado - CORREGIDO
    if ordenar_prioridad:
        df_filtrado = ordenar_dataframe_por_prioridad_y_antiguedad(df_filtrado)
        st.sidebar.success("✅ Ordenado por prioridad")
    else:
        # Si no está activado, ordenar solo por antigüedad descendente
        columnas_antiguedad = [col for col in df_filtrado.columns if 'ANTIGÜEDAD' in col.upper() or 'DÍAS' in col.upper()]
        if columnas_antiguedad:
            columna_antiguedad = columnas_antiguedad[0]
            df_filtrado = df_filtrado.sort_values(columna_antiguedad, ascending=False)
            st.sidebar.info("📊 Ordenado solo por antigüedad")

    # Auto-filtros para mostrar solo filas con formato condicional
    st.sidebar.markdown("**Auto-filtros:**")

    # Usar keys únicos para los checkboxes y definir las variables
    mostrar_solo_amarillos = st.sidebar.checkbox("Solo RUE prioritarios", value=False, key="filtro_amarillos")
    mostrar_solo_rojos = st.sidebar.checkbox("Solo USUARIO-CSV discrepantes", value=False, key="filtro_rojos") 
    mostrar_solo_90_incdocu = st.sidebar.checkbox("Solo con 90 INCDOCU", value=False, key="filtro_90_incdocu")

    # Aplicar auto-filtros - VERSIÓN CORREGIDA
    if mostrar_solo_amarillos or mostrar_solo_rojos or mostrar_solo_90_incdocu:
        # Crear una COPIA para los filtros sin afectar el DataFrame principal
        df_filtrado_temp = df_filtrado.copy(deep=True)
        
        filtros_aplicados = []
        
        if mostrar_solo_amarillos:
            # Filtrar solo RUE amarillos usando la función corregida
            df_priorizado_temp = identificar_filas_prioritarias(df_filtrado_temp)
            mask_amarillo = df_priorizado_temp['_prioridad'] == 1
            if mask_amarillo.any():
                df_filtrado_temp = df_filtrado_temp[mask_amarillo]
                filtros_aplicados.append(f"RUE prioritarios: {mask_amarillo.sum()}")
        
        if mostrar_solo_rojos:
            # Filtrar solo USUARIO-CSV rojos
            if 'USUARIO' in df_filtrado_temp.columns and 'USUARIO-CSV' in df_filtrado_temp.columns:
                # Comparación segura
                usuario_principal = df_filtrado_temp['USUARIO'].astype(str).str.strip().fillna('')
                usuario_csv = df_filtrado_temp['USUARIO-CSV'].astype(str).str.strip().fillna('')
                mask_rojo = usuario_principal != usuario_csv
                
                if mask_rojo.any():
                    df_filtrado_temp = df_filtrado_temp[mask_rojo]
                    filtros_aplicados.append(f"USUARIO-CSV discrepantes: {mask_rojo.sum()}")
        
        if mostrar_solo_90_incdocu:
            # Filtrar solo expedientes con "ETIQ. PENÚLTIMO TRAM." = "90 INCDOCU"
            if 'ETIQ. PENÚLTIMO TRAM.' in df_filtrado_temp.columns:
                mask_90_incdocu = df_filtrado_temp['ETIQ. PENÚLTIMO TRAM.'] == "90 INCDOCU"
                
                if mask_90_incdocu.any():
                    df_filtrado_temp = df_filtrado_temp[mask_90_incdocu]
                    filtros_aplicados.append(f"Con 90 INCDOCU: {mask_90_incdocu.sum()}")
                else:
                    st.sidebar.warning("No hay expedientes con ETIQ. PENÚLTIMO TRAM. = '90 INCDOCU'")
            else:
                st.sidebar.warning("Columna 'ETIQ. PENÚLTIMO TRAM.' no disponible")
        
        # Solo actualizar si se aplicaron filtros
        if filtros_aplicados:
            df_filtrado = df_filtrado_temp
            st.sidebar.success(" | ".join(filtros_aplicados))

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
    
    # Contar 90 INCDOCU (sustituye a DOCUM.INCORP.)
    if 'ETIQ. PENÚLTIMO TRAM.' in df_filtrado.columns:
        filas_90_incdocu = (df_filtrado['ETIQ. PENÚLTIMO TRAM.'] == "90 INCDOCU").sum()
        st.sidebar.write(f"Con 90 INCDOCU: {filas_90_incdocu}")

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

    # NUEVOS GRÁFICOS PARA LAS ETIQUETAS
    if 'ETIQ. PENÚLTIMO TRAM.' in df_filtrado.columns:
        conteo_penultimo = df_filtrado['ETIQ. PENÚLTIMO TRAM.'].value_counts().reset_index()
        conteo_penultimo.columns = ['ETIQ. PENÚLTIMO TRAM.', 'Cantidad']
        fig_penultimo = crear_grafico_dinamico(conteo_penultimo, 'ETIQ. PENÚLTIMO TRAM.', 'Distribución por ETIQ. PENÚLTIMO TRAM.')
        if fig_penultimo:
            st.plotly_chart(fig_penultimo, use_container_width=True)

    if 'ETIQ. ÚLTIMO TRAM.' in df_filtrado.columns:
        conteo_ultimo = df_filtrado['ETIQ. ÚLTIMO TRAM.'].value_counts().reset_index()
        conteo_ultimo.columns = ['ETIQ. ÚLTIMO TRAM.', 'Cantidad']
        fig_ultimo = crear_grafico_dinamico(conteo_ultimo, 'ETIQ. ÚLTIMO TRAM.', 'Distribución por ETIQ. ÚLTIMO TRAM.')
        if fig_ultimo:
            st.plotly_chart(fig_ultimo, use_container_width=True)

    if "NOTIFICADO" in df_filtrado.columns:
        conteo_notificado = df_filtrado["NOTIFICADO"].value_counts().reset_index()
        conteo_notificado.columns = ["NOTIFICADO", "Cantidad"]
        fig = crear_grafico_dinamico(conteo_notificado, "NOTIFICADO", "Expedientes notificados")
        if fig:
            st.plotly_chart(fig, use_container_width=True)

    # VISTA GENERAL - CON AGGRID
    st.subheader("📋 Vista general de expedientes")

    # Crear copia y formatear datos para AgGrid
    df_mostrar = df_filtrado.copy()

    # Formatear TODAS las columnas de fecha
    for col in df_mostrar.select_dtypes(include='datetime').columns:
        df_mostrar[col] = df_mostrar[col].dt.strftime("%d/%m/%Y")

    # 🔥 CORRECCIÓN: Redondear columnas numéricas con decimales
    columnas_antiguedad = [col for col in df_mostrar.columns if 'ANTIGÜEDAD' in col.upper() or 'DÍAS' in col.upper()]

    for col in df_mostrar.columns:
        if df_mostrar[col].dtype in ['float64', 'float32']:
            if col in columnas_antiguedad:
                # Redondear antigüedad y convertir a entero
                df_mostrar[col] = df_mostrar[col].apply(
                    lambda x: int(round(x)) if pd.notna(x) else 0
                )
            else:
                # Redondear otras columnas flotantes
                df_mostrar[col] = df_mostrar[col].apply(
                    lambda x: int(round(x)) if pd.notna(x) else 0
                )

    registros_mostrados = f"{len(df_mostrar):,}".replace(",", ".")
    registros_totales = f"{len(df):,}".replace(",", ".")
    st.write(f"Mostrando {registros_mostrados} de {registros_totales} registros")
    
    # CONFIGURACIÓN AGGRID CON FILTROS MEJORADOS
    gb = GridOptionsBuilder.from_dataframe(df_mostrar)

    # Configurar todas las columnas
    gb.configure_default_column(
        filterable=True,
        sortable=True,
        resizable=True,
        editable=False,
        groupable=False,
        min_column_width=100
    )

    # 🔥 CONFIGURACIÓN ESPECÍFICA PARA FECHAS CON MEJOR FORMATEO
    for col in columnas_fechas:
        if col in df_mostrar.columns:
            gb.configure_column(
                col,
                type=["dateColumn", "filterDateColumn"],
                filter="agDateColumnFilter",
                filterParams={
                    "buttons": ['apply', 'reset'],
                    "closeOnApply": True,
                    "browserDatePicker": True,
                },
                # 🔥 MEJOR FORMATEADOR PARA FECHAS
                valueFormatter="""
                function(params) {
                    if (!params.value) return '';
                    const date = new Date(params.value);
                    const day = date.getDate().toString().padStart(2, '0');
                    const month = (date.getMonth() + 1).toString().padStart(2, '0');
                    const year = date.getFullYear();
                    return `${day}/${month}/${year}`;
                }
                """
            )

    # 🔥 CONFIGURACIÓN SIMPLIFICADA DEL PANEL LATERAL (CORREGIDA)
    gb.configure_side_bar()  # ← Solo esto, sin parámetros

    # Configurar paginación
    gb.configure_pagination(
        paginationAutoPageSize=False,
        paginationPageSize=50
    )

    # Configurar selección
    gb.configure_selection(
        selection_mode="multiple",
        use_checkbox=True,
        groupSelectsChildren=True,
        groupSelectsFiltered=True
    )

    grid_options = gb.build()

    # 🔥 AGREGAR CONFIGURACIÓN DEL PANEL LATERAL DIRECTAMENTE EN grid_options
    grid_options.update({
        "sideBar": {
            "toolPanels": [
                {
                    "id": "filters",
                    "labelDefault": "Filtros",
                    "labelKey": "filters",
                    "iconKey": "filter",
                    "toolPanel": "agFiltersToolPanel",
                    "toolPanelParams": {
                        "expandFilters": True
                    }
                },
                {
                    "id": "columns",
                    "labelDefault": "Columnas",
                    "labelKey": "columns", 
                    "iconKey": "columns",
                    "toolPanel": "agColumnsToolPanel",
                    "toolPanelParams": {
                        "suppressRowGroups": True,
                        "suppressValues": True,
                        "suppressPivots": True,
                        "suppressPivotMode": True
                    }
                }
            ],
            "position": "right",
            "defaultToolPanel": "filters"
        }
    })

    # Mostrar tabla con AgGrid
    try:
        grid_response = AgGrid(
            df_mostrar,
            gridOptions=grid_options,
            height=600,
            width='100%',
            data_return_mode='AS_INPUT',
            update_mode='MODEL_CHANGED',
            fit_columns_on_grid_load=False,
            allow_unsafe_jscode=True,
            enable_enterprise_modules=True,
            theme='streamlit'
        )
        
        # OBTENER FILAS SELECCIONADAS
        selected_rows = grid_response.get('selected_rows', [])
        
        if selected_rows:
            st.info(f"📌 {len(selected_rows)} fila(s) seleccionada(s)")
            
    except Exception as e:
        st.error(f"❌ Error en AgGrid: {e}")
        selected_rows = []

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
        # Contar 90 INCDOCU (sustituye a DOCUM.INCORP.)
        if 'ETIQ. PENÚLTIMO TRAM.' in df_filtrado.columns:
            mask_90_incdocu = (df_filtrado['ETIQ. PENÚLTIMO TRAM.'] == "90 INCDOCU").sum()
            st.metric("Con 90 INCDOCU", f"{mask_90_incdocu:,}".replace(",", "."))
        else:
            st.metric("Con 90 INCDOCU", "N/A")

    # NUEVA SECCIÓN: GESTIÓN DE DOCUMENTACIÓN INCORPORADA
    st.markdown("---")
    st.header("📄 Gestión de Documentación Incorporada")

    if datos_documentos:
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
                        df_combinado = st.session_state["df_combinado"]
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
                        df_combinado = st.session_state["df_combinado"]
                        df_documentos_actualizado = df_combinado[['RUE', 'DOCUM.INCORP.']].copy()
                        df_documentos_actualizado = df_documentos_actualizado.dropna(subset=['DOCUM.INCORP.'])
                        df_documentos_actualizado = df_documentos_actualizado[df_documentos_actualizado['DOCUM.INCORP.'] != '']
                        
                        # Guardar en el archivo DOCUMENTOS.xlsx (esto reemplazará completamente el contenido anterior)
                        contenido_actualizado = guardar_documentos_actualizados(
                            datos_documentos['archivo'], 
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

# =============================================
# PÁGINA 3: INDICADORES CLAVE (KPI)
# =============================================
elif eleccion == "Indicadores clave (KPI)":
    st.header("📊 Indicadores clave (KPI)")
    
    if "df_combinado" not in st.session_state:
        st.warning("⚠️ Primero carga los archivos en la sección 'Carga de Archivos'")
        st.stop()
    
    # Usar df_combinado en lugar de df
    df = st.session_state["df_combinado"]
    
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

# =============================================
# PÁGINA 4: ANÁLISIS DEL RENDIMIENTO
# =============================================
elif eleccion == "Análisis del Rendimiento":
    st.header("📈 Análisis del Rendimiento")
    
    if "df_combinado" not in st.session_state:
        st.warning("⚠️ Primero carga los archivos en la sección 'Carga de Archivos'")
        st.stop()
    
    st.info("🔧 **Esta sección está en desarrollo.**")
    st.write("Próximamente se incluirán análisis avanzados de rendimiento y productividad.")
    
    # Aquí puedes agregar el contenido específico para el análisis de rendimiento
    st.write("Funcionalidades planificadas:")
    st.write("- Análisis comparativo entre equipos")
    st.write("- Tendencias de productividad")
    st.write("- Análisis de bottlenecks")
    st.write("- Recomendaciones de optimización")

# =============================================
# PÁGINA 5: INFORMES Y CORREOS
# =============================================
elif eleccion == "Informes y Correos":
    st.header("📧 Informes y Correos")
    
    if "df_combinado" not in st.session_state:
        st.warning("⚠️ Primero carga los archivos en la sección 'Carga de Archivos'")
        st.stop()
    
    # Usar df_combinado en lugar de df
    df = st.session_state["df_combinado"]
    df_usuarios = st.session_state.get("df_usuarios", None)
    
    # Obtener información de la semana actual
    columna_fecha = df.columns[13]
    df[columna_fecha] = pd.to_datetime(df[columna_fecha], errors='coerce')
    fecha_max = df[columna_fecha].max()
    dias_transcurridos = (fecha_max - FECHA_REFERENCIA).days
    num_semana = dias_transcurridos // 7 + 1
    fecha_max_str = fecha_max.strftime("%d/%m/%Y") if pd.notna(fecha_max) else "Sin fecha"
    
    # Descarga de informes
    st.subheader("📄 Generación de Informes PDF")
    
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
                    df, 
                    semanas_disponibles, 
                    FECHA_REFERENCIA, 
                    fecha_max
                )
                if pdf_resumen:
                    file_name = f"{num_semana}RESUMEN_KPI.pdf"
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

    # SECCIÓN: ENVÍO DE CORREOS INTEGRADA - VERSIÓN CORREGIDA
    st.markdown("---")
    st.subheader("📧 Envío de Correos Electrónicos")
    
    # CORRECCIÓN: Filtrar usuarios activos de forma más robusta
    try:
        if df_usuarios is None or df_usuarios.empty:
            st.error("❌ No se ha cargado el archivo USUARIOS o está vacío")
            st.stop()

        df_usuarios = df_usuarios.copy()

        # Normalizar campos de control
        for col in ['ENVIAR', 'ENVÍO RESUMEN']:
            if col not in df_usuarios.columns:
                df_usuarios[col] = ""
            df_usuarios[col] = df_usuarios[col].astype(str).str.upper().str.strip()

        valores_si = ['SÍ', 'SI', 'S', 'YES', 'Y', 'TRUE', '1', 'VERDADERO']

        # Usuarios que reciben su listado individual
        usuarios_enviar = df_usuarios[df_usuarios['ENVIAR'].isin(valores_si)]

        # Usuarios que reciben resumen KPI
        usuarios_resumen = df_usuarios[df_usuarios['ENVÍO RESUMEN'].isin(valores_si)]

        st.success(f"✅ {len(usuarios_enviar)} usuarios recibirán expedientes individuales")
        st.success(f"📊 {len(usuarios_resumen)} usuarios recibirán resumen KPI")

        if usuarios_enviar.empty and usuarios_resumen.empty:
            st.warning("⚠️ No hay usuarios con ENVIAR o ENVÍO RESUMEN = 'SÍ'")
            st.stop()

    except Exception as e:
        st.error(f"❌ Error procesando usuarios: {e}")
        st.stop()

    # Función para verificar envío de resumen - MÁS ROBUSTA
    def verificar_envio_resumen(usuario_row):
        """Verifica si el usuario debe recibir resumen KPI"""
        try:
            if 'ENVÍO RESUMEN' not in usuario_row.index:
                return False
            
            valor = str(usuario_row['ENVÍO RESUMEN']).upper().strip() if pd.notna(usuario_row['ENVÍO RESUMEN']) else ""
            valores_positivos = ['SÍ', 'SI', 'S', 'YES', 'Y', 'TRUE', '1', 'VERDADERO', 'OK', 'X', '✔', '✅']
            return valor in valores_positivos
        except:
            return False

    # Procesar usuarios para envío - SEPARAR CLARAMENTE LOS DOS GRUPOS
    usuarios_para_envio_individual = []  # Usuarios con expedientes que reciben su PDF
    usuarios_para_resumen_solo = []       # Usuarios que solo reciben resumen (jefes sin expedientes)

    # INICIALIZAR Y DEFINIR usuarios_con_pendientes DE FORMA SEGURA
    usuarios_con_pendientes = []
    try:
        # Primero, identificar todos los usuarios con expedientes pendientes
        if not df_pendientes.empty and 'USUARIO' in df_pendientes.columns:
            usuarios_con_pendientes = df_pendientes['USUARIO'].dropna().unique().tolist()
            st.info(f"📋 {len(usuarios_con_pendientes)} usuarios tienen expedientes pendientes")
        else:
            st.warning("ℹ️ No se encontraron expedientes pendientes o falta la columna 'USUARIO'")
            usuarios_con_pendientes = []
    except Exception as e:
        st.error(f"❌ Error al obtener usuarios con expedientes pendientes: {e}")
        usuarios_con_pendientes = []

    # VERIFICAR COLUMNAS REQUERIDAS EN EL ARCHIVO DE USUARIOS
    columnas_requeridas = ['USUARIOS', 'EMAIL']
    columnas_faltantes = [col for col in columnas_requeridas if col not in df_usuarios.columns]

    if columnas_faltantes:
        st.error(f"❌ Faltan columnas requeridas en el archivo USUARIOS: {', '.join(columnas_faltantes)}")
        st.stop()

    # PROCESAR CADA USUARIO ACTIVO - CORRECCIÓN IMPORTANTE
    for _, usuario_row in usuarios_enviar.iterrows():
        try:
            usuario_nombre = usuario_row['USUARIOS']  # Del archivo USUARIOS
            usuario_email = usuario_row['EMAIL']
            
            # Verificar datos básicos
            if pd.isna(usuario_nombre) or pd.isna(usuario_email):
                st.warning(f"⚠️ Usuario con nombre o email vacío: {usuario_nombre}")
                continue
                
            # Normalizar para comparación
            usuario_normalizado = str(usuario_nombre).strip().upper()
            
            # Verificar si tiene expedientes - COMPARAR CON 'USUARIO' de df_pendientes
            tiene_expedientes = False
            num_expedientes = 0
            
            # CORREGIDO: Verificar de forma segura si usuarios_con_pendientes está definido y no está vacío
            if usuarios_con_pendientes and len(usuarios_con_pendientes) > 0:
                try:
                    # CORREGIDO: Comparar con 'USUARIO' de los expedientes pendientes
                    tiene_expedientes = any(
                        str(user_pendiente).strip().upper() == usuario_normalizado 
                        for user_pendiente in usuarios_con_pendientes
                    )
                    
                    if tiene_expedientes:
                        # Contar expedientes del usuario - USANDO 'USUARIO' de df_pendientes
                        num_expedientes = len(df_pendientes[
                            df_pendientes['USUARIO'].apply(
                                lambda x: str(x).strip().upper() if pd.notna(x) else ''
                            ) == usuario_normalizado
                        ])
                except Exception as e:
                    st.error(f"❌ Error al verificar expedientes para {usuario_nombre}: {e}")
                    tiene_expedientes = False
                    num_expedientes = 0
            
            # Verificar si debe recibir resumen - CORRECCIÓN IMPORTANTE
            # Usamos la función verificar_envio_resumen para determinar esto
            recibir_resumen = verificar_envio_resumen(usuario_row)
            
            # Procesar asunto y mensaje
            asunto_template = usuario_row['ASUNTO'] if pd.notna(usuario_row.get('ASUNTO', '')) else f"Situación RECTAUTO asignados en la semana {num_semana} a {fecha_max_str}"
            asunto_procesado = asunto_template.replace("&num_semana&", str(num_semana)).replace("&fecha_max&", fecha_max_str)
            
            # Generar cuerpo del mensaje
            mensaje_base = ""
            if pd.notna(usuario_row.get('MENSAJE1', '')):
                mensaje_base += f"{usuario_row['MENSAJE1']}\n\n"
            if pd.notna(usuario_row.get('MENSAJE2', '')):
                mensaje_base += f"{usuario_row['MENSAJE2']}\n\n"
            if pd.notna(usuario_row.get('MENSAJE3', '')):
                mensaje_base += f"{usuario_row['MENSAJE3']}\n\n"
            if pd.notna(usuario_row.get('DESPEDIDA', '')):
                mensaje_base += f"{usuario_row['DESPEDIDA']}\n\n"
            
            if not mensaje_base.strip():
                mensaje_base = "Se adjunta informe de expedientes pendientes."
            
            mensaje_base += "__________________\n\nEquipo RECTAUTO."
            cuerpo_mensaje = f"Buenos días,\n\n{mensaje_base}"
            
            # CORRECCIÓN CRÍTICA: DETERMINAR A QUÉ LISTA PERTENECE
            # Primero verificar si el usuario está marcado para recibir resumen
            debe_recibir_resumen = verificar_envio_resumen(usuario_row)
            
            # Ahora determinar a qué lista va
            if tiene_expedientes:
                # Usuario con expedientes: va a la lista de envío individual
                usuarios_para_envio_individual.append({
                    'usuario': usuario_nombre,
                    'email': usuario_email,
                    'cc': usuario_row.get('CC', ''),
                    'bcc': usuario_row.get('BCC', ''),
                    'expedientes': num_expedientes,
                    'asunto': asunto_procesado,
                    'cuerpo_mensaje': cuerpo_mensaje,
                    'recibir_resumen': debe_recibir_resumen  # Puede que también reciba resumen
                })
                st.success(f"✅ {usuario_nombre} - Tiene {num_expedientes} expedientes" + 
                        (" y recibirá resumen" if debe_recibir_resumen else ""))
            
            # CORRECCIÓN: Usuarios que NO tienen expedientes pero SÍ deben recibir resumen
            elif debe_recibir_resumen:
                # Usuario sin expedientes pero que debe recibir resumen
                usuarios_para_resumen_solo.append({
                    'usuario': usuario_nombre,
                    'email': usuario_email,
                    'cc': usuario_row.get('CC', ''),
                    'bcc': usuario_row.get('BCC', ''),
                    'asunto': f"Resumen KPI Semana {num_semana} - {fecha_max_str}",
                    'cuerpo_mensaje': f"Buenos días,\n\nSe adjunta el resumen de KPIs de la semana {num_semana} y los listados de expedientes prioritarios de todos los equipos.\n\n__________________\n\nEquipo RECTAUTO.",
                    'recibir_resumen': True
                })
                st.info(f"📊 {usuario_nombre} - Recibirá resumen KPI + expedientes prioritarios de todos los equipos")
            
            else:
                # Usuario sin expedientes y sin marcar para resumen
                st.warning(f"⚠️ {usuario_nombre} - No tiene expedientes ni está marcado para resumen")
                
        except Exception as e:
            st.error(f"❌ Error procesando usuario {usuario_row.get('USUARIOS', 'Desconocido')}: {e}")
            continue

    # Procesar también usuarios que solo reciben resumen (no están en ENVIAR pero sí en ENVÍO RESUMEN)
    for _, usuario_row in usuarios_resumen.iterrows():
        try:
            usuario_nombre = usuario_row['USUARIOS']
            usuario_email = usuario_row['EMAIL']
            
            # Verificar que no esté ya en la lista de envío individual
            if any(u['usuario'] == usuario_nombre for u in usuarios_para_envio_individual):
                continue  # Ya está procesado
                
            if any(u['usuario'] == usuario_nombre for u in usuarios_para_resumen_solo):
                continue  # Ya está procesado
                
            # Verificar datos básicos
            if pd.isna(usuario_nombre) or pd.isna(usuario_email):
                st.warning(f"⚠️ Usuario con nombre o email vacío: {usuario_nombre}")
                continue
            
            # Solo agregar si no tiene expedientes (para evitar duplicados)
            usuario_normalizado = str(usuario_nombre).strip().upper()
            tiene_expedientes = any(
                str(user_pendiente).strip().upper() == usuario_normalizado 
                for user_pendiente in usuarios_con_pendientes
            ) if usuarios_con_pendientes else False
            
            if not tiene_expedientes:
                usuarios_para_resumen_solo.append({
                    'usuario': usuario_nombre,
                    'email': usuario_email,
                    'cc': usuario_row.get('CC', ''),
                    'bcc': usuario_row.get('BCC', ''),
                    'asunto': f"Resumen KPI Semana {num_semana} - {fecha_max_str}",
                    'cuerpo_mensaje': f"Buenos días,\n\nSe adjunta el resumen de KPIs de la semana {num_semana} y los listados de expedientes prioritarios de todos los equipos.\n\n__________________\n\nEquipo RECTAUTO.",
                    'recibir_resumen': True
                })
                st.info(f"📊 {usuario_nombre} - Recibirá solo resumen KPI (sin expedientes propios)")
                
        except Exception as e:
            st.error(f"❌ Error procesando usuario resumen {usuario_row.get('USUARIOS', 'Desconocido')}: {e}")
            continue

    # MOSTRAR RESUMEN DE LO QUE SE VA A ENVIAR
    st.markdown("---")
    st.subheader("📋 Resumen de Envíos Programados")

    if usuarios_para_envio_individual:
        st.success(f"✅ {len(usuarios_para_envio_individual)} usuarios recibirán sus expedientes pendientes")
        with st.expander("📋 Ver usuarios con expedientes"):
            df_envio = pd.DataFrame(usuarios_para_envio_individual)
            st.dataframe(df_envio[['usuario', 'email', 'expedientes', 'recibir_resumen']], use_container_width=True)

    if usuarios_para_resumen_solo:
        st.success(f"📊 {len(usuarios_para_resumen_solo)} usuarios recibirán solo el resumen KPI")
        with st.expander("📋 Ver usuarios para resumen"):
            df_resumen = pd.DataFrame(usuarios_para_resumen_solo)
            st.dataframe(df_resumen[['usuario', 'email']], use_container_width=True)

    if not usuarios_para_envio_individual and not usuarios_para_resumen_solo:
        st.warning("⚠️ No hay usuarios para enviar correos")
        st.stop()

    # BOTÓN DE ENVÍO - SIEMPRE VISIBLE SI HAY USUARIOS EN ALGUNA LISTA
    st.markdown("---")
    st.subheader("🚀 Envío de Correos")

    st.warning("""
    **⚠️ Importante antes de enviar:**
    - Se usará la cuenta de Outlook predeterminada
    - No es necesario tener Outlook abierto
    - Los correos se enviarán inmediatamente
    - **Usuarios con expedientes pendientes:** Recibirán su PDF individual
    - **Usuarios Gerente y Jefes de Equipo:** Recibirán el resumen KPI y los Expedientes Prioritarios de todos los equipos
    - **Verifica que los datos sean correctos**
    """)

    if st.button("📤 Enviar Todos los Correos", type="primary", key="enviar_correos"):
        progress_bar = st.progress(0)
        status_text = st.empty()
        results_container = st.container()
        
        correos_enviados = 0
        correos_fallidos = 0
        
        # Generar PDF de resumen KPI (una sola vez para todos)
        pdf_resumen = None
        with st.spinner("Generando resumen KPI..."):
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
                df, 
                semanas_disponibles, 
                FECHA_REFERENCIA, 
                fecha_max
            )
        
        total_a_procesar = len(usuarios_para_envio_individual) + len(usuarios_para_resumen_solo)
        
        # PRIMERO: Enviar a usuarios con expedientes
        for i, usuario_info in enumerate(usuarios_para_envio_individual):
            status_text.text(f"📨 Enviando a: {usuario_info['usuario']}")
            
            # Generar PDF individual
            pdf_individual = generar_pdf_usuario(usuario_info['usuario'], df_pendientes, num_semana, fecha_max_str)
            
            if pdf_individual:
                archivos_adjuntos = []
                
                # 1. PDF individual
                nombre_individual = f"Expedientes_Pendientes_{usuario_info['usuario']}_Semana_{num_semana}.pdf"
                archivos_adjuntos.append((nombre_individual, pdf_individual))
                
                # 2. Resumen KPI si corresponde
                if usuario_info['recibir_resumen'] and pdf_resumen:
                    nombre_resumen = f"Resumen_KPI_Semana_{num_semana}.pdf"
                    archivos_adjuntos.append((nombre_resumen, pdf_resumen))
                
                # Enviar correo
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
                    with results_container:
                        st.success(f"✅ {usuario_info['usuario']} - Expedientes" + 
                                (" + Resumen" if usuario_info['recibir_resumen'] and pdf_resumen else ""))
                else:
                    correos_fallidos += 1
                    with results_container:
                        st.error(f"❌ Falló: {usuario_info['usuario']}")
            else:
                correos_fallidos += 1
                with results_container:
                    st.error(f"❌ No se pudo generar PDF para {usuario_info['usuario']}")
            
            progress_bar.progress((i + 1) / total_a_procesar)
        
        # SEGUNDO: Enviar solo resúmenes a usuarios sin expedientes - CON EXPEDIENTES PRIORITARIOS DE TODOS LOS EQUIPOS
        for i, usuario_info in enumerate(usuarios_para_resumen_solo):
            status_text.text(f"📊 Enviando resumen a: {usuario_info['usuario']}")
            
            # Crear lista de adjuntos FRESCA para cada usuario
            archivos_adjuntos = []
            
            if pdf_resumen:
                # 1. Adjuntar resumen KPI
                nombre_resumen = f"Resumen_KPI_Semana_{num_semana}.pdf"
                archivos_adjuntos.append((nombre_resumen, pdf_resumen))
                
                # 🔥 NUEVO: Adjuntar expedientes prioritarios de TODOS los equipos
                # Obtener la lista de equipos únicos con expedientes pendientes
                equipos = df_pendientes['EQUIPO'].dropna().unique()
                
                for equipo in equipos:
                    with st.spinner(f"Generando expedientes prioritarios para {equipo}..."):
                        pdf_prioritarios_equipo = generar_pdf_equipo_prioritarios(
                            equipo, 
                            df_pendientes, 
                            num_semana, 
                            fecha_max_str
                        )
                        
                        if pdf_prioritarios_equipo:
                            nombre_prioritarios = f"Expedientes_Prioritarios_{equipo}_Semana_{num_semana}.pdf"
                            archivos_adjuntos.append((nombre_prioritarios, pdf_prioritarios_equipo))
                            st.success(f"✅ Expedientes prioritarios de {equipo} generados")
                        else:
                            st.warning(f"⚠️ No hay expedientes prioritarios para el equipo {equipo}")
                
                # Actualizar el cuerpo del mensaje para reflejar los nuevos adjuntos
                cuerpo_mensaje_actualizado = f"Buenos días,\n\nSe adjunta el resumen de KPIs de la semana {num_semana} y los listados de expedientes prioritarios de todos los equipos.\n\n__________________\n\nEquipo RECTAUTO."
                
                # Enviar correo con todos los adjuntos
                exito = enviar_correo_outlook(
                    destinatario=usuario_info['email'],
                    asunto=usuario_info['asunto'],
                    cuerpo_mensaje=cuerpo_mensaje_actualizado,
                    archivos_adjuntos=archivos_adjuntos,
                    cc=usuario_info.get('cc'),
                    bcc=usuario_info.get('bcc')
                )
                
                if exito:
                    correos_enviados += 1
                    with results_container:
                        st.success(f"📊 Resumen KPI y {len(equipos)} equipos de expedientes prioritarios enviados a {usuario_info['usuario']}")
                else:
                    correos_fallidos += 1
                    with results_container:
                        st.error(f"❌ Falló resumen: {usuario_info['usuario']}")
            else:
                correos_fallidos += 1
                with results_container:
                    st.error(f"❌ No hay resumen para {usuario_info['usuario']}")
            
            progress_bar.progress((len(usuarios_para_envio_individual) + i + 1) / total_a_procesar)
        
        status_text.text("")
        
        # RESUMEN FINAL
        st.markdown("---")
        st.subheader("📊 Resumen del Envío")
        
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total procesados", total_a_procesar)
        with col2:
            st.metric("Correos enviados", correos_enviados)
        with col3:
            st.metric("Correos fallidos", correos_fallidos)
        
        # Calcular resúmenes enviados
        resumenes_enviados = sum(1 for u in usuarios_para_envio_individual if u.get('recibir_resumen', False))
        resumenes_enviados += len(usuarios_para_resumen_solo)
        
        st.info(f"📊 Resúmenes KPI enviados: {resumenes_enviados}")
        
        if correos_enviados > 0:
            st.balloons()
            st.success("🎉 ¡Envío de correos completado!")