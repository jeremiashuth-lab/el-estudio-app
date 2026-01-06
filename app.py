import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import os
import altair as alt
from io import BytesIO
from docx import Document
from docx.shared import Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
import calendar
import extra_streamlit_components as stx
import time
import re

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="El Estudio", page_icon="🔥", layout="wide")

# --- LISTA DE MESES ---
MESES_ESP = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# --- ESTILOS CSS (OPTIMIZADO PARA IOS) ---
def cargar_estilos():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        html, body, [class*="css"] { font-family: 'Montserrat', sans-serif; }
        
        header {visibility: hidden;}
        .stApp > header {background-color: transparent;}
        footer {visibility: hidden;}
        
        /* Títulos */
        h1 { color: #E63946 !important; font-weight: 800 !important; letter-spacing: -1px; }
        h2, h3 { font-weight: 600 !important; letter-spacing: -0.5px; }
        
        /* Calendario */
        .calendar-container {
            display: grid; grid-template-columns: repeat(7, 1fr); gap: 5px; margin-top: 10px; margin-bottom: 20px;
        }
        .calendar-day-header { text-align: center; font-weight: bold; color: #888; font-size: 0.8rem; }
        .calendar-day {
            aspect-ratio: 1; display: flex; align-items: center; justify-content: center;
            border-radius: 50%; font-weight: bold; font-size: 0.9rem; color: #FFF; background-color: #262730;
        }
        .day-completed { background-color: #2ECC71 !important; color: #000 !important; }
        .day-incomplete { background-color: #F39C12 !important; color: #000 !important; }
        .day-empty { background-color: transparent; }
        
        /* Métricas y Contenedores */
        div[data-testid="stExpander"] { border: none; box-shadow: none; }
        div[data-testid="stMetric"] { background-color: #1A1C24; border: 1px solid #333; padding: 10px; border-radius: 8px; }
        div[data-testid="stMetricValue"] { color: #E63946 !important; font-weight: 800; font-size: 1.8rem !important; }
        
        /* Botones (Toque rápido en iOS) */
        div.stButton > button:first-child {
            background-color: #E63946; color: white; border-radius: 8px; border: none;
            font-weight: 600; text-transform: uppercase; letter-spacing: 1px;
            padding-top: 12px; padding-bottom: 12px;
            touch-action: manipulation;
            -webkit-tap-highlight-color: transparent;
        }
        div.stButton > button:first-child:hover { background-color: #FF4D5A; }
        
        /* Ajuste de Tablas para Móvil */
        div[data-testid="stDataFrame"] { width: 100%; }
        </style>
    """, unsafe_allow_html=True)

cargar_estilos()

# --- CONEXIÓN GOOGLE SHEETS ---
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def conectar_google_sheet():
    if os.path.exists("mis_secretos.json"):
        creds = Credentials.from_service_account_file("mis_secretos.json", scopes=SCOPES)
    else:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds).open("El Estudio DB")

# --- FUNCIONES DE LECTURA ---
@st.cache_data(ttl=600)
def obtener_todos_usuarios():
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Usuarios").get_all_records())
    df.columns = [c.strip().capitalize() for c in df.columns]
    if "Rol" in df.columns: df["Rol"] = df["Rol"].astype(str).str.strip().str.lower()
    if "Usuario" in df.columns: df["Usuario"] = df["Usuario"].astype(str).str.strip()
    return df

@st.cache_data(ttl=600)
def leer_rutina(alumno):
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Rutinas").get_all_records())
    if df.empty: return df
    df.columns = [c.strip().capitalize() for c in df.columns]
    df["Alumno_Norm"] = df["Alumno"].astype(str).str.strip().str.lower()
    alumno_norm = alumno.strip().lower()
    df["Seccion"] = df["Seccion"].astype(str).str.strip().str.capitalize()
    for col in ["Link", "Series", "Reps", "Kg"]:
        if col not in df.columns: df[col] = ""
        else: df[col] = df[col].astype(str)
    return df[df["Alumno_Norm"] == alumno_norm].drop(columns=["Alumno_Norm"])

@st.cache_data(ttl=600)
def leer_sesiones_alumno(alumno):
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Sesiones").get_all_records())
    if df.empty: return df
    df.columns = [c.strip().capitalize() for c in df.columns]
    if "Usuario" not in df.columns: return pd.DataFrame()
    df["Usuario_Norm"] = df["Usuario"].astype(str).str.strip().str.lower()
    alumno_norm = alumno.strip().lower()
    df_alumno = df[df["Usuario_Norm"] == alumno_norm].copy()
    if not df_alumno.empty and "Fecha" in df_alumno.columns:
        df_alumno["Fecha"] = pd.to_datetime(df_alumno["Fecha"], format='mixed').dt.normalize()
    return df_alumno

@st.cache_data(ttl=600)
def leer_registros_full():
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Registros").get_all_records())
    if not df.empty: df.columns = [c.strip().capitalize() for c in df.columns]
    return df

@st.cache_data(ttl=600)
def leer_registros_alumno(alumno):
    df = leer_registros_full()
    if df.empty: return df
    col_map = {c: c for c in df.columns}
    for c in df.columns:
        if c.lower() in ['user', 'alumno']: col_map[c] = 'Usuario'
        if c.lower() in ['date', 'day']: col_map[c] = 'Fecha'
    df = df.rename(columns=col_map)
    if "Usuario" not in df.columns: return pd.DataFrame()
    
    df["Usuario_Norm"] = df["Usuario"].astype(str).str.strip().str.lower()
    alumno_norm = alumno.strip().lower()
    df_alumno = df[df["Usuario_Norm"] == alumno_norm].copy()
    
    if not df_alumno.empty and "Fecha" in df_alumno.columns:
        df_alumno["Fecha"] = pd.to_datetime(df_alumno["Fecha"], format='mixed', errors='coerce').dt.normalize()
        df_alumno = df_alumno.dropna(subset=["Fecha"])
        def extraer_numero(texto):
            try:
                match = re.search(r"(\d+[.,]?\d*)", str(texto))
                return float(match.group(1).replace(",", ".")) if match else 0.0
            except: return 0.0
        if "Peso" in df_alumno.columns:
            df_alumno["Peso_Grafico"] = df_alumno["Peso"].apply(extraer_numero)
        else: df_alumno["Peso_Grafico"] = 0.0
    return df_alumno

# --- FUNCIONES DE ESCRITURA ---
def obtener_usuario(usuario_input, password_input):
    try:
        df = obtener_todos_usuarios()
        u_input = usuario_input.strip().lower()
        if "Usuario" not in df.columns: return None
        df["User_Lower"] = df["Usuario"].astype(str).str.strip().str.lower()
        usuario = df[(df["User_Lower"] == u_input) & (df["Password"].astype(str).str.strip() == password_input.strip())]
        return usuario.iloc[0] if not usuario.empty else None
    except: return None

def obtener_usuario_por_cookie(usuario_input):
    try:
        df = obtener_todos_usuarios()
        if "Usuario" not in df.columns: return None
        u_input = usuario_input.strip().lower()
        df["User_Lower"] = df["Usuario"].astype(str).str.strip().str.lower()
        usuario = df[df["User_Lower"] == u_input]
        return usuario.iloc[0] if not usuario.empty else None
    except: return None

def guardar_rutina_actualizada(alumno, dia, df_calentamiento, df_fuerza, df_cardio):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Rutinas")
    all_data = ws.get_all_records()
    cols = ["Alumno", "Dia", "Seccion", "Orden", "Ejercicio", "Link", "Series", "Reps", "Kg", "Notas"]
    nuevas_filas = []
    def get_val(row, key): return str(row.get(key, "")).strip()
    df_calentamiento = df_calentamiento.fillna("")
    df_fuerza = df_fuerza.fillna("")
    df_cardio = df_cardio.fillna("")

    for _, row in df_calentamiento.iterrows():
        if row["Ejercicio"] or row["Notas"]: 
             nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Calentamiento", "Orden": "-", "Ejercicio": row["Ejercicio"], "Link": get_val(row, "Link"), "Series": get_val(row, "Series"), "Reps": get_val(row, "Reps"), "Kg": "-", "Notas": row["Notas"]})
    for _, row in df_fuerza.iterrows():
        if row["Ejercicio"] or row["Notas"] or row["Orden"]:
            nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Fuerza", "Orden": row.get("Orden", "-"), "Ejercicio": row["Ejercicio"], "Link": get_val(row, "Link"), "Series": get_val(row, "Series"), "Reps": get_val(row, "Reps"), "Kg": get_val(row, "Kg"), "Notas": row["Notas"]})
    for _, row in df_cardio.iterrows():
        if row["Ejercicio"] or row["Notas"]:
            nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Cardio", "Orden": "-", "Ejercicio": row["Ejercicio"], "Link": get_val(row, "Link"), "Series": get_val(row, "Series"), "Reps": get_val(row, "Reps"), "Kg": "-", "Notas": row["Notas"]})

    if not all_data:
        df_final = pd.DataFrame(nuevas_filas)
        for c in cols: 
            if c not in df_final.columns: df_final[c] = ""
        df_final = df_final[cols]
        ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())
        return

    df_old = pd.DataFrame(all_data)
    df_old.columns = [c.strip().capitalize() for c in df_old.columns]
    if "Alumno" in df_old.columns: df_old["Alumno_Norm"] = df_old["Alumno"].astype(str).str.strip().str.lower()
    else: df_old["Alumno_Norm"] = ""
    if "Dia" in df_old.columns: df_old["Dia_Norm"] = df_old["Dia"].astype(str).str.strip().str.lower()
    else: df_old["Dia_Norm"] = ""
    a_norm = alumno.strip().lower()
    d_norm = dia.strip().lower()
    mask = ~((df_old["Alumno_Norm"] == a_norm) & (df_old["Dia_Norm"] == d_norm))
    df_clean = df_old[mask].drop(columns=["Alumno_Norm", "Dia_Norm"], errors='ignore')
    df_nuevas = pd.DataFrame(nuevas_filas)
    df_final = pd.concat([df_clean, df_nuevas], ignore_index=True)
    df_final = df_final.fillna("")
    for c in cols: 
        if c not in df_final.columns: df_final[c] = ""
    ws.clear(); ws.update([df_final[cols].columns.values.tolist()] + df_final[cols].values.tolist())

def guardar_registro(usuario, ejercicio, peso, reps, rpe, notas, fecha_input=None):
    sh = conectar_google_sheet()
    if fecha_input:
        fecha = fecha_input.strftime("%Y-%m-%d")
    else:
        fecha = datetime.now().strftime("%Y-%m-%d") 
    sh.worksheet("Registros").append_row([fecha, usuario, ejercicio, peso, reps, rpe, notas])

def actualizar_registros_usuario(usuario, df_editado):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Registros")
    all_data = ws.get_all_records()
    df_all = pd.DataFrame(all_data)
    if df_all.empty: return 
    
    df_all.columns = [c.strip().capitalize() for c in df_all.columns]
    col_map = {c: c for c in df_all.columns}
    for c in df_all.columns:
        if c.lower() in ['user', 'alumno']: col_map[c] = 'Usuario'
    df_all = df_all.rename(columns=col_map)

    df_all["Usuario_Norm"] = df_all["Usuario"].astype(str).str.strip().str.lower()
    usuario_norm = usuario.strip().lower()
    df_otros = df_all[df_all["Usuario_Norm"] != usuario_norm].drop(columns=["Usuario_Norm"])
    
    df_usuario_nuevo = df_editado.copy()
    if "Fecha" in df_usuario_nuevo.columns:
        df_usuario_nuevo["Fecha"] = pd.to_datetime(df_usuario_nuevo["Fecha"]).dt.strftime("%Y-%m-%d")
    
    df_final = pd.concat([df_otros, df_usuario_nuevo], ignore_index=True)
    df_final = df_final.fillna("")
    
    if "Fecha" in df_final.columns:
        df_final = df_final.sort_values("
