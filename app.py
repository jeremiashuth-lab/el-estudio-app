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
import calendar
import extra_streamlit_components as stx

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="El Estudio", page_icon="🔥", layout="wide")

# --- LISTA DE MESES ---
MESES_ESP = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

# --- ESTILOS CSS ---
def cargar_estilos():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        html, body, [class*="css"] { font-family: 'Montserrat', sans-serif; }
        
        header {visibility: visible;}
        .stApp > header {background-color: transparent;}
        footer {visibility: hidden;}
        
        h1 { color: #E63946 !important; font-weight: 800 !important; letter-spacing: -1px; }
        h2, h3 { font-weight: 600 !important; letter-spacing: -0.5px; }
        
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
        
        div[data-testid="stExpander"] { border: none; box-shadow: none; }
        div[data-testid="stMetric"] { background-color: #1A1C24; border: 1px solid #333; padding: 10px; border-radius: 8px; }
        div[data-testid="stMetricValue"] { color: #E63946 !important; font-weight: 800; font-size: 1.8rem !important; }
        
        div.stButton > button:first-child {
            background-color: #E63946; color: white; border-radius: 8px; border: none;
            font-weight: 600; text-transform: uppercase; letter-spacing: 1px;
            padding-top: 12px; padding-bottom: 12px;
        }
        div.stButton > button:first-child:hover { background-color: #FF4D5A; }
        </style>
    """, unsafe_allow_html=True)

cargar_estilos()

# --- CONEXIÓN GOOGLE SHEETS (CON CACHÉ RESOURCE) ---
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def conectar_google_sheet():
    if os.path.exists("mis_secretos.json"):
        creds = Credentials.from_service_account_file("mis_secretos.json", scopes=SCOPES)
    else:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds).open("El Estudio DB")

# --- FUNCIONES DE LECTURA (CON CACHÉ DATA) ---
@st.cache_data(ttl=600)
def obtener_todos_usuarios():
    sh = conectar_google_sheet()
    return pd.DataFrame(sh.worksheet("Usuarios").get_all_records())

@st.cache_data(ttl=600)
def leer_rutina(alumno):
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Rutinas").get_all_records())
    if df.empty: return df
    df["Alumno"] = df["Alumno"].astype(str).str.strip()
    df["Seccion"] = df["Seccion"].astype(str).str.strip().str.capitalize()
    return df[df["Alumno"] == alumno.strip()]

@st.cache_data(ttl=600)
def leer_sesiones_alumno(alumno):
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Sesiones").get_all_records())
    if df.empty: return df
    df_alumno = df[df["Usuario"].astype(str).str.strip() == alumno.strip()].copy()
    if not df_alumno.empty:
        df_alumno["Fecha"] = pd.to_datetime(df_alumno["Fecha"], format='mixed').dt.normalize()
    return df_alumno

@st.cache_data(ttl=600)
def leer_registros_alumno(alumno):
    sh = conectar_google_sheet()
    raw_data = sh.worksheet("Registros").get_all_records()
    df = pd.DataFrame(raw_data)
    if df.empty: return df
    df_alumno = df[df["Usuario"].astype(str).str.strip() == alumno.strip()].copy()
    if not df_alumno.empty:
        df_alumno["Fecha"] = pd.to_datetime(df_alumno["Fecha"], format='mixed', errors='coerce').dt.normalize()
        df_alumno["Peso"] = pd.to_numeric(df_alumno["Peso"], errors='coerce').fillna(0)
        if "Repeticiones" in df_alumno.columns:
            df_alumno["Repeticiones"] = pd.to_numeric(df_alumno["Repeticiones"], errors='coerce').fillna(0)
        else: df_alumno["Repeticiones"] = 0
    return df_alumno

# --- FUNCIONES DE ESCRITURA (SIN CACHÉ) ---
def obtener_usuario(usuario_input, password_input):
    try:
        df = obtener_todos_usuarios()
        usuario = df[
            (df["Usuario"].astype(str).str.strip() == usuario_input.strip()) & 
            (df["Password"].astype(str).str.strip() == password_input.strip())
        ]
        return usuario.iloc[0] if not usuario.empty else None
    except: return None

def obtener_usuario_por_cookie(usuario_input):
    try:
        df = obtener_todos_usuarios()
        usuario = df[df["Usuario"].astype(str).str.strip() == usuario_input.strip()]
        return usuario.iloc[0] if not usuario.empty else None
    except: return None

def guardar_rutina_actualizada(alumno, dia, df_calentamiento, df_fuerza, df_cardio):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Rutinas")
    all_data = ws.get_all_records()
    cols = ["Alumno", "Dia", "Seccion", "Orden", "Ejercicio", "Series", "Reps", "Kg", "Notas"]

    nuevas_filas = []
    for _, row in df_calentamiento.iterrows():
        if row["Ejercicio"]: 
             s = row.get("Series", "2"); r = row.get("Reps", "10")
             nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Calentamiento", "Orden": "-", "Ejercicio": row["Ejercicio"], "Series": s, "Reps": r, "Kg": "-", "Notas": row["Notas"]})
    for _, row in df_fuerza.iterrows():
        if row["Ejercicio"]:
            s = row.get("Series", "3"); r = row.get("Reps", "10"); o = row.get("Orden", "-"); ej = row["Ejercicio"]
            nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Fuerza", "Orden": o, "Ejercicio": ej, "Series": s, "Reps": r, "Kg": row["Kg"], "Notas": row["Notas"]})
    for _, row in df_cardio.iterrows():
        if row["Ejercicio"]:
            s = row.get("Series", "-"); r = row.get("Reps", "-")
            nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Cardio", "Orden": "-", "Ejercicio": row["Ejercicio"], "Series": s, "Reps": r, "Kg": "-", "Notas": row["Notas"]})

    if not all_data:
        df_final = pd.DataFrame(nuevas_filas)
        for c in cols: 
            if c not in df_final.columns: df_final[c] = ""
        df_final = df_final[cols]
        ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())
        return

    df_old = pd.DataFrame(all_data)
    df_old["Alumno"] = df_old["Alumno"].astype(str).str.strip()
    df_old["Dia"] = df_old["Dia"].astype(str).str.strip()
    mask = ~((df_old["Alumno"] == alumno) & (df_old["Dia"] == dia))
    df_clean = df_old[mask]
    df_nuevas = pd.DataFrame(nuevas_filas)
    df_final = pd.concat([df_clean, df_nuevas], ignore_index=True)
    df_final = df_final.fillna("")
    for c in cols: 
        if c not in df_final.columns: df_final[c] = ""
    df_final = df_final[cols]
    ws.clear()
    ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())

def guardar_registro(usuario, ejercicio, peso, reps, rpe, notas):
    sh = conectar_google_sheet()
    fecha = datetime.now().strftime("%Y-%m-%d") 
    sh.worksheet("Registros").append_row([fecha, usuario, ejercicio, peso, reps, rpe, notas])

def guardar_estado_sesion(usuario, estado):
    sh = conectar_google_sheet()
    fecha = datetime.now().strftime("%Y-%m-%d")
    sh.worksheet("Sesiones").append_row([fecha, usuario, estado])

def generar_word(alumno, df_rutina):
    doc = Document()
    doc.add_heading(f'Rutina: {alumno}', 0)
    dias = df_rutina["Dia"].unique()
    for dia in dias:
        doc.add_heading(dia, level=1)
        rutina_dia = df_rutina[df_rutina["Dia"] == dia]
        c = rutina_dia[rutina_dia["Seccion"] == "Calentamiento"]
        if not c.empty:
            doc.add_heading('Calentamiento', 2)
            t = doc.add_table(1,3); t.style='Table Grid'
            r=t.rows[0].cells; r[0].text='Ejer'; r[1].text='Series'; r[2].text='Notas'
            for _, row in c.iterrows(): 
                rr=t.add_row().cells; rr[0].text=str(row["Ejercicio"]); rr[1].text=f"{row['Series']}x{row['Reps']}"; rr[2].text=str(row["Notas"])
        f = rutina_dia[rutina_dia["Seccion"] == "Fuerza"]
        if not f.empty:
            doc.add_heading('Fuerza', 2)
            t = doc.add_table(1,5); t.style='Table Grid'
            r=t.rows[0].cells; r[0].text='Ord'; r[1].text='Ejer'; r[2].text='Ser'; r[3].text='Rep'; r[4].text='Kg'
            for _, row in f.iterrows():
                rr=t.add_row().cells; rr[0].text=str(row["Orden"]); rr[1].text=str(row["Ejercicio"]); rr[2].text=str(row["Series"]); rr[3].text=str(row["Reps"]); rr[4].text=str(row["Kg"])
        ca = rutina_dia[rutina_dia["Seccion"] == "Cardio"]
        if not ca.empty:
            doc.add_heading('Cardio', 2)
            t = doc.add_table(1,3); t.style='Table Grid'
            r=t.rows[0].cells; r[0].text='Ejer'; r[1].text='Tiempo/Int'; r[2].text='Notas'
            for _, row in ca.iterrows():
                rr=t.add_row().cells; rr[0].text=str(row["Ejercicio"]); rr[1].text=f
