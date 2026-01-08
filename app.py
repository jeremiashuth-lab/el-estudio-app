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

# --- ESTILOS CSS (SÚPER OPTIMIZADO PARA IOS) ---
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
        
        /* FIX CRÍTICO IPHONE */
        div.stButton > button:first-child {
            background-color: #E63946; color: white; border-radius: 8px; border: none;
            font-weight: 600; text-transform: uppercase; letter-spacing: 1px;
            padding-top: 15px; padding-bottom: 15px;
            touch-action: manipulation;
            -webkit-tap-highlight-color: transparent;
        }
        input, textarea, select { font-size: 16px !important; }
        div[data-testid="stVerticalBlock"] { gap: 1rem; }
        .stSpinner > div { border-top-color: #E63946 !important; }
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
        df_final = df_final.sort_values("Fecha", ascending=True)

    ws.clear()
    ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())

def guardar_estado_sesion(usuario, estado):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Sesiones")
    fecha_hoy = datetime.now().strftime("%Y-%m-%d")
    records = ws.get_all_records()
    df = pd.DataFrame(records)
    ya_registrado = False
    if not df.empty:
        df.columns = [c.strip().capitalize() for c in df.columns]
        df["Usuario_Norm"] = df["Usuario"].astype(str).str.strip().str.lower()
        usuario_norm = usuario.strip().lower()
        df["Fecha_Str"] = pd.to_datetime(df["Fecha"], format='mixed').dt.strftime("%Y-%m-%d")
        filtro = (df["Usuario_Norm"] == usuario_norm) & (df["Fecha_Str"] == fecha_hoy)
        if filtro.any():
            ya_registrado = True
    if not ya_registrado:
        ws.append_row([fecha_hoy, usuario, estado])

def preparar_df_editor(df_input, columnas, filas_minimas=4):
    if df_input.empty: df_input = pd.DataFrame(columns=columnas)
    for c in ["Series", "Reps", "Kg"]:
        if c in df_input.columns: df_input[c] = df_input[c].astype(str)
    filas_actuales = len(df_input)
    if filas_actuales < filas_minimas:
        filas_a_agregar = filas_minimas - filas_actuales
        empty_rows = pd.DataFrame([[""] * len(columnas)] * filas_a_agregar, columns=columnas)
        df_input = pd.concat([df_input, empty_rows], ignore_index=True)
    return df_input

def generar_word(alumno, df_rutina):
    doc = Document()
    p_titulo = doc.add_heading(f'RUTINA DE ENTRENAMIENTO: {alumno.upper()}', 0)
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Generado el: {datetime.now().strftime('%d/%m/%Y')}")
    doc.add_paragraph("-" * 50)
    dias = df_rutina["Dia"].unique()
    for dia in dias:
        doc.add_heading(dia, level=1)
        rutina_dia = df_rutina[df_rutina["Dia"] == dia]
        def format_link(link): return str(link) if link and str(link).strip() != "" else "-"
        c = rutina_dia[rutina_dia["Seccion"] == "Calentamiento"]
        if not c.empty:
            doc.add_heading('🔥 Entrada en Calor', level=2)
            table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
            hdr = table.rows[0].cells; hdr[0].text='EJERCICIO'; hdr[1].text='PAUTA'; hdr[2].text='LINK'; hdr[3].text='NOTAS'
            for _, row in c.iterrows():
                cells = table.add_row().cells; cells[0].text=str(row["Ejercicio"]); cells[1].text=f"{row.get('Series','')} x {row.get('Reps','')}"; cells[2].text=format_link(row.get("Link", "")); cells[3].text=str(row["Notas"])
        doc.add_paragraph("")
        f = rutina_dia[rutina_dia["Seccion"] == "Fuerza"]
        if not f.empty:
            doc.add_heading('🏋️‍♂️ Fuerza', level=2)
            table = doc.add_table(rows=1, cols=5); table.style = 'Table Grid'
            hdr = table.rows[0].cells; hdr[0].text='EJERCICIO'; hdr[1].text='SER x REP'; hdr[2].text='KG'; hdr[3].text='LINK'; hdr[4].text='NOTAS'
            for _, row in f.iterrows():
                cells = table.add_row().cells
                o = str(row.get("Orden","")).strip(); e = str(row["Ejercicio"]).strip()
                cells[0].text = f"{o}. {e}" if o and o != "-" else e
                cells[1].text = f"{row.get('Series','')} x {row.get('Reps','')}"; cells[2].text = str(row.get("Kg", "")); cells[3].text = format_link(row.get("Link", "")); cells[4].text = str(row["Notas"])
        doc.add_paragraph("")
        ca = rutina_dia[rutina_dia["Seccion"] == "Cardio"]
        if not ca.empty:
            doc.add_heading('🏃‍♂️ Cardio', level=2)
            table = doc.add_table(rows=1, cols=4); table.style = 'Table Grid'
            hdr = table.rows[0].cells; hdr[0].text='EJERCICIO'; hdr[1].text='TIEMPO/INT'; hdr[2].text='LINK'; hdr[3].text='NOTAS'
            for _, row in ca.iterrows():
                cells = table.add_row().cells; cells[0].text=str(row["Ejercicio"]); cells[1].text=f"{row.get('Series','')} | {row.get('Reps','')}"; cells[2].text=format_link(row.get("Link", "")); cells[3].text=str(row["Notas"])
        doc.add_page_break()
    b = BytesIO(); doc.save(b); b.seek(0); return b

def render_calendar(year, month, df_sesiones):
    cal = calendar.monthcalendar(year, month)
    month_name = MESES_ESP[month]
    asistencia_map = {}
    if not df_sesiones.empty:
        mask = (df_sesiones["Fecha"].dt.year == year) & (df_sesiones["Fecha"].dt.month == month)
        df_mes = df_sesiones[mask]
        for _, row in df_mes.iterrows(): asistencia_map[row["Fecha"].day] = row["Estado"]
    html = f"""<div style="text-align:center; margin-bottom:10px; font-weight:bold; font-size:1.2rem; color:#E63946;">{month_name} {year}</div>
    <div class="calendar-container"><div class="calendar-day-header">L</div><div class="calendar-day-header">M</div><div class="calendar-day-header">M</div><div class="calendar-day-header">J</div><div class="calendar-day-header">V</div><div class="calendar-day-header">S</div><div class="calendar-day-header">D</div>"""
    for week in cal:
        for day in week:
            if day == 0: html += '<div class="calendar-day day-empty"></div>'
            else:
                css = "calendar-day"
                if asistencia_map.get(day) == "Completado": css += " day-completed"
                elif asistencia_map.get(day) == "Incompleto": css += " day-incomplete"
                html += f'<div class="{css}">{day}</div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

# --- GESTIÓN DE SESIÓN ROBUSTA (VERSIÓN 34.0: DOBLE VERIFICACIÓN) ---
cookie_manager = stx.CookieManager(key="login_cookies")

if 'logueado' not in st.session_state: 
    st.session_state['logueado'] = False

# FASE 1: Verificación Agresiva de Cookie
if not st.session_state['logueado']:
    # INTENTO 1
    c_user = cookie_manager.get(cookie="gym_user")
    
    # Si falla, PAUSAMOS 1.5s (Trampa de tiempo para móviles) y reintentamos
    if not c_user:
        with st.spinner('Cargando...'):
            time.sleep(1.5)
            c_user = cookie_manager.get(cookie="gym_user")
            
    if c_user:
        user = obtener_usuario_por_cookie(c_user)
        if user is not None:
            st.session_state['logueado'] = True
            st.session_state['usuario_info'] = user
            # Renovamos cookie
            exp_date = datetime.now() + timedelta(days=30)
            cookie_manager.set("gym_user", c_user, expires_at=exp_date)
            st.rerun()

# FASE 2: Mostrar Login
if not st.session_state['logueado']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><h1 style='text-align: center;'>EL ESTUDIO 🔥</h1><br>", unsafe_allow_html=True)
        with st.container(border=True):
            with st.form("login"):
                u = st.text_input("USUARIO"); p = st.text_input("CONTRASEÑA", type="password")
                st.markdown("<br>", unsafe_allow_html=True)
                if st.form_submit_button("ENTRAR", use_container_width=True):
                    user = obtener_usuario(u, p)
                    if user is not None: 
                        exp_date = datetime.now() + timedelta(days=30)
                        cookie_manager.set("gym_user", u, expires_at=exp_date)
                        st.session_state['logueado'] = True; st.session_state['usuario_info'] = user; st.rerun()
                    else: st.error("❌ Error")
else:
    datos = st.session_state['usuario_info']
    # --- BLINDAJE ANTI-ERROR (Safe Unpack) ---
    rol = str(datos.get('Rol', 'alumno')).strip()
    nombre = str(datos.get('Nombre', 'Usuario')).strip()
    alias = str(datos.get('Usuario', '')).strip()
    # Si nombre está vacío o es 'nan', poner default
    if nombre.lower() in ['nan', 'none', '']: nombre = "Usuario"

    with st.sidebar:
        st.markdown(f"## {nombre.upper()}"); st.caption(f"ROL: {rol.upper()}")
        st.markdown("---")
        if st.button("🔴 LIMPIAR CACHÉ", use_container_width=True): st.cache_data.clear(); st.rerun()
        if st.button("SALIR", use_container_width=True):
            try: cookie_manager.delete("gym_user")
            except: pass
            st.session_state['logueado'] = False; st.rerun()

    if rol == "admin":
        st.title("PANEL DE CONTROL")
        tab1, tab2 = st.tabs(["DISEÑO", "ESTADÍSTICAS"])
        with tab1:
            c1, c2, c3 = st.columns([3, 3, 1]) 
            with c3: 
                st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                if st.button("🔄", key="refresh_users"): st.cache_data.clear(); st.rerun()
            us = obtener_todos_usuarios()
            als = us[us["Rol"] == "alumno"]["Usuario"].tolist()
            with c1: alu = st.selectbox("ALUMNO", als)
            with c2: dia = st.selectbox("DÍA", ["Día 1", "Día 2", "Día 3", "Día 4"])
            rut = leer_rutina(alu)
            
            cols_c = ["Ejercicio", "Link", "Series", "Reps", "Notas"]
            cols_f = ["Orden", "Ejercicio", "Link", "Series", "Reps", "Kg", "Notas"]
            cols_ca = ["Ejercicio", "Link", "Series", "Reps", "Notas"]
            d_cal, d_fue, d_car = pd.DataFrame(columns=cols_c), pd.DataFrame(columns=cols_f), pd.DataFrame(columns=cols_ca)
            if not rut.empty:
                r_dia = rut[rut["Dia"] == dia]
                if not r_dia.empty:
                    c = r_dia[r_dia["Seccion"] == "Calentamiento"]
                    if not c.empty: d_cal = c[cols_c]
                    f = r_dia[r_dia["Seccion"] == "Fuerza"]
                    if not f.empty: d_fue = f[cols_f]; d_fue["Kg"] = d_fue["Kg"].astype(str)
                    ca = r_dia[r_dia["Seccion"] == "Cardio"]
                    if not ca.empty: d_car = ca[cols_ca]
            d_cal = preparar_df_editor(d_cal, cols_c, filas_minimas=4)
            d_fue = preparar_df_editor(d_fue, cols_f, filas_minimas=8)
            d_car = preparar_df_editor(d_car, cols_ca, filas_minimas=3)
            
            st.markdown("---")
            with st.container(border=True):
                col_link = st.column_config.LinkColumn("Link", display_text="🔗 Video")
                col_text = st.column_config.TextColumn()
                st.caption("CALENTAMIENTO")
                ed_c = st.data_editor(d_cal, num_rows="dynamic", use_container_width=True, key=f"c_{alu}_{dia}", column_config={"Link": col_link, "Series": col_text, "Reps": col_text})
                st.caption("FUERZA")
                ed_f = st.data_editor(d_fue, num_rows="dynamic", use_container_width=True, height=400, key=f"f_{alu}_{dia}", column_config={"Kg": col_text, "Link": col_link, "Series": col_text, "Reps": col_text})
                st.caption("CARDIO")
                ed_ca = st.data_editor(d_car, num_rows="dynamic", use_container_width=True, key=f"ca_{alu}_{dia}", column_config={"Series": st.column_config.TextColumn("Tiempo/Dist"), "Reps": st.column_config.TextColumn("Intensidad"), "Link": col_link})
            c_g, c_r, c_d = st.columns([1, 1, 1])
            with c_g:
                if st.button("💾 GUARDAR", type="primary", use_container_width=True):
                    guardar_rutina_actualizada(alu, dia, ed_c, ed_f, ed_ca); st.cache_data.clear(); st.success("Guardado."); st.rerun()
            with c_r:
                 if st.button("🔄 RECARGAR", use_container_width=True): st.cache_data.clear(); st.rerun()
            with c_d:
                if not rut.empty: st.download_button("📥 WORD", generar_word(alu, rut), f"{alu}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

        with tab2:
            us = obtener_todos_usuarios()
            als = us[us["Rol"] == "alumno"]["Usuario"].tolist()
            alu_s = st.selectbox("VER DATOS DE:", als)
            st.markdown("### 📅 Rendimiento")
            df_s = leer_sesiones_alumno(alu_s)
            now = datetime.now()
            c_mes, c_anio = st.columns([2, 1])
            sel_mes = c_mes.selectbox("Mes", MESES_ESP[1:], index=now.month-1)
            sel_anio = c_anio.number_input("Año", value=now.year, step=1)
            mes_idx = MESES_ESP.index(sel_mes)
            if not df_s.empty:
                df_year = df_s[df_s["Fecha"].dt.year == sel_anio]
                count_year = df_year["Fecha"].dt.date.nunique()
                df_month = df_year[df_year["Fecha"].dt.month == mes_idx]
                count_month = df_month["Fecha"].dt.date.nunique()
                m1, m2 = st.columns(2)
                m1.metric(f"Total {sel_mes}", count_month)
                m2.metric(f"Total Año {sel_anio}", count_year)
                render_calendar(sel_anio, mes_idx, df_s)
            else: st.info("Sin datos."); render_calendar(sel_anio, mes_idx, pd.DataFrame())
            st.markdown("---")
            st.markdown("### 📈 Historial y Edición")
            df_r = leer_registros_alumno(alu_s)
            with st.expander("📝 GESTIONAR BITÁCORA (Borrar/Editar Registros)"):
                if not df_r.empty:
                    st.info("Selecciona las filas y presiona 'Supr' para borrar. Edita valores y luego presiona 'ACTUALIZAR'.")
                    cols_show = ["Fecha", "Usuario", "Ejercicio", "Peso", "Repeticiones", "Rpe", "Notas"]
                    for c in cols_show:
                        if c not in df_r.columns: df_r[c] = ""
                    df_to_edit = df_r[cols_show].copy()
                    edited_df = st.data_editor(df_to_edit, num_rows="dynamic", use_container_width=True, key=f"edit_hist_{alu_s}")
                    if st.button("💾 ACTUALIZAR HISTORIAL"):
                        actualizar_registros_usuario(alu_s, edited_df)
                        st.cache_data.clear(); st.success("Actualizado."); st.rerun()
                else: st.warning("No hay registros.")
            if not df_r.empty and "Peso_Grafico" in df_r.columns:
                lista_ejercicios = df_r["Ejercicio"].unique()
                if len(lista_ejercicios) > 0:
                    ej_v = st.selectbox("Ejercicio", lista_ejercicios)
                    df_plt = df_r[df_r["Ejercicio"] == ej_v].sort_values("Fecha", ascending=False)
                    st.line_chart(df_plt.set_index("Fecha")["Peso_Grafico"], color="#E63946")
                    
                    st.markdown("#### 🗂️ Bitácora Visual")
                    for idx, row in df_plt.iterrows():
                        with st.container(border=True):
                            c_date, c_data = st.columns([1, 3])
                            with c_date: 
                                st.markdown(f"**{row['Fecha'].strftime('%d/%m')}**")
                                st.caption(f"{row['Fecha'].year}")
                            with c_data:
                                reps_val = row.get('Reps', row.get('Repeticiones', 0))
                                st.markdown(f"💪 **{row['Peso']}** x  **{reps_val} reps**")
                                st.markdown(f"🔥 RPE: {row.get('Rpe', '-')}")
                                if str(row.get('Notas', '')) != "": st.info(f"📝 {row['Notas']}")
    else:
        st.title(f"RUTINA DE {nombre.upper()}")
        t1, t2 = st.tabs(["ENTRENAR", "PROGRESO"])
        with t1:
            rut = leer_rutina(alias)
            if not rut.empty:
                st.download_button("📥 Descargar Word", generar_word(alias, rut), "Rutina.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                dias = rut["Dia"].unique()
                d_hoy = st.selectbox("DÍA", dias)
                r_hoy = rut[rut["Dia"] == d_hoy]
                cfg_link = st.column_config.LinkColumn("Ver Video", display_text="📺 Ver Video", width="small")
                col_text_alu = st.column_config.TextColumn()
                c = r_hoy[r_hoy["Seccion"] == "Calentamiento"]
                if not c.empty:
                    st.markdown("### 🔥 Entrada en Calor")
                    ed_c_alumno = st.data_editor(c[["Ejercicio", "Link", "Series", "Reps", "Notas"]], hide_index=True, use_container_width=True, disabled=["Ejercicio","Link"], key=f"cal_alu_{d_hoy}", column_config={"Link": cfg_link, "Series": col_text_alu, "Reps": col_text_alu})
                else: ed_c_alumno = pd.DataFrame()
                f = r_hoy[r_hoy["Seccion"] == "Fuerza"]
                if not f.empty:
                    st.markdown("### 🏋️‍♂️ Fuerza")
                    f_display = f.copy()
                    f_display["Ejercicio_Full"] = f_display["Orden"] + ". " + f_display["Ejercicio"]
                    f_display["SxR"] = f_display["Series"].astype(str) + " x " + f_display["Reps"].astype(str)
                    f_final = f_display[["Ejercicio_Full", "Link", "SxR", "Kg", "Notas"]]
                    f_final["_Orden_Original"] = f_display["Orden"]
                    f_final["_Ejercicio_Original"] = f_display["Ejercicio"]
                    f_final["_Link_Original"] = f_display["Link"]
                    f_final["_Series_Original"] = f_display["Series"]
                    f_final["_Reps_Original"] = f_display["Reps"]
                    ed_f_alumno = st.data_editor(f_final, column_order=["Ejercicio_Full", "Link", "SxR", "Kg", "Notas"], column_config={"Ejercicio_Full": st.column_config.TextColumn("Ejercicio", disabled=True), "Link": cfg_link, "SxR": st.column_config.TextColumn("SxR", disabled=True), "Kg": col_text_alu, "Notas": st.column_config.TextColumn("Notas")}, hide_index=True, use_container_width=True, key=f"fue_alu_{d_hoy}")
                else: ed_f_alumno = pd.DataFrame()
                ca = r_hoy[r_hoy["Seccion"] == "Cardio"]
                if not ca.empty: 
                    st.markdown("### 🏃‍♂️ Cardio")
                    ed_ca_alumno = st.data_editor(ca[["Ejercicio", "Link", "Series", "Reps", "Notas"]], column_config={"Series":st.column_config.TextColumn("Tiempo/Dist"), "Reps":st.column_config.TextColumn("Intensidad", disabled=True), "Ejercicio":st.column_config.TextColumn(disabled=True), "Link": cfg_link}, hide_index=True, use_container_width=True, key=f"car_alu_{d_hoy}")
                else: ed_ca_alumno = pd.DataFrame()
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 GUARDAR CAMBIOS RUTINA", type="primary", use_container_width=True):
                    df_f_save = pd.DataFrame()
                    if not ed_f_alumno.empty:
                        df_f_save = ed_f_alumno.copy()
                        df_f_save["Orden"] = df_f_save["_Orden_Original"]
                        df_f_save["Ejercicio"] = df_f_save["_Ejercicio_Original"]
                        df_f_save["Link"] = df_f_save["_Link_Original"]
                        df_f_save["Series"] = df_f_save["_Series_Original"]
                        df_f_save["Reps"] = df_f_save["_Reps_Original"]
                    guardar_rutina_actualizada(alias, d_hoy, ed_c_alumno, df_f_save, ed_ca_alumno); st.cache_data.clear(); st.success("✅ Notas guardadas"); st.rerun()
                st.markdown("---")
                with st.expander("➕ AGREGAR REGISTRO EXTRA (Día Actual o Pasado)"):
                    with st.form("reg"):
                        lista_ej = f["Ejercicio"].unique() if 'f' in locals() and not f.empty else ["Varios"]
                        ej = st.selectbox("Ejercicio", lista_ej)
                        c_date, c_k = st.columns(2)
                        fecha_manual = c_date.date_input("Fecha", value=datetime.now())
                        k = c_k.text_input("Kilos (ej: 20kg)")
                        c_r, c_rp = st.columns(2)
                        reps = c_r.number_input("Reps", step=1, min_value=1)
                        rpe = c_rp.slider("RPE", 1, 10)
                        n = st.text_area("Notas")
                        if st.form_submit_button("GUARDAR REGISTRO"): 
                            guardar_registro(alias, ej, k, reps, rpe, n, fecha_manual); st.cache_data.clear(); st.success("Listo")
                c1, c2 = st.columns(2)
                with c1: 
                    if st.button("✅ LISTO", type="primary", use_container_width=True): guardar_estado_sesion(alias, "Completado"); st.cache_data.clear(); st.balloons()
                with c2: 
                    if st.button("⚠️ INCOMPLETO", use_container_width=True): guardar_estado_sesion(alias, "Incompleto"); st.cache_data.clear()
            else: st.info("No tienes rutina cargada.")
        with t2:
            st.markdown("### 📅 Mi Constancia")
            dfs = leer_sesiones_alumno(alias)
            now = datetime.now()
            c_mes, c_anio = st.columns([2, 1])
            sel_mes_al = c_mes.selectbox("Mes", MESES_ESP[1:], index=now.month-1, key="mes_al")
            sel_anio_al = c_anio.number_input("Año", value=now.year, step=1, key="anio_al")
            mes_idx_al = MESES_ESP.index(sel_mes_al)
            if not dfs.empty:
                df_year = dfs[dfs["Fecha"].dt.year == sel_anio_al]
                count_year = df_year["Fecha"].dt.date.nunique()
                df_month = df_year[df_year["Fecha"].dt.month == mes_idx_al]
                count_month = df_month["Fecha"].dt.date.nunique()
                m1, m2 = st.columns(2)
                m1.metric(f"Entrenos {sel_mes_al}", count_month)
                m2.metric(f"Total Año {sel_anio_al}", count_year)
                render_calendar(sel_anio_al, mes_idx_al, dfs)
            else: st.info("Sin datos."); render_calendar(sel_anio_al, mes_idx_al, pd.DataFrame())
            st.markdown("---")
            st.markdown("### 📈 Historial y Edición")
            df_r = leer_registros_alumno(alias)
            with st.expander("📝 CORREGIR MIS REGISTROS"):
                if not df_r.empty:
                    st.caption("Si te equivocaste, corrige el valor aquí y presiona ACTUALIZAR.")
                    cols_show = ["Fecha", "Usuario", "Ejercicio", "Peso", "Repeticiones", "Rpe", "Notas"]
                    for c in cols_show: 
                         if c not in df_r.columns: df_r[c] = ""
                    df_to_edit_al = df_r[cols_show].copy()
                    edited_df_al = st.data_editor(df_to_edit_al, num_rows="dynamic", use_container_width=True, key=f"edit_hist_al_{alias}")
                    if st.button("💾 ACTUALIZAR MIS REGISTROS"):
                        actualizar_registros_usuario(alias, edited_df_al); st.cache_data.clear(); st.success("Corregido."); st.rerun()
            if not df_r.empty and "Peso_Grafico" in df_r.columns:
                 lista_ej = df_r["Ejercicio"].unique()
                 if len(lista_ej) > 0:
                     ej_sel = st.selectbox("Ver progreso en:", lista_ej)
                     df_plt = df_r[df_r["Ejercicio"] == ej_sel].sort_values("Fecha", ascending=False)
                     st.line_chart(df_plt.set_index("Fecha")["Peso_Grafico"], color="#E63946")
                     
                     st.markdown("#### 🗂️ Bitácora Visual")
                     for idx, row in df_plt.iterrows():
                        with st.container(border=True):
                            c_date, c_data = st.columns([1, 3])
                            with c_date: 
                                st.markdown(f"**{row['Fecha'].strftime('%d/%m')}**")
                                st.caption(f"{row['Fecha'].year}")
                            with c_data:
                                reps_val = row.get('Reps', row.get('Repeticiones', 0))
                                st.markdown(f"💪 **{row['Peso']}** x  **{reps_val} reps**")
                                st.markdown(f"🔥 RPE: {row.get('Rpe', '-')}")
                                if str(row.get('Notas', '')) != "": st.info(f"📝 {row['Notas']}")
