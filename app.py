import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import os
from io import BytesIO
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import calendar
import time
import re
import itertools
import math
import altair as alt

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="El Estudio", 
    page_icon="🔥", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- 2. ESTILOS CSS ---
def cargar_estilos():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        html, body, [class*="css"] { 
            font-family: 'Montserrat', sans-serif; 
            -webkit-text-size-adjust: 100%;
        }
        
        .block-container { padding-top: 1.5rem; padding-bottom: 5rem; }

        h1 { color: #E63946 !important; font-weight: 800 !important; letter-spacing: -1px; margin-bottom: 0.5rem; }
        h2, h3 { font-weight: 600 !important; margin-top: 1rem; }
        
        /* Botones */
        div.stButton > button:first-child {
            background-color: #E63946; color: white; border-radius: 12px; border: none;
            font-weight: 600; text-transform: uppercase; letter-spacing: 1px;
            padding: 18px 20px; width: 100%;
            transition: all 0.2s ease;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        div.stButton > button:first-child:active { transform: scale(0.98); }
        
        /* Inputs */
        input, textarea, select, div[data-baseweb="select"] { 
            font-size: 16px !important; 
            border-radius: 8px !important; 
            min-height: 45px;
        }
        
        /* Estilo de Expander (Tarjetas) */
        .streamlit-expanderHeader {
            font-weight: 600;
            background-color: #262730;
            border-radius: 8px;
        }
        
        /* Calendario */
        .calendar-container {
            display: grid; grid-template-columns: repeat(7, 1fr); gap: 4px; margin-top: 10px; margin-bottom: 20px;
        }
        .calendar-day {
            aspect-ratio: 1; display: flex; align-items: center; justify-content: center;
            border-radius: 6px; font-weight: bold; font-size: 0.85rem; color: #FFF; background-color: #262730;
        }
        .day-completed { background-color: #2ECC71 !important; color: #000 !important; }
        .day-incomplete { background-color: #F39C12 !important; color: #000 !important; }
        
        div[data-testid="stVerticalBlock"] { gap: 1rem; }

        /* --- MODO KIOSCO / OCULTAR INTERFAZ --- */
        #MainMenu {visibility: hidden !important;}
        footer {visibility: hidden !important; display: none !important;}
        header {visibility: hidden !important;}
        
        [data-testid="stToolbar"] {
            visibility: hidden !important;
            display: none !important;
        }
        
        [data-testid="stHeader"] {
            visibility: hidden !important;
            background: transparent !important;
        }
        
        .stAppDeployButton {
            display: none !important;
            visibility: hidden !important;
        }

        /* Estilo para las tarjetas de RM */
        .rm-card {
            background-color: #262730;
            padding: 15px;
            border-radius: 10px;
            text-align: center;
            border: 1px solid #333;
        }
        .rm-title { font-size: 0.9rem; color: #aaa; }
        .rm-value { font-size: 1.5rem; font-weight: 800; color: #E63946; }
        .rm-unit { font-size: 0.8rem; color: #aaa; }
        
        /* Métricas personalizadas */
        .big-metric {
            font-size: 3rem;
            font-weight: 800;
            color: #E63946;
            text-align: center;
        }
        .sub-metric {
            font-size: 1rem;
            color: #aaa;
            text-align: center;
            margin-bottom: 20px;
        }

        /* --- BORDE ROJO GRUESO PARA BLOQUES AGRUPADOS --- */
        div[data-testid="stVerticalBlockBorderWrapper"]:has(.red-border-trigger) {
            border: 3px solid #E63946 !important;
            border-radius: 10px !important;
        }
        
        /* --- CONGELAR GRÁFICOS (Hacerlos no clickeables ni arrastrables) --- */
        [data-testid="stArrowVegaLiteChart"], [data-testid="stVegaLiteChart"], canvas.marks {
            pointer-events: none !important;
        }
        </style>
    """, unsafe_allow_html=True)

cargar_estilos()

# --- 3. CONEXIÓN GOOGLE SHEETS ---
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

@st.cache_resource
def conectar_google_sheet():
    if os.path.exists("mis_secretos.json"):
        creds = Credentials.from_service_account_file("mis_secretos.json", scopes=SCOPES)
    else:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds).open("El Estudio DB")

def natural_sort_key(s):
    return [int(text) if text.isdigit() else text.lower() for text in re.split('([0-9]+)', str(s))]

# --- 4. FUNCIONES DE LECTURA ---
@st.cache_data(ttl=600)
def obtener_todos_usuarios():
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Usuarios").get_all_records())
    if df.empty: return df
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
        df_alumno["Fecha"] = pd.to_datetime(df_alumno["Fecha"], format='mixed', errors='coerce').dt.normalize()
    return df_alumno

@st.cache_data(ttl=600)
def leer_registros_alumno(alumno):
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Registros").get_all_records())
    if df.empty: return df
    col_map = {c: c.strip().capitalize() for c in df.columns}
    for k, v in col_map.items():
        if v.lower() in ['user', 'alumno']: col_map[k] = 'Usuario'
        if v.lower() in ['date', 'day']: col_map[k] = 'Fecha'
    df = df.rename(columns=col_map)
    if "Usuario" not in df.columns: return pd.DataFrame()
    df["Usuario_Norm"] = df["Usuario"].astype(str).str.strip().str.lower()
    alumno_norm = alumno.strip().lower()
    df_alumno = df[df["Usuario_Norm"] == alumno_norm].copy()
    if not df_alumno.empty and "Fecha" in df_alumno.columns:
        df_alumno["Fecha"] = pd.to_datetime(df_alumno["Fecha"], format='mixed', errors='coerce').dt.normalize()
        df_alumno = df_alumno.dropna(subset=["Fecha"])
        df_alumno = df_alumno.reset_index(drop=True)
        def extraer_numero(texto):
            try: return float(re.search(r"(\d+[.,]?\d*)", str(texto)).group(1).replace(",", "."))
            except: return 0.0
        if "Peso" in df_alumno.columns: df_alumno["Peso_Grafico"] = df_alumno["Peso"].apply(extraer_numero)
        col_reps = next((c for c in df_alumno.columns if "rep" in c.lower()), None)
        if col_reps: df_alumno["Reps_Grafico"] = df_alumno[col_reps].apply(extraer_numero)
        else: df_alumno["Reps_Grafico"] = 0.0
    return df_alumno

# --- LÓGICA DE CÁLCULO 1RM PROMEDIO ---
def calcular_1rm_promedio(peso, reps):
    if reps == 1: return peso
    if reps == 0: return 0
    # Fórmulas
    brzycki = peso / (1.0278 - 0.0278 * reps)
    epley = peso * (1 + 0.0333 * reps)
    lander = (100 * peso) / (101.3 - 2.67123 * reps)
    lombardi = peso * (reps ** 0.10)
    mayhew = (100 * peso) / (52.2 + (41.9 * math.exp(-0.055 * reps)))
    oconner = peso * (1 + 0.025 * reps)
    wathen = (100 * peso) / (48.8 + (53.8 * math.exp(-0.075 * reps)))
    # Promedio
    promedio = (brzycki + epley + lander + lombardi + mayhew + oconner + wathen) / 7
    return promedio

# --- 5. FUNCIONES DE ESCRITURA (BLINDADAS) ---
def obtener_usuario(usuario_input, password_input):
    try:
        df = obtener_todos_usuarios()
        if df.empty: return None
        u_input = usuario_input.strip().lower()
        if "Usuario" not in df.columns: return None
        df["User_Lower"] = df["Usuario"].astype(str).str.strip().str.lower()
        usuario = df[(df["User_Lower"] == u_input) & (df["Password"].astype(str).str.strip() == password_input.strip())]
        return usuario.iloc[0] if not usuario.empty else None
    except: return None

def recuperar_usuario_por_nombre(usuario_input):
    try:
        df = obtener_todos_usuarios()
        if df.empty: return None
        if "Usuario" not in df.columns: return None
        u_input = usuario_input.strip().lower()
        df["User_Lower"] = df["Usuario"].astype(str).str.strip().str.lower()
        usuario = df[df["User_Lower"] == u_input]
        return usuario.iloc[0] if not usuario.empty else None
    except: return None

def guardar_rutina_actualizada(alumno, dia, df_c, df_f, df_ca):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Rutinas")
    all_data = ws.get_all_records()
    cols = ["Alumno", "Dia", "Seccion", "Orden", "Ejercicio", "Link", "Series", "Reps", "Kg", "Notas"]
    nuevas_filas = []
    
    def procesar_df(df, seccion):
        df = df.fillna("")
        for _, row in df.iterrows():
            if str(row.get("Ejercicio","")).strip() or str(row.get("Notas","")).strip():
                nuevas_filas.append({
                    "Alumno": alumno, "Dia": dia, "Seccion": seccion,
                    "Orden": str(row.get("Orden","-")), "Ejercicio": str(row.get("Ejercicio","")),
                    "Link": str(row.get("Link","")), "Series": str(row.get("Series","")),
                    "Reps": str(row.get("Reps","")), "Kg": str(row.get("Kg","-")),
                    "Notas": str(row.get("Notas",""))
                })
    
    procesar_df(df_c, "Calentamiento")
    procesar_df(df_f, "Fuerza")
    procesar_df(df_ca, "Cardio")

    df_old = pd.DataFrame(all_data)
    
    if not df_old.empty:
        df_old.columns = [c.strip().capitalize() for c in df_old.columns]
        mask = ~((df_old["Alumno"].astype(str).str.lower() == alumno.lower()) & (df_old["Dia"].astype(str) == dia))
        df_clean = df_old[mask]
        df_final = pd.concat([df_clean, pd.DataFrame(nuevas_filas)], ignore_index=True)
    else:
        df_final = pd.DataFrame(nuevas_filas)

    # PROTOCOLO DE SEGURIDAD
    filas_antes = len(df_old)
    filas_despues = len(df_final)
    if filas_antes > 20 and filas_despues < 10:
        st.error(f"🚨 ERROR DE SEGURIDAD CRÍTICO: La aplicación intentó borrar datos. Guardado cancelado.")
        return 

    df_final = df_final.fillna("")
    for c in cols: 
        if c not in df_final.columns: df_final[c] = ""
        
    ws.clear()
    ws.update([cols] + df_final[cols].values.tolist())

def guardar_desde_tarjetas(alumno, dia, rut_hoy, session_state):
    rows_c = []; rows_f = []; rows_ca = []
    for idx, row in rut_hoy.iterrows():
        seccion = row["Seccion"]
        key_kg = f"kg_{idx}"
        key_notas = f"notas_{idx}"
        
        nuevo_kg = str(session_state.get(key_kg, row.get("Kg", "")))
        nueva_nota = str(session_state.get(key_notas, row.get("Notas", "")))
        
        new_row = row.copy()
        new_row["Kg"] = nuevo_kg
        new_row["Notas"] = nueva_nota
        if seccion == "Calentamiento": rows_c.append(new_row)
        elif seccion == "Fuerza": rows_f.append(new_row)
        elif seccion == "Cardio": rows_ca.append(new_row)
    df_c = pd.DataFrame(rows_c) if rows_c else pd.DataFrame()
    df_f = pd.DataFrame(rows_f) if rows_f else pd.DataFrame()
    df_ca = pd.DataFrame(rows_ca) if rows_ca else pd.DataFrame()
    guardar_rutina_actualizada(alumno, dia, df_c, df_f, df_ca)

def guardar_registro(usuario, ejercicio, peso, reps, rpe, notas, fecha_input=None):
    sh = conectar_google_sheet()
    fecha = fecha_input.strftime("%Y-%m-%d") if fecha_input else datetime.now().strftime("%Y-%m-%d") 
    sh.worksheet("Registros").append_row([fecha, usuario, ejercicio, peso, reps, rpe, notas])

def editar_un_registro_especifico(usuario, fecha_original, ejercicio_original, nuevo_peso, nuevas_reps, nueva_nota):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Registros")
    all_data = ws.get_all_records()
    fila_idx = -1
    for i, row in enumerate(all_data):
        r_user = str(row.get("Usuario", "")).strip().lower()
        r_date = str(row.get("Fecha", "")).strip()
        r_ej = str(row.get("Ejercicio", "")).strip()
        if r_user == usuario.strip().lower() and r_date == fecha_original and r_ej == ejercicio_original:
            fila_idx = i + 2
            break
    if fila_idx != -1:
        ws.update_cell(fila_idx, 4, nuevo_peso)
        ws.update_cell(fila_idx, 5, nuevas_reps)
        ws.update_cell(fila_idx, 7, nueva_nota)
        return True
    return False

def actualizar_registros_masivo(usuario, df_editado):
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
    df_otros = df_all[df_all["Usuario"].astype(str).str.lower() != usuario.lower()]
    df_nuevos = df_editado.copy()
    if "Fecha" in df_nuevos.columns:
        df_nuevos["Fecha"] = pd.to_datetime(df_nuevos["Fecha"]).dt.strftime("%Y-%m-%d")
    df_final = pd.concat([df_otros, df_nuevos], ignore_index=True).fillna("").sort_values("Fecha")
    ws.clear()
    ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())

def guardar_estado_sesion(usuario, estado, fecha_dt=None):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Sesiones")
    fecha_str = fecha_dt.strftime("%Y-%m-%d") if fecha_dt else datetime.now().strftime("%Y-%m-%d")
    ws.append_row([fecha_str, usuario, estado])

def preparar_df_editor(df_input, columnas, filas_minimas=4):
    if df_input.empty: df_input = pd.DataFrame(columns=columnas)
    for c in ["Series", "Reps", "Kg"]:
        if c in df_input.columns: df_input[c] = df_input[c].astype(str)
    filas_actuales = len(df_input)
    if filas_actuales < filas_minimas:
        extras = pd.DataFrame([[""] * len(columnas)] * (filas_minimas - filas_actuales), columns=columnas)
        df_input = pd.concat([df_input, extras], ignore_index=True)
    return df_input

def generar_word(alumno, df_rutina):
    doc = Document()
    p_titulo = doc.add_heading(f'RUTINA: {alumno.upper()}', 0)
    p_titulo.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph(f"Generado: {datetime.now().strftime('%d/%m/%Y')}")
    doc.add_paragraph("-" * 50)
    dias_ordenados = sorted(df_rutina["Dia"].unique(), key=natural_sort_key)
    for dia in dias_ordenados:
        doc.add_heading(dia, level=1)
        rutina_dia = df_rutina[df_rutina["Dia"] == dia]
        emoji_map = {"Calentamiento": "🔥", "Fuerza": "🏋️‍♂️", "Cardio": "🏃‍♂️"}
        for sec in ["Calentamiento", "Fuerza", "Cardio"]:
            df_sec = rutina_dia[rutina_dia["Seccion"] == sec]
            if not df_sec.empty:
                titulo_sec = f"{emoji_map.get(sec, '')} {sec}"
                doc.add_heading(titulo_sec, level=2)
                cols_count = 5 if sec == "Fuerza" else 4
                table = doc.add_table(rows=1, cols=cols_count)
                table.style = 'Table Grid'
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = 'EJERCICIO'
                hdr_cells[1].text = 'SERIES'
                hdr_cells[2].text = 'REPS'
                if sec == "Fuerza":
                    hdr_cells[3].text = 'KG'
                    hdr_cells[4].text = 'NOTAS'
                else:
                    hdr_cells[3].text = 'NOTAS'
                for _, row in df_sec.iterrows():
                    row_cells = table.add_row().cells
                    nom_ej = str(row['Ejercicio'])
                    if sec == "Fuerza":
                        orden = str(row.get('Orden', '')).strip()
                        if orden and orden != '-': nom_ej = f"{orden}. {nom_ej}"
                    row_cells[0].text = nom_ej
                    row_cells[1].text = str(row.get('Series', ''))
                    row_cells[2].text = str(row.get('Reps', ''))
                    if sec == "Fuerza":
                        row_cells[3].text = str(row.get('Kg', ''))
                        row_cells[4].text = str(row.get('Notas', ''))
                    else:
                        row_cells[3].text = str(row.get('Notas', ''))
                doc.add_paragraph("")
    b = BytesIO(); doc.save(b); b.seek(0); return b

def render_calendar(year, month, df_sesiones):
    nombres_meses = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
    month_name = nombres_meses[month]
    cal = calendar.monthcalendar(year, month)
    asistencia_map = {}
    if not df_sesiones.empty:
        mask = (df_sesiones["Fecha"].dt.year == year) & (df_sesiones["Fecha"].dt.month == month)
        df_mes = df_sesiones[mask]
        for _, row in df_mes.iterrows(): asistencia_map[row["Fecha"].day] = row["Estado"]
    html = f"""<div style="text-align:center; font-weight:bold; font-size:1.1rem; color:#E63946; margin-bottom:8px;">{month_name} {year}</div>
    <div class="calendar-container">
    <div style="text-align:center; color:#888;">L</div><div style="text-align:center; color:#888;">M</div><div style="text-align:center; color:#888;">M</div><div style="text-align:center; color:#888;">J</div><div style="text-align:center; color:#888;">V</div><div style="text-align:center; color:#888;">S</div><div style="text-align:center; color:#888;">D</div>"""
    for week in cal:
        for day in week:
            if day == 0: html += '<div></div>'
            else:
                css = "calendar-day"
                if asistencia_map.get(day) == "Completado": css += " day-completed"
                elif asistencia_map.get(day) == "Incompleto": css += " day-incomplete"
                html += f'<div class="{css}">{day}</div>'
    html += "</div>"
    st.markdown(html, unsafe_allow_html=True)

def renderizar_bloque_seccion(titulo, df_seccion):
    if df_seccion.empty: return
    st.markdown(f"#### {titulo}")
    def obtener_bloque(orden):
        o = str(orden).strip().upper()
        if len(o) > 0 and o[0].isalpha() and o != "-":
            return o[0] 
        return "SIN_BLOQUE"
    datos_con_indice = []
    for idx, row in df_seccion.iterrows():
        datos_con_indice.append((idx, row, obtener_bloque(row.get("Orden", ""))))
    for bloque, grupo in itertools.groupby(datos_con_indice, key=lambda x: x[2]):
        items = list(grupo)
        if bloque != "SIN_BLOQUE":
            with st.container(border=True):
                # --- HACK SILENCIOSO PARA EL BORDE ROJO ---
                st.markdown("<div class='red-border-trigger'></div>", unsafe_allow_html=True)
                for idx, row, _ in items:
                    render_tarjeta_individual(idx, row)
        else:
            for idx, row, _ in items:
                render_tarjeta_individual(idx, row)

def render_tarjeta_individual(idx, row):
    ej_nombre = row['Ejercicio']
    series_reps = f"{row.get('Series','?')} x {row.get('Reps','?')}"
    orden_prefix = ""
    if "Orden" in row and str(row["Orden"]).strip() not in ["", "-"]:
        orden_prefix = f"**{row['Orden']}** | "
    titulo_card = f"{orden_prefix}{ej_nombre} ({series_reps})"
    with st.expander(titulo_card):
        link = str(row.get('Link', '')).strip()
        if link: st.link_button("📺 Ver Video", link, use_container_width=True)
        c_kg, c_notas = st.columns([1, 2])
        k_kg = f"kg_{idx}"
        k_notas = f"notas_{idx}"
        if k_kg in st.session_state: val_kg = st.session_state[k_kg]
        else:
            val_kg = str(row.get('Kg', ''))
            if val_kg == "nan": val_kg = ""
        if k_notas in st.session_state: val_notas = st.session_state[k_notas]
        else:
            val_notas = str(row.get('Notas', ''))
            if val_notas == "nan": val_notas = ""
        with c_kg: st.text_input("Kg Realizados", value=val_kg, key=k_kg)
        with c_notas: st.text_area("Notas / Instrucciones", value=val_notas, key=k_notas, height=68)

# --- 6. GESTIÓN DE LOGIN ---
if 'logueado' not in st.session_state: st.session_state['logueado'] = False

if not st.session_state['logueado']:
    params = st.query_params
    user_url = params.get("u", None)
    if user_url:
        u_data = recuperar_usuario_por_nombre(user_url)
        if u_data is not None:
            st.session_state['logueado'] = True
            st.session_state['usuario_info'] = u_data
        else: st.query_params.clear()

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
                        st.session_state['logueado'] = True
                        st.session_state['usuario_info'] = user
                        if str(user.get('Rol','')).lower() == 'alumno':
                            st.query_params["u"] = user['Usuario']
                        st.rerun()
                    else: st.error("❌ Error")
else:
    datos = st.session_state['usuario_info']
    rol = str(datos.get('Rol', 'alumno')).strip()
    nombre = str(datos.get('Nombre', 'Usuario')).strip()
    alias = str(datos.get('Usuario', '')).strip()

    with st.sidebar:
        st.markdown(f"## {nombre.upper()}")
        st.caption(f"ROL: {rol.upper()}")
        st.divider()
        if st.button("🔴 LIMPIAR CACHÉ", use_container_width=True): 
            st.cache_data.clear(); st.rerun()
        if st.button("🚪 CERRAR SESIÓN", use_container_width=True):
            st.session_state['logueado'] = False; st.query_params.clear(); st.rerun()

    if rol == "admin":
        st.title("🎛️ PANEL DE CONTROL")
        tab1, tab2 = st.tabs(["📝 DISEÑO", "📊 ESTADÍSTICAS"])
        with tab1:
            with st.container(border=True):
                c1, c2, c3 = st.columns([3, 3, 1]) 
                with c3: 
                    st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
                    if st.button("🔄", key="ref_admin"): st.cache_data.clear(); st.rerun()
                us = obtener_todos_usuarios()
                als = us[us["Rol"] == "alumno"]["Usuario"].tolist()
                with c1: alu = st.selectbox("ALUMNO", als)
                with c2: dia = st.selectbox("DÍA", ["Día 1", "Día 2", "Día 3", "Día 4"])

            rut = leer_rutina(alu)
            cols_c = ["Ejercicio", "Link", "Series", "Reps", "Notas"]
            cols_f = ["Orden", "Ejercicio", "Link", "Series", "Reps", "Kg", "Notas"]
            cols_ca = ["Ejercicio", "Link", "Series", "Reps", "Notas"]
            d_cal = pd.DataFrame(columns=cols_c)
            d_fue = pd.DataFrame(columns=cols_f)
            d_car = pd.DataFrame(columns=cols_ca)

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
            col_link = st.column_config.LinkColumn("Link", display_text="🔗")
            
            with st.expander("✏️ EDITOR DE RUTINA", expanded=True):
                st.caption("CALENTAMIENTO")
                ed_c = st.data_editor(d_cal, num_rows="dynamic", use_container_width=True, key=f"c_{alu}_{dia}", column_config={"Link": col_link})
                st.caption("FUERZA")
                ed_f = st.data_editor(d_fue, num_rows="dynamic", use_container_width=True, height=400, key=f"f_{alu}_{dia}", column_config={"Link": col_link})
                st.caption("CARDIO")
                ed_ca = st.data_editor(d_car, num_rows="dynamic", use_container_width=True, key=f"ca_{alu}_{dia}", column_config={"Link": col_link})
            
            c_g, c_d = st.columns([1, 1])
            with c_g:
                if st.button("💾 GUARDAR RUTINA", type="primary", use_container_width=True):
                    guardar_rutina_actualizada(alu, dia, ed_c, ed_f, ed_ca)
                    st.cache_data.clear(); st.success("Guardado."); time.sleep(1); st.rerun()
            with c_d:
                if not rut.empty: 
                    st.download_button("📥 DESCARGAR WORD", generar_word(alu, rut), f"{alu}.docx", use_container_width=True)

        with tab2:
            us = obtener_todos_usuarios()
            als = us[us["Rol"] == "alumno"]["Usuario"].tolist()
            alu_s = st.selectbox("VER DATOS DE:", als, key="stats_sel")
            
            c_cal, c_hist = st.columns([1, 2])
            with c_cal:
                st.markdown("### 📅 Asistencia")
                dfs = leer_sesiones_alumno(alu_s)
                
                MESES_LIST = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
                c_sel_m, c_sel_y = st.columns([2, 1])
                mes_actual_idx = datetime.now().month
                sel_mes_nom = c_sel_m.selectbox("Mes", MESES_LIST[1:], index=mes_actual_idx-1, key="admin_month")
                sel_anio = c_sel_y.number_input("Año", value=datetime.now().year, key="admin_year")
                sel_mes_idx = MESES_LIST.index(sel_mes_nom)

                count_month = 0; count_year = 0
                if not dfs.empty:
                     count_month = dfs[(dfs["Fecha"].dt.year == sel_anio) & (dfs["Fecha"].dt.month == sel_mes_idx)]["Fecha"].nunique()
                     count_year = dfs[dfs["Fecha"].dt.year == sel_anio]["Fecha"].nunique()
                
                km1, km2 = st.columns(2)
                km1.metric(f"Mes", count_month)
                km2.metric(f"Año", count_year)

                render_calendar(int(sel_anio), sel_mes_idx, dfs)

            with c_hist:
                st.markdown("### 📈 Historial")
                df_r = leer_registros_alumno(alu_s)
                with st.expander("Gestión de Registros"):
                     if not df_r.empty:
                        cols_show = ["Fecha", "Usuario", "Ejercicio", "Peso", "Repeticiones", "Rpe", "Notas"]
                        for c in cols_show: 
                            if c not in df_r.columns: df_r[c] = ""
                        edited_hist = st.data_editor(df_r[cols_show], num_rows="dynamic", key=f"hist_edit_{alu_s}")
                        if st.button("Actualizar Historial"):
                            actualizar_registros_masivo(alu_s, edited_hist)
                            st.cache_data.clear(); st.rerun()
                     else: st.info("Sin registros.")
                if not df_r.empty and "Peso_Grafico" in df_r.columns:
                    lista_ejercicios = df_r["Ejercicio"].unique()
                    if len(lista_ejercicios) > 0:
                        ej_v = st.selectbox("Gráfico de:", lista_ejercicios)
                        df_plt = df_r[df_r["Ejercicio"] == ej_v].sort_values("Fecha", ascending=True)
                        if not df_plt.empty:
                            df_plt['1RM'] = df_plt['Peso_Grafico'] * (1 + (df_plt['Reps_Grafico'] / 30))
                            st.bar_chart(df_plt.set_index("Fecha")["1RM"], color="#E63946")

    else:
        # VISTA ALUMNO
        st.title(f"RUTINA DE {nombre.upper()}")
        t1, t2, t3 = st.tabs(["💪 ENTRENAR", "📅 ASISTENCIA", "📝 BITÁCORA"])
        
        with t1:
            rut = leer_rutina(alias)
            if not rut.empty:
                with st.container(border=True):
                    col_d, col_w = st.columns([3, 1])
                    with col_d:
                        dias_disponibles = sorted(rut["Dia"].unique(), key=natural_sort_key)
                        d_hoy = st.selectbox("Selecciona Día", dias_disponibles, key="selector_dia_alumno")

                    with col_w:
                         st.markdown("<br>", unsafe_allow_html=True)
                         st.download_button("📥 Word", generar_word(alias, rut), "Rutina.docx", use_container_width=True)
                
                r_hoy = rut[rut["Dia"] == d_hoy]
                
                notas_importantes = []
                for idx, row in r_hoy.iterrows():
                    if str(row.get('Notas','')).strip() != "":
                        notas_importantes.append(f"**{row['Ejercicio']}**: {row['Notas']}")
                
                if notas_importantes:
                    st.info("📝 **NOTAS:**\n\n" + "\n".join([f"- {n}" for n in notas_importantes]))

                df_sec_c = r_hoy[r_hoy["Seccion"] == "Calentamiento"]
                df_sec_f = r_hoy[r_hoy["Seccion"] == "Fuerza"]
                df_sec_ca = r_hoy[r_hoy["Seccion"] == "Cardio"]

                renderizar_bloque_seccion("🔥 Calentamiento", df_sec_c)
                renderizar_bloque_seccion("🏋️‍♂️ Fuerza", df_sec_f)
                renderizar_bloque_seccion("🏃‍♂️ Cardio", df_sec_ca)

                st.markdown("<br>", unsafe_allow_html=True)
                
                if st.button("💾 GUARDAR CAMBIOS EN LA RUTINA", type="primary", use_container_width=True):
                    guardar_desde_tarjetas(alias, d_hoy, r_hoy, st.session_state)
                    st.cache_data.clear()
                    st.toast("✅ Notas y Kilos guardados!")
                    time.sleep(1)
                    st.rerun()

                st.markdown("---")
                with st.expander("📝 Registro Rápido (Guardar serie en historial)", expanded=False):
                    with st.form("registro_rapido"):
                        ejercicios_fuerza = r_hoy[r_hoy["Seccion"] == "Fuerza"]["Ejercicio"].unique()
                        if len(ejercicios_fuerza) == 0:
                            ej_sel = st.text_input("Ejercicio")
                        else:
                            c_ej, c_kg = st.columns([2, 1])
                            ej_sel = c_ej.selectbox("Ejercicio", ejercicios_fuerza)
                            kg_in = c_kg.text_input("Kilos", placeholder="Ej: 50")
                        c_reps, c_rpe = st.columns(2)
                        reps_in = c_reps.text_input("Reps", placeholder="Ej: 10")
                        rpe_in = c_rpe.slider("RPE", 1, 10, 7)
                        notas_in = st.text_area("Notas extra", height=80)
                        if st.form_submit_button("GUARDAR EN HISTORIAL", use_container_width=True):
                            guardar_registro(alias, ej_sel, kg_in, reps_in, rpe_in, notas_in)
                            st.cache_data.clear(); st.success("Guardado!"); time.sleep(1)

                c_ok, c_fail = st.columns(2)
                if c_ok.button("✅ TERMINÉ POR HOY", use_container_width=True):
                    guardar_estado_sesion(alias, "Completado"); st.cache_data.clear(); st.balloons()
                if c_fail.button("⚠️ INCOMPLETO", use_container_width=True):
                    guardar_estado_sesion(alias, "Incompleto"); st.cache_data.clear()
            else: st.info("No tienes rutina asignada aún.")
            
        with t2:
            st.markdown("### 📅 Asistencia")
            dfs = leer_sesiones_alumno(alias)
            
            MESES_LIST = ["", "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]
            c_sel_m, c_sel_y = st.columns([2, 1])
            mes_actual_idx = datetime.now().month
            sel_mes_nom = c_sel_m.selectbox("Mes", MESES_LIST[1:], index=mes_actual_idx-1)
            sel_anio = c_sel_y.number_input("Año", value=datetime.now().year)
            sel_mes_idx = MESES_LIST.index(sel_mes_nom)
            
            c_kpi1, c_kpi2 = st.columns(2)
            count_month = 0; count_year = 0
            if not dfs.empty:
                 count_month = dfs[(dfs["Fecha"].dt.year == sel_anio) & (dfs["Fecha"].dt.month == sel_mes_idx)]["Fecha"].nunique()
                 count_year = dfs[dfs["Fecha"].dt.year == sel_anio]["Fecha"].nunique()
            c_kpi1.metric(f"Entrenos {sel_mes_nom}", count_month)
            c_kpi2.metric(f"Total {sel_anio}", count_year)
            render_calendar(int(sel_anio), sel_mes_idx, dfs)
            
            with st.expander("🗓️ ¿Olvidaste dar el presente?"):
                st.caption("Registra un entrenamiento que hiciste antes y olvidaste marcar.")
                col_d_pas, col_b_pas = st.columns([2,1])
                fecha_pasada = col_d_pas.date_input("Fecha del entreno", value=datetime.now().date() - timedelta(days=1), max_value=datetime.now().date())
                if col_b_pas.button("Marcar Completado", use_container_width=True):
                    guardar_estado_sesion(alias, "Completado", fecha_pasada)
                    st.success(f"Entreno del {fecha_pasada.strftime('%d/%m')} guardado.")
                    st.cache_data.clear(); time.sleep(1); st.rerun()

        with t3:
            st.markdown("### 📝 Bitácora")
            
            df_r = leer_registros_alumno(alias)
            ultimo_1rm_val = 0
            
            if not df_r.empty and "Peso_Grafico" in df_r.columns:
                lista_ej = df_r["Ejercicio"].unique()
                if len(lista_ej) > 0:
                    ej_sel = st.selectbox("Selecciona Ejercicio:", lista_ej)
                    df_plt = df_r[df_r["Ejercicio"] == ej_sel].sort_values("Fecha", ascending=True)
                    
                    if not df_plt.empty:
                        st.caption(f"Historial de Cargas: {ej_sel}")
                        
                        # --- GRÁFICO 100% ESTÁTICO Y CONGELADO ---
                        df_chart = df_plt.copy()
                        df_chart["Reps_Label"] = df_chart["Reps_Grafico"].astype(int).astype(str) + " reps"
                        
                        chart = alt.Chart(df_chart).mark_bar(color="#E63946").encode(
                            x=alt.X('Reps_Label:O', sort=None, title="Repeticiones", axis=alt.Axis(labelAngle=0)),
                            y=alt.Y('Peso_Grafico:Q', title="Kilos")
                        ).properties(height=300) 
                        
                        st.altair_chart(chart, use_container_width=True)
                        
                        cols_view = ["Fecha", "Peso", "Repeticiones", "Rpe", "Notas"]
                        for c in cols_view:
                            if c not in df_plt.columns: df_plt[c] = "-"
                        st.dataframe(df_plt[cols_view].sort_values("Fecha", ascending=False), use_container_width=True, hide_index=True)
                        
                        last_row = df_plt.iloc[-1]
                        ultimo_1rm_val = calcular_1rm_promedio(last_row["Peso_Grafico"], last_row["Reps_Grafico"])
                    else:
                        st.info("Sin datos para este ejercicio.")
                else:
                    st.info("No hay ejercicios registrados.")
            else:
                st.info("Aún no tienes registros en tu bitácora.")

            st.markdown("---")
            
            st.markdown(f"<div class='sub-metric'>ESTIMACIÓN ACTUAL (Basado en tu último registro)</div>", unsafe_allow_html=True)
            st.markdown(f"<div class='big-metric'>{int(ultimo_1rm_val)} kg</div>", unsafe_allow_html=True)
            
            st.markdown("---")

            st.markdown("### 🧮 Calculadora")
            with st.container(border=True):
                c_peso, c_reps = st.columns(2)
                p_in = c_peso.number_input("Peso (kg)", min_value=0.0, step=1.0, value=0.0)
                r_in = c_reps.number_input("Repeticiones", min_value=1, max_value=30, step=1, value=5)
            
            rm_calculo = 0
            if p_in > 0:
                rm_calculo = calcular_1rm_promedio(p_in, r_in)
            elif ultimo_1rm_val > 0:
                rm_calculo = ultimo_1rm_val
            
            if rm_calculo > 0:
                # --- TABLAS DENTRO DEL DESPLEGABLE ---
                with st.expander("📊 Ver Tablas de RM y % de Fuerza", expanded=False):
                    st.caption(f"TABLA DE FUERZA (Base: {int(rm_calculo)}kg)")
                    cols_grid = st.columns(4)
                    for i in range(1, 13):
                        peso_rm_i = rm_calculo * (1.0278 - 0.0278 * i)
                        if i == 1: peso_rm_i = rm_calculo
                        
                        with cols_grid[(i-1)%4]:
                            st.markdown(f"""
                            <div class="rm-card">
                                <div class="rm-title">{i}RM</div>
                                <div class="rm-value">{int(peso_rm_i)}</div>
                                <div class="rm-unit">kg</div>
                            </div>
                            <div style="margin-bottom:10px;"></div>
                            """, unsafe_allow_html=True)

                    st.markdown("<br>", unsafe_allow_html=True)
                    st.markdown("#### % Cargas de Trabajo")
                    
                    # --- CALCULADORA DE PORCENTAJE PERSONALIZADO ---
                    with st.container(border=True):
                        st.caption("🎯 **Calculadora de % Exacto**")
                        col_calc1, col_calc2 = st.columns([1, 2])
                        porc_custom = col_calc1.number_input("Ingresa %", min_value=1.0, max_value=200.0, value=72.5, step=0.5)
                        val_custom = rm_calculo * (porc_custom / 100)
                        col_calc2.info(f"🔥 **{porc_custom}%** = {val_custom:.1f} kg")
                    
                    st.markdown("<br>", unsafe_allow_html=True)
                    
                    # --- TABLA DE PORCENTAJES FIJOS ---
                    col_p1, col_p2 = st.columns(2)
                    porcentajes = [125, 120, 115, 110, 105, 100, 95, 90, 85, 80, 75, 70, 65, 60, 55, 50, 45, 40, 35, 30]
                    
                    for idx, porc in enumerate(porcentajes):
                        val_p = rm_calculo * (porc / 100)
                        texto = f"**{porc}%** : {int(val_p)} kg"
                        if idx % 2 == 0: col_p1.info(texto)
                        else: col_p2.info(texto)
