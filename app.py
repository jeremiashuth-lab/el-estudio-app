import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import gspread
from google.oauth2.service_account import Credentials
import os
from io import BytesIO
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
import calendar
import time
import re

# --- 1. CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(
    page_title="El Estudio", 
    page_icon="🔥", 
    layout="wide", 
    initial_sidebar_state="expanded"
)

# --- 2. ESTILOS CSS (OPTIMIZADO PARA INPUTS MÓVILES) ---
def cargar_estilos():
    st.markdown("""
        <style>
        @import url('https://fonts.googleapis.com/css2?family=Montserrat:wght@300;400;600;800&display=swap');
        html, body, [class*="css"] { 
            font-family: 'Montserrat', sans-serif; 
            /* Se eliminó scroll-behavior para arreglar el teclado en móviles */
        }
        
        .block-container { padding-top: 1.5rem; padding-bottom: 5rem; }

        h1 { color: #E63946 !important; font-weight: 800 !important; letter-spacing: -1px; margin-bottom: 0.5rem; }
        h2, h3 { font-weight: 600 !important; margin-top: 1rem; }
        
        /* Métricas */
        div[data-testid="stMetric"] { background-color: #1A1C24; border: 1px solid #333; padding: 10px; border-radius: 8px; }
        div[data-testid="stMetricValue"] { color: #E63946 !important; font-weight: 800; font-size: 1.8rem !important; }

        /* Botones */
        div.stButton > button:first-child {
            background-color: #E63946; color: white; border-radius: 12px; border: none;
            font-weight: 600; text-transform: uppercase; letter-spacing: 1px;
            padding: 18px 20px; width: 100%;
            transition: all 0.2s ease;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        div.stButton > button:first-child:active { transform: scale(0.98); }
        
        /* Inputs: Forzamos tamaño de fuente para evitar zoom en iOS */
        input, textarea, select, div[data-baseweb="select"] { 
            font-size: 16px !important; 
            border-radius: 8px !important; 
        }
        
        /* Calendario Compacto */
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

# --- 4. FUNCIONES DE LECTURA (CACHED) ---
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
        
        def extraer_numero(texto):
            try: return float(re.search(r"(\d+[.,]?\d*)", str(texto)).group(1).replace(",", "."))
            except: return 0.0
        
        # Procesamiento seguro de columnas numéricas
        if "Peso" in df_alumno.columns: 
            df_alumno["Peso_Grafico"] = df_alumno["Peso"].apply(extraer_numero)
        
        col_reps = next((c for c in df_alumno.columns if "rep" in c.lower()), None)
        if col_reps: 
            df_alumno["Reps_Grafico"] = df_alumno[col_reps].apply(extraer_numero)
        else: 
            df_alumno["Reps_Grafico"] = 0.0
            
    return df_alumno

# --- 5. FUNCIONES DE ESCRITURA ---
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

    df_final = df_final.fillna("")
    for c in cols: 
        if c not in df_final.columns: df_final[c] = ""
        
    ws.clear()
    ws.update([cols] + df_final[cols].values.tolist())

def guardar_registro(usuario, ejercicio, peso, reps, rpe, notas, fecha_input=None):
    sh = conectar_google_sheet()
    fecha = fecha_input.strftime("%Y-%m-%d") if fecha_input else datetime.now().strftime("%Y-%m-%d") 
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
    
    df_otros = df_all[df_all["Usuario"].astype(str).str.lower() != usuario.lower()]
    
    df_nuevos = df_editado.copy()
    if "Fecha" in df_nuevos.columns:
        df_nuevos["Fecha"] = pd.to_datetime(df_nuevos["Fecha"]).dt.strftime("%Y-%m-%d")
        
    df_final = pd.concat([df_otros, df_nuevos], ignore_index=True).fillna("").sort_values("Fecha")
    
    ws.clear()
    ws.update([df_final.columns.values.tolist()] + df_final.values.tolist())

def guardar_estado_sesion(usuario, estado):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Sesiones")
    hoy = datetime.now().strftime("%Y-%m-%d")
    ws.append_row([hoy, usuario, estado])

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
    doc.add_paragraph(f"Fecha: {datetime.now().strftime('%d/%m/%Y')}")
    doc.add_paragraph("-" * 50)
    
    for dia in df_rutina["Dia"].unique():
        doc.add_heading(dia, level=1)
        rutina_dia = df_rutina[df_rutina["Dia"] == dia]
        
        for sec in ["Calentamiento", "Fuerza", "Cardio"]:
            df_sec = rutina_dia[rutina_dia["Seccion"] == sec]
            if not df_sec.empty:
                doc.add_heading(sec, level=2)
                for _, row in df_sec.iterrows():
                    ej = str(row['Ejercicio'])
                    if ej and ej != "":
                        detalle = f"{ej}"
                        if str(row.get('Series','')): detalle += f" | {row.get('Series','')} series"
                        if str(row.get('Reps','')): detalle += f" | {row.get('Reps','')} reps"
                        if str(row.get('Kg','')) and str(row.get('Kg','')) != "-": detalle += f" | {row.get('Kg','')} kg"
                        if str(row.get('Notas','')): detalle += f" ({row.get('Notas','')})"
                        doc.add_paragraph(detalle, style='List Bullet')
    
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

# --- 6. GESTIÓN DE LOGIN (URL PERMANENTE) ---
if 'logueado' not in st.session_state: st.session_state['logueado'] = False

if not st.session_state['logueado']:
    params = st.query_params
    user_url = params.get("u", None)
    
    if user_url:
        u_data = recuperar_usuario_por_nombre(user_url)
        if u_data is not None:
            st.session_state['logueado'] = True
            st.session_state['usuario_info'] = u_data
        else:
            st.query_params.clear()

if not st.session_state['logueado']:
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("<br><br><h1 style='text-align: center;'>EL ESTUDIO 🔥</h1><br>", unsafe_allow_html=True)
        with st.container(border=True):
            with st.form("login"):
                u = st.text_input("USUARIO")
                p = st.text_input("CONTRASEÑA", type="password")
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
    # --- APP PRINCIPAL ---
    datos = st.session_state['usuario_info']
    rol = str(datos.get('Rol', 'alumno')).strip()
    nombre = str(datos.get('Nombre', 'Usuario')).strip()
    alias = str(datos.get('Usuario', '')).strip()

    # SIDEBAR
    with st.sidebar:
        st.markdown(f"## {nombre.upper()}")
        st.caption(f"ROL: {rol.upper()}")
        st.divider()
        if st.button("🔴 LIMPIAR CACHÉ", use_container_width=True): 
            st.cache_data.clear()
            st.rerun()
        if st.button("CERRAR SESIÓN", use_container_width=True):
            st.session_state['logueado'] = False
            st.query_params.clear()
            st.rerun()

    if rol == "admin":
        st.title("PANEL DE CONTROL")
        tab1, tab2 = st.tabs(["DISEÑO", "ESTADÍSTICAS"])
        
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
            
            with st.expander("EDITOR DE RUTINA", expanded=True):
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
                    st.cache_data.clear()
                    st.success("Guardado.")
                    time.sleep(1)
                    st.rerun()
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
                df_s = leer_sesiones_alumno(alu_s)
                now = datetime.now()
                render_calendar(now.year, now.month, df_s)
            
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
                            actualizar_registros_usuario(alu_s, edited_hist)
                            st.cache_data.clear()
                            st.rerun()
                     else: st.info("Sin registros.")
                if not df_r.empty and "Peso_Grafico" in df_r.columns:
                    lista_ejercicios = df_r["Ejercicio"].unique()
                    if len(lista_ejercicios) > 0:
                        ej_v = st.selectbox("Gráfico de:", lista_ejercicios)
                        df_plt = df_r[df_r["Ejercicio"] == ej_v].sort_values("Fecha", ascending=True)
                        st.line_chart(df_plt.set_index("Fecha")["Peso_Grafico"], color="#E63946")

    else:
        # VISTA ALUMNO
        st.title(f"RUTINA DE {nombre.upper()}")
        t1, t2 = st.tabs(["ENTRENAR", "PROGRESO"])
        
        with t1:
            rut = leer_rutina(alias)
            if not rut.empty:
                # Header Rutina
                with st.container(border=True):
                    col_d, col_w = st.columns([3, 1])
                    with col_d:
                        dias = rut["Dia"].unique()
                        d_hoy = st.selectbox("Selecciona Día", dias)
                    with col_w:
                         st.markdown("<br>", unsafe_allow_html=True)
                         st.download_button("📥 Word", generar_word(alias, rut), "Rutina.docx", use_container_width=True)
                
                r_hoy = rut[rut["Dia"] == d_hoy]
                cfg_link = st.column_config.LinkColumn("Video", display_text="📺")
                
                # Tablas de lectura
                for seccion in ["Calentamiento", "Fuerza", "Cardio"]:
                    df_sec = r_hoy[r_hoy["Seccion"] == seccion]
                    if not df_sec.empty:
                        st.markdown(f"#### {seccion}")
                        cols_show = ["Ejercicio", "Series", "Reps", "Link", "Notas"]
                        if seccion == "Fuerza": cols_show.insert(3, "Kg")
                        st.dataframe(df_sec[cols_show], hide_index=True, use_container_width=True, column_config={"Link": cfg_link})

                st.markdown("---")
                st.subheader("📝 Registrar Serie")
                
                with st.form("registro_rapido"):
                    # FILTRO: SOLO EJERCICIOS DE FUERZA
                    ejercicios_fuerza = r_hoy[r_hoy["Seccion"] == "Fuerza"]["Ejercicio"].unique()
                    
                    if len(ejercicios_fuerza) == 0:
                        st.warning("No hay ejercicios de fuerza para este día.")
                        ej_sel = st.text_input("Ejercicio") # Fallback manual
                    else:
                        c_ej, c_kg = st.columns([2, 1])
                        ej_sel = c_ej.selectbox("Ejercicio (Solo Fuerza)", ejercicios_fuerza)
                        kg_in = c_kg.text_input("Kilos", placeholder="Ej: 50")
                    
                    c_reps, c_rpe = st.columns(2)
                    reps_in = c_reps.text_input("Reps", placeholder="Ej: 10")
                    rpe_in = c_rpe.slider("RPE (Esfuerzo)", 1, 10, 7)
                    notas_in = st.text_area("Notas", height=80)
                    
                    guardar = st.form_submit_button("GUARDAR REGISTRO", use_container_width=True)
                    if guardar:
                        guardar_registro(alias, ej_sel, kg_in, reps_in, rpe_in, notas_in)
                        st.cache_data.clear()
                        st.success("Registrado!")
                        time.sleep(1)

                c_ok, c_fail = st.columns(2)
                if c_ok.button("✅ FINALIZAR ENTRENO", use_container_width=True, type="primary"):
                    guardar_estado_sesion(alias, "Completado")
                    st.cache_data.clear()
                    st.balloons()
                if c_fail.button("⚠️ INCOMPLETO", use_container_width=True):
                    guardar_estado_sesion(alias, "Incompleto")
                    st.cache_data.clear()
            else:
                st.info("No tienes rutina asignada aún.")

        with t2:
            st.markdown("### Mi Constancia")
            dfs = leer_sesiones_alumno(alias)
            now = datetime.now()
            
            # KPI Metrics Restaurados
            c_kpi1, c_kpi2 = st.columns(2)
            count_month = 0
            count_year = 0
            if not dfs.empty:
                 count_month = dfs[(dfs["Fecha"].dt.year == now.year) & (dfs["Fecha"].dt.month == now.month)]["Fecha"].nunique()
                 count_year = dfs[dfs["Fecha"].dt.year == now.year]["Fecha"].nunique()
            
            c_kpi1.metric("Entrenos Mes", count_month)
            c_kpi2.metric("Total Año", count_year)

            render_calendar(now.year, now.month, dfs)
            
            st.markdown("### Mi Progreso")
            
            # SECCIÓN DE EDICIÓN PARA ALUMNO (NUEVA)
            with st.expander("🛠️ Corregir Historial"):
                df_r_edit = leer_registros_alumno(alias)
                if not df_r_edit.empty:
                    cols_show = ["Fecha", "Ejercicio", "Peso", "Repeticiones", "Rpe", "Notas"]
                    # Normalización
                    for c in cols_show: 
                        if c not in df_r_edit.columns: df_r_edit[c] = ""
                    
                    # Editor
                    edited_data_alu = st.data_editor(
                        df_r_edit[cols_show], 
                        num_rows="dynamic", 
                        key=f"editor_alumno_{alias}",
                        use_container_width=True
                    )
                    
                    if st.button("Guardar Correcciones"):
                        # Re-inyectar usuario para guardar
                        edited_data_alu["Usuario"] = alias
                        actualizar_registros_usuario(alias, edited_data_alu)
                        st.cache_data.clear()
                        st.success("Historial corregido.")
                        time.sleep(1)
                        st.rerun()
                else:
                    st.info("No hay registros para editar.")

            # GRÁFICO Y BITÁCORA VISUAL
            df_r = leer_registros_alumno(alias)
            if not df_r.empty and "Peso_Grafico" in df_r.columns:
                 lista_ej = df_r["Ejercicio"].unique()
                 if len(lista_ej) > 0:
                     ej_sel = st.selectbox("Ver:", lista_ej)
                     df_plt = df_r[df_r["Ejercicio"] == ej_sel].sort_values("Fecha", ascending=True)
                     
                     # Calculo 1RM
                     df_plt['1RM'] = df_plt['Peso_Grafico'] * (1 + (df_plt['Reps_Grafico'] / 30))
                     
                     st.line_chart(df_plt.set_index("Fecha")["1RM"], color="#E63946")
                     st.caption("📈 La línea muestra tu fuerza estimada (1RM). Si sube, ¡estás más fuerte!")
                     
                     st.markdown("#### 🗂️ Bitácora Visual")
                     # Bitácora visual tipo "Feed" restaurada
                     df_feed = df_plt.sort_values("Fecha", ascending=False)
                     for idx, row in df_feed.iterrows():
                         with st.container(border=True):
                             c_1, c_2 = st.columns([1, 4])
                             with c_1:
                                 st.write(f"**{row['Fecha'].strftime('%d/%m')}**")
                             with c_2:
                                 reps_txt = str(row.get('Repeticiones', row.get('Reps', '')))
                                 peso_txt = str(row.get('Peso', ''))
                                 st.write(f"💪 **{peso_txt}** |  🔄 **{reps_txt} reps**")
                                 if str(row.get('Notas','')).strip():
                                     st.caption(f"📝 {row['Notas']}")
