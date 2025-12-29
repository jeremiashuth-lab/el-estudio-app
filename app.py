import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import os
import altair as alt
from io import BytesIO
from docx import Document
from docx.shared import Inches

# --- CONFIGURACIÓN DE PÁGINA ---
st.set_page_config(page_title="El Estudio", page_icon="🔥", layout="wide")

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
        
        div[data-testid="stMetric"] {
            background-color: #1A1C24; border: 1px solid #333; padding: 15px; border-radius: 8px;
        }
        div[data-testid="stMetricValue"] { color: #E63946 !important; font-weight: 700; }
        
        div.stButton > button:first-child {
            background-color: #E63946; color: white; border-radius: 6px; border: none;
            font-weight: 600; text-transform: uppercase; letter-spacing: 1px;
            padding-top: 10px; padding-bottom: 10px;
        }
        div.stButton > button:first-child:hover {
            background-color: #FF4D5A; box-shadow: 0 4px 10px rgba(230, 57, 70, 0.4);
        }
        </style>
    """, unsafe_allow_html=True)

cargar_estilos()

# --- CONEXIÓN GOOGLE SHEETS ---
SCOPES = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]

def conectar_google_sheet():
    if os.path.exists("mis_secretos.json"):
        creds = Credentials.from_service_account_file("mis_secretos.json", scopes=SCOPES)
    else:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds).open("El Estudio DB")

# --- FUNCIONES ---
def obtener_usuario(usuario_input, password_input):
    try:
        sh = conectar_google_sheet()
        df = pd.DataFrame(sh.worksheet("Usuarios").get_all_records())
        usuario = df[
            (df["Usuario"].astype(str).str.strip() == usuario_input.strip()) & 
            (df["Password"].astype(str).str.strip() == password_input.strip())
        ]
        return usuario.iloc[0] if not usuario.empty else None
    except: return None

def leer_rutina(alumno):
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Rutinas").get_all_records())
    if df.empty: return df
    
    df["Alumno"] = df["Alumno"].astype(str).str.strip()
    df["Seccion"] = df["Seccion"].astype(str).str.strip().str.capitalize()
    
    return df[df["Alumno"] == alumno.strip()]

# Función universal para guardar cambios (Sirve para Admin y Alumno)
def guardar_rutina_actualizada(alumno, dia, df_calentamiento, df_fuerza, df_cardio):
    sh = conectar_google_sheet()
    ws = sh.worksheet("Rutinas")
    all_data = ws.get_all_records()
    cols = ["Alumno", "Dia", "Seccion", "Orden", "Ejercicio", "Series", "Reps", "Kg", "Notas"]

    nuevas_filas = []
    
    # Procesamos Calentamiento
    for _, row in df_calentamiento.iterrows():
        if row["Ejercicio"]: 
             # Manejo de compatibilidad si usamos la vista comprimida
             s = row.get("Series", "2")
             r = row.get("Reps", "10")
             nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Calentamiento", "Orden": "-", "Ejercicio": row["Ejercicio"], "Series": s, "Reps": r, "Kg": "-", "Notas": row["Notas"]})
    
    # Procesamos Fuerza
    for _, row in df_fuerza.iterrows():
        if row["Ejercicio"]:
            # Si viene de la vista comprimida del alumno, necesitamos recuperar Series/Reps
            # Si existen columnas ocultas en el dataframe editado, las usamos
            s = row.get("Series", "3")
            r = row.get("Reps", "10")
            o = row.get("Orden", "-")
            
            # Limpieza del nombre si venía con el orden pegado (A1. Sentadilla -> Sentadilla)
            ej_nombre = row["Ejercicio"]
            
            nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Fuerza", "Orden": o, "Ejercicio": ej_nombre, "Series": s, "Reps": r, "Kg": row["Kg"], "Notas": row["Notas"]})
    
    # Procesamos Cardio
    for _, row in df_cardio.iterrows():
        if row["Ejercicio"]:
            s = row.get("Series", "-")
            r = row.get("Reps", "-")
            nuevas_filas.append({"Alumno": alumno, "Dia": dia, "Seccion": "Cardio", "Orden": "-", "Ejercicio": row["Ejercicio"], "Series": s, "Reps": r, "Kg": "-", "Notas": row["Notas"]})

    # Guardado seguro
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

def leer_sesiones_alumno(alumno):
    sh = conectar_google_sheet()
    df = pd.DataFrame(sh.worksheet("Sesiones").get_all_records())
    if df.empty: return df
    df_alumno = df[df["Usuario"].astype(str).str.strip() == alumno.strip()].copy()
    if not df_alumno.empty:
        df_alumno["Fecha"] = pd.to_datetime(df_alumno["Fecha"], format='mixed').dt.normalize()
    return df_alumno

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
                rr=t.add_row().cells; rr[0].text=str(row["Ejercicio"]); rr[1].text=f"{row['Series']} | {row['Reps']}"; rr[2].text=str(row["Notas"])
        doc.add_paragraph("\n")
    
    b = BytesIO(); doc.save(b); b.seek(0); return b

# --- INTERFAZ ---
if 'logueado' not in st.session_state:
    st.session_state['logueado'] = False

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
                        st.rerun()
                    else: st.error("❌ Error")

else:
    datos = st.session_state['usuario_info']
    rol, nombre, alias = datos['Rol'], datos['Nombre'], datos['Usuario']
    
    with st.sidebar:
        st.markdown(f"## {nombre.upper()}")
        st.caption(f"ROL: {rol.upper()}")
        st.markdown("---")
        if st.button("SALIR", use_container_width=True):
            st.session_state['logueado'] = False
            st.rerun()

    # --- ADMIN VIEW ---
    if rol == "admin":
        st.title("PANEL DE CONTROL")
        tab1, tab2 = st.tabs(["DISEÑO", "ESTADÍSTICAS"])
        with tab1:
            c1, c2 = st.columns(2)
            sh = conectar_google_sheet()
            us = pd.DataFrame(sh.worksheet("Usuarios").get_all_records())
            als = us[us["Rol"] == "alumno"]["Usuario"].tolist()
            alu = c1.selectbox("ALUMNO", als)
            dia = c2.selectbox("DÍA", ["Día 1", "Día 2", "Día 3", "Día 4"])
            
            rut = leer_rutina(alu)
            # Default templates
            d_cal = pd.DataFrame([{"Ejercicio": "", "Series": "2", "Reps": "10", "Notas": ""}]*4)
            d_fue = pd.DataFrame({"Orden":["A1","A2","B1","B2","C1","C2","D1","D2"], "Ejercicio":[""]*8, "Series":["3"]*8, "Reps":["8-12"]*8, "Kg":[0.0]*8, "Notas":[""]*8})
            d_car = pd.DataFrame([{"Ejercicio": "", "Series": "10'", "Reps": "RPE 6", "Notas": ""}]*2)
            
            if not rut.empty:
                r_dia = rut[rut["Dia"] == dia]
                if not r_dia.empty:
                    c = r_dia[r_dia["Seccion"] == "Calentamiento"]
                    if not c.empty: d_cal = c[["Ejercicio","Series","Reps","Notas"]]
                    f = r_dia[r_dia["Seccion"] == "Fuerza"]
                    if not f.empty: 
                        d_fue = f[["Orden","Ejercicio","Series","Reps","Kg","Notas"]]
                        d_fue["Kg"] = pd.to_numeric(d_fue["Kg"], errors='coerce').fillna(0.0)
                    ca = r_dia[r_dia["Seccion"] == "Cardio"]
                    if not ca.empty: d_car = ca[["Ejercicio","Series","Reps","Notas"]]
            
            st.markdown("---")
            with st.container(border=True):
                st.caption("CALENTAMIENTO")
                ed_c = st.data_editor(d_cal, num_rows="dynamic", use_container_width=True, key=f"c_{alu}_{dia}")
                st.caption("FUERZA")
                ed_f = st.data_editor(d_fue, num_rows="dynamic", use_container_width=True, height=350, key=f"f_{alu}_{dia}", column_config={"Kg": st.column_config.NumberColumn(format="%.1f")})
                st.caption("CARDIO")
                ed_ca = st.data_editor(d_car, num_rows="dynamic", use_container_width=True, key=f"ca_{alu}_{dia}", column_config={"Series": st.column_config.TextColumn("Tiempo/Dist"), "Reps": st.column_config.TextColumn("Intensidad")})

            c_g, c_d = st.columns([1,1])
            with c_g:
                if st.button("💾 GUARDAR", type="primary", use_container_width=True):
                    guardar_rutina_actualizada(alu, dia, ed_c, ed_f, ed_ca)
                    st.success("Guardado y reparado.")
                    st.rerun()
            with c_d:
                if not rut.empty:
                    st.download_button("📥 WORD", generar_word(alu, rut), f"{alu}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", use_container_width=True)

        with tab2:
            alu_s = st.selectbox("VER DATOS DE:", als)
            df_s = leer_sesiones_alumno(alu_s)
            if not df_s.empty:
                st.metric("Sesiones", len(df_s))
                ch = alt.Chart(df_s).mark_rect().encode(x='week(Fecha):O', y='day(Fecha):O', color=alt.Color('Estado', scale=alt.Scale(range=['#2ECC71', '#F39C12']))).properties(height=200)
                st.altair_chart(ch, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### 📈 Cargas")
            df_r = leer_registros_alumno(alu_s)
            if not df_r.empty and "Peso" in df_r.columns:
                lista_ejercicios = df_r["Ejercicio"].unique()
                if len(lista_ejercicios) > 0:
                    ej_v = st.selectbox("Ejercicio", lista_ejercicios)
                    df_plt = df_r[df_r["Ejercicio"] == ej_v].sort_values("Fecha")
                    
                    st.line_chart(df_plt.set_index("Fecha")["Peso"])
                    if "Repeticiones" in df_plt.columns:
                        st.line_chart(df_plt.set_index("Fecha")["Repeticiones"])
                else: st.info("Hay registros pero sin ejercicios válidos.")
            else: st.info("El alumno no ha registrado pesos todavía.")

    # --- ALUMNO VIEW ---
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
                
                # --- CALENTAMIENTO (Editable para que puedan poner notas) ---
                c = r_hoy[r_hoy["Seccion"] == "Calentamiento"]
                if not c.empty:
                    st.markdown("### 🔥 Entrada en Calor")
                    # Vista comprimida
                    c_display = c.copy()
                    ed_c_alumno = st.data_editor(
                        c_display[["Ejercicio", "Series", "Reps", "Notas"]],
                        hide_index=True,
                        use_container_width=True,
                        disabled=["Ejercicio", "Series", "Reps"], # Solo Notas editable
                        key=f"cal_alu_{d_hoy}"
                    )
                else: ed_c_alumno = pd.DataFrame() # Vacío
                
                # --- FUERZA (COMPRIMIDO Y EDITABLE) ---
                f = r_hoy[r_hoy["Seccion"] == "Fuerza"]
                if not f.empty:
                    st.markdown("### 🏋️‍♂️ Fuerza")
                    
                    # 1. Crear vista COMPRIMIDA para celular
                    f_display = f.copy()
                    
                    # Fusionamos ORDEN con EJERCICIO (Ej: "A1. Sentadilla")
                    f_display["Ejercicio_Full"] = f_display["Orden"] + ". " + f_display["Ejercicio"]
                    
                    # Fusionamos SERIES y REPS (Ej: "3 x 10")
                    f_display["SxR"] = f_display["Series"].astype(str) + " x " + f_display["Reps"].astype(str)
                    
                    # Columnas finales: Ejercicio | SxR | Kg | Notas
                    f_final = f_display[["Ejercicio_Full", "SxR", "Kg", "Notas"]]
                    
                    # Columnas ocultas (para guardar después)
                    f_final["_Orden_Original"] = f_display["Orden"]
                    f_final["_Ejercicio_Original"] = f_display["Ejercicio"]
                    f_final["_Series_Original"] = f_display["Series"]
                    f_final["_Reps_Original"] = f_display["Reps"]
                    
                    ed_f_alumno = st.data_editor(
                        f_final,
                        column_order=["Ejercicio_Full", "SxR", "Kg", "Notas"], # Solo mostramos estas 4
                        column_config={
                            "Ejercicio_Full": st.column_config.TextColumn("Ejercicio (Bloque)", disabled=True),
                            "SxR": st.column_config.TextColumn("Series x Reps", disabled=True),
                            "Kg": st.column_config.NumberColumn("Kg (Real)", format="%.1f"),
                            "Notas": st.column_config.TextColumn("Mis Notas")
                        },
                        hide_index=True,
                        use_container_width=True,
                        key=f"fue_alu_{d_hoy}"
                    )
                else: ed_f_alumno = pd.DataFrame()

                # --- CARDIO (Editable) ---
                ca = r_hoy[r_hoy["Seccion"] == "Cardio"]
                if not ca.empty: 
                    st.markdown("### 🏃‍♂️ Cardio")
                    ca_display = ca.copy()
                    ed_ca_alumno = st.data_editor(
                        ca_display[["Ejercicio", "Series", "Reps", "Notas"]],
                        column_config={
                            "Series": st.column_config.TextColumn("Tiempo/Dist", disabled=True),
                            "Reps": st.column_config.TextColumn("Intensidad", disabled=True),
                            "Ejercicio": st.column_config.TextColumn("Ejercicio", disabled=True)
                        },
                        hide_index=True,
                        use_container_width=True,
                        key=f"car_alu_{d_hoy}"
                    )
                else: ed_ca_alumno = pd.DataFrame()
                
                # --- BOTÓN GUARDAR CAMBIOS ---
                st.markdown("<br>", unsafe_allow_html=True)
                if st.button("💾 GUARDAR MIS NOTAS Y PESOS", type="primary", use_container_width=True):
                    # Reconstruimos los dataframes originales para guardar
                    
                    # 1. Fuerza: Recuperamos columnas ocultas
                    df_f_save = pd.DataFrame()
                    if not ed_f_alumno.empty:
                        df_f_save = ed_f_alumno.copy()
                        # Mapeamos de vuelta a lo que espera la base de datos
                        df_f_save["Orden"] = df_f_save["_Orden_Original"]
                        df_f_save["Ejercicio"] = df_f_save["_Ejercicio_Original"]
                        df_f_save["Series"] = df_f_save["_Series_Original"]
                        df_f_save["Reps"] = df_f_save["_Reps_Original"]
                        # Kg y Notas ya están editados
                    
                    # 2. Calentamiento y Cardio (directos)
                    guardar_rutina_actualizada(alias, d_hoy, ed_c_alumno, df_f_save, ed_ca_alumno)
                    st.success("✅ ¡Tus anotaciones se han guardado en la rutina!")
                    st.rerun()

                # --- REGISTRO EXTRA ---
                st.markdown("---")
                with st.expander("➕ AGREGAR REGISTRO EXTRA (Opcional)"):
                    with st.form("reg"):
                        lista_ej = f["Ejercicio"].unique() if 'f' in locals() and not f.empty else ["Varios"]
                        ej = st.selectbox("Ejercicio", lista_ej)
                        c1, c2 = st.columns(2)
                        k = c1.number_input("Kilos", step=1.0)
                        reps = c2.number_input("Reps", step=1, min_value=1)
                        rpe = st.slider("RPE", 1, 10)
                        n = st.text_area("Notas")
                        if st.form_submit_button("GUARDAR"):
                            guardar_registro(alias, ej, k, reps, rpe, n)
                            st.success("Listo")

                c1, c2 = st.columns(2)
                with c1: 
                    if st.button("✅ LISTO", type="primary", use_container_width=True):
                        guardar_estado_sesion(alias, "Completado")
                        st.balloons()
                with c2: 
                    if st.button("⚠️ INCOMPLETO", use_container_width=True):
                        guardar_estado_sesion(alias, "Incompleto")
            else:
                st.info("No tienes rutina cargada.")
        
        with t2:
            dfs = leer_sesiones_alumno(alias)
            if not dfs.empty:
                st.metric("Entrenamientos", len(dfs))
                ch = alt.Chart(dfs).mark_rect().encode(x='monthdate(Fecha):O', y='month(Fecha):O', color=alt.Color('Estado', scale=alt.Scale(range=['#2ECC71', '#F39C12']))).properties(height=250)
                st.altair_chart(ch, use_container_width=True)
            
            st.markdown("---")
            st.markdown("### 📈 Mis Cargas y Reps")
            df_r = leer_registros_alumno(alias)
            if not df_r.empty and "Peso" in df_r.columns:
                 lista_ej = df_r["Ejercicio"].unique()
                 if len(lista_ej) > 0:
                     ej_sel = st.selectbox("Ver progreso en:", lista_ej)
                     df_plt = df_r[df_r["Ejercicio"] == ej_sel].sort_values("Fecha")
                     
                     st.caption("Evolución del Peso (Kg)")
                     st.line_chart(df_plt.set_index("Fecha")["Peso"], color="#E63946")
                     
                     if "Repeticiones" in df_plt.columns:
                         st.caption("Evolución de Repeticiones")
                         st.line_chart(df_plt.set_index("Fecha")["Repeticiones"], color="#ffffff")
                 else: st.write("Datos insuficientes.")
            else: st.write("Registra tus series para ver gráficos.")
