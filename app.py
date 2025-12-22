import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials
import os # <--- ¡AGREGA ESTO!

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="El Estudio", page_icon="💪", layout="centered")

# --- CONEXIÓN GOOGLE SHEETS ---
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

def conectar_google_sheet():
    # LÓGICA INTELIGENTE:
    # 1. ¿Existe el archivo en mi carpeta? (Estamos en tu PC)
    if os.path.exists("mis_secretos.json"):
        creds = Credentials.from_service_account_file("mis_secretos.json", scopes=SCOPES)
    
    # 2. Si no existe, asumimos que estamos en Internet (Streamlit Cloud)
    else:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    
    cliente = gspread.authorize(creds)
    return cliente.open("El Estudio DB")

# --- FUNCIONES DE BASE DE DATOS ---

def obtener_usuario(usuario_input, password_input):
    """Verifica si el usuario y contraseña existen en la hoja 'Usuarios'"""
    try:
        sh = conectar_google_sheet()
        worksheet = sh.worksheet("Usuarios")
        datos = worksheet.get_all_records()
        df = pd.DataFrame(datos)
        
        # Filtramos buscando usuario y contraseña
        usuario_encontrado = df[
            (df["Usuario"].astype(str) == usuario_input) & 
            (df["Password"].astype(str) == password_input)
        ]
        
        if not usuario_encontrado.empty:
            return usuario_encontrado.iloc[0] # Retorna los datos del usuario
        return None
    except Exception as e:
        st.error(f"Error leyendo usuarios: {e}")
        return None

def guardar_rutina(alumno, dia, ejercicio, series, reps, notas):
    sh = conectar_google_sheet()
    worksheet = sh.worksheet("Rutinas")
    worksheet.append_row([alumno, dia, ejercicio, series, reps, notas])

def leer_rutina(alumno):
    sh = conectar_google_sheet()
    worksheet = sh.worksheet("Rutinas")
    datos = worksheet.get_all_records()
    df = pd.DataFrame(datos)
    if df.empty: return df
    # Filtramos solo las rutinas de este alumno
    return df[df["Alumno"] == alumno]

def guardar_registro(usuario, ejercicio, peso, rpe, notas):
    sh = conectar_google_sheet()
    worksheet = sh.worksheet("Registros") # OJO: Ahora se llama Registros
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    worksheet.append_row([fecha, usuario, ejercicio, peso, rpe, notas])

def leer_registros():
    sh = conectar_google_sheet()
    worksheet = sh.worksheet("Registros")
    return pd.DataFrame(worksheet.get_all_records())

# --- GESTIÓN DE SESIÓN (LOGIN) ---
if 'logueado' not in st.session_state:
    st.session_state['logueado'] = False
if 'usuario_info' not in st.session_state:
    st.session_state['usuario_info'] = None

# --- PANTALLA DE LOGIN ---
if not st.session_state['logueado']:
    st.title("🔐 El Estudio - Acceso")
    with st.form("login_form"):
        user = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")
        submit = st.form_submit_button("Ingresar")
        
        if submit:
            usuario_validado = obtener_usuario(user, password)
            if usuario_validado is not None:
                st.session_state['logueado'] = True
                st.session_state['usuario_info'] = usuario_validado
                st.success("¡Bienvenido!")
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos")

else:
    # --- APLICACIÓN PRINCIPAL (YA LOGUEADO) ---
    datos_usuario = st.session_state['usuario_info']
    rol = datos_usuario['Rol']
    nombre = datos_usuario['Nombre']
    alias = datos_usuario['Usuario']

    # Barra lateral con botón de salir
    with st.sidebar:
        st.write(f"Hola, **{nombre}**")
        st.write(f"Rol: {rol}")
        if st.button("Cerrar Sesión"):
            st.session_state['logueado'] = False
            st.session_state['usuario_info'] = None
            st.rerun()

    # --- VISTA DE ENTRENADOR (ADMIN) ---
    if rol == "admin":
        st.title("👨‍🏫 Panel de Entrenador")
        
        tab1, tab2 = st.tabs(["📝 Crear Rutinas", "📊 Ver Progresos"])
        
        with tab1:
            st.subheader("Asignar nueva rutina")
            # Buscamos lista de alumnos (truco: leemos la hoja usuarios)
            sh = conectar_google_sheet()
            df_users = pd.DataFrame(sh.worksheet("Usuarios").get_all_records())
            lista_alumnos = df_users[df_users["Rol"] == "alumno"]["Usuario"].tolist()
            
            with st.form("form_asignar"):
                alumno_elegido = st.selectbox("Selecciona Alumno", lista_alumnos)
                dia = st.selectbox("Día", ["Lunes", "Martes", "Miércoles", "Jueves", "Viernes", "Sábado"])
                ejercicio = st.text_input("Ejercicio (Ej: Sentadilla)")
                col1, col2 = st.columns(2)
                series = col1.text_input("Series (Ej: 3)")
                reps = col2.text_input("Reps (Ej: 10-12)")
                notas_rutina = st.text_area("Instrucciones")
                
                if st.form_submit_button("Guardar en Rutina"):
                    guardar_rutina(alumno_elegido, dia, ejercicio, series, reps, notas_rutina)
                    st.success(f"Ejercicio asignado a {alumno_elegido}")

            st.write("---")
            st.write("### Rutinas vigentes:")
            st.dataframe(pd.DataFrame(sh.worksheet("Rutinas").get_all_records()))

        with tab2:
            st.write("Historial de entrenamientos realizados:")
            st.dataframe(leer_registros())

    # --- VISTA DE ALUMNO ---
    else:
        st.title(f"🚀 Vamos a entrenar, {nombre}")
        
        # 1. MOSTRAR RUTINA ASIGNADA
        st.info("👇 Esta es la rutina que te asignó tu coach:")
        mi_rutina = leer_rutina(alias)
        
        if not mi_rutina.empty:
            # Mostramos la rutina agrupada por días o simple
            st.table(mi_rutina[["Dia", "Ejercicio", "Series", "Reps", "Notas"]])
        else:
            st.warning("Aún no tienes rutina asignada. ¡Escribe al coach!")
        
        st.write("---")
        
        # 2. REGISTRAR EL ENTRENAMIENTO (Lo que ya tenías)
        st.subheader("✍️ Registrar Resultados")
        with st.form("form_alumno"):
            ejercicio_hecho = st.text_input("Ejercicio Realizado")
            c1, c2 = st.columns(2)
            peso = c1.number_input("Peso (kg)", step=1.0)
            rpe = c2.slider("RPE", 1, 10)
            notas = st.text_area("Notas / Sensaciones")
            
            if st.form_submit_button("Guardar Serie"):
                guardar_registro(alias, ejercicio_hecho, peso, rpe, notas)
                st.success("¡Guardado!")
