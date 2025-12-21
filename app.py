import streamlit as st
import pandas as pd
from datetime import datetime
import gspread
from google.oauth2.service_account import Credentials

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="El Estudio", page_icon="💪", layout="centered")

# --- CONEXIÓN CON GOOGLE SHEETS ---
# Definimos los permisos que necesita el robot
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive"
]

# Asegúrate de importar esto arriba del todo:
# import json (Agrégalo junto a los otros imports)

def conectar_google_sheet():
    # Opción A: Estamos en la Nube (Streamlit Cloud)
    if "gcp_service_account" in st.secrets:
        creds_dict = st.secrets["gcp_service_account"]
        creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    
    # Opción B: Estamos en local (Tu computadora)
    else:
        creds = Credentials.from_service_account_file("mis_secretos.json", scopes=SCOPES)
        
    cliente = gspread.authorize(creds)
    sheet = cliente.open("El Estudio DB").sheet1
    return sheet

# --- FUNCIONES NUEVAS (El Cerebro en la Nube) ---
def cargar_datos():
    try:
        sheet = conectar_google_sheet()
        # Bajamos todos los datos de la hoja
        datos = sheet.get_all_records()
        df = pd.DataFrame(datos)
        
        # Si la hoja está vacía (solo encabezados), pandas a veces se lía, así que aseguramos
        if df.empty:
            return pd.DataFrame(columns=["Fecha", "Usuario", "Ejercicio", "Peso", "RPE", "Notas"])
            
        return df
    except Exception as e:
        st.error(f"Error al conectar con Google Sheets: {e}")
        return pd.DataFrame()

def guardar_datos(usuario, ejercicio, peso, rpe, notas):
    sheet = conectar_google_sheet()
    fecha = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    # Agregamos una fila al final de la hoja (append_row)
    sheet.append_row([fecha, usuario, ejercicio, peso, rpe, notas])

# Lista de ejercicios
LISTA_EJERCICIOS = [
    "Sentadilla Hack", "Press de Banca", "Peso Muerto Rumano", 
    "Hip Thrust", "Press Militar", "Jalón al Pecho"
]

# --- INTERFAZ (Esto casi no cambia) ---
st.title("🏋️‍♂️ El Estudio - Cloud Edition ☁️")

usuario_actual = st.sidebar.text_input("👤 Tu ID de Alumno:")

if not usuario_actual:
    st.info("👈 Identifícate para entrar.")
    st.stop()

# --- VISTA ADMIN ---
if usuario_actual.lower() in ["admin", "coach"]:
    st.header("Base de Datos en Vivo (Google Sheets)")
    
    # Botón para forzar actualización
    if st.button("🔄 Refrescar Datos"):
        st.rerun()
        
    df = cargar_datos()
    st.dataframe(df)

# --- VISTA ALUMNO ---
else:
    st.subheader(f"Hola, {usuario_actual}")
    
    ejercicio_seleccionado = st.selectbox("Ejercicio", LISTA_EJERCICIOS)

    # Gráficos
    df = cargar_datos()
    if not df.empty and "Usuario" in df.columns:
        datos_ejercicio = df[
            (df["Usuario"] == usuario_actual) & 
            (df["Ejercicio"] == ejercicio_seleccionado)
        ]
        
        if not datos_ejercicio.empty:
            st.line_chart(datos_ejercicio.set_index("Fecha")["Peso"])
            record = datos_ejercicio["Peso"].max()
            st.metric("Récord Personal", f"{record} kg")

    # Formulario
    with st.form("form_cloud"):
        c1, c2 = st.columns(2)
        peso = c1.number_input("Peso (kg)", step=1.0)
        rpe = c2.slider("RPE", 1, 10)
        notas = st.text_area("Notas")
        
        if st.form_submit_button("Guardar en la Nube 🚀"):
            with st.spinner("Enviando a Google..."):
                guardar_datos(usuario_actual, ejercicio_seleccionado, peso, rpe, notas)
            st.success("¡Guardado! Revisa tu Google Sheet.")
            st.rerun()