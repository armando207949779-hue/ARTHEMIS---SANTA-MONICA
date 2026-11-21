import importlib
import streamlit as st

st.set_page_config(
    page_title="Panel Principal",
    layout="wide"
)

# Menú en la barra lateral
st.sidebar.title("Navegación")
opcion = st.sidebar.radio(
    "Selecciona una app:",
    [
        "Pestaña 1: Generador de Guías",
        "Pestaña 2",
        "Pestaña 3",
        "Pestaña 4",
        "Pestaña 5",
        "Pestaña 6",
        "Pestaña 7",
        "Pestaña 8",
        "Pestaña 9",
        "Pestaña 10"
    ]
)

# Contenido principal según la selección
if opcion == "Pestaña 1: Generador de Guías":
    app_01 = importlib.import_module("01_APP")
    app_01.run()

elif opcion == "Pestaña 2":
    st.title("Pestaña 2")
    st.info("Espacio disponible para conectar app 2.")

elif opcion == "Pestaña 3":
    st.title("Pestaña 3")
    st.info("Espacio disponible para conectar app 3.")

elif opcion == "Pestaña 4":
    st.title("Pestaña 4")
    st.info("Espacio disponible para conectar app 4.")

elif opcion == "Pestaña 5":
    st.title("Pestaña 5")
    st.info("Espacio disponible para conectar app 5.")

elif opcion == "Pestaña 6":
    st.title("Pestaña 6")
    st.info("Espacio disponible para conectar app 6.")

elif opcion == "Pestaña 7":
    st.title("Pestaña 7")
    st.info("Espacio disponible para conectar app 7.")

elif opcion == "Pestaña 8":
    st.title("Pestaña 8")
    st.info("Espacio disponible para conectar app 8.")

elif opcion == "Pestaña 9":
    st.title("Pestaña 9")
    st.info("Espacio disponible para conectar app 9.")

elif opcion == "Pestaña 10":
    st.title("Pestaña 10")
    st.info("Espacio disponible para conectar app 10.")
