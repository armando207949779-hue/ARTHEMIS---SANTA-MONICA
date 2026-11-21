import streamlit as st

st.set_page_config(
    page_title="Panel Principal",
    layout="wide"
)

st.title("Panel Principal")

# Definición de las 10 pestañas
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8, tab9, tab10 = st.tabs([
    "Pestaña 1",
    "Pestaña 2",
    "Pestaña 3",
    "Pestaña 4",
    "Pestaña 5",
    "Pestaña 6",
    "Pestaña 7",
    "Pestaña 8",
    "Pestaña 9",
    "Pestaña 10"
])

with tab1:
    st.subheader("Pestaña 1")
    st.info("Espacio disponible para conectar app 1.")

with tab2:
    st.subheader("Pestaña 2")
    st.info("Espacio disponible para conectar app 2.")

with tab3:
    st.subheader("Pestaña 3")
    st.info("Espacio disponible para conectar app 3.")

with tab4:
    st.subheader("Pestaña 4")
    st.info("Espacio disponible para conectar app 4.")

with tab5:
    st.subheader("Pestaña 5")
    st.info("Espacio disponible para conectar app 5.")

with tab6:
    st.subheader("Pestaña 6")
    st.info("Espacio disponible para conectar app 6.")

with tab7:
    st.subheader("Pestaña 7")
    st.info("Espacio disponible para conectar app 7.")

with tab8:
    st.subheader("Pestaña 8")
    st.info("Espacio disponible para conectar app 8.")

with tab9:
    st.subheader("Pestaña 9")
    st.info("Espacio disponible para conectar app 9.")

with tab10:
    st.subheader("Pestaña 10")
    st.info("Espacio disponible para conectar app 10.")
