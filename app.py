import streamlit as st
from tarificador_core import ejecutar_tarificador
import tempfile
import os

st.title("Tarificador SiceTAC")

# Configuración de archivos maestros
maestros = {
    "2": "Maestro_SiceTAC_RNDC.xlsx",
    "2L1": "Maestro_SiceTAC_RNDC 2L1.xlsx",
    "3": "Maestro_SiceTAC_RNDC 3.xlsx",
    "2S2": "Maestro_SiceTAC_RNDC 2S2.xlsx",
    "3S2": "Maestro_SiceTAC_RNDC 3S2.xlsx"
}

# Formulario
tipo_vehiculo = st.selectbox("Tipo de Vehículo", list(maestros.keys()))
tipo_carga = st.selectbox("Tipo de Carga", ["2", "5", "12", "13", "1003"])
unidad_transporte = st.selectbox("Unidad de Transporte", [
    "ESTACAS", "FURGON", "ESTIBAS", "PLATAFORMA", "PORTACONTENEDORES", "TANQUE", "FURGON REFRIGERADO"
])

archivo = st.file_uploader("Sube tu archivo Excel tarifario", type=["xlsx"])

if st.button("Ejecutar Tarificador") and archivo:
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(archivo.read())
        tmp_path = tmp.name

    try:
        resultados = ejecutar_tarificador(tipo_vehiculo, tipo_carga, unidad_transporte, tmp_path, maestros)
        st.success("Tarifas calculadas correctamente.")

        with open(tmp_path, "rb") as f:
            st.download_button("Descargar tarifario modificado", f.read(), file_name="Tarifario_Modificado.xlsx")

    except Exception as e:
        st.error(f"Error: {e}")

    os.remove(tmp_path)
