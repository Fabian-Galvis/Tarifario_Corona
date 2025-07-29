import streamlit as st
from tarificador_core import ejecutar_tarificador
import tempfile
import os

st.title("Tarificador SiceTAC")

# Configuración de archivos maestros
maestros = {
    "Camión dos ejes - PBV mas de 10500 Kg (2)": "Maestro_SiceTAC_RNDC 2.xlsx",
    "Camión dos ejes - Livianos PBV 7500 - 8000 Kg (2L3)": "Maestro_SiceTAC_RNDC 2L3.xlsx",
    "Camión dos ejes - Livianos PBV 8001 - 9000 Kg (2L2)": "Maestro_SiceTAC_RNDC 2L2.xlsx",
    "Camión dos ejes - Livianos PBV 9001 - 10500 Kg (2L1)": "Maestro_SiceTAC_RNDC 2L1.xlsx",
    "Tractocamión dos ejes con semiremolque de dos ejes (2S2)": "Maestro_SiceTAC_RNDC 2S2.xlsx",
    "Tractocamión dos ejes con semiremolque de tres ejes (2S3)": "Maestro_SiceTAC_RNDC 2S3.xlsx",
    "Camión 3 ejes": "Maestro_SiceTAC_RNDC 3.xlsx",
    "Tractocamión tres ejes con semiremolque de dos ejes (3S2)": "Maestro_SiceTAC_RNDC 3S2.xlsx",
    "Tractocamión tres ejes con semiremolque de tres ejes (3S3)": "Maestro_SiceTAC_RNDC 3S3.xlsx",
    "Volqueta dos ejes (V2)": "Maestro_SiceTAC_RNDC V2.xlsx",
    "Volqueta tres ejes (V3)": "Maestro_SiceTAC_RNDC V3.xlsx"
}

t_carga = {
    "General": 12,
    "Granel sólido": 5,
    "Granel líquido": 1003,
    "Contenedor": 13,
    "Carga refrigerada": 2
}

# Formulario
tipo_vehiculo = st.selectbox("Tipo de Vehículo", list(maestros.keys()))
tipo_carga = st.selectbox("Tipo de Carga", list(t_carga.keys()))
unidad_transporte = st.selectbox("Unidad de Transporte", [
    "ESTACAS", "ESTIBAS", "TANQUE", "FURGON", "PORTACONTENEDORES", "PLATAFORMA", "TRAYLER", "VOLCO", "FURGON REFRIGERADO"
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
