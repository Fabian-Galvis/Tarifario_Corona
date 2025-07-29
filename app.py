import streamlit as st
from tarificador_core import ejecutar_tarificador
import tempfile
import os
import pandas as pd

st.title("Tarificador SiceTAC")

# Diccionarios
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

# Selectores
tipo_vehiculo = st.selectbox("Tipo de Vehículo", list(maestros.keys()))
nombre_carga = st.selectbox("Tipo de Carga", list(t_carga.keys()))
tipo_carga = t_carga[nombre_carga]

unidad_transporte = st.selectbox("Unidad de Transporte", [
    "ESTACAS", "ESTIBAS", "TANQUE", "FURGON", "PORTACONTENEDORES", 
    "PLATAFORMA", "TRAYLER", "VOLCO", "FURGON REFRIGERADO"
])

# Botón para descargar plantilla
with open("Plantilla Tarifario.xlsx", "rb") as plantilla_file:
    st.download_button("📥 Descargar plantilla admitida", plantilla_file, file_name="Plantilla Tarifario.xlsx")

# Instrucciones
st.markdown("""
<div style="background-color:#fff3cd; padding:10px; border-left:5px solid #ffc107; margin-bottom:15px;">
    <strong>⚠️ Antes de subir:</strong><br>
    - Verificar la plantilla admitida antes de subir el archivo<br>
    - Ingresar el <b>origen</b> y el <b>destino</b> en sus respectivas columnas:<br>
</div>
""", unsafe_allow_html=True)

st.image("info.jpg", caption="Ejemplo correcto de encabezados en la plantilla", use_column_width=True)

st.markdown("""
<div style="background-color:#fff3cd; padding:10px; border-left:5px solid #ffc107; margin-top:-10px; margin-bottom:15px;">
    - Solo subir archivos de Excel <code>.xlsx</code>
</div>
""", unsafe_allow_html=True)

# Carga de archivo
archivo = st.file_uploader("Sube tu archivo Excel tarifario", type=["xlsx"])

if archivo is not None:
    # Mostrar vista previa
    try:
        df_preview = pd.read_excel(archivo)
        st.subheader("📄 Vista previa del archivo")
        st.dataframe(df_preview.head())
    except Exception as e:
        st.error(f"Error al leer el archivo: {e}")

# Botón de ejecución
if st.button("Ejecutar Tarificador"):
    if archivo is None:
        st.warning("Por favor sube un archivo .xlsx antes de ejecutar.")
    elif not archivo.name.endswith(".xlsx"):
        st.error("Formato inválido. Solo se permiten archivos con extensión .xlsx.")
    else:
        try:
            df = pd.read_excel(archivo)
            columnas = [col.strip().lower() for col in df.columns]

            if "origen" not in columnas or "destino" not in columnas:
                faltantes = []
                if "origen" not in columnas:
                    faltantes.append("Origen")
                if "destino" not in columnas:
                    faltantes.append("Destino")
                st.error(f"❌ El archivo no contiene las columnas requeridas: {', '.join(faltantes)}.")
            else:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    archivo.seek(0)
                    tmp.write(archivo.read())
                    tmp_path = tmp.name

                with st.spinner("⏳ Generando tarifas, por favor espera..."):
                    resultados = ejecutar_tarificador(tipo_vehiculo, tipo_carga, unidad_transporte, tmp_path, maestros)

                st.success("✅ Tarifas calculadas correctamente.")

                with open(tmp_path, "rb") as f:
                    st.download_button("⬇️ Descargar tarifario modificado", f.read(), file_name="Tarifario_Modificado.xlsx")

                os.remove(tmp_path)

        except Exception as e:
            st.error(f"❌ Error durante el procesamiento: {e}")
