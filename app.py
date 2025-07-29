import streamlit as st
from tarificador_core import ejecutar_tarificador
import tempfile
import os
import pandas as pd
import base64

st.title("Tarificador")

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
col1, col2 = st.columns([1, 1])
with col1:
    with open("Plantilla Tarifario.xlsx", "rb") as plantilla_file:
        st.download_button("📥 Descargar plantilla admitida", plantilla_file, file_name="Plantilla Tarifario.xlsx")

with col2:
    mostrar_info = st.toggle("📋 Ver antes de subir")

# Si se activa el toggle, se muestra el contenido emergente
if mostrar_info:
    st.markdown("""
    <div style='background-color:#fff3cd; padding: 15px; border-radius: 5px; border: 1px solid #ffeeba; margin-top:10px;'>
      <strong>⚠️ Antes de subir:</strong><br>
      - Verificar la plantilla admitida antes de subir el archivo<br>
      - Ingresar el <strong>origen</strong> y el <strong>destino</strong> en sus respectivas columnas:
    </div>
    """, unsafe_allow_html=True)

    # Imagen informativa
    st.markdown("""
    <div style='text-align: center; margin-top: 10px;'>
      <img src='data:image/png;base64,{}' style='width: 25%; border: 1px solid #ccc; border-radius: 4px;'><br>
      <small>Ejemplo correcto de encabezados en la plantilla</small>
    </div>
    """.format(base64.b64encode(open("info.jpg", "rb").read()).decode()), unsafe_allow_html=True)

    # Aviso tipo de archivo
    st.markdown("""
    <div style='background-color:#fff3cd; padding: 10px; border-radius: 5px; border: 1px solid #ffeeba; margin-top:10px;'>
    - Solo subir archivos de Excel <code>.xlsx</code>
    </div>
    """, unsafe_allow_html=True)

# Carga de archivo
archivo = st.file_uploader("Sube tu archivo Excel tarifario", type=["xlsx"])

# Botón de ejecución
if st.button("Ejecutar Tarificador"):
    if archivo is None:
        st.warning("Por favor sube un archivo .xlsx antes de ejecutar.")
    elif not archivo.name.endswith(".xlsx"):
        st.error("Formato inválido. Solo se permiten archivos con extensión .xlsx.")
    else:
        try:
            df = pd.read_excel(archivo, header=None)
            b3 = str(df.iloc[2, 1]).strip().lower()
            c3 = str(df.iloc[2, 2]).strip().lower()

            errores = []
            if b3 != "origen":
                errores.append("celda B3 debe decir 'origen'")
            if c3 != "destino":
                errores.append("celda C3 debe decir 'destino'")

            if errores:
                st.error("❌ Errores en la plantilla:\n- " + "\n- ".join(errores))
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
