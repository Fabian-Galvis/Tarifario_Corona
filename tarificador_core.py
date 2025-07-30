
import openpyxl
import unicodedata
import re
import os

def normalize_text(texto):
    if texto is None:
        return ""
    texto = str(texto).lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = re.sub(r'[\u0300-\u036f]', '', texto)
    texto = re.sub(r'[^a-z0-9\s-]', '', texto)
    return texto.strip()

def extraer_periferias(texto):
    if texto is None:
        return 0
    texto = texto.lower()
    return texto.count("periferia")

def leer_tarifario(hoja_tarifario):
    datos = []
    for fila in range(4, hoja_tarifario.max_row + 1):  # Desde fila 4
        origen = hoja_tarifario[f'B{fila}'].value
        destino = hoja_tarifario[f'C{fila}'].value
        if origen and destino:
            datos.append((fila, origen, destino))
    return datos

def buscar_en_maestro(hoja_maestro, datos, tipo_carga, unidad_transporte, horas_logisticas):
    resultados = []

    try:
        tipo_carga = int(tipo_carga)
    except ValueError:
        raise ValueError("El tipo de carga debe ser un número entero (mes).")

    try:
        horas_logisticas = int(horas_logisticas)
    except ValueError:
        raise ValueError("Las horas logísticas deben ser un número.")

    for fila_tarifario, origen_tarifario, destino_tarifario in datos:
        ot_norm = normalize_text(origen_tarifario).replace("-", " ")
        dt_norm = normalize_text(destino_tarifario)

        tokens_origen = ot_norm.split()
        n_periferias = extraer_periferias(dt_norm)
        destino_tratado = "urbano" if n_periferias > 0 else dt_norm

        encontrado = False

        for fila in range(2, hoja_maestro.max_row + 1):
            mes = hoja_maestro[f'H{fila}'].value
            tipo = normalize_text(hoja_maestro[f'K{fila}'].value)

            if mes != tipo_carga or tipo != normalize_text(unidad_transporte):
                continue

            origen_maestro = normalize_text(hoja_maestro[f'D{fila}'].value)
            destino_maestro = normalize_text(hoja_maestro[f'F{fila}'].value)

            origen_match = any(m in origen_maestro for m in tokens_origen)

            if destino_tratado == "urbano":
                destino_match = any(m in destino_maestro for m in tokens_origen)
            else:
                tokens_dest = dt_norm.split()
                destino_final = tokens_dest[-1] if tokens_dest else ""
                destino_match = destino_final in destino_maestro

            if origen_match and destino_match:
                valor_base = hoja_maestro[f'N{fila}'].value or 0
                adicional = hoja_maestro[f'O{fila}'].value or 0
                valor_total = (
                    valor_base * (n_periferias if n_periferias > 0 else 1)
                    + (adicional * horas_logisticas)
                )

                resultados.append((fila_tarifario, valor_total))
                encontrado = True
                break

        if not encontrado:
            resultados.append((fila_tarifario, "No encontrado"))

    return resultados

def ejecutar_tarificador(tipo_vehiculo, tipo_carga, unidad_transporte, archivo_tarifario, maestros, horas_logisticas):
    ruta_maestro = maestros.get(tipo_vehiculo)
    if not ruta_maestro or not os.path.exists(ruta_maestro):
        raise FileNotFoundError("Tipo de vehículo no válido o archivo no encontrado.")

    libro_tarifario = openpyxl.load_workbook(archivo_tarifario)
    hoja_tarifario = libro_tarifario.active
    libro_maestro = openpyxl.load_workbook(ruta_maestro)
    hoja_maestro = libro_maestro.active

    datos_tarifario = leer_tarifario(hoja_tarifario)
    resultados = buscar_en_maestro(hoja_maestro, datos_tarifario, tipo_carga, unidad_transporte, horas_logisticas)

    hoja_tarifario['B2'] = f"VEHICULO: {tipo_vehiculo}"
    hoja_tarifario['E2'] = f"HORAS LOGISTICAS: {horas_logisticas}"

    for fila, tarifa in resultados:
        hoja_tarifario[f'E{fila}'] = tarifa

    libro_tarifario.save(archivo_tarifario)
    return resultados