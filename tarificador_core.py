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
    return texto.count("periferia") if "periferia" in texto else 0

def leer_tarifario(hoja_tarifario):
    datos = []
    for fila in hoja_tarifario.iter_rows(min_row=4, max_row=hoja_tarifario.max_row, values_only=True):
        origen, destino = fila[2], fila[3]
        if origen and destino:
            datos.append((origen, destino))
    return datos

def buscar_en_maestro(hoja_maestro, datos, tipo_carga, unidad_transporte):
    resultados = []
    for origen_tarifario, destino_tarifario in datos:
        ot_norm = normalize_text(origen_tarifario).replace("-", " ")
        dt_norm = normalize_text(destino_tarifario)

        tokens_origen = ot_norm.split()
        n_periferias = extraer_periferias(dt_norm)
        destino_tratado = "urbano" if n_periferias > 0 else dt_norm

        encontrado = False

        for fila in range(2, hoja_maestro.max_row + 1):
            mes = hoja_maestro[f'H{fila}'].value
            tipo = normalize_text(hoja_maestro[f'K{fila}'].value)

            if mes != int(tipo_carga) or tipo != normalize_text(unidad_transporte):
                continue

            origen_maestro = normalize_text(hoja_maestro[f'D{fila}'].value)
            destino_maestro = normalize_text(hoja_maestro[f'F{fila}'].value)

            origen_match = any(m in origen_maestro for m in tokens_origen)

            destino_match = False
            if destino_tratado == "urbano":
                destino_match = any(m in destino_maestro for m in tokens_origen)
            else:
                tokens_dest = dt_norm.split()
                destino_final = tokens_dest[-1] if tokens_dest else ""
                destino_match = destino_final in destino_maestro

            if origen_match and destino_match:
                valor_base = hoja_maestro[f'N{fila}'].value or 0
                adicional = hoja_maestro[f'O{fila}'].value or 0
                valor_total = valor_base * (n_periferias if n_periferias > 0 else 1) + (adicional * 8)

                resultados.append((fila, origen_tarifario, destino_tarifario, valor_total))
                encontrado = True
                break

        if not encontrado:
            resultados.append((None, origen_tarifario, destino_tarifario, "No encontrado"))

    return resultados

def ejecutar_tarificador(tipo_vehiculo, tipo_carga, unidad_transporte, archivo_tarifario, maestros):
    ruta_maestro = maestros.get(tipo_vehiculo)
    if not ruta_maestro or not os.path.exists(ruta_maestro):
        raise ValueError("Tipo de vehículo no válido o archivo no encontrado.")

    libro_tarifario = openpyxl.load_workbook(archivo_tarifario)
    hoja_tarifario = libro_tarifario['Tarifario']
    libro_maestro = openpyxl.load_workbook(ruta_maestro)
    hoja_maestro = libro_maestro.active

    datos_tarifario = leer_tarifario(hoja_tarifario)
    resultados = buscar_en_maestro(hoja_maestro, datos_tarifario, tipo_carga, unidad_transporte)

    hoja_tarifario["B2"] = f"Tarifa para vehículo {tipo_vehiculo}"

    for i, (_, _, _, tarifa) in enumerate(resultados, start=4):
        if i > 80:
            break
        hoja_tarifario[f'D{i}'] = tarifa

    libro_tarifario.save(archivo_tarifario)
    return resultados
