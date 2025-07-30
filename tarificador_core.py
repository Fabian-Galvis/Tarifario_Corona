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
    for fila in range(4, hoja_tarifario.max_row + 1):
        origen = hoja_tarifario[f'B{fila}'].value
        destino = hoja_tarifario[f'C{fila}'].value
        if origen and destino:
            datos.append((fila, origen, destino))
    return datos

def obtener_candidatos(hoja_ubicaciones, texto):
    texto_norm = normalize_text(texto)
    candidatos = []
    for fila in range(2, hoja_ubicaciones.max_row + 1):
        cod_dep = str(hoja_ubicaciones[f'A{fila}'].value).strip()
        nom_dep = normalize_text(hoja_ubicaciones[f'B{fila}'].value)
        cod_mun = str(hoja_ubicaciones[f'C{fila}'].value).strip()
        nom_mun = normalize_text(hoja_ubicaciones[f'D{fila}'].value)

        if nom_dep in texto_norm or nom_mun in texto_norm:
            candidatos.append((fila, cod_mun + '000', hoja_ubicaciones[f'B{fila}'].value, hoja_ubicaciones[f'D{fila}'].value))

    return candidatos

def buscar_en_maestro_con_ubicaciones(hoja_maestro, hoja_ubicaciones, datos, tipo_carga, unidad_transporte, horas_logisticas, hoja_tarifario):

    try:
        tipo_carga = int(tipo_carga)
    except ValueError:
        raise ValueError("El tipo de carga debe ser un número entero (mes).")

    try:
        horas_logisticas = int(horas_logisticas)
    except ValueError:
        raise ValueError("Tipo de carga y horas logísticas deben ser números.")

    offset = 0
    for fila_tarifario, origen_tarifario, destino_tarifario in datos:
        fila_tarifario += offset
        n_periferias = extraer_periferias(destino_tarifario)
        destino_tratado = "urbano" if n_periferias > 0 else destino_tarifario

        origenes = obtener_candidatos(hoja_ubicaciones, origen_tarifario)
        destinos = obtener_candidatos(hoja_ubicaciones, destino_tratado)

        if not origenes or not destinos:
            hoja_tarifario[f'E{fila_tarifario}'] = "No encontrado"
            continue

        total_iteraciones = len(origenes) * len(destinos)
        iteracion = 0

        for cod_ori, cod_ori_str, dep_ori, mun_ori in origenes:
            for cod_dest, cod_dest_str, dep_dest, mun_dest in destinos:
                iteracion += 1
                encontrado = False
                for fila in range(2, hoja_maestro.max_row + 1):
                    mes = hoja_maestro[f'H{fila}'].value
                    tipo = normalize_text(hoja_maestro[f'K{fila}'].value)

                    if mes != tipo_carga or tipo != normalize_text(unidad_transporte):
                        continue

                    cod_ori_maestro = str(hoja_maestro[f'C{fila}'].value).strip()
                    cod_dest_maestro = str(hoja_maestro[f'E{fila}'].value).strip()

                    if cod_ori_str == cod_ori_maestro and cod_dest_str == cod_dest_maestro:
                        valor_base = hoja_maestro[f'N{fila}'].value or 0
                        adicional = hoja_maestro[f'O{fila}'].value or 0
                        valor_total = (
                            valor_base * (n_periferias if n_periferias > 0 else 1)
                            + (adicional * horas_logisticas)
                        )

                        if total_iteraciones == 1:
                            hoja_tarifario[f'B{fila_tarifario}'] = f"{origen_tarifario} ({dep_ori})"
                            hoja_tarifario[f'C{fila_tarifario}'] = f"{destino_tarifario} ({dep_dest})"
                            hoja_tarifario[f'E{fila_tarifario}'] = valor_total
                        else:
                            hoja_tarifario.insert_rows(fila_tarifario + 1)
                            hoja_tarifario[f'B{fila_tarifario + 1}'] = f"{origen_tarifario} ({dep_ori})"
                            hoja_tarifario[f'C{fila_tarifario + 1}'] = f"{destino_tarifario} ({dep_dest})"
                            hoja_tarifario[f'E{fila_tarifario + 1}'] = valor_total
                            offset += 1
                        encontrado = True
                        break

                if not encontrado and total_iteraciones == 1:
                    hoja_tarifario[f'E{fila_tarifario}'] = "No encontrado"

    return True

def ejecutar_tarificador(tipo_vehiculo, tipo_carga, unidad_transporte, archivo_tarifario, maestros, horas_logisticas):
    ruta_maestro = maestros.get(tipo_vehiculo)
    if not ruta_maestro or not os.path.exists(ruta_maestro):
        raise FileNotFoundError("Tipo de vehículo no válido o archivo no encontrado.")

    libro_tarifario = openpyxl.load_workbook(archivo_tarifario)
    hoja_tarifario = libro_tarifario.active
    libro_maestro = openpyxl.load_workbook(ruta_maestro)
    hoja_maestro = libro_maestro.active
    libro_ubicaciones = openpyxl.load_workbook("Maestro_ubicaciones.xlsx")
    hoja_ubicaciones = libro_ubicaciones.active

    datos_tarifario = leer_tarifario(hoja_tarifario)
    buscar_en_maestro_con_ubicaciones(
        hoja_maestro,
        hoja_ubicaciones,
        datos_tarifario,
        tipo_carga,
        unidad_transporte,
        horas_logisticas,
        hoja_tarifario
    )

    hoja_tarifario['B2'] = f"VEHICULO: {tipo_vehiculo}"
    hoja_tarifario['E2'] = f"HORAS LOGISTICAS: {horas_logisticas}"

    libro_tarifario.save(archivo_tarifario)
    return True
