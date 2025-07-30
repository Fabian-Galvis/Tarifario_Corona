import openpyxl
import unicodedata
import re
import os

# Constantes
DEPARTAMENTOS = {
    "amazonas", "antioquia", "arauca", "atlántico", "bogotá d.c", "bolívar", "boyacá", "caldas",
    "caquetá", "casanare", "cauca", "cesar", "chocó", "córdoba", "cundinamarca", "guainía",
    "guaviare", "huila", "la guajira", "magdalena", "meta", "nariño", "norte de santander",
    "putumayo", "quindío", "risaralda", "santander", "sucre", "tolima", "valle del cauca",
    "vaupés", "vichada"
}

def normalize_text(texto):
    if texto is None:
        return ""
    texto = str(texto).lower()
    texto = unicodedata.normalize("NFD", texto)
    texto = re.sub(r'[\u0300-\u036f]', '', texto)
    texto = re.sub(r'[^a-z0-9\s-]', '', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

def extraer_periferias(texto):
    if texto is None:
        return 0
    texto = normalize_text(texto)
    return texto.count("periferia")

def leer_tarifario(hoja_tarifario):
    datos = []
    for fila in range(4, hoja_tarifario.max_row + 1):  # Desde fila 4
        origen = hoja_tarifario[f'B{fila}'].value
        destino = hoja_tarifario[f'C{fila}'].value
        if isinstance(origen, str) and isinstance(destino, str):
            datos.append((fila, origen, destino))
    return datos

def obtener_candidatos(hoja_ubicaciones, texto):
    texto_norm = normalize_text(texto)
    candidatos = []

    for fila in range(2, hoja_ubicaciones.max_row + 1):
        depto = normalize_text(hoja_ubicaciones[f'B{fila}'].value)
        municipio = normalize_text(hoja_ubicaciones[f'D{fila}'].value)

        if depto in texto_norm or municipio in texto_norm:
            codigo_divipola = str(hoja_ubicaciones[f'C{fila}'].value) + '000'
            dep_original = hoja_ubicaciones[f'B{fila}'].value
            mun_original = hoja_ubicaciones[f'D{fila}'].value
            candidatos.append((fila, codigo_divipola, dep_original, mun_original))

    return candidatos

def buscar_en_maestro_con_ubicaciones(hoja_maestro, hoja_ubicaciones, datos, tipo_carga, unidad_transporte, horas_logisticas, hoja_tarifario):
    try:
        tipo_carga = int(tipo_carga)
    except ValueError:
        raise ValueError("El tipo de carga debe ser un número entero (mes).")

    try:
        horas_logisticas = int(horas_logisticas)
    except ValueError:
        raise ValueError("Las horas logísticas deben ser un número.")

    offset = 0  # Controla filas insertadas

    for fila_tarifario, origen_tarifario, destino_tarifario in datos:
        fila_tarifario += offset
        n_periferias = extraer_periferias(destino_tarifario)
        destino_tratado = "urbano" if n_periferias > 0 else destino_tarifario

        origenes = obtener_candidatos(hoja_ubicaciones, origen_tarifario)
        destinos = obtener_candidatos(hoja_ubicaciones, destino_tratado)

        if not origenes or not destinos:
            hoja_tarifario[f'E{fila_tarifario}'] = "No encontrado"
            continue

        es_primera_coincidencia = True

        for _, cod_ori_str, dep_ori, mun_ori in origenes:
            for _, cod_dest_str, dep_dest, mun_dest in destinos:
                encontrado = False

                for fila in range(2, hoja_maestro.max_row + 1):
                    mes = hoja_maestro[f'H{fila}'].value
                    tipo = normalize_text(hoja_maestro[f'K{fila}'].value)
                    cod_ori_maestro = str(hoja_maestro[f'C{fila}'].value or '').strip()
                    cod_dest_maestro = str(hoja_maestro[f'E{fila}'].value or '').strip()

                    if mes != tipo_carga or tipo != normalize_text(unidad_transporte):
                        continue

                    if cod_ori_str == cod_ori_maestro and cod_dest_str == cod_dest_maestro:
                        valor_base = hoja_maestro[f'N{fila}'].value or 0
                        adicional = hoja_maestro[f'O{fila}'].value or 0
                        valor_total = (
                            valor_base * (n_periferias if n_periferias > 0 else 1)
                            + (adicional * horas_logisticas)
                        )

                        if es_primera_coincidencia:
                            hoja_tarifario[f'E{fila_tarifario}'] = valor_total
                            hoja_tarifario[f'C{fila_tarifario}'] = f"{destino_tarifario} ({dep_dest})"
                            es_primera_coincidencia = False
                        else:
                            hoja_tarifario.insert_rows(fila_tarifario + 1)
                            hoja_tarifario[f'B{fila_tarifario + 1}'] = origen_tarifario
                            hoja_tarifario[f'C{fila_tarifario + 1}'] = f"{destino_tarifario} ({dep_dest})"
                            hoja_tarifario[f'E{fila_tarifario + 1}'] = valor_total
                            offset += 1

                        encontrado = True
                        break

                if not encontrado and es_primera_coincidencia:
                    hoja_tarifario[f'E{fila_tarifario}'] = "No encontrado"

def ejecutar_tarificador(tipo_vehiculo, tipo_carga, unidad_transporte, archivo_tarifario, maestros, horas_logisticas):
    ruta_maestro = maestros.get(tipo_vehiculo)
    if not ruta_maestro or not os.path.exists(ruta_maestro):
        raise FileNotFoundError("Tipo de vehículo no válido o archivo no encontrado.")

    if not archivo_tarifario.lower().endswith(".xlsx"):
        raise ValueError("El archivo del tarifario debe ser un .xlsx")

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
