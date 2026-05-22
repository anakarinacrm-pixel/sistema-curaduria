import os
import logging
from datetime import datetime
import openpyxl
from pypdf import PdfReader, PdfWriter

# =========================================================
# LOGS
# =========================================================

os.makedirs("logs", exist_ok=True)
logging.basicConfig(
    filename="logs/procesos.log",
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# =========================================================
# CACHE GLOBAL (🔥 CLAVE)
# =========================================================

CACHE_CATASTRO = {}

# =========================================================
# CARGA INICIAL (SOLO 1 VEZ)
# =========================================================

def cargar_catastro_en_memoria():
    global CACHE_CATASTRO

    if CACHE_CATASTRO:
        return

    ruta = os.path.join("excels", "CB1_REGISTRO CATASTRO BCA.xlsx")

    logging.info("Cargando catastro en memoria...")

    wb = openpyxl.load_workbook(ruta, data_only=True, read_only=True)

    for hoja in wb.sheetnames:
        ws = wb[hoja]
        filas = ws.iter_rows(values_only=True)

        encabezados = next(filas)
        encabezados_lower = [str(h).lower() if h else "" for h in encabezados]

        idx_cedula = next((i for i, h in enumerate(encabezados_lower)
                           if "cedula" in h or "documento" in h or "nit" in h), None)

        idx_codigo = next((i for i, h in enumerate(encabezados_lower)
                           if "codigo" in h or "predial" in h or "catastral" in h), None)

        for fila in filas:
            if not fila:
                continue

            fila_dict = {
                encabezados[i]: fila[i] if i < len(fila) else None
                for i in range(len(encabezados))
            }

            if idx_cedula is not None and fila[idx_cedula]:
                CACHE_CATASTRO[str(fila[idx_cedula]).strip()] = fila_dict

            if idx_codigo is not None and fila[idx_codigo]:
                CACHE_CATASTRO[str(fila[idx_codigo]).strip()] = fila_dict

    wb.close()

    logging.info(f"Catastro cargado: {len(CACHE_CATASTRO)} registros")


# =========================================================
# CONSULTA RÁPIDA ⚡
# =========================================================

def obtener_datos_predio(criterio: str):
    datos = CACHE_CATASTRO.get(str(criterio).strip())

    if not datos:
        return None

    direccion = ""
    barrio = ""
    vereda = ""
    matricula = ""
    codigo = ""
    nombre = ""
    tipo_suelo = "URBANO"

    for k, v in datos.items():
        clave = str(k).lower()

        if "direccion" in clave:
            direccion = str(v)

        elif "barrio" in clave:
            barrio = str(v)

        elif "vereda" in clave:
            vereda = str(v)
            tipo_suelo = "RURAL"

        elif "matricula" in clave:
            matricula = str(v)

        elif "codigo" in clave or "predial" in clave:
            codigo = str(v)

        elif "nombre" in clave or "propietario" in clave:
            nombre = str(v)

    return {
        "nombre": nombre,
        "direccion": direccion,
        "barrio_vereda": vereda if tipo_suelo == "RURAL" else barrio,
        "matricula": matricula,
        "codigo_predial": codigo,
        "tipo_suelo": tipo_suelo
    }

# =========================================================
# PDF
# =========================================================

VALOR_ACTIVACION = {
    "chk_suelo_a": "/Yes_xuhu",
    "chk_suelo_b": "/Yes_xuhu",
}

def llenar_pdf(ruta_plantilla, ruta_salida, campos_texto, casillas, radios):
    reader = PdfReader(ruta_plantilla)
    writer = PdfWriter(clone_from=reader)

    writer.update_page_form_field_values(writer.pages[0], campos_texto)

    valores = {c: VALOR_ACTIVACION.get(c) for c in casillas if c in VALOR_ACTIVACION}
    if valores:
        writer.update_page_form_field_values(writer.pages[0], valores)

    if radios:
        writer.update_page_form_field_values(writer.pages[0], radios)

    writer.set_need_appearances_writer(True)

    with open(ruta_salida, "wb") as f:
        writer.write(f)


def ejecutar_llenado_pdf(criterio: str, datos_usuario: dict):

    ruta_pdf = os.path.join("plantillas_pdf", "formulario_base.pdf")

    campos = {
        "txt_direccion": datos_usuario.get("direccion_predio", ""),
        "txt_barrio": datos_usuario.get("barrio", "") if datos_usuario.get("tipo_suelo") != "RURAL" else "",
        "txt_vereda": datos_usuario.get("barrio", "") if datos_usuario.get("tipo_suelo") == "RURAL" else "",
        "txt_matricula": datos_usuario.get("cedula_catastral", ""),
        "txt_codigo_predial": datos_usuario.get("codigo_predial", ""),
    }

    casillas = ["chk_suelo_b" if datos_usuario.get("tipo_suelo") == "RURAL" else "chk_suelo_a"]

    if datos_usuario.get("casillas_extra"):
        casillas.extend(datos_usuario["casillas_extra"])

    radios = {}
    if datos_usuario.get("grupo_clima"):
        radios["grupo_clima"] = datos_usuario["grupo_clima"]

    nombre = f"formulario_{criterio}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
    ruta_salida = os.path.join("generados", nombre)

    llenar_pdf(ruta_pdf, ruta_salida, campos, casillas, radios)

    return ruta_salida, nombre
