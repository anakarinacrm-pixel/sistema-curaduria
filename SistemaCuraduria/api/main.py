from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, FileResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Optional
from servicios.generar_word import generar_documento_word
from servicios.generar_pdf import obtener_datos_predio, ejecutar_llenado_pdf
import os

# =========================================================
# CONFIGURACIÓN BASE (IMPORTANTE PARA RENDER)
# =========================================================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
GENERADOS_DIR = os.path.join(BASE_DIR, "generados")

os.makedirs(GENERADOS_DIR, exist_ok=True)

app = FastAPI(title="Sistema Curaduría")
templates = Jinja2Templates(directory=TEMPLATES_DIR)

# =========================================================
# MODELOS
# =========================================================

class DatosFormulario(BaseModel):
    nombre_propietario: str
    cedula_propietario: str
    direccion_predio: str
    nombre_apoderado: str
    cedula_apoderado: str
    lugar_expedicion_apoderado: str

class FormularioPdfFinal(BaseModel):
    criterio: str
    nombre_propietario: Optional[str] = ""
    direccion_predio: Optional[str] = ""
    barrio: Optional[str] = ""
    cedula_catastral: Optional[str] = ""
    codigo_predial: Optional[str] = ""
    tipo_suelo: Optional[str] = "URBANO"
    casillas_extra: Optional[List[str]] = []

# =========================================================
# ENDPOINTS
# =========================================================

@app.get("/", response_class=HTMLResponse)
async def leer_formulario(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


@app.post("/consultar-catastro")
async def consultar_catastro(data: dict = None):

    if not data:
        return {"status": "success"}

    criterio_valor = data.get("text-1") or data.get("criterio")

    if not criterio_valor and isinstance(data, dict):
        criterio_valor = next(iter(data.values()), None)

    if not criterio_valor:
        raise HTTPException(status_code=400, detail="Criterio vacío")

    resultado = obtener_datos_predio(str(criterio_valor).strip())

    if not resultado:
        raise HTTPException(status_code=404, detail="No encontrado")

    return resultado


@app.post("/generar-word")
async def procesar_formulario(datos: DatosFormulario):
    try:
        ruta, nombre = generar_documento_word(datos.model_dump())

        return FileResponse(
            path=ruta,
            filename=nombre,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@app.post("/generar-pdf")
async def generar_pdf_final(data: FormularioPdfFinal):
    try:
        ruta, nombre = ejecutar_llenado_pdf(data.model_dump())

        return FileResponse(
            path=ruta,
            filename=nombre,
            media_type='application/pdf'
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))