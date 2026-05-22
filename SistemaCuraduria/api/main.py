import os
from fastapi import FastAPI, Request, HTTPException
from fastapi.responses import HTMLResponse, FileResponse, JSONResponse
from fastapi.templating import Jinja2Templates
from pydantic import BaseModel
from typing import List, Optional

from servicios.generar_word import generar_documento_word
from servicios.generar_pdf import (
    obtener_datos_predio,
    ejecutar_llenado_pdf,
    cargar_catastro_en_memoria
)

# =========================================================
# CONFIGURACIÓN BASE
# =========================================================

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
TEMPLATES_DIR = os.path.join(BASE_DIR, "templates")
GENERADOS_DIR = os.path.join(BASE_DIR, "generados")

os.makedirs(GENERADOS_DIR, exist_ok=True)

app = FastAPI(title="Sistema Curaduría PRO")
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
    grupo_clima: Optional[str] = None


class GenerarTodo(BaseModel):
    criterio: str
    datos_formulario: DatosFormulario
    casillas_extra: Optional[List[str]] = []
    grupo_clima: Optional[str] = None


# =========================================================
# STARTUP (CARGA EN MEMORIA 🔥)
# =========================================================

@app.on_event("startup")
def startup():
    cargar_catastro_en_memoria()
    print("✅ Catastro cargado en memoria")

# =========================================================
# VISTA PRINCIPAL
# =========================================================

@app.get("/", response_class=HTMLResponse)
async def leer_formulario(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})


# =========================================================
# 🔍 CONSULTA RÁPIDA (YA EN MEMORIA)
# =========================================================

@app.post("/consultar-catastro")
async def consultar_catastro(data: dict = None):

    if not data:
        raise HTTPException(status_code=400, detail="No se enviaron datos")

    criterio = data.get("text-1") or data.get("criterio")

    if not criterio:
        criterio = next(iter(data.values()), None)

    if not criterio or str(criterio).strip() == "":
        raise HTTPException(status_code=400, detail="Criterio vacío")

    resultado = obtener_datos_predio(str(criterio).strip())

    if not resultado:
        raise HTTPException(status_code=404, detail="No encontrado")

    return resultado


# =========================================================
# 📄 WORD
# =========================================================

@app.post("/generar-word")
async def generar_word(datos: DatosFormulario):
    try:
        ruta, nombre = generar_documento_word(datos.model_dump())

        return FileResponse(
            path=ruta,
            filename=nombre,
            media_type='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# =========================================================
# 📄 PDF
# =========================================================

@app.post("/generar-pdf")
async def generar_pdf(data: FormularioPdfFinal):
    try:
        datos = data.model_dump()
        criterio = datos.get("criterio")

        if not criterio:
            raise HTTPException(status_code=400, detail="Criterio vacío")

        ruta, nombre = ejecutar_llenado_pdf(criterio, datos)

        return FileResponse(
            path=ruta,
            filename=nombre,
            media_type='application/pdf'
        )

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


# =========================================================
# 💥 GENERAR TODO (PDF + WORD EN UNA SOLA LLAMADA)
# =========================================================

@app.post("/generar-todo")
async def generar_todo(data: GenerarTodo):
    try:
        criterio = data.criterio

        if not criterio:
            raise HTTPException(status_code=400, detail="Criterio vacío")

        # 🔎 Buscar datos reales del predio
        datos_predio = obtener_datos_predio(criterio)

        if not datos_predio:
            raise HTTPException(status_code=404, detail="Predio no encontrado")

        # 🔗 Unir TODO
        datos_formulario = data.datos_formulario.model_dump()
        datos_completos = {**datos_predio, **datos_formulario}

        # Extras PDF
        datos_completos["casillas_extra"] = data.casillas_extra or []
        datos_completos["grupo_clima"] = data.grupo_clima

        # 📄 Generar PDF
        pdf_ruta, pdf_nombre = ejecutar_llenado_pdf(criterio, datos_completos)

        # 📝 Generar Word
        word_ruta, word_nombre = generar_documento_word(datos_completos)

        return JSONResponse({
            "ok": True,
            "pdf": pdf_nombre,
            "word": word_nombre
        })

    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))
