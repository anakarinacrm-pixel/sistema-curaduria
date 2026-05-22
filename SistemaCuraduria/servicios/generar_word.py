from docx import Document
from datetime import datetime
import os

def generar_documento_word(datos):
    # Asegurar que exista la carpeta de salida
    os.makedirs("generados", exist_ok=True)
    
    ruta_plantilla = "plantillas/persona_natural.docx"
    if not os.path.exists(ruta_plantilla):
        raise Exception("No existe la plantilla persona_natural.docx en la carpeta plantillas/")
        
    doc = Document(ruta_plantilla)
    
    # Mapeamos tus variables con los datos que llegan del formulario web
    variables = {
        "{{nombre_propietario}}": datos["nombre_propietario"],
        "{{cedula_propietario}}": datos["cedula_propietario"],
        "{{direccion_predio}}": datos["direccion_predio"],
        "{{nombre_apoderado}}": datos["nombre_apoderado"],
        "{{cedula_apoderado}}": datos["cedula_apoderado"],
        "{{lugar_expedicion_apoderado}}": datos["lugar_expedicion_apoderado"]
    }
    
    # ================= TU REEMPLAZO XML (EL BUENO) =================
    def reemplazar_en_elemento(element):
        for clave, valor in variables.items():
            if clave in element.text:
                element.text = element.text.replace(clave, valor)

    # Recorrer párrafos normales
    for p in doc.paragraphs:
        reemplazar_en_elemento(p)
        
    # Recorrer tablas (crucial para que no se salte nada)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    reemplazar_en_elemento(p)
                    
    # Generar nombre único con marca de tiempo para evitar sobreescritura
    fecha = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_archivo = f"poder_{datos['cedula_propietario']}_{fecha}.docx"
    ruta_salida = os.path.join("generados", nombre_archivo)
    
    # Guardar documento
    doc.save(ruta_salida)
    
    return ruta_salida, nombre_archivo