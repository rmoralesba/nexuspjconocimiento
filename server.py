import os
import traceback
import pandas as pd
from typing import Optional
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(
    title="Nexus-BOT API & MCP Server",
    description="Servidor de herramientas y consulta jurisprudencial independiente",
    version="1.5.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_RESOLUCIONES = os.path.join(BASE_DIR, "datos", "Reporte Temas Subtemas Analizados.xlsx")
PATH_TESAUROS = os.path.join(BASE_DIR, "datos", "TESAURO-LABORAL.xlsx")

ERROR_CARGA = None


def cargar_resoluciones_local(ruta: str) -> pd.DataFrame:
    """Lee la primera hoja del Excel de resoluciones de forma segura."""
    global ERROR_CARGA
    if not os.path.exists(ruta):
        ERROR_CARGA = f"El archivo no existe en la ruta: {ruta}"
        return pd.DataFrame()
    
    try:
        # sheet_name=0 lee siempre la primera pestaña del libro sin importar su nombre
        df = pd.read_excel(ruta, sheet_name=0, engine="openpyxl")
        df = df.fillna("")

        mapeo_columnas = {}
        for col in df.columns:
            col_limpia = (
                str(col).strip().lower()
                .replace("á", "a").replace("é", "e")
                .replace("í", "i").replace("ó", "o").replace("ú", "u")
            )
            if "abogado" in col_limpia:
                mapeo_columnas[col] = "abogado"
            elif "tema" in col_limpia and "sub" not in col_limpia:
                mapeo_columnas[col] = "tema"
            elif "subtema" in col_limpia:
                mapeo_columnas[col] = "subtema"
            elif "voto" in col_limpia or "resolucion" in col_limpia or "sentencia" in col_limpia:
                mapeo_columnas[col] = "numero_voto"
            elif "fecha" in col_limpia:
                mapeo_columnas[col] = "fecha_voto"
            elif "despacho" in col_limpia:
                mapeo_columnas[col] = "despacho"
            elif "expediente" in col_limpia:
                mapeo_columnas[col] = "no_expediente"

        df = df.rename(columns=mapeo_columnas)
        
        # Conservar únicamente las columnas útiles para ahorrar memoria RAM en Render
        columnas_utiles = [c for c in ["abogado", "tema", "subtema", "numero_voto", "fecha_voto", "despacho", "no_expediente"] if c in df.columns]
        if columnas_utiles:
            df = df[columnas_utiles]
            
        return df
    except Exception as e:
        ERROR_CARGA = f"Error al leer Excel: {str(e)}"
        return pd.DataFrame()


def cargar_tesauros_local(ruta: str) -> pd.DataFrame:
    if not os.path.exists(ruta):
        return pd.DataFrame()
    try:
        return pd.read_excel(ruta, sheet_name=0, engine="openpyxl").fillna("")
    except Exception:
        return pd.DataFrame()


df_resoluciones = cargar_resoluciones_local(PATH_RESOLUCIONES)
df_tesauros = cargar_tesauros_local(PATH_TESAUROS)


@app.get("/", tags=["Estado"])
def inicio():
    return {
        "servicio": "Nexus-BOT Knowledge Server",
        "estado": "En línea y operativo",
        "archivo_existe": os.path.exists(PATH_RESOLUCIONES),
        "error_carga": ERROR_CARGA,
        "columnas_detectadas": list(df_resoluciones.columns),
        "registros_cargados": {
            "resoluciones": len(df_resoluciones),
            "tesauros": len(df_tesauros)
        }
    }


@app.get("/api/buscar-jurisprudencia", tags=["Herramientas Jurisprudenciales"])
def buscar_jurisprudencia(
    termino: str = Query(..., description="Término clave, tema, subtema, abogado, voto o expediente"),
    despacho: Optional[str] = Query(None, description="Filtrar por despacho judicial"),
    limite: int = Query(5, description="Cantidad máxima de resoluciones a devolver")
):
    try:
        if df_resoluciones.empty:
            return {
                "total_encontrados": 0,
                "resultados": [],
                "detalle_error": ERROR_CARGA or "La tabla de resoluciones está vacía."
            }

        termino_norm = termino.lower().strip()
        resultados = []

        # Búsqueda fila por fila optimizada (evita duplicar memoria en Render)
        for _, fila in df_resoluciones.iterrows():
            texto_fila = " ".join(str(val) for val in fila.values).lower()
            
            if termino_norm in texto_fila:
                despacho_val = str(fila.get("despacho", "N/D"))
                if despacho and despacho.lower().strip() not in despacho_val.lower():
                    continue
                    
                resultados.append({
                    "numero_voto": str(fila.get("numero_voto", "N/D")),
                    "numero_expediente": str(fila.get("no_expediente", "N/D")),
                    "fecha_voto": str(fila.get("fecha_voto", "N/D")),
                    "despacho": despacho_val,
                    "abogado": str(fila.get("abogado", "N/D")),
                    "tema": str(fila.get("tema", "N/D")),
                    "subtema": str(fila.get("subtema", "N/D"))
                })
                
                if len(resultados) >= limite:
                    break

        return {
            "termino_buscado": termino,
            "despacho_filtrado": despacho,
            "total_encontrados": len(resultados),
            "resultados": resultados
        }
    except Exception as e:
        return {
            "error_ejecucion": str(e),
            "traza": traceback.format_exc().splitlines()[-3:]
        }


@app.get("/api/consultar-tesauro", tags=["Herramientas Jurisprudenciales"])
def consultar_tesauro(
    concepto: str = Query(..., description="Concepto o figura jurídica a normalizar")
):
    if df_tesauros.empty:
        return {"coincidencias": [{"mensaje": "No se encontró el archivo 'datos/tesauros.xlsx' en el repositorio."}]}

    concepto_norm = concepto.lower().strip()
    coincidencias = []

    for _, fila in df_tesauros.iterrows():
        texto_fila = " ".join([str(val) for val in fila.values]).lower()
        if concepto_norm in texto_fila:
            coincidencias.append(fila.to_dict())
            if len(coincidencias) >= 10:
                break

    return {
        "concepto_consultado": concepto,
        "coincidencias": coincidencias if coincidencias else [
            {"mensaje": f"No se encontró término normalizado exacto para '{concepto}'."}
        ]
    }


@app.get("/mcp/tools", tags=["MCP Tools"])
def obtener_herramientas_disponibles():
    return {
        "tools": [
            {
                "name": "buscar_jurisprudencia",
                "description": "Busca resoluciones judiciales por tema, subtema, abogado, número de voto/sentencia, despacho o expediente.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "termino": {"type": "string", "description": "Texto clave a buscar"},
                        "despacho": {"type": "string", "description": "Filtro opcional por despacho judicial"},
                        "limite": {"type": "integer", "description": "Cantidad de resultados (por defecto 5)", "default": 5}
                    },
                    "required": ["termino"]
                }
            },
            {
                "name": "consultar_tesauro",
                "description": "Obtiene conceptos y sinónimos estandarizados para normalizar términos de búsqueda.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "concepto": {"type": "string", "description": "Término legal a consultar"}
                    },
                    "required": ["concepto"]
                }
            }
        ]
    }