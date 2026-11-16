import os
import pandas as pd
from typing import Optional
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(
    title="Nexus-BOT API & MCP Server",
    description="Servidor de herramientas y consulta jurisprudencial independiente",
    version="1.3.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==============================================================================
# RUTA LOCAL DENTRO DEL SERVIDOR RENDER (Sin bloqueos de red ni firewalls)
# ==============================================================================
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PATH_RESOLUCIONES = os.path.join(BASE_DIR, "Reporte Temas Subtemas Analizados.xlsx")
PATH_TESAUROS = os.path.join(BASE_DIR, "TESAURO-LABORAL.xlsx")


def cargar_resoluciones_local(ruta: str) -> pd.DataFrame:
    """Lee el archivo Excel alojado directamente en el repositorio."""
    if not os.path.exists(ruta):
        return pd.DataFrame()
    
    try:
        # Se leen únicamente las columnas necesarias para optimizar memoria en Render
        df = pd.read_excel(ruta, engine="openpyxl")
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

        return df.rename(columns=mapeo_columnas)
    except Exception as e:
        print(f"Error leyendo Excel local: {e}")
        return pd.DataFrame()


def cargar_tesauros_local(ruta: str) -> pd.DataFrame:
    if not os.path.exists(ruta):
        return pd.DataFrame()
    try:
        return pd.read_excel(ruta, engine="openpyxl").fillna("")
    except Exception:
        return pd.DataFrame()


# Carga en memoria al encender el servidor
df_resoluciones = cargar_resoluciones_local(PATH_RESOLUCIONES)
df_tesauros = cargar_tesauros_local(PATH_TESAUROS)


@app.get("/", tags=["Estado"])
def inicio():
    return {
        "servicio": "Nexus-BOT Knowledge Server (Local Storage)",
        "estado": "En línea y operativo",
        "archivo_encontrado": os.path.exists(PATH_RESOLUCIONES),
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
    if df_resoluciones.empty:
        return {
            "total_encontrados": 0,
            "resultados": [],
            "aviso": "No se encontró el archivo 'resoluciones.xlsx' junto a server.py en el repositorio."
        }

    termino_norm = termino.lower().strip()
    resultados = []

    mask = pd.Series(False, index=df_resoluciones.index)
    campos_busqueda = ["tema", "subtema", "abogado", "numero_voto", "despacho", "no_expediente"]

    for campo in campos_busqueda:
        if campo in df_resoluciones.columns:
            mask = mask | df_resoluciones[campo].astype(str).str.lower().str.contains(termino_norm, regex=False)

    if despacho and "despacho" in df_resoluciones.columns:
        mask = mask & df_resoluciones["despacho"].astype(str).str.lower().str.contains(despacho.lower().strip(), regex=False)

    coincidencias = df_resoluciones[mask].head(limite)

    for _, fila in coincidencias.iterrows():
        resultados.append({
            "numero_voto": str(fila.get("numero_voto", "N/D")),
            "numero_expediente": str(fila.get("no_expediente", "N/D")),
            "fecha_voto": str(fila.get("fecha_voto", "N/D")),
            "despacho": str(fila.get("despacho", "N/D")),
            "abogado": str(fila.get("abogado", "N/D")),
            "tema": str(fila.get("tema", "N/D")),
            "subtema": str(fila.get("subtema", "N/D"))
        })

    return {
        "termino_buscado": termino,
        "despacho_filtrado": despacho,
        "total_encontrados": len(resultados),
        "resultados": resultados
    }


@app.get("/api/consultar-tesauro", tags=["Herramientas Jurisprudenciales"])
def consultar_tesauro(
    concepto: str = Query(..., description="Concepto o figura jurídica a normalizar")
):
    if df_tesauros.empty:
        return {"coincidencias": [{"mensaje": "No se encontró el archivo 'tesauros.xlsx' en el repositorio."}]}

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