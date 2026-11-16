import os
import io
import urllib3
import requests
import pandas as pd
from typing import Optional
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware

# Desactivar advertencias de certificados SSL institucionales
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

app = FastAPI(
    title="Nexus-BOT API & MCP Server",
    description="Servidor de herramientas y consulta jurisprudencial independiente",
    version="1.2.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ==============================================================================
# CONFIGURACIÓN DE RUTAS HTTPS EN EL SERVIDOR INSTITUCIONAL
# Verifica que el nombre del archivo coincida exactamente con el subido a /files/BI/
# ==============================================================================
BASE_URL_BI = "https://cij.poder-judicial.go.cr/files/BI/"

URL_RESOLUCIONES = f"{BASE_URL_BI}Reporte%20Temas%20Subtemas%20Analizados.xlsx"
URL_TESAUROS = f"{BASE_URL_BI}tesauros.xlsx"


def descargar_y_leer_excel(url: str) -> pd.DataFrame:
    """Descarga el archivo Excel desde cij.poder-judicial.go.cr simulando un navegador estándar."""
    try:
        cabeceras = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
        }
        # verify=False evita errores si el servidor institucional usa una CA intermedia no estándar
        respuesta = requests.get(url, headers=cabeceras, timeout=45, verify=False)
        respuesta.raise_for_status()
        
        archivo_memoria = io.BytesIO(respuesta.content)
        # engine="openpyxl" lee el archivo de forma eficiente
        df = pd.read_excel(archivo_memoria, engine="openpyxl")
        return df.fillna("")
    except Exception as e:
        print(f"Error al descargar desde {url}: {e}")
        return pd.DataFrame()


def cargar_resoluciones_desde_url(url: str) -> pd.DataFrame:
    """Descarga y estandariza las columnas del Excel de resoluciones."""
    df = descargar_y_leer_excel(url)
    if df.empty:
        return df

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


# Carga inicial en memoria al encender el servidor
df_resoluciones = cargar_resoluciones_desde_url(URL_RESOLUCIONES)
df_tesauros = descargar_y_leer_excel(URL_TESAUROS)


# --- Endpoints de Estado y Recarga ---
@app.get("/", tags=["Estado"])
def inicio():
    return {
        "servicio": "Nexus-BOT Knowledge Server",
        "estado": "En línea y operativo",
        "url_origen": URL_RESOLUCIONES,
        "registros_cargados": {
            "resoluciones": len(df_resoluciones),
            "tesauros": len(df_tesauros)
        }
    }


@app.post("/api/recargar-datos", tags=["Mantenimiento"])
def recargar_datos():
    """Vuelve a leer los archivos Excel desde cij.poder-judicial.go.cr sin reiniciar el servidor."""
    global df_resoluciones, df_tesauros
    df_resoluciones = cargar_resoluciones_desde_url(URL_RESOLUCIONES)
    df_tesauros = descargar_y_leer_excel(URL_TESAUROS)
    
    return {
        "mensaje": "Bases de datos actualizadas correctamente",
        "resoluciones_activas": len(df_resoluciones),
        "tesauros_activos": len(df_tesauros)
    }


# --- Herramienta 1: Búsqueda Jurisprudencial ---
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
            "aviso": f"No se pudieron cargar datos desde {URL_RESOLUCIONES}. Verifica que el archivo exista en esa ruta."
        }

    termino_norm = termino.lower().strip()
    resultados = []

    # Búsqueda vectorizada rápida en Pandas (ideal para archivos grandes)
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


# --- Herramienta 2: Consulta de Tesauro Jurídico ---
@app.get("/api/consultar-tesauro", tags=["Herramientas Jurisprudenciales"])
def consultar_tesauro(
    concepto: str = Query(..., description="Concepto o figura jurídica a normalizar")
):
    if df_tesauros.empty:
        return {"coincidencias": [{"mensaje": "No hay archivo de tesauros cargado en /files/BI/."}]}

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


# --- Declaración del Manifiesto MCP / Herramientas ---
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