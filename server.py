import os
import glob
import traceback
import pandas as pd
from typing import Optional, List, Dict, Any
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(
    title="Nexus-BOT API & MCP Server",
    description="Servidor de herramientas jurisprudenciales y tesauros multidisciplinares",
    version="2.0.0"
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "datos")
PATH_RESOLUCIONES = os.path.join(DATA_DIR, "Reporte Temas Subtemas Analizados.xlsx")

ERROR_CARGA_RESOLUCIONES = None
ARCHIVOS_TESAUROS_CARGADOS: List[str] = []


def cargar_resoluciones_local(ruta: str) -> pd.DataFrame:
    """Lee la primera hoja del Excel de resoluciones de forma segura."""
    global ERROR_CARGA_RESOLUCIONES
    if not os.path.exists(ruta):
        ERROR_CARGA_RESOLUCIONES = f"El archivo no existe en: {ruta}"
        return pd.DataFrame()
    
    try:
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
        
        columnas_utiles = [
            c for c in ["abogado", "tema", "subtema", "numero_voto", "fecha_voto", "despacho", "no_expediente"]
            if c in df.columns
        ]
        if columnas_utiles:
            df = df[columnas_utiles]
            
        return df
    except Exception as e:
        ERROR_CARGA_RESOLUCIONES = f"Error al leer Excel de resoluciones: {str(e)}"
        return pd.DataFrame()


def cargar_todos_los_tesauros(carpeta_datos: str) -> List[Dict[str, str]]:
    """
    Busca automáticamente TODOS los archivos Excel de tesauros en la carpeta 'datos/'
    y reconstruye su jerarquía (Primer Nivel -> Segundo Nivel -> Tercer Nivel).
    """
    global ARCHIVOS_TESAUROS_CARGADOS
    ARCHIVOS_TESAUROS_CARGADOS = []
    registros_unificados: List[Dict[str, str]] = []

    if not os.path.exists(carpeta_datos):
        return registros_unificados

    # Buscar cualquier .xlsx en datos/ que contenga 'tesauro' en su nombre
    todos_xlsx = glob.glob(os.path.join(carpeta_datos, "*.xlsx"))
    archivos_tesauro = [f for f in todos_xlsx if "tesauro" in os.path.basename(f).lower()]

    for ruta_archivo in sorted(archivos_tesauro):
        nombre_archivo = os.path.basename(ruta_archivo)
        try:
            df = pd.read_excel(ruta_archivo, sheet_name=0, engine="openpyxl")
            
            # Identificar columnas de nivel sin importar espacios o mayúsculas
            col_l1, col_l2, col_l3 = None, None, None
            for col in df.columns:
                c_norm = str(col).strip().lower()
                if "primer" in c_norm:
                    col_l1 = col
                elif "segundo" in c_norm:
                    col_l2 = col
                elif "tercer" in c_norm:
                    col_l3 = col

            # Materia por defecto derivada del nombre del archivo (ej. TESAURO-AGRARIO.xlsx -> AGRARIO)
            materia_default = (
                nombre_archivo.replace(".xlsx", "")
                .replace("TESAURO-", "").replace("tesauro-", "")
                .replace("_", " ").strip().title()
            )
            
            nivel_1_actual = materia_default
            nivel_2_actual = "General"

            for _, fila in df.iterrows():
                if col_l1 and pd.notna(fila[col_l1]) and str(fila[col_l1]).strip():
                    nivel_1_actual = str(fila[col_l1]).strip()
                    nivel_2_actual = "General"  # Reiniciar nivel 2 al cambiar de materia principal
                
                if col_l2 and pd.notna(fila[col_l2]) and str(fila[col_l2]).strip():
                    nivel_2_actual = str(fila[col_l2]).strip()
                
                if col_l3 and pd.notna(fila[col_l3]) and str(fila[col_l3]).strip():
                    termino_l3 = str(fila[col_l3]).strip()
                    registros_unificados.append({
                        "materia_primer_nivel": nivel_1_actual,
                        "categoria_segundo_nivel": nivel_2_actual,
                        "descriptor_tercer_nivel": termino_l3,
                        "archivo_origen": nombre_archivo
                    })
                elif (col_l2 and pd.notna(fila[col_l2]) and str(fila[col_l2]).strip()) and not col_l3:
                    # Caso en que el archivo solo tenga hasta segundo nivel
                    registros_unificados.append({
                        "materia_primer_nivel": nivel_1_actual,
                        "categoria_segundo_nivel": nivel_2_actual,
                        "descriptor_tercer_nivel": nivel_2_actual,
                        "archivo_origen": nombre_archivo
                    })

            ARCHIVOS_TESAUROS_CARGADOS.append(nombre_archivo)
        except Exception as e:
            print(f"Error cargando tesauro {nombre_archivo}: {e}")

    return registros_unificados


# Carga en memoria al iniciar el servidor
df_resoluciones = cargar_resoluciones_local(PATH_RESOLUCIONES)
lista_tesauros = cargar_todos_los_tesauros(DATA_DIR)


# --- Endpoint de Estado ---
@app.get("/", tags=["Estado"])
def inicio():
    return {
        "servicio": "Nexus-BOT Knowledge Server",
        "estado": "En línea y operativo",
        "resoluciones": {
            "cargadas": len(df_resoluciones),
            "error": ERROR_CARGA_RESOLUCIONES
        },
        "tesauros": {
            "archivos_detectados": len(ARCHIVOS_TESAUROS_CARGADOS),
            "lista_archivos": ARCHIVOS_TESAUROS_CARGADOS,
            "total_descriptores_indexados": len(lista_tesauros)
        }
    }


# --- Herramienta 1: Búsqueda Jurisprudencial ---
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
                "detalle_error": ERROR_CARGA_RESOLUCIONES or "La tabla de resoluciones está vacía."
            }

        termino_norm = termino.lower().strip()
        resultados = []

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


# --- Herramienta 2: Consulta en Todos los Tesauros Jurídicos ---
@app.get("/api/consultar-tesauro", tags=["Herramientas Jurisprudenciales"])
def consultar_tesauro(
    concepto: str = Query(..., description="Concepto o figura jurídica a buscar en todos los tesauros"),
    materia: Optional[str] = Query(None, description="Filtrar por materia (ej. Agrario, Civil, Penal, Disciplinario)"),
    limite: int = Query(15, description="Cantidad máxima de descriptores a devolver")
):
    """
    Busca términos en todos los archivos Excel de tesauro cargados en 'datos/'.
    """
    if not lista_tesauros:
        return {
            "concepto_consultado": concepto,
            "coincidencias": [],
            "aviso": "No se encontraron archivos Excel con 'TESAURO' en el nombre dentro de la carpeta 'datos/'."
        }

    concepto_norm = concepto.lower().strip()
    coincidencias = []

    for item in lista_tesauros:
        texto_completo = f"{item['materia_primer_nivel']} {item['categoria_segundo_nivel']} {item['descriptor_tercer_nivel']}".lower()
        
        if concepto_norm in texto_completo:
            if materia and materia.lower().strip() not in item["materia_primer_nivel"].lower() and materia.lower().strip() not in item["archivo_origen"].lower():
                continue
            
            coincidencias.append(item)
            if len(coincidencias) >= limite:
                break

    return {
        "concepto_consultado": concepto,
        "materia_filtrada": materia,
        "total_encontrados": len(coincidencias),
        "coincidencias": coincidencias if coincidencias else [
            {"mensaje": f"No se encontraron descriptores en los tesauros para '{concepto}'."}
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
                "description": "Busca descriptores oficiales y su jerarquía (Primer, Segundo y Tercer Nivel) en todos los tesauros jurídicos.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "concepto": {"type": "string", "description": "Término o figura jurídica a consultar"},
                        "materia": {"type": "string", "description": "Filtro opcional por rama jurídica (Agrario, Civil, Penal, etc.)"},
                        "limite": {"type": "integer", "description": "Cantidad máxima de coincidencias (por defecto 15)", "default": 15}
                    },
                    "required": ["concepto"]
                }
            }
        ]
    }