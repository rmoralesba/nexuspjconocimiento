import os
import json
import sqlite3
import traceback
import unicodedata
from contextlib import asynccontextmanager
from typing import Optional
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware
from mcp.server.fastmcp import FastMCP

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "nexus_knowledge.db")


def normalizar_texto(texto: str) -> str:
    s = str(texto).strip().lower()
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def leer_metadata() -> dict:
    if not os.path.exists(DB_PATH):
        return {
            "total_resoluciones": 0,
            "total_tesauros": 0,
            "error_resoluciones": "Base de datos SQLite no encontrada. Ejecuta build_db.py en el Build Command.",
            "archivos_tesauros": []
        }
    with sqlite3.connect(DB_PATH) as conn:
        filas = dict(conn.execute("SELECT clave, valor FROM metadata").fetchall())
    return {
        "total_resoluciones": int(filas.get("total_resoluciones", 0)),
        "total_tesauros": int(filas.get("total_tesauros", 0)),
        "error_resoluciones": filas.get("error_resoluciones") or None,
        "archivos_tesauros": json.loads(filas.get("archivos_tesauros", "[]"))
    }


META = leer_metadata()


# ==============================================================================
# LÓGICA COMPARTIDA (USADA TANTO POR REST COMO POR MCP)
# ==============================================================================
def ejecutar_busqueda_jurisprudencia(
    termino: str,
    despacho: Optional[str] = None,
    limite: int = 5,
    offset: int = 0
) -> dict:
    limite = max(1, min(limite, 5000))
    offset = max(0, offset)

    if META["total_resoluciones"] == 0:
        return {
            "total_registros_coincidentes": 0,
            "total_resoluciones_distintas": 0,
            "total_devueltos": 0,
            "offset": offset,
            "limite": limite,
            "tiene_mas": False,
            "siguiente_offset": None,
            "resultados": [],
            "detalle_error": META["error_resoluciones"] or "La tabla de resoluciones está vacía."
        }

    termino_norm = normalizar_texto(termino)
    params = [f"%{termino_norm}%"]
    where_clause = "texto_busqueda LIKE ?"

    if despacho and despacho.strip():
        where_clause += " AND despacho_norm LIKE ?"
        params.append(f"%{normalizar_texto(despacho)}%")

    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()

        cur.execute(
            f"SELECT COUNT(*) AS total, COUNT(DISTINCT numero_voto || '|' || no_expediente) AS distintas "
            f"FROM resoluciones WHERE {where_clause}",
            params
        )
        stats = cur.fetchone()

        cur.execute(
            f"""
            SELECT id, numero_voto, no_expediente, fecha_voto, despacho, abogado, tema, subtema
            FROM resoluciones
            WHERE {where_clause}
            ORDER BY id ASC
            LIMIT ? OFFSET ?
            """,
            params + [limite, offset]
        )
        filas = cur.fetchall()

    resultados = [
        {
            "id_registro": f["id"],
            "numero_voto": f["numero_voto"],
            "numero_expediente": f["no_expediente"],
            "fecha_voto": f["fecha_voto"],
            "despacho": f["despacho"],
            "abogado": f["abogado"],
            "tema": f["tema"],
            "subtema": f["subtema"]
        }
        for f in filas
    ]

    total_coincidentes = stats["total"]
    tiene_mas = (offset + len(resultados)) < total_coincidentes
    siguiente_offset = (offset + len(resultados)) if tiene_mas else None

    return {
        "termino_buscado": termino,
        "despacho_filtrado": despacho,
        "total_registros_coincidentes": total_coincidentes,
        "total_resoluciones_distintas": stats["distintas"],
        "offset": offset,
        "limite": limite,
        "total_devueltos": len(resultados),
        "tiene_mas": tiene_mas,
        "siguiente_offset": siguiente_offset,
        "resultados": resultados
    }


def ejecutar_consulta_tesauro(
    concepto: str,
    materia: Optional[str] = None,
    limite: int = 15
) -> dict:
    limite = max(1, min(limite, 500))

    if META["total_tesauros"] == 0:
        return {
            "concepto_consultado": concepto,
            "coincidencias": [],
            "aviso": "No se encontraron archivos Excel con 'TESAURO' en el nombre dentro de la carpeta 'datos/'."
        }

    concepto_norm = normalizar_texto(concepto)
    params = [f"%{concepto_norm}%"]
    where_clause = "texto_busqueda LIKE ?"

    if materia and materia.strip():
        mat_norm = f"%{normalizar_texto(materia)}%"
        where_clause += " AND (materia_norm LIKE ? OR archivo_norm LIKE ?)"
        params.extend([mat_norm, mat_norm])

    with sqlite3.connect(DB_PATH) as conn:
        conn.row_factory = sqlite3.Row
        cur = conn.cursor()
        cur.execute(
            f"""
            SELECT materia_primer_nivel, categoria_segundo_nivel, descriptor_tercer_nivel, archivo_origen
            FROM tesauros
            WHERE {where_clause}
            LIMIT ?
            """,
            params + [limite]
        )
        filas = cur.fetchall()

    coincidencias = [dict(f) for f in filas]

    return {
        "concepto_consultado": concepto,
        "materia_filtrada": materia,
        "total_encontrados": len(coincidencias),
        "coincidencias": coincidencias if coincidencias else [
            {"mensaje": f"No se encontraron descriptores en los tesauros para '{concepto}'."}
        ]
    }


# ==============================================================================
# SERVIDOR MCP REAL (CONEXIÓN DIRECTA PARA CHATGPT EN /mcp)
# ==============================================================================
mcp = FastMCP(
    "Nexus-BOT Knowledge Server",
    stateless_http=True,
    json_response=True
)


@mcp.tool()
def buscar_jurisprudencia_mcp(
    termino: str,
    despacho: Optional[str] = None,
    limite: int = 20,
    offset: int = 0
) -> dict:
    """
    Busca resoluciones judiciales por tema, subtema, abogado, número de voto/sentencia, despacho o expediente.
    Soporta paginación estable mediante 'limite' (máximo 5000) y 'offset' (ej. 0, 20, 40).
    """
    return ejecutar_busqueda_jurisprudencia(termino=termino, despacho=despacho, limite=limite, offset=offset)


@mcp.tool()
def consultar_tesauro_mcp(
    concepto: str,
    materia: Optional[str] = None,
    limite: int = 15
) -> dict:
    """
    Busca descriptores oficiales y su jerarquía (Primer, Segundo y Tercer Nivel) en todos los tesauros jurídicos.
    """
    return ejecutar_consulta_tesauro(concepto=concepto, materia=materia, limite=limite)


# ==============================================================================
# APLICACIÓN FASTAPI + MONTAJE DEL ENDPOINT /mcp
# ==============================================================================
@asynccontextmanager
async def lifespan(app: FastAPI):
    async with mcp.session_manager.run():
        yield


app = FastAPI(
    title="Nexus-BOT API & MCP Server",
    description="Servidor REST y MCP de herramientas jurisprudenciales y tesauros multidisciplinares",
    version="2.3.0",
    lifespan=lifespan
)

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/", tags=["Estado"])
def inicio():
    return {
        "servicio": "Nexus-BOT Knowledge Server",
        "estado": "En línea y operativo",
        "motor_consulta": "SQLite precompilado en Build",
        "mcp_endpoint": "/mcp",
        "resoluciones": {
            "cargadas": META["total_resoluciones"],
            "error": META["error_resoluciones"]
        },
        "tesauros": {
            "archivos_detectados": len(META["archivos_tesauros"]),
            "lista_archivos": META["archivos_tesauros"],
            "total_descriptores_indexados": META["total_tesauros"]
        }
    }


@app.get("/api/buscar-jurisprudencia", tags=["Herramientas Jurisprudenciales"])
def buscar_jurisprudencia(
    termino: str = Query(..., description="Término clave, tema, subtema, abogado, voto o expediente"),
    despacho: Optional[str] = Query(None, description="Filtrar por despacho judicial"),
    limite: int = Query(5, ge=1, le=5000, description="Cantidad máxima de registros a devolver (máx. 5000)"),
    offset: int = Query(0, ge=0, description="Desplazamiento inicial para paginación (ej. 0, 5000, 10000)")
):
    try:
        return ejecutar_busqueda_jurisprudencia(termino=termino, despacho=despacho, limite=limite, offset=offset)
    except Exception as e:
        return {
            "error_ejecucion": str(e),
            "traza": traceback.format_exc().splitlines()[-3:]
        }


@app.get("/api/consultar-tesauro", tags=["Herramientas Jurisprudenciales"])
def consultar_tesauro(
    concepto: str = Query(..., description="Concepto o figura jurídica a buscar en todos los tesauros"),
    materia: Optional[str] = Query(None, description="Filtrar por materia (ej. Agrario, Civil, Penal, Disciplinario)"),
    limite: int = Query(15, ge=1, le=500, description="Cantidad máxima de descriptores a devolver")
):
    return ejecutar_consulta_tesauro(concepto=concepto, materia=materia, limite=limite)


@app.get("/mcp/tools", tags=["MCP Tools"])
def obtener_herramientas_disponibles():
    return {
        "mcp_endpoint": "/mcp",
        "tools": [
            {
                "name": "buscar_jurisprudencia",
                "description": "Busca resoluciones judiciales por tema, subtema, abogado, número de voto/sentencia, despacho o expediente.",
                "parameters": {
                    "type": "object",
                    "properties": {
                        "termino": {"type": "string", "description": "Texto clave a buscar"},
                        "despacho": {"type": "string", "description": "Filtro opcional por despacho judicial"},
                        "limite": {"type": "integer", "description": "Cantidad de resultados por página (por defecto 5, máximo 5000)", "default": 5},
                        "offset": {"type": "integer", "description": "Desplazamiento inicial para paginación (por defecto 0)", "default": 0}
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


# Monta la app oficial de MCP al final para exponer exactamente la ruta /mcp sin redirecciones 307
app.mount("/", mcp.streamable_http_app())