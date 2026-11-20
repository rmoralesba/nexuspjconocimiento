import os
import gc
import glob
import sqlite3
import tempfile
import traceback
import unicodedata
import pandas as pd
from typing import Optional, List
from fastapi import FastAPI, Query
from fastapi.middleware.cors import CORSMiddleware

app = FastAPI(
    title="Nexus-BOT API & MCP Server",
    description="Servidor de herramientas jurisprudenciales y tesauros multidisciplinares",
    version="2.1.0"
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
DB_PATH = os.path.join(tempfile.gettempdir(), "nexus_knowledge.db")

ERROR_CARGA_RESOLUCIONES: Optional[str] = None
ARCHIVOS_TESAUROS_CARGADOS: List[str] = []
TOTAL_RESOLUCIONES_CARGADAS: int = 0
TOTAL_TESAUROS_CARGADOS: int = 0


def normalizar_texto(texto: object) -> str:
    """Convierte a minúsculas, elimina tildes y espacios sobrantes para búsquedas rápidas."""
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    s = str(texto).strip().lower()
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def limpiar_valor(val: object) -> str:
    """Garantiza un string limpio sin valores NaN/NaT ni estructuras de Pandas."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "N/D"
    if isinstance(val, pd.Timestamp):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return "N/D"
    return s


def inicializar_db() -> sqlite3.Connection:
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)
    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE resoluciones (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            numero_voto TEXT,
            no_expediente TEXT,
            fecha_voto TEXT,
            despacho TEXT,
            abogado TEXT,
            tema TEXT,
            subtema TEXT,
            despacho_norm TEXT,
            texto_busqueda TEXT
        )
    """)
    cur.execute("""
        CREATE TABLE tesauros (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            materia_primer_nivel TEXT,
            categoria_segundo_nivel TEXT,
            descriptor_tercer_nivel TEXT,
            archivo_origen TEXT,
            materia_norm TEXT,
            archivo_norm TEXT,
            texto_busqueda TEXT
        )
    """)
    cur.execute("CREATE INDEX idx_res_despacho ON resoluciones(despacho_norm)")
    cur.execute("CREATE INDEX idx_tes_materia ON tesauros(materia_norm)")
    conn.commit()
    return conn


def cargar_resoluciones_a_sqlite(conn: sqlite3.Connection, ruta: str) -> int:
    """Lee el Excel de resoluciones, corrige el mapeo de columnas y lo vuelca a SQLite."""
    global ERROR_CARGA_RESOLUCIONES
    if not os.path.exists(ruta):
        ERROR_CARGA_RESOLUCIONES = f"El archivo no existe en: {ruta}"
        return 0

    try:
        df = pd.read_excel(ruta, sheet_name=0, engine="openpyxl")

        # Mapeo seguro: evaluar 'fecha' ANTES que 'voto' y evitar columnas duplicadas
        mapeo_columnas = {}
        destinos_usados = set()

        for col in df.columns:
            col_limpia = normalizar_texto(col)
            destino = None

            if "fecha" in col_limpia:
                destino = "fecha_voto"
            elif "abogado" in col_limpia:
                destino = "abogado"
            elif "subtema" in col_limpia:
                destino = "subtema"
            elif "tema" in col_limpia:
                destino = "tema"
            elif "expediente" in col_limpia:
                destino = "no_expediente"
            elif "despacho" in col_limpia:
                destino = "despacho"
            elif "voto" in col_limpia or "resolucion" in col_limpia or "sentencia" in col_limpia:
                destino = "numero_voto"

            if destino and destino not in destinos_usados:
                mapeo_columnas[col] = destino
                destinos_usados.add(destino)

        df = df.rename(columns=mapeo_columnas)

        columnas_esperadas = [
            "numero_voto", "no_expediente", "fecha_voto",
            "despacho", "abogado", "tema", "subtema"
        ]
        for c in columnas_esperadas:
            if c not in df.columns:
                df[c] = "N/D"

        df = df[columnas_esperadas].drop_duplicates()

        filas_insertar = []
        for row in df.itertuples(index=False):
            num_voto = limpiar_valor(row.numero_voto)
            no_exp = limpiar_valor(row.no_expediente)
            fecha = limpiar_valor(row.fecha_voto)
            desp = limpiar_valor(row.despacho)
            abog = limpiar_valor(row.abogado)
            tema = limpiar_valor(row.tema)
            subtema = limpiar_valor(row.subtema)

            texto_busq = normalizar_texto(f"{num_voto} {no_exp} {fecha} {desp} {abog} {tema} {subtema}")
            filas_insertar.append((
                num_voto, no_exp, fecha, desp, abog, tema, subtema,
                normalizar_texto(desp), texto_busq
            ))

        conn.executemany("""
            INSERT INTO resoluciones (
                numero_voto, no_expediente, fecha_voto, despacho, abogado,
                tema, subtema, despacho_norm, texto_busqueda
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, filas_insertar)
        conn.commit()

        total = len(filas_insertar)
        del df, filas_insertar
        gc.collect()
        return total

    except Exception as e:
        ERROR_CARGA_RESOLUCIONES = f"Error al leer Excel de resoluciones: {str(e)}"
        return 0


def cargar_tesauros_a_sqlite(conn: sqlite3.Connection, carpeta_datos: str) -> int:
    """Carga los archivos Excel de tesauros a SQLite y libera la memoria de cada archivo."""
    global ARCHIVOS_TESAUROS_CARGADOS
    ARCHIVOS_TESAUROS_CARGADOS = []
    total_insertados = 0

    if not os.path.exists(carpeta_datos):
        return 0

    todos_xlsx = glob.glob(os.path.join(carpeta_datos, "*.xlsx"))
    archivos_tesauro = [f for f in todos_xlsx if "tesauro" in os.path.basename(f).lower()]

    for ruta_archivo in sorted(archivos_tesauro):
        nombre_archivo = os.path.basename(ruta_archivo)
        try:
            df = pd.read_excel(ruta_archivo, sheet_name=0, engine="openpyxl")

            col_l1, col_l2, col_l3 = None, None, None
            for col in df.columns:
                c_norm = normalizar_texto(col)
                if "primer" in c_norm:
                    col_l1 = col
                elif "segundo" in c_norm:
                    col_l2 = col
                elif "tercer" in c_norm:
                    col_l3 = col

            materia_default = (
                nombre_archivo.replace(".xlsx", "")
                .replace("TESAURO-", "").replace("tesauro-", "")
                .replace("_", " ").strip().title()
            )

            nivel_1_actual = materia_default
            nivel_2_actual = "General"
            lote = []

            for _, fila in df.iterrows():
                if col_l1 and pd.notna(fila[col_l1]) and str(fila[col_l1]).strip():
                    nivel_1_actual = str(fila[col_l1]).strip()
                    nivel_2_actual = "General"

                if col_l2 and pd.notna(fila[col_l2]) and str(fila[col_l2]).strip():
                    nivel_2_actual = str(fila[col_l2]).strip()

                termino_l3 = None
                if col_l3 and pd.notna(fila[col_l3]) and str(fila[col_l3]).strip():
                    termino_l3 = str(fila[col_l3]).strip()
                elif col_l2 and pd.notna(fila[col_l2]) and str(fila[col_l2]).strip() and not col_l3:
                    termino_l3 = nivel_2_actual

                if termino_l3:
                    texto_busq = normalizar_texto(f"{nivel_1_actual} {nivel_2_actual} {termino_l3}")
                    lote.append((
                        nivel_1_actual, nivel_2_actual, termino_l3, nombre_archivo,
                        normalizar_texto(nivel_1_actual), normalizar_texto(nombre_archivo), texto_busq
                    ))

            if lote:
                conn.executemany("""
                    INSERT INTO tesauros (
                        materia_primer_nivel, categoria_segundo_nivel, descriptor_tercer_nivel,
                        archivo_origen, materia_norm, archivo_norm, texto_busqueda
                    ) VALUES (?, ?, ?, ?, ?, ?, ?)
                """, lote)
                conn.commit()
                total_insertados += len(lote)

            ARCHIVOS_TESAUROS_CARGADOS.append(nombre_archivo)
            del df, lote
            gc.collect()
        except Exception as e:
            print(f"Error cargando tesauro {nombre_archivo}: {e}")

    return total_insertados


# Inicialización al arrancar el servidor
_conn_init = inicializar_db()
TOTAL_RESOLUCIONES_CARGADAS = cargar_resoluciones_a_sqlite(_conn_init, PATH_RESOLUCIONES)
TOTAL_TESAUROS_CARGADOS = cargar_tesauros_a_sqlite(_conn_init, DATA_DIR)
_conn_init.close()
gc.collect()


# --- Endpoint de Estado ---
@app.get("/", tags=["Estado"])
def inicio():
    return {
        "servicio": "Nexus-BOT Knowledge Server",
        "estado": "En línea y operativo",
        "motor_consulta": "SQLite (bajo consumo de RAM)",
        "resoluciones": {
            "cargadas": TOTAL_RESOLUCIONES_CARGADAS,
            "error": ERROR_CARGA_RESOLUCIONES
        },
        "tesauros": {
            "archivos_detectados": len(ARCHIVOS_TESAUROS_CARGADOS),
            "lista_archivos": ARCHIVOS_TESAUROS_CARGADOS,
            "total_descriptores_indexados": TOTAL_TESAUROS_CARGADOS
        }
    }


# --- Herramienta 1: Búsqueda Jurisprudencial ---
@app.get("/api/buscar-jurisprudencia", tags=["Herramientas Jurisprudenciales"])
def buscar_jurisprudencia(
    termino: str = Query(..., description="Término clave, tema, subtema, abogado, voto o expediente"),
    despacho: Optional[str] = Query(None, description="Filtrar por despacho judicial"),
    limite: int = Query(5, ge=1, le=5000, description="Cantidad máxima de registros a devolver")
):
    try:
        if TOTAL_RESOLUCIONES_CARGADAS == 0:
            return {
                "total_encontrados": 0,
                "total_resoluciones_distintas": 0,
                "resultados": [],
                "detalle_error": ERROR_CARGA_RESOLUCIONES or "La tabla de resoluciones está vacía."
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

            # Conteo exacto de registros y resoluciones distintas antes del LIMIT
            cur.execute(
                f"SELECT COUNT(*) AS total, COUNT(DISTINCT numero_voto || '|' || no_expediente) AS distintas "
                f"FROM resoluciones WHERE {where_clause}",
                params
            )
            stats = cur.fetchone()

            cur.execute(
                f"""
                SELECT numero_voto, no_expediente, fecha_voto, despacho, abogado, tema, subtema
                FROM resoluciones
                WHERE {where_clause}
                LIMIT ?
                """,
                params + [limite]
            )
            filas = cur.fetchall()

        resultados = [
            {
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

        return {
            "termino_buscado": termino,
            "despacho_filtrado": despacho,
            "total_registros_coincidentes": stats["total"],
            "total_resoluciones_distintas": stats["distintas"],
            "total_devueltos": len(resultados),
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
    limite: int = Query(15, ge=1, le=500, description="Cantidad máxima de descriptores a devolver")
):
    if TOTAL_TESAUROS_CARGADOS == 0:
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