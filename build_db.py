import os
import json
import glob
import sqlite3
import unicodedata
import pandas as pd

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(BASE_DIR, "datos")
PATH_RESOLUCIONES = os.path.join(DATA_DIR, "Reporte Temas Subtemas Analizados.xlsx")
DB_PATH = os.path.join(BASE_DIR, "nexus_knowledge.db")


def normalizar_texto(texto: object) -> str:
    if texto is None or (isinstance(texto, float) and pd.isna(texto)):
        return ""
    s = str(texto).strip().lower()
    return "".join(
        c for c in unicodedata.normalize("NFD", s)
        if unicodedata.category(c) != "Mn"
    )


def limpiar_valor(val: object) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return "N/D"
    if isinstance(val, pd.Timestamp):
        return val.strftime("%Y-%m-%d")
    s = str(val).strip()
    if not s or s.lower() in ("nan", "nat", "none"):
        return "N/D"
    return s


def construir_base():
    if os.path.exists(DB_PATH):
        os.remove(DB_PATH)

    conn = sqlite3.connect(DB_PATH)
    cur = conn.cursor()

    cur.execute("""
        CREATE TABLE metadata (
            clave TEXT PRIMARY KEY,
            valor TEXT
        )
    """)
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

    # --- 1. Cargar Resoluciones ---
    error_res = ""
    total_res = 0
    if not os.path.exists(PATH_RESOLUCIONES):
        error_res = f"El archivo no existe en: {PATH_RESOLUCIONES}"
    else:
        try:
            df = pd.read_excel(PATH_RESOLUCIONES, sheet_name=0, engine="openpyxl")
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

            filas = []
            for row in df.itertuples(index=False):
                num_voto = limpiar_valor(row.numero_voto)
                no_exp = limpiar_valor(row.no_expediente)
                fecha = limpiar_valor(row.fecha_voto)
                desp = limpiar_valor(row.despacho)
                abog = limpiar_valor(row.abogado)
                tema = limpiar_valor(row.tema)
                subtema = limpiar_valor(row.subtema)
                texto_busq = normalizar_texto(f"{num_voto} {no_exp} {fecha} {desp} {abog} {tema} {subtema}")
                filas.append((
                    num_voto, no_exp, fecha, desp, abog, tema, subtema,
                    normalizar_texto(desp), texto_busq
                ))

            cur.executemany("""
                INSERT INTO resoluciones (
                    numero_voto, no_expediente, fecha_voto, despacho, abogado,
                    tema, subtema, despacho_norm, texto_busqueda
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            """, filas)
            total_res = len(filas)
            del df, filas
        except Exception as e:
            error_res = f"Error al leer Excel de resoluciones: {str(e)}"

    # --- 2. Cargar Tesauros ---
    archivos_tesauros = []
    total_tes = 0
    if os.path.exists(DATA_DIR):
        todos_xlsx = glob.glob(os.path.join(DATA_DIR, "*.xlsx"))
        archivos_tesauro = [f for f in todos_xlsx if "tesauro" in os.path.basename(f).lower()]

        for ruta_archivo in sorted(archivos_tesauro):
            nombre_archivo = os.path.basename(ruta_archivo)
            try:
                df = pd.read_excel(ruta_archivo, sheet_name=0, engine="openpyxl")
                col_l1 = col_l2 = col_l3 = None
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
                    cur.executemany("""
                        INSERT INTO tesauros (
                            materia_primer_nivel, categoria_segundo_nivel, descriptor_tercer_nivel,
                            archivo_origen, materia_norm, archivo_norm, texto_busqueda
                        ) VALUES (?, ?, ?, ?, ?, ?, ?)
                    """, lote)
                    total_tes += len(lote)
                archivos_tesauros.append(nombre_archivo)
            except Exception as e:
                print(f"Error cargando tesauro {nombre_archivo}: {e}")

    # Índices y metadatos
    cur.execute("CREATE INDEX idx_res_despacho ON resoluciones(despacho_norm)")
    cur.execute("CREATE INDEX idx_tes_materia ON tesauros(materia_norm)")

    meta = [
        ("total_resoluciones", str(total_res)),
        ("total_tesauros", str(total_tes)),
        ("error_resoluciones", error_res),
        ("archivos_tesauros", json.dumps(archivos_tesauros, ensure_ascii=False)),
    ]
    cur.executemany("INSERT INTO metadata (clave, valor) VALUES (?, ?)", meta)

    conn.commit()
    conn.close()
    print(f"Base SQLite generada: {total_res} resoluciones, {total_tes} descriptores de tesauros.")


if __name__ == "__main__":
    construir_base()