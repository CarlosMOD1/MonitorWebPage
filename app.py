#!/usr/bin/env python3
"""
Manufacturing Dashboard - Live Web App
Ejecutar: python app.py
Abrir en navegador: http://localhost:5000
"""
from flask import Flask, jsonify, request, send_from_directory
import pyodbc
import pandas as pd
import json
import re
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv

# Cargar variables de entorno desde el archivo .env
load_dotenv()

import threading
import time

app = Flask(__name__, static_folder="static")

# ─────────────────────────────────────────────────────────────
# CONFIGURACION
# ─────────────────────────────────────────────────────────────
DB_SERVER  = os.getenv("DB_SERVER")
DB_NAME    = os.getenv("DB_NAME")
DB_USER    = os.getenv("DB_USER")
DB_PASS    = os.getenv("DB_PASS")
DB_NAME_QR = "trkprdshippapp"   # Base de datos de tracking de QR/modelNumber

# Mapeo de modelNumber → producto. Agregar nuevos modelos aqui.
MODEL_NUMBER_TO_PRODUCT = {
    "GBP-3001": "SPSF",
    "GBP-3003": "Blade",
}

# Estaciones compartidas SPSF/Blade que NO pueden diferenciarse por modelNumber.
# Sus registros aparecerán duplicados en ambos productos.
UNDIFF_STATIONS = {406, 393}

# Constantes de comportamiento del dashboard
HISTORICAL_LOAD_DAYS = 31
RECENT_REFRESH_DAYS = 2
REFRESH_INTERVAL_MINUTES = 15
REAL_FAILURE_HOURS = 3
DEFAULT_QUERY_RANGE_DAYS = 30
DEBUG_ENDPOINT_RANGE_DAYS = 7
DETAILS_MAX_RECORDS = 500
DB_CONNECT_TIMEOUT_SECONDS = 60
DEFAULT_APP_PORT = 5000

_SPSF_STATIONS = [406,393,178,183,184,185,186,187,244,201,202,190,191,192,193,194,195,196,199,208,211,215,220,232]

GROUPS = {
    "SPSF":        _SPSF_STATIONS,
    "Blade":       _SPSF_STATIONS,   # comparte las mismas estaciones; se diferencia por modelNumber
    "C-Track":     [238,239,241,333],
    "Solar":       [279,275,277,278,279,280,281,284],
    "Black Box":   [203,204,233,237,245],
    "OPAL 3":      [260,261,262,263,264,265,266],
    "Black Box 2": [297,298,300,301,302,303,304,305,306],
    "Opal 4":      [289,290,291,292,293,294,295],
}

# Estaciones que SPSF y Blade comparten (incluyendo las no-diferenciables)
SPSF_BLADE_STATIONS      = set(_SPSF_STATIONS)
# Estaciones donde SI podemos resolver el producto via modelNumber
SPSF_BLADE_DIFF_STATIONS = SPSF_BLADE_STATIONS - UNDIFF_STATIONS

# STATION_TO_GROUP: primer grupo que registra la estacion gana (SPSF para estaciones compartidas).
# Para las estaciones compartidas, prepare_dataframe sobreescribe con el valor de modelNumber.
STATION_TO_GROUP    = {}
ALL_STATION_IDS_SET = set()
for grp, ids in GROUPS.items():
    for sid in ids:
        if sid not in STATION_TO_GROUP:   # first-assignment-wins → SPSF gana sobre Blade
            STATION_TO_GROUP[sid] = grp
        ALL_STATION_IDS_SET.add(sid)

ALL_STATION_IDS = list(ALL_STATION_IDS_SET)
IDS_STR = ",".join(str(s) for s in ALL_STATION_IDS)
GROUP_NAMES = list(GROUPS.keys())

# Orden de estaciones por producto: (producto, stationid) → posicion en la lista hardcodeada.
# Usado en compute_stats para que el sidebar respete el orden definido arriba.
STATION_ORDER = {}
for grp, ids in GROUPS.items():
    for pos, sid in enumerate(ids):
        STATION_ORDER[(grp, sid)] = pos

SQL = f"""
SELECT
    e.endtime,
    s.stationName,
    s.stationid,
    e.failureCode,
    e.notes,
    c.prevQr,
    c.currQr,
    CASE
        WHEN e.failureCode LIKE 'ALL-PASS%%'    THEN 'PASSED'
        WHEN e.failureCode LIKE '%%-PASS-%%'    THEN 'PASSED'
        WHEN e.failureCode = 'MAXPASSEXCEEDED'  THEN 'PASSED'
        ELSE 'FAILED'
    END AS resultado
FROM trk.manufacturingEvents AS e
INNER JOIN trk.qrCorrelation AS c ON e.correlationID = c.correlationID
INNER JOIN trk.stationConfig  AS s ON e.stationid    = s.stationid
WHERE s.stationid IN ({IDS_STR})
"""


# ─────────────────────────────────────────────────────────────
# EXTRACCION DE TIPO DE FALLA DESDE NOTES
# Lógica basada exactamente en los archivos Excel de referencia:
#   Solar       → JSON: comment.fail_tests[0], fallback failure_codes[0]
#   BlackBox    → texto plano: split(" | ")[0] → split(",")[0] → split(":")[0]
#   BlackBox2   → JSON: fail_tests[0] o station.split(" - ")[-1]; texto plano → station name
#   OPAL4       → JSON anidado {"notes":{...}}: failure_codes[0] solo si NO es código patrón
# ─────────────────────────────────────────────────────────────

# Patrón de código estructurado: ALL-FAIL-000, OPAL-BATT-000, CT-PCBA-002, etc.
_CODE_RE = re.compile(r'^[A-Z0-9]+-[A-Z0-9]+-\d+$')

# Patrón de metadata de dispositivo: QR=..., MAC=..., UID=..., etc.
_METADATA_RE = re.compile(r'^[A-Z_]{1,6}=')


def _clean(text):
    """Elimina detalles específicos de dispositivo entre paréntesis/corchetes."""
    text = re.split(r'[\(\[]', text)[0].strip().rstrip(',').strip()
    return text


def extract_tipo_falla(failure_code, notes_raw, station_name="", producto=""):
    fc      = str(failure_code).strip() if failure_code and str(failure_code) not in ("nan", "None") else ""
    station = str(station_name).strip() if station_name and str(station_name) not in ("nan", "None") else ""
    ns      = str(notes_raw).strip()    if notes_raw    and str(notes_raw)    not in ("nan", "None", "") else ""

    # Códigos de pase — devolver tal cual
    if "ALL-PASS" in fc or "-PASS-" in fc or fc == "MAXPASSEXCEEDED":
        return fc

    # CT- son descriptivos por sí solos (C-Track)
    if fc.startswith("CT-"):
        return fc

    if ns and ns != "undefined":

        if '"fail_tests": ["Sleep Current", "Non_FEM Current", "FEM Current"]' in ns:
            return "No current measured"

        # ── Búsqueda de cadena estilo fórmula Excel ────────────
        # Igual que: MID(I2, FIND("""fail_tests"": [""", I2)+16, FIND(...)-(FIND(...)+16))
        # Busca el marcador en el texto crudo y extrae hasta el siguiente "
        # Se intentan variantes con y sin espacio por si el JSON viene compacto
        for marker in ('"fail_tests": ["', '"fail_tests":["', '"failed": ["', '"failed":["'):
            idx = ns.find(marker)
            if idx >= 0:
                start = idx + len(marker)
                end   = ns.find('"', start)
                if end > start:
                    val = _clean(ns[start:end].strip())
                    if val:
                        return val

        # ── JSON para casos especiales ─────────────────────────
        if ns.startswith("{"):
            try:
                parsed = json.loads(ns)

                # OPAL4: wrapper {"notes": {...}}
                if "notes" in parsed and isinstance(parsed["notes"], dict):
                    inner   = parsed["notes"]
                    fc_list = inner.get("failure_codes", [])
                    fc0     = str(fc_list[0]).strip() if fc_list else ""
                    if fc0 and not _CODE_RE.match(fc0):
                        return fc0
                    return station

                # BlackBox2 RF Chamber: {"station": "... - RF Chamber BLE Test"}
                if "station" in parsed:
                    parts = str(parsed["station"]).split(" - ")
                    return parts[-1].strip() if len(parts) > 1 else parts[0].strip()

                # failure_codes[0] como último recurso JSON
                fc_list = parsed.get("failure_codes", [])
                if fc_list:
                    return str(fc_list[0]).strip()

            except (json.JSONDecodeError, TypeError):
                pass

        # ── Texto plano ────────────────────────────────────────

        # Solar Panel Test: "FAIL - V:2.349V I:0.915A P:2.149W" → nombre de estación
        if ns.startswith("FAIL - "):
            return station if station else fc

        # Black Box: "MAC, Sleep Current | sleep: 0.53 | bat: 3.1"
        if "Black Box" in producto:
            left   = ns.split(" | ")[0].strip()
            first  = left.split(",")[0].strip()
            result = _clean(first.split(":")[0].strip())
            if result and not _METADATA_RE.match(result):
                return result

        # SPSF y otros: primera línea no vacía, split por " | " y luego ":"
        # "MCU Failed:SOM Could Not..."           → "MCU Failed"
        # "\r\nLogin incorrect\r\n..."            → "Login incorrect"
        # "Cell | charge: 100%"                  → "Cell"
        # "404 - No transmission | FAILURE_CODES" → "404 - No transmission"
        lines = [
            l.strip() for l in ns.replace("\r", "\n").split("\n")
            if l.strip() and l.strip().lower() not in ("undefined", "fail")
        ]
        if lines:
            segment = lines[0].split(" | ")[0].strip()
            result  = _clean(segment.split(":")[0].strip())
            if result and not _METADATA_RE.match(result):
                return result

    # Sin notas útiles → failureCode
    return fc

# ─────────────────────────────────────────────────────────────
# DB Y PROCESAMIENTO
# ─────────────────────────────────────────────────────────────

GLOBAL_CACHE = {
    "df": pd.DataFrame(),
    "last_updated": None
}

def _get_odbc_driver():
    """Detecta automaticamente ODBC Driver 17 o 18 segun lo que este instalado."""
    drivers = pyodbc.drivers()
    for drv in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
        if drv in drivers:
            return drv
    return "ODBC Driver 17 for SQL Server"  # fallback

def connect():
    conn_str = (
        f"DRIVER={{{_get_odbc_driver()}}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASS};"
        f"Encrypt=yes;TrustServerCertificate=no;"
    )
    return pyodbc.connect(conn_str, timeout=DB_CONNECT_TIMEOUT_SECONDS)

def connect_qr():
    """Conexion a la BD de tracking de QR (trkprdshipapp)."""
    conn_str = (
        f"DRIVER={{{_get_odbc_driver()}}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_NAME_QR};"
        f"UID={DB_USER};"
        f"PWD={DB_PASS};"
        f"Encrypt=yes;TrustServerCertificate=no;"
    )
    return pyodbc.connect(conn_str, timeout=DB_CONNECT_TIMEOUT_SECONDS)

def lookup_model_numbers(qr_codes):
    """
    Consulta trkprdshipapp.trk.qrmac_db y devuelve {qrcode: modelNumber}
    para los QR codes proporcionados. Usa parametros seguros en lotes de 500.
    Si falla la conexion, devuelve {} y los registros quedan en "SPSF" (fallback).
    """
    unique_qrs = [
        str(q) for q in qr_codes
        if q and str(q) not in ("nan", "None", "")
    ]
    if not unique_qrs:
        return {}

    result = {}
    batch_size = 500
    try:
        conn = connect_qr()
        for i in range(0, len(unique_qrs), batch_size):
            chunk = unique_qrs[i : i + batch_size]
            placeholders = ",".join("?" for _ in chunk)
            sql = f"SELECT qrcode, modelNumber FROM trk.qrmac_db WHERE qrcode IN ({placeholders})"
            for row in conn.execute(sql, chunk).fetchall():
                if row[1] is not None:
                    result[str(row[0])] = str(row[1]).strip()
        conn.close()
    except Exception as e:
        print(f"[WARNING] lookup_model_numbers: no se pudo consultar {DB_NAME_QR}: {e}")

    return result

def prepare_dataframe(df):
    """Pre-calcula columnas pesadas para evitar demoras en los endpoints."""
    if df.empty:
        return df
    if not pd.api.types.is_datetime64_any_dtype(df['endtime']):
        df['endtime'] = pd.to_datetime(df['endtime'])

    # Asignacion inicial por estacion (comportamiento base para productos no-compartidos)
    df["producto"] = df["stationid"].map(STATION_TO_GROUP).fillna("Other")

    # ── Resolver SPSF vs Blade para estaciones compartidas ──────────────────
    shared_mask = df["stationid"].isin(SPSF_BLADE_STATIONS)
    if shared_mask.any():
        # Consultar modelNumber en trkprdshipapp para todos los QR's involucrados
        shared_qrs = df.loc[shared_mask, "currQr"].dropna().unique().tolist()
        model_map  = lookup_model_numbers(shared_qrs)

        # Estaciones diferenciables: sobreescribir producto por modelNumber
        # Fallback a "SPSF" si el QR no existe en trkprdshipapp o el modelNumber es desconocido
        diff_mask = shared_mask & df["stationid"].isin(SPSF_BLADE_DIFF_STATIONS)
        if diff_mask.any():
            df.loc[diff_mask, "producto"] = df.loc[diff_mask, "currQr"].map(
                lambda qr: MODEL_NUMBER_TO_PRODUCT.get(model_map.get(str(qr), ""), "SPSF")
            )

        # Estaciones NO diferenciables (406, 393):
        # Duplicar filas → original queda como "SPSF", copia queda como "Blade"
        undiff_mask = shared_mask & df["stationid"].isin(UNDIFF_STATIONS)
        if undiff_mask.any():
            df.loc[undiff_mask, "producto"] = "SPSF"
            blade_rows = df.loc[undiff_mask].copy()
            blade_rows["producto"] = "Blade"
            df = pd.concat([df, blade_rows], ignore_index=True)

    # Pre-calcular tipo de falla una sola vez
    df["tipoFalla"] = df.apply(
        lambda r: extract_tipo_falla(r["failureCode"], r["notes"], r["stationName"], r["producto"]), axis=1
    )
    return df

def load_historical_data():
    """Carga 1 año de datos desde la BD al arrancar el servidor."""
    print(" [DATABASE] Iniciando carga de datos historicos (ultimo año)...")
    start_date = (datetime.now() - timedelta(days=HISTORICAL_LOAD_DAYS)).strftime("%Y-%m-%d")
    conn = connect()
    
    query = SQL + f" AND e.endtime >= '{start_date}'"
    df = pd.read_sql(query, conn)
    conn.close()
    
    df = prepare_dataframe(df)
        
    GLOBAL_CACHE["df"] = df
    GLOBAL_CACHE["last_updated"] = datetime.now()
    print(f" [DATABASE] Carga inicial completa. {len(df)} registros alojados en Memoria RAM.")

def background_update_worker():
    """
    Hilo en segundo plano: Cada 15 min solo pide a la base de datos 
    los ultimos 2 dias y los inyecta a la memoria global.
    """
    while True:
        time.sleep(REFRESH_INTERVAL_MINUTES * 60)
        try:
            print("[BACKGROUND] Refrescando registros recientes de DB...")
            cutoff_dt = datetime.now() - timedelta(days=RECENT_REFRESH_DAYS)
            cutoff_str = cutoff_dt.strftime("%Y-%m-%d")
            cutoff_ts  = pd.to_datetime(cutoff_str)   # medianoche, sin hora
            
            conn = connect()
            query = SQL + f" AND e.endtime >= '{cutoff_str}'"
            df_recent = pd.read_sql(query, conn)
            conn.close()
            
            df_recent = prepare_dataframe(df_recent)
                
            current_df = GLOBAL_CACHE["df"]
            
            # Borrar los datos viejos de los ultimos 2 dias del historico
            # Usar cutoff_ts (medianoche) para que coincida con el filtro SQL
            old_data = current_df[current_df['endtime'] < cutoff_ts]
            
            # Unir los datos recien consultados
            new_df = pd.concat([old_data, df_recent], ignore_index=True)
            
            GLOBAL_CACHE["df"] = new_df
            GLOBAL_CACHE["last_updated"] = datetime.now()
            print(f"✅ [BACKGROUND] Tabla en RAM Actualizada. Total registros: {len(new_df)}")
            
        except Exception as e:
            print(f"❌ [BACKGROUND] Error en hilo de actualizacion: {str(e)}")


def filter_real_failures(df):
    """
    Filtra 'fallas reales': para cada currQr, toma el último registro.
    Si ese último registro es FAILED y tiene más de 24h de antigüedad,
    se considera falla real (producto que no pudo ser reprocesado).
    """
    if df.empty:
        return df
    now = datetime.now()
    cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)

    # Último registro por currQr
    idx = df.groupby("currQr")["endtime"].idxmax()
    last_per_qr = df.loc[idx]

    # Solo los que su último intento fue FAILED y tiene >24h
    real = last_per_qr[
        (last_per_qr["resultado"] == "FAILED") &
        (last_per_qr["endtime"] < cutoff)
    ]
    return real


def compute_stats(df, date_from, date_to, real_failures=False):
    now = datetime.now()
    cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)

    def get_kpi_data(data, group_cols):
        if not real_failures:
            return data
        # Unique QRs passed or failed > time
        idx = data.groupby(group_cols + ["currQr"])["endtime"].idxmax()
        last = data.loc[idx]
        return last[
            (last["resultado"] == "PASSED") | 
            ((last["resultado"] == "FAILED") & (last["endtime"] < cutoff))
        ]

    def kpis(d):
        t  = len(d)
        p  = (d["resultado"] == "PASSED").sum()
        f  = (d["resultado"] == "FAILED").sum()
        fp = d["failureCode"].str.contains(r".*-PASS-.*", na=False).sum()
        return {
            "total":  int(t),
            "failed": int(f),
            "passed": int(p),
            "yield":  round(p / t * 100, 1) if t else 0,
            "fpy":    round(fp / t * 100, 1) if t else 0,
        }

    # Data scopeado globalmente
    data_overall = get_kpi_data(df, [])
    overall = kpis(data_overall)
    overall["dateFrom"] = date_from
    overall["dateTo"]   = date_to

    # Data scopeado por producto
    data_prod = get_kpi_data(df, ["producto"])
    by_product = {prod: kpis(grp) for prod, grp in data_prod.groupby("producto")}

    # Data scopeado por estación (y producto referencial)
    data_stat = get_kpi_data(df, ["stationid", "stationName", "producto"])
    stations = []
    for (sid, sname, prod), grp in data_stat.groupby(["stationid", "stationName", "producto"]):
        k = kpis(grp)
        stations.append({"id": int(sid), "name": sname, "producto": prod, **k})
    stations.sort(key=lambda x: STATION_ORDER.get((x["producto"], x["id"]), 9999))

    # Data de fail modes: scopeados por estación, producto y Falla
    if real_failures:
        # Usamos data_stat (que contiene el último intento por QR por estación).
        # Si su estado final fue PASSED, ya no aparecerá como falla.
        data_f = data_stat
    else:
        data_f = df

    df_f = data_f[
        (data_f["resultado"] == "FAILED") &
        (data_f["tipoFalla"] != "") &
        data_f["tipoFalla"].notna()
    ]
    failures = []
    for (sid, sname, prod, falla), grp in df_f.groupby(
        ["stationid", "stationName", "producto", "tipoFalla"]
    ):
        failures.append({
            "sid":      int(sid),
            "sname":    sname,
            "producto": prod,
            "falla":    falla,
            "qty":      len(grp),
        })
    failures.sort(key=lambda x: x["qty"], reverse=True)

    return {
        "overall":    overall,
        "byProduct":  by_product,
        "stations":   stations,
        "failures":   failures[:300],
        "groups":     GROUP_NAMES,
        "lastUpdated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


# ─────────────────────────────────────────────────────────────
# RUTAS
# ─────────────────────────────────────────────────────────────
@app.route("/")
def index():
    return send_from_directory("static", "index.html")

@app.route("/monitor")
def monitor_page():
    return send_from_directory("static", "monitor.html")

@app.route("/api/monitor_data")
def api_monitor_data():
    target_date_str = request.args.get("date", datetime.now().strftime("%Y-%m-%d"))
    product_filter = request.args.get("product", "all")
    
    try:
        df = GLOBAL_CACHE["df"]
        if df.empty:
            return jsonify({"date": target_date_str, "product": product_filter, "error": "System booting", "data": {}})
            
        # Filtro en RAM
        target_dt = pd.to_datetime(target_date_str)
        mask = (df['endtime'] >= target_dt) & (df['endtime'] < target_dt + timedelta(days=1))
        df_day = df[mask].copy()
        
        # Filter by product if specified
        if product_filter != "all" and product_filter in GROUPS:
            df_day = df_day[df_day['producto'] == product_filter]
        
        if df_day.empty:
            return jsonify({"date": target_date_str, "product": product_filter, "data": {}})
            
        def extract_hardware_id(notes_str, station_name):
            try:
                if not notes_str or pd.isna(notes_str): return station_name
                if isinstance(notes_str, str) and notes_str.startswith("{"):
                    data = json.loads(notes_str)
                    hw_id = data.get('hardware_id')
                    if hw_id: return hw_id
                return station_name
            except:
                return station_name

        def determine_group_key(row):
            st_name = str(row['stationName']).strip()
            if st_name == "Jasper MPS":
                return extract_hardware_id(row['notes'], st_name)
            return st_name

        df_day['is_pass'] = (df_day['resultado'] == 'PASSED').astype(int)
        df_day['GroupKey'] = df_day.apply(determine_group_key, axis=1)
        df_day = df_day[df_day['GroupKey'] != "MPS_CTRACK_00"]
        
        response_data = {}
        for group, df_g in df_day.groupby('GroupKey'):
            hourly_counts = [0] * 24
            hourly_sums = [0] * 24
            hourly_fpy = [0.0] * 24
            
            for _, row in df_g.iterrows():
                h = row['endtime'].hour
                hourly_counts[h] += 1
                hourly_sums[h] += row['is_pass']
                
            for h in range(24):
                if hourly_counts[h] > 0:
                    hourly_fpy[h] = (hourly_sums[h] / hourly_counts[h]) * 100
                    
            total_att = int(df_g['is_pass'].count())
            total_pass = int(df_g['is_pass'].sum())
            st_name = str(df_g['stationName'].iloc[0])
            
            response_data[str(group)] = {
                "stationName": st_name,
                "hourly": {
                    "fpy": hourly_fpy,
                    "count": hourly_counts,
                    "sum": hourly_sums
                },
                "total_att": total_att,
                "total_pass": total_pass,
                "global_fpy": (total_pass / total_att * 100) if total_att > 0 else 0
            }
            
        return jsonify({"date": target_date_str, "product": product_filter, "data": response_data})
        
    except pyodbc.Error as e:
        return jsonify({"error": f"Error de conexion a DB: {str(e)}"}), 500
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/debug/<grupo>")
def api_debug(grupo):
    """Muestra notas crudas de filas FAILED para un grupo dado. Solo para desarrollo."""
    default_from = (datetime.now() - timedelta(days=DEBUG_ENDPOINT_RANGE_DAYS)).strftime("%Y-%m-%d")
    default_to   = datetime.now().strftime("%Y-%m-%d")
    ids = GROUPS.get(grupo, [])
    if not ids:
        return jsonify({"error": f"Grupo '{grupo}' no existe. Opciones: {GROUP_NAMES}"}), 404
    ids_str = ",".join(str(i) for i in ids)
    sql = f"""
        SELECT TOP 20
            s.stationName, e.failureCode, e.notes
        FROM trk.manufacturingEvents AS e
        INNER JOIN trk.stationConfig  AS s ON e.stationid = s.stationid
        WHERE s.stationid IN ({ids_str})
          AND e.endtime >= ?
          AND e.failureCode NOT LIKE 'ALL-PASS%%'
          AND e.failureCode NOT LIKE '%%-PASS-%%'
          AND e.failureCode <> 'MAXPASSEXCEEDED'
          AND e.notes IS NOT NULL
        ORDER BY e.endtime DESC
    """
    try:
        conn = connect()
        rows = conn.execute(sql, [default_from]).fetchall()
        conn.close()
        result = [
            {
                "station":    r[0],
                "failureCode": r[1],
                "tipoFalla":  extract_tipo_falla(r[1], r[2], r[0], grupo),
                "notes":      r[2],
            }
            for r in rows
        ]
        return jsonify(result)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/debug_station/<int:sid>")
def api_debug_station(sid):
    """
    Diagnostico para una estacion especifica.
    Consulta la BD directamente sin filtro de fecha para revelar si hay datos
    y cuando fue el ultimo registro. Util para estaciones que no aparecen en el dashboard.
    Ejemplo: /api/debug_station/406
    """
    sql = """
        SELECT TOP 10
            e.endtime,
            s.stationName,
            c.currQr,
            c.prevQr,
            e.failureCode,
            CASE
                WHEN e.failureCode LIKE 'ALL-PASS%'   THEN 'PASSED'
                WHEN e.failureCode LIKE '%-PASS-%'    THEN 'PASSED'
                WHEN e.failureCode = 'MAXPASSEXCEEDED' THEN 'PASSED'
                ELSE 'FAILED'
            END AS resultado
        FROM trk.manufacturingEvents AS e
        INNER JOIN trk.qrCorrelation AS c ON e.correlationID = c.correlationID
        INNER JOIN trk.stationConfig  AS s ON e.stationid    = s.stationid
        WHERE s.stationid = ?
        ORDER BY e.endtime DESC
    """
    try:
        conn = connect()
        rows = conn.execute(sql, [sid]).fetchall()
        conn.close()

        if not rows:
            return jsonify({
                "stationid": sid,
                "warning": "Sin registros en la BD (con INNER JOIN qrCorrelation). "
                           "Puede que los eventos de esta estacion no tengan correlationID valido.",
                "records": []
            })

        days_since_last = (datetime.now() - rows[0][0]).days if rows[0][0] else None
        in_cache = days_since_last is not None and days_since_last <= HISTORICAL_LOAD_DAYS

        records = [
            {
                "endtime":     r[0].strftime("%Y-%m-%d %H:%M") if r[0] else None,
                "stationName": r[1],
                "currQr":      str(r[2]) if r[2] is not None else "NULL",
                "prevQr":      str(r[3]) if r[3] is not None else "NULL",
                "failureCode": r[4],
                "resultado":   r[5],
            }
            for r in rows
        ]
        return jsonify({
            "stationid":       sid,
            "days_since_last": days_since_last,
            "in_cache_window": in_cache,
            "cache_window_days": HISTORICAL_LOAD_DAYS,
            "diagnosis": (
                "OK - datos dentro del rango del cache" if in_cache
                else f"FUERA DEL CACHE: ultimo registro hace {days_since_last} dias, pero HISTORICAL_LOAD_DAYS={HISTORICAL_LOAD_DAYS}"
            ),
            "records": records,
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/failure_details")
def api_failure_details():
    """Devuelve registros individuales de fallas filtrados por producto y tipo de falla."""
    date_from_str = request.args.get("from", (datetime.now() - timedelta(days=DEFAULT_QUERY_RANGE_DAYS)).strftime("%Y-%m-%d"))
    date_to_str   = request.args.get("to",   datetime.now().strftime("%Y-%m-%d"))
    product       = request.args.get("product", "all")
    station_id_str = request.args.get("stationid", "all")
    falla         = request.args.get("falla", "")
    real_failures = request.args.get("real_failures", "false").lower() == "true"
    export_csv    = request.args.get("csv", "false").lower() == "true"

    try:
        df = GLOBAL_CACHE["df"]
        if df.empty:
            return jsonify({"error": "System is booting"}), 503

        dt_from = pd.to_datetime(date_from_str)
        dt_to   = pd.to_datetime(date_to_str) + timedelta(days=1)

        mask = (df['endtime'] >= dt_from) & (df['endtime'] < dt_to)
        sub = df[mask].copy()

        if product != "all":
            sub = sub[sub["producto"] == product]

        if station_id_str != "all":
            try:
                station_id = int(station_id_str)
                sub = sub[sub["stationid"] == station_id]
            except ValueError:
                return jsonify({"error": "stationid invalido"}), 400

        if real_failures:
            # Obtener el último estatus (sea PASS o FAIL) de cada QR en este cruce de est/prod
            now = datetime.now()
            cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)
            idx = sub.groupby("currQr")["endtime"].idxmax()
            last_per_qr = sub.loc[idx]
            # De ese último estatus, quedarse solo con los que hayan fallado hace más de X horas
            sub = last_per_qr[
                (last_per_qr["resultado"] == "FAILED") & 
                (last_per_qr["endtime"] < cutoff)
            ]
        else:
            # En modo completo, solo registros FAILED
            sub = sub[sub["resultado"] == "FAILED"]

        if falla:
            sub = sub[sub["tipoFalla"] == falla]

        if export_csv:
            from flask import Response
            # Aseguramos columnas útiles para exportar
            cols = ["endtime", "currQr", "prevQr", "stationName", "failureCode", "tipoFalla", "resultado", "notes"]
            csv_data = sub[[c for c in cols if c in sub.columns]].to_csv(index=False)
            return Response(
                csv_data,
                mimetype="text/csv",
                headers={"Content-disposition": "attachment; filename=failure_details.csv"}
            )

        # Limitar cantidad de registros de respuesta
        sub = sub.head(DETAILS_MAX_RECORDS)

        records = []
        for _, r in sub.iterrows():
            records.append({
                "prevQr":      str(r.get("prevQr", "") or ""),
                "currQr":      str(r.get("currQr", "") or ""),
                "tipoFalla":   str(r.get("tipoFalla", "")),
                "failureCode": str(r.get("failureCode", "")),
                "stationName": str(r.get("stationName", "")),
                "notes":       str(r.get("notes", "") or ""),
            })

        return jsonify(records)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/passfail_details")
def api_passfail_details():
    """Devuelve registros del último intento por QR, filtrados por resultado (PASSED/FAILED)."""
    date_from_str = request.args.get("from", (datetime.now() - timedelta(days=DEFAULT_QUERY_RANGE_DAYS)).strftime("%Y-%m-%d"))
    date_to_str   = request.args.get("to",   datetime.now().strftime("%Y-%m-%d"))
    product       = request.args.get("product", "all")
    station_id_str = request.args.get("stationid", "all")
    resultado     = request.args.get("resultado", "")  # PASSED o FAILED
    real_failures = request.args.get("real_failures", "false").lower() == "true"
    export_csv    = request.args.get("csv", "false").lower() == "true"

    try:
        df = GLOBAL_CACHE["df"]
        if df.empty:
            return jsonify({"error": "System is booting"}), 503

        dt_from = pd.to_datetime(date_from_str)
        dt_to   = pd.to_datetime(date_to_str) + timedelta(days=1)

        mask = (df['endtime'] >= dt_from) & (df['endtime'] < dt_to)
        sub = df[mask].copy()

        if product != "all":
            sub = sub[sub["producto"] == product]

        if station_id_str != "all":
            try:
                station_id = int(station_id_str)
                sub = sub[sub["stationid"] == station_id]
            except ValueError:
                return jsonify({"error": "stationid invalido"}), 400

        if sub.empty:
            return jsonify([])

        if real_failures:
            # Último registro por currQr
            now = datetime.now()
            cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)
            idx = sub.groupby("currQr")["endtime"].idxmax()
            last_per_qr = sub.loc[idx]
            # Solo los relevantes (passed o failed >24h)
            sub = last_per_qr[
                (last_per_qr["resultado"] == "PASSED") |
                ((last_per_qr["resultado"] == "FAILED") & (last_per_qr["endtime"] < cutoff))
            ]

        if resultado:
            sub = sub[sub["resultado"] == resultado]

        sub = sub.sort_values("endtime", ascending=False)

        if export_csv:
            from flask import Response
            cols = ["endtime", "currQr", "prevQr", "stationName", "failureCode", "tipoFalla", "resultado", "notes"]
            csv_data = sub[[c for c in cols if c in sub.columns]].to_csv(index=False)
            return Response(
                csv_data,
                mimetype="text/csv",
                headers={"Content-disposition": f"attachment; filename=passfail_{resultado}.csv"}
            )

        sub = sub.head(DETAILS_MAX_RECORDS)

        records = []
        for _, r in sub.iterrows():
            records.append({
                "prevQr":      str(r.get("prevQr", "") or ""),
                "currQr":      str(r.get("currQr", "") or ""),
                "resultado":   str(r.get("resultado", "")),
                "tipoFalla":   str(r.get("tipoFalla", "")),
                "failureCode": str(r.get("failureCode", "")),
                "stationName": str(r.get("stationName", "")),
                "notes":       str(r.get("notes", "") or ""),
                "endtime":     r["endtime"].strftime("%Y-%m-%d %H:%M") if pd.notna(r.get("endtime")) else "",
            })

        return jsonify(records)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/export_all_csv")
def api_export_all_csv():
    """Descarga de CSV completo para todas las pruebas filtradas por los controles actuales."""
    date_from_str = request.args.get("from", (datetime.now() - timedelta(days=DEFAULT_QUERY_RANGE_DAYS)).strftime("%Y-%m-%d"))
    date_to_str   = request.args.get("to",   datetime.now().strftime("%Y-%m-%d"))
    product       = request.args.get("product", "all")
    station_id_str = request.args.get("stationid", "all")
    real_failures = request.args.get("real_failures", "false").lower() == "true"

    try:
        df = GLOBAL_CACHE["df"]
        if df.empty:
            return "System is booting", 503

        dt_from = pd.to_datetime(date_from_str)
        dt_to   = pd.to_datetime(date_to_str) + timedelta(days=1)

        mask = (df['endtime'] >= dt_from) & (df['endtime'] < dt_to)
        sub = df[mask].copy()

        if product != "all":
            sub = sub[sub["producto"] == product]

        if station_id_str != "all":
            try:
                station_id = int(station_id_str)
                sub = sub[sub["stationid"] == station_id]
            except ValueError:
                return "stationid invalido", 400

        if sub.empty:
            return "No data found", 404

        if real_failures:
            now = datetime.now()
            cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)
            idx = sub.groupby("currQr")["endtime"].idxmax()
            last_per_qr = sub.loc[idx]
            sub = last_per_qr[
                (last_per_qr["resultado"] == "PASSED") |
                ((last_per_qr["resultado"] == "FAILED") & (last_per_qr["endtime"] < cutoff))
            ]

        sub = sub.sort_values("endtime", ascending=False)
        
        from flask import Response
        cols = ["endtime", "currQr", "prevQr", "stationName", "producto", "failureCode", "tipoFalla", "resultado", "notes"]
        csv_data = sub[[c for c in cols if c in sub.columns]].to_csv(index=False)
        
        return Response(
            csv_data,
            mimetype="text/csv",
            headers={"Content-disposition": "attachment; filename=all_tests_export.csv"}
        )
    except Exception as e:
        return str(e), 500


@app.route("/api/data")
def api_data():
    date_from_str = request.args.get("from", (datetime.now() - timedelta(days=DEFAULT_QUERY_RANGE_DAYS)).strftime("%Y-%m-%d"))
    date_to_str   = request.args.get("to",   datetime.now().strftime("%Y-%m-%d"))
    real_failures = request.args.get("real_failures", "false").lower() == "true"

    try:
        df = GLOBAL_CACHE["df"]
        if df.empty:
            return jsonify({"error": "System is booting up database into RAM. Try again in a minute.", "cacheBooting": True}), 503
            
        dt_from = pd.to_datetime(date_from_str)
        dt_to   = pd.to_datetime(date_to_str) + timedelta(days=1)
        
        # Filtro en RAM 
        mask = (df['endtime'] >= dt_from) & (df['endtime'] < dt_to)
        df_filtered = df[mask].copy()

        # Calcular pass/fail pie con desglose por producto y estación
        if real_failures and not df_filtered.empty:
            now = datetime.now()
            cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)

            def get_relevant(sub_df, by):
                if not by:
                    idx = sub_df.groupby("currQr")["endtime"].idxmax()
                else:
                    idx = sub_df.groupby(by + ["currQr"])["endtime"].idxmax()
                last = sub_df.loc[idx]
                return last[
                    (last["resultado"] == "PASSED") |
                    ((last["resultado"] == "FAILED") & (last["endtime"] < cutoff))
                ]

            relevant_overall = get_relevant(df_filtered, [])
            relevant_product = get_relevant(df_filtered, ["producto"])
            relevant_station = get_relevant(df_filtered, ["stationid"])

            def _pf(d):
                p = int((d["resultado"] == "PASSED").sum())
                f = int((d["resultado"] == "FAILED").sum())
                t = p + f
                return {
                    "passed": p,
                    "failed": f,
                    "total":  t,
                    "yield":  round(p / t * 100, 1) if t else 0,
                }

            pf_overall = _pf(relevant_overall)
            
            pf_by_product = {}
            for prod, grp in relevant_product.groupby("producto"):
                pf_by_product[prod] = _pf(grp)

            pf_by_station = {}
            for sid, grp in relevant_station.groupby("stationid"):
                pf_by_station[int(sid)] = _pf(grp)

            pass_fail_pie = {
                "overall": pf_overall,
                "byProduct": pf_by_product,
                "byStation": pf_by_station,
            }

            # Ya no filtramos df_filtered porque se enviará entero a compute_stats, 
            # y compute_stats hará el filtrado por sus propios axis.
        else:
            pass_fail_pie = None

        stats = compute_stats(df_filtered, date_from_str, date_to_str, real_failures=real_failures)
        stats["realFailures"] = real_failures
        if pass_fail_pie:
            stats["passFailPie"] = pass_fail_pie
        if GLOBAL_CACHE["last_updated"]:
            stats["cacheUpdated"] = GLOBAL_CACHE["last_updated"].strftime("%Y-%m-%d %H:%M:%S")
        return jsonify(stats)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    # Iniciar la carga fuerte de 1 año al arrancar el py
    load_historical_data()
    
    # Iniciar el recolector de 2do plano cada 15 min
    bg_thread = threading.Thread(target=background_update_worker, daemon=True)
    bg_thread.start()
    
    port = int(os.environ.get("PORT", DEFAULT_APP_PORT))
    print(f"\n  Manufacturing Dashboard iniciado")
    print(f"  Abre en tu navegador: http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
