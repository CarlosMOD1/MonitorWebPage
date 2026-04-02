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
DB_SERVER = os.getenv("DB_SERVER")
DB_NAME   = os.getenv("DB_NAME")
DB_USER   = os.getenv("DB_USER")
DB_PASS   = os.getenv("DB_PASS")

GROUPS = {
    "SPSF":        [178,183,184,185,186,187,244,201,202,190,191,192,193,194,195,196,199,208,211,215,220,232],
    "C-Track":     [238,239,241,333],
    "Solar":       [275,277,278,279,280,281,284],
    "Black Box":   [203,204,233,237,245],
    "OPAL 3":      [260,261,262,263,264,265,266],
    "Black Box 2": [297,298,300,301,302,303,304,305,306],
    "Opal 4":      [289,290,291,292,293,294,295],
}

STATION_TO_GROUP = {}
ALL_STATION_IDS  = []
for grp, ids in GROUPS.items():
    for sid in ids:
        STATION_TO_GROUP[sid] = grp
        ALL_STATION_IDS.append(sid)

IDS_STR = ",".join(str(s) for s in ALL_STATION_IDS)
GROUP_NAMES = list(GROUPS.keys())

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

def connect():
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASS};"
        f"Encrypt=yes;TrustServerCertificate=no;"
    )
    return pyodbc.connect(conn_str, timeout=60)

def load_historical_data():
    """Carga 1 año de datos desde la BD al arrancar el servidor."""
    print("⏳ [DATABASE] Iniciando carga de datos historicos (ultimo año)...")
    start_date = (datetime.now() - timedelta(days=365)).strftime("%Y-%m-%d")
    conn = connect()
    
    query = SQL + f" AND e.endtime >= '{start_date}'"
    df = pd.read_sql(query, conn)
    conn.close()
    
    if not pd.api.types.is_datetime64_any_dtype(df['endtime']):
        df['endtime'] = pd.to_datetime(df['endtime'])
        
    GLOBAL_CACHE["df"] = df
    GLOBAL_CACHE["last_updated"] = datetime.now()
    print(f"✅ [DATABASE] Carga inicial completa. {len(df)} registros alojados en Memoria RAM.")

def background_update_worker():
    """
    Hilo en segundo plano: Cada 15 min solo pide a la base de datos 
    los ultimos 2 dias y los inyecta a la memoria global.
    """
    while True:
        time.sleep(15 * 60) # 15 min de espera
        try:
            print("🔄 [BACKGROUND] Refrescando registros recientes de DB...")
            cutoff_dt = datetime.now() - timedelta(days=2)
            cutoff_str = cutoff_dt.strftime("%Y-%m-%d")
            
            conn = connect()
            query = SQL + f" AND e.endtime >= '{cutoff_str}'"
            df_recent = pd.read_sql(query, conn)
            conn.close()
            
            if not df_recent.empty and not pd.api.types.is_datetime64_any_dtype(df_recent['endtime']):
                df_recent['endtime'] = pd.to_datetime(df_recent['endtime'])
                
            current_df = GLOBAL_CACHE["df"]
            
            # Borrar los datos viejos de los ultimos 2 dias del historico
            old_data = current_df[current_df['endtime'] < cutoff_dt]
            
            # Unir los datos recien consultados
            new_df = pd.concat([old_data, df_recent], ignore_index=True)
            
            GLOBAL_CACHE["df"] = new_df
            GLOBAL_CACHE["last_updated"] = datetime.now()
            print(f"✅ [BACKGROUND] Tabla en RAM Actualizada. Total registros: {len(new_df)}")
            
        except Exception as e:
            print(f"❌ [BACKGROUND] Error en hilo de actualizacion: {str(e)}")


def compute_stats(df, date_from, date_to):
    df["producto"]  = df["stationid"].map(STATION_TO_GROUP).fillna("Other")
    df["tipoFalla"] = df.apply(
        lambda r: extract_tipo_falla(r["failureCode"], r["notes"], r["stationName"], r["producto"]), axis=1
    )

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

    overall = kpis(df)
    overall["dateFrom"] = date_from
    overall["dateTo"]   = date_to

    by_product = {prod: kpis(grp) for prod, grp in df.groupby("producto")}

    stations = []
    for (sid, sname, prod), grp in df.groupby(["stationid", "stationName", "producto"]):
        k = kpis(grp)
        stations.append({"id": int(sid), "name": sname, "producto": prod, **k})
    stations.sort(key=lambda x: x["failed"], reverse=True)

    df_f = df[
        (df["resultado"] == "FAILED") &
        (df["tipoFalla"] != "") &
        df["tipoFalla"].notna()
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
            group_ids = GROUPS[product_filter]
            df_day = df_day[df_day['stationid'].isin(group_ids)]
        
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
    default_from = (datetime.now() - timedelta(days=7)).strftime("%Y-%m-%d")
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


@app.route("/api/failure_details")
def api_failure_details():
    """Devuelve registros individuales de fallas filtrados por producto y tipo de falla."""
    date_from_str = request.args.get("from", (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"))
    date_to_str   = request.args.get("to",   datetime.now().strftime("%Y-%m-%d"))
    product       = request.args.get("product", "all")
    falla         = request.args.get("falla", "")

    try:
        df = GLOBAL_CACHE["df"]
        if df.empty:
            return jsonify({"error": "System is booting"}), 503

        dt_from = pd.to_datetime(date_from_str)
        dt_to   = pd.to_datetime(date_to_str) + timedelta(days=1)

        mask = (df['endtime'] >= dt_from) & (df['endtime'] < dt_to)
        sub = df[mask].copy()

        sub["producto"] = sub["stationid"].map(STATION_TO_GROUP).fillna("Other")
        if product != "all":
            sub = sub[sub["producto"] == product]

        # Solo registros FAILED
        sub = sub[sub["resultado"] == "FAILED"]

        # Calcular tipo de falla
        sub["tipoFalla"] = sub.apply(
            lambda r: extract_tipo_falla(r["failureCode"], r["notes"], r["stationName"], r["producto"]), axis=1
        )

        if falla:
            sub = sub[sub["tipoFalla"] == falla]

        # Limitar a 500 registros
        sub = sub.head(500)

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


@app.route("/api/data")
def api_data():
    date_from_str = request.args.get("from", (datetime.now() - timedelta(days=30)).strftime("%Y-%m-%d"))
    date_to_str   = request.args.get("to",   datetime.now().strftime("%Y-%m-%d"))

    try:
        df = GLOBAL_CACHE["df"]
        if df.empty:
            return jsonify({"error": "System is booting up database into RAM. Try again in a minute.", "cacheBooting": True}), 503
            
        dt_from = pd.to_datetime(date_from_str)
        dt_to   = pd.to_datetime(date_to_str) + timedelta(days=1)
        
        # Filtro en RAM 
        mask = (df['endtime'] >= dt_from) & (df['endtime'] < dt_to)
        df_filtered = df[mask].copy()
        
        stats = compute_stats(df_filtered, date_from_str, date_to_str)
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
    
    port = int(os.environ.get("PORT", 5000))
    print(f"\n  Manufacturing Dashboard iniciado")
    print(f"  Abre en tu navegador: http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
