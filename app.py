#!/usr/bin/env python3
"""
Manufacturing Dashboard - Live Web App
Ejecutar: python app.py
Abrir en navegador: http://localhost:5000
"""
from flask import Flask, jsonify, request, send_from_directory
import pyodbc
import pandas as pd
from scipy import stats as scipy_stats
import json
import re
import logging
import threading
import time
from datetime import datetime, timedelta
import os
from dotenv import load_dotenv

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

app = Flask(__name__, static_folder="static")

# ─────────────────────────────────────────────────────────────
# CONFIGURACION
# ─────────────────────────────────────────────────────────────
DB_SERVER = os.getenv("DB_SERVER")
DB_NAME   = os.getenv("DB_NAME")
DB_USER   = os.getenv("DB_USER")
DB_PASS   = os.getenv("DB_PASS")

# Constantes de comportamiento del dashboard
HISTORICAL_LOAD_DAYS = 30
RECENT_REFRESH_DAYS = 2
REFRESH_INTERVAL_MINUTES = 15
REAL_FAILURE_HOURS = 3
DEFAULT_QUERY_RANGE_DAYS = 30
DEBUG_ENDPOINT_RANGE_DAYS = 7
DETAILS_MAX_RECORDS = 500
DB_CONNECT_TIMEOUT_SECONDS = 60
DEFAULT_APP_PORT = 5000
DB_RETRY_ATTEMPTS = 3
DB_RETRY_BASE_DELAY_SECONDS = 30

# Anomaly detection
ANOMALY_LOOKBACK_DAYS    = 7     # baseline window
ANOMALY_ZSCORE_THRESHOLD = 3.0   # sigma below mean to flag a drop (increased from 2.0 to avoid false positives)
ANOMALY_MIN_SAMPLES      = 12    # minimum active hours needed to compute baseline
ANOMALY_STREAK_MIN_HOURS    = 4  # consecutive declining hours to flag a streak (increased from 3)
ANOMALY_MAX_STREAK_GAP_HOURS = 2 # idle gap (h) that breaks a streak — hours further apart are not "consecutive"
ANOMALY_WORK_HOUR_START     = 7  # ignore hours 00:00–06:59 (no production)
ANOMALY_MIN_TESTS_PER_HOUR  = 10  # min tests in a bucket to be analysed (filters sparse/startup hours)

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

    # Notes were present but no pattern matched — log so format changes are visible.
    if ns:
        logger.debug(
            "extract_tipo_falla: unrecognized notes format for station=%r fc=%r notes_prefix=%r",
            station, fc, ns[:120],
        )
    return fc

# ─────────────────────────────────────────────────────────────
# DB Y PROCESAMIENTO
# ─────────────────────────────────────────────────────────────

GLOBAL_CACHE = {
    "df": pd.DataFrame(),
    "last_updated": None,
}

# Keyed by (date_from, date_to, real_failures, cache_timestamp); cleared on every successful DB refresh.
_STATS_CACHE: dict = {}

# Populated by compute_anomalies() after each successful DB refresh.
_ANOMALY_CACHE: dict = {"alerts": [], "computed_at": None}

def connect():
    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={DB_NAME};"
        f"UID={DB_USER};"
        f"PWD={DB_PASS};"
        f"Encrypt=yes;TrustServerCertificate=no;"
    )
    return pyodbc.connect(conn_str, timeout=DB_CONNECT_TIMEOUT_SECONDS)

def prepare_dataframe(df):
    """Pre-calcula columnas pesadas para evitar demoras en los endpoints."""
    if df.empty:
        return df
    if not pd.api.types.is_datetime64_any_dtype(df['endtime']):
        df['endtime'] = pd.to_datetime(df['endtime'])
    
    # Pre-calcular producto y tipo de falla una sola vez
    df["producto"]  = df["stationid"].map(STATION_TO_GROUP).fillna("Other")
    df["tipoFalla"] = df.apply(
        lambda r: extract_tipo_falla(r["failureCode"], r["notes"], r["stationName"], r["producto"]), axis=1
    )
    return df

def load_historical_data():
    """Load the last HISTORICAL_LOAD_DAYS of data into the in-memory cache at startup."""
    logger.info("Loading historical data (last %d days)...", HISTORICAL_LOAD_DAYS)
    start_date = (datetime.now() - timedelta(days=HISTORICAL_LOAD_DAYS)).strftime("%Y-%m-%d")
    query = SQL + f" AND e.endtime >= '{start_date}'"
    delay = DB_RETRY_BASE_DELAY_SECONDS
    for attempt in range(DB_RETRY_ATTEMPTS):
        try:
            conn = connect()
            df = pd.read_sql(query, conn)
            conn.close()
            df = prepare_dataframe(df)
            GLOBAL_CACHE["df"] = df
            GLOBAL_CACHE["last_updated"] = datetime.now()
            logger.info("Initial load complete: %d records in memory.", len(df))
            _refresh_anomalies(df, "startup")
            return
        except Exception as exc:
            logger.error(
                "Historical load failed (attempt %d/%d): %s",
                attempt + 1, DB_RETRY_ATTEMPTS, exc,
            )
            if attempt < DB_RETRY_ATTEMPTS - 1:
                logger.info("Retrying in %ds...", delay)
                time.sleep(delay)
                delay = min(delay * 2, 300)
    logger.critical(
        "All %d historical load attempts failed. Server will start with an empty cache.",
        DB_RETRY_ATTEMPTS,
    )

def _run_background_update():
    """Execute one DB refresh cycle with exponential-backoff retries."""
    delay = DB_RETRY_BASE_DELAY_SECONDS
    for attempt in range(DB_RETRY_ATTEMPTS):
        try:
            logger.info("[BACKGROUND] Refreshing recent records from DB (attempt %d/%d)...",
                        attempt + 1, DB_RETRY_ATTEMPTS)
            cutoff_dt  = datetime.now() - timedelta(days=RECENT_REFRESH_DAYS)
            cutoff_str = cutoff_dt.strftime("%Y-%m-%d")
            cutoff_ts  = pd.to_datetime(cutoff_str)

            conn = connect()
            query = SQL + f" AND e.endtime >= '{cutoff_str}'"
            df_recent = pd.read_sql(query, conn)
            conn.close()

            df_recent  = prepare_dataframe(df_recent)
            current_df = GLOBAL_CACHE["df"]
            old_data   = current_df[current_df["endtime"] < cutoff_ts]
            new_df     = pd.concat([old_data, df_recent], ignore_index=True)

            GLOBAL_CACHE["df"]          = new_df
            GLOBAL_CACHE["last_updated"] = datetime.now()
            _STATS_CACHE.clear()
            logger.info("[BACKGROUND] Cache updated: %d total records.", len(new_df))
            _refresh_anomalies(new_df, "background")
            return
        except Exception as exc:
            logger.error(
                "[BACKGROUND] Update failed (attempt %d/%d): %s",
                attempt + 1, DB_RETRY_ATTEMPTS, exc,
            )
            if attempt < DB_RETRY_ATTEMPTS - 1:
                logger.info("[BACKGROUND] Retrying in %ds...", delay)
                time.sleep(delay)
                delay = min(delay * 2, 300)

    last_updated = GLOBAL_CACHE.get("last_updated")
    if last_updated:
        stale_mins = (datetime.now() - last_updated).total_seconds() / 60
        logger.warning(
            "[BACKGROUND] All retries exhausted — data is %.0f minutes stale.", stale_mins
        )


def background_update_worker():
    """Background thread: refresh the last RECENT_REFRESH_DAYS of data every REFRESH_INTERVAL_MINUTES."""
    while True:
        time.sleep(REFRESH_INTERVAL_MINUTES * 60)
        _run_background_update()


def filter_real_failures(df):
    """
    Filtra 'fallas reales': para cada prevQr, toma el último registro.
    Si ese último registro es FAILED y tiene más de 24h de antigüedad,
    se considera falla real (producto que no pudo ser reprocesado).
    """
    if df.empty:
        return df
    now = datetime.now()
    cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)

    # Último registro por prevQr
    idx = df.groupby("prevQr")["endtime"].idxmax()
    last_per_qr = df.loc[idx]

    # Solo los que su último intento fue FAILED y tiene >24h
    real = last_per_qr[
        (last_per_qr["resultado"] == "FAILED") &
        (last_per_qr["endtime"] < cutoff)
    ]
    return real


def compute_stats(df, date_from, date_to):
    def kpis(d):
        t  = len(d)
        p  = (d["resultado"] == "PASSED").sum()
        f  = (d["resultado"] == "FAILED").sum()
        fp = d["failureCode"].str.contains(r"ALL-PASS|.*-PASS-.*|MAXPASSEXCEEDED", na=False, case=False).sum()
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


def compute_stats_cached(df, date_from: str, date_to: str, real_failures: bool):
    """Return compute_stats() result, reusing the cached value when the cache is still valid."""
    cache_ts = GLOBAL_CACHE.get("last_updated")
    key = (date_from, date_to, real_failures, cache_ts)
    if key in _STATS_CACHE:
        return _STATS_CACHE[key]
    result = compute_stats(df, date_from, date_to)
    _STATS_CACHE[key] = result
    return result


# ─────────────────────────────────────────────────────────────
# ANOMALY DETECTION
# ─────────────────────────────────────────────────────────────

def compute_anomalies(df: pd.DataFrame) -> list:
    """
    Scan every station for two signal types over the past 24 h:

    1. Sudden drop  — hourly FPY falls below (7-day rolling mean − 2σ).
                      scipy.stats.norm.sf gives the one-sided probability so
                      operators can see how improbable the reading is.

    2. Degradation streak — 3+ consecutive hours of monotonically declining FPY,
                            regardless of absolute level.

    Returns a list of alert dicts sorted critical-first then by station name.
    """
    if df.empty:
        return []

    now        = datetime.now()
    cutoff_7d  = now - timedelta(days=ANOMALY_LOOKBACK_DAYS)
    cutoff_24h = now - timedelta(hours=24)
    cutoff_4h  = now - timedelta(hours=4)

    df_week = df[df["endtime"] >= cutoff_7d].copy()
    if df_week.empty:
        return []

    df_week["hour_bucket"] = df_week["endtime"].dt.floor("h")

    alerts = []

    for (sid, sname, prod), sdf in df_week.groupby(
        ["stationid", "stationName", "producto"]
    ):
        # Hourly FPY series for the full 7-day window
        hourly = (
            sdf.groupby("hour_bucket")
            .agg(
                total  =("resultado", "count"),
                passed =("resultado", lambda x: (x == "PASSED").sum()),
            )
            .sort_index()
        )
        hourly["fpy"] = hourly["passed"] / hourly["total"] * 100

        # Drop overnight dead hours (00:00–06:59) and sparse hours (shift start, breaks).
        # Hours with < ANOMALY_MIN_TESTS_PER_HOUR units give unstable FPY (e.g. 0/1 = 0%)
        # that would corrupt the rolling baseline and trigger false drop/streak alerts.
        hourly = hourly[
            (hourly.index.hour >= ANOMALY_WORK_HOUR_START) &
            (hourly["total"] >= ANOMALY_MIN_TESTS_PER_HOUR)
        ]

        if len(hourly) < ANOMALY_MIN_SAMPLES:
            continue

        # Rolling baseline (up to 168 active hours; requires ANOMALY_MIN_SAMPLES to activate)
        roll           = hourly["fpy"].rolling(168, min_periods=ANOMALY_MIN_SAMPLES)
        hourly["mean"] = roll.mean()
        hourly["std"]  = roll.std()

        recent = hourly[hourly.index >= pd.Timestamp(cutoff_24h)]
        if recent.empty:
            continue

        # ── 1. Sudden-drop detection ───────────────────────────
        drop_alerts = []
        for ts, row in recent.iterrows():
            mean, std = row["mean"], row["std"]
            if pd.isna(mean) or pd.isna(std):
                continue  # no baseline yet
            
            # Si una estación es siempre perfecta (100%), std puede ser 0 o casi 0.
            # Imponemos un std artificial mínimo de 1.0 para permitir que saltos graves generen alertas.
            eff_std = max(std, 1.0)
            threshold = mean - ANOMALY_ZSCORE_THRESHOLD * eff_std
            if row["fpy"] < threshold:
                # Solo alertar si la caída sucedió en las últimas 4 horas
                if ts >= pd.Timestamp(cutoff_4h):
                    z     = (mean - row["fpy"]) / eff_std
                    p_val = float(scipy_stats.norm.sf(z))
                    drop_alerts.append({
                        "hour":      ts.strftime("%H:%M"),
                        "fpy":       round(float(row["fpy"]), 1),
                        "threshold": round(float(threshold), 1),
                        "mean":      round(float(mean), 1),
                        "sigma":     round(float(std), 1),
                        "z_score":   round(float(z), 2),
                        "p_value":   round(p_val, 4),
                    })

        # ── 2. Consecutive degradation streak ─────────────────
        streak_alert = None
        fpy_vals = recent["fpy"].values
        ts_vals  = list(recent.index)

        if len(fpy_vals) >= ANOMALY_STREAK_MIN_HOURS:
            best_len = cur_len = 1
            best_start = cur_start = 0

            for i in range(1, len(fpy_vals)):
                gap_h = (ts_vals[i] - ts_vals[i - 1]).total_seconds() / 3600
                # An idle gap larger than ANOMALY_MAX_STREAK_GAP_HOURS means the station
                # stopped testing between these two hours — they are NOT consecutive
                # production hours and must not form part of the same streak.
                is_consecutive = gap_h <= ANOMALY_MAX_STREAK_GAP_HOURS
                if is_consecutive and fpy_vals[i] < fpy_vals[i - 1]:
                    cur_len += 1
                    if cur_len > best_len:
                        best_len   = cur_len
                        best_start = cur_start
                else:
                    cur_len   = 1
                    cur_start = i

            if best_len >= ANOMALY_STREAK_MIN_HOURS:
                end_idx = best_start + best_len - 1
                drop_amt = round(float(fpy_vals[best_start]) - float(fpy_vals[end_idx]), 1)
                # Solo alertar si la racha terminó en las últimas 4 horas y la caída es >= 5%
                if drop_amt >= 5.0 and ts_vals[end_idx] >= pd.Timestamp(cutoff_4h):
                    streak_alert = {
                        "hours":     best_len,
                        "from":      ts_vals[best_start].strftime("%H:%M"),
                        "to":        ts_vals[end_idx].strftime("%H:%M"),
                        "fpy_start": round(float(fpy_vals[best_start]), 1),
                        "fpy_end":   round(float(fpy_vals[end_idx]), 1),
                        "drop_pts":  drop_amt,
                    }

        if not drop_alerts and streak_alert is None:
            continue

        severity = (
            "critical"
            if len(drop_alerts) >= 2 or (streak_alert and streak_alert["hours"] >= 5)
            else "warning"
        )

        alerts.append({
            "station_id":   int(sid),
            "station_name": sname,
            "product":      prod,
            "severity":     severity,
            "drop_alerts":  drop_alerts,
            "streak":       streak_alert,
        })

    alerts.sort(key=lambda a: (0 if a["severity"] == "critical" else 1, a["station_name"]))
    return alerts


def _refresh_anomalies(df: pd.DataFrame, context: str) -> None:
    """Run compute_anomalies and store results; swallows exceptions so DB refresh always completes."""
    try:
        alerts = compute_anomalies(df)
        _ANOMALY_CACHE["alerts"]      = alerts
        _ANOMALY_CACHE["computed_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        logger.info("[ANOMALY:%s] Scan complete: %d alert(s).", context, len(alerts))
    except Exception as exc:
        logger.error("[ANOMALY:%s] Scan failed: %s", context, exc)


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
    station_filter = request.args.get("stationid", "all")
    
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
            
        # Filter by station if specified
        if station_filter != "all":
            try:
                df_day = df_day[df_day['stationid'] == int(station_filter)]
            except ValueError:
                pass
        
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


@app.route("/api/failure_details")
def api_failure_details():
    """Devuelve registros individuales de fallas filtrados por producto y tipo de falla."""
    date_from_str = request.args.get("from", (datetime.now() - timedelta(days=DEFAULT_QUERY_RANGE_DAYS)).strftime("%Y-%m-%d"))
    date_to_str   = request.args.get("to",   datetime.now().strftime("%Y-%m-%d"))
    product       = request.args.get("product", "all")
    station_id_str = request.args.get("stationid", "all")
    falla         = request.args.get("falla", "")
    real_failures = request.args.get("real_failures", "false").lower() == "true"

    try:
        df = GLOBAL_CACHE["df"]
        if df.empty:
            return jsonify({"error": "System is booting"}), 503

        dt_from = pd.to_datetime(date_from_str)
        dt_to   = pd.to_datetime(date_to_str) + timedelta(days=1)

        mask = (df['endtime'] >= dt_from) & (df['endtime'] < dt_to)
        sub = df[mask].copy()

        if real_failures:
            sub = filter_real_failures(sub)

        if product != "all":
            sub = sub[sub["producto"] == product]

        if station_id_str != "all":
            try:
                station_id = int(station_id_str)
                sub = sub[sub["stationid"] == station_id]
            except ValueError:
                return jsonify({"error": "stationid invalido"}), 400

        # Solo registros FAILED
        sub = sub[sub["resultado"] == "FAILED"]

        if falla:
            sub = sub[sub["tipoFalla"] == falla]

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
    resultado     = request.args.get("resultado", "")  # PASSED o FAILED
    real_failures = request.args.get("real_failures", "false").lower() == "true"
    station_id_str = request.args.get("stationid", "all")

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
            # Último registro por prevQr
            now = datetime.now()
            cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)
            idx = sub.groupby("prevQr")["endtime"].idxmax()
            last_per_qr = sub.loc[idx]
            # Solo los relevantes (passed o failed >24h)
            sub = last_per_qr[
                (last_per_qr["resultado"] == "PASSED") |
                ((last_per_qr["resultado"] == "FAILED") & (last_per_qr["endtime"] < cutoff))
            ]

        if resultado:
            sub = sub[sub["resultado"] == resultado]

        sub = sub.sort_values("endtime", ascending=False).head(DETAILS_MAX_RECORDS)

        records = []
        for _, r in sub.iterrows():
            records.append({
                "prevQr":      str(r.get("prevQr", "") or ""),
                "currQr":      str(r.get("currQr", "") or ""),
                "resultado":   str(r.get("resultado", "")),
                "tipoFalla":   str(r.get("tipoFalla", "")),
                "failureCode": str(r.get("failureCode", "")),
                "stationName": str(r.get("stationName", "")),
                "endtime":     r["endtime"].strftime("%Y-%m-%d %H:%M") if pd.notna(r.get("endtime")) else "",
            })

        return jsonify(records)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


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
            # Último registro por prevQr (incluye PASSED y FAILED)
            now = datetime.now()
            cutoff = now - timedelta(hours=REAL_FAILURE_HOURS)
            idx = df_filtered.groupby("prevQr")["endtime"].idxmax()
            last_per_qr = df_filtered.loc[idx]

            # Solo contar como "failed" los que tienen >24h (misma lógica que filter_real_failures)
            # Los FAILED recientes (<24h) no se cuentan ni como pass ni como fail — están en proceso
            relevant = last_per_qr[
                (last_per_qr["resultado"] == "PASSED") |
                ((last_per_qr["resultado"] == "FAILED") & (last_per_qr["endtime"] < cutoff))
            ]

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

            pf_overall = _pf(relevant)
            pf_by_product = {}
            for prod, grp in relevant.groupby("producto"):
                pf_by_product[prod] = _pf(grp)

            pf_by_station = {}
            for sid, grp in relevant.groupby("stationid"):
                pf_by_station[int(sid)] = _pf(grp)

            pass_fail_pie = {
                "overall": pf_overall,
                "byProduct": pf_by_product,
                "byStation": pf_by_station,
            }

            # Ahora filtrar solo fallas reales para el resto de stats
            df_filtered = filter_real_failures(df_filtered)
        else:
            pass_fail_pie = None
        
        stats = compute_stats_cached(df_filtered, date_from_str, date_to_str, real_failures)
        stats["realFailures"] = real_failures
        if pass_fail_pie:
            stats["passFailPie"] = pass_fail_pie
        if GLOBAL_CACHE["last_updated"]:
            stats["cacheUpdated"] = GLOBAL_CACHE["last_updated"].strftime("%Y-%m-%d %H:%M:%S")
        return jsonify(stats)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@app.route("/api/anomalies")
def api_anomalies():
    product = request.args.get("product", "all")
    alerts  = _ANOMALY_CACHE["alerts"]
    if product != "all":
        alerts = [a for a in alerts if a["product"] == product]
    return jsonify({
        "alerts":      alerts,
        "count":       len(alerts),
        "computed_at": _ANOMALY_CACHE["computed_at"],
    })


if __name__ == "__main__":
    # Iniciar la carga fuerte de 1 año al arrancar el py
    load_historical_data()
    
    # Iniciar el recolector de 2do plano cada 15 min
    bg_thread = threading.Thread(target=background_update_worker, daemon=True)
    bg_thread.start()
    
    port = int(os.environ.get("PORT", DEFAULT_APP_PORT))
    logger.info("Manufacturing Dashboard started — http://localhost:%d", port)
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
