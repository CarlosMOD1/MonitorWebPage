#!/usr/bin/env python3
"""
processing.py — Procesamiento de DataFrames, carga de datos y calculo de KPIs.

Responsabilidades:
  - extract_tipo_falla(): clasifica el tipo de falla a partir de failureCode + notes
  - prepare_dataframe():  enriquece el DataFrame con columnas 'producto' y 'tipoFalla'
  - load_historical_data(): carga inicial desde la BD al arrancar
  - background_update_worker(): hilo que refresca datos cada N minutos
  - compute_stats(): calcula KPIs por producto, estacion y modos de falla
"""

import json
import re
import threading
import time

import pandas as pd
from datetime import datetime, timedelta

from config import (
    SQL,
    HISTORICAL_LOAD_DAYS, RECENT_REFRESH_DAYS, REFRESH_INTERVAL_MINUTES,
    INCREMENTAL_REFRESH_MINUTES, FULL_RESYNC_INTERVAL_MINUTES,
    REAL_FAILURE_HOURS,
    SPSF_BLADE_STATIONS, SPSF_BLADE_DIFF_STATIONS, UNDIFF_STATIONS,
    MODEL_NUMBER_TO_PRODUCT, STATION_TO_GROUP, STATION_ORDER,
    CODE_RE, METADATA_RE, GROUP_NAMES,
)
from db import connect, lookup_model_and_tape

# ─────────────────────────────────────────────────────────────
# CACHE GLOBAL EN RAM
# ─────────────────────────────────────────────────────────────
GLOBAL_CACHE = {
    "df":           pd.DataFrame(),
    "last_updated": None,
}


# ─────────────────────────────────────────────────────────────
# EXTRACCION DE TIPO DE FALLA
# Logica basada en los archivos Excel de referencia por producto.
# ─────────────────────────────────────────────────────────────

def _clean_falla_text(text):
    """Elimina detalles especificos de dispositivo entre parentesis/corchetes."""
    text = re.split(r'[\(\[]', text)[0].strip().rstrip(',').strip()
    return text


def _extract_station406(ns):
    """
    Estacion 406: notas con formato 'Test failed: <Componente> (<Detalle>)'.
    Ejemplos:
      'Test failed: Cellular (NOT_REGISTERED_CGREG)'   → 'Cellular (NOT_REGISTERED_CGREG)'
      'Test failed: GPS Test (Sats: 3, Fix: True)'     → 'GPS Test'
      'Test failed: USB Spec (Missing: Modem [vid:pid])' → 'USB Spec (Missing: Modem)'
    Retorna None si el formato no coincide.
    """
    if not ns.startswith("Test failed: "):
        return None
    content = ns[len("Test failed: "):].strip()
    if content.startswith("GPS Test"):
        return "GPS Test"
    if content.startswith("USB Spec"):
        cleaned = re.sub(r'\s*\([A-Za-z0-9][\w\s\-/]*\)', '', content)
        cleaned = re.sub(r'\s*\[[0-9a-fA-F:]+\]', '', cleaned)
        cleaned = re.sub(r',\s*\)', ')', cleaned).strip()
        return cleaned
    return content


def _extract_fail_tests_marker(ns):
    """
    Busca el marcador '"fail_tests": [" (con variantes compactas) en texto plano/JSON
    y extrae el primer elemento de la lista.
    Equivale a la formula Excel: MID(I2, FIND(...)+16, FIND(...)-...).
    Retorna None si no encuentra el marcador.
    """
    for marker in ('"fail_tests": ["', '"fail_tests":["', '"failed": ["', '"failed":["'):
        idx = ns.find(marker)
        if idx >= 0:
            start = idx + len(marker)
            end   = ns.find('"', start)
            if end > start:
                val = _clean_falla_text(ns[start:end].strip())
                if val:
                    return val
    return None


def _extract_from_json(ns, station, producto=None):
    """
    Parsea notas JSON y extrae el tipo de falla segun el esquema del grupo:
      - OPAL4:     wrapper {"notes": {"failure_codes": [...]}}
      - BlackBox2: {"station": "... - RF Chamber BLE Test"}
      - Generico:  {"failure_codes": [...]}
    Retorna None si la cadena no empieza con '{' o no hay coincidencia util.
    """
    if not ns.startswith("{"):
        return None
    try:
        parsed = json.loads(ns)

        # OPAL4: wrapper {"notes": {...}}
        if "notes" in parsed and isinstance(parsed["notes"], dict):
            inner   = parsed["notes"]
            fc_list = inner.get("failure_codes", [])
            fc0     = str(fc_list[0]).strip() if fc_list else ""
            return fc0 if (fc0 and not CODE_RE.match(fc0)) else station

        # BlackBox2 RF Chamber: special classification logic when data fields present
        if "station" in parsed:
            # Si este JSON corresponde a Black Box 2 (por producto o por nombre de station),
            # aplicar reglas especiales solicitadas por el equipo de validación.
            station_str = str(parsed.get("station", ""))
            is_blackbox2 = (producto and "Black Box 2" in str(producto)) or ("Black Box 2" in station_str)
            if is_blackbox2:
                # 1) MACID unknown -> MACID failure
                mac = parsed.get("mac_id")
                try:
                    if isinstance(mac, str) and mac.strip().upper() == "UNKNOWN":
                        return "MACID failure"
                except Exception:
                    pass

                # 2) failure_reason contains 'backward_check_failed' -> backward fail
                failure_reason = parsed.get("failure_reason")
                try:
                    if isinstance(failure_reason, str) and "backward_check_failed" in failure_reason:
                        return "backward fail"
                except Exception:
                    pass

                # 3) PCBA unknown -> pcba_qr unknown (user requested label)
                pcba = parsed.get("pcba_qr")
                try:
                    if isinstance(pcba, str) and pcba.strip().upper() == "UNKNOWN":
                        return "pcba_qr unknown"
                except Exception:
                    pass

                # 3) RSSI range check -> RSSI failure
                rssi_val = parsed.get("rssi_dbm")
                rssi_range = parsed.get("rssi_pass_range") or ""
                try:
                    m = re.search(r'(-?\d+)\s*to\s*(-?\d+)', str(rssi_range))
                    if m and rssi_val is not None:
                        low = int(m.group(1))
                        high = int(m.group(2))
                        try:
                            rnum = float(rssi_val)
                            if rnum < low or rnum > high:
                                return "RSSI failure"
                        except Exception:
                            pass
                except Exception:
                    pass

                # 4) Battery range check -> Battery failure
                try:
                    batt = parsed.get("battery_voltage_v")
                    bmin = parsed.get("battery_min_v")
                    bmax = parsed.get("battery_max_v")
                    if batt is not None and bmin is not None and bmax is not None:
                        try:
                            bf = float(batt)
                            bminf = float(bmin)
                            bmaxf = float(bmax)
                            if bf < bminf or bf > bmaxf:
                                return "Battery failure"
                        except Exception:
                            pass
                except Exception:
                    pass

                # Si ninguna regla matchea, devolver el nombre de la prueba tal cual (default)
                parts = station_str.split(" - ")
                return parts[-1].strip() if len(parts) > 1 else parts[0].strip()

            # No es BlackBox2 — comportamiento original: devolver la parte despues de ' - '
            parts = station_str.split(" - ")
            return parts[-1].strip() if len(parts) > 1 else parts[0].strip()

        # failure_codes[0] como ultimo recurso JSON
        fc_list = parsed.get("failure_codes", [])
        if fc_list:
            return str(fc_list[0]).strip()

    except (json.JSONDecodeError, TypeError):
        pass
    return None


def _extract_from_plain_text(ns, station, producto, fc):
    """
    Extrae tipo de falla de notas en texto plano (fallback final).
    Logica diferenciada por producto (Black Box vs resto).
    Retorna None si no puede extraer algo significativo.
    """
    # Solar Panel Test: 'FAIL - V:2.349V I:0.915A P:2.149W' → nombre de estacion
    if ns.startswith("FAIL - "):
        return station if station else fc

    # Black Box: 'MAC, Sleep Current | sleep: 0.53 | bat: 3.1'
    if "Black Box" in producto:
        left   = ns.split(" | ")[0].strip()
        first  = left.split(",")[0].strip()
        result = _clean_falla_text(first.split(":")[0].strip())
        if result and not METADATA_RE.match(result):
            return result

    # SPSF y otros: primera linea no vacia → split por ' | ' → split por ':'
    lines = [
        ln.strip() for ln in ns.replace("\r", "\n").split("\n")
        if ln.strip() and ln.strip().lower() not in ("undefined", "fail")
    ]
    if lines:
        segment = lines[0].split(" | ")[0].strip()
        result  = _clean_falla_text(segment.split(":")[0].strip())
        if result and not METADATA_RE.match(result):
            return result

    return None


def _extract_tipo_falla_impl(failure_code, notes_raw, station_name="", producto=""):
    """
    Determina el tipo de falla legible a partir de failureCode y notes.

    Orden de prioridad:
      1. Codigos de pase (ALL-PASS, -PASS-, MAXPASSEXCEEDED)  → devuelve el failureCode
      2. C-Track (prefijo CT-)                                 → devuelve el failureCode
      3. Caso especial de corriente conocido
      4. Estacion 406 ('Test failed: ...')
      5. Marcador fail_tests en texto/JSON
      6. JSON estructurado (OPAL4, BlackBox2, generico)
      7. Texto plano (Solar, Black Box, SPSF/otros)
      8. failureCode como ultimo recurso
    """
    fc      = str(failure_code).strip() if failure_code and str(failure_code) not in ("nan", "None") else ""
    station = str(station_name).strip() if station_name and str(station_name) not in ("nan", "None") else ""
    ns      = str(notes_raw).strip()    if notes_raw    and str(notes_raw)    not in ("nan", "None", "") else ""

    # 1. Codigos de pase
    if "ALL-PASS" in fc or "-PASS-" in fc or fc == "MAXPASSEXCEEDED":
        return fc

    # 2. C-Track (descriptivos por si solos)
    if fc.startswith("CT-"):
        return fc

    if ns and ns != "undefined":

        # 3. Caso especial de corriente (texto exacto conocido)
        if '"fail_tests": ["Sleep Current", "Non_FEM Current", "FEM Current"]' in ns:
            return "No current measured"

        # 4. Estacion 406
        result = _extract_station406(ns)
        if result:
            return result

        # 5. Marcador fail_tests en texto
        result = _extract_fail_tests_marker(ns)
        if result:
            return result

        # 6. JSON estructurado
        result = _extract_from_json(ns, station, producto)
        if result:
            return result

        # 7. Texto plano (omitir si 'ns' es JSON: ya se intento extraer en el
        # paso 6; parsearlo como texto plano produce basura como '{"failure_codes').
        # Si el JSON no tuvo nada util (ej. failure_codes vacio), cae al paso 8
        # y se usa el failureCode (ej. ALL-FAIL-000) como tipo de falla.
        if not ns.startswith("{"):
            result = _extract_from_plain_text(ns, station, producto, fc)
            if result:
                return result

    # 8. Ultimo recurso
    return fc


# ─────────────────────────────────────────────────────────────
# CACHE DE TIPO DE FALLA
# ─────────────────────────────────────────────────────────────
# Evita re-parsear failureCode/notes para registros ya vistos.
# Clave: (failureCode, notes, stationName, producto) normalizados.
# Carga inicial: se rellena progresivamente (~261k filas la primera vez).
# Refresco de fondo: la mayoria ya existe → lookup O(1) por fila → segundos.
_TIPO_FALLA_CACHE: dict = {}


def extract_hardware_id(notes_raw):
    """
    Extrae el campo 'hardware_id' de notas en formato JSON.
    Usado para diferenciar equipos dentro de una misma estacion
    (ej. MPS-Jasper / stationid=238, que reporta varios hardware_id distintos).
    Retorna '' si las notas no son JSON o no contienen 'hardware_id'.
    """
    try:
        if notes_raw and isinstance(notes_raw, str) and notes_raw.strip().startswith("{"):
            hw_id = json.loads(notes_raw).get("hardware_id")
            if hw_id:
                return str(hw_id)
    except Exception:
        pass
    return ""


def extract_tipo_falla(failure_code, notes_raw, station_name="", producto=""):
    """Wrapper publico con cache. Llama a _extract_tipo_falla_impl solo la primera vez."""
    key = (
        str(failure_code).strip() if failure_code and str(failure_code) not in ("nan", "None") else "",
        str(notes_raw).strip()    if notes_raw    and str(notes_raw)    not in ("nan", "None", "") else "",
        str(station_name).strip() if station_name and str(station_name) not in ("nan", "None") else "",
        str(producto).strip()     if producto     and str(producto)     not in ("nan", "None") else "",
    )
    if key in _TIPO_FALLA_CACHE:
        return _TIPO_FALLA_CACHE[key]
    result = _extract_tipo_falla_impl(failure_code, notes_raw, station_name, producto)
    _TIPO_FALLA_CACHE[key] = result
    return result


# ─────────────────────────────────────────────────────────────
# PREPARACION DEL DATAFRAME
# ─────────────────────────────────────────────────────────────

def prepare_dataframe(df):
    """
    Enriquece el DataFrame con las columnas 'producto' y 'tipoFalla'.

    Pasos:
      1. Convertir endtime a datetime si hace falta
      2. Asignar 'producto' por estacion (regla base)
      3. Resolver SPSF vs Blade en estaciones compartidas via modelNumber
         (QRs vacios/nulos quedan con fallback 'SPSF' y se muestran en el portal)
      4. Duplicar filas en estaciones no-diferenciables (406, 393)
      5. Pre-calcular 'tipoFalla' llamando a extract_tipo_falla
    """
    if df.empty:
        return df

    if not pd.api.types.is_datetime64_any_dtype(df["endtime"]):
        # Coerce invalid values to NaT to avoid .dt accessor errors on some
        # environments where the DB driver returns strings or mixed types.
        df["endtime"] = pd.to_datetime(df["endtime"], errors="coerce")
        # Drop rows without a valid timestamp as they are not usable for time-series
        df = df.dropna(subset=["endtime"]) if not df.empty else df

    # Paso 2: asignacion inicial por estacion
    df["producto"] = df["stationid"].map(STATION_TO_GROUP).fillna("Other")

    # Paso 2.5: resolver Lima vs White Tape por tapeColor
    # Estacion 90 esta compartida por Lima y White Tape.
    from config import GROUPS
    lima_white_stations = set(GROUPS.get("Lima", [])) & set(GROUPS.get("White tape", []))
    if lima_white_stations:
        lw_mask = df["stationid"].isin(lima_white_stations)
        if lw_mask.any():
            from db import lookup_model_and_tape
            # Consultar QRs para saber el tapeColor
            lw_qrs = df.loc[lw_mask, "currQr"].dropna().unique().tolist()
            qr_data = lookup_model_and_tape(lw_qrs)
            
            def _assign_lw_product(row):
                qr = str(row["currQr"])
                if qr in qr_data:
                    tape = qr_data[qr].get("tapeColor", "").lower()
                    if tape == "white":
                        return "White tape"
                    elif tape == "lime":
                        return "Lima"
                # Fallback predeterminado si no hay QR o no tiene color
                return "Lima"

            df.loc[lw_mask, "producto"] = df.loc[lw_mask].apply(_assign_lw_product, axis=1)

    # Paso 3 y 4: resolver SPSF vs Blade
    shared_mask = df["stationid"].isin(SPSF_BLADE_STATIONS)
    if shared_mask.any():

        # ── CONFIGURACIÓN ACTUAL ───────────────────────────────────────
        # Se ignora el filtro por modelNumber y se duplican TODOS los
        # registros de estaciones compartidas (SPSF y Blade) para que
        # ambas pestañas muestren exactamente los mismos datos.
        # (Esto incluye a las estaciones 406 y 393 por definición).
        df.loc[shared_mask, "producto"] = "SPSF"
        blade_rows = df.loc[shared_mask].copy()
        blade_rows["producto"] = "Blade"
        df = pd.concat([df, blade_rows], ignore_index=True)
        # ───────────────────────────────────────────────────────────────

        """
        # --- CÓDIGO ORIGINAL (DESCOMENTAR PARA REVERTIR AL FILTRO POR MODELO) ---
        # Consultar modelNumber para estaciones diferenciables.
        # lookup_model_numbers ya ignora QRs vacios/nulos internamente.
        shared_qrs = df.loc[shared_mask, "currQr"].dropna().unique().tolist()
        model_map  = lookup_model_numbers(shared_qrs)

        diff_mask = shared_mask & df["stationid"].isin(SPSF_BLADE_DIFF_STATIONS)
        if diff_mask.any():
            df.loc[diff_mask, "producto"] = df.loc[diff_mask, "currQr"].map(
                lambda qr: MODEL_NUMBER_TO_PRODUCT.get(
                    model_map.get(str(qr), ""), "SPSF"
                )
            )

        # Estaciones NO diferenciables (406, 393): duplicar → SPSF + Blade
        undiff_mask = shared_mask & df["stationid"].isin(UNDIFF_STATIONS)
        if undiff_mask.any():
            df.loc[undiff_mask, "producto"] = "SPSF"
            blade_rows = df.loc[undiff_mask].copy()
            blade_rows["producto"] = "Blade"
            df = pd.concat([df, blade_rows], ignore_index=True)
        # --------------------------------------------------------------------------
        """

    # Paso 5: pre-calcular tipo de falla (costoso; se hace una sola vez aqui)
    df["tipoFalla"] = df.apply(
        lambda r: extract_tipo_falla(
            r["failureCode"], r["notes"], r["stationName"], r["producto"]
        ),
        axis=1,
    )

    # Paso 6: pre-calcular hardwareId (ej. MPS-Jasper reporta varios hardware_id
    # distintos bajo la misma estacion). Vacio ('') si no aplica.
    df["hardwareId"] = df["notes"].apply(extract_hardware_id)

    return df


# ─────────────────────────────────────────────────────────────
# CARGA INICIAL Y ACTUALIZACION EN SEGUNDO PLANO
# ─────────────────────────────────────────────────────────────

def load_historical_data():
    """Carga los ultimos HISTORICAL_LOAD_DAYS dias desde la BD al arrancar el servidor."""
    t0 = time.time()
    print("[DATABASE] Iniciando carga de datos historicos...")
    start_date = (datetime.now() - timedelta(days=HISTORICAL_LOAD_DAYS)).strftime("%Y-%m-%d")

    conn = connect()
    df   = pd.read_sql(SQL + f" AND e.endtime >= '{start_date}'", conn)
    conn.close()
    t_sql = time.time() - t0
    print(f"[DATABASE] SQL completo: {len(df)} filas en {t_sql:.1f}s. Procesando tipoFalla...")

    df = prepare_dataframe(df)

    GLOBAL_CACHE["df"]           = df
    GLOBAL_CACHE["last_updated"] = datetime.now()
    print(f"[DATABASE] Carga inicial completa. {len(df)} registros en RAM. "
          f"Total: {time.time()-t0:.1f}s | Cache tipoFalla: {len(_TIPO_FALLA_CACHE)} entradas unicas.")


def background_update_worker():
    """
    Hilo en segundo plano con refresco en dos niveles:
      - Cada REFRESH_INTERVAL_MINUTES: refresco incremental, solo consulta los
        ultimos INCREMENTAL_REFRESH_MINUTES (ventana chica -> query rapida y liviana).
      - Cada FULL_RESYNC_INTERVAL_MINUTES: resync completo de RECENT_REFRESH_DAYS,
        para capturar ediciones/correcciones tardias en registros ya cacheados.
    """
    ticks_since_full_resync = 0
    ticks_for_full_resync   = max(1, FULL_RESYNC_INTERVAL_MINUTES // REFRESH_INTERVAL_MINUTES)

    while True:
        time.sleep(REFRESH_INTERVAL_MINUTES * 60)
        try:
            t0 = time.time()
            ticks_since_full_resync += 1
            do_full_resync = ticks_since_full_resync >= ticks_for_full_resync

            if do_full_resync:
                window_days_or_minutes = f"{RECENT_REFRESH_DAYS} dias (resync completo)"
                cutoff_str = (datetime.now() - timedelta(days=RECENT_REFRESH_DAYS)).strftime("%Y-%m-%d %H:%M:%S")
            else:
                window_days_or_minutes = f"{INCREMENTAL_REFRESH_MINUTES} minutos (incremental)"
                cutoff_str = (datetime.now() - timedelta(minutes=INCREMENTAL_REFRESH_MINUTES)).strftime("%Y-%m-%d %H:%M:%S")

            print(f"[BACKGROUND] Refrescando registros recientes de DB ({window_days_or_minutes})...")
            cutoff_ts = pd.to_datetime(cutoff_str)

            conn      = connect()
            df_recent = pd.read_sql(SQL + f" AND e.endtime >= '{cutoff_str}'", conn)
            conn.close()

            df_recent = prepare_dataframe(df_recent)

            # Conservar datos anteriores al cutoff y unir con los frescos
            old_data = GLOBAL_CACHE["df"]
            old_data = old_data[old_data["endtime"] < cutoff_ts]
            new_df   = pd.concat([old_data, df_recent], ignore_index=True)

            GLOBAL_CACHE["df"]           = new_df
            GLOBAL_CACHE["last_updated"] = datetime.now()

            if do_full_resync:
                ticks_since_full_resync = 0

            print(f"[BACKGROUND] Actualizacion completa. {len(df_recent)} filas en "
                  f"{time.time()-t0:.1f}s | Total RAM: {len(new_df)} | Cache tipoFalla: {len(_TIPO_FALLA_CACHE)}")

        except Exception as exc:
            print(f"[BACKGROUND] Error en hilo de actualizacion: {exc}")


# ─────────────────────────────────────────────────────────────
# FILTROS Y CALCULO DE KPIs
# ─────────────────────────────────────────────────────────────

def filter_real_failures(df):
    """
    Retorna solo las 'fallas reales': el ultimo intento por QR es FAILED
    y tiene mas de REAL_FAILURE_HOURS de antiguedad (unidad que no pudo ser
    reprocesada en el turno).
    """
    if df.empty:
        return df
    cutoff      = datetime.now() - timedelta(hours=REAL_FAILURE_HOURS)
    idx         = df.groupby("currQr")["endtime"].idxmax()
    last_per_qr = df.loc[idx]
    return last_per_qr[
        (last_per_qr["resultado"] == "FAILED") &
        (last_per_qr["endtime"] < cutoff)
    ]


def _kpis(data):
    """
    Calcula metricas basicas: total, passed, failed, yield%, fpy%.
    FPY (First Pass Yield) cuenta los codigos con patron '-PASS-' en failureCode.
    """
    t  = len(data)
    p  = int((data["resultado"] == "PASSED").sum())
    f  = int((data["resultado"] == "FAILED").sum())
    fp = int(data["failureCode"].str.contains(r".*-PASS-.*", na=False).sum())
    return {
        "total":  t,
        "failed": f,
        "passed": p,
        "yield":  round(p / t * 100, 1) if t else 0,
        "fpy":    round(fp / t * 100, 1) if t else 0,
    }


def _get_kpi_subset(data, group_cols, real_failures):
    """
    En modo real_failures: filtra al ultimo intento por QR y conserva
    solo los que pasaron o fallaron hace mas de REAL_FAILURE_HOURS.
    En modo normal: retorna el DataFrame sin modificar.
    """
    if not real_failures:
        return data
    cutoff = datetime.now() - timedelta(hours=REAL_FAILURE_HOURS)
    key    = (group_cols + ["currQr"]) if group_cols else ["currQr"]
    idx    = data.groupby(key)["endtime"].idxmax()
    last   = data.loc[idx]
    return last[
        (last["resultado"] == "PASSED") |
        ((last["resultado"] == "FAILED") & (last["endtime"] < cutoff))
    ]


def compute_stats(df, date_from, date_to, real_failures=False):
    """
    Calcula KPIs generales, por producto, por estacion y modos de falla.
    Retorna un dict completamente serializable a JSON.
    """
    # KPIs globales
    data_overall = _get_kpi_subset(df, [], real_failures)
    overall      = _kpis(data_overall)
    overall["dateFrom"] = date_from
    overall["dateTo"]   = date_to

    # KPIs por producto
    data_prod  = _get_kpi_subset(df, ["producto"], real_failures)
    by_product = {
        prod: _kpis(grp)
        for prod, grp in data_prod.groupby("producto")
    }

    # KPIs por estacion
    data_stat = _get_kpi_subset(df, ["stationid", "stationName", "producto"], real_failures)
    stations  = []
    for (sid, sname, prod), grp in data_stat.groupby(["stationid", "stationName", "producto"]):
        k = _kpis(grp)
        stations.append({"id": int(sid), "name": sname, "producto": prod, **k})
    stations.sort(key=lambda x: STATION_ORDER.get((x["producto"], x["id"]), 9999))

    # Modos de falla
    data_f = data_stat if real_failures else df
    df_f   = data_f[
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
        "overall":     overall,
        "byProduct":   by_product,
        "stations":    stations,
        "failures":    failures[:300],
        "groups":      GROUP_NAMES,
        "lastUpdated": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    }


def compute_report_data(df, date_from_str, date_to_str):
    """
    Calcula metricas resumidas, series de tiempo y top de fallas
    para el panel de reporte en la UI. Retorna un dict serializable a JSON.
    """
    if df.empty:
        return {"empty": True}

    total   = len(df)
    passed  = int((df["resultado"] == "PASSED").sum())
    failed  = int((df["resultado"] == "FAILED").sum())
    yld_pct = round(passed / total * 100, 1) if total else 0.0
    fpy_cnt = int(df["failureCode"].str.contains(r".*-PASS-.*", na=False).sum())
    fpy_pct = round(fpy_cnt / total * 100, 1) if total else 0.0

    # RTY (Rolled Throughput Yield): producto de los FPY por estacion
    # Usar solo el PRIMER intento por unidad en cada estación (ignora reintentos
    # al calcular FPY, aunque los reintentos sí reducen el FPY porque no se
    # cuentan como primera pasada). Excluir estaciones con muy pocas unidades.
    RTY_MIN_UNITS_PER_STATION = 5
    rty_pct = None
    rty_by_station = []

    groups = [g for _, g in df.groupby(["stationid", "stationName", "producto"]) if len(g) > 0]
    if len(groups) > 1:
        rty = 1.0
        included = 0
        for g in groups:
            # Primer intento por unidad (prevQr) ordenado por endtime
            first = g.sort_values("endtime").drop_duplicates(subset="prevQr", keep="first")
            total_first = len(first)
            if total_first < RTY_MIN_UNITS_PER_STATION:
                # registrar pero no incluir en el producto del RTY
                sid = int(g["stationid"].iloc[0])
                rty_by_station.append({
                    "id": sid,
                    "name": str(g["stationName"].iloc[0]),
                    "producto": g["producto"].iloc[0] if "producto" in g.columns else None,
                    "units_first_attempt": total_first,
                    "first_pass": int((first["resultado"] == "PASSED").sum()),
                    "not_first_pass": int((first["resultado"] != "PASSED").sum()),
                    "fpy_pct": round((int((first["resultado"] == "PASSED").sum()) / total_first * 100), 1) if total_first else 0.0,
                    "included_in_rty": False,
                })
                continue

            first_pass = int((first["resultado"] == "PASSED").sum())
            not_first_pass = total_first - first_pass
            fpy_pct_station = round(first_pass / total_first * 100, 1) if total_first else 0.0

            rty *= (first_pass / total_first) if total_first else 1.0
            included += 1

            sid = int(g["stationid"].iloc[0])
            rty_by_station.append({
                "id": sid,
                "name": str(g["stationName"].iloc[0]),
                "producto": g["producto"].iloc[0] if "producto" in g.columns else None,
                "units_first_attempt": total_first,
                "first_pass": first_pass,
                "not_first_pass": not_first_pass,
                "fpy_pct": fpy_pct_station,
                "included_in_rty": True,
            })

        if included > 0:
            rty_pct = round(rty * 100, 1)
        else:
            rty_pct = None

    # Series de tiempo — columnas sin guion bajo inicial para evitar problemas con itertuples
    df2 = df.copy()
    # Ensure endtime is datetime-like before using .dt (defensive for Pi envs)
    if not pd.api.types.is_datetime64_any_dtype(df2["endtime"]):
        df2["endtime"] = pd.to_datetime(df2["endtime"], errors="coerce")
    df2 = df2.dropna(subset=["endtime"]) if not df2.empty else df2
    df2["period_d"] = df2["endtime"].dt.date.astype(str)
    df2["period_w"] = df2["endtime"].dt.to_period("W").astype(str)
    df2["period_m"] = df2["endtime"].dt.to_period("M").astype(str)

    def _ts_agg(grp_col, limit):
        agg = (
            df2.groupby(grp_col)
            .agg(
                total  =("resultado", "count"),
                passed =("resultado", lambda x: (x == "PASSED").sum()),
            )
            .reset_index()
        )
        agg["yield_pct"] = (
            (agg["passed"] / agg["total"] * 100).round(1).where(agg["total"] > 0, 0)
        )
        agg = agg.tail(limit)
        return {
            "labels": agg[grp_col].astype(str).tolist(),
            "yields": agg["yield_pct"].tolist(),
            "totals": agg["total"].tolist(),
        }

    # Top failures
    df_fail   = df[df["resultado"] == "FAILED"]
    tot_fails = len(df_fail)
    top_failures = []
    if not df_fail.empty:
        agg_f = (
            df_fail.groupby(["tipoFalla", "stationName", "producto"])
            .size()
            .reset_index(name="count")
            .sort_values("count", ascending=False)
            .head(20)
        )
        for _, row in agg_f.iterrows():
            top_failures.append({
                "falla":    str(row["tipoFalla"]),
                "station":  str(row["stationName"]),
                "producto": str(row["producto"]),
                "count":    int(row["count"]),
                "pct":      round(row["count"] / tot_fails * 100, 1) if tot_fails else 0,
            })

    return {
        "dateFrom":    date_from_str,
        "dateTo":      date_to_str,
        "summary": {
            "total":     total,
            "passed":    passed,
            "failed":    failed,
            "yield_pct": yld_pct,
            "fpy_pct":   fpy_pct,
            "rty_pct":   rty_pct,
        },
        "rty_by_station": rty_by_station,
        "daily":   _ts_agg("period_d", 30),
        "weekly":  _ts_agg("period_w", 12),
        "monthly": _ts_agg("period_m", 12),
        "topFailures": top_failures,
    }
