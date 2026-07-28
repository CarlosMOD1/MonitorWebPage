#!/usr/bin/env python3
"""
app.py — Manufacturing Dashboard: servidor Flask y rutas de API.

Arquitectura del proyecto:
  config.py     — Constantes, grupos de estaciones, SQL base
  db.py         — Conexiones a SQL Server y consultas seguras
  processing.py — Procesamiento de DataFrames, cache en RAM, KPIs
  reports.py    — Generacion de reportes Excel
  app.py        — Rutas Flask (este archivo)

Ejecutar: python app.py
Navegador: http://localhost:5000
"""

import os
import json
import threading
from datetime import datetime, timedelta

import pandas as pd
from flask import Flask, jsonify, request, send_from_directory, Response

from config import (
    DEFAULT_QUERY_RANGE_DAYS, DEBUG_ENDPOINT_RANGE_DAYS, DETAILS_MAX_RECORDS,
    REAL_FAILURE_HOURS, GROUPS, GROUP_NAMES, HISTORICAL_LOAD_DAYS,
    DEFAULT_APP_PORT,
)
from db import connect
from processing import (
    GLOBAL_CACHE,
    compute_stats,
    compute_report_data,
    load_historical_data,
    background_update_worker,
    extract_tipo_falla,
)
from reports import generate_excel_report

app = Flask(__name__, static_folder="static")


# ─────────────────────────────────────────────────────────────
# UTILIDADES COMPARTIDAS DE RUTAS
# ─────────────────────────────────────────────────────────────

def _default_dates():
    now = datetime.now()
    return {
        "from": (now - timedelta(days=DEFAULT_QUERY_RANGE_DAYS)).strftime("%Y-%m-%d"),
        "to":   now.strftime("%Y-%m-%d"),
    }


def _get_cached_df():
    """Retorna el DataFrame del cache global. Retorna None si el cache aun no esta listo."""
    df = GLOBAL_CACHE["df"]
    return df if not df.empty else None


def _filter_by_dates(df, date_from_str, date_to_str):
    dt_from = pd.to_datetime(date_from_str)
    dt_to   = pd.to_datetime(date_to_str) + timedelta(days=1)
    return df[(df["endtime"] >= dt_from) & (df["endtime"] < dt_to)].copy()


def _filter_by_product(df, product):
    if product and product not in ("all", "") and product in GROUPS:
        return df[df["producto"] == product]
    return df


def _filter_by_station(df, station_id_str):
    """Filtra por stationid. Lanza ValueError si el ID no es entero valido."""
    if station_id_str and station_id_str != "all":
        try:
            sid = int(station_id_str)
        except ValueError:
            raise ValueError(f"stationid invalido: '{station_id_str}'")
        return df[df["stationid"] == sid]
    return df


def _filter_by_hardware_id(df, hardware_id):
    """Filtra por hardwareId (ej. MPS-Jasper / stationid=238 con varios hardware_id)."""
    if hardware_id and hardware_id not in ("all", "") and "hardwareId" in df.columns:
        return df[df["hardwareId"] == hardware_id]
    return df


def _apply_real_failures_filter(df):
    """Retorna solo ultimo intento por QR (PASSED o FAILED > REAL_FAILURE_HOURS)."""
    cutoff = datetime.now() - timedelta(hours=REAL_FAILURE_HOURS)
    idx    = df.groupby("currQr")["endtime"].idxmax()
    last   = df.loc[idx]
    return last[
        (last["resultado"] == "PASSED") |
        ((last["resultado"] == "FAILED") & (last["endtime"] < cutoff))
    ]


# ─────────────────────────────────────────────────────────────
# RUTAS ESTATICAS
# ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_from_directory("static", "index.html")


@app.route("/monitor")
def monitor_page():
    return send_from_directory("static", "monitor.html")


# ─────────────────────────────────────────────────────────────
# RUTA: DATOS PRINCIPALES DEL DASHBOARD
# ─────────────────────────────────────────────────────────────

@app.route("/api/data")
def api_data():
    dates         = _default_dates()
    date_from_str = request.args.get("from", dates["from"])
    date_to_str   = request.args.get("to",   dates["to"])
    real_failures = request.args.get("real_failures", "false").lower() == "true"
    hardware_id   = request.args.get("hardware_id", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return jsonify({
                "error": "System is booting up database into RAM. Try again in a minute.",
                "cacheBooting": True,
            }), 503

        df_filtered = _filter_by_dates(df, date_from_str, date_to_str)
        df_filtered = _filter_by_hardware_id(df_filtered, hardware_id)

        # Calculo QR-level para el pie de pass/fail en modo real_failures
        pass_fail_pie = None
        if real_failures and not df_filtered.empty:
            cutoff = datetime.now() - timedelta(hours=REAL_FAILURE_HOURS)

            def _get_relevant(sub_df, group_by_cols):
                key  = (group_by_cols + ["currQr"]) if group_by_cols else ["currQr"]
                idx  = sub_df.groupby(key)["endtime"].idxmax()
                last = sub_df.loc[idx]
                return last[
                    (last["resultado"] == "PASSED") |
                    ((last["resultado"] == "FAILED") & (last["endtime"] < cutoff))
                ]

            def _pf_kpi(d, original_subset=None):
                p = int((d["resultado"] == "PASSED").sum())
                f_df = d[d["resultado"] == "FAILED"]
                f = len(f_df)
                t = p + f
                
                fail_bins = {"1": 0, "2": 0, "3plus": 0}
                if not f_df.empty and original_subset is not None and not original_subset.empty:
                    fails_only = original_subset[
                        original_subset["currQr"].isin(f_df["currQr"]) & 
                        (original_subset["resultado"] == "FAILED")
                    ]
                    counts = fails_only.groupby("currQr").size()
                    for count in counts:
                        if count == 1:
                            fail_bins["1"] += 1
                        elif count == 2:
                            fail_bins["2"] += 1
                        else:
                            fail_bins["3plus"] += 1
                            
                return {
                    "passed": p, "failed": f, "total": t,
                    "yield": round(p / t * 100, 1) if t else 0,
                    "failBins": fail_bins
                }

            rel_overall      = _get_relevant(df_filtered, [])
            rel_product      = _get_relevant(df_filtered, ["producto"])
            rel_station      = _get_relevant(df_filtered, ["stationid"])
            rel_station_prod = _get_relevant(df_filtered, ["stationid", "producto"])

            pass_fail_pie = {
                "overall":   _pf_kpi(rel_overall, df_filtered),
                "byProduct": {
                    p: _pf_kpi(g, df_filtered[df_filtered["producto"] == p]) 
                    for p, g in rel_product.groupby("producto")
                },
                "byStation": {
                    int(s): _pf_kpi(g, df_filtered[df_filtered["stationid"] == s]) 
                    for s, g in rel_station.groupby("stationid")
                },
                "byStationProduct": {
                    f"{int(s)}_{prod}": _pf_kpi(g, df_filtered[(df_filtered["stationid"] == s) & (df_filtered["producto"] == prod)])
                    for (s, prod), g in rel_station_prod.groupby(["stationid", "producto"])
                },
            }

        stats = compute_stats(df_filtered, date_from_str, date_to_str, real_failures=real_failures)
        stats["realFailures"] = real_failures
        stats["hardwareId"]   = hardware_id
        if pass_fail_pie:
            stats["passFailPie"] = pass_fail_pie
        if GLOBAL_CACHE["last_updated"]:
            stats["cacheUpdated"] = GLOBAL_CACHE["last_updated"].strftime("%Y-%m-%d %H:%M:%S")
        return jsonify(stats)

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ─────────────────────────────────────────────────────────────
# RUTA: MONITOR EN TIEMPO REAL (FPY por hora)
# ─────────────────────────────────────────────────────────────

@app.route("/api/monitor_data")
def api_monitor_data():
    target_date_str = request.args.get("date", datetime.now().strftime("%Y-%m-%d"))
    product_filter  = request.args.get("product", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return jsonify({"date": target_date_str, "product": product_filter,
                            "error": "System booting", "data": {}})

        target_dt = pd.to_datetime(target_date_str)
        df_day    = df[
            (df["endtime"] >= target_dt) &
            (df["endtime"] < target_dt + timedelta(days=1))
        ].copy()
        df_day = _filter_by_product(df_day, product_filter)

        if df_day.empty:
            return jsonify({"date": target_date_str, "product": product_filter, "data": {}})

        def _extract_hardware_id(notes_str, station_name):
            try:
                if notes_str and isinstance(notes_str, str) and notes_str.startswith("{"):
                    hw_id = json.loads(notes_str).get("hardware_id")
                    if hw_id:
                        return hw_id
            except Exception:
                pass
            return station_name

        def _group_key(row):
            if str(row["stationName"]).strip() == "Jasper MPS":
                return _extract_hardware_id(row["notes"], row["stationName"])
            return str(row["stationName"]).strip()

        df_day["is_pass"]  = (df_day["resultado"] == "PASSED").astype(int)
        df_day["GroupKey"] = df_day.apply(_group_key, axis=1)
        df_day = df_day[df_day["GroupKey"] != "MPS_CTRACK_00"]

        response_data = {}
        for group, df_g in df_day.groupby("GroupKey"):
            hourly_counts = [0] * 24
            hourly_sums   = [0] * 24
            hourly_fpy    = [0.0] * 24

            for _, row in df_g.iterrows():
                h = row["endtime"].hour
                hourly_counts[h] += 1
                hourly_sums[h]   += row["is_pass"]

            for h in range(24):
                if hourly_counts[h] > 0:
                    hourly_fpy[h] = hourly_sums[h] / hourly_counts[h] * 100

            total_att  = int(df_g["is_pass"].count())
            total_pass = int(df_g["is_pass"].sum())
            response_data[str(group)] = {
                "stationName": str(df_g["stationName"].iloc[0]),
                "hourly": {
                    "fpy":   hourly_fpy,
                    "count": hourly_counts,
                    "sum":   hourly_sums,
                },
                "total_att":  total_att,
                "total_pass": total_pass,
                "global_fpy": (total_pass / total_att * 100) if total_att > 0 else 0,
            }

        return jsonify({"date": target_date_str, "product": product_filter,
                        "data": response_data})

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ─────────────────────────────────────────────────────────────
# RUTA: MONITOR FPY — RANGO FLEXIBLE
# ─────────────────────────────────────────────────────────────

@app.route("/api/fpy_range")
def api_fpy_range():
    """
    Historial de yield para el rango de fechas del dashboard principal.
    Granularidad automatica:
      <= 3 dias  -> horario  (1 punto/hora)
      > 3 dias   -> diario   (1 punto/dia)
    Parametros:
      from_dt   — fecha inicio (YYYY-MM-DD)
      to_dt     — fecha fin    (YYYY-MM-DD)
      product   — filtro de producto ('all' o nombre)
      stationid — filtro de estacion (int o 'all')
      hardware_id — filtro opcional de hardware (ej. MPS-Jasper / stationid=238)
    """
    from_dt_str = request.args.get("from_dt", None)
    to_dt_str   = request.args.get("to_dt",   None)
    product     = request.args.get("product",  "all")
    station_id  = request.args.get("stationid","all")
    hardware_id = request.args.get("hardware_id", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return jsonify({"error": "System booting", "labels": [], "fpy": [],
                            "count": [], "passed": []}), 503

        now     = datetime.now()
        dt_from = pd.to_datetime(from_dt_str) if from_dt_str else now - timedelta(days=30)
        dt_to   = (pd.to_datetime(to_dt_str) + timedelta(hours=23, minutes=59, seconds=59)
                   if to_dt_str else now)

        delta_days = max((dt_to - dt_from).days, 1)
        use_hourly = delta_days <= 3

        df_range = df[(df["endtime"] >= dt_from) & (df["endtime"] <= dt_to)].copy()
        df_range = _filter_by_product(df_range, product)
        df_range = _filter_by_hardware_id(df_range, hardware_id)
        df_range = df_range.dropna(subset=["endtime"])

        station_name = product if product != "all" else "All Stations"

        if station_id and station_id != "all":
            try:
                sid      = int(station_id)
                df_range = df_range[df_range["stationid"] == sid]
                station_name = (str(df_range["stationName"].iloc[0])
                                if not df_range.empty else f"Station {sid}")
            except (ValueError, KeyError, IndexError):
                pass

        if df_range.empty:
            return jsonify({
                "labels": [], "fpy": [], "count": [], "passed": [],
                "station_name": station_name,
            })

        if use_hourly:
            df_range["bucket"] = df_range["endtime"].dt.floor("h")
            all_buckets = pd.date_range(
                start=pd.to_datetime(dt_from).floor("h"), end=pd.to_datetime(dt_to).floor("h"), freq="h"
            )
            lbl_fmt = "%m/%d %H:%M"
        else:
            df_range["bucket"] = df_range["endtime"].dt.floor("D")
            all_buckets = pd.date_range(
                start=pd.to_datetime(dt_from).floor("D"), end=pd.to_datetime(dt_to).floor("D"), freq="D"
            )
            lbl_fmt = "%b %d"

        agg = (
            df_range.groupby("bucket")
            .agg(
                total =("resultado", "count"),
                passed=("resultado", lambda x: int((x == "PASSED").sum())),
            )
            .reset_index()
        )

        full   = pd.DataFrame({"bucket": all_buckets})
        merged = full.merge(agg, on="bucket", how="left")
        merged["total"]  = merged["total"].fillna(0).astype(int)
        merged["passed"] = merged["passed"].fillna(0).astype(int)
        merged["fpy"]    = (
            (merged["passed"] / merged["total"] * 100)
            .where(merged["total"] > 0)
            .round(1)
        )

        return jsonify({
            "labels":       [b.strftime(lbl_fmt) for b in merged["bucket"]],
            "fpy":          [None if pd.isna(v) else float(v) for v in merged["fpy"]],
            "count":        merged["total"].tolist(),
            "passed":       merged["passed"].tolist(),
            "station_name": station_name,
        })

    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ─────────────────────────────────────────────────────────────
# RUTA: HARDWARE IDs POR ESTACION (ej. MPS-Jasper / stationid=238)
# ─────────────────────────────────────────────────────────────

@app.route("/api/hardware_ids")
def api_hardware_ids():
    """
    Devuelve la lista de hardware_id distintos reportados por una estacion
    (extraidos del JSON en 'notes'). Se usa para poblar el filtro opcional
    de hardware en el frontend cuando una estacion reporta varios equipos
    (ej. MPS-Jasper). Retorna lista vacia si la estacion no usa este campo.
    """
    station_id = request.args.get("stationid", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return jsonify({"hardwareIds": []})

        sub = _filter_by_station(df, station_id)
        if sub.empty or "hardwareId" not in sub.columns:
            return jsonify({"hardwareIds": []})

        ids = sorted(v for v in sub["hardwareId"].dropna().unique().tolist() if v)
        return jsonify({"hardwareIds": ids})

    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ─────────────────────────────────────────────────────────────
# RUTA: DETALLES DE FALLAS
# ─────────────────────────────────────────────────────────────

@app.route("/api/failure_details")
def api_failure_details():
    dates         = _default_dates()
    date_from_str = request.args.get("from", dates["from"])
    date_to_str   = request.args.get("to",   dates["to"])
    product       = request.args.get("product", "all")
    station_id    = request.args.get("stationid", "all")
    falla         = request.args.get("falla", "")
    real_failures = request.args.get("real_failures", "false").lower() == "true"
    export_csv    = request.args.get("csv", "false").lower() == "true"
    hardware_id   = request.args.get("hardware_id", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return jsonify({"error": "System is booting"}), 503

        sub = _filter_by_dates(df, date_from_str, date_to_str)
        sub = _filter_by_product(sub, product)
        sub = _filter_by_station(sub, station_id)
        sub = _filter_by_hardware_id(sub, hardware_id)

        if real_failures:
            cutoff = datetime.now() - timedelta(hours=REAL_FAILURE_HOURS)
            idx    = sub.groupby("currQr")["endtime"].idxmax()
            last   = sub.loc[idx]
            sub    = last[(last["resultado"] == "FAILED") & (last["endtime"] < cutoff)]
        else:
            sub = sub[sub["resultado"] == "FAILED"]

        if falla:
            sub = sub[sub["tipoFalla"] == falla]

        # Ordenar del mas nuevo al mas viejo
        sub = sub.sort_values("endtime", ascending=False)

        if export_csv:
            cols     = ["endtime", "currQr", "prevQr", "stationName",
                        "failureCode", "tipoFalla", "resultado", "notes"]
            csv_data = sub[[c for c in cols if c in sub.columns]].to_csv(index=False)
            return Response(csv_data, mimetype="text/csv",
                            headers={"Content-Disposition":
                                     "attachment; filename=failure_details.csv"})

        sub     = sub.head(DETAILS_MAX_RECORDS)
        records = [
            {
                "prevQr":      str(r.get("prevQr", "") or ""),
                "currQr":      str(r.get("currQr", "") or ""),
                "endtime":     (r["endtime"].strftime("%Y-%m-%d %H:%M") if pd.notna(r.get("endtime")) else ""),
                "tipoFalla":   str(r.get("tipoFalla", "")),
                "failureCode": str(r.get("failureCode", "")),
                "stationName": str(r.get("stationName", "")),
                "notes":       str(r.get("notes", "") or ""),
            }
            for _, r in sub.iterrows()
        ]
        return jsonify(records)

    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ─────────────────────────────────────────────────────────────
# RUTA: DETALLES PASS / FAIL
# ─────────────────────────────────────────────────────────────

@app.route("/api/passfail_details")
def api_passfail_details():
    dates         = _default_dates()
    date_from_str = request.args.get("from", dates["from"])
    date_to_str   = request.args.get("to",   dates["to"])
    product       = request.args.get("product", "all")
    station_id    = request.args.get("stationid", "all")
    resultado     = request.args.get("resultado", "")
    real_failures = request.args.get("real_failures", "false").lower() == "true"
    export_csv    = request.args.get("csv", "false").lower() == "true"
    hardware_id   = request.args.get("hardware_id", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return jsonify({"error": "System is booting"}), 503

        sub = _filter_by_dates(df, date_from_str, date_to_str)
        sub = _filter_by_product(sub, product)
        sub = _filter_by_station(sub, station_id)
        sub = _filter_by_hardware_id(sub, hardware_id)

        # Keep a copy prior to applying the real_failures/result filters
        original_sub = sub.copy()

        # Optional: filter by failure-count bin (1,2,3plus)
        fail_bin = request.args.get("fail_bin", None)
        qr_set = None
        if fail_bin:
            # Compute counts of FAILED rows per currQr within the selected scope
            try:
                counts = (
                    original_sub[original_sub["resultado"] == "FAILED"]
                    .dropna(subset=["currQr"])  # ignore missing QRs
                    .groupby("currQr").size()
                )
                if fail_bin == "1":
                    qr_set = counts[counts == 1].index.tolist()
                elif fail_bin == "2":
                    qr_set = counts[counts == 2].index.tolist()
                elif fail_bin in ("3", "3plus", "3+"):
                    qr_set = counts[counts >= 3].index.tolist()
            except Exception:
                qr_set = []

        if sub.empty:
            return jsonify([])

        if real_failures:
            sub = _apply_real_failures_filter(sub)

        if resultado:
            sub = sub[sub["resultado"] == resultado]

        # If requested, keep only rows whose currQr belongs to the selected bin
        if qr_set is not None:
            sub = sub[sub["currQr"].isin(qr_set)]

        sub = sub.sort_values("endtime", ascending=False)

        if export_csv:
            cols     = ["endtime", "currQr", "prevQr", "stationName",
                        "failureCode", "tipoFalla", "resultado", "notes"]
            csv_data = sub[[c for c in cols if c in sub.columns]].to_csv(index=False)
            return Response(csv_data, mimetype="text/csv",
                            headers={"Content-Disposition":
                                     f"attachment; filename=passfail_{resultado}.csv"})

        sub     = sub.head(DETAILS_MAX_RECORDS)
        records = [
            {
                "prevQr":      str(r.get("prevQr", "") or ""),
                "currQr":      str(r.get("currQr", "") or ""),
                "resultado":   str(r.get("resultado", "")),
                "tipoFalla":   str(r.get("tipoFalla", "")),
                "failureCode": str(r.get("failureCode", "")),
                "stationName": str(r.get("stationName", "")),
                "notes":       str(r.get("notes", "") or ""),
                "endtime":     (r["endtime"].strftime("%Y-%m-%d %H:%M")
                                if pd.notna(r.get("endtime")) else ""),
            }
            for _, r in sub.iterrows()
        ]
        return jsonify(records)

    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ─────────────────────────────────────────────────────────────
# RUTA: EXPORT CSV COMPLETO
# ─────────────────────────────────────────────────────────────

@app.route("/api/export_all_csv")
def api_export_all_csv():
    dates         = _default_dates()
    date_from_str = request.args.get("from", dates["from"])
    date_to_str   = request.args.get("to",   dates["to"])
    product       = request.args.get("product", "all")
    station_id    = request.args.get("stationid", "all")
    real_failures = request.args.get("real_failures", "false").lower() == "true"
    hardware_id   = request.args.get("hardware_id", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return "System is booting", 503

        sub = _filter_by_dates(df, date_from_str, date_to_str)
        sub = _filter_by_product(sub, product)
        sub = _filter_by_station(sub, station_id)
        sub = _filter_by_hardware_id(sub, hardware_id)

        if sub.empty:
            return "No data found", 404

        if real_failures:
            sub = _apply_real_failures_filter(sub)

        sub = sub.sort_values("endtime", ascending=False)
        cols = ["endtime", "currQr", "prevQr", "stationName", "producto",
                "failureCode", "tipoFalla", "resultado", "notes"]
        csv_data = sub[[c for c in cols if c in sub.columns]].to_csv(index=False)
        return Response(csv_data, mimetype="text/csv",
                        headers={"Content-Disposition":
                                 "attachment; filename=all_tests_export.csv"})

    except ValueError as ve:
        return str(ve), 400
    except Exception as exc:
        return str(exc), 500


# ─────────────────────────────────────────────────────────────
# RUTA: DATOS PARA EL PANEL DE REPORTE EN LA UI
# ─────────────────────────────────────────────────────────────

@app.route("/api/report_data")
def api_report_data():
    """
    Devuelve JSON con metricas resumidas, series de tiempo y top de fallas
    para el panel de reporte de la UI. Acepta los mismos filtros que /api/data.
    """
    dates         = _default_dates()
    date_from_str = request.args.get("from", dates["from"])
    date_to_str   = request.args.get("to",   dates["to"])
    product       = request.args.get("product", "all")
    station_id    = request.args.get("stationid", "all")
    hardware_id   = request.args.get("hardware_id", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return jsonify({"error": "System is booting"}), 503

        sub = _filter_by_dates(df, date_from_str, date_to_str)
        sub = _filter_by_product(sub, product)
        sub = _filter_by_station(sub, station_id)
        sub = _filter_by_hardware_id(sub, hardware_id)

        if sub.empty:
            return jsonify({"empty": True, "label": "No data"})

        label_parts = []
        if product and product != "all":
            label_parts.append(product)
        if station_id and station_id != "all":
            sname = sub["stationName"].iloc[0] if not sub.empty else f"Station {station_id}"
            label_parts.append(sname)
        label = " — ".join(label_parts) if label_parts else "All Products"

        data         = compute_report_data(sub, date_from_str, date_to_str)
        data["label"] = label
        return jsonify(data)

    except ValueError as ve:
        return jsonify({"error": str(ve)}), 400
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ─────────────────────────────────────────────────────────────
# RUTA: REPORTE EXCEL
# ─────────────────────────────────────────────────────────────

@app.route("/api/excel_report")
def api_excel_report():
    """
    Genera y descarga un reporte Excel con historial de yield.

    Parametros de query:
      product   — nombre del producto (ej. 'SPSF', 'Blade') o 'all'
      stationid — ID de estacion (int) o 'all'
      from      — fecha inicio YYYY-MM-DD
      to        — fecha fin   YYYY-MM-DD
    """
    dates         = _default_dates()
    date_from_str = request.args.get("from", dates["from"])
    date_to_str   = request.args.get("to",   dates["to"])
    product       = request.args.get("product", "all")
    station_id    = request.args.get("stationid", "all")
    hardware_id   = request.args.get("hardware_id", "all")

    try:
        df = _get_cached_df()
        if df is None:
            return "System is booting", 503

        sub = _filter_by_dates(df, date_from_str, date_to_str)
        sub = _filter_by_product(sub, product)
        sub = _filter_by_station(sub, station_id)
        sub = _filter_by_hardware_id(sub, hardware_id)

        if sub.empty:
            return "No data for the selected filters", 404

        # Etiqueta descriptiva para el reporte
        label_parts = []
        if product and product != "all":
            label_parts.append(product)
        if station_id and station_id != "all":
            sname = sub["stationName"].iloc[0] if not sub.empty else f"Station {station_id}"
            label_parts.append(sname)
        label = " — ".join(label_parts) if label_parts else "All Products"

        output = generate_excel_report(sub, label, date_from_str, date_to_str)

        safe_label = (label.replace(" ", "_").replace("—", "-")
                          .replace("/", "-").replace(" ", "_"))
        filename = f"report_{safe_label}_{date_from_str}_{date_to_str}.xlsx"

        return Response(
            output.read(),
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers={"Content-Disposition": f"attachment; filename={filename}"},
        )

    except ValueError as ve:
        return str(ve), 400
    except Exception as exc:
        return str(exc), 500


# ─────────────────────────────────────────────────────────────
# RUTAS DE DEBUG (solo para desarrollo)
# ─────────────────────────────────────────────────────────────

@app.route("/api/debug/<grupo>")
def api_debug(grupo):
    """Muestra notas crudas de filas FAILED para un grupo. Solo para desarrollo."""
    ids = GROUPS.get(grupo, [])
    if not ids:
        return jsonify({"error": f"Grupo '{grupo}' no existe. Opciones: {GROUP_NAMES}"}), 404

    default_from = (datetime.now() - timedelta(days=DEBUG_ENDPOINT_RANGE_DAYS)).strftime("%Y-%m-%d")
    ids_str      = ",".join(str(i) for i in ids)
    sql = f"""
        SELECT TOP 20
            s.stationName, e.failureCode, e.notes
        FROM trk.manufacturingEvents AS e
        INNER JOIN trk.stationConfig AS s ON e.stationid = s.stationid
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
                "station":     r[0],
                "failureCode": r[1],
                "tipoFalla":   extract_tipo_falla(r[1], r[2], r[0], grupo),
                "notes":       r[2],
            }
            for r in rows
        ]
        return jsonify(result)
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


@app.route("/api/debug_station/<int:sid>")
def api_debug_station(sid):
    """
    Diagnostico para una estacion especifica (consulta BD directamente).
    Util para estaciones que no aparecen en el dashboard.
    Ejemplo: /api/debug_station/406
    """
    sql = """
        SELECT TOP 10
            e.endtime, s.stationName, c.currQr, c.prevQr, e.failureCode,
            CASE
                WHEN e.failureCode LIKE 'ALL-PASS%'    THEN 'PASSED'
                WHEN e.failureCode LIKE '%-PASS-%'     THEN 'PASSED'
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
                "warning":   "Sin registros en BD (INNER JOIN qrCorrelation sin match).",
                "records":   [],
            })

        days_since = (datetime.now() - rows[0][0]).days if rows[0][0] else None
        in_cache   = days_since is not None and days_since <= HISTORICAL_LOAD_DAYS

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
            "stationid":         sid,
            "days_since_last":   days_since,
            "in_cache_window":   in_cache,
            "cache_window_days": HISTORICAL_LOAD_DAYS,
            "diagnosis": (
                "OK - datos dentro del rango del cache" if in_cache
                else f"FUERA DEL CACHE: ultimo registro hace {days_since} dias "
                     f"(HISTORICAL_LOAD_DAYS={HISTORICAL_LOAD_DAYS})"
            ),
            "records": records,
        })
    except Exception as exc:
        return jsonify({"error": str(exc)}), 500


# ─────────────────────────────────────────────────────────────
# PUNTO DE ENTRADA
# ─────────────────────────────────────────────────────────────

if __name__ == "__main__":
    # Carga inicial de historia (bloquea hasta completar)
    load_historical_data()

    # Hilo de refresco en segundo plano cada REFRESH_INTERVAL_MINUTES
    bg_thread = threading.Thread(target=background_update_worker, daemon=True)
    bg_thread.start()

    port = int(os.environ.get("PORT", DEFAULT_APP_PORT))
    print(f"\n  Manufacturing Dashboard iniciado")
    print(f"  Abre en tu navegador: http://localhost:{port}\n")
    app.run(host="0.0.0.0", port=port, debug=False, threaded=True)
