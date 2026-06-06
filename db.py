#!/usr/bin/env python3
"""
db.py — Conexiones a SQL Server y consultas de base de datos.
"""

import pyodbc
from config import (
    DB_SERVER, DB_NAME, DB_USER, DB_PASS,
    DB_NAME_QR, DB_CONNECT_TIMEOUT_SECONDS,
)


def _get_odbc_driver():
    """Detecta automaticamente ODBC Driver 18 o 17 segun lo instalado."""
    for drv in ("ODBC Driver 18 for SQL Server", "ODBC Driver 17 for SQL Server"):
        if drv in pyodbc.drivers():
            return drv
    return "ODBC Driver 17 for SQL Server"  # fallback


def _build_conn_str(database):
    return (
        f"DRIVER={{{_get_odbc_driver()}}};"
        f"SERVER={DB_SERVER};"
        f"DATABASE={database};"
        f"UID={DB_USER};"
        f"PWD={DB_PASS};"
        f"Encrypt=yes;TrustServerCertificate=no;"
    )


def connect():
    """Conexion a la base de datos principal de manufactura."""
    return pyodbc.connect(_build_conn_str(DB_NAME), timeout=DB_CONNECT_TIMEOUT_SECONDS)


def connect_qr():
    """Conexion a la base de datos de tracking de QR (trkprdshippapp)."""
    return pyodbc.connect(_build_conn_str(DB_NAME_QR), timeout=DB_CONNECT_TIMEOUT_SECONDS)


def lookup_model_numbers(qr_codes):
    """
    Consulta trk.qrmac_db y devuelve {qrcode: modelNumber} para los QRs dados.
    Filtra QRs nulos/vacios antes de consultar. Si la conexion falla, devuelve {}
    y los registros quedaran con fallback "SPSF" en prepare_dataframe.
    Usa parametros seguros en lotes de 500 para evitar inyeccion SQL.
    """
    _EMPTY_VALUES = {"", "nan", "None", "null", "undefined"}
    unique_qrs = [
        str(q) for q in qr_codes
        if q and str(q).strip() not in _EMPTY_VALUES
    ]
    if not unique_qrs:
        return {}

    result     = {}
    batch_size = 500
    try:
        conn = connect_qr()
        for i in range(0, len(unique_qrs), batch_size):
            chunk        = unique_qrs[i : i + batch_size]
            placeholders = ",".join("?" for _ in chunk)
            sql = (
                f"SELECT qrcode, modelNumber "
                f"FROM trk.qrmac_db WHERE qrcode IN ({placeholders})"
            )
            for row in conn.execute(sql, chunk).fetchall():
                if row[1] is not None:
                    result[str(row[0])] = str(row[1]).strip()
        conn.close()
    except Exception as exc:
        print(f"[WARNING] lookup_model_numbers: no se pudo consultar {DB_NAME_QR}: {exc}")

    return result
