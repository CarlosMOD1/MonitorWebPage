#!/usr/bin/env python3
"""
config.py — Constantes globales, grupos de estaciones y query SQL base.
"""

import re
import os
from dotenv import load_dotenv

load_dotenv()

# ─────────────────────────────────────────────────────────────
# CREDENCIALES DE BASE DE DATOS
# ─────────────────────────────────────────────────────────────
DB_SERVER  = os.getenv("DB_SERVER")
DB_NAME    = os.getenv("DB_NAME")
DB_USER    = os.getenv("DB_USER")
DB_PASS    = os.getenv("DB_PASS")
DB_NAME_QR = "trkprdshippapp"

# ─────────────────────────────────────────────────────────────
# MAPEO modelNumber → producto
# Agregar nuevos modelos aqui.
# ─────────────────────────────────────────────────────────────
MODEL_NUMBER_TO_PRODUCT = {
    "GBP-3001": "SPSF",
    "GBP-3003": "Blade",
}

# Estaciones compartidas SPSF/Blade que NO pueden diferenciarse por modelNumber.
# Sus registros se duplican: uno como SPSF y otro como Blade.
UNDIFF_STATIONS = {406, 393}

# ─────────────────────────────────────────────────────────────
# CONSTANTES DE COMPORTAMIENTO DEL DASHBOARD
# ─────────────────────────────────────────────────────────────
HISTORICAL_LOAD_DAYS       = 30    # Dias de historia a cargar al iniciar
RECENT_REFRESH_DAYS        = 2     # Ventana de refresco en segundo plano
REFRESH_INTERVAL_MINUTES   = 15    # Frecuencia del hilo de refresco
REAL_FAILURE_HOURS         = 1     # Horas minimas para considerar una falla "real"
DEFAULT_QUERY_RANGE_DAYS   = 30    # Rango de fechas por defecto en las rutas
DEBUG_ENDPOINT_RANGE_DAYS  = 7     # Rango de fechas para endpoints de debug
DETAILS_MAX_RECORDS        = 500   # Maximo de filas en respuestas de detalle
DB_CONNECT_TIMEOUT_SECONDS = 60    # Timeout de conexion a DB
DEFAULT_APP_PORT           = 5000  # Puerto del servidor Flask

# ─────────────────────────────────────────────────────────────
# GRUPOS DE ESTACIONES
# ─────────────────────────────────────────────────────────────
_SPSF_STATIONS = [
    406, 393, 178, 183, 184, 185, 186, 187, 244,
    201, 202, 190, 191, 192, 193, 194, 195, 196,
    199, 208, 211, 215, 220, 232,
]

GROUPS = {
    "SPSF":        _SPSF_STATIONS,
    "Blade":       _SPSF_STATIONS,   # Comparte estaciones; se diferencia por modelNumber
    "C-Track":     [238, 239, 241, 333],
    "Solar":       [279, 275, 277, 278, 279, 280, 281, 284],
    "Black Box":   [203, 204, 233, 237, 245],
    "OPAL 3":      [260, 261, 262, 263, 264, 265, 266],
    "Black Box 2": [297, 298, 300, 301, 302, 303, 304, 305, 306],
    "Opal 4":      [289, 290, 291, 292, 293, 294, 295],
    "Lima":        [254, 90, 247, 251],
}

# Conjunto de todas las estaciones compartidas SPSF/Blade
SPSF_BLADE_STATIONS      = set(_SPSF_STATIONS)
# Estaciones donde SI podemos resolver el producto via modelNumber
SPSF_BLADE_DIFF_STATIONS = SPSF_BLADE_STATIONS - UNDIFF_STATIONS

# STATION_TO_GROUP: primera asignacion gana (SPSF gana sobre Blade en estaciones compartidas)
STATION_TO_GROUP    = {}
ALL_STATION_IDS_SET = set()
for _grp, _ids in GROUPS.items():
    for _sid in _ids:
        if _sid not in STATION_TO_GROUP:
            STATION_TO_GROUP[_sid] = _grp
        ALL_STATION_IDS_SET.add(_sid)

ALL_STATION_IDS = list(ALL_STATION_IDS_SET)
IDS_STR         = ",".join(str(s) for s in ALL_STATION_IDS)
GROUP_NAMES     = list(GROUPS.keys())

# Orden de estaciones por producto: (producto, stationid) → posicion en la lista original.
# Usado en compute_stats para respetar el orden de la UI.
STATION_ORDER = {}
for _grp, _ids in GROUPS.items():
    for _pos, _sid in enumerate(_ids):
        STATION_ORDER[(_grp, _sid)] = _pos

# ─────────────────────────────────────────────────────────────
# QUERY SQL BASE
# ─────────────────────────────────────────────────────────────
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
# PATRONES REGEX REUTILIZABLES
# ─────────────────────────────────────────────────────────────

# Codigo estructurado tipo: ALL-FAIL-000, OPAL-BATT-000, CT-PCBA-002
CODE_RE = re.compile(r'^[A-Z0-9]+-[A-Z0-9]+-\d+$')

# Metadata de dispositivo: QR=..., MAC=..., UID=...
METADATA_RE = re.compile(r'^[A-Z_]{1,6}=')
