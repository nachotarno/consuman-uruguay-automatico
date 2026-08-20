from __future__ import annotations

import json
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

import google.auth
from google.auth.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SPREADSHEET_ID = os.getenv(
    "GOOGLE_SHEET_ID", "1FuowVEsGZDKUjHedfOgPuMSSSFEmfr1itIgW-TMuQEM"
).strip()
SHEET_NAME = os.getenv("GOOGLE_SHEET_TAB", "Órdenes de trabajo").strip()
INPUT_PATH = Path(__file__).resolve().parent / "ordenes_trabajo_proevo.json"

HEADERS = [
    "Número",
    "Fecha de emisión",
    "Generado por",
    "Tipo",
    "Activo",
    "Tipo de mantenimiento",
    "Estadio",
    "Nivel de urgencia",
    "Estado",
    "%",
    "Porcentaje de asignación",
]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

STATUS_NAMES = {
    0: "Sin iniciar",
    1: "Finalizada",
    2: "En progreso",
    3: "Sin iniciar",
    4: "En progreso",
    5: "Cancelada",
}

TYPE_NAMES = {
    1: "Interna",
    2: "Externa",
}

ACTIVITY_TYPE_NAMES = {
    "1": "Solicitud",
}


def _number(value: Any, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def status_name(order: dict[str, Any]) -> str:
    try:
        status = int(order.get("status"))
    except (TypeError, ValueError):
        return str(order.get("status") or "")

    percentage = _number(order.get("percentage"), 0)
    if status == 0 and percentage >= 100 and order.get("endDate"):
        return "Finalizada"
    return STATUS_NAMES.get(status, str(order.get("status") or ""))


def type_name(order: dict[str, Any]) -> str:
    raw = order.get("type")
    try:
        key = int(raw)
    except (TypeError, ValueError):
        return str(raw or "")
    return TYPE_NAMES.get(key, str(raw or ""))


def activity_type_name(order: dict[str, Any]) -> str:
    raw = str(order.get("activityType") or "").strip()
    return ACTIVITY_TYPE_NAMES.get(raw, raw)


def google_date_serial(value: Any) -> float | str:
    """Convierte fecha ISO de Consuman a serial de Google Sheets."""
    raw = str(value or "").strip()
    if not raw:
        return ""
    try:
        dt = datetime.fromisoformat(raw.replace("Z", "+00:00"))
        # Google Sheets/Excel usan 1899-12-30 como base práctica.
        base = datetime(1899, 12, 30, tzinfo=dt.tzinfo)
        delta = dt - base
        return delta.days + (delta.seconds + delta.microseconds / 1_000_000) / 86400
    except ValueError:
        return raw


def load_rows() -> list[list[Any]]:
    if not INPUT_PATH.is_file():
        raise RuntimeError(
            f"No existe {INPUT_PATH.name}; no se modifica Google Sheets."
        )

    root = json.loads(INPUT_PATH.read_text(encoding="utf-8"))
    if isinstance(root, list):
        orders = root
    elif isinstance(root, dict):
        orders = root.get("data") or root.get("rows") or root.get("items") or []
    else:
        orders = []

    if not isinstance(orders, list):
        raise RuntimeError("El JSON de Consuman no contiene una lista de órdenes.")

    rows: list[list[Any]] = []
    for order in orders:
        if not isinstance(order, dict):
            continue

        rows.append(
            [
                order.get("number") if order.get("number") is not None else "",
                google_date_serial(order.get("date")),
                str(order.get("generatedUser") or ""),
                type_name(order),
                str(order.get("equipment") or ""),
                activity_type_name(order),
                str(order.get("stadium") or ""),
                str(order.get("levelOfUrgency") or ""),
                status_name(order),
                _number(order.get("percentage"), 0),
                _number(order.get("percentageAssingation"), 0) / 100.0,
            ]
        )

    if not rows:
        raise RuntimeError(
            "La extracción no devolvió órdenes. Para proteger la hoja válida, no se borra ni reemplaza nada."
        )
    return rows


def credentials() -> Credentials:
    creds, _ = google.auth.default(scopes=SCOPES)
    if not creds.valid:
        creds.refresh(Request())
    return creds


def get_sheet_id(service) -> int:
    meta = (
        service.spreadsheets()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            fields="sheets(properties(sheetId,title,gridProperties(frozenRowCount)))",
        )
        .execute()
    )
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == SHEET_NAME:
            return int(props["sheetId"])
    raise RuntimeError(
        f"No existe la pestaña {SHEET_NAME!r} en el spreadsheet indicado."
    )


def sync() -> None:
    rows = load_rows()
    creds = credentials()
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    sheet_id = get_sheet_id(service)

    existing_header = (
        service.spreadsheets()
        .values()
        .get(spreadsheetId=SPREADSHEET_ID, range=f"'{SHEET_NAME}'!A1:K1")
        .execute()
        .get("values", [])
    )

    if existing_header and existing_header[0] != HEADERS:
        raise RuntimeError(
            f"Encabezados inesperados. Se esperaba {HEADERS!r} y se leyó {existing_header[0]!r}. "
            "No se modifica la hoja."
        )

    # Borra solo valores; conserva formato, filtros, anchos y estilos existentes.
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A2:K",
        body={},
    ).execute()

    values = [HEADERS] + rows
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A1:K{len(values)}",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

    requests: list[dict[str, Any]] = [
        {
            "updateSheetProperties": {
                "properties": {
                    "sheetId": sheet_id,
                    "gridProperties": {"frozenRowCount": 1},
                },
                "fields": "gridProperties.frozenRowCount",
            }
        }
    ]

    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={"requests": requests},
    ).execute()

    check_end = min(len(rows) + 1, 6)
    sample = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A1:K{check_end}",
        )
        .execute()
        .get("values", [])
    )

    if not sample or sample[0] != HEADERS:
        raise RuntimeError(f"Falló la verificación de encabezados. Leído: {sample[:1]!r}")

    print(f"OK: {len(rows)} filas sincronizadas en {SHEET_NAME!r}.")
    print("Muestra verificada:")
    for row in sample[:6]:
        print(row)


def main() -> int:
    try:
        sync()
        return 0
    except (RuntimeError, HttpError, OSError, ValueError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
