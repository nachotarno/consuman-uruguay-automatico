from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

from order_status import status_name

import google.auth
from google.auth.credentials import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SPREADSHEET_ID = os.getenv(
    "GOOGLE_SHEET_ID", "1fwuWOnZeiOaY0g8RV0VZVAgJqwAsLjVujnZ4z0gySiA"
).strip()
SHEET_NAME = os.getenv("GOOGLE_SHEET_TAB", "Ordenes de trabajo").strip()
INPUT_PATH = Path(__file__).resolve().parent / "ordenes_trabajo.json"

OLD_HEADERS = ["OT", "Activo", "Tarea", "Fecha planificada", "Estado"]
HEADERS = ["OT", "Activo", "Tarea", "Fecha planificada", "Estado", "Fecha cierre"]

SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

def task_name(order: dict[str, Any]) -> str:
    tasks = order.get("tasks")
    if isinstance(tasks, list) and tasks:
        unique: list[str] = []
        for item in tasks:
            value = str(item or "").strip()
            if value and value not in unique:
                unique.append(value)
        return " | ".join(unique)
    return str(order.get("task") or "").strip()


def planned_date(order: dict[str, Any]) -> str:
    return str(order.get("date") or "")


def close_date(order: dict[str, Any]) -> str:
    # Consuman devuelve la fecha de cierre en endDate.
    return str(order.get("endDate") or "")


def load_rows() -> list[list[str]]:
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

    rows: list[list[str]] = []
    for order in orders:
        if not isinstance(order, dict):
            continue
        rows.append(
            [
                "" if order.get("number") is None else str(order.get("number")),
                str(order.get("equipment") or ""),
                task_name(order),
                planned_date(order),
                status_name(order),
                close_date(order),
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


def sync() -> None:
    rows = load_rows()
    creds = credentials()
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)

    existing_header = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A1:F1",
        )
        .execute()
        .get("values", [])
    )

    if existing_header:
        read_header = existing_header[0]

        legacy_columna1 = OLD_HEADERS.copy()
        legacy_columna1[0] = "Columna 1"

        legacy_columna1_fecha_cierre = HEADERS.copy()
        legacy_columna1_fecha_cierre[0] = "Columna 1"

        valid_headers = (
            OLD_HEADERS,
            HEADERS,
            legacy_columna1,
            legacy_columna1_fecha_cierre,
        )

        if read_header not in valid_headers:
            raise RuntimeError(
                f"Encabezados inesperados. Se esperaba {HEADERS!r} "
                f"(o una versión anterior compatible) y se leyó {read_header!r}. "
                "No se modifica la hoja."
            )

    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A2:F",
        body={},
    ).execute()

    values = [HEADERS] + rows

    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A1:F{len(values)}",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

    # Solo modifica valores; no toca tablas, filtros ni formato.

    check_end = min(len(rows) + 1, 6)
    sample = (
        service.spreadsheets()
        .values()
        .get(
            spreadsheetId=SPREADSHEET_ID,
            range=f"'{SHEET_NAME}'!A1:F{check_end}",
        )
        .execute()
        .get("values", [])
    )

    if not sample or sample[0] != HEADERS:
        raise RuntimeError(f"Falló la verificación de encabezados. Leído: {sample[:1]!r}")

    print(f"OK: {len(rows)} filas sincronizadas en {SHEET_NAME!r}, incluyendo Fecha cierre.")
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
