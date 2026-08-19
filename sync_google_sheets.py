from __future__ import annotations

import json
import os
import sys
from pathlib import Path
from typing import Any

from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

SPREADSHEET_ID = os.getenv(
    "GOOGLE_SHEET_ID", "1fwuWOnZeiOaY0g8RV0VZVAgJqwAsLjVujnZ4z0gySiA"
).strip()
SHEET_NAME = os.getenv("GOOGLE_SHEET_TAB", "Ordenes de trabajo").strip()
INPUT_PATH = Path(__file__).resolve().parent / "ordenes_trabajo.json"
HEADERS = ["OT", "Activo", "Tarea", "Fecha planificada", "Estado"]
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

STATUS_NAMES = {
    0: "Sin iniciar",
    1: "Finalizada",
    2: "En progreso",
    3: "Sin iniciar",
    5: "Cancelada",
}


def status_name(order: dict[str, Any]) -> str:
    try:
        status = int(order.get("status"))
    except (TypeError, ValueError):
        return str(order.get("status") or "")
    try:
        percentage = float(order.get("percentage") or 0)
    except (TypeError, ValueError):
        percentage = 0
    if status == 0 and percentage >= 100 and order.get("endDate"):
        return "Finalizada"
    return STATUS_NAMES.get(status, str(order.get("status") or ""))


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
    # Keep the value as returned by Consuman. Google Sheets parses ISO dates reliably
    # with USER_ENTERED while preserving the existing column formatting.
    return str(order.get("date") or "")


def load_rows() -> list[list[str]]:
    if not INPUT_PATH.is_file():
        raise RuntimeError(f"No existe {INPUT_PATH.name}; no se modifica Google Sheets.")
    root = json.loads(INPUT_PATH.read_text(encoding="utf-8"))
    orders = root if isinstance(root, list) else root.get("data") or root.get("rows") or root.get("items") or []
    if not isinstance(orders, list):
        raise RuntimeError("El JSON de Consuman no contiene una lista de ordenes.")

    rows: list[list[str]] = []
    for order in orders:
        if not isinstance(order, dict):
            continue
        rows.append([
            "" if order.get("number") is None else str(order.get("number")),
            str(order.get("equipment") or ""),
            task_name(order),
            planned_date(order),
            status_name(order),
        ])
    return rows


def credentials() -> Credentials:
    raw = os.getenv("GOOGLE_SERVICE_ACCOUNT_JSON", "").strip()
    if not raw:
        raise RuntimeError("Falta el secret GOOGLE_SERVICE_ACCOUNT_JSON.")
    try:
        info = json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError("GOOGLE_SERVICE_ACCOUNT_JSON no contiene JSON valido.") from exc
    return Credentials.from_service_account_info(info, scopes=SCOPES)


def get_sheet_id(service) -> int:
    meta = service.spreadsheets().get(
        spreadsheetId=SPREADSHEET_ID,
        fields="sheets(properties(sheetId,title,gridProperties(frozenRowCount)))",
    ).execute()
    for sheet in meta.get("sheets", []):
        props = sheet.get("properties", {})
        if props.get("title") == SHEET_NAME:
            return int(props["sheetId"])
    raise RuntimeError(f"No existe la pestana {SHEET_NAME!r} en el spreadsheet indicado.")


def sync() -> None:
    rows = load_rows()
    creds = credentials()
    service = build("sheets", "v4", credentials=creds, cache_discovery=False)
    sheet_id = get_sheet_id(service)

    # Clear only values, not formatting, widths, frozen rows, filters, or other sheet properties.
    service.spreadsheets().values().clear(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A2:E",
        body={},
    ).execute()

    values = [HEADERS] + rows
    service.spreadsheets().values().update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A1:E{len(values)}",
        valueInputOption="USER_ENTERED",
        body={"values": values},
    ).execute()

    # Ensure header is frozen; this does not touch widths or formatting.
    service.spreadsheets().batchUpdate(
        spreadsheetId=SPREADSHEET_ID,
        body={
            "requests": [
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
        },
    ).execute()

    check_end = min(len(rows) + 1, 6)
    sample = service.spreadsheets().values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"'{SHEET_NAME}'!A1:E{check_end}",
    ).execute().get("values", [])

    if not sample or sample[0] != HEADERS:
        raise RuntimeError(f"Verificacion de encabezados fallo. Leido: {sample[:1]!r}")

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
