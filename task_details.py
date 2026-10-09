"""Detalle de tareas para el reporte, sin agrupar ni perder realizaciones."""

from __future__ import annotations

import json
import unicodedata
from datetime import datetime
from pathlib import Path
from typing import Any
from zoneinfo import ZoneInfo

HEADERS = [
    "ID tarea", "OT", "Activo", "Parte", "Tarea", "Fecha planificada",
    "Fecha realizacion", "Estado tarea", "Referencia tarea", "Tipo OT",
]


def normalized(value: Any) -> str:
    text = unicodedata.normalize("NFKD", str(value or ""))
    return " ".join("".join(c for c in text if not unicodedata.combining(c)).lower().split())


def resolve_date_fields(columns: list[dict[str, Any]]) -> dict[str, str]:
    """Usa la configuración de columnas de Consuman, no nombres supuestos de API."""
    aliases = {
        "planned": {"fecha planificada", "fecha de planificacion", "fecha planificacion"},
        "realized": {"fecha realizacion", "fecha de realizacion"},
    }
    result = {}
    for kind, captions in aliases.items():
        fields = {
            str(col["dataField"])
            for col in columns
            if col.get("dataField") and normalized(col.get("caption")) in captions
        }
        if len(fields) != 1:
            raise ValueError(f"No se pudo identificar un único campo para fecha {kind}: {sorted(fields)}")
        result[kind] = fields.pop()
    if result["planned"] == result["realized"]:
        raise ValueError("Consuman asignó el mismo campo a planificación y realización.")
    return result


def text(value: Any) -> str:
    if isinstance(value, dict):
        return str(value.get("name") or value.get("description") or value.get("number") or "")
    return "" if value is None else str(value).strip()


def key(value: Any) -> str:
    return text(value)


def build_task_records(orders_data: Any, tasks_data: Any, columns: list[dict[str, Any]]) -> dict[str, Any]:
    orders = orders_data.get("data", []) if isinstance(orders_data, dict) else orders_data
    tasks = tasks_data.get("data", []) if isinstance(tasks_data, dict) else tasks_data
    if not isinstance(orders, list) or not isinstance(tasks, list):
        raise ValueError("Formato de órdenes/tareas no reconocido.")
    if not tasks:
        raise ValueError("No hay tareas. No se reemplaza el detalle existente.")
    fields = resolve_date_fields(columns)
    # También resolver los rótulos de los datos opcionales desde la grilla.
    display_fields = {normalized(c.get("caption")): str(c["dataField"]) for c in columns if c.get("dataField")}
    orders_by_id: dict[str, dict[str, Any]] = {}
    orders_by_number: dict[str, dict[str, Any]] = {}
    for order in orders:
        if not isinstance(order, dict):
            continue
        for index, value in ((orders_by_id, order.get("id")), (orders_by_number, order.get("number"))):
            if key(value):
                previous = index.get(key(value))
                if previous is not None and previous is not order:
                    raise ValueError(f"Referencia de OT ambigua: {value}")
                index[key(value)] = order
    if not any(isinstance(task, dict) and fields["realized"] in task for task in tasks):
        raise ValueError("El campo de Fecha realización de la grilla no aparece en la API de tareas. No se publica un detalle con fechas falsas.")
    records = []
    seen: dict[str, dict[str, Any]] = {}
    for task in tasks:
        if not isinstance(task, dict):
            raise ValueError("Consuman devolvió una tarea que no es un registro.")
        task_id = key(task.get("id"))
        if not task_id:
            raise ValueError("Una tarea no tiene ID. No se puede distinguir de otras de la misma OT.")
        if task_id in seen:
            if seen[task_id] != task:
                raise ValueError(f"Dos versiones contradictorias de la tarea {task_id}.")
            continue  # Solamente una repetición idéntica del mismo ID por paginación.
        seen[task_id] = task
        order = orders_by_id.get(key(task.get("workOrderId"))) or orders_by_number.get(key(task.get("workOrderNumber")))
        ot_field = display_fields.get("ot", "workOrderNumber")
        ot = text(order.get("number")) if order else text(task.get(ot_field))
        if not ot:
            raise ValueError(f"No se pudo identificar la OT de la tarea {task_id}.")
        planned = task.get(fields["planned"])
        if not planned:
            raise ValueError(f"La tarea {task_id} no tiene fecha planificada.")
        # Una realización ausente sigue siendo ausente aunque la OT marque 100 %.
        realized = task.get(fields["realized"])
        asset = text(task.get(display_fields.get("activo", "equipment"))) or (text(order.get("equipment")) if order else "")
        task_name = text(task.get(display_fields.get("tarea", "task")))
        if not asset or not task_name:
            raise ValueError(f"Falta activo o tarea en registro {task_id}.")
        part = text(task.get(display_fields.get("parte", "part")))
        reference_field = next((v for k, v in display_fields.items() if k.startswith("numero /") or k.startswith("numero/")), "number")
        records.append({
            "id": task_id, "workOrderNumber": ot, "equipment": asset,
            "part": part, "task": task_name, "plannedDate": planned,
            "realizationDate": realized or None, "state": task.get("state"),
            "reference": text(task.get(reference_field)),
            "orderType": text(order.get("type")) if order else "",
        })
    return {"dateFields": fields, "records": records}


def sheet_date(value: Any, timezone_name: str = "America/Montevideo") -> float | str:
    if value in (None, ""):
        return ""
    raw = str(value).strip()
    try:
        date = datetime.fromisoformat(raw.replace("Z", "+00:00"))
    except ValueError:
        date = None
        for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%d/%m/%Y"):
            try:
                date = datetime.strptime(raw, fmt)
                break
            except ValueError:
                pass
        if date is None:
            raise ValueError(f"Fecha de tarea no reconocida: {raw!r}")
    if date.tzinfo is not None:
        date = date.astimezone(ZoneInfo(timezone_name)).replace(tzinfo=None)
    # Número de fecha local explícito. No depende de la zona horaria del Sheet.
    return (date - datetime(1899, 12, 30)).total_seconds() / 86400


def load_task_rows(path: Path) -> list[list[Any]]:
    if not path.is_file():
        raise ValueError("Falta el detalle de tareas; no se actualiza Google Sheets.")
    payload = json.loads(path.read_text(encoding="utf-8"))
    rows = []
    for record in payload.get("records", []):
        state = text(record.get("state"))
        rows.append([
            text(record.get("id")), text(record.get("workOrderNumber")),
            text(record.get("equipment")), text(record.get("part")),
            text(record.get("task")), sheet_date(record.get("plannedDate")),
            sheet_date(record.get("realizationDate")), state,
            text(record.get("reference")), text(record.get("orderType")),
        ])
    if not rows:
        raise ValueError("El detalle de tareas está vacío; no se reemplaza Google Sheets.")
    return rows
