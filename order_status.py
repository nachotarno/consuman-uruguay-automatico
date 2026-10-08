"""Estado compartido por la exportacion JSON/Excel y Google Sheets."""

from __future__ import annotations

from typing import Any

STATUS_NAMES = {
    0: "Sin iniciar",
    1: "Finalizada",
    2: "En progreso",
    3: "Sin iniciar",
    4: "En progreso",
    5: "Cancelada",
}


def status_name(order: dict[str, Any]) -> str:
    try:
        status = int(order.get("status"))
    except (TypeError, ValueError):
        status = None

    try:
        percentage = float(order.get("percentage") or 0)
    except (TypeError, ValueError):
        percentage = 0

    # Una cancelacion explicita no debe convertirse en finalizacion.
    if status == 5:
        return "Cancelada"
    # Consuman muestra Finalizada al 100 %, incluso sin endDate.
    # El avance de la OT prevalece sobre estados pendientes de sus tareas.
    if percentage >= 100 or status == 1:
        return "Finalizada"

    task_states = order.get("taskStates")
    if isinstance(task_states, list) and task_states:
        try:
            states = {int(value) for value in task_states}
        except (TypeError, ValueError):
            states = set()
        if states and states <= {0, 1, 2, 3}:
            if states == {1}:
                return "Finalizada"
            if states == {2}:
                return "Cancelada"
            if states == {0} and percentage <= 0:
                return "Sin iniciar"
            return "En progreso"

    if percentage > 0:
        return "En progreso"
    return STATUS_NAMES.get(status, str(order.get("status") or ""))
