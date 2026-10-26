import json
import tempfile
import unittest
from pathlib import Path

from task_details import build_task_records, load_task_rows, resolve_date_fields, sheet_date

COLUMNS = [{"caption": "Fecha planificada", "dataField": "plannedDate"}, {"caption": "Fecha realización", "dataField": "actualExecution"}]
ORDERS = [{"id": 10, "number": 4833, "equipment": "CP SIM 3068", "type": "ME"}, {"id": 11, "number": 4832, "equipment": "CP SIM 3021", "type": "MP"}]


class TaskDetailsTests(unittest.TestCase):
    def task(self, id, order=10, realized=None):
        return {"id": id, "workOrderId": order, "task": "CALIBRACION", "plannedDate": "2026-10-08T03:00:00Z", "actualExecution": realized, "state": 1 if realized else 0}

    def test_two_tasks_same_order_keep_completed_and_pending(self):
        result = build_task_records(ORDERS, [self.task(4081, realized="2026-10-08T13:06:00Z"), self.task(4085)], COLUMNS)
        self.assertEqual(len(result["records"]), 2)
        self.assertEqual([r["workOrderNumber"] for r in result["records"]], ["4833", "4833"])
        self.assertIsNone(result["records"][1]["realizationDate"])

    def test_realization_next_day_not_replaced_by_plan(self):
        result = build_task_records(ORDERS, [self.task(5000, order=11, realized="2026-10-09T10:48:00Z")], COLUMNS)
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp)/"tasks.json"
            path.write_text(json.dumps(result), encoding="utf-8")
            row = load_task_rows(path)[0]
        self.assertEqual(row[5], 46303)
        self.assertAlmostEqual(row[6], 46304+7.8/24)

    def test_missing_date_does_not_close_task_even_final_order(self):
        orders = [dict(ORDERS[0], percentage=100, status=1)]
        task = self.task(4085)
        task["state"] = 1
        result = build_task_records(orders, [task], COLUMNS)
        self.assertIsNone(result["records"][0]["realizationDate"])

    def test_same_task_name_different_ids_not_deduplicated(self):
        tasks = [self.task(1), self.task(2)]
        self.assertEqual(len(build_task_records(ORDERS, tasks, COLUMNS)["records"]), 2)

    def test_only_identical_pagination_duplicate_removed(self):
        task = self.task(1)
        self.assertEqual(len(build_task_records(ORDERS, [task, dict(task)], COLUMNS)["records"]), 1)
        with self.assertRaises(ValueError):
            build_task_records(ORDERS, [task, dict(task, actualExecution="2026-10-08")], COLUMNS)

    def test_field_names_resolved_from_ui_and_not_guessed(self):
        self.assertEqual(resolve_date_fields(COLUMNS)["realized"], "actualExecution")
        with self.assertRaises(ValueError):
            resolve_date_fields([COLUMNS[0]])
        with self.assertRaises(ValueError):
            build_task_records(ORDERS, [{"id": 1, "workOrderId": 10, "plannedDate": "2026-10-08", "task": "x"}], COLUMNS)

    def test_two_orders_same_internal_both_retained_for_report_grouping(self):
        orders = [ORDERS[0], dict(ORDERS[1], equipment=ORDERS[0]["equipment"])]
        result = build_task_records(orders, [self.task(1), self.task(2, order=11)], COLUMNS)
        self.assertEqual([r["equipment"] for r in result["records"]], ["CP SIM 3068"]*2)
        self.assertEqual(len({r["workOrderNumber"] for r in result["records"]}), 2)

    def test_incomplete_record_is_skipped_not_fatal(self):
        good = self.task(1, realized="2026-10-08T13:06:00Z")
        bad = {"id": 999, "workOrderNumber": "7777", "plannedDate": "2026-10-08T03:00:00Z", "actualExecution": "2026-10-08T13:06:00Z", "task": "", "state": 1}
        result = build_task_records(ORDERS, [good, bad], COLUMNS)
        self.assertEqual(len(result["records"]), 1)
        self.assertEqual(len(result["skipped"]), 1)
        self.assertEqual(result["skipped"][0]["id"], "999")

    def test_bad_date_and_missing_task_id_fail(self):
        with self.assertRaises(ValueError):
            sheet_date("fecha inválida")
        with self.assertRaises(ValueError):
            build_task_records(ORDERS, [dict(self.task(1), id=None)], COLUMNS)
        self.assertEqual(sheet_date(None), "")


if __name__ == "__main__":
    unittest.main()
