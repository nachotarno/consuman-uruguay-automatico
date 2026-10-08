import unittest

from order_status import status_name


class StatusTests(unittest.TestCase):
    def test_completed_progress_without_close_date(self):
        for status in (0, 3, "3", None):
            with self.subTest(status=status):
                self.assertEqual(
                    status_name({"status": status, "percentage": "100", "taskStates": [0]}),
                    "Finalizada",
                )

    def test_task_completion_survives_pending_order_status(self):
        self.assertEqual(status_name({"status": 3, "taskStates": [1]}), "Finalizada")

    def test_partial_and_pending_orders(self):
        self.assertEqual(
            status_name({"status": 3, "percentage": 50, "taskStates": [0]}),
            "En progreso",
        )
        self.assertEqual(
            status_name({"status": 3, "percentage": 0, "taskStates": [0]}),
            "Sin iniciar",
        )

    def test_cancelled_orders_and_tasks(self):
        self.assertEqual(status_name({"status": 5, "percentage": 100}), "Cancelada")
        self.assertEqual(status_name({"status": 3, "taskStates": [2]}), "Cancelada")

    def test_mixed_tasks(self):
        self.assertEqual(status_name({"status": 3, "taskStates": [0, 1]}), "En progreso")

    def test_legacy_status_and_invalid_optional_fields(self):
        self.assertEqual(status_name({"status": "1"}), "Finalizada")
        self.assertEqual(
            status_name({"status": "3", "percentage": "N/A", "taskStates": ["bad"]}),
            "Sin iniciar",
        )
        self.assertEqual(status_name({"status": 4}), "En progreso")


if __name__ == "__main__":
    unittest.main()
