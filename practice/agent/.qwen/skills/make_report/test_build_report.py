"""Тесты для build_report: расчёт метрик и сортировка.

Дата формирования фиксирована (2026-11-10), чтобы результаты не зависели
от дня запуска. Задачи создаются относительно фиксированной даты.
"""
import sys
import unittest
from datetime import date

sys.path.insert(0, __file__.rsplit("/", 1)[0])

from build_report import (
    PRIORITIES,
    compute_metrics,
    normalize_task,
    sort_key,
)

REPORT_DATE = date(2026, 11, 10)
REPORT_ISO = "2026-11-10"

# Даты относительно REPORT_DATE
DUE_OVERDUE = "2026-11-01"    # раньше даты отчёта
DUE_TODAY = "2026-11-10"     # равен дате отчёта
DUE_AHEAD = "2026-11-20"     # позже даты отчёта


def task(**kwargs) -> dict:
    """Собирает задачу с любыми переопределениями полей."""
    defaults = {
        "id": 1,
        "title": "Тестовая задача",
        "done": False,
        "priority": "normal",
        "due": None,
        "created_at": "2026-11-01T10:00:00",
    }
    return {**defaults, **kwargs}


class ComputeMetricsTests(unittest.TestCase):

    def test_empty_list(self):
        metrics = compute_metrics([], REPORT_DATE)
        self.assertEqual(metrics["total"], 0)
        self.assertEqual(metrics["done"], 0)
        self.assertEqual(metrics["not_done"], 0)
        self.assertEqual(metrics["pct"], 0.0)
        self.assertEqual(metrics["overdue"], 0)
        self.assertEqual(metrics["due_today"], 0)
        self.assertEqual(metrics["no_due_not_done"], 0)
        for p in PRIORITIES:
            self.assertEqual(metrics["by_priority"][p], 0)
        self.assertEqual(metrics["other_count"], 0)

    def test_done_task(self):
        t = task(done=True)
        metrics = compute_metrics([t], REPORT_DATE)
        self.assertEqual(metrics["total"], 1)
        self.assertEqual(metrics["done"], 1)
        self.assertEqual(metrics["not_done"], 0)
        self.assertEqual(metrics["pct"], 100.0)
        self.assertEqual(metrics["overdue"], 0)
        self.assertEqual(metrics["no_due_not_done"], 0)

    def test_not_done_overdue(self):
        t = task(done=False, due=DUE_OVERDUE)
        metrics = compute_metrics([t], REPORT_DATE)
        self.assertEqual(metrics["not_done"], 1)
        self.assertEqual(metrics["overdue"], 1)
        self.assertEqual(metrics["due_today"], 0)
        self.assertEqual(metrics["pct"], 0.0)

    def test_not_done_due_today(self):
        t = task(done=False, due=DUE_TODAY)
        metrics = compute_metrics([t], REPORT_DATE)
        self.assertEqual(metrics["not_done"], 1)
        self.assertEqual(metrics["overdue"], 0)
        self.assertEqual(metrics["due_today"], 1)

    def test_not_done_no_due(self):
        t = task(done=False)
        metrics = compute_metrics([t], REPORT_DATE)
        self.assertEqual(metrics["not_done"], 1)
        self.assertEqual(metrics["overdue"], 0)
        self.assertEqual(metrics["due_today"], 0)
        self.assertEqual(metrics["no_due_not_done"], 1)

    def test_each_priority(self):
        for i, p in enumerate(PRIORITIES):
            t = task(id=i + 1, done=False, priority=p, due=DUE_OVERDUE)
            metrics = compute_metrics([t], REPORT_DATE)
            with self.subTest(priority=p):
                self.assertEqual(metrics["by_priority"][p], 1)
                self.assertEqual(metrics["overdue"], 1)
                self.assertEqual(metrics["other_count"], 0)

    def test_unknown_priority_counts_as_other(self):
        t = task(done=False, priority="critical")
        metrics = compute_metrics([t], REPORT_DATE)
        self.assertEqual(metrics["other_count"], 1)
        for p in PRIORITIES:
            self.assertEqual(metrics["by_priority"][p], 0)

    def test_priority_only_counts_not_done(self):
        high_done = task(done=True, priority="high")
        high_not_done = task(done=False, priority="high", id=2)
        metrics = compute_metrics([high_done, high_not_done], REPORT_DATE)
        self.assertEqual(metrics["by_priority"]["high"], 1)
        self.assertEqual(metrics["done"], 1)
        self.assertEqual(metrics["not_done"], 1)

    def test_mixed_metrics(self):
        tasks = [
            task(id=1, done=True, due=DUE_OVERDUE),
            task(id=2, done=False, due=DUE_OVERDUE),
            task(id=3, done=False, due=DUE_TODAY),
            task(id=4, done=False, due=DUE_AHEAD),
            task(id=5, done=False),
        ]
        metrics = compute_metrics(tasks, REPORT_DATE)
        self.assertEqual(metrics["total"], 5)
        self.assertEqual(metrics["done"], 1)
        self.assertEqual(metrics["not_done"], 4)
        self.assertEqual(metrics["pct"], 20.0)
        self.assertEqual(metrics["overdue"], 1)
        self.assertEqual(metrics["due_today"], 1)
        self.assertEqual(metrics["no_due_not_done"], 1)

    def test_rounding(self):
        # 1/3 = 33.333... -> 33.3
        tasks = [task(done=True), task(), task()]
        metrics = compute_metrics(tasks, REPORT_DATE)
        self.assertEqual(metrics["pct"], 33.3)
        # 2/3 = 66.666... -> 66.7
        tasks2 = [task(done=True), task(done=True), task()]
        metrics2 = compute_metrics(tasks2, REPORT_DATE)
        self.assertEqual(metrics2["pct"], 66.7)


class SortKeyTests(unittest.TestCase):

    def test_empty_list(self):
        self.assertEqual(sorted([], key=sort_key), [])

    def test_not_done_before_done(self):
        done = task(done=True, id=1)
        not_done = task(done=False, id=2)
        result = sorted([done, not_done], key=sort_key)
        self.assertEqual([t["id"] for t in result], [2, 1])

    def test_not_done_by_date(self):
        overdue = task(done=False, id=1, due=DUE_OVERDUE)
        today = task(done=False, id=2, due=DUE_TODAY)
        ahead = task(done=False, id=3, due=DUE_AHEAD)
        result = sorted([overdue, today, ahead], key=sort_key)
        self.assertEqual([t["id"] for t in result], [1, 2, 3])

    def test_no_date_at_end(self):
        no_due = task(done=False, id=1)
        with_due = task(done=False, id=2, due=DUE_AHEAD)
        result = sorted([no_due, with_due], key=sort_key)
        self.assertEqual([t["id"] for t in result], [2, 1])

    def test_done_group_by_date(self):
        done_1 = task(done=True, id=1, due=DUE_OVERDUE)
        done_2 = task(done=True, id=2, due=DUE_TODAY)
        result = sorted([done_2, done_1], key=sort_key)
        self.assertEqual([t["id"] for t in result], [1, 2])

    def test_no_date_in_done_group_at_end(self):
        done_no_due = task(done=True, id=1)
        done_with_due = task(done=True, id=2, due=DUE_AHEAD)
        result = sorted([done_no_due, done_with_due], key=sort_key)
        self.assertEqual([t["id"] for t in result], [2, 1])

    def test_then_by_id(self):
        a = task(done=False, id=2, due=DUE_AHEAD)
        b = task(done=False, id=1, due=DUE_AHEAD)
        result = sorted([a, b], key=sort_key)
        self.assertEqual([t["id"] for t in result], [1, 2])

    def test_mixed_sort_order(self):
        tasks = [
            task(id=1, done=True, due=DUE_OVERDUE),
            task(id=2, done=False),
            task(id=3, done=False, due=DUE_AHEAD),
            task(id=4, done=True, due=DUE_TODAY),
            task(id=5, done=False, due=DUE_OVERDUE),
        ]
        result = sorted(tasks, key=sort_key)
        self.assertEqual([t["id"] for t in result], [5, 3, 2, 1, 4])


class NormalizeTaskTests(unittest.TestCase):

    def test_normal_task(self):
        t = normalize_task(task(due=DUE_TODAY))
        self.assertEqual(t["due"], DUE_TODAY)
        self.assertEqual(t["done"], False)

    def test_missing_due(self):
        t = normalize_task(task(due=None))
        self.assertIsNone(t["due"])

    def test_empty_due_string(self):
        t = normalize_task(task(due="  "))
        self.assertIsNone(t["due"])

    def test_invalid_due_string(self):
        t = normalize_task(task(due="not-a-date"))
        self.assertIsNone(t["due"])

    def test_missing_title(self):
        t = normalize_task({**task(), "title": None})
        self.assertEqual(t["title"], "")

    def test_missing_priority(self):
        t = normalize_task({**task(), "priority": None})
        self.assertEqual(t["priority"], "")

    def test_unknown_priority_preserved(self):
        t = normalize_task(task(priority="critical"))
        self.assertEqual(t["priority"], "critical")

    def test_due_date_object_stripped(self):
        t = normalize_task(task(due=DUE_TODAY))
        self.assertEqual(t["due"], DUE_TODAY)


if __name__ == "__main__":
    unittest.main(verbosity=2)