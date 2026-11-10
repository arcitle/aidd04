#!/usr/bin/env python3
"""Детерминированная сборка Excel-отчёта по задачам TODO-приложения.

Получает полный список задач через REST API (GET /api/tasks), считает метрики
и создаёт Excel-файл с двумя листами: «Сводка» и «Задачи».

Пример:
    python3 build_report.py --url http://127.0.0.1:5000 --output report.xlsx
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import urllib.error
import urllib.request
from datetime import date, datetime

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    sys.stderr.write("Нужна библиотека openpyxl: pip install openpyxl\n")
    sys.exit(1)

DEFAULT_URL = "http://127.0.0.1:5000"
API_PATH = "/api/tasks"
PRIORITIES = ("high", "normal", "low")

STATUS_DONE = "выполнена"
STATUS_NOT_DONE = "не выполнена"
TERM_OVERDUE = "просрочена"
TERM_TODAY = "сегодня"
TERM_AHEAD = "впереди"
TERM_NONE = "не задан"

HEADER_FILL = PatternFill("solid", fgColor="D9E1F2")
STATUS_OK_FILL = PatternFill("solid", fgColor="C6EFCE")
STATUS_FAIL_FILL = PatternFill("solid", fgColor="FFEB9C")
TERM_OVERDUE_FILL = PatternFill("solid", fgColor="FFC7CE")

HEADER_FONT = Font(bold=True)
HEADER_ALIGNMENT = Alignment(wrap_text=True, vertical="center", horizontal="center")


def parse_args(argv=None) -> argparse.Namespace:
    p = argparse.ArgumentParser(
        description="Excel-отчёт по задачам TODO-приложения (GET /api/tasks)")
    p.add_argument(
        "--url",
        default=os.environ.get("TODO_API_URL", DEFAULT_URL),
        help="адрес TODO-приложения (по умолчанию: TODO_API_URL или http://127.0.0.1:5000)",
    )
    p.add_argument(
        "--output",
        default="report.xlsx",
        help="путь к создаваемому файлу отчёта (по умолчанию: report.xlsx)",
    )
    p.add_argument(
        "--date",
        default=None,
        help="дата формирования отчёта в формате YYYY-MM-DD (по умолчанию: сегодня)",
    )
    return p.parse_args(argv)


def fetch_tasks(url: str) -> list:
    """Получает список задач из GET {url}/api/tasks."""
    api_url = url.rstrip("/") + API_PATH
    try:
        with urllib.request.urlopen(api_url, timeout=15) as resp:
            raw = resp.read().decode("utf-8")
    except urllib.error.HTTPError as e:
        detail = ""
        try:
            detail = json.loads(e.read()).get("error", "")
        except (ValueError, OSError):
            pass
        msg = f"API вернул ошибку {e.code}" + (f": {detail}" if detail else "")
        raise RuntimeError(msg) from e
    except (urllib.error.URLError, TimeoutError) as e:
        reason = getattr(e, "reason", str(e))
        raise RuntimeError(
            f"TODO-приложение недоступно по адресу {api_url} ({reason})"
        ) from e

    try:
        data = json.loads(raw)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"Ответ API не является корректным JSON: {e}") from e

    if not isinstance(data, list):
        raise RuntimeError(
            f"Ожидался JSON-массив задач, получен {type(data).__name__}"
        )
    return data


def normalize_task(raw: dict) -> dict:
    """Копия задачи в рабочем формате; исходные данные задачи не изменяются."""
    due = raw.get("due")
    if isinstance(due, str):
        due = due.strip()
        if not due:
            due = None
        else:
            try:
                date.fromisoformat(due)
            except ValueError:
                due = None
    else:
        due = None

    title = raw.get("title")
    if not isinstance(title, str):
        title = "" if title is None else str(title)

    task_id = raw.get("id")
    if isinstance(task_id, bool) or not isinstance(task_id, int):
        task_id = 0

    priority = raw.get("priority")
    if priority is None:
        priority = ""

    return {
        "id": task_id,
        "title": title,
        "done": bool(raw.get("done")),
        "priority": priority,
        "due": due,
        "created_at": raw.get("created_at", ""),
    }


def sort_key(t: dict):
    """Сортировка: невыполненные сначала, затем выполненные;
    внутри группы — по дате дедлайна (без даты — в конец), затем по ID."""
    task_id = t["id"] if isinstance(t["id"], int) else 0
    return (
        0 if not t["done"] else 1,
        0 if t["due"] else 1,
        t["due"] or "",
        task_id,
    )


def compute_metrics(tasks: list, report_date: date) -> dict:
    """Считает метрики для списка задач на дату report_date."""
    total = len(tasks)
    done_count = sum(1 for t in tasks if t["done"])
    not_done_count = total - done_count
    pct = round(done_count / total * 100, 1) if total > 0 else 0.0

    report_iso = report_date.isoformat()
    overdue = sum(
        1 for t in tasks
        if not t["done"] and t["due"] and t["due"] < report_iso
    )
    due_today = sum(
        1 for t in tasks
        if not t["done"] and t["due"] == report_iso
    )
    no_due_not_done = sum(
        1 for t in tasks
        if not t["done"] and t["due"] is None
    )

    by_priority = {p: 0 for p in PRIORITIES}
    other = 0
    for t in tasks:
        if not t["done"]:
            p = t["priority"]
            if p in by_priority:
                by_priority[p] += 1
            else:
                other += 1

    return {
        "total": total,
        "done": done_count,
        "not_done": not_done_count,
        "pct": pct,
        "overdue": overdue,
        "due_today": due_today,
        "no_due_not_done": no_due_not_done,
        "by_priority": by_priority,
        "other_count": other,
    }


def _build_summary_sheet(ws, metrics: dict, report_date: date) -> None:
    rows = [
        ("Дата формирования", report_date, "текущая локальная дата"),
        ("Всего задач", metrics["total"],
         "число задач в полном списке GET /api/tasks"),
        ("Выполнено", metrics["done"], "задачи с done == true"),
        ("Невыполнено", metrics["not_done"], "задачи с done == false"),
        ("% выполненных", metrics["pct"],
         "выполнено / всего × 100, округление до одного знака; 0 при пустом списке"),
        ("Просроченные невыполненные", metrics["overdue"],
         "невыполненные задачи с дедлайном раньше даты формирования отчёта"),
        ("Срок сегодня и невыполненные", metrics["due_today"],
         "невыполненные задачи с дедлайном, равным дате формирования отчёта"),
        ("Без дедлайна и невыполненные", metrics["no_due_not_done"],
         "невыполненные задачи без дедлайна (дедлайн не задан)"),
    ]
    for p in PRIORITIES:
        rows.append((
            f"Невыполненные ({p})",
            metrics["by_priority"][p],
            f"невыполненные задачи с приоритетом {p}",
        ))
    if metrics["other_count"] > 0:
        rows.append((
            "Невыполненные (прочие)",
            metrics["other_count"],
            "невыполненные задачи с приоритетом, отличным от high/normal/low",
        ))

    ws["A1"] = "Сводка по задачам TODO"
    ws["A1"].font = Font(bold=True, size=14)
    ws.merge_cells("A1:C1")

    headers = ["Метрика", "Значение", "Правило расчёта"]
    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT

    for r, (name, value, rule) in enumerate(rows, 4):
        ws.cell(row=r, column=1, value=name)
        val_cell = ws.cell(row=r, column=2, value=value)
        if name == "Дата формирования":
            val_cell.number_format = "YYYY-MM-DD"
        ws.cell(row=r, column=3, value=rule)

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 56


def _build_tasks_sheet(ws, tasks: list, report_date: date) -> None:
    headers = ["ID", "Название", "Статус", "Приоритет", "Дедлайн",
               "Дней до дедлайна", "Срок", "Примечание"]

    report_iso = report_date.isoformat()
    sorted_tasks = sorted(tasks, key=sort_key)

    data_rows = []
    for t in sorted_tasks:
        due_str = t["due"]
        days = None
        if due_str:
            try:
                days = (date.fromisoformat(due_str) - report_date).days
            except ValueError:
                days = None

        if t["done"]:
            status = STATUS_DONE
            term = TERM_NONE
        elif not due_str:
            status = STATUS_NOT_DONE
            term = TERM_NONE
        elif due_str < report_iso:
            status = STATUS_NOT_DONE
            term = TERM_OVERDUE
        elif due_str == report_iso:
            status = STATUS_NOT_DONE
            term = TERM_TODAY
        else:
            status = STATUS_NOT_DONE
            term = TERM_AHEAD

        data_rows.append([
            t["id"],
            t["title"],
            status,
            t["priority"],
            due_str,
            days,
            term,
            t["created_at"],
        ])

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = HEADER_FONT
        cell.fill = HEADER_FILL
        cell.alignment = HEADER_ALIGNMENT

    for r, row in enumerate(data_rows, 2):
        for c_idx, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            if c_idx == 3:
                cell.fill = STATUS_OK_FILL if val == STATUS_DONE else STATUS_FAIL_FILL
            elif c_idx == 7 and val == TERM_OVERDUE:
                cell.fill = TERM_OVERDUE_FILL

    widths = [8, 40, 15, 12, 12, 18, 15, 25]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    ws.freeze_panes = "A2"
    last_row = len(data_rows) + 1
    ws.auto_filter.ref = f"A1:H{last_row}"


def build_workbook(tasks: list, report_date: date) -> Workbook:
    wb = Workbook()
    wb.remove(wb.active)

    metrics = compute_metrics(tasks, report_date)

    ws_summary = wb.create_sheet("Сводка")
    _build_summary_sheet(ws_summary, metrics, report_date)

    ws_tasks = wb.create_sheet("Задачи")
    _build_tasks_sheet(ws_tasks, tasks, report_date)

    return wb


def main(argv=None) -> None:
    args = parse_args(argv)

    if args.date:
        try:
            report_date = date.fromisoformat(args.date)
        except ValueError:
            raise SystemExit(f"Некорректная дата '{args.date}', ожидается формат YYYY-MM-DD")
    else:
        report_date = date.today()

    print(f"Получаю список задач из {args.url.rstrip('/') + API_PATH} ...")
    raw_tasks = fetch_tasks(args.url)
    tasks = [normalize_task(t) for t in raw_tasks]
    print(f"Загружено задач: {len(tasks)}")

    wb = build_workbook(tasks, report_date)
    wb.save(args.output)
    print(f"Отчёт сохранён: {os.path.abspath(args.output)}")


if __name__ == "__main__":
    main()