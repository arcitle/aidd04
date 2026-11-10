#!/usr/bin/env python3
"""Generate an Excel report from the TODO MCP server data."""
import json
import os
from datetime import date

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

OUTPUT_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "todo_report.xlsx")

# Dark theme colors
BG = "FF1e1f24"
BG_ALT = "FF26272c"
HEADER_BG = "FF0f3d5c"
TEXT = "FFe6e1cf"
MUTED = "FF94a3b8"
ACCENT = "FF4cc9f0"
DONE_GREEN = "FF56cc9d"
OVERDUE_RED = "FFf17285"
TODAY_YELLOW = "FFf9d368"
HIGH_ORANGE = "FFf97316"

THIN = Side(style="thin", color="FF3a3b40")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)


def parse_due(raw):
    if not raw:
        return None
    return date.fromisoformat(raw)


def days_until(d):
    if d is None:
        return None
    return (d - date.today()).days


def category(d):
    if d is None:
        return "не задан"
    diff = days_until(d)
    if diff < 0:
        return "просрочена"
    if diff == 0:
        return "сегодня"
    return "впереди"


def category_color(cat):
    if cat == "просрочена":
        return OVERDUE_RED
    if cat == "сегодня":
        return TODAY_YELLOW
    if cat == "впереди":
        return ACCENT
    return MUTED


def priority_rank(p):
    return {"high": 0, "normal": 1, "low": 2}.get(p, 3)


def status_text(done):
    return "выполнено" if done else "не выполнено"


def style_header(ws):
    for cell in ws[1]:
        cell.font = Font(bold=True, color=TEXT, size=11)
        cell.fill = PatternFill("solid", fgColor=HEADER_BG)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER


def style_rows(ws, start=2):
    row = start
    while row <= ws.max_row:
        is_odd = (row - start) % 2 == 0
        fill = PatternFill("solid", fgColor=BG_ALT) if is_odd else PatternFill("solid", fgColor=BG)
        for cell in ws[row]:
            cell.fill = fill
            cell.font = Font(color=TEXT, size=11)
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            cell.border = BORDER
        row += 1


def build_summary(tasks):
    total = len(tasks)
    done = sum(1 for t in tasks if t["done"])
    todo = total - done
    pct = round(done / total * 100, 1) if total else 0.0

    overdue = 0
    today = 0
    no_due = 0
    high = 0
    normal = 0
    low = 0

    for t in tasks:
        if t["done"]:
            continue
        d = parse_due(t.get("due"))
        cat = category(d)
        if cat == "просрочена":
            overdue += 1
        elif cat == "сегодня":
            today += 1
        if d is None:
            no_due += 1
        p = t["priority"]
        if p == "high":
            high += 1
        elif p == "normal":
            normal += 1
        elif p == "low":
            low += 1

    return [
        ("Всего задач", total, "Все задачи из списка"),
        ("Выполнено", done, "Задачи с done=true"),
        ("Не выполнено", todo, "Задачи с done=false"),
        ("Выполнено, %", pct, f"Выполнено / Всего * 100, округление до 1 знака"),
        ("Просрочено и не выполнено", overdue, "Невыполненная задача с дедлайном раньше сегодняшней даты"),
        ("Дедлайн сегодня и не выполнено", today, "Невыполненная задача с дедлайном = сегодня"),
        ("Без срока и не выполнено", no_due, "Невыполненная задача без дедлайна"),
        ("Незакрытые с высоким приоритетом", high, "done=false и priority=high"),
        ("Незакрытые с обычным приоритетом", normal, "done=false и priority=normal"),
        ("Незакрытые с низким приоритетом", low, "done=false и priority=low"),
    ]


def build_detail_rows(tasks):
    todo_items = [t for t in tasks if not t["done"]]
    done_items = [t for t in tasks if t["done"]]

    def sort_key(t):
        d = parse_due(t.get("due"))
        # Items without a deadline go last, then by due date, then by ID
        if d is None:
            return (1, date.max, t["id"])
        return (0, d, t["id"])

    todo_items.sort(key=sort_key)
    done_items.sort(key=sort_key)

    rows = []
    for t in todo_items + done_items:
        d = parse_due(t.get("due"))
        cat = category(d)
        rows.append([
            t["id"],
            t["title"],
            status_text(t["done"]),
            t["priority"],
            d.isoformat() if d else "",
            days_until(d) if d is not None else "",
            cat,
            "",
        ])
    return rows


def write_summary(ws, summary):
    today_str = date.today().isoformat()
    ws["A1"] = "Метрика"
    ws["B1"] = "Значение"
    ws["C1"] = "Правило расчёта"
    style_header(ws)

    # Title row
    title_cell = ws.cell(row=2, column=1, value=f"Отчёт сформирован: {today_str}")
    title_cell.font = Font(bold=True, color=ACCENT, size=12)
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    start = 3
    for i, (label, value, rule) in enumerate(summary):
        r = start + i
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=value)
        ws.cell(row=r, column=3, value=rule)

    # Style metric rows
    for r in range(start, start + len(summary)):
        is_odd = i % 2 == 0 if (i := r - start) is not None else True
        fill = PatternFill("solid", fgColor=BG_ALT) if (r - start) % 2 == 0 else PatternFill("solid", fgColor=BG)
        for cell in ws[r]:
            cell.fill = fill
            cell.font = Font(color=TEXT, size=11)
            cell.alignment = Alignment(vertical="center")
            cell.border = BORDER
        ws.cell(row=r, column=2).font = Font(color=ACCENT, bold=True, size=11)

    ws.column_dimensions["A"].width = 42
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 62


def write_tasks(ws, rows):
    headers = ["ID", "Название", "Статус", "Приоритет", "Дедлайн", "Дней до дедлайна", "Категория срока", "Примечание"]
    for col, h in enumerate(headers, 1):
        ws.cell(row=1, column=col, value=h)
    style_header(ws)

    for i, row in enumerate(rows):
        r = i + 2
        cat = row[6]
        cat_color = category_color(cat)
        for c, val in enumerate(row, 1):
            cell = ws.cell(row=r, column=c, value=val)
            if c == 3:  # Status
                cell.font = Font(color=DONE_GREEN if val == "выполнено" else TODAY_YELLOW, bold=True)
            elif c == 4:  # Priority
                p = row[2]
                if p == "high":
                    cell.font = Font(color=HIGH_ORANGE, bold=True)
                elif p == "normal":
                    cell.font = Font(color=ACCENT, bold=True)
                else:
                    cell.font = Font(color=MUTED)
            elif c == 6:  # Days until
                cell.font = Font(color=MUTED)
            elif c == 7:  # Category
                cell.font = Font(color=cat_color, bold=True)
            else:
                cell.font = Font(color=TEXT)

    ws.column_dimensions["A"].width = 6
    ws.column_dimensions["B"].width = 40
    ws.column_dimensions["C"].width = 16
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 12
    ws.column_dimensions["F"].width = 16
    ws.column_dimensions["G"].width = 14
    ws.column_dimensions["H"].width = 10


def main():
    # The MCP todo server is accessed via the tool, not the DB file.
    # Data was collected via mcp__todo__list_tasks and is embedded below
    # because the MCP tool result cannot be read directly by this script.
    # See todo_data.json for the raw MCP response.
    data_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "todo_data.json")
    with open(data_path, "r", encoding="utf-8") as f:
        tasks = json.load(f)

    wb = openpyxl.Workbook()
    wb.active.title = "Сводка"
    summary_ws = wb.active
    summary = build_summary(tasks)
    write_summary(summary_ws, summary)

    detail_ws = wb.create_sheet("Задачи")
    detail_rows = build_detail_rows(tasks)
    write_tasks(detail_ws, detail_rows)

    wb.save(OUTPUT_PATH)
    print(f"Saved: {OUTPUT_PATH}")
    print(f"Tasks: {len(tasks)}")
    for label, value, _ in summary:
        print(f"  {label}: {value}")


if __name__ == "__main__":
    main()