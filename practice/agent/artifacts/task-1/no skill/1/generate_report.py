#!/usr/bin/env python3
"""Генерация Excel-отчёта по задачам TODO-приложения."""

import datetime as dt
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# Данные получены через MCP-сервер todo (list_tasks), 2026-11-10
TASKS = [
    (1, "Записаться на документацию", True, "low", "2026-11-01"),
    (2, "Распечатать презентацию", False, "high", None),
    (3, "Согласовать документацию", False, "high", None),
    (4, "Заказать документацию", False, "normal", "2026-12-04"),
    (5, "Позвонить по поводу обновления системы", True, "normal", "2026-10-31"),
    (6, "Обновить презентацию", True, "normal", "2026-11-21"),
    (7, "Записаться на договор", False, "high", "2026-12-08"),
    (8, "Починить презентацию", True, "high", None),
    (9, "Протестировать резервную копию", True, "high", "2026-11-23"),
    (10, "Проверить документацию", False, "low", "2026-11-18"),
    (11, "Позвонить по поводу подарок коллеге", True, "normal", "2026-11-23"),
    (12, "Доработать визит к врачу", False, "high", "2026-11-04"),
    (13, "Починить билеты на поезд", False, "high", None),
    (14, "Отправить документацию", False, "low", "2026-11-20"),
    (15, "Согласовать посылку", True, "high", "2026-11-04"),
    (16, "Прочитать молоко и хлеб", True, "normal", "2026-11-25"),
    (17, "Записаться на посылку", False, "normal", "2026-12-04"),
    (18, "Купить доклад для лекции", True, "normal", "2026-11-14"),
    (19, "Обновить молоко и хлеб", True, "high", None),
    (20, "Проверить заявку на отпуск", False, "normal", "2026-12-08"),
    (21, "Распечатать отчёт за квартал", False, "high", "2026-11-16"),
    (22, "Записаться на план на неделю", True, "low", "2026-12-04"),
    (23, "Подготовить билеты на поезд", False, "high", None),
    (24, "Разобрать доклад для лекции", True, "high", "2026-11-06"),
    (25, "Распечатать доклад для лекции", False, "low", "2026-11-10"),
    (26, "Забронировать ревью кода", True, "low", "2026-11-07"),
    (27, "Проверить встречу с командой", False, "low", "2026-11-05"),
    (28, "Купить встречу с командой", True, "high", "2026-11-08"),
    (29, "Заказать резервную копию", False, "high", "2026-12-08"),
    (30, "Починить подарок коллеге", False, "normal", "2026-11-19"),
    (31, "Оплатить презентацию", True, "normal", "2026-11-15"),
    (32, "Оплатить отчёт за квартал", False, "low", "2026-12-07"),
    (33, "Проверить договор", True, "normal", None),
    (34, "Согласовать анкету", False, "low", "2026-12-01"),
    (35, "Спланировать документацию", False, "normal", "2026-12-06"),
    (36, "Обновить обновления системы", False, "normal", None),
    (37, "Обновить договор", False, "normal", "2026-11-03"),
    (38, "Оплатить резервную копию", False, "low", "2026-11-06"),
    (39, "Позвонить по поводу билеты на поезд", True, "normal", None),
    (40, "Распечатать счёт за интернет", False, "low", "2026-11-28"),
    (41, "Проверить документацию", False, "low", "2026-10-31"),
    (42, "Отправить визит к врачу", True, "normal", "2026-11-25"),
    (43, "Разобрать обновления системы", False, "high", "2026-11-29"),
    (44, "Спланировать молоко и хлеб", True, "low", "2026-12-05"),
    (45, "Распечатать договор", True, "high", "2026-12-07"),
    (46, "Спланировать доклад для лекции", True, "high", None),
    (47, "Позвонить по поводу презентацию", True, "high", "2026-11-05"),
    (48, "Обновить посылку", False, "low", "2026-11-25"),
    (49, "Проверить обновления системы", False, "high", None),
    (50, "Согласовать план на неделю", True, "low", "2026-11-16"),
    (51, "Разобрать список покупок", False, "normal", None),
    (52, "Прочитать тесты для приложения", False, "high", "2026-10-31"),
    (53, "Согласовать доклад для лекции", True, "normal", None),
    (54, "Оплатить подарок коллеге", True, "normal", "2026-11-04"),
    (55, "Доработать доклад для лекции", False, "low", "2026-11-19"),
    (56, "Купить билеты на поезд", False, "low", "2026-11-06"),
    (57, "Записаться на ревью кода", True, "high", "2026-11-09"),
    (58, "Записаться на доклад для лекции", True, "normal", "2026-12-10"),
    (59, "Проверить обновления системы", True, "normal", "2026-11-03"),
    (60, "Записаться на визит к врачу", False, "high", "2026-11-08"),
    (61, "Подготовить счёт за интернет", False, "low", "2026-12-05"),
    (62, "Распечатать договор", False, "normal", "2026-11-09"),
    (63, "Купить договор", False, "normal", None),
    (64, "Забронировать резервную копию", True, "high", "2026-11-02"),
    (65, "Распечатать обновления системы", False, "high", None),
    (66, "Позвонить по поводу визит к врачу", True, "normal", "2026-11-15"),
    (67, "Заказать документацию", True, "normal", "2026-11-21"),
    (68, "Отправить встречу с командой", True, "low", "2026-11-22"),
    (69, "Разобрать документацию", True, "low", "2026-11-24"),
    (70, "Написать билеты на поезд", False, "low", None),
    (71, "Починить анкету", False, "normal", "2026-12-02"),
    (72, "Записаться на визит к врачу", False, "high", None),
    (73, "Доработать обновления системы", False, "normal", "2026-11-06"),
    (74, "Закрыть счёт за интернет", False, "normal", "2026-12-08"),
    (75, "Отправить обновления системы", False, "low", None),
    (76, "Забронировать обновления системы", False, "low", "2026-11-12"),
    (77, "Написать счёт за интернет", False, "high", "2026-12-09"),
    (78, "Написать заявку на отпуск", True, "high", "2026-11-26"),
    (79, "Заказать заявку на отпуск", False, "high", None),
    (80, "Протестировать ревью кода", False, "normal", None),
    (81, "Разобрать резервную копию", False, "normal", None),
    (82, "Прочитать список покупок", False, "normal", "2026-11-20"),
    (83, "Проверить ревью кода", False, "high", "2026-11-10"),
    (84, "Оплатить ревью кода", False, "low", "2026-11-05"),
    (85, "Отправить документацию", False, "normal", None),
    (86, "Заказать посылку", False, "low", "2026-11-29"),
    (87, "Оплатить молоко и хлеб", False, "high", "2026-11-25"),
    (88, "Заказать документацию", False, "low", "2026-11-06"),
    (89, "Отправить анкету", False, "low", "2026-11-29"),
    (90, "Купить список покупок", False, "high", None),
    (91, "Прочитать тесты для приложения", False, "high", "2026-11-20"),
    (92, "Прочитать визит к врачу", False, "high", "2026-12-05"),
    (93, "Записаться на доклад для лекции", False, "normal", "2026-11-15"),
    (94, "Разобрать документацию", True, "normal", "2026-11-04"),
    (95, "Купить молоко и хлеб", False, "low", "2026-11-05"),
    (96, "Согласовать презентацию", False, "normal", None),
    (97, "Отправить резервную копию", False, "normal", "2026-11-26"),
    (98, "Протестировать заявку на отпуск", True, "normal", "2026-11-01"),
    (99, "Заказать анкету", False, "high", "2026-11-24"),
    (100, "Спланировать документацию", False, "normal", "2026-11-14"),
]

TODAY = dt.date(2026, 11, 10)
PRIORITIES_RU = {"high": "высокий", "normal": "обычный", "low": "низкий"}
STATUS_RU = {True: "выполнено", False: "не выполнено"}


def days_until(deadline_str):
    if deadline_str is None:
        return None
    y, m, d = map(int, deadline_str.split("-"))
    return (dt.date(y, m, d) - TODAY).days


def deadline_category(deadline_str, done):
    if deadline_str is None:
        return "не задан"
    y, m, d = map(int, deadline_str.split("-"))
    delta = (dt.date(y, m, d) - TODAY).days
    if delta < 0:
        return "просрочена"
    elif delta == 0:
        return "сегодня"
    else:
        return "впереди"


# Сортировка: невыполненные первыми, внутри группы по дедлайну (None — в конец), затем по ID
def sort_key(task):
    id_, title, done, priority, due = task
    due_days = days_until(due)
    if due_days is None:
        # Бесплатный байт, чтобы задачи без срока шли после всех с дедлайном
        return (0 if not done else 1, 10**9, id_)
    return (0 if not done else 1, due_days, id_)


sorted_tasks = sorted(TASKS, key=sort_key)

# Метрики
total = len(TASKS)
done_count = sum(1 for t in TASKS if t[2])
not_done_count = total - done_count
pct_done = round(done_count / total * 100, 1)

overdue = [t for t in TASKS if not t[2] and t[4] is not None and days_until(t[4]) < 0]
today_deadline = [t for t in TASKS if not t[2] and t[4] is not None and days_until(t[4]) == 0]
no_deadline = [t for t in TASKS if not t[2] and t[4] is None]

high_unclosed = [t for t in TASKS if not t[2] and t[3] == "high"]
normal_unclosed = [t for t in TASKS if not t[2] and t[3] == "normal"]
low_unclosed = [t for t in TASKS if not t[2] and t[3] == "low"]

# ── Создание книги ──
wb = Workbook()

# ── Лист «Сводка» ──
ws = wb.active
ws.title = "Сводка"

thin = Side(style="thin", color="999999")
border = Border(left=thin, right=thin, top=thin, bottom=thin)
header_fill = PatternFill("solid", fgColor="4472C4")
header_font = Font(color="FFFFFF", bold=True)
rule_fill = PatternFill("solid", fgColor="D9E1F2")

ws["A1"] = "Дата формирования отчёта"
ws["A1"].font = Font(bold=True)
ws["B1"] = TODAY.isoformat()
ws["B1"].number_format = "yyyy-mm-dd"

headers = ["Метрика", "Значение", "Правило расчёта"]
for col, h in enumerate(headers, 1):
    cell = ws.cell(row=3, column=col, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border

rows = [
    ("Всего задач", total, "Подсчёт всех записей в TODO-приложении"),
    ("Выполнено", done_count, "Задачи с done = true"),
    ("Не выполнено", not_done_count, "Всего задач − Выполнено"),
    ("Выполнено в %", pct_done, f"(Выполнено / Всего задач) × 100 = ({done_count} / {total}) × 100, округление до одного знака после запятой"),
    ("Просрочено и не выполнено", len(overdue),
     "Невыполненные задачи с дедлайном раньше сегодняшней даты " + TODAY.isoformat()),
    ("Дедлайн сегодня и не выполнено", len(today_deadline),
     "Невыполненные задачи с дедлайном, совпадающим с сегодняшней датой " + TODAY.isoformat()),
    ("Без срока и не выполнено", len(no_deadline),
     "Невыполненные задачи без дедлайна (дедлайн не задан)"),
    ("Незакрытые с высоким приоритетом", len(high_unclosed),
     "Невыполненные задачи с приоритетом «high»"),
    ("Незакрытые с обычным приоритетом", len(normal_unclosed),
     "Невыполненные задачи с приоритетом «normal»"),
    ("Незакрытые с низким приоритетом", len(low_unclosed),
     "Невыполненные задачи с приоритетом «low»"),
]

ws["B7"].number_format = "0.0"

for i, (metric, value, rule) in enumerate(rows):
    row = 4 + i
    c1 = ws.cell(row=row, column=1, value=metric)
    c2 = ws.cell(row=row, column=2, value=value)
    c3 = ws.cell(row=row, column=3, value=rule)
    for c in (c1, c2, c3):
        c.border = border
    c2.alignment = Alignment(horizontal="center")
    if i % 2 == 1:
        for c in (c1, c2, c3):
            c.fill = rule_fill

ws.column_dimensions["A"].width = 32
ws.column_dimensions["B"].width = 14
ws.column_dimensions["C"].width = 90

# ── Лист «Задачи» ──
ws2 = wb.create_sheet("Задачи")

task_headers = ["ID", "Название", "Статус", "Приоритет", "Дедлайн",
                "Дней до дедлайна", "Категория срока", "Примечание"]
for col, h in enumerate(task_headers, 1):
    cell = ws2.cell(row=1, column=col, value=h)
    cell.font = header_font
    cell.fill = header_fill
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = border

for idx, task in enumerate(sorted_tasks):
    id_, title, done, priority, due = task
    row = 2 + idx
    cells = [
        (1, id_),
        (2, title),
        (3, STATUS_RU[done]),
        (4, PRIORITIES_RU[priority]),
        (5, due if due else "—"),
        (6, days_until(due) if due else "—"),
        (7, deadline_category(due, done)),
        (8, "—"),
    ]
    for col, val in cells:
        cell = ws2.cell(row=row, column=col, value=val)
        cell.border = border

ws2.column_dimensions["A"].width = 6
ws2.column_dimensions["B"].width = 42
ws2.column_dimensions["C"].width = 14
ws2.column_dimensions["D"].width = 14
ws2.column_dimensions["E"].width = 12
ws2.column_dimensions["F"].width = 16
ws2.column_dimensions["G"].width = 16
ws2.column_dimensions["H"].width = 12

output_path = "/home/arcle/mcs/aidd04/manual/practice/agent/todo_report.xlsx"
wb.save(output_path)

print(f"Файл сохранён: {output_path}")
print(f"Всего задач: {total}")
print(f"Выполнено: {done_count}")
print(f"Не выполнено: {not_done_count}")
print(f"Выполнено: {pct_done}%")
print(f"Просрочено и не выполнено: {len(overdue)}")
print(f"Дедлайн сегодня и не выполнено: {len(today_deadline)}")
print(f"Без срока и не выполнено: {len(no_deadline)}")
print(f"Незакрытые с высоким приоритетом: {len(high_unclosed)}")
print(f"Незакрытые с обычным приоритетом: {len(normal_unclosed)}")
print(f"Незакрытые с низким приоритетом: {len(low_unclosed)}")