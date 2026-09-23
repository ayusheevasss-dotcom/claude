#!/usr/bin/env python3
"""Сборка демо-данных участка «Вилюй» для дашборда «Учёт объёмов работ».

Читает пять отчётов из папки «исходные», по каждому проекту строит
месячный план (как если бы его загрузил администратор) и дневной факт
(как если бы его вносил ответственный проекта), затем:
  1) пишет план-файл Excel «Планы_Вилюй_2026-09.xlsx» — один файл, лист на проект;
  2) вшивает планы, факт, комментарии, причины и журнал в HTML дашборда
     (блок между маркерами <VIL-DATA> … </VIL-DATA>).

Запуск:  python3 vilyuy/сборка/build_vilyuy.py
"""
import calendar
import datetime as dt
import json
import math
import re
import statistics
from pathlib import Path

import openpyxl
import xlrd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / 'исходные'
HTML = ROOT / 'Учёт_объёмов_работ_Вилюй.html'
MONTHS = ['январь', 'февраль', 'март', 'апрель', 'май', 'июнь', 'июль', 'август', 'сентябрь', 'октябрь', 'ноябрь', 'декабрь']
DOW = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс']

# ---------------------------------------------------------------- справочники
# новые виды работ (к стандартным: Строительство, Ремонт, Содержание, Асфальтирование, Транспортные услуги, Дробление)
EXTRA_CATS = [
    {'id': 'haul_t', 'name': 'Перевозка материалов (т)', 'unit': 'т'},
    {'id': 'haul_m3', 'name': 'Перевозка материалов (м³)', 'unit': 'м³'},
    {'id': 'asph_t', 'name': 'Асфальтобетон (т)', 'unit': 'т'},
]
CAT_NAMES = {'build': 'Строительство', 'repair': 'Ремонт', 'maint': 'Содержание', 'asph': 'Асфальтирование',
             'trans': 'Транспортные услуги', 'crush': 'Дробление', '-': 'Не учитывать',
             **{c['id']: c['name'] for c in EXTRA_CATS}}

# проекты участка «Вилюй»: у каждого свой ответственный и пароль (ФИО — условные, заменить на реальные)
PROJECTS = [
    {'id': 'vil_drob', 'sheet': 'Дробление', 'title': 'Дробление щебня (УДСК Лапчан, УДСК Мирный)', 'short': 'Вилюй · Дробление',
     'person': 'Кузнецов Д. А.', 'pass': '2201', 'source': 'Дробление_22.09.2026.xlsx'},
    {'id': 'vil_perev', 'sheet': 'Перевозка', 'title': 'Перевозка материалов (месячно-суточный график)', 'short': 'Вилюй · Перевозка',
     'person': 'Семёнова О. В.', 'pass': '2202', 'source': 'Месячно-суточный_график_перевозки_сентябрь.xls'},
    {'id': 'vil_smen', 'sheet': 'Сменные работы', 'title': 'Сменные работы км 1180–1190 (накопительная сводка)', 'short': 'Вилюй · Смены',
     'person': 'Николаев П. С.', 'pass': '2203', 'source': 'Накопительная_сводка_км1180-1190_19.08.2026.xlsx'},
    {'id': 'vil_ats', 'sheet': 'Автотранспорт', 'title': 'Автотранспорт (рейсы по машинам)', 'short': 'Вилюй · Рейсы',
     'person': 'Егорова М. И.', 'pass': '2204', 'source': 'Суточный_отчёт_автотранспорт_сентябрь.xlsx'},
    {'id': 'vil_gk', 'sheet': 'Капремонт ГК', 'title': 'Капремонт км 1180–1190 (ход работ по ГК)', 'short': 'Вилюй · Капремонт',
     'person': 'Попов А. Г.', 'pass': '2205', 'source': 'Исполнительная_схема_км1180-1190_22.09.2026.xlsx'},
]
PRJ = {p['id']: p for p in PROJECTS}


def dim(ym):
    y, m = map(int, ym.split('-'))
    return calendar.monthrange(y, m)[1]


def nice(x):
    """Норма на день — «круглое» число, как её задал бы плановик."""
    if x <= 0:
        return 0
    step = 1 if x < 20 else 5 if x < 200 else 10 if x < 1000 else 50
    return max(step, round(x / step) * step)


def num(v):
    if v is None or v == '':
        return None
    if isinstance(v, (int, float)):
        return float(v) if math.isfinite(v) else None
    s = str(v).replace(' ', '').replace('\xa0', '').replace(',', '.')
    try:
        return float(s)
    except ValueError:
        return None


def rnd(v, k=2):
    v = round(v, k)
    return int(v) if v == int(v) else v


def derive_plan(fact, n, as_of, start_early=5, tail=3):
    """План по работе, у которой в отчёте нет дневного плана.
    Окно работ — от первого дня с фактом (с 1-го, если работа шла с начала месяца) до последнего;
    если работа шла в последние дни отчёта — до конца месяца.
    Норма на день: для работ, которые шли почти ежедневно, — средняя сменная выработка;
    для эпизодических — средний темп за окно + 10 % (напряжённый план). Норма округляется."""
    nz = {d: v for d, v in fact.items() if v > 0}
    if not nz:
        return {}
    first, last = min(nz), max(nz)
    start = 1 if first <= start_early else first
    end = n if last >= as_of - tail else last
    span = min(end, as_of) - start + 1
    if len(nz) / span >= 0.6:
        norm = nice(statistics.mean(nz.values()))
    else:
        norm = nice(sum(nz.values()) / span * 1.1)
    return {d: norm for d in range(start, end + 1)}


class Work:
    def __init__(self, name, unit, cls, group=''):
        self.name, self.unit, self.cls, self.group = name, unit, cls, group
        self.plan, self.fact = {}, {}


class Month:
    """План и факт проекта за месяц."""
    def __init__(self, ym, as_of):
        self.ym, self.as_of, self.works, self.notes = ym, as_of, [], {}

    def add(self, w):
        self.works.append(w)
        return w

    def finalize(self, zero_fill=True):
        """Ответственный вносит 0 по плановой работе, если в этот день она не выполнялась."""
        for w in self.works:
            w.plan = {d: rnd(v) for d, v in w.plan.items() if v}
            w.fact = {d: rnd(v) for d, v in w.fact.items() if v is not None and d <= self.as_of}
            if zero_fill:
                for d in w.plan:
                    if d <= self.as_of and d not in w.fact:
                        w.fact[d] = 0
        self.works = [w for w in self.works if w.plan or any(w.fact.values())]
        return self


# ---------------------------------------------------------------- 1. Дробление
def read_drob():
    ws = openpyxl.load_workbook(SRC / PRJ['vil_drob']['source'], data_only=True)['Дробление Август']
    # лист назван «Август», но в шапке — «Сентябрь» и «по состоянию на 21.09.2026»: это сентябрь
    ym, as_of = '2026-09', 22
    n = dim(ym)
    M = Month(ym, as_of)
    site = crusher = ''
    for r in range(7, ws.max_row + 1):
        b, c = ws.cell(r, 2).value, ws.cell(r, 3).value
        b = str(b).strip() if b else ''
        c = str(c).strip() if c else ''
        if b.startswith('УДСК'):
            site = b.title().replace('Удск', 'УДСК')
            continue
        if b.startswith('Дробление') or b.startswith('Отгружено'):
            crusher = b
            continue
        days = {d: num(ws.cell(r, 5 + d).value) for d in range(1, n + 1)}
        days = {d: v for d, v in days.items() if v}
        if not days or c in ('Остаток', 'Всего', 'план') or b == '%':
            continue
        if c.startswith('фр') or c.startswith('ЩПС'):
            frac = c
        elif c.startswith('факт'):
            frac = ''
        else:
            continue
        unit_name = crusher.replace('Дробление фр.щебня', 'Дробление').replace('Дробление ', '')
        if 'Дробмаш' in crusher:
            name = f'Дробмаш: щебень {frac}'
        elif 'Локотрак' in crusher:
            name = f'Локотрак: {frac.replace("фр.", "щебень фр. ")}'
        elif site.endswith('Мирный'):
            name = f'УДСК Мирный: щебень {frac}'
        else:
            name = f'{site}: {unit_name}'.replace('ЩПС  ', 'ЩПС ')
        name = re.sub(r'\s+', ' ', name.replace('фр.', 'фр. ')).replace('фр. фр.', 'фр.').strip()
        w = M.add(Work(name, 'м³', 'crush', site))
        w.fact = days
    for w in M.works:
        w.plan = derive_plan(w.fact, n, as_of)
    return {ym: M.finalize()}


# ---------------------------------------------------------------- 2. Перевозка (месячно-суточный график)
def read_perev():
    sh = xlrd.open_workbook(SRC / PRJ['vil_perev']['source']).sheet_by_index(0)
    ym, as_of = '2026-09', 21  # факт в отчёте — по 21.09
    n = dim(ym)
    M = Month(ym, as_of)
    c1 = 14  # столбец O = 1-е число
    names = {
        'Перевозка ЩПС 0/63  / с Лапчана': 'Перевозка ЩПС 0/63 с к. Лапчан',
        'Перевозка асфальта с АБЗ': 'Перевозка асфальтобетонной смеси с АБЗ',
        'Щебень фр.  / с Лапчана': 'Перевозка щебня с к. Лапчан',
        'Щебень фр. / со Светлого': 'Перевозка щебня с п. Светлый',
    }
    for r in range(sh.nrows):
        raw = str(sh.cell_value(r, 1)).replace('\n', ' / ').strip()
        key = re.sub(r'\s+', ' ', raw)
        match = next((v for k, v in names.items() if re.sub(r'\s+', ' ', k) == key), None)
        if match:
            w = M.add(Work(match, 'т', 'haul_t'))
            w.plan = {d: num(sh.cell_value(r, c1 + d - 1)) for d in range(1, n + 1)}
            w.fact = {d: num(sh.cell_value(r + 1, c1 + d - 1)) for d in range(1, n + 1)}
            w.plan = {d: v for d, v in w.plan.items() if v}
            w.fact = {d: v for d, v in w.fact.items() if v is not None}
        if key.startswith('Наличие автосамосвалов'):
            w = M.add(Work('Автосамосвалы на линии (собств. и привлечённые)', 'ед.', '-'))
            w.plan = {d: num(sh.cell_value(r, c1 + d - 1)) for d in range(1, n + 1) if num(sh.cell_value(r, c1 + d - 1))}
            w.plan = {d: v for d, v in w.plan.items() if v <= 200}  # в последнем столбце — «хвост» формулы (1120)
            w.fact = {d: num(sh.cell_value(r + 1, c1 + d - 1)) for d in range(1, n + 1) if num(sh.cell_value(r + 1, c1 + d - 1))}
        if str(sh.cell_value(r, 13)).strip().lower() == 'погода':
            for d in range(1, n + 1):
                v = str(sh.cell_value(r, c1 + d - 1)).strip()
                if v:
                    M.notes[d] = f'Погода: {v}. Перевозка по графику не выполнена в полном объёме.'
    return {ym: M.finalize()}


# ---------------------------------------------------------------- 3. Сменные работы км 1180–1190
def read_smen():
    ws = openpyxl.load_workbook(SRC / PRJ['vil_smen']['source'], data_only=True)['Август']
    ym, as_of = '2026-08', 19  # «по состоянию на 19.08.2026»
    n = dim(ym)
    M = Month(ym, as_of)
    names = [
        ('по вывозу непригодного грунта в отвал', 'Вывоз непригодного грунта в отвал', 'м³', 'haul_m3'),
        ('по устройству ЩПС 0/63', 'Устройство слоя основания из ЩПС 0/63', 'м³', 'build'),
        ('покрытия А32НН', 'Устройство нижнего слоя покрытия А32НН', 'т', 'asph_t'),
        ('покрытия А16ВН', 'Устройство нижнего слоя покрытия А16ВН', 'т', 'asph_t'),
        ('с "Лапчан" на накопитель км1182', 'Вывоз ЩПС 0/63 с к. Лапчан на накопитель км 1182', 'м³', 'haul_m3'),
        ('с карьера "км 179"', 'Вывоз с карьера «км 179» на дорогу ООО «Айхал-Центр»', 'м³', 'haul_m3'),
    ]
    cur = None
    for r in range(7, ws.max_row + 1):
        b = str(ws.cell(r, 2).value or '')
        c = str(ws.cell(r, 3).value or '')
        hit = next((x for x in names if x[0] in b), None)
        if b.startswith('Накопительная сводка'):
            cur = hit
            continue
        if cur and c.startswith('факт') and 'накоп' in c:
            w = M.add(Work(cur[1], cur[2], cur[3]))
            w.fact = {d: num(ws.cell(r, 5 + d).value) for d in range(1, n + 1)}
            w.fact = {d: v for d, v in w.fact.items() if v}
            cur = None
    for w in M.works:
        w.plan = derive_plan(w.fact, n, as_of)
    return {ym: M.finalize()}


# ---------------------------------------------------------------- 4. Автотранспорт (рейсы)
def read_ats():
    wb = openpyxl.load_workbook(SRC / PRJ['vil_ats']['source'], data_only=True)
    ym, as_of = '2026-09', 22
    n = dim(ym)
    M = Month(ym, as_of)
    works = {}
    last_plan = {}
    for ws in wb.worksheets:
        if ws.sheet_state != 'visible':  # скрытые листы 11–13.09 — копии листа 10.09 (незаполненные заготовки)
            continue
        d = int(ws.title[:2])
        route = None
        notes = []
        day_plan = {}
        for r in range(5, ws.max_row + 1):
            a, b = ws.cell(r, 1).value, ws.cell(r, 2).value
            if isinstance(a, str) and 'смена' in a:
                continue
            if isinstance(a, str) and not a.strip().isdigit():
                route = re.sub(r'\s+', ' ', a.strip()).replace('км1182', 'км 1182')
                continue
            m = ws.cell(r, 13).value
            if m and b != 'Всего:' and 'разные расстояния' not in str(m).lower():
                notes.append(f'{ws.cell(r, 3).value} {ws.cell(r, 4).value} — {str(m).strip()}')
            if b == 'Всего:' and route:
                asph = 'асфальт' in route
                unit = 'т' if asph else 'м³'
                w = works.get(route) or M.add(Work(route, unit, 'haul_t' if asph else 'haul_m3'))
                works[route] = w
                p, f = num(ws.cell(r, 8).value) or 0, num(ws.cell(r, 11).value) or 0
                w.plan[d] = w.plan.get(d, 0) + p
                w.fact[d] = w.fact.get(d, 0) + f
                day_plan[route] = day_plan.get(route, 0) + p
        if notes:
            M.notes[d] = 'Сход с линии / замечания: ' + '; '.join(notes) + '.'
        last_plan = day_plan
    # план на оставшиеся дни месяца — по последнему суточному заданию (22.09)
    for d in range(as_of + 1, n + 1):
        for route, p in last_plan.items():
            if p:
                works[route].plan[d] = p
    for w in M.works:
        w.fact = {d: v for d, v in w.fact.items() if v or w.plan.get(d)}
    return {ym: M.finalize()}


# ---------------------------------------------------------------- 5. Капремонт км 1180–1190 по ГК
GK_NAMES = {
    'Устройство земляного полотна': ('Устройство земляного полотна', 'м³', 'build'),
    'Устройство слоя основания (щпс 0/63). Слой 1 (толщ.20 см)': ('Устройство основания из ЩПС 0/63, слой 1 (20 см)', 'м³', 'build'),
    'Устройство слоя основания (щпс 0/63). Слой 2 (толщ.13 см)': ('Устройство основания из ЩПС 0/63, слой 2 (13 см)', 'м³', 'build'),
    'Устройство нижнего слоя покрытия': ('Устройство нижнего слоя покрытия', 'м²', 'asph'),
    'Устройство верхнего слоя покрытия': ('Устройство верхнего слоя покрытия', 'м²', 'asph'),
    'Устройство нижнего слоя покрытия(примыкания)': ('Устройство нижнего слоя покрытия (примыкания)', 'м²', 'asph'),
    'Устройство верхнего слоя покрытия(примыкания)': ('Устройство верхнего слоя покрытия (примыкания)', 'м²', 'asph'),
    'Устройство присыпных обочин (20 см)': ('Устройство присыпных обочин (20 см)', 'м³', 'build'),
    'Устройство присыпных обочин (13 см)#1': ('Устройство присыпных обочин (13 см), позиция 1', 'м³', 'build'),
    'Устройство присыпных обочин (13 см)#2': ('Устройство присыпных обочин (13 см), позиция 2', 'м³', 'build'),
    'Укрепление откосов земляного полотна посевом трав': ('Укрепление откосов посевом трав', 'м²', '-'),
    'Планировка откосов земляного полотна': ('Планировка откосов земляного полотна', 'м²', '-'),
    'Установка МБО': ('Установка металлического барьерного ограждения', 'м', '-'),
}
# значения «за сутки», которые не подтверждаются нарастающим итогом
GK_SKIP = {('Планировка откосов земляного полотна', '2026-09-17'): 'за 17.09 указано 12 303 м², но нарастающий итог не изменился; объём учтён 21.09'}


def read_gk():
    wb = openpyxl.load_workbook(SRC / PRJ['vil_gk']['source'], data_only=True)
    series = {}  # дата -> {ключ: (нарастающий итог, за сутки, по ГК)}
    for ws in wb.worksheets:
        t = ws.title.strip()
        if not re.match(r'\d\d\.\d\d\.\d{4}$', t):
            continue
        date = dt.date(int(t[6:]), int(t[3:5]), int(t[:2]))
        hdr = next(r for r in range(25, 40) if ws.cell(r, 14).value == 'Вид работ')
        seen, rows = {}, {}
        for r in range(hdr + 1, ws.max_row + 1):
            nm = ws.cell(r, 14).value
            if not nm:
                continue
            nm = re.sub(r'\s+', ' ', str(nm)).strip()
            if nm.startswith('Устройство присыпных обочин (13'):
                seen[nm] = seen.get(nm, 0) + 1
                nm = f'{nm}#{seen[nm]}'
            rows[nm] = (num(ws.cell(r, 43).value), num(ws.cell(r, 59).value), num(ws.cell(r, 36).value))
        series[date] = rows
    dates = sorted(series)
    out = {}
    for ym, as_of in (('2026-08', 31), ('2026-09', 22)):
        y, m = map(int, ym.split('-'))
        n = dim(ym)
        M = Month(ym, as_of)
        for key, (name, unit, cls) in GK_NAMES.items():
            w = Work(name, unit, cls)
            for d in range(1, as_of + 1):
                day = dt.date(y, m, d)
                if day not in series or key not in series[day]:
                    continue
                cum, daily, _ = series[day][key]
                prev = [x for x in dates if x < day]
                pcum, pdaily = (series[prev[-1]].get(key, (None, None, None))[:2] if prev else (None, None))
                stale = daily and pdaily == daily and cum == pcum  # значение «за сутки» перенесено со вчерашнего листа
                v = None
                if daily and daily > 0 and not stale:
                    v = daily
                elif cum is not None and pcum is not None and 0 < cum - pcum and not (daily == 0 and pdaily is None):
                    v = cum - pcum  # «за сутки» не заполнено, но нарастающий итог вырос
                if v and (name, day.isoformat()) not in GK_SKIP:
                    w.fact[d] = v
            w.plan = derive_plan(w.fact, n, as_of)
            M.works.append(w)
        out[ym] = M.finalize()
    return out


# ---------------------------------------------------------------- сборка
def build():
    data = {
        'vil_drob': read_drob(),
        'vil_perev': read_perev(),
        'vil_smen': read_smen(),
        'vil_ats': read_ats(),
        'vil_gk': read_gk(),
    }
    # у «Сменных работ» последний отчёт — на 19.08: план на сентябрь администратор составляет по выработке августа
    aug = data['vil_smen']['2026-08']
    sep = Month('2026-09', 0)
    for w in aug.works:
        nw = sep.add(Work(w.name, w.unit, w.cls, w.group))
        nw.plan = {d: max(w.plan.values()) for d in range(1, 31)} if w.plan else {}
    data['vil_smen']['2026-09'] = sep.finalize()
    return data


# ---------------------------------------------------------------- причины, журнал, данные для HTML
def ddmm(ym, d):
    return f'{d:02d}.{ym[5:]}'


def day_list(days):
    """[3,4,5,9] -> «3–5, 9»"""
    out, days = [], sorted(days)
    i = 0
    while i < len(days):
        j = i
        while j + 1 < len(days) and days[j + 1] == days[j] + 1:
            j += 1
        out.append(f'{days[i]}' if i == j else f'{days[i]}–{days[j]}')
        i = j + 1
    return ', '.join(out)


def reasons_for(sid, M):
    """Причины отклонений — из данных отчёта: работы ниже 90 % плана на дату и дни без выработки."""
    if not M.as_of:
        return {}
    by_cat = {}
    for w in M.works:
        if w.cls == '-':
            continue
        pt = sum(v for d, v in w.plan.items() if d <= M.as_of)
        ft = sum(w.fact.values())
        if not pt or ft >= pt * 0.9:
            continue
        zero = [d for d, v in w.plan.items() if d <= M.as_of and not w.fact.get(d)]
        t = f'{w.name} — {ft / pt * 100:.0f} % плана на {ddmm(M.ym, M.as_of)}'
        if zero:
            t += f', без выработки {day_list(zero)}.{M.ym[5:]}'
        by_cat.setdefault(w.cls, []).append(t)
    res = {c: '; '.join(v[:4]) + ('; …' if len(v) > 4 else '') for c, v in by_cat.items()}
    general = []
    rain = [d for d, t in M.notes.items() if 'дождь' in t.lower()]
    if rain:
        general.append(f'дожди {day_list(rain)}.{M.ym[5:]}')
    trucks = next((w for w in M.works if w.unit == 'ед.'), None)
    if trucks:
        f = [v for v in trucks.fact.values() if v]
        p = [v for d, v in trucks.plan.items() if d <= M.as_of]
        if f and p:
            general.append(f'нехватка автосамосвалов: на линии в среднем {statistics.mean(f):.0f} ед. при потребности {max(p):.0f}')
    breaks = [f'{ddmm(M.ym, d)} — {t.split(": ", 1)[1].rstrip(".")}' for d, t in M.notes.items() if t.startswith('Сход')]
    if breaks:
        general.append('сход техники с линии: ' + '; '.join(breaks))
    if general:
        g = '; '.join(general)
        res['_all'] = g[0].upper() + g[1:] + '.'
    return res


LOADED = {'2026-08': '2026-07-30T16:{m:02d}:00', '2026-09': '2026-08-28T15:{m:02d}:00'}
CLOSED = {('vil_gk', '2026-08'): {'at': '2026-09-02T10:15:00', 'by': 'Попов А. Г.'}}
EXTRA_NOTES = {('vil_gk', '2026-09', 21): 'Планировка откосов 10 647 м²: объём, указанный в суточном отчёте за 17.09 (12 303 м²), '
                                          'не подтверждён нарастающим итогом — учтён по факту 21.09.'}


def plan_file(ym):
    return f'Планы_Вилюй_{ym}.xlsx'


def to_vil(data):
    sections = []
    for p in PROJECTS:
        sections.append({'id': p['id'], 'name': f'Вилюй · {p["title"]}', 'short': p['short'], 'area': 'Вилюй', 'project': p['title'],
                         'person': p['person'], 'passPlain': p['pass'], 'aliases': [p['title'], p['sheet']]})
    months, log = [], []
    for i, p in enumerate(PROJECTS):
        sid = p['id']
        who = f'{p["person"]} ({p["short"]})'
        for ym, M in sorted(data[sid].items()):
            for (s2, y2, d), t in EXTRA_NOTES.items():
                if s2 == sid and y2 == ym:
                    M.notes[d] = t
            loaded = LOADED[ym].format(m=i * 3)
            source = f'{plan_file(ym)} / лист «{p["sheet"]}»'
            last_input = None
            for d in range(1, M.as_of + 1):
                k = sum(1 for w in M.works if d in w.fact)
                if not k:
                    continue
                y, m = map(int, ym.split('-'))
                at = (dt.datetime(y, m, d) + dt.timedelta(days=1, hours=8, minutes=20 + i * 7 + d % 5)).isoformat()
                last_input = at
                log.append({'at': at, 'who': who, 'role': 'sec', 'sid': sid, 'action': 'Внесение объёмов',
                            'details': f'{p["short"]}, {ddmm(ym, d)}.{ym[:4]}: изменено значений — {k}' + (', комментарий' if d in M.notes else '')})
            log.append({'at': loaded, 'who': 'Администратор', 'role': 'admin', 'sid': sid, 'action': 'Загрузка плана',
                        'details': f'{p["short"]}, {MONTHS[int(ym[5:]) - 1].capitalize()} {ym[:4]}: {len(M.works)} работ · источник: {source}'})
            reasons = reasons_for(sid, M)
            closure = CLOSED.get((sid, ym))
            if reasons and M.as_of:
                y, m = map(int, ym.split('-'))
                log.append({'at': (dt.datetime(y, m, M.as_of) + dt.timedelta(days=1, hours=9, minutes=5 + i)).isoformat(), 'who': who, 'role': 'sec',
                            'sid': sid, 'action': 'Причины отклонений', 'details': f'{p["short"]}, {MONTHS[int(ym[5:]) - 1].capitalize()} {ym[:4]}: изменено — {len(reasons)}'})
            if closure:
                log.append({'at': closure['at'], 'who': who, 'role': 'sec', 'sid': sid, 'action': 'Закрытие месяца',
                            'details': f'{p["short"]}: {MONTHS[int(ym[5:]) - 1].capitalize()} {ym[:4]} закрыт'})
            months.append({'sid': sid, 'ym': ym, 'loadedAt': loaded, 'source': source, 'lastInput': last_input, 'closure': closure,
                           'notes': {str(d): t for d, t in sorted(M.notes.items())}, 'reasons': reasons,
                           'works': [{'name': w.name, 'unit': w.unit, 'cls': w.cls,
                                      'plan': {str(d): v for d, v in sorted(w.plan.items())},
                                      'fact': {str(d): v for d, v in sorted(w.fact.items())}} for w in M.works]})
    log.sort(key=lambda e: e['at'], reverse=True)
    return {'sections': sections, 'categories': EXTRA_CATS, 'months': months, 'log': log}


def inject(vil):
    html = HTML.read_text(encoding='utf-8')
    js = 'const VIL = ' + json.dumps(vil, ensure_ascii=False, separators=(',', ':')) + ';'
    new, n = re.subn(r'(// <VIL-DATA>\n).*?(\n// </VIL-DATA>)', lambda m: m.group(1) + js + m.group(2), html, flags=re.S)
    assert n == 1, 'маркеры <VIL-DATA> не найдены'
    HTML.write_text(new, encoding='utf-8')


# ---------------------------------------------------------------- план-файл Excel
THIN = Side(style='thin', color='C9CED4')
BOX = Border(top=THIN, bottom=THIN, left=THIN, right=THIN)
F_HEAD = PatternFill('solid', fgColor='1F4E8C')
F_WE = PatternFill('solid', fgColor='F3E9D2')
F_GRP = PatternFill('solid', fgColor='DCE8F5')
F_TOT = PatternFill('solid', fgColor='EEF2F6')
HOW = {
    'vil_drob': 'В отчёте нет дневного плана: норма на день — средняя сменная выработка по каждой фракции, в дни работы установки.',
    'vil_perev': 'План по дням перенесён из месячно-суточного графика перевозки без изменений.',
    'vil_smen': 'Последняя сводка — на 19.08: норма на сентябрь — по выработке августа.',
    'vil_ats': 'План — суточные задания (рейсы × объём кузова) с листов 09–22.09; на 23–30.09 — задание 22.09.',
    'vil_gk': 'Норма на день — средний суточный прирост по исполнительной схеме (нарастающий итог и «за сутки»).',
}


def write_plan_xlsx(ym, data):
    y, m = map(int, ym.split('-'))
    n = dim(ym)
    month = f'{MONTHS[m - 1]} {y}'
    wb = openpyxl.Workbook()
    ins = wb.active
    ins.title = 'Инструкция'
    ins.sheet_view.showGridLines = False
    rows = [
        (f'Планы работ участка «Вилюй» — {month} г.', 'title'),
        ('Формат: один файл — один месяц, каждый проект участка — на своём листе.', ''),
        ('', ''),
        ('Как заполнить лист проекта', 'h'),
        ('1. Шапка листа: «План работ на <месяц> <год> г.», «Участок: Вилюй», «Проект: <название>», «Ответственный: <ФИО>».', ''),
        ('    По строке «Проект» программа сама сопоставляет лист с проектом, название листа может быть любым.', ''),
        ('2. Таблица: № | Наименование работ | Классификация | Ед. изм. | дни месяца 1 … 30 (31) | Итого за месяц.', ''),
        ('3. В ячейке дня — плановый объём на этот день. Пустая ячейка — работа в этот день не планируется.', ''),
        ('4. Классификация — выбор из списка (лист «Справочник»): определяет, в какой показатель отчёта попадёт работа.', ''),
        ('    Объём суммируется в показатель, только если единица работы совпадает с единицей вида работ.', ''),
        ('5. Строки-заголовки групп (например, «УДСК Лапчан») без чисел программа пропускает. «Итого» считается формулой.', ''),
        ('6. Новая работа — новая строка. Названия работ не меняйте в течение месяца: по ним факт связывается с планом.', ''),
        ('', ''),
        ('Загрузка в программу', 'h'),
        ('Администратор → «Планы» → «Загрузка из Excel» → месяц → перетащить файл → проверить распознавание → «Загрузить».', ''),
        ('Факт вносит ответственный проекта: «Ответственный» → участок «Вилюй» → свой проект → пароль → «Внесение объёмов».', ''),
        ('', ''),
        ('Проекты участка', 'h'),
    ]
    for i, (t, kind) in enumerate(rows, 1):
        c = ins.cell(i, 1, t)
        if kind == 'title':
            c.font = Font(bold=True, size=15, color='15191E')
        elif kind == 'h':
            c.font = Font(bold=True, size=12, color='1F4E8C')
        else:
            c.font = Font(size=10.5)
    r0 = len(rows) + 1
    hdr = ['Лист', 'Проект', 'Ответственный', 'Пароль (демо)', 'Работ', 'Исходный отчёт', 'Как построен план']
    for j, h in enumerate(hdr, 1):
        c = ins.cell(r0, j, h)
        c.font, c.fill, c.border = Font(bold=True, color='FFFFFF'), F_HEAD, BOX
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    for k, p in enumerate(PROJECTS, 1):
        M = data[p['id']].get(ym)
        if not M:
            continue
        vals = [p['sheet'], p['title'], p['person'], p['pass'], len(M.works), p['source'], HOW[p['id']]]
        for j, v in enumerate(vals, 1):
            c = ins.cell(r0 + k, j, v)
            c.border, c.alignment, c.font = BOX, Alignment(vertical='top', wrap_text=True), Font(size=10)
    for col, w in zip('ABCDEFG', (16, 40, 18, 11, 8, 40, 60)):
        ins.column_dimensions[col].width = w

    ref = wb.create_sheet('Справочник')
    ref['A1'], ref['B1'] = 'Вид работ (Классификация)', 'Ед. изм. в отчёте'
    cats = [('Строительство', 'м³'), ('Ремонт', 'м³'), ('Содержание', 'км'), ('Асфальтирование', 'м²'), ('Транспортные услуги', 'маш.-ч'),
            ('Дробление', 'м³'), *[(c['name'], c['unit']) for c in EXTRA_CATS], ('Не учитывать', '—')]
    for i, (a, b) in enumerate(cats, 2):
        ref.cell(i, 1, a), ref.cell(i, 2, b)
    for c in ref[1]:
        c.font, c.fill = Font(bold=True, color='FFFFFF'), F_HEAD
    ref.column_dimensions['A'].width, ref.column_dimensions['B'].width = 30, 18
    cat_range = f'Справочник!$A$2:$A${len(cats) + 1}'

    for p in PROJECTS:
        M = data[p['id']].get(ym)
        if not M:
            continue
        ws = wb.create_sheet(p['sheet'])
        ws.sheet_view.showGridLines = False
        head = [f'План работ на {MONTHS[m - 1]} {y} г.', 'Участок: Вилюй', f'Проект: {p["title"]}', f'Ответственный: {p["person"]}']
        for i, t in enumerate(head, 1):
            ws.cell(i, 1, t).font = Font(bold=(i == 1), size=14 if i == 1 else 10.5, color='15191E' if i == 1 else '46505B')
        H = 6  # строка заголовка таблицы
        cols = ['№', 'Наименование работ', 'Классификация', 'Ед. изм.', *range(1, n + 1), 'Итого за месяц']
        for j, v in enumerate(cols, 1):
            c = ws.cell(H, j, v)
            c.font, c.fill, c.border = Font(bold=True, color='FFFFFF', size=10), F_HEAD, BOX
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        for d in range(1, n + 1):  # строка дней недели
            wd = dt.date(y, m, d).weekday()
            c = ws.cell(H + 1, 4 + d, DOW[wd])
            c.font, c.alignment, c.border = Font(size=8.5, color='8A939D'), Alignment(horizontal='center'), BOX
            if wd >= 5:
                c.fill = F_WE
        for j in (1, 2, 3, 4, n + 5):
            ws.cell(H + 1, j).border = BOX
        r, no, group = H + 2, 0, None
        dv = DataValidation(type='list', formula1=cat_range, allow_blank=True)
        ws.add_data_validation(dv)
        for w in M.works:
            if w.group and w.group != group:
                group = w.group
                ws.cell(r, 2, group).font = Font(bold=True, size=10)
                for j in range(1, n + 6):
                    ws.cell(r, j).fill, ws.cell(r, j).border = F_GRP, BOX
                r += 1
            no += 1
            vals = [no, w.name, CAT_NAMES[w.cls], w.unit]
            for j, v in enumerate(vals, 1):
                c = ws.cell(r, j, v)
                c.border, c.font = BOX, Font(size=10)
                c.alignment = Alignment(vertical='center', wrap_text=(j == 2), horizontal='center' if j in (1, 4) else None)
            dv.add(ws.cell(r, 3))
            for d in range(1, n + 1):
                c = ws.cell(r, 4 + d, w.plan.get(d))
                c.border, c.number_format, c.font = BOX, '#,##0.##', Font(size=9.5)
                if dt.date(y, m, d).weekday() >= 5:
                    c.fill = F_WE
            c = ws.cell(r, n + 5, f'=SUM({get_column_letter(5)}{r}:{get_column_letter(n + 4)}{r})')
            c.border, c.number_format, c.font, c.fill = BOX, '#,##0.##', Font(bold=True, size=10), F_TOT
            r += 1
        for extra in range(3):  # пустые строки для новых работ
            for j in range(1, n + 6):
                ws.cell(r, j).border = BOX
            dv.add(ws.cell(r, 3))
            ws.cell(r, n + 5, f'=SUM({get_column_letter(5)}{r}:{get_column_letter(n + 4)}{r})').number_format = '#,##0.##'
            r += 1
        ws.column_dimensions['A'].width = 4.5
        ws.column_dimensions['B'].width = 52
        ws.column_dimensions['C'].width = 24
        ws.column_dimensions['D'].width = 8
        for d in range(1, n + 1):
            ws.column_dimensions[get_column_letter(4 + d)].width = 7.2
        ws.column_dimensions[get_column_letter(n + 5)].width = 12
        ws.row_dimensions[H].height = 30
        ws.freeze_panes = ws.cell(H + 2, 5)
    # проекты — сразу после инструкции, справочник — в конце
    wb.move_sheet('Справочник', offset=len(wb.sheetnames) - 2)
    wb.active = 1
    out = ROOT / plan_file(ym)
    wb.save(out)
    return out


def summary(data):
    for sid, months in data.items():
        for ym, M in sorted(months.items()):
            print(f'\n== {PRJ[sid]["short"]} · {ym} · данные по {M.as_of}-е · работ {len(M.works)}')
            for w in M.works:
                sp = sum(w.plan.values())
                spt = sum(v for d, v in w.plan.items() if d <= M.as_of)
                sf = sum(w.fact.values())
                pct = f'{sf / spt * 100:5.1f}%' if spt else '   — '
                print(f'   {w.name[:62]:62} {w.unit:4} план {sp:9.1f} (на дату {spt:9.1f}) факт {sf:9.1f} {pct}  [{CAT_NAMES[w.cls]}]')
            if M.notes:
                print('   комментарии:', M.notes)


if __name__ == '__main__':
    data = build()
    summary(data)
    for ym in ('2026-09', '2026-08'):
        print('план-файл:', write_plan_xlsx(ym, data).name)
    inject(to_vil(data))
    print('данные вшиты в', HTML.name)
