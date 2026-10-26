"""Сборка презентации «Сходимость ОУ и маркзамера» по образцу template.pptx.
Δ = учёт − замер; отклонение, % = Δ / замер (как в образце). Допуск ±5%."""
import copy, json, os
from pptx import Presentation
from pptx.util import Emu, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

HERE = os.path.dirname(os.path.abspath(__file__))
D = json.load(open(os.path.join(HERE, '..', 'calc', 'data.json')))
MZ = {int(k): v for k, v in D['mz'].items()}
OU = {int(k): v for k, v in D['ou'].items()}
MONTHS = range(1, 10)
MN = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь', 'Июль', 'Август', 'Сентябрь']
TOL = 5.0
IN = 914400

# ---------- данные: (замер_т, учёт_т, замер_м3, учёт_м3) ----------
def obj_month(key, m):
    z, o = MZ[m], OU[m]
    if key == 'gross':
        d = o['gross'][0] / o['gross'][1]
        return (z['gross_t'] or z['gross_m3'] * d, o['gross'][0], z['gross_m3'], o['gross'][1])
    if key == 'kv':
        return (z['kv_t'], o['kv'][0], z['kv_m3'], o['kv'][1])
    if key == 'tab':
        m3 = z['tab_m3'] + z['vrez_m3']
        d = o['tab'][0] / o['tab'][1]
        t = (z['tab_t'] + z['vrez_t']) if z['tab_t'] else m3 * d
        return (t, o['tab'][0], m3, o['tab'][1])
    if key == 'tokko':
        return (z['tokko0_t'], o['tokko'][0], z['tokko0_m3'], o['tokko'][1])
    if key.startswith('tk'):
        s = int(key[2]); b = o['k_ST-%d' % s]
        return (z['tokko%d_t' % s], b[0], z['tokko%d_m3' % s], b[1])

def stage_gross(m, st):
    b = OU[m]['g_' + st]; v = MZ[m]['grst'][st]
    return (v * b[0] / b[1], b[0], v, b[1])

def stage_tab(m, st):
    b = OU[m]['t_' + st]
    v = MZ[m]['vrez_m3'] if st == 'VREZ' else MZ[m]['tabst'][st]
    t = MZ[m]['vrez_t'] if (st == 'VREZ' and MZ[m]['vrez_t']) else v * b[0] / b[1]
    return (t, b[0], v, b[1])

def add(*rows):
    return tuple(sum(r[i] for r in rows) for i in range(4))

# ---------- форматирование ----------
def num(x, sign=False):
    s = f'{abs(x):,.1f}'.replace(',', ' ').replace('.', ',')
    if sign:
        return ('+' if x > 0 else '−' if x < 0 else '') + s
    return ('−' if x < 0 else '') + s

def status(p):
    a = abs(p)
    if a <= TOL: return 'g'
    if a <= TOL + 3: return 'o'
    return 'r'

COL = {'g': ('2E7D6B', 'E4EFEC', '2E7D6B', 'Норма'),
       'o': ('D98C1F', 'F8ECDB', 'A8650F', 'Внимание'),
       'r': ('B53F2F', 'F6E2DF', 'B53F2F', 'Критично')}

# ---------- примитивы ----------
def rgb(h): return RGBColor.from_string(h)

def box(sl, x, y, w, h, shape=MSO_SHAPE.RECTANGLE, fill=None, line=None, adj=None, name=None):
    s = sl.shapes.add_shape(shape, Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    if adj is not None: s.adjustments[0] = adj
    if fill: s.fill.solid(); s.fill.fore_color.rgb = rgb(fill)
    else: s.fill.background()
    if line: s.line.color.rgb = rgb(line); s.line.width = Emu(9525)
    else: s.line.fill.background()
    s.shadow.inherit = False
    if name: s.name = name
    return s

def text(sl, x, y, w, h, t, size, color='1F2933', bold=False, align=PP_ALIGN.CENTER,
         anchor=MSO_ANCHOR.MIDDLE, shape=None, name=None):
    s = shape or sl.shapes.add_textbox(Emu(int(x)), Emu(int(y)), Emu(int(w)), Emu(int(h)))
    tf = s.text_frame; tf.word_wrap = True
    tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = 0
    tf.vertical_anchor = anchor
    p = tf.paragraphs[0]; p.alignment = align
    r = p.add_run(); r.text = t
    f = r.font; f.size = Pt(size); f.bold = bold; f.name = 'Arial'; f.color.rgb = rgb(color)
    if name: s.name = name
    return s

# колонки образца (x, ширина) в EMU
CX = {'name': (786384, 3108960), 'z': (4023360, 1097280), 'u': (5120640, 1097280),
      'd1': (6217920, 960120), 'd2': (7178040, 960120), 'tol': (8183879, 868680),
      'pct': (9098280, 1005840), 'st': (10314432, 1060704)}
LEFT, WIDTH = 603504, 10972800
TOP, BOTTOM = 1481328, 5212080  # область строк; ниже — комментарии

def row(sl, y, h, label, r, unit, total=False, caption=None):
    zt, ut, zm, um = r
    z, u, z2, u2 = (zt, ut, zm, um) if unit == 't' else (zm, um, zt, ut)
    d, d2 = u - z, u2 - z2
    p = d / z * 100
    st = status(p); bar, pf, pt, lab = COL[st]
    full = h >= 0.6 * IN
    box(sl, LEFT, y, WIDTH, h, MSO_SHAPE.ROUNDED_RECTANGLE,
        fill='F4F5F6' if total else 'FFFFFF', line='DDE1E4', adj=0.08)
    box(sl, LEFT, y + h * 0.12, 54864, h * 0.76, fill=bar)
    nsz = 11.5 if full else 10.5
    vsz = 13 if full else 11.5
    if caption and full:
        text(sl, CX['name'][0], y + h * 0.1, CX['name'][1], 0.15 * IN, caption, 7.5, '9AA5AD', True, PP_ALIGN.LEFT, MSO_ANCHOR.TOP)
        text(sl, CX['name'][0], y + h * 0.3, CX['name'][1], 0.25 * IN, label, nsz, bold=True, align=PP_ALIGN.LEFT)
    else:
        text(sl, CX['name'][0], y, CX['name'][1], h, label, nsz, bold=True, align=PP_ALIGN.LEFT)
    b = total
    text(sl, *CX['z'][:1], y, CX['z'][1], h, num(z), vsz, bold=b)
    text(sl, CX['u'][0], y, CX['u'][1], h, num(u), vsz, bold=b)
    text(sl, CX['d1'][0], y, CX['d1'][1], h, num(d, True), vsz, bold=b)
    text(sl, CX['d2'][0], y, CX['d2'][1], h, num(d2, True), vsz, '6B7785', bold=b)
    text(sl, CX['tol'][0], y, CX['tol'][1], h, '± 5,0%', vsz, '6B7785')
    ph = min(438911, h * 0.78)
    pill = box(sl, CX['pct'][0], y + (h - ph) / 2, CX['pct'][1], ph, MSO_SHAPE.ROUNDED_RECTANGLE, fill=pf, adj=0.2)
    pct = ('+' if p > 0 else '−' if p < 0 else '') + f'{abs(p):.1f}'.replace('.', ',') + '%'
    text(sl, 0, 0, 0, 0, pct, 17 if full else 13, pt, True, shape=pill)
    # светофор
    if full:
        lw, lh, lx, ly = 932688, 329184, 10378440, y + h * 0.16
    else:
        lw, lh = 640080, h * 0.68
        lx, ly = CX['st'][0] - 30000, y + (h - lh) / 2
    box(sl, lx, ly, lw, lh, MSO_SHAPE.ROUNDED_RECTANGLE, fill='3A4550', adj=0.5)
    od = lh * 0.72; gap = (lw - 3 * od) / 4
    for i, key in enumerate('rog'):
        c = COL[key][0] if key == st else '5A6570'
        box(sl, lx + gap + i * (od + gap), ly + (lh - od) / 2, od, od, MSO_SHAPE.OVAL, fill=c)
    if full:
        text(sl, CX['st'][0], ly + lh + 30000, CX['st'][1], 130805, lab, 8.5, pt, True)
    else:
        text(sl, lx + lw + 50000, y, 600000, h, lab, 8, pt, True, PP_ALIGN.LEFT)
    return p

# ---------- шаблонные элементы ----------
prs = Presentation(os.path.join(HERE, 'template.pptx'))
tpl = prs.slides[0]
KEEP = {'Freeform 6', 'Group 10', 'TextBox 913', 'Rectangle 915', 'TextBox 916', 'Rectangle 917', 'TextBox 918',
        'Rectangle 919', 'TextBox 920', 'TextBox 921', 'TextBox 922', 'TextBox 923', 'TextBox 924', 'TextBox 925',
        'TextBox 926', 'TextBox 927', 'TextBox 928'}
proto = [copy.deepcopy(s._element) for s in tpl.shapes if s.name in KEEP]

def set_run_text(shape, t):
    runs = [r for p in shape.text_frame.paragraphs for r in p.runs]
    runs[0].text = t
    for r in runs[1:]: r.text = ''

def new_slide(title, first_col, unit, rows, comment):
    sl = prs.slides.add_slide(tpl.slide_layout)
    for el in proto: sl.shapes._spTree.append(copy.deepcopy(el))
    sh = {s.name: s for s in sl.shapes}
    set_run_text(sh['TextBox 913'], title)
    set_run_text(sh['TextBox 921'], first_col)
    u1, u2 = ('тыс. т', 'тыс. м³') if unit == 't' else ('тыс. м³', 'тыс. т')
    set_run_text(sh['TextBox 922'], 'Замер, ' + u1)
    set_run_text(sh['TextBox 923'], 'Учёт, ' + u1)
    set_run_text(sh['TextBox 924'], 'Δ, ' + u1)
    set_run_text(sh['TextBox 925'], 'Δ, ' + u2)
    set_run_text(sh['TextBox 927'], 'Отклонение, ' + ('т' if unit == 't' else 'м³'))
    n = len(rows)
    pitch = min(0.9 * IN, (BOTTOM - TOP) / n)
    h = pitch * (0.91 if pitch >= 0.6 * IN else 0.88)
    for i, r in enumerate(rows):
        row(sl, TOP + i * pitch, h, r[0], r[1], unit, total=r[2] if len(r) > 2 else False,
            caption=r[3] if len(r) > 3 else None)
    cy = TOP + n * pitch + 0.08 * IN
    ch = 6.62 * IN - cy - 0.0 * IN
    ch = min(ch, 1161288)
    cb = box(sl, LEFT, cy, 9198864, ch, MSO_SHAPE.ROUNDED_RECTANGLE, fill='FFFFFF', line='DDE1E4', adj=0.08)
    tf = cb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = Emu(118872); tf.margin_top = Emu(64008)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    r = p.add_run(); r.text = 'Комментарии: '
    r.font.size = Pt(9); r.font.bold = True; r.font.name = 'Arial'; r.font.color.rgb = rgb('6B7785')
    r = p.add_run(); r.text = comment
    r.font.size = Pt(9); r.font.name = 'Arial'; r.font.color.rgb = rgb('404040')
    return sl

PER = 'январь–сентябрь 2026'
NOTE_T = 'Тонны замера за январь–июль по Гроссу и Таборному пересчитаны из м³ по плотности ОУ (в справках МЗ тонны есть только с августа).'

def pct_of(r, unit):
    z, u = (r[0], r[1]) if unit == 't' else (r[2], r[3])
    return (u - z) / z * 100

OBJ = [('gross', 'Гросс'), ('kv', 'Карты выщелачивания Гросс'), ('tab', 'Таборный'), ('tokko', 'Токко')]

# 1. Сводка за 9 месяцев
for unit in ('t', 'm3'):
    rs = {k: add(*[obj_month(k, m) for m in MONTHS]) for k, _ in OBJ}
    tot = add(rs['gross'], rs['tab'], rs['tokko'])
    rows = [(nm, rs[k]) for k, nm in OBJ] + [('Всего', tot, True, 'ТРИ КАРЬЕРА')]
    ul = 'тыс. т' if unit == 't' else 'тыс. м³'
    cm = ('Итог за 9 месяцев по трём карьерам в допуске. Вне допуска только карты выщелачивания Гросс по тоннам: '
          'учёт выше замера на %s%%. «Всего» — Гросс + Таборный + Токко, без КВ. ' % num(pct_of(rs['kv'], 't'))) + \
         ('Помесячно разброс заметно больше — см. следующие слайды. ' + (NOTE_T if unit == 't' else ''))
    new_slide('Сходимость оперативного учёта и маркзамера · итог %s · %s' % (PER, ul), 'Карьер', unit, rows, cm)

# 2. По месторождениям, по месяцам
COMM = {
 'gross': None,
 'kv': 'Помесячное расхождение до ±60%%: ОУ и маркзамер по укладке относят объёмы к разным месяцам (январь в ОУ — 0,9 млн т против 2,3 млн т по замеру). Объём замера — сумма замеренных панелей. ',
 'tab': 'Таборный с Врезанным. Вне допуска январь и март (сдвиг объёмов между месяцами), с апреля — в допуске. ',
 'tokko': 'Замер — маркшейдерская справка Токко (горная масса в целике). Вне допуска 4–5 месяцев: февраль–март замер выше учёта, июнь–июль ниже. ',
}
for k, nm in OBJ:
    for unit in ('t', 'm3'):
        rows = [(MN[m - 1], obj_month(k, m)) for m in MONTHS]
        rows.append(('Итого 9 мес.', add(*[r[1] for r in rows]), True))
        ul = 'тыс. т' if unit == 't' else 'тыс. м³'
        extra = NOTE_T if (unit == 't' and k in ('gross', 'tab')) else ''
        if k == 'kv' and unit == 't':
            extra = 'Плотность в ОУ 1,66 т/м³ против 1,51 т/м³ по замеру: в м³ итог в допуске, в тоннах — нет.'
        cm = COMM[k]
        if cm is None:
            ps = [pct_of(r[1], unit) for r in rows[:-1]]
            cm = ('В допуске %d месяцев из 9; наибольшее отклонение — %s (%s%%). Учёт выше замера в %d месяцах из 9 — ОУ систематически завышает выемку. '
                  % (sum(abs(x) <= TOL for x in ps), MN[max(range(9), key=lambda i: abs(ps[i]))].lower(),
                     num(max(ps, key=abs), True), sum(x > 0 for x in ps)))
        new_slide('%s · сходимость ОУ и маркзамера по месяцам · %s' % (nm, ul), 'Месяц', unit, rows, cm + extra)

# 3. По стадиям
for unit in ('t', 'm3'):
    rows = []
    for m in (7, 8):
        sts = ['ST3', 'ST4', 'ST5', 'ST6']
        rs = [stage_gross(m, s) for s in sts]
        rows += [('%s · %s' % (MN[m - 1], s), r) for s, r in zip(sts, rs)]
        rows.append(('%s · итого' % MN[m - 1], add(*rs), True))
    ul = 'тыс. т' if unit == 't' else 'тыс. м³'
    new_slide('Гросс · сходимость по стадиям · июль–август 2026 · %s' % ul, 'Месяц · стадия', unit, rows,
              'Разбивка замера по стадиям есть только в справках за июль и август (в сентябрьском файле лист стадий — копия августа). '
              'Тонны по стадиям — объём замера × плотность ОУ стадии. Июль ST3 и август ST5: учёт выше замера на 9–22%.')
for unit in ('t', 'm3'):
    rows = []
    for m in (7, 8):
        sts = [('WST4', 'W4'), ('EST3', 'E3'), ('EST4', 'E4'), ('VREZ', 'Врезанный')]
        rs = [stage_tab(m, s) for s, _ in sts]
        rows += [('%s · %s' % (MN[m - 1], lb), r) for (_, lb), r in zip(sts, rs)]
        rows.append(('%s · итого' % MN[m - 1], add(*rs), True))
    ul = 'тыс. т' if unit == 't' else 'тыс. м³'
    new_slide('Таборный · сходимость по стадиям · июль–август 2026 · %s' % ul, 'Месяц · стадия', unit, rows,
              'Стадии W4, E3, E4 по замеру — только июль и август. Врезанный: замер выше учёта на 9–12%, E3 — учёт выше замера на 8–15%. '
              'Тонны W4/E3/E4 и июль Врезанного — объём замера × плотность ОУ.')
for s in (1, 2, 3):
    for unit in ('t', 'm3'):
        rows = [(MN[m - 1], obj_month('tk%d' % s, m)) for m in MONTHS]
        rows.append(('Итого 9 мес.', add(*[r[1] for r in rows]), True))
        ul = 'тыс. т' if unit == 't' else 'тыс. м³'
        new_slide('Токко · стадия %d · сходимость ОУ и маркзамера по месяцам · %s' % (s, ul), 'Месяц', unit, rows,
                  'Замер — маркшейдерская справка Токко по стадии %d; учёт — место погрузки ST-%d. '
                  'По стадиям отклонения больше, чем по карьеру в целом: часть объёма отнесена в ОУ к другой стадии.' % (s, s))

# удаляем исходный слайд-образец
sldIdLst = prs.slides._sldIdLst
first = sldIdLst[0]
prs.part.drop_rel(first.rId); sldIdLst.remove(first)
out = os.path.join(HERE, 'Сходимость_ОУ_и_маркзамера_2026.pptx')
prs.save(out)
print('saved', out, len(prs.slides))
