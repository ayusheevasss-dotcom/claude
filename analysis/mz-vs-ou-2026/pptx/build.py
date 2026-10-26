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

# ---------- сетка «месяцы × объекты» ----------
def matrix_slide(title, first_col, cols, data, comment):
    """cols: [(заголовок, подзаголовок)], data: [(метка, [r…], итог?)]; r = (зт, ут, зм, ум); всё в м³."""
    sl = prs.slides.add_slide(tpl.slide_layout)
    for el in proto:
        sl.shapes._spTree.append(copy.deepcopy(el))
    sh = {s.name: s for s in sl.shapes}
    set_run_text(sh['TextBox 913'], title)
    set_run_text(sh['TextBox 921'], first_col)
    for nm in ('TextBox 922', 'TextBox 923', 'TextBox 924', 'TextBox 925', 'TextBox 926', 'TextBox 927', 'TextBox 928'):
        sh[nm]._element.getparent().remove(sh[nm]._element)
    x0 = 1.8 * IN
    cw = (LEFT + WIDTH - 0.08 * IN - x0) / len(cols)
    hy = 1.17 * IN
    for j, (h1, h2) in enumerate(cols):
        text(sl, x0 + j * cw, hy, cw, 0.17 * IN, h1, 9.5, '1F2933', True)
        text(sl, x0 + j * cw, hy + 0.17 * IN, cw, 0.13 * IN, h2, 7.5, '6B7785')
    top = 1.55 * IN
    n = len(data)
    pitch = min(0.6 * IN, (BOTTOM + 0.25 * IN - top) / n)
    h = pitch * 0.88
    for i, (label, rs, total) in enumerate(data):
        y = top + i * pitch
        box(sl, LEFT, y, WIDTH, h, MSO_SHAPE.ROUNDED_RECTANGLE,
            fill='F4F5F6' if total else 'FFFFFF', line='DDE1E4', adj=0.12)
        text(sl, CX['name'][0], y, 1.25 * IN, h, label, 10.5, bold=True, align=PP_ALIGN.LEFT)
        for j, r in enumerate(rs):
            cx = x0 + j * cw
            if r is None:
                text(sl, cx, y, cw, h, '—', 10, '9AA5AD')
                continue
            z, u = r[2], r[3]
            d = u - z; pc = d / z * 100
            st = status(pc); _, pf, pt, _ = COL[st]
            light = j == len(rs) - 1  # светофор только в итоговом столбце
            pw, ph = (0.64 if len(cols) > 4 else 0.70) * IN, h * 0.74
            pill = box(sl, cx + 0.04 * IN, y + (h - ph) / 2, pw, ph, MSO_SHAPE.ROUNDED_RECTANGLE, fill=pf, adj=0.25)
            sg = '' if abs(pc) < 0.05 else ('+' if pc > 0 else '−')
            text(sl, 0, 0, 0, 0, sg + f'{abs(pc):.1f}'.replace('.', ',') + '%', 10.5, pt, True, shape=pill)
            if light:
                lw, lh = 0.40 * IN, 0.16 * IN
                lx, ly = cx + 0.04 * IN + pw + 0.05 * IN, y + (h - lh) / 2
                box(sl, lx, ly, lw, lh, MSO_SHAPE.ROUNDED_RECTANGLE, fill='3A4550', adj=0.5)
                od = lh * 0.72; gp = (lw - 3 * od) / 4
                for k, key in enumerate('rog'):
                    box(sl, lx + gp + k * (od + gp), ly + (lh - od) / 2, od, od, MSO_SHAPE.OVAL,
                        fill=COL[key][0] if key == st else '5A6570')
                tx = lx + lw + 0.06 * IN
            else:
                tx = cx + 0.04 * IN + pw + 0.10 * IN
            tw = cw - (tx - cx) - 0.02 * IN
            text(sl, tx, y + h * 0.08, tw, h * 0.46, 'Δ ' + num(d, True), 9, '1F2933', total, PP_ALIGN.LEFT)
            text(sl, tx, y + h * 0.52, tw + 0.1 * IN, h * 0.42, num(z) + ' → ' + num(u), 7, '6B7785', False, PP_ALIGN.LEFT)
    cy = top + n * pitch + 0.05 * IN
    cb = box(sl, LEFT, cy, 9198864, min(6.62 * IN - cy, 1161288), MSO_SHAPE.ROUNDED_RECTANGLE, fill='FFFFFF', line='DDE1E4', adj=0.08)
    tf = cb.text_frame; tf.word_wrap = True; tf.vertical_anchor = MSO_ANCHOR.TOP
    tf.margin_left = tf.margin_right = Emu(118872); tf.margin_top = Emu(54864)
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.LEFT
    for t, b, c in (('Комментарии: ', True, '6B7785'), (comment, False, '404040')):
        rr = p.add_run(); rr.text = t
        rr.font.size = Pt(9); rr.font.bold = b; rr.font.name = 'Arial'; rr.font.color.rgb = rgb(c)
    return sl

def pct3(r): return (r[3] - r[2]) / r[2] * 100

# 1. Общий слайд: горная масса, тыс. м³, по месяцам
OBJ = [('gross', 'Гросс'), ('kv', 'КВ Гросс'), ('tab', 'Таборный'), ('tokko', 'Токко')]
data = []
for m in MONTHS:
    rs = [obj_month(k, m) for k, _ in OBJ]
    data.append((MN[m - 1], rs + [add(rs[0], rs[2], rs[3])], False))
tot = [add(*[d[1][j] for d in data]) for j in range(5)]
data.append(('Итого 9 мес.', tot, True))
cnt = lambda j: sum(abs(pct3(d[1][j])) <= TOL for d in data[:-1])
matrix_slide('Сходимость оперативного учёта и маркзамера · горная масса, тыс. м³ · январь–сентябрь 2026', 'Месяц',
    [('Гросс', 'Δ %  ·  Δ  ·  замер → учёт'), ('КВ Гросс', 'укладка на карты'), ('Таборный', 'с Врезанным'),
     ('Токко', 'горная масса в целике'), ('Всего', 'три карьера, без КВ')], data,
    'Δ = учёт − замер, %% от замера; допуск ±5%%. За 9 месяцев все объекты в допуске: Гросс %s%%, КВ %s%%, Таборный %s%%, Токко %s%%, '
    'всего %s%%. Помесячно в допуске: Гросс %d из 9, Таборный %d, Токко %d, КВ — %d: укладку ОУ и замер относят к разным месяцам '
    '(январь в ОУ 578 тыс. м³ против 1 565 по замеру). По трём карьерам вне допуска январь (+7,5%%) и март (−6,1%%): объёмы сдвигаются между месяцами.'
    % tuple([num(pct3(tot[j]), True) for j in range(5)] + [cnt(0), cnt(2), cnt(3), cnt(1)]))

# 2–3. Стадии Гросс и Таборный, июль–август (сетка «месяцы × стадии»)
def stage_matrix(title, cols, fn, keys, comment):
    data = []
    for m in (7, 8):
        rs = [fn(m, k) for k in keys]
        data.append((MN[m - 1], rs + [add(*rs)], False))
    data.append(('Итого 2 мес.', [add(*[d[1][j] for d in data]) for j in range(len(keys) + 1)], True))
    matrix_slide(title, 'Месяц', cols, data, comment)

stage_matrix('Гросс · сходимость ОУ и маркзамера по стадиям · тыс. м³ · июль–август 2026',
    [('ST3', 'Δ %  ·  Δ  ·  замер → учёт'), ('ST4', 'стадия 4'), ('ST5', 'стадия 5'), ('ST6', 'стадия 6'), ('Карьер', 'стадии 3–6')],
    stage_gross, ['ST3', 'ST4', 'ST5', 'ST6'],
    'Разбивка замера по стадиям есть только в справках за июль и август (в сентябрьском файле лист стадий — копия августа). '
    'Вне допуска: июль ST3 (учёт выше замера на 8,9%) и август ST5 (+21,8%), при этом по карьеру в целом — в допуске. '
    'Похоже, часть объёма в ОУ отнесена к соседней стадии (август: ST4 −3,7%, ST5 +21,8%).')
stage_matrix('Таборный · сходимость ОУ и маркзамера по стадиям · тыс. м³ · июль–август 2026',
    [('W4', 'Δ %  ·  Δ  ·  замер → учёт'), ('E3', 'стадия E3'), ('E4', 'стадия E4'), ('Врезанный', 'VREZ'), ('Карьер', 'W4, E3, E4, Врезанный')],
    stage_tab, ['WST4', 'EST3', 'EST4', 'VREZ'],
    'Стадии W4, E3, E4 по замеру есть только за июль и август. Врезанный: замер выше учёта на 9–12% — ОУ занижает выемку; '
    'E3: учёт выше замера на 8–15%. Итог по карьеру в допуске за счёт взаимной компенсации стадий.')

# 4. Токко по стадиям: месяцы × стадии
data = []
for m in MONTHS:
    rs = [obj_month('tk%d' % s, m) for s in (1, 2, 3)]
    data.append((MN[m - 1], rs + [add(*rs)], False))
tot = [add(*[d[1][j] for d in data]) for j in range(4)]
data.append(('Итого 9 мес.', tot, True))
matrix_slide('Токко · сходимость ОУ и маркзамера по стадиям · тыс. м³ · январь–сентябрь 2026', 'Месяц',
    [('Стадия 1', 'Δ %  ·  Δ  ·  замер → учёт'), ('Стадия 2', 'ST-2'), ('Стадия 3', 'ST-3'), ('Карьер', 'стадии 1–3')], data,
    'Замер — маркшейдерская справка Токко по стадиям; учёт — места погрузки ST-1…ST-3. По стадиям отклонения больше, чем по карьеру: '
    'в феврале стадия 2 −32%%, стадия 3 +32%% — объём отнесён в ОУ к другой стадии. В июне–августе стадия 3 завышена в ОУ '
    '(+13…19%%); за 9 месяцев: стадия 1 %s%%, стадия 2 %s%%, стадия 3 %s%%.' % tuple(num(pct3(tot[j]), True) for j in range(3)))

# удаляем исходный слайд-образец
sldIdLst = prs.slides._sldIdLst
first = sldIdLst[0]
prs.part.drop_rel(first.rId); sldIdLst.remove(first)
out = os.path.join(HERE, 'Сходимость_ОУ_и_маркзамера_2026.pptx')
prs.save(out)
print('saved', out, len(prs.slides))
