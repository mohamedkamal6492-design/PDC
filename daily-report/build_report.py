"""PDC Mecca daily report builder.
Usage: python3 build_report.py template.pptx data.json out.pptx
Any key missing/null in data.json leaves that part of the template untouched
and is listed in the printed "MISSING" report.
"""
import copy, json, sys
from pptx import Presentation
from pptx.util import Pt, Emu
from pptx.oxml.ns import qn

src, data_path, out = sys.argv[1:4]
D = json.load(open(data_path, encoding='utf-8'))
p = Presentation(src)
S1, S2, S3 = p.slides[0], p.slides[1], p.slides[2]
missing = []

def sh(slide, sid):
    for s in slide.shapes:
        if s.shape_id == sid:
            return s
    raise KeyError(sid)

def set_para(para, *texts):
    for i, r in enumerate(para.runs):
        r.text = texts[i] if i < len(texts) else ''

def st(slide, sid, *texts, para=0):
    set_para(sh(slide, sid).text_frame.paragraphs[para], *texts)

def set_paras(shape, items):
    tf = shape.text_frame
    paras = tf.paragraphs
    t0 = copy.deepcopy(paras[0]._p)
    t1 = copy.deepcopy(paras[1]._p) if len(paras) > 1 else t0
    body = tf._txBody
    for pp in list(paras):
        body.remove(pp._p)
    for i, _ in enumerate(items):
        body.append(copy.deepcopy(t0 if i == 0 else t1))
    for pp, it in zip(tf.paragraphs, items):
        set_para(pp, *(it if isinstance(it, (tuple, list)) else (it,)))

def g(path):
    """Fetch nested value; record as missing if absent/None/'—'."""
    cur = D
    for k in path.split('.'):
        if not isinstance(cur, dict) or cur.get(k) is None:
            missing.append(path); return None
        cur = cur[k]
    if cur in ('', '—', '-'):
        missing.append(path); return None
    return cur

def f(v):  # number formatting
    return f'{v:,}' if isinstance(v, int) else str(v)

def put(slide, sid, path, fmt=lambda v: f(v), para=0):
    v = g(path)
    if v is not None:
        st(slide, sid, fmt(v), para=para)

date = g('date')  # e.g. "الإثنين 10/11/2026"
if date:
    st(S1, 40, date); st(S2, 40, date)

# ---------- Slide 1: construction ----------
c = 'construction'
put(S1, 48, f'{c}.assigned_total')
v = g(f'{c}.assigned_connections')
if v is not None: st(S1, 50, f'التوصيلات {v}')
v = g(f'{c}.assigned_projects')
if v is not None:
    runs = sh(S1, 51).text_frame.paragraphs[0].runs
    st(S1, 51, *(('المشاريع ', str(v)) if len(runs) > 1 else (f'المشاريع {v}',)))
put(S1, 55, f'{c}.wo_190_203')
put(S1, 57, f'{c}.wo_155')
v = g(f'{c}.cert_accepted_today')
if v is not None: st(S1, 62, f'قُبلت {v} اليوم')
put(S1, 63, f'{c}.cert_ready_156')
put(S1, 65, f'{c}.cert_accepted_170')
put(S1, 67, f'{c}.cert_rejected_171')
put(S1, 72, f'{c}.fines_amount', lambda v: f'{v:,.2f}')
v = g(f'{c}.fines_certificates')
if v is not None: st(S1, 73, f'ريال · {v} شهادة')
put(S1, 403, f'{c}.close_300')
put(S1, 405, f'{c}.close_230')
a, b = g(f'{c}.close_with_fine'), g(f'{c}.close_without_fine')
if a is not None and b is not None: st(S1, 407, f'بغرامة {a}/بدون {b}')

# asset receipt (overall)
r = 'assets'
for sid, k in [(90, 'accepted'), (92, 'in_progress'), (94, 'late'), (96, 'rejected_with_notes')]:
    put(S1, sid, f'{r}.overall.{k}')
for sid, k in [(101, 'projects'), (103, 'connections'), (105, 'haram'), (107, 'khulais'), (109, 'allith')]:
    put(S1, sid, f'{r}.by_entity.{k}')
v = g(f'{r}.pdc.total')
if v is not None: st(S1, 114, f'الإجمالي {v}')
for sid, k in [(116, 'accepted'), (119, 'in_progress'), (122, 'ready_to_send'), (125, 'need_review'),
               (128, 'rejected_with_notes'), (131, 'late')]:
    v = g(f'{r}.pdc.{k}')
    st(S1, sid, f(v) if v is not None else '—')

# attachments pending table
att = g('attachments_pending')  # {"consultant": [at_consultant, at_contractor], ...}
if att:
    order = sorted(att, key=lambda k: (-att[k][0], -att[k][1]))
    st(S1, 411, f'{sum(x[0] + x[1] for x in att.values())} مرفقاً')
    tbl = sh(S1, 412).table
    for row, name in zip(list(tbl.rows)[1:], order):
        for cell, val in zip(row.cells, (name, str(att[name][0]), str(att[name][1]))):
            set_para(cell.text_frame.paragraphs[0], val)

# ---------- Slide 2: SAP ----------
sap = g('sap')
if sap:
    N = sum(sap[k]['count'] for k in ('completed', 'in_review', 'not_started'))
    A = sum(sap[k]['amount'] for k in ('completed', 'in_review', 'not_started'))
    st(S2, 48, f'{N:,}'); st(S2, 50, f'{A:,.2f}')
    pct = lambda x: f'{x*100:.1f}%' if x < .1 else f'{round(x*100)}%'
    for k, (pid, cid, mid, shid) in zip(('completed', 'in_review', 'not_started'),
                                         [(55, 56, 57, 58), (62, 63, 64, 65), (69, 70, 71, 72)]):
        n, am = sap[k]['count'], sap[k]['amount']
        st(S2, pid, f'{round(n / N * 100)}%')
        st(S2, cid, str(n), '  أمر')
        st(S2, mid, f'{am / 1e6:.2f} مليون ﷼')
        st(S2, shid, pct(am / A))
put(S2, 76, 'sap_extra.generator_receipts')
put(S2, 81, 'sap_extra.temp_cables')

# holy sites
m = 'mashaer'
v = g(f'{m}.new_today')
if v is not None: st(S2, 86, 'الجديد اليوم', ' ', f': {v}', '     ')
for sid, k in [(87, 'total'), (89, 'replacements_total'), (91, 'cabinets_total'), (93, 'meters_total'), (95, 'stations_total')]:
    put(S2, sid, f'{m}.{k}')
rep = [g(f'{m}.replacements.{k}') for k in ('done', 'in_progress', 'awaiting_permits', 'surveyed')]
tot = g(f'{m}.replacements_total')
if None not in rep and tot:
    for sid, v in zip((101, 103, 105, 107), rep):
        st(S2, sid, str(v))
    st(S2, 100, f'              {rep[0] / tot * 100:.1f}%  ')
    track, fill = sh(S2, 109), sh(S2, 110)
    w = int(track.width * rep[0] / tot)
    fill.width = Emu(w); fill.left = Emu(track.left + track.width - w)
cd, ct = g(f'{m}.cabinets_done'), g(f'{m}.cabinets_total')
if cd is not None and ct is not None:
    st(S2, 115, str(cd)); st(S2, 116, f'منجز {cd} من {ct}')
put(S2, 120, f'{m}.meters_done')
put(S2, 125, f'{m}.stations_done')

# ---------- Slide 3: safety ----------
sf = g('safety')
if sf:
    st(S3, 940, sf['report_label'])
    t, cu = sf['today'], sf['cumulative']
    k = lambda sid, i, txt: st(S3, sid, txt, para=i)
    k(202, 0, f(t['visits']))
    k(202, 2, f"سلامة: PMSS ‏{t['pmss_safety']} · CMS ‏{t['cms']}")
    k(202, 3, f"تنفيذ: PMSS ‏{t['pmss_exec']}")
    k(203, 0, f(t['workshops']))
    k(204, 0, f(t['observations'])); k(204, 2, f"{t['obs_closed']} مغلقة · {t['obs_open']} مفتوحة")
    k(205, 0, f(t['violations'])); k(205, 2, f"سلامة {t['viol_safety']} · جودة {t['viol_quality']}")
    k(206, 0, f(t['viol_amount'])); k(206, 2, f"سلامة {t['amount_safety']:,} · جودة {t['amount_quality']:,}")
    k(207, 0, f(t['stops']))
    k(209, 0, f(cu['visits'])); k(210, 0, f(cu['workshops']))
    k(211, 0, f(cu['observations'])); k(211, 2, f"{cu['obs_closed']:,} مغلقة · {cu['obs_open']:,} مفتوحة")
    k(211, 3, f"نسبة الإغلاق {cu['close_rate']}")
    k(212, 0, f(cu['violations'])); k(212, 2, f"سلامة {cu['viol_safety']} · جودة {cu['viol_quality']}")
    k(213, 0, f(cu['viol_amount'])); k(213, 0, f(cu['viol_amount']))
    k(214, 0, f(cu['stops']))

    # observations: split into two side-by-side boxes
    obs = sf.get('observations_list') or []
    q = sf.get('quality_observations_list') or []
    st(S3, 34, 'أبرز ملاحظات ', 'الجودة', ' المرصودة (', str(len(q)), ')')
    if q: set_paras(sh(S3, 36), q)
    if obs:
        st(S3, 215, f'أبرز ملاحظات السلامة المرصودة ({len(obs)})')
        ob = sh(S3, 219)
        if len(obs) <= 6:
            set_paras(ob, obs)
        else:
            gap = Emu(152400); half = (ob.width - gap) // 2
            clone = copy.deepcopy(ob._element); ob._element.addnext(clone)
            ob2 = [s for s in S3.shapes if s._element is clone][0]
            clone.find('.//' + qn('p:cNvPr')).set('id', '970')
            clone.find('.//' + qn('p:cNvPr')).set('name', 'Observations 2')
            x0 = ob.left
            ob.left = Emu(x0 + half + gap); ob.width = Emu(half)
            ob2.left = Emu(x0); ob2.width = Emu(half)
            h = (len(obs) + 1) // 2
            for box, items in ((ob, obs[:h]), (ob2, obs[h:])):
                set_paras(box, items)
                for pp in box.text_frame.paragraphs:
                    sb = pp._p.find(qn('a:pPr')).find(qn('a:spcBef'))
                    if sb is not None: sb[0].set('val', '250')
                    for rr in pp.runs: rr.font.size = Pt(9.5 if len(obs) > 8 else 11)

    viol = sf.get('violations_list') or []
    if viol:
        tshape = sh(S3, 220); tb = tshape.table; tbl_el = tb._tbl
        rows = list(tb.rows)
        tmpl_tr = copy.deepcopy(rows[1]._tr); total_tr = rows[-1]._tr
        for row in rows[1:-1]: tbl_el.remove(row._tr)
        for _ in viol: total_tr.addprevious(copy.deepcopy(tmpl_tr))
        rows = list(tb.rows)
        for i, (row, v) in enumerate(zip(rows[1:-1], viol), 1):
            vals = (str(i), v['date'], v['description'], v['class'], v['consultant'], v['contractor'], f"{v['amount']:,}")
            for cell, val in zip(row.cells, vals):
                set_para(cell.text_frame.paragraphs[0], val)
        last = rows[-1].cells[len(rows[-1].cells) - 1]
        set_para(last.text_frame.paragraphs[0], f"{sum(v['amount'] for v in viol):,}")
        if len(viol) > 2:
            avail = sh(S3, 960).top - tshape.top - Emu(80000)
            rows[0].height = Emu(250000); rows[-1].height = Emu(230000)
            bh = (avail - 480000) // len(viol)
            for row in rows[1:-1]: row.height = Emu(bh)
            for row in rows:
                for cell in row.cells:
                    cell.margin_top = cell.margin_bottom = Emu(12700)
                    for pp in cell.text_frame.paragraphs:
                        for rr in pp.runs: rr.font.size = Pt(8 if len(viol) > 4 else 9.5)
            tshape.height = Emu(sum(r.height for r in rows))
        dist = {}
        for v in viol: dist[v['contractor']] = dist.get(v['contractor'], 0) + 1
        card = sh(S3, 221)
        set_paras(card, [(str(n), f'  مخالفة · {cn}') for cn, n in sorted(dist.items(), key=lambda x: -x[1])])
        if len(dist) > 2:
            for pp in card.text_frame.paragraphs:
                pp.runs[0].font.size = Pt(15); pp.runs[1].font.size = Pt(10.5)

p.save(out)
print('SAVED', out)
print('MISSING:', ', '.join(missing) if missing else 'none')
