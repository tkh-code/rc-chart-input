"""midas の MGTX と、板要素の断面力を含む ANL から、耐震壁（板要素）の応力を Excel にまとめる。

py -3 make_wall.py model.mgtx result.anl [-o 出力フォルダ]

出力：壁の応力.xlsx（概要／壁全体の合力／最大値／要素別）
値は ANL の要素中央の値（kN/m、kN·m/m）。壁全体の合力は、要素中央の値×要素幅を足し合わせたもの。
"""
import argparse
import collections
import math
import os
import re
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.dont_write_bytecode = True
import make_excel as mx  # noqa: E402

SIGMA = mx.SIGMA
NUM = re.compile(r'^[-+]?(?:\d+\.?\d*|\.\d+)(?:[Ee][-+]?\d+)?$')
FKEYS = ['Fxx', 'Fyy', 'Fxy']
MKEYS = ['Mxx', 'Myy', 'Mxy']
VKEYS = ['Vxx', 'Vyy']


def parse_plate(path):
    """板要素の断面力。{要素: {ケース: {'F': {位置: [Fxx,Fyy,Fxy,Fmax,Fmin,角度]}, 'M': …, 'V': …}}}"""
    data = open(path, 'rb').read().replace(b'\x83\xb0', SIGMA.encode()).replace(b'\xa5\xd2', SIGMA.encode())
    lines = [b.decode('cp932', 'replace').replace(SIGMA, 'Σ') for b in data.split(b'\n')]
    del data
    # 単位
    f_unit, l_unit = 'kN', 'm'
    for l in lines[:600]:
        m = re.search(r':\s*(kN|N|tonf|kgf)\s*,\s*(mm|cm|m)\b', l)
        if m:
            f_unit, l_unit = m.group(1), m.group(2)
            break
    # 短縮名 → 正式名（要約表）
    full = {}
    for i, l in enumerate(lines[:3000]):
        if '要約表' in l and '<<' in l:
            rules = 0
            for l2 in lines[i + 1:i + 200]:
                s = l2.strip()
                if s.startswith('-----'):
                    rules += 1
                    if rules == 2:
                        break
                    continue
                if rules == 1 and s:
                    t = s.split()
                    if len(t) >= 2:
                        full[t[0]] = t[1]
            break
    first = next((i for i, l in enumerate(lines) if 'Fxx' in l), None)   # 板要素の表の見出しは文字コードが違うので Fxx で探す
    start = None if first is None else max(0, first - 30)
    if start is None:
        raise ValueError('ANL に板要素の断面力の表が見つかりません（板要素の断面力を出力した ANL か確認してください）。')
    res = collections.OrderedDict()
    cases = []
    elem = case = None
    mode = None
    for l in lines[start:]:
        if 'Fxx' in l:
            mode = 'F'
        elif 'Mxx' in l:
            mode = 'M'
        elif 'Vxx' in l:
            mode = 'V'
        t = l.split()
        if len(t) < 3:
            continue
        if len(t) >= 10 and t[0].isdigit() and t[1].isdigit() and t[2].isdigit() and not NUM.match(t[3]) and all(NUM.match(x) for x in t[5:11]):
            elem, case, pos, vals, mode = int(t[0]), t[3], t[4], t[5:11], 'F'
            pos = 'C'
        elif mode == 'F' and len(t) >= 8 and not NUM.match(t[0]) and t[0] != 'LC' and all(NUM.match(x) for x in t[2:8]):
            case, pos, vals = t[0], 'C', t[2:8]
        elif elem is not None and all(NUM.match(x) for x in t[1:]) and ((mode in 'FM' and len(t) == 7) or (mode == 'V' and len(t) == 3)):
            pos = int(t[0]) if t[0].isdigit() else 'C'
            vals = t[1:]
        elif elem is not None and not NUM.match(t[0]) and all(NUM.match(x) for x in t[1:]) and ((mode in 'FM' and len(t) == 7) or (mode == 'V' and len(t) == 3)):
            pos, vals = 'C', t[1:]
        else:
            continue
        name = full.get(case, case)
        if name not in cases:
            cases.append(name)
        res.setdefault(elem, {}).setdefault(name, {}).setdefault(mode, {})[pos] = [float(v) for v in vals]
    return dict(results=res, cases=cases, units=(f_unit, l_unit))


def scale_factors(units):
    f = {'kN': 1.0, 'N': 0.001, 'tonf': 9.80665, 'kgf': 0.00980665}[units[0]]
    l = {'mm': 0.001, 'cm': 0.01, 'm': 1.0}[units[1]]
    return f / l, f   # (kN/m に直す係数, kN·m/m に直す係数)


def unit(v):
    n = math.sqrt(sum(a * a for a in v))
    return tuple(a / n for a in v)


def geometry(model, e):
    N = [model['nodes'][n] for n in e['nodes']]   # mm
    if len(N) != 4:
        return None
    sub = lambda a, b: tuple(x - y for x, y in zip(a, b))
    ex = unit(sub(N[1], N[0]))
    ey0 = sub(N[3], N[0])
    ez = unit((ex[1] * ey0[2] - ex[2] * ey0[1], ex[2] * ey0[0] - ex[0] * ey0[2], ex[0] * ey0[1] - ex[1] * ey0[0]))
    ey = (ez[1] * ex[2] - ez[2] * ex[1], ez[2] * ex[0] - ez[0] * ex[2], ez[0] * ex[1] - ez[1] * ex[0])
    width = (math.dist(N[0], N[1]) + math.dist(N[3], N[2])) / 2000
    height = (math.dist(N[0], N[3]) + math.dist(N[1], N[2])) / 2000
    cen = tuple(sum(p[i] for p in N) / 4 for i in range(3))
    return dict(ex=ex, ey=ey, ez=ez, width=width, height=height, center=cen)


def classify(name):
    if re.search(r'\+QU[XY]', name):   # 保有水平耐力時（Qu）のケースは終局
        return '終局'
    return mx.classify(name)


HEAD = mx.HEAD
BORDER = mx.BORDER
BOLD = mx.BOLD
SUMF = mx.SUMF
CLASS_COLOR = mx.CLASS_COLOR


def header(ws, row, heads, widths=None):
    for j, h in enumerate(heads):
        c = ws.cell(row, 1 + j, h)
        c.font, c.fill, c.border = BOLD, HEAD, BORDER
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
    if widths:
        for j, w in enumerate(widths):
            ws.column_dimensions[get_column_letter(1 + j)].width = w


def build(model_path, anl_path, out_path, log=print, fc=None, names=None):
    """names＝面外の柱テンプレートのファイル名 {'短期': …, '終局': …}。戻り値は面内のチャート入力用の値（chart_input）。"""
    import make_template as mt
    if fc is None:
        fc = mt.read_fc(model_path) or 24.0
    model = mx.read_mgt(model_path)
    pl = parse_plate(anl_path)
    res, cases = pl['results'], pl['cases']
    kf, km = scale_factors(pl['units'])
    plates = [e for e in sorted(model['elements'].values(), key=lambda x: x['id']) if e['type'] == 'PLATE' and e['id'] in res]
    if not plates:
        raise ValueError('ANL に載っている板要素が MGTX に見つかりません。')
    log(f'板要素 {len(plates)} 枚・荷重ケース {len(cases)}：' + ' / '.join(cases))
    gi = mx.GridIndex(model['grids'])
    geo = {e['id']: geometry(model, e) for e in plates}
    thick = {e['id']: model['thickness'].get(e['sec'], 0.0) for e in plates}
    long = mt.find_long(cases)
    classes = {c: '長期' if c == long else classify(c) for c in cases}

    # 要素×ケース：中央値
    rows = []
    for e in plates:
        g = geo[e['id']]
        label = gi.label(g['center'], g['center'][2]) if gi.available else ''
        for c in cases:
            r = res[e['id']].get(c)
            if not r or 'F' not in r or 'M' not in r or 'V' not in r:
                continue
            F, M, V = r['F']['C'], r['M']['C'], r['V']['C']
            t = thick[e['id']] or 1.0
            fxx, fyy, fxy = F[0] * kf, F[1] * kf, F[2] * kf
            rows.append(dict(e=e['id'], label=label, case=c, cls=classes[c], t=t, g=g,
                             Fxx=fxx, Fyy=fyy, Fxy=fxy, Mxx=M[0] * km, Myy=M[1] * km, Mxy=M[2] * km,
                             Vxx=V[0] * kf, Vyy=V[1] * kf,
                             tau=abs(fxy) / t, sx=fxx / t, sy=fyy / t))   # kN/m ÷ mm = N/mm2（kN/m/mm = N/mm²）
    # 壁全体の合力（壁がベースから受ける力＝要素中央の力×幅の和）
    tot = collections.OrderedDict()
    for c in cases:
        R = [0.0, 0.0, 0.0]
        sumF = {'Fxy': 0.0, 'Fyy': 0.0, 'Vyy': 0.0}
        n = 0
        for r in (x for x in rows if x['case'] == c):
            g, w = r['g'], r['g']['width']
            for i in range(3):
                R[i] += w * (r['Fxy'] * g['ex'][i] + r['Fyy'] * g['ey'][i] + r['Vyy'] * g['ez'][i])
            sumF['Fxy'] += w * r['Fxy']
            sumF['Fyy'] += w * r['Fyy']
            sumF['Vyy'] += w * r['Vyy']
            n += 1
        if n:
            tot[c] = dict(R=R, n=n, **sumF)
    length = sum(geo[e['id']]['width'] for e in plates)
    height = sum(geo[e['id']]['height'] for e in plates) / len(plates)
    t0 = collections.Counter(thick.values()).most_common(1)[0][0]

    wb = Workbook()
    ws = wb.active
    ws.title = '概要'
    info = [
        '耐震壁（板要素）の応力　midas の ANL（板要素の断面力）から抜き出し',
        '',
        f'モデル：{os.path.basename(model_path)}',
        f'ANL：{os.path.basename(anl_path)}（単位 {pl["units"][0]}, {pl["units"][1]}）',
        f'壁：板要素 {len(plates)} 枚（要素 {plates[0]["id"]}〜{plates[-1]["id"]}）、壁厚 t={t0:g} mm、壁の長さ（要素幅の合計）{length:.2f} m、高さ約 {height:.2f} m',
        '',
        '【シート】',
        '壁全体(ベース) … 壁下端の断面での N・Q・M（面内／面外）。耐震壁の N・M・Q の検定用（下端節点の値を足し合わせ）',
        '面内検討用 … 壁全体を 厚さ×弦長 の壁として検定するための NL・NE・ML・ME・QL・QE（短期・終局）',
        '面外検討用(1m柱) … 要素ごと・脚部／頂部の、幅1mの柱としての NL・NE・ML・ME・QL・QE（同じ内容を 壁面外_柱_短期／終局.xlsx に RC8 柱テンプレートの形で別出力）',
        '壁全体の合力 … 荷重ケースごとに、全要素の力を足した壁全体の鉛直力・水平力（耐震壁の N・Q の目安）',
        '最大値 … 区分（長期／短期／終局）ごとの、要素の最大・最小応力とそのケース',
        '要素別 … 全要素・全荷重ケースの要素中央の応力（Fxx・Fyy・Fxy＝面内力 kN/m、Mxx・Myy・Mxy＝曲げ kN·m/m、Vxx・Vyy＝面外せん断 kN/m）',
        '',
        '【読み方・注意】',
        '要素座標：x＝節点1→2（壁の水平方向）、y＝節点1→4（壁の鉛直方向、下向き）、z＝壁の面外方向。',
        'Fxx＝水平方向の軸力、Fyy＝鉛直方向の軸力（＋引張、−圧縮）、Fxy＝面内せん断力（単位幅当たり）。応力度 σ・τ は 力÷壁厚（N/mm²）。',
        '壁全体の合力は、要素中央の値×要素幅を足したもの（壁のベース付近の断面ではなく、各要素の中央の値）。符号は「壁がベースから受ける力」の全体座標（X・Y・Z）。',
        '区分（長期／短期／終局）は荷重ケース名から仮決め：長期（NΣDL など「+」の付かないケース）、「終局」を含む／QUX・QUY＝終局、それ以外＝短期。短期の各ケースは長期荷重を含む組合せ。',
        '耐震壁の検定そのものは未計算。',
    ]
    for r, s in enumerate(info, 1):
        ws.cell(r, 1, s)
    ws['A1'].font = Font(bold=True, size=13)
    ws.column_dimensions['A'].width = 130

    # 壁全体の合力
    w2 = wb.create_sheet('壁全体の合力')
    w2['A1'] = '壁全体の合力（荷重ケースごと）'
    w2['A1'].font = Font(bold=True, size=13)
    w2['A2'] = '壁がベースから受ける力の合計（全体座標、kN）。鉛直反力は上向き＋。面内せん断合計＝Σ(Fxy×幅)（壁に沿う水平力）。'
    heads = ['荷重ケース', '区分', '鉛直反力 Rz (kN)', '水平力 Rx (kN)', '水平力 Ry (kN)', '水平力合計 |RxRy| (kN)', '面内せん断 ΣFxy·b (kN)', '軸力 ΣFyy·b (kN)', '面外せん断 ΣVyy·b (kN)']
    header(w2, 4, heads, [26, 8, 16, 14, 14, 18, 18, 16, 18])
    for i, (c, v) in enumerate(tot.items(), 5):
        vals = [c, classes[c], v['R'][2], v['R'][0], v['R'][1], math.hypot(v['R'][0], v['R'][1]), v['Fxy'], v['Fyy'], v['Vyy']]
        for j, x in enumerate(vals):
            cell = w2.cell(i, 1 + j, round(x, 2) if isinstance(x, float) else x)
            cell.border = BORDER
            if j == 1:
                cell.fill = mx.PatternFill('solid', fgColor=CLASS_COLOR[x])
            elif j >= 2:
                cell.number_format = '0.0'
    w2.freeze_panes = 'C5'

    # 最大値
    w3 = wb.create_sheet('最大値')
    w3['A1'] = '要素の最大・最小応力（区分ごと）'
    w3['A1'].font = Font(bold=True, size=13)
    w3['A2'] = '各区分の全要素・全荷重ケースの中で、その項目が最大になる要素とケース。右はその要素・ケースの他の成分（同じ組）。'
    keys = ['Fxx', 'Fyy', 'Fxy', 'Mxx', 'Myy', 'Mxy', 'Vxx', 'Vyy']
    heads = ['区分', '項目', '要素', '通り', '荷重ケース'] + [k + (' (kN/m)' if k[0] in 'FV' else ' (kN·m/m)') for k in keys] + ['τ=|Fxy|/t (N/mm²)', 'σy=Fyy/t (N/mm²)']
    header(w3, 4, heads, [8, 22, 8, 22, 22] + [11] * 8 + [16, 16])
    picks = [('Fyy 最大（引張）', 'Fyy', 'max'), ('Fyy 最小（圧縮）', 'Fyy', 'min'), ('Fxy 絶対最大（面内せん断）', 'Fxy', 'abs'),
             ('Fxx 最大', 'Fxx', 'max'), ('Fxx 最小', 'Fxx', 'min'), ('Myy 絶対最大', 'Myy', 'abs'), ('Mxx 絶対最大', 'Mxx', 'abs'),
             ('Vyy 絶対最大', 'Vyy', 'abs')]
    r = 5
    for cls in ('長期', '短期', '終局'):
        sub = [x for x in rows if x['cls'] == cls]
        if not sub:
            continue
        for label, k, how in picks:
            f = {'max': lambda x: x[k], 'min': lambda x: -x[k], 'abs': lambda x: abs(x[k])}[how]
            best = max(sub, key=f)
            vals = [cls, label, best['e'], best['label'], best['case']] + [round(best[x], 2) for x in keys] + [round(best['tau'], 3), round(best['sy'], 3)]
            for j, x in enumerate(vals):
                cell = w3.cell(r, 1 + j, x)
                cell.border = BORDER
                if j == 0:
                    cell.fill = mx.PatternFill('solid', fgColor=CLASS_COLOR[cls])
                if 5 <= j < 5 + 8 and keys[j - 5] == k:
                    cell.fill = SUMF
                    cell.font = BOLD
            r += 1
    w3.freeze_panes = 'F5'

    # 要素別
    w4 = wb.create_sheet('要素別')
    w4['A1'] = '要素別の応力（要素中央、全荷重ケース）'
    w4['A1'].font = Font(bold=True, size=13)
    heads = ['要素', '通り', '壁厚 (mm)', '荷重ケース', '区分'] + [k + (' (kN/m)' if k[0] in 'FV' else ' (kN·m/m)') for k in keys] + ['τ=|Fxy|/t (N/mm²)', 'σx=Fxx/t (N/mm²)', 'σy=Fyy/t (N/mm²)']
    header(w4, 3, heads, [8, 22, 10, 22, 8] + [11] * 8 + [16, 16, 16])
    for i, x in enumerate(rows, 4):
        vals = [x['e'], x['label'], x['t'], x['case'], x['cls']] + [round(x[k], 2) for k in keys] + [round(x['tau'], 3), round(x['sx'], 3), round(x['sy'], 3)]
        for j, v in enumerate(vals):
            cell = w4.cell(i, 1 + j, v)
            cell.border = BORDER
            if j == 4:
                cell.fill = mx.PatternFill('solid', fgColor=CLASS_COLOR[v])
    w4.freeze_panes = 'F4'
    w4.auto_filter.ref = f'A3:{get_column_letter(len(heads))}{len(rows) + 3}'
    # ベース断面（下端節点 n3・n4）の合力と曲げモーメント
    nodes = model['nodes']
    base = []
    for e in plates:
        n3, n4 = e['nodes'][2], e['nodes'][3]
        p3, p4 = nodes[n3], nodes[n4]
        base.append(dict(e=e['id'], b=math.dist(p3, p4) / 1000, mid=tuple((p3[i] + p4[i]) / 2 for i in range(3)), n=(n3, n4), g=geo[e['id']]))
    L_arc = sum(x['b'] for x in base)
    ref = tuple(sum(x['b'] * x['mid'][i] for x in base) / L_arc for i in range(2)) + (base[0]['mid'][2],)
    first, last = base[0]['mid'], base[-1]['mid']
    ch = unit((last[0] - first[0], last[1] - first[1], 0.0))
    pe = (-ch[1], ch[0], 0.0)
    chord = math.hypot(last[0] - first[0], last[1] - first[1]) / 1000
    cross = lambda a, b: (a[1] * b[2] - a[2] * b[1], a[2] * b[0] - a[0] * b[2], a[0] * b[1] - a[1] * b[0])
    dot = lambda a, b: sum(x * y for x, y in zip(a, b))
    bres = collections.OrderedDict()
    for c in cases:
        R, Mf, Mp = [0.0] * 3, [0.0] * 3, [0.0] * 3
        ok = True
        for x in base:
            r = res[x['e']].get(c)
            if not r or any(k not in r for k in 'FMV') or any(n not in r[k] for k in 'FMV' for n in x['n']):
                ok = False
                break
            avg = lambda k, i, sc: sum(r[k][n][i] for n in x['n']) / 2 * sc
            fxy, fyy = avg('F', 2, kf), avg('F', 1, kf)
            vyy = avg('V', 1, kf)
            myy, mxy = avg('M', 1, km), avg('M', 2, km)
            g = x['g']
            f = tuple(fxy * g['ex'][i] + fyy * g['ey'][i] + vyy * g['ez'][i] for i in range(3))
            rr = tuple((x['mid'][i] - ref[i]) / 1000 for i in range(2)) + (0.0,)
            mm = cross(rr, f)
            for i in range(3):
                R[i] += x['b'] * f[i]
                Mf[i] += x['b'] * mm[i]
                Mp[i] += x['b'] * (-myy * g['ex'][i] + mxy * g['ey'][i])
        if ok:
            bres[c] = dict(R=R, Mf=Mf, Mp=Mp)
    w5 = wb.create_sheet('壁全体(ベース)', 2)
    w5['A1'] = '壁全体のベース断面の力とモーメント（下端節点の値）'
    w5['A1'].font = Font(bold=True, size=13)
    notes = [
        f'壁下端（節点3・4）の値を幅方向に足し合わせたもの。基準点は下端の壁長さ方向の重心（X={ref[0] / 1000:.3f} m, Y={ref[1] / 1000:.3f} m, Z={ref[2] / 1000:.3f} m）。',
        f'壁面に沿う方向 c＝始端→終端の水平方向（弦長 {chord:.2f} m）、直角方向 p。面内の転倒モーメント M面内＝p まわり、面外 M面外＝c まわり、ねじり＝Z まわり。',
        'M(力)＝膜力・せん断力の基準点まわりモーメント。M(板曲げ)＝板の曲げ Myy・ねじり Mxy の幅方向の合計（符号は −Myy·ex＋Mxy·ey、ユーザー確認済み）。合計＝両者の和。',
        'Q面内＝c 方向の水平力、Q面外＝p 方向の水平力、N＝鉛直反力（上向き＋、圧縮側）。単位 kN、kN·m。',
    ]
    for i, t in enumerate(notes, 2):
        w5.cell(i, 1, t)
    heads = ['荷重ケース', '区分', 'N 鉛直 (kN)', 'Q面内 (kN)', 'Q面外 (kN)', 'M面内 力 (kN·m)', 'M面内 板曲げ', 'M面内 合計', 'M面外 力', 'M面外 板曲げ', 'M面外 合計', 'ねじり Mz 合計', 'M面内/(Q面内·弦長)', 'τ平均 Q面内/(t·弦長) (N/mm²)']
    header(w5, 7, heads, [24, 8] + [14] * 12)
    for i, (c, v) in enumerate(bres.items(), 8):
        R, Mf, Mp = v['R'], v['Mf'], v['Mp']
        N = R[2]
        Qc, Qp = dot(R, ch), dot(R, pe)
        Min_f, Min_p = dot(Mf, pe), dot(Mp, pe)
        Mout_f, Mout_p = dot(Mf, ch), dot(Mp, ch)
        Mz = Mf[2] + Mp[2]
        ratio = (Min_f + Min_p) / (Qc * chord) if abs(Qc) > 1 else ''
        vals = [c, classes[c], N, Qc, Qp, Min_f, Min_p, Min_f + Min_p, Mout_f, Mout_p, Mout_f + Mout_p, Mz, ratio, abs(Qc) / (t0 * chord)]
        for j, x in enumerate(vals):
            cell = w5.cell(i, 1 + j, round(x, 3) if isinstance(x, float) else x)
            cell.border = BORDER
            if j == 1:
                cell.fill = mx.PatternFill('solid', fgColor=CLASS_COLOR[x])
            elif j >= 2:
                cell.number_format = '0.0' if j < 12 else '0.00'
            if j in (2, 3, 7, 10):
                cell.font = BOLD
    w5.freeze_panes = 'C8'
    r0 = 8 + len(bres) + 2
    w5.cell(r0, 1, '区分ごとの最大（絶対値）').font = BOLD
    header(w5, r0 + 1, ['区分', '項目', '荷重ケース', '値'])
    rr = r0 + 2
    for cls in ('長期', '短期', '終局'):
        cs = [c for c in bres if classes[c] == cls]
        if not cs:
            continue
        defs = [('M面内 合計の絶対最大 (kN·m)', lambda c: abs(dot(bres[c]['Mf'], pe) + dot(bres[c]['Mp'], pe))),
                ('Q面内の絶対最大 (kN)', lambda c: abs(dot(bres[c]['R'], ch))),
                ('N 最大（圧縮）(kN)', lambda c: bres[c]['R'][2]), ('N 最小 (kN)', lambda c: -bres[c]['R'][2])]
        for label, f in defs:
            best = max(cs, key=f)
            v = f(best) if not label.startswith('N 最小') else -f(best)
            for j, x in enumerate([cls, label, best, round(v, 1)]):
                cell = w5.cell(rr, 1 + j, x)
                cell.border = BORDER
                if j == 0:
                    cell.fill = mx.PatternFill('solid', fgColor=CLASS_COLOR[cls])
            rr += 1
    # 耐震壁チャートの基本入力用：スパン＝壁下端の両端節点間（柱芯間）の水平距離、階高＝壁の平均高さ
    bn = {n for x in base for n in x['n']}
    span = max(math.hypot(nodes[a][0] - nodes[b][0], nodes[a][1] - nodes[b][1]) for a in bn for b in bn) / 1000
    d = dict(res=res, cases=cases, classes=classes, plates=plates, geo=geo, gi=gi, kf=kf, km=km, bres=bres,
             ch=ch, pe=pe, chord=chord, t0=t0, height=height, out_dir=os.path.dirname(out_path), fc=fc, log=log,
             names=names or {k: f'壁面外_柱_{k}.xlsx' for k in ('短期', '終局')}, inplane=[])
    add_check_sheets(wb, d)
    wb.save(out_path)
    log(f'保存しました：{out_path}')
    return dict(span=span, height=height, t=t0, fc=fc, arc=L_arc, long=long, rows=d['inplane'], files=d.get('files', {}))


def nodeval(res, eid, case, nds, kf, km):
    """指定節点（下端 n3・n4、上端 n1・n2）の値の平均。1m 幅あたり。"""
    r = res[eid].get(case)
    if not r or any(k not in r for k in 'FMV') or any(n not in r[k] for k in 'FMV' for n in nds):
        return None
    av = lambda k, i, sc: sum(r[k][n][i] for n in nds) / len(nds) * sc
    return dict(Fyy=av('F', 1, kf), Myy=av('M', 1, km), Mxx=av('M', 0, km), Vyy=av('V', 1, kf))


def add_check_sheets(wb, d):
    """検定用：面内（壁全体の N・M・Q を L と E に分ける）と、面外（要素ごとの 1m 幅の柱）。"""
    from openpyxl import load_workbook
    import copy
    res, cases, classes, kf, km, bres = d['res'], d['cases'], d['classes'], d['kf'], d['km'], d['bres']
    ch, pe, chord, t0, log = d['ch'], d['pe'], d['chord'], d['t0'], d['log']
    import make_template as mt
    LONG = mt.find_long(cases)
    dot = lambda a, b: sum(x * y for x, y in zip(a, b))
    # ---- 面内（壁全体） ----
    ws = wb.create_sheet('面内検討用', 3)
    ws['A1'] = '耐震壁 面内の検討用（壁全体を 厚さ t × 弦長 の長方形断面の壁として見る）'
    ws['A1'].font = Font(bold=True, size=13)
    ws['A2'] = (f'断面：t={t0:g} mm、弦長 l={chord * 1000:,.0f} mm、高さ h={d["height"] * 1000:,.0f} mm。L＝長期 {LONG}、E＝候補ケースと長期の差の絶対値。'
                'N は圧縮＋（kN）、M は M面内 合計（kN·m）、Q は面内せん断（kN）。')
    ws['A3'] = '曲げの最大ケースとせん断の最大ケースが違うので、区分ごとに別の行にしています。採用ケースの L±E で N の最小・最大も見られます。'
    if not LONG or LONG not in bres:
        ws['A5'] = '長期ケースの結果がないため作れませんでした。'
        return
    gv = lambda c: (bres[c]['R'][2], dot(bres[c]['R'], ch), dot([a + b for a, b in zip(bres[c]['Mf'], bres[c]['Mp'])], pe))
    Ln, Lq, Lm = gv(LONG)
    heads = ['区分', '何で選んだか', '採用ケース', 'NL (kN)', 'NE (kN)', 'N 最大=NL+NE', 'N 最小=NL−NE', 'ML (kN·m)', 'ME (kN·m)', 'QL (kN)', 'QE (kN)',
             't (mm)', '弦長 l (mm)', 'τL+τE 平均 (N/mm²)']
    header(ws, 5, heads, [8, 22, 22] + [13] * 11)
    r = 6
    for kind in ('短期', '終局'):
        cs = [c for c in bres if classes[c] == kind]
        if not cs:
            continue
        dl = {c: tuple(abs(a - b) for a, b in zip(gv(c), (Ln, Lq, Lm))) for c in cs}   # (dN, dQ, dM)
        for why, key in (('曲げ ΔM 最大', lambda c: dl[c][2]), ('せん断 ΔQ 最大', lambda c: dl[c][1]), ('軸力 ΔN 最大', lambda c: dl[c][0])):
            c = max(cs, key=key)
            dn, dq, dm = dl[c]
            vals = [kind, why, c, Ln, dn, Ln + dn, Ln - dn, abs(Lm), dm, abs(Lq), dq, t0, chord * 1000, (abs(Lq) + dq) / (t0 * chord)]
            d['inplane'].append(dict(kind=kind, why=why, case=c, NL=Ln, NE=dn, ML=abs(Lm), ME=dm, QL=abs(Lq), QE=dq))
            for j, x in enumerate(vals):
                cell = ws.cell(r, 1 + j, round(x, 3) if isinstance(x, float) else x)
                cell.border = BORDER
                if j == 0:
                    cell.fill = mx.PatternFill('solid', fgColor=CLASS_COLOR[kind])
                elif j >= 3:
                    cell.number_format = '0.0' if j < 13 else '0.000'
            r += 1
    ws.freeze_panes = 'D6'
    # ---- 面外（1m 幅の柱） ----
    wo = wb.create_sheet('面外検討用(1m柱)', 4)
    wo['A1'] = '耐震壁 面外の検討用（要素ごとに、幅 1,000 mm × 厚さ t の柱として見る）'
    wo['A1'].font = Font(bold=True, size=13)
    wo['A2'] = ('脚部＝要素下端の節点（3・4）の平均、頂部＝上端の節点（1・2）の平均。N は圧縮＋（kN/m）＝1m 幅の kN、M は kN·m/m、Q は kN/m。'
                'L＝長期、E＝採用ケースと長期の差の絶対値。採用ケースは、その要素・位置で縦筋方向の曲げ Myy の変化が最大のケース。')
    wo['A3'] = '縦筋は Myy、横筋は Mxx（同じ採用ケースでの値）。Myy・Mxx のどちらが縦筋方向かは midas の定義を確認すること。'
    heads = ['区分', '要素', '通り', '位置', '採用ケース', 'NL', 'NE', 'MyyL', 'MyyE', 'VyyL', 'VyyE', 'MxxL', 'MxxE', 't (mm)']
    header(wo, 5, heads, [8, 8, 22, 8, 22] + [11] * 9)
    tpl = {'短期': [], '終局': []}
    r = 6
    for kind in ('短期', '終局'):
        cs = [c for c in cases if classes[c] == kind]
        for e in d['plates']:
            g = d['geo'][e['id']]
            label = d['gi'].label(g['center'], g['center'][2]) if d['gi'].available else ''
            n = e['nodes']
            for pos, nds in (('脚部', (n[2], n[3])), ('頂部', (n[0], n[1]))):
                lg = nodeval(res, e['id'], LONG, nds, kf, km)
                vs = {c: nodeval(res, e['id'], c, nds, kf, km) for c in cs}
                vs = {c: v for c, v in vs.items() if v}
                if not lg or not vs:
                    continue
                c = max(vs, key=lambda x: (abs(vs[x]['Myy'] - lg['Myy']), abs(vs[x]['Fyy'] - lg['Fyy'])))
                v = vs[c]
                vals = [kind, e['id'], label, pos, c, -lg['Fyy'], abs(v['Fyy'] - lg['Fyy']), abs(lg['Myy']), abs(v['Myy'] - lg['Myy']),
                        abs(lg['Vyy']), abs(v['Vyy'] - lg['Vyy']), abs(lg['Mxx']), abs(v['Mxx'] - lg['Mxx']), t0]
                for j, x in enumerate(vals):
                    cell = wo.cell(r, 1 + j, round(x, 3) if isinstance(x, float) else x)
                    cell.border = BORDER
                    if j == 0:
                        cell.fill = mx.PatternFill('solid', fgColor=CLASS_COLOR[kind])
                    elif j >= 5 and j < 13:
                        cell.number_format = '0.0'
                r += 1
                # 柱テンプレート：Dx＝壁厚（曲げの方向）、Dy＝1000。X 方向の曲げだけ入れる
                tpl[kind].append([f'W{e["id"]} {label} {pos}'.replace('  ', ' '), t0, 1000.0, -lg['Fyy'], abs(v['Fyy'] - lg['Fyy']), 0.0,
                                  abs(lg['Myy']), 0.0, abs(v['Myy'] - lg['Myy']), 0.0, abs(lg['Vyy']), 0.0, abs(v['Vyy'] - lg['Vyy']), 0.0, d['fc']])
    wo.freeze_panes = 'F6'
    wo.auto_filter.ref = f'A5:{get_column_letter(len(heads))}{r - 1}'
    # 柱テンプレート（RC8）の形で
    import make_template as mt
    for kind, rows in tpl.items():
        if rows:
            out = os.path.normpath(os.path.join(d['out_dir'], d['names'][kind]))
            mt.fill(os.path.join(HERE, 'template', 'template_column.xlsx'), rows, out)
            d.setdefault('files', {})[kind] = out
            log(f'面外の柱テンプレート {len(rows)} 行 → {out}')


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('model')
    ap.add_argument('anl')
    ap.add_argument('-o', '--out', default=os.path.join(HERE, '出力'))
    ap.add_argument('--fc', type=float, default=None, help='省略時は MGTX の材料名（Fc27 など）から')
    a = ap.parse_args()
    os.makedirs(a.out, exist_ok=True)
    build(a.model, a.anl, os.path.join(a.out, '壁の応力.xlsx'), fc=a.fc)


if __name__ == '__main__':
    main()
