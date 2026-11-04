"""midas の MGTX と ANL から、RC柱・RC梁の設計ツール用テンプレート（template_column.xlsx / template_beam.xlsx）を埋める。

py -3 make_template.py model.mgtx result.anl [-o 出力フォルダ] [--kind 短期|終局] [--fc 27] [--swap-xy]
（4-1～4-3 をまとめて出すときは make_all.py）

出力：柱_<kind>.xlsx / 梁_<kind>.xlsx（テンプレートの 2 行目以降に 1 行 1 識別名）と、採用した荷重ケースの根拠 根拠_<kind>.xlsx。
"""
import argparse
import collections
import copy
import math
import os
import re
import sys

from openpyxl import Workbook, load_workbook

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import make_excel as mx  # noqa: E402

LONG = 'NΣDL'   # run() で ANL のケース名から決め直す


def find_long(cases):
    """長期のケース名。NΣDL があればそれ。無ければ「+」を含まず、他のケースが「その名前+…」になっているもの（NTL など）。"""
    if 'NΣDL' in cases:
        return 'NΣDL'
    return max((c for c in cases if '+' not in c), key=lambda c: sum(x.startswith(c + '+') for x in cases), default=None)


def read_extra(path):
    """要素のベータ角と、四角断面（SB）の寸法 H・B を読む。make_excel の読み込みでは拾っていない分。"""
    text = open(path, 'rb').read().decode('cp932', 'replace')
    key, beta, dims, mm = None, {}, {}, 1.0
    for line in text.split('\n'):
        m = re.match(r'\s*\*([A-Z0-9_-]+)', line)
        if m:
            key = m.group(1).upper()
            continue
        p = [x.strip() for x in line.split(',')]
        if key == 'UNIT' and len(p) >= 2 and not line.lstrip().startswith(';'):
            mm = {'MM': 1.0, 'CM': 10.0, 'M': 1000.0}.get(p[1].upper(), 1.0)
        elif key == 'ELEMENT' and len(p) >= 7 and p[0].isdigit() and p[1].upper() == 'BEAM':
            try:
                beta[int(p[0])] = float(p[6])
            except ValueError:
                beta[int(p[0])] = 0.0
        elif key == 'SECTION' and len(p) >= 16 and p[0].isdigit() and p[12].upper() == 'SB':
            dims[int(p[0])] = (float(p[14]) * mm, float(p[15]) * mm)   # (H, B) mm。H は要素 z 方向、B は y 方向
    return beta, dims


def candidate_cases(cases, kind):
    """E の候補。短期は 長期以外の（終局でない）全ケース＝風・雪・地震。終局は「終局」の付いたケースと QUX・QUY。"""
    ult = lambda c: '終局' in c or re.search(r'QU[XY]$', c) is not None
    return [c for c in cases if c != LONG and ult(c) == (kind == '終局')]


def vec(res, eid, case, pt):
    return res['results'].get(eid, {}).get(case, {}).get(pt)


def beam_rows(model, res, chains, dims, gi, cands, fc, log):
    rows, notes = [], []
    for k, line in enumerate(chains, 1):
        e = line.elems[0]
        H, B = dims.get(e['sec'], (0.0, 0.0))
        L = line.total
        n0, n1 = line.node_ids[0], line.node_ids[-1]
        mg = gi.member(model['nodes'][n0], model['nodes'][n1], line.min_z(model)) if gi.available else ''
        name = f'B{k:02d} {mg}'.strip()
        ci_q, ci_m = 2, 4   # Qz, My

        def mom(case):
            return [line.value_at(res['results'], case, ci_m, s) for s in (0, L / 2, L)]

        def shr(case):
            return [line.value_at(res['results'], case, ci_q, s) for s in (0, L / 2, L)]
        ml, ql = mom(LONG), shr(LONG)
        if None in ml or None in ql:
            log(f'  {name}: 長期の結果なし → 飛ばしました')
            continue
        best = None
        for case in cands:
            m, q = mom(case), shr(case)
            if None in m or None in q:
                continue
            dm = [abs(a - b) for a, b in zip(m, ml)]
            dq = [abs(a - b) for a, b in zip(q, ql)]
            score = max(dm[0], dm[2]) + max(dq[0], dq[2])
            if best is None or score > best[0]:
                best = (score, case, dm, dq)
        if best is None:
            log(f'  {name}: 地震ケースの結果なし → 飛ばしました')
            continue
        _, case, dm, dq = best
        rows.append([name, B, H, ml[0], ml[1], ml[2], dm[0], dm[2], abs(ql[0]), abs(ql[2]), max(dq[0], dq[2]), fc])
        notes.append((name, f'要素{line.elems[0]["id"]}〜{line.elems[-1]["id"]}', f'{B:g}×{H:g}', case))
    return rows, notes


def column_rows(model, res, columns, dims, beta, gi, cands, fc, swap, log):
    rows, notes = [], []
    nodes = model['nodes']
    for k, line in enumerate(columns, 1):
        e = line.elems[0]
        H, B = dims.get(e['sec'], (0.0, 0.0))
        bt = beta.get(e['id'], 0.0) % 180
        if abs(bt) > 1e-6 and abs(bt - 90) > 1e-6:
            log(f'  要素{e["id"]}: ベータ角 {bt:g}° は X・Y への分解に未対応 → 0° とみなしました')
            bt = 0.0
        # ベータ 0°の鉛直材は 要素 z' ∥ 全体 X、要素 y' ∥ 全体 Y。90°で入れ替わる。
        Dx, Dy = (H, B) if bt == 0 else (B, H)
        # 要素 My（y' まわり）＝ X 方向に倒れる曲げ、Qz ＝ X 方向せん断。Mz・Qy が Y 方向。
        ix = (4, 2) if bt == 0 else (5, 1)   # (X 方向の曲げ成分, X 方向のせん断成分)
        iy = (5, 1) if bt == 0 else (4, 2)
        if swap:
            ix, iy = iy, ix
            Dx, Dy = Dy, Dx
        a, b = e['nodes'][0], e['nodes'][1]
        top_is_i = nodes[a][2] >= nodes[b][2]
        for lab, pt, nd in (('柱頭', 'I' if top_is_i else 'J', a if top_is_i else b), ('柱脚', 'J' if top_is_i else 'I', b if top_is_i else a)):
            mg = gi.label(nodes[nd], line.min_z(model)) if gi.available else ''
            name = f'C{k:02d} {mg} {lab}'.replace('  ', ' ').strip()
            base = vec(res, e['id'], LONG, pt)
            vs = {c: vec(res, e['id'], c, pt) for c in cands}
            vs = {c: v for c, v in vs.items() if v is not None}
            if base is None or not vs:
                log(f'  {name}: 結果なし → 飛ばしました')
                continue
            # X 方向の E は X 方向の曲げ変化が最大のケース、Y 方向は Y 方向の曲げ変化が最大のケース（同点なら軸力変化）
            case_x = max(vs, key=lambda c: (abs(vs[c][ix[0]] - base[ix[0]]), abs(vs[c][0] - base[0])))
            case_y = max(vs, key=lambda c: (abs(vs[c][iy[0]] - base[iy[0]]), abs(vs[c][0] - base[0])))
            dx = [abs(p - q) for p, q in zip(vs[case_x], base)]
            dy = [abs(p - q) for p, q in zip(vs[case_y], base)]
            rows.append([name, Dx, Dy, -base[0], dx[0], dy[0], abs(base[ix[0]]), abs(base[iy[0]]), dx[ix[0]], dy[iy[0]],
                         abs(base[ix[1]]), abs(base[iy[1]]), dx[ix[1]], dy[iy[1]], fc])
            notes.append((name, f'要素{e["id"]}', f'{Dx:g}×{Dy:g}', case_x, case_y))
    return rows, notes


def fill(template, rows, out):
    wb = load_workbook(template)
    ws = wb.worksheets[0]
    proto = [copy.copy(c._style) for c in ws[2]]
    for r, vals in enumerate(rows, 2):
        for j, v in enumerate(vals):
            c = ws.cell(r, 1 + j, round(v, 3) if isinstance(v, float) else v)
            c._style = copy.copy(proto[j])
    wb.save(out)


def write_notes(path, kind, cands, beam_notes, col_notes):
    wb = Workbook()
    ws = wb.active
    ws.title = '根拠'
    ws.append([f'E ＝ {kind}の候補ケース {" / ".join(cands)} のうち最大のものと、長期 {LONG} との差（絶対値）'])
    ws.append([])
    ws.append(['梁'])
    ws.append(['識別名', '要素', 'b×D (mm)', '採用ケース'])
    for n in beam_notes:
        ws.append(list(n))
    ws.append([])
    ws.append(['柱'])
    ws.append(['識別名', '要素', 'Dx×Dy (mm)', 'X方向に採用', 'Y方向に採用'])
    for n in col_notes:
        ws.append(list(n))
    ws.column_dimensions['A'].width = 40
    ws.column_dimensions['E'].width = 28
    wb.save(path)


def read_fc(path):
    """RC のコンクリート強度。MGTX の材料名（Fc27 など）から読む。見つからなければ None。"""
    text = open(path, 'rb').read().decode('cp932', 'replace')
    m = re.search(r'^\s*\d+\s*,\s*CONC\s*,\s*Fc\s*(\d+(?:\.\d+)?)', text, re.M | re.I)
    return float(m.group(1)) if m else None


def run(model_path, anl_path, out, kind='短期', fc=None, swap=False, names=None, log=print):
    """柱・梁のテンプレートと根拠を出力し、出力したファイルのパスを返す。names＝{'梁','柱','根拠'} のファイル名。"""
    names = names or {'梁': f'梁_{kind}.xlsx', '柱': f'柱_{kind}.xlsx', '根拠': f'根拠_{kind}.xlsx'}
    model = mx.read_mgt(model_path)
    beta, dims = read_extra(model_path)
    if fc is None:
        fc = read_fc(model_path) or 24.0
    res = mx.read_anl(anl_path)
    cases = res['cases']
    global LONG
    LONG = find_long(cases)
    if not LONG:
        raise ValueError(f'長期ケース（NΣDL など「+」の付かないケース）が ANL にありません：{cases}')
    log(f'長期ケース：{LONG}')
    cands = candidate_cases(cases, kind)
    log(f'E の候補ケース（{kind}）：{cands}')
    rc_secs = {i for i, n in model['sections'].items() if n.upper().startswith('RC')}
    rc = [e for e in model['elements'].values() if e['type'] == 'BEAM' and e['sec'] in rc_secs]

    def vertical(e):
        p, q = (model['nodes'][n] for n in e['nodes'][:2])
        return abs(q[2] - p[2]) > 0.7 * math.dist(p, q)
    columns = [mx.Line([e], model) for e in sorted((e for e in rc if vertical(e)), key=lambda x: x['id'])]
    chains = mx.build_chains([e for e in rc if not vertical(e)], model)
    gi = mx.GridIndex(model['grids'])
    log(f'RC柱 {len(columns)} 本・RC梁 {len(chains)} 部材、Fc{fc:g}')
    brow, bnote = beam_rows(model, res, chains, dims, gi, cands, fc, log)
    crow, cnote = column_rows(model, res, columns, dims, beta, gi, cands, fc, swap, log)
    os.makedirs(out, exist_ok=True)
    paths = {k: os.path.join(out, v) for k, v in names.items()}
    fill(os.path.join(HERE, 'template', 'template_beam.xlsx'), brow, paths['梁'])
    fill(os.path.join(HERE, 'template', 'template_column.xlsx'), crow, paths['柱'])
    write_notes(paths['根拠'], kind, cands, bnote, cnote)
    log(f'梁 {len(brow)} 行・柱 {len(crow)} 行')
    return paths


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('model')
    ap.add_argument('anl')
    ap.add_argument('-o', '--out', default=os.path.join(HERE, '出力'))
    ap.add_argument('--kind', choices=['短期', '終局'], default='短期', help='E（地震）に使うケース')
    ap.add_argument('--fc', type=float, default=None, help='省略時は MGTX の材料名（Fc27 など）から')
    ap.add_argument('--swap-xy', action='store_true', help='Mx/Qx の向きの取り方を入れ替える')
    a = ap.parse_args()
    try:
        run(a.model, a.anl, a.out, a.kind, a.fc, a.swap_xy)
    except ValueError as e:
        sys.exit(str(e))


if __name__ == '__main__':
    main()
