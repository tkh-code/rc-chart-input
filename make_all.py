"""MGTX と ANL から、4-1～4-3（と 5-4 の終局）の RC チャートに入れる Excel をまとめて作る。

py -3 make_all.py model.mgtx result.anl [-o 出力フォルダ] [--fc 27] [--swap-xy]

出力フォルダ
  4-1_RC柱_短期.xlsx           … RC柱チャート（柱テンプレート）
  4-2_RC梁_短期.xlsx           … RC梁チャート（梁テンプレート）
  4-3_耐震壁_面外_短期.xlsx    … RC柱チャート（壁を 1m 幅の柱とみなした柱テンプレート）
  4-3_耐震壁_面内_チャート入力.xlsx … 耐震壁チャートの「基本入力」に打ち込む値
  5-4_終局/                    … 同じものの終局（QUX・QUY）版
  参考/                        … 採用ケースの根拠、壁の応力の詳細
"""
import argparse
import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
sys.dont_write_bytecode = True
import make_excel as mx  # noqa: E402
import make_template as mt  # noqa: E402
import make_wall as mw  # noqa: E402

INPUT = PatternFill('solid', fgColor='FFF2CC')   # 手で入れる欄


def column_size(model_path):
    """RC柱の断面（mm）。壁の両端の柱としてチャートに入れる。"""
    model = mx.read_mgt(model_path)
    _, dims = mt.read_extra(model_path)
    secs = sorted({(H, B) for i, n in model['sections'].items() if n.upper().startswith('RC柱') for H, B in [dims.get(i, (0.0, 0.0))]})
    return secs[0] if secs else (0.0, 0.0)


def wall_chart_input(path, w, col):
    wb = Workbook()
    ws = wb.active
    ws.title = '耐震壁チャート入力'
    ws['A1'] = '4-3 耐震壁（面内）　耐震壁チャートの「基本入力」に入れる値'
    ws['A1'].font = Font(bold=True, size=13)
    ws['A2'] = '黄色の欄はモデルに無いので、4-3 の断面諸元・構造図から手で入れる。せん断力は下の表（midas の壁全体の合力）から。'
    for c, wd in zip('ABCD', (26, 14, 10, 70)):
        ws.column_dimensions[c].width = wd
    for j, h in enumerate(['項目', '値', '単位', '出どころ・注意']):
        c = ws.cell(4, 1 + j, h)
        c.font, c.fill, c.border = mx.BOLD, mx.HEAD, mx.BORDER
    # 根拠の表（下）を先に位置決めして、せん断力はそのセルを参照する
    t0 = 26
    pick = {r['kind']: i for i, r in enumerate(w['rows']) if r['why'].startswith('せん断')}
    ref = lambda kind, col_: f'{col_}{t0 + 1 + pick[kind]}' if kind in pick else None
    items = [
        ('スパン長', round(w['span'], 3), 'm', '壁の両端の柱芯間の水平距離（壁下端の両端節点間、midas モデル）'),
        ('階高', round(w['height'], 3), 'm', '壁の平均高さ（板要素の高さの平均）。基礎梁上端～上部の梁上端と違う場合は直す'),
        ('壁厚', w['t'], 'mm', 'midas の板厚'),
        ('そで壁長（左・右）', 0, 'm', 'そで壁なし'),
        ('柱 B（左・右）', col[1], 'mm', 'midas の RC柱断面'),
        ('柱 D（左・右）', col[0], 'mm', 'midas の RC柱断面'),
        ('開口', 'なし', '', 'モデルに開口なし（ある場合は lop・hop を入れる）'),
        ('横筋（径・ピッチ）', None, '', '4-3 断面諸元による（例：D13@100）'),
        ('縦筋（径・ピッチ）', None, '', '4-3 断面諸元による（縦横同配筋ならチェックのみ）'),
        ('シングル配筋', None, '', 'ダブルならチェックしない'),
        ('柱の HOOP 筋（本数・径・ピッチ）', None, '', '4-1 RC柱の断面諸元による'),
        ('短期せん断力', f'={ref("短期", "H")}+{ref("短期", "I")}' if '短期' in pick else None, 'kN', 'QL＋QE（下の表の短期・せん断 ΔQ 最大の行）'),
        ('応力割増率', 1.5, '', 'チャートの初期値。4-3 本文の QS＝QL＋QE に合わせるなら 1.0 にする'),
        ('コンクリート Fc', w['fc'], 'N/mm²', 'midas の材料名から'),
    ]
    for i, (k, v, u, note) in enumerate(items, 5):
        for j, x in enumerate((k, v, u, note)):
            c = ws.cell(i, 1 + j, x)
            c.border = mx.BORDER
        if v is None:
            ws.cell(i, 2).fill = INPUT
        ws.cell(i, 2).alignment = Alignment(horizontal='right')
        if k == '短期せん断力':
            ws.cell(i, 1).font = ws.cell(i, 2).font = mx.BOLD
            ws.cell(i, 2).number_format = '0.0'
    r = 5 + len(items) + 1
    if '終局' in pick:
        ws.cell(r, 1, '（5-4 用）終局時せん断力').font = mx.BOLD
        c = ws.cell(r, 2, f'={ref("終局", "H")}+{ref("終局", "I")}')
        c.font, c.number_format = mx.BOLD, '0.0'
        ws.cell(r, 3, 'kN')
        ws.cell(r, 4, 'QL＋QE（下の表の終局・せん断 ΔQ 最大の行）。5-4 で使う')
    ws.cell(t0 - 1, 1, '壁全体の N・M・Q（壁下端の断面、midas の板要素から）').font = mx.BOLD
    heads = ['区分', '何で選んだか', '採用ケース', 'NL (kN)', 'NE (kN)', 'ML (kN·m)', 'ME (kN·m)', 'QL (kN)', 'QE (kN)']
    for j, h in enumerate(heads):
        c = ws.cell(t0, 1 + j, h)
        c.font, c.fill, c.border = mx.BOLD, mx.HEAD, mx.BORDER
    for i, x in enumerate(w['rows'], t0 + 1):
        for j, v in enumerate([x['kind'], x['why'], x['case'], x['NL'], x['NE'], x['ML'], x['ME'], x['QL'], x['QE']]):
            c = ws.cell(i, 1 + j, round(v, 1) if isinstance(v, float) else v)
            c.border = mx.BORDER
            if j >= 3:
                c.number_format = '0.0'
        ws.cell(i, 1).fill = PatternFill('solid', fgColor=mx.CLASS_COLOR[x['kind']])
    ws.cell(t0 + len(w['rows']) + 2, 1, f'L＝長期 {w["long"]}、E＝採用ケースと長期の差の絶対値。N は圧縮＋。M は面内の転倒モーメント、Q は壁に沿う水平力。')
    for c in 'EFGHI':
        ws.column_dimensions[c].width = 12
    wb.save(path)


def run(model_path, anl_path, out, fc=None, swap=False, log=print):
    if fc is None:
        fc = mt.read_fc(model_path) or 24.0
    sub_u, sub_r = os.path.join(out, '5-4_終局'), os.path.join(out, '参考')
    for p in (out, sub_u, sub_r):
        os.makedirs(p, exist_ok=True)
    made = []
    for kind, folder, pre in (('短期', out, '4'), ('終局', sub_u, '5-4')):
        log(f'--- 柱・梁（{kind}）')
        names = {'柱': f'{pre}-1_RC柱_{kind}.xlsx' if pre == '4' else f'5-4_RC柱_{kind}.xlsx',
                 '梁': f'{pre}-2_RC梁_{kind}.xlsx' if pre == '4' else f'5-4_RC梁_{kind}.xlsx',
                 '根拠': f'根拠_柱梁_{kind}.xlsx'}
        p = mt.run(model_path, anl_path, folder, kind, fc, swap, names, log)
        os.replace(p['根拠'], os.path.join(sub_r, names['根拠']))
        made += [p['柱'], p['梁']]
    log('--- 耐震壁')
    w = mw.build(model_path, anl_path, os.path.join(sub_r, '壁の応力.xlsx'), log=log, fc=fc,
                 names={'短期': os.path.join('..', '4-3_耐震壁_面外_短期.xlsx'), '終局': os.path.join('..', '5-4_終局', '5-4_耐震壁_面外_終局.xlsx')})
    made += list(w['files'].values())
    col = column_size(model_path)
    p = os.path.join(out, '4-3_耐震壁_面内_チャート入力.xlsx')
    wall_chart_input(p, w, col)
    made.append(p)
    log('')
    log('できたファイル：')
    for m in made:
        log('  ' + os.path.relpath(os.path.normpath(m), out))
    return made


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument('model')
    ap.add_argument('anl')
    ap.add_argument('-o', '--out', default=os.path.join(HERE, '出力'))
    ap.add_argument('--fc', type=float, default=None, help='省略時は MGTX の材料名（Fc27 など）から')
    ap.add_argument('--swap-xy', action='store_true')
    a = ap.parse_args()
    try:
        run(a.model, a.anl, a.out, a.fc, a.swap_xy)
    except ValueError as e:
        sys.exit(str(e))


if __name__ == '__main__':
    main()
