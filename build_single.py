"""index.html・worker.js・Python 本体・テンプレートを 1 つの HTML にまとめる。py -3 build_single.py → RCチャート入力.html
（ダブルクリックでブラウザで開ける。初回だけ計算エンジン Pyodide をネットから読み込む）"""
import base64, json, os
HERE = os.path.dirname(os.path.abspath(__file__))
rd = lambda p, m='r': open(os.path.join(HERE, p), m, **({} if 'b' in m else {'encoding': 'utf-8'})).read()
src = ['make_excel.py', 'make_template.py', 'make_wall.py', 'make_all.py']
tpl = ['template/template_beam.xlsx', 'template/template_column.xlsx']
bundle = {'worker': rd('worker.js'), 'sources': {f: rd(f) for f in src},
          'templates': {f: base64.b64encode(rd(f, 'rb')).decode() for f in tpl}}
js = 'window.__BUNDLE=' + json.dumps(bundle, ensure_ascii=False).replace(r'</', r'<\/') + ';'
html = rd('index.html').replace('<script>\nconst $ =', '<script>' + js + '</script>\n<script>\nconst $ =', 1)
assert js in html
open(os.path.join(HERE, 'RCチャート入力.html'), 'w', encoding='utf-8').write(html)
print('RCチャート入力.html', len(html) // 1024, 'KB')
