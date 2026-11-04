# 73_日吉津_4章ブラウザ版（RC チャート入力 4-1～4-3 のブラウザ版）

MGTX と ANL を 2 つまとめてドロップすると、計算書 4-1 RC柱・4-2 RC梁・4-3 耐震壁（と 5-4 終局）の RC チャートに入れる Excel を作る。bat 不要。

- 仕組み：`index.html`＋`worker.js` が Pyodide（ブラウザ内 Python）で `make_all.py` をそのまま実行。ファイルは外へ送られない。
- Python 本体（`make_excel.py`・`make_template.py`・`make_wall.py`・`make_all.py`）と `template/` は、`50_RC8テンプレート入力` の原本のコピー。
  **原本（50）を直したらここへ上書きコピーする。**
- ローカルで試す：`py -m http.server 8767` してから http://localhost:8767/ を開く（ファイルを直接開くと動かない）。
- 出力は ZIP（`4-1…`、`4-2…`、`4-3…`、`5-4_終局/`、`参考/`）と、1 ファイルずつのダウンロード。Fc 空欄は MGTX の材料名から自動、X・Y 入れ替えは PC 版の `--swap-xy` と同じ。
- 確認：日吉津 グランルーフの MGTX（9 MB）＋ANL（33 MB）で 19 秒、10 ファイル生成（2026-11-04）。
- 1 ファイル版：`py -3 build_single.py` で `RCチャート入力.html` を作る（ダブルクリックで開ける。初回だけ Pyodide をネットから読む）。G の 02_計算書/midas→エクセル/RCチャート入力_4章.html にコピー済み。原本を直したら再ビルドして再コピー。
- 公開：https://tkh-code.github.io/rc-chart-input/ （GitHub: tkh-code/rc-chart-input、push で更新。公開版は index.html＋worker.js の複数ファイル版）
