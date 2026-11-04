/* RC チャート入力（4-1～4-3）ワーカー
 * Pyodide（ブラウザ内 Python）上で make_all.py の run() をそのまま実行する。
 * ファイルはワーカー内の仮想ファイルシステムだけで処理され、外部には送信されない。
 */
importScripts("https://cdn.jsdelivr.net/pyodide/v0.26.4/full/pyodide.js");

let pyodide = null;
const HOME = "/home/pyodide";
const SOURCES = ["make_excel.py", "make_template.py", "make_wall.py", "make_all.py"];
const TEMPLATES = ["template/template_beam.xlsx", "template/template_column.xlsx"];

async function init(bundle) {
  pyodide = await loadPyodide();
  await pyodide.loadPackage("micropip");
  const micropip = pyodide.pyimport("micropip");
  await micropip.install("openpyxl==3.1.5");
  pyodide.FS.mkdirTree(HOME + "/template");
  if (bundle) {   // 1 ファイル版：ソースとテンプレートは HTML に埋め込み済み
    for (const f of SOURCES) pyodide.FS.writeFile(HOME + "/" + f, bundle.sources[f]);
    for (const f of TEMPLATES) {
      const bin = atob(bundle.templates[f]);
      const u8 = new Uint8Array(bin.length);
      for (let i = 0; i < bin.length; i++) u8[i] = bin.charCodeAt(i);
      pyodide.FS.writeFile(HOME + "/" + f, u8);
    }
  } else {
    const v = "?v=" + Date.now();
    for (const f of SOURCES) {
      const r = await fetch(f + v);
      if (!r.ok) throw new Error(f + " の取得に失敗しました");
      pyodide.FS.writeFile(HOME + "/" + f, await r.text());
    }
    for (const f of TEMPLATES) {
      const r = await fetch(f + v);
      if (!r.ok) throw new Error(f + " の取得に失敗しました");
      pyodide.FS.writeFile(HOME + "/" + f, new Uint8Array(await r.arrayBuffer()));
    }
  }
  pyodide.runPython("import sys; sys.path.insert(0, '" + HOME + "'); sys.dont_write_bytecode = True; import make_all");
  postMessage({ type: "ready" });
}

function convert(msg) {
  pyodide.FS.writeFile("/tmp/model" + msg.modelExt, new Uint8Array(msg.model));
  pyodide.FS.writeFile("/tmp/result.anl", new Uint8Array(msg.anl));
  pyodide.globals.set("js_log", (s) => postMessage({ type: "log", text: String(s) }));
  pyodide.globals.set("model_path", "/tmp/model" + msg.modelExt);
  pyodide.globals.set("fc_value", msg.fc);   // null なら MGTX から読む
  pyodide.globals.set("swap_xy", !!msg.swap);
  try {
    pyodide.runPython(`
import os, shutil, zipfile, make_all
shutil.rmtree("/tmp/out", ignore_errors=True)
made = make_all.run(model_path, "/tmp/result.anl", "/tmp/out", fc_value, swap_xy, log=js_log)
_files = []
for root, _, names in os.walk("/tmp/out"):
    for n in sorted(names):
        p = os.path.join(root, n)
        _files.append(os.path.relpath(p, "/tmp/out").replace(os.sep, "/"))
_files.sort()
with zipfile.ZipFile("/tmp/out.zip", "w", zipfile.ZIP_DEFLATED) as z:
    for f in _files:
        z.write(os.path.join("/tmp/out", f), f)
`);
    const names = pyodide.globals.get("_files").toJs();
    const files = [];
    const transfer = [];
    for (const n of names) {
      const b = pyodide.FS.readFile("/tmp/out/" + n);
      files.push({ path: n, bytes: b.buffer });
      transfer.push(b.buffer);
    }
    const z = pyodide.FS.readFile("/tmp/out.zip");
    transfer.push(z.buffer);
    postMessage({ type: "done", files, zip: z.buffer }, transfer);
  } finally {
    try { pyodide.runPython("import shutil; shutil.rmtree('/tmp/out', ignore_errors=True)"); } catch (e) { /* noop */ }
    for (const p of ["/tmp/model" + msg.modelExt, "/tmp/result.anl", "/tmp/out.zip"]) {
      try { pyodide.FS.unlink(p); } catch (e) { /* noop */ }
    }
  }
}

self.onmessage = async (e) => {
  const msg = e.data;
  try {
    if (msg.type === "init") await init(msg.bundle);
    else if (msg.type === "convert") convert(msg);
  } catch (err) {
    const text = (err && err.message) || String(err);
    const lastLine = text.trim().split(/\r?\n/).pop();
    postMessage({ type: "error", text, summary: lastLine.replace(/^(?:\w+\.)*\w*(?:Error|Exception):\s*/, "") });
  }
};
