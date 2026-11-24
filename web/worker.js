// Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
// Web Worker: nạp Pyodide + thư viện + mã nguồn app/, rồi xử lý yêu cầu từ trang.
const PYODIDE = "https://cdn.jsdelivr.net/pyodide/v0.26.4/full/";
importScripts(PYODIDE + "pyodide.js");

let bridge = null;
const baoTien = (msg) => postMessage({ kieu: "tien-do", msg });

async function khoiDong(goc) {
  baoTien("Đang tải Python (lần đầu mất khoảng 30–60 giây)...");
  const py = await loadPyodide({ indexURL: PYODIDE });
  baoTien("Đang tải thư viện xử lý Word/PowerPoint...");
  await py.loadPackage(["micropip", "lxml", "pillow", "pyyaml", "pydantic", "typing-extensions"]);
  await py.runPythonAsync(`
import micropip
await micropip.install(["python-docx", "python-pptx", "pypdf"])
`);
  baoTien("Đang nạp ứng dụng...");
  const r = await fetch(goc + "app.zip", { cache: "no-cache" });
  if (!r.ok) throw new Error("Không tải được app.zip (" + r.status + ")");
  py.unpackArchive(await r.arrayBuffer(), "zip", { extractDir: "/home/pyodide" });
  const b = await fetch(goc + "web_bridge.py", { cache: "no-cache" });
  py.FS.writeFile("/home/pyodide/web_bridge.py", await b.text());
  await py.runPythonAsync(`
import sys
sys.path.insert(0, "/home/pyodide")
`);
  bridge = py.pyimport("web_bridge");
}

function layBytes(p) {
  let b;
  try { b = p.toJs(); } catch (e) { b = null; }
  if (!(b instanceof Uint8Array)) { const buf = p.getBuffer(); b = buf.data.slice(); buf.release(); }
  p.destroy && p.destroy();
  return b;
}

onmessage = async (ev) => {
  const m = ev.data;
  if (m.kieu === "khoi-dong") {
    try { await khoiDong(m.goc); postMessage({ kieu: "san-sang" }); }
    catch (e) { postMessage({ kieu: "loi-khoi-dong", msg: String((e && e.message) || e) }); }
    return;
  }
  if (m.kieu === "goi") {
    try {
      const body = m.body ? new Uint8Array(m.body) : null;
      const r = bridge.handle(m.method, m.url, body);
      const status = r.get(0), headers = JSON.parse(r.get(1)), data = layBytes(r.get(2));
      r.destroy();
      postMessage({ kieu: "tra-loi", id: m.id, status, headers, data }, [data.buffer]);
    } catch (e) {
      postMessage({ kieu: "tra-loi", id: m.id, status: 500, headers: { "Content-Type": "application/json" },
        data: new TextEncoder().encode(JSON.stringify({ detail: "Lỗi: " + ((e && e.message) || e) })) });
    }
  }
};
