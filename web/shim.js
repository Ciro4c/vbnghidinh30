// Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
// Chặn các lời gọi fetch("/api/...") của giao diện và chuyển cho Python chạy trong trình duyệt.
(() => {
  const goc = new URL("./", document.currentScript.src).href;
  const worker = new Worker(goc + "worker.js");
  let sanSang, loiSanSang;
  const cho = new Promise((ok, hong) => { sanSang = ok; loiSanSang = hong; });
  const doi = new Map();
  let dem = 0;

  const lop = document.createElement("div");
  lop.style.cssText = "position:fixed;inset:0;z-index:9999;background:rgba(255,255,255,.96);display:flex;" +
    "flex-direction:column;align-items:center;justify-content:center;font:16px system-ui,sans-serif;text-align:center;padding:24px";
  lop.innerHTML = '<div style="font-weight:600;font-size:18px;margin-bottom:8px">Chuẩn hóa văn bản NĐ 30</div>' +
    '<div id="tien-do-py">Đang khởi động...</div>' +
    '<progress style="margin-top:14px;width:260px"></progress>';
  document.addEventListener("DOMContentLoaded", () => document.body.appendChild(lop));

  worker.onmessage = (ev) => {
    const m = ev.data;
    if (m.kieu === "tien-do") { const e = document.getElementById("tien-do-py"); if (e) e.textContent = m.msg; }
    else if (m.kieu === "san-sang") { lop.remove(); sanSang(); }
    else if (m.kieu === "loi-khoi-dong") {
      lop.innerHTML = '<div style="color:#b00020;max-width:560px"><b>Không khởi động được.</b><br>Kiểm tra kết nối mạng rồi tải lại trang (Ctrl+F5).<br><small>' +
        m.msg.replace(/</g, "&lt;") + "</small></div>";
      loiSanSang(new Error(m.msg));
    } else if (m.kieu === "tra-loi") {
      const f = doi.get(m.id); doi.delete(m.id);
      if (f) f(new Response(m.data, { status: m.status, headers: m.headers }));
    }
  };
  worker.postMessage({ kieu: "khoi-dong", goc });

  const fetchGoc = window.fetch.bind(window);
  window.fetch = async (input, init = {}) => {
    const url = typeof input === "string" ? input : input.url;
    if (!url.startsWith("/api/")) return fetchGoc(input, init);
    await cho;
    let body = null;
    if (init.body != null) body = init.body instanceof ArrayBuffer ? init.body : await new Response(init.body).arrayBuffer();
    const id = ++dem;
    return new Promise((ok) => {
      doi.set(id, ok);
      worker.postMessage({ kieu: "goi", id, method: (init.method || "GET").toUpperCase(), url, body }, body ? [body] : []);
    });
  };
})();
