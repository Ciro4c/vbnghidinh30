# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Đóng gói bản chạy trong trình duyệt (GitHub Pages + Pyodide) vào thư mục _site/.

Chạy:  python scripts/build_pages.py
"""
import shutil
import sys
import zipfile
from pathlib import Path

GOC = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(GOC))
from app import ban_quyen  # noqa: E402

DICH = GOC / "_site"


def thay(html: str, cu: str, moi: str) -> str:
    if cu not in html:
        raise SystemExit(f"Không tìm thấy đoạn cần sửa trong index.html:\n{cu[:80]}...")
    return html.replace(cu, moi, 1)


def main():
    if DICH.exists():
        shutil.rmtree(DICH)
    DICH.mkdir()
    ban_quyen.kiem_tra()

    # 1) mã nguồn Python (không kèm static) -> app.zip
    with zipfile.ZipFile(DICH / "app.zip", "w", zipfile.ZIP_DEFLATED) as z:
        for f in sorted((GOC / "app").rglob("*")):
            if f.is_file() and "__pycache__" not in f.parts and "static" not in f.relative_to(GOC / "app").parts:
                z.write(f, f.relative_to(GOC).as_posix())

    # 2) trang web: thêm shim + chân trang bản quyền + tải ZIP bằng fetch
    html = (GOC / "app" / "static" / "index.html").read_text(encoding="utf-8")
    html = thay(html, '<script>', '<script src="shim.js"></script>\n<script>')
    html = thay(
        html,
        '''const a = document.createElement("a"); a.href = `/api/lo/${loId}/zip`; a.download = "ket-qua-chuan-hoa.zip";
  document.body.appendChild(a); a.click(); a.remove();''',
        '''fetch(`/api/lo/${loId}/zip`).then(async (r) => { if (!r.ok) throw new Error(await loiTuServer(r)); taiVe(await r.blob(), "ket-qua-chuan-hoa.zip"); })
    .catch((e) => setMsg("msg_lo", e.message, true));''',
    )
    html = thay(html, 'const PDFJS_GOC = "/static/vendor/";', 'const PDFJS_GOC = new URL("vendor/", location.href).href;')
    chan = ('<footer id="chu-ky" style="text-align:center;padding:18px 12px 28px;font:13px system-ui,sans-serif;'
            f'opacity:.7">{ban_quyen.BAN_QUYEN} · <span>{ban_quyen.CHU_KY}</span></footer>')
    html = thay(html, "</body>", chan + "\n</body>")
    (DICH / "index.html").write_text(html, encoding="utf-8")

    for ten in ("shim.js", "worker.js", "web_bridge.py"):
        shutil.copy(GOC / "web" / ten, DICH / ten)
    shutil.copytree(GOC / "app" / "static" / "vendor", DICH / "vendor")
    (DICH / ".nojekyll").write_text("", encoding="utf-8")
    print("Da tao", DICH)


if __name__ == "__main__":
    main()
