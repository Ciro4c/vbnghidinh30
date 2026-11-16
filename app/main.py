# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Web app chuẩn hóa văn bản theo NĐ 30/2020/NĐ-CP và slide chuẩn.

Chạy:  python -m uvicorn app.main:app --reload
"""
import json
import os
from pathlib import Path

from fastapi import FastAPI, HTTPException, Query, Request
from fastapi.responses import FileResponse, HTMLResponse, Response
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool

from . import ban_quyen, lo as lo_mod
from .importers.docx_reader import doc_docx_ex, doc_txt
from .importers.phan_tich import phan_tich
from .importers.pptx_reader import doc_pptx
from .profile import DEFAULT_PROFILE
from .renderers.cong_van import cong_van_to_bytes
from .renderers.slide_hoc_thuat import ThongTin, dung_slide, ten_profile
from .schemas import TEN_LOAI, CongVan

STATIC_DIR = Path(__file__).parent / "static"
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
PPTX_MIME = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
try:                                    # chạy trên máy cá nhân nên cho phép file lớn; chỉnh bằng biến môi trường
    MAX_MB = max(1, int(os.environ.get("VANBAN_MAX_MB", "500")))
except ValueError:
    MAX_MB = 500
MAX_BYTES = MAX_MB * 1024 * 1024

ban_quyen.kiem_tra()

app = FastAPI(title="Chuẩn hóa văn bản NĐ 30/2020 và slide chuẩn")


@app.middleware("http")
async def them_tac_gia(request: Request, call_next):
    resp = await call_next(request)
    resp.headers["X-Tac-Gia"] = ban_quyen.CHU_KY
    return resp

app.mount("/static", StaticFiles(directory=STATIC_DIR), name="static")


async def _doc_file(request: Request) -> bytes:
    data = await request.body()
    if not data:
        raise HTTPException(400, "Chưa nhận được nội dung file.")
    if len(data) > MAX_BYTES:
        raise HTTPException(413, f"File quá lớn (tối đa {MAX_MB} MB). Có thể nâng giới hạn bằng biến "
                                 "môi trường VANBAN_MAX_MB.")
    return data


def _dau_header(warns: list[str]) -> str:
    gon = [w[:300] for w in warns[:15]]
    return json.dumps(gon, ensure_ascii=True)


CHAN_TRANG = (
    '<footer id="chu-ky" style="text-align:center;padding:18px 12px 28px;font:13px system-ui,sans-serif;'
    'opacity:.7">{ban_quyen} · <span>{chu_ky}</span></footer>'
)


@app.get("/")
def trang_chu():
    ban_quyen.kiem_tra()
    html = (STATIC_DIR / "index.html").read_text(encoding="utf-8")
    chan = CHAN_TRANG.format(ban_quyen=ban_quyen.BAN_QUYEN, chu_ky=ban_quyen.CHU_KY)
    html = html.replace("</body>", chan + "\n</body>") if "</body>" in html else html + chan
    return HTMLResponse(html)


@app.get("/api/tac-gia")
def tac_gia():
    return {"chu_ky": ban_quyen.CHU_KY, "ban_quyen": ban_quyen.BAN_QUYEN}


@app.get("/api/loai-van-ban")
def loai_van_ban():
    return {k: (v or "Công văn") for k, v in TEN_LOAI.items()}


@app.post("/api/import")
async def nhap_file(request: Request, ten: str = Query("")):
    """Đọc file Word/TXT thô, nhận diện cấu trúc, trả về các trường để rà soát."""
    data = await _doc_file(request)
    ext = Path(ten).suffix.lower()
    try:
        hinh: dict = {}
        if ext == ".docx":
            lines, w0, hinh = doc_docx_ex(data)
        elif ext == ".txt":
            lines, w0 = doc_txt(data)
        elif ext == ".doc":
            raise HTTPException(400, "File .doc cũ chưa đọc được. Hãy mở bằng Word, chọn "
                                     "Lưu thành (Save As) định dạng .docx rồi tải lại.")
        else:
            raise HTTPException(400, "Chỉ nhận file .docx hoặc .txt (PowerPoint: dùng tab Slide).")
    except HTTPException:
        raise
    except Exception:
        raise HTTPException(400, "Không đọc được file (có thể bị hỏng hoặc đặt mật khẩu).")
    if not lines:
        raise HTTPException(400, "File không có nội dung văn bản.")
    fields, w1 = phan_tich(lines)
    return {"fields": fields, "canh_bao": w0 + w1, "hinh": hinh}


@app.post("/api/cong-van/docx")
def tao_van_ban(data: CongVan, profile: str = Query(DEFAULT_PROFILE)):
    try:
        content = cong_van_to_bytes(data, profile)
    except (FileNotFoundError, ValueError, KeyError) as e:
        raise HTTPException(status_code=400, detail=str(e))
    return Response(content=content, media_type=DOCX_MIME,
                    headers={"Content-Disposition": 'attachment; filename="van-ban.docx"'})


@app.post("/api/slide/chuan-hoa")
async def chuan_hoa_slide(request: Request, ten: str = Query(""), ten_bai: str = "",
                          hoc_phan: str = "", giang_vien: str = "", don_vi: str = "",
                          nam: str = "", kieu: str = "chuan"):
    """Nhận file .pptx thô, trả về .pptx đã dựng lại theo chuẩn slide (kieu: chuan | hoc_thuat)."""
    data = await _doc_file(request)
    ext = Path(ten).suffix.lower()
    if ext == ".ppt":
        raise HTTPException(400, "File .ppt cũ chưa đọc được. Hãy mở bằng PowerPoint, chọn "
                                 "Lưu thành định dạng .pptx rồi tải lại.")
    if ext != ".pptx":
        raise HTTPException(400, "Chỉ nhận file .pptx.")
    try:
        slides, w0 = doc_pptx(data)
    except Exception:
        raise HTTPException(400, "Không đọc được file (có thể bị hỏng hoặc đặt mật khẩu).")
    if not slides:
        raise HTTPException(400, "File không có slide nào có nội dung.")
    out, w1, n = dung_slide(slides, ThongTin(ten_bai.strip(), hoc_phan.strip(),
                                             giang_vien.strip(), don_vi.strip(), nam.strip()),
                               profile=ten_profile(kieu))
    return Response(content=out, media_type=PPTX_MIME, headers={
        "Content-Disposition": 'attachment; filename="slide-chuan.pptx"',
        "X-Canh-Bao": _dau_header(w0 + w1),
        "X-So-Slide": str(n),
    })


# ---------------------------------------------------------------- xử lý hàng loạt
@app.post("/api/lo/tao")
def lo_tao():
    """Mở một lô mới; các file thêm vào lô qua /api/lo/{lo}/them."""
    return {"lo": lo_mod.tao_lo()}


@app.post("/api/lo/{lo}/them")
async def lo_them(lo: str, request: Request, ten: str = Query(""), hoc_phan: str = "",
                  giang_vien: str = "", don_vi: str = "", nam: str = "", kieu: str = "chuan"):
    """Xử lý một file của lô (Word -> NĐ 30, PowerPoint -> slide chuẩn), trả về kết quả của file đó."""
    data = await _doc_file(request)
    tt = {"hoc_phan": hoc_phan, "giang_vien": giang_vien, "don_vi": don_vi, "nam": nam, "kieu": kieu}
    try:
        return await run_in_threadpool(lo_mod.xu_ly_tep, lo, ten, data, tt)
    except (ValueError, FileNotFoundError) as e:
        raise HTTPException(400, str(e))


@app.get("/api/lo/{lo}/zip")
def lo_zip(lo: str):
    """Tải toàn bộ kết quả của lô (kèm BAO-CAO.csv) dưới dạng một file ZIP."""
    try:
        path = lo_mod.tao_zip(lo)
    except (ValueError, FileNotFoundError) as e:
        raise HTTPException(404, str(e))
    return FileResponse(path, media_type="application/zip", filename="ket-qua-chuan-hoa.zip")
