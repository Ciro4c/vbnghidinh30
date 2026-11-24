# Bản quyền © Thevancorp-1801808400. Bảo lưu mọi quyền.
"""Điểm khởi động của ứng dụng (dùng cho cả chạy trực tiếp và file .exe).

- Tự tìm cổng trống, khởi động máy chủ, mở trình duyệt.
- Hiện một cửa sổ nhỏ; đóng cửa sổ là thoát ứng dụng.
"""
import os
import socket
import sys
import threading
import time
import webbrowser

# Bản .exe không có cửa sổ dòng lệnh nên stdout/stderr có thể là None.
if sys.stdout is None:
    sys.stdout = open(os.devnull, "w", encoding="utf-8")
if sys.stderr is None:
    sys.stderr = open(os.devnull, "w", encoding="utf-8")

import uvicorn  # noqa: E402

from app import ban_quyen  # noqa: E402
from app.main import app  # noqa: E402

HOST = "127.0.0.1"
CONG_UU_TIEN = list(range(8000, 8011))


def tim_cong(danh_sach=CONG_UU_TIEN, host=HOST):
    """Trả về cổng đầu tiên còn trống, hoặc None."""
    for port in danh_sach:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind((host, port))
                return port
            except OSError:
                continue
    return None


class MayChu(uvicorn.Server):
    def install_signal_handlers(self):  # chạy trong luồng phụ nên không bắt tín hiệu
        pass


def mo_cua_so(url, may_chu):
    import tkinter as tk

    root = tk.Tk()
    root.title("Chuẩn hóa văn bản NĐ 30")
    root.geometry("440x210")
    root.resizable(False, False)
    tk.Label(root, text="Ứng dụng đang chạy", font=("Segoe UI", 13, "bold")).pack(pady=(16, 2))
    tk.Label(root, text=url, fg="#1f5fbf", font=("Segoe UI", 11)).pack()
    tk.Label(root, text="Giữ cửa sổ này mở khi sử dụng.\nĐóng cửa sổ để thoát ứng dụng.",
             fg="#555555", font=("Segoe UI", 9)).pack(pady=8)
    tk.Label(root, text=ban_quyen.BAN_QUYEN, fg="#777777", font=("Segoe UI", 8)).pack(side="bottom", pady=6)
    bar = tk.Frame(root)
    bar.pack(pady=4)
    tk.Button(bar, text="Mở trình duyệt", width=16, command=lambda: webbrowser.open(url)).pack(side="left", padx=6)
    tk.Button(bar, text="Thoát", width=10, command=root.destroy).pack(side="left", padx=6)
    root.protocol("WM_DELETE_WINDOW", root.destroy)
    root.mainloop()


def bao_loi(noi_dung):
    try:
        import tkinter as tk
        from tkinter import messagebox
        r = tk.Tk()
        r.withdraw()
        messagebox.showerror("Chuẩn hóa văn bản NĐ 30", noi_dung)
    except Exception:
        print(noi_dung)


def main():
    port = tim_cong()
    if port is None:
        bao_loi("Không tìm được cổng trống (8000-8010). Hãy đóng bản đang chạy rồi mở lại.")
        return
    url = f"http://{HOST}:{port}"
    may_chu = MayChu(uvicorn.Config(app, host=HOST, port=port, log_config=None))
    luong = threading.Thread(target=may_chu.run, daemon=True)
    luong.start()
    for _ in range(100):                       # chờ tối đa ~10 giây
        if may_chu.started:
            break
        time.sleep(0.1)
    else:
        bao_loi("Máy chủ không khởi động được.")
        return
    webbrowser.open(url)
    try:
        mo_cua_so(url, may_chu)
    except Exception:                          # không có tkinter: chạy ở chế độ dòng lệnh
        print(f"Đang chạy tại {url}  (Ctrl+C để dừng)")
        try:
            while luong.is_alive():
                time.sleep(0.5)
        except KeyboardInterrupt:
            pass
    may_chu.should_exit = True
    luong.join(timeout=5)


if __name__ == "__main__":
    main()
