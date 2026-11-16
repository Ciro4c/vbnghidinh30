# vanban-nd30 — Chuẩn hóa văn bản NĐ 30/2020/NĐ-CP và slide học thuật

Web app chạy trên máy cá nhân, **không cần mạng, không tốn phí, không dùng AI**.

| Chức năng | Đầu vào | Đầu ra |
|---|---|---|
| Văn bản hành chính | file **Word thô** (.docx, .txt) hoặc nhập form | file Word đúng thể thức NĐ 30 |
| Slide học thuật | file **PowerPoint thô** (.pptx) | file PowerPoint dựng lại theo chuẩn học thuật |
| **Xử lý hàng loạt** | nhiều file hoặc cả thư mục (.docx, .txt, .pptx) | một file ZIP kết quả + `BAO-CAO.csv` |

## Chạy (Windows)

1. Cài Python 3.11+ từ https://www.python.org/downloads/ (tick **Add python.exe to PATH**).
2. Nhấp đúp `CHAY.bat` (lần đầu tự cài thư viện, cần mạng một lần).
3. Trình duyệt mở `http://127.0.0.1:8000`.

## Đóng gói thành file .exe (không cần cài Python trên máy dùng)

1. Trên máy Windows có Python, nhấp đúp `BUILD_EXE.bat` (cần mạng; mất vài phút).
2. Kết quả: thư mục `dist\VanBanND30\`. Chạy `VanBanND30.exe` trong đó. Copy cả thư mục sang máy khác là dùng được.
3. Muốn gọn một file duy nhất: nhấp đúp `BUILD_EXE_1FILE.bat` (khởi động chậm hơn vài giây).

Chạy .exe sẽ mở một cửa sổ nhỏ và trình duyệt tại `http://127.0.0.1:8000`. Đóng cửa sổ nhỏ để thoát.
File .exe chưa ký số nên Windows SmartScreen có thể cảnh báo: chọn *More info → Run anyway*.
Một số phần mềm diệt virus báo nhầm với file đóng gói bằng PyInstaller; dạng thư mục ít bị hơn dạng một file.

Thử ngay với các file thô trong thư mục `mau-tho/`. Kết quả mong đợi để so sánh nằm trong `out/`.

## Cách dùng

**Văn bản hành chính:** tab 1 → chọn file .docx → *Đọc file và điền form* → rà soát các ô
(cảnh báo màu vàng chỉ ra chỗ hệ thống đoán) → *Tạo file Word*.
Tick “Tải file chuẩn về ngay” nếu muốn chuyển một lần không cần rà soát.

Hệ thống tự nhận diện: cơ quan chủ quản/ban hành, số và ký hiệu, địa danh và ngày tháng, loại văn bản
(công văn, quyết định, tờ trình, thông báo, báo cáo, kế hoạch...), trích yếu, kính gửi, mục/tiểu mục,
“Căn cứ”, “Điều N.”, nơi nhận, quyền hạn ký, chức vụ, họ tên. Đọc được cả đầu văn bản dựng bằng
bảng hoặc gõ bằng phím Tab. Tự dọn khoảng trắng thừa, dấu câu, đổi gạch đầu dòng.

**Slide:** tab 2 → chọn file .pptx → điền tên bài, giảng viên, đơn vị → *Chuyển thành slide chuẩn*.
Hệ thống dựng slide mở đầu, mục lục, tiêu đề thống nhất, gạch đầu dòng đúng cấp, tự tách slide dài
thành “(tiếp)” (kể cả một đoạn văn quá dài: tách theo câu, không tràn khỏi slide), tách đoạn văn nằm nhầm trong ô tiêu đề
xuống phần nội dung, giữ chỗ xuống dòng mềm (Shift+Enter) để chữ không bị dính nhau, đánh số slide, chân trang đơn vị, đưa “Tài liệu tham khảo” xuống cuối, giữ hình, bảng,
ghi chú của diễn giả.

## Xử lý hàng loạt (tab 3)

Chọn nhiều file cùng lúc, hoặc chọn cả một thư mục (lấy cả thư mục con). Hệ thống xử lý lần lượt từng file,
file lỗi không làm dừng cả lô. Kết thúc bấm **Tải tất cả** để lấy file ZIP gồm các file đã chuẩn hóa (giữ cấu trúc
thư mục, tên thêm `-chuan`) và `BAO-CAO.csv` (mở bằng Excel) ghi rõ file nào Đạt / Cần rà soát / Lỗi và lý do.
Với Word, thông tin thiếu (cơ quan ban hành, ngày) được điền chỗ giữ chỗ và ghi vào báo cáo để sửa sau.
Thông tin học phần, giảng viên, đơn vị, năm học ở tab này áp cho mọi slide; tên bài lấy từ slide đầu của từng file.

## Bảng và hình

- **Word:** bảng dữ liệu (có viền) được giữ thành bảng có viền, hàng đầu in đậm và lặp lại khi sang trang, ô gộp
  ngang/dọc được giữ; bảng chỉ dùng để dàn trang (đầu văn bản, khối chữ ký) vẫn đọc như cũ. Hình được giữ, tự thu nhỏ
  nếu quá lớn, không vượt bề rộng vùng chữ. Hộp văn bản (text box) được đọc ngay sau đoạn chứa hộp.
  Trong form, bảng viết dạng `| ô | ô |` (ô gộp: `<` gộp trái, `^` gộp trên), hình là dòng `[[hinh:N]]`.
- **PowerPoint:** bảng dài tự tách nhiều slide (lặp lại hàng tiêu đề, không cắt ngang ô gộp dọc), tự chọn cỡ chữ cho vừa;
  ô gộp được giữ; nhiều hình trên một slide được xếp lưới, hình quá lớn được thu nhỏ; biểu đồ được đổi thành bảng số liệu.

## Kiểu slide

Có hai kiểu, chọn trong tab Slide (và ở tab Xử lý hàng loạt); mỗi kiểu là một file trong `app/profiles/`:

- **Tiêu chuẩn** (`slide_chuan.yaml`, mặc định): dùng cho báo cáo ủy ban/hội đồng, kinh doanh, hành chính. Nền trắng,
  phông Calibri, tông xanh navy; slide mở đầu canh trái có khối trang trí; mục lục “Nội dung chính”; chân trang có
  đường kẻ mảnh và số slide; có slide kết “Xin trân trọng cảm ơn” (nếu file gốc đã có slide cảm ơn/hỏi đáp thì dùng
  luôn slide đó làm slide kết, không thêm lần nữa).
- **Học thuật** (`slide_hoc_thuat.yaml`): bài giảng, phông Times New Roman, không có slide kết.

Muốn đổi màu, phông, cỡ chữ hay bố cục: sửa file yaml tương ứng, không sửa code. Muốn thêm một kiểu mới: chép một
file yaml, đổi tên, rồi thêm vào `KIEU_SLIDE` trong `renderers/slide_hoc_thuat.py`.

## Phụ lục sau chữ ký (quyết định kèm quy trình, biểu mẫu...)

Mọi thứ nằm sau khối chữ ký (tiêu đề phụ lục, bảng, biểu mẫu, hình) được giữ nguyên ở ô **Phụ lục** của form,
không còn bị nhét vào khối chữ ký. Phụ lục luôn bắt đầu ở trang riêng, tiếp tục đánh số trang; nếu file gốc
dùng giấy A4 nằm ngang cho phần này thì file kết quả cũng ngang (có thể tick/bỏ tick “giấy nằm ngang” trong form).
Ngắt trang giữa các biểu mẫu được giữ (dòng `[[ngat-trang]]`). Bảng trong phụ lục dùng cỡ chữ nhỏ hơn
(`table.phu_luc_size_pt` trong `nd30.yaml`). Công thức LaTeX đơn giản dán từ chat AI
(`$\phi$`, `$\le 15$`, `$^\circ\text{C}$`) được đổi thành ký tự thật (φ, ≤, °C).

## Giới hạn dung lượng

Mặc định mỗi file tối đa **500 MB** (app chạy trên máy cá nhân). Muốn đổi, đặt biến môi trường trước khi chạy,
ví dụ `set VANBAN_MAX_MB=1000` rồi chạy `CHAY.bat`.

## Về số trang và canh lề

- Số trang đặt giữa lề trên, **bắt đầu từ trang 2** (đúng NĐ 30). Văn bản 1 trang sẽ không hiện số.
  Muốn hiện cả trang 1 thì tick ô tương ứng trong form (không đúng quy định).
- Lề: trên 20, dưới 20, trái 30, phải 15 mm; nội dung canh đều hai bên. Chỉnh trong
  `app/profiles/nd30.yaml`.

## Cấu trúc

```
vanban-nd30/
├─ CHAY.bat                  # chạy bằng Python
├─ BUILD_EXE.bat             # đóng gói thành .exe
├─ chay_app.py               # điểm khởi động (cửa sổ nhỏ + mở trình duyệt)
├─ mau-tho/                  # file thô mẫu để thử
├─ out/                      # kết quả mẫu đã chuẩn hóa
├─ app/
│  ├─ main.py                # API: /api/import, /api/cong-van/docx, /api/slide/chuan-hoa
│  ├─ schemas.py             # dữ liệu văn bản + danh sách loại văn bản
│  ├─ profiles/
│  │  ├─ nd30.yaml           # thể thức NĐ 30 (lề, font, cỡ chữ...)
│  │  └─ slide_hoc_thuat.yaml# chuẩn slide (màu, cỡ chữ, số dòng tối đa...)
│  ├─ importers/
│  │  ├─ docx_reader.py      # đọc Word/TXT thô
│  │  ├─ phan_tich.py        # nhận diện cấu trúc văn bản hành chính
│  │  └─ pptx_reader.py      # đọc PowerPoint thô
│  ├─ renderers/
│  │  ├─ common.py
│  │  ├─ cong_van.py         # dựng văn bản NĐ 30
│  │  └─ slide_hoc_thuat.py  # dựng slide chuẩn
│  ├─ hinh.py                # xử lý hình dùng chung (thu nhỏ, mã hóa)
│  ├─ lo.py                  # xử lý hàng loạt + ZIP + báo cáo
│  └─ static/index.html
├─ scripts/tao_mau.py
└─ tests/test_cong_van.py    # python -m pytest -q
```

## Giới hạn hiện tại (nói thẳng)

- Nhận diện cấu trúc dùng **quy tắc**, không phải AI nên có thể đoán sai với file quá khác thường
  (đầu văn bản không có bảng và không có “Số:”, chữ ký đặt trước nơi nhận...). Luôn xem cảnh báo và rà soát.
- Chưa có: văn bản có Chương/Mục/Phần phức tạp (nghị quyết dài, quy chế nhiều chương), biên bản
  có bảng ký nhiều bên. Phụ lục được giữ nguyên nội dung nhưng chỉ dàn lại theo kiểu chung (tiêu đề, bảng, gạch đầu
  dòng), không sao chép từng chi tiết định dạng của file gốc; biểu mẫu không có ngắt trang trong file gốc thì không
  tự ngắt trang.
- Danh sách đánh số tự động của Word (1., 2., a, b) bị đọc thành gạch đầu dòng.
- File .doc/.ppt cũ phải Lưu thành .docx/.pptx trước.
  Hình dạng EMF/WMF, hình liên kết ngoài, bảng lồng trong bảng, hình ở khối chữ ký/con dấu và logo đầu văn bản không được chuyển (có cảnh báo).
- Slide: SmartArt, video, hiệu ứng, chuyển cảnh **không** được chuyển; hệ thống báo số slide cần chèn lại.
  Biểu đồ chỉ giữ số liệu (dạng bảng). Định dạng riêng trong ô bảng (màu, đậm từng chữ) không được giữ.
- Chưa có phân hệ **giáo trình, bài giảng bằng Word** theo thể thức học thuật.
- Bề rộng đường kẻ tính theo font Times New Roman trên máy; không có font này thì ước lượng, có thể lệch vài mm.
- **Luôn mở file bằng Word/PowerPoint kiểm tra lại trước khi ban hành.** Đơn vị có quy chế riêng thì
  tạo profile mới trong `app/profiles/`.

## Bản quyền

© Thevancorp-1801808400. Xem `LICENSE.txt`. Chữ ký tác giả nằm ở: đầu mọi file mã nguồn, `app/ban_quyen.py`,
chân trang giao diện (do máy chủ chèn), header `X-Tac-Gia`, thuộc tính (Comments) của file Word/PowerPoint
tạo ra, và thông tin phiên bản của file .exe (chuột phải → Properties → Details).

## Bản chạy trên GitHub Pages (Pyodide)
Workflow `.github/workflows/pages.yml` chạy `scripts/build_pages.py`, đóng gói `app/` + `web/` vào `_site/`
và đưa lên GitHub Pages. Python chạy ngay trong trình duyệt (Pyodide), không cần máy chủ.
Bật một lần: Settings → Pages → Source: **GitHub Actions**.
