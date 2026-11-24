# vanban-nd30 — Chuẩn hóa văn bản (NĐ 30/2020 và thể thức khác), PDF và slide

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

## Xử lý hàng loạt (tab cuối)

Mỗi lô chỉ làm **một việc**, chọn ở Bước 1; các tùy chọn bên dưới đổi theo việc đã chọn:

| Việc | File nhận | File ra | Tùy chọn |
|---|---|---|---|
| Chuẩn hóa văn bản Word | .docx, .txt | .docx | thể thức, thành phần bỏ/giữ, dọn ký tự lạ |
| Xử lý PDF | .pdf | .pdf | dấu trang (giữ gốc / tự thêm theo thể thức), watermark chữ và hình, lưu dạng ảnh |
| Chuyển PowerPoint | .pptx | .pptx | kiểu slide (mục đích), chủ đề, người trình bày, đơn vị |

File không đúng loại của việc đã chọn được **bỏ qua** (ghi “Bỏ qua” trong báo cáo, không tính là lỗi). Một file lỗi không làm dừng cả lô. Kết quả gom thành ZIP kèm `BAO-CAO.csv`. Chữ ký/mộc trong PDF đặt vị trí riêng từng văn bản nên không có trong chế độ lô.

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

## Tính năng nâng cấp (phiên bản này)

**Thể thức tùy chọn.** Chọn thể thức ở đầu form (mặc định NĐ 30/2020; kèm: Đảng 05-HD/VPTW, Đoàn TNCS HCM 29-HD/TWĐTN-VP, Công đoàn QĐ 933/QĐ-TLĐ, Quân đội TT 92/2012/TT-BQP). Mỗi thành phần (quốc hiệu, tiêu ngữ, cơ quan, số ký hiệu, địa danh ngày, tên loại, trích yếu, kính gửi, chữ ký, nơi nhận, số trang, phụ lục) có ô chọn; có mẫu nhanh như “bỏ quốc hiệu và tiêu ngữ”. Thông số chi tiết (lề, cỡ chữ...) ghi đè bằng JSON, được kiểm tra khoảng hợp lệ. Cấu hình lưu/nạp ra file JSON. Thể thức mới: thêm một file vào `app/profiles/` (có thể `extends: nd30` rồi chỉ ghi phần khác). **Lưu ý:** các thể thức ngoài NĐ 30 được dựng từ bản tóm tắt hoặc từ mẫu văn bản thực tế (Đoàn: lề và cỡ chữ chưa đọc được từ bản gốc; Công đoàn: chưa nạp Phụ lục 1 nên dùng thông số NĐ 30; Quân đội: số trang đặt giữa lề trên thay vì góc phải chân trang, chưa kiểm tra văn bản còn hiệu lực), cần đối chiếu bản gốc trước khi dùng chính thức.

**Dò ký tự lạ do AI.** Phát hiện ký tự vô hình (zero-width, điều khiển chiều chữ), khoảng trắng lạ, ký hiệu Markdown (`**`, `#`, bảng `|---|`), mã trích dẫn chatbot (`[oaicite]`, `turn0search0`, `【…†source】`), emoji, chữ giả dạng (homoglyph), lỗi mã hóa tiếng Việt, NFD/NFC, ô giữ chỗ chưa điền, câu dẫn của chatbot. Nút “Dò ký tự lạ” trong ô nội dung, tùy chọn tự dọn khi đọc file Word/TXT/PDF, và mục “Ký tự lạ” trong tab PDF. Mục nào không tự sửa an toàn được thì chỉ báo.

**Chữ ký và con dấu.** Tải ảnh chữ ký/mộc (xóa nền trắng, chỉnh độ đậm, nhuộm màu mực). Trong Word: chữ ký nằm trên họ tên, con dấu là ảnh nổi chồm khoảng 1/3 lên chữ ký về bên trái. Trong PDF: chọn trang, vị trí theo %, kích thước mm. Đây là ảnh hình thức, **không** thay cho chữ ký số có giá trị pháp lý.

**Tab PDF.** Phân tích (lớp chữ, khổ giấy, kiểm tra thành phần thể thức theo chữ, dấu trang hiện có và gợi ý), sửa danh sách dấu trang theo thể thức, watermark chữ và hình (độ đậm/mờ 3–100%, giữa trang/lát đều/đặt vị trí, chọn trang), chèn chữ ký/mộc, và lưu dạng ảnh (phủ phẳng: gộp mọi lớp vào ảnh JPEG từng trang, chọn dpi, chất lượng, đen trắng). Trang được vẽ bằng pdf.js (đã kèm trong `app/static/vendor/`, giấy phép Apache-2.0) ngay trong trình duyệt. Khi lưu dạng ảnh **văn bản không còn chọn/tìm kiếm được**, dung lượng thường tăng.

**Slide đa mục đích.** Kiểu slide chọn theo mục đích: Chuẩn Action Title, Báo cáo, Thuyết trình, Bán hàng/đề xuất (và hai kiểu cũ). Các tiêu chuẩn bạn nêu được đưa thành thông số trong `app/profiles/slide_action.yaml`, **không khóa cứng** (đổi trong yaml hoặc từng lần ở “Tùy biến nâng cao”):

| Tiêu chuẩn | Cách xử lý |
|---|---|
| Action Title (câu kết luận, 28–36pt đậm, trái, vùng 15–20% trên) | Dựng: cỡ chữ, vùng, canh trái. Kiểm tra: gợi ý khi tiêu đề là nhãn chủ đề; bạn sửa tiêu đề ngay trong giao diện |
| Chunking (3–5 ý, 5–7 từ/dòng) | Dựng: tự tách slide khi vượt số ý; tùy chọn rút ý dài thành ý chính + câu đầy đủ vào ghi chú. Kiểm tra: báo ý dài |
| WCAG AA/AAA, đơn sắc | Kiểm tra từng cặp màu (tỷ lệ tương phản và khi in đen trắng); màu vi phạm được báo, không tự đổi |
| 60-30-10 | Dựng: nền/chủ đạo/điểm nhấn/nhấn từ khóa là thông số màu; số liệu (%, tỷ, triệu...) tô màu nhấn |
| Gestalt, 3 vùng, lề 5–8% | Dựng: vùng tiêu đề/trọng tâm/chân trang tính theo %; kiểm tra khoảng tham khảo |
| Chữ không chân, 18–24pt, chân trang 12–14pt | Dựng theo thông số; kiểm tra khoảng tham khảo |
| Bố cục thẻ cho 3–5 ý ngắn | Dựng tự động (tắt được) |
| Hiệu ứng xuất hiện dần, làm mờ 30% | **Chưa làm**: chưa có cách kiểm tra hiệu ứng trong PowerPoint thật; thông số `animation` trong yaml hiện chưa tác dụng |
| Biểu đồ Data-Ink, ảnh minh chứng/APA, hướng nhìn, quy tắc 1/3, ghép ảnh-chữ | **Không tự động hóa được** (cần nhìn nội dung); chưa có kiểm tra |

**API chung.** Mọi chức năng đi qua `app/api_core.py` (FastAPI và bản trình duyệt dùng chung). Thêm các endpoint: `/api/thong-tin`, `/api/profile/<id>`, `/api/ky-tu-la`, `/api/pdf/phan-tich`, `/api/pdf/xu-ly`, `/api/slide/phan-tich`; gói nhiều tệp dùng khung `VBN1` (`app/khung.py`).

**Chưa kiểm chứng.** Bản GitHub Pages (Pyodide) đã được thêm thư viện `pypdf` và vẽ chữ bằng Pillow nhưng **chưa chạy thử** trong Pyodide thật; bản máy cá nhân (FastAPI/.exe) đã chạy thử đầy đủ trên trình duyệt. Bản Pages có thể cần chỉnh nếu Pillow ở Pyodide thiếu freetype.

## Bản quyền

© Thevancorp-1801808400. Xem `LICENSE.txt`. Chữ ký tác giả nằm ở: đầu mọi file mã nguồn, `app/ban_quyen.py`,
chân trang giao diện (do máy chủ chèn), header `X-Tac-Gia`, thuộc tính (Comments) của file Word/PowerPoint
tạo ra, và thông tin phiên bản của file .exe (chuột phải → Properties → Details).

## Bản chạy trên GitHub Pages (Pyodide)
Workflow `.github/workflows/pages.yml` chạy `scripts/build_pages.py`, đóng gói `app/` + `web/` vào `_site/`
và đưa lên GitHub Pages. Python chạy ngay trong trình duyệt (Pyodide), không cần máy chủ.
Bật một lần: Settings → Pages → Source: **GitHub Actions**.
