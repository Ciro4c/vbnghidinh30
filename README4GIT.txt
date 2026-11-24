================================================================================
     HƯỚNG DẪN CHI TIẾT: ĐƯA MÃ NGUỒN LÊN REPO VÀ CẤU HÌNH TỰ ĐỘNG XUẤT BẢN WEB
================================================================================

Tài liệu này hướng dẫn bạn từ khâu giải nén tệp tin, đưa toàn bộ mã nguồn (bao gồm 
cả thư mục ẩn) lên kho lưu trữ trên GitHub, cho đến khi trang web chính thức hoạt động.


================================================================================
BƯỚC 1: GIẢI NÉN VÀ ĐƯA TOÀN BỘ MÃ NGUỒN LÊN REPO "vanban-nd30"
================================================================================
LƯU Ý QUAN TRỌNG: Tránh dùng tính năng Upload trên trang web vì giao diện web 
thường tự động bỏ sót các thư mục ẩn bắt đầu bằng dấu chấm (như thư mục `.github`). 
Do đó, bạn phải dùng phần mềm quản lý mã nguồn trên máy tính để thực hiện.

1. Giải nén tệp tin mã nguồn:
   - Nhấp chuột phải vào tệp nén (file `.zip`) của bạn -> Chọn giải nén (Extract All / Extract Here).
   - Mở thư mục vừa giải nén ra để đảm bảo bên trong có chứa toàn bộ các tệp dự án 
     và đặc biệt là **phải nhìn thấy thư mục ẩn tên là `.github`**. 
     *(Nếu không thấy thư mục `.github`, hãy bật chế độ hiển thị file ẩn trên máy tính: 
     Trên Windows vào View -> tích chọn "Hidden items"; trên Mac nhấn `Cmd + Shift + .`).*

2. Mở thư mục dự án vừa giải nén bằng phần mềm quản lý mã nguồn trên máy tính:
   - Mở phần mềm quản lý mã nguồn lên.
   - Nhấn tổ hợp phím mở thư mục (Ví dụ trên Windows: nhấn giữ `Ctrl + K` rồi nhấn `Ctrl + O`; 
     trên Mac: nhấn giữ `Cmd + K` rồi nhấn `Cmd + O`).
   - Tìm và chọn thư mục dự án vừa giải nén, bấm xác nhận mở (Select Folder / Open).
   - Nếu phần mềm hiển thị bảng cảnh báo bảo mật ("Trust" / Tin cậy thư mục), 
     hãy tích chọn đồng ý tin cậy và bấm Xác nhận (Trust / Yes).

3. Khởi tạo kho lưu trữ và lưu mốc dữ liệu (Commit):
   - Nhìn sang thanh công cụ quản lý thay đổi (Source Control) của phần mềm.
   - Nếu phần mềm yêu cầu khởi tạo kho lưu trữ, hãy bấm nút khởi tạo.
   - **BẮT BUỘC:** Nhấp chuột vào ô nhập nội dung mô tả (Message / Summary) và gõ một dòng chữ bất kỳ 
     (ví dụ: `Up code full thu muc`). *Nếu để trống ô này, phần mềm sẽ không cho phép lưu mốc.*
   - Bấm nút xác nhận lưu mốc (Commit / Commit to main).

4. Đẩy mã nguồn lên kho lưu trữ trực tuyến:
   - Tìm và bấm vào nút đẩy mã nguồn lên mạng (Publish / Publish Repository / Publish Branch).
   - Khi phần mềm yêu cầu đặt tên kho lưu trữ, hãy nhập chính xác tên: **`vanban-nd30`**.
   - Chọn chế độ hiển thị là **Công khai (Public)** để hệ thống có thể xây dựng website miễn phí.
   - Bấm xác nhận và đợi phần mềm chạy hoàn tất (thanh tiến trình kết thúc), đưa toàn bộ 
     tệp tin và thư mục `.github` lên mây thành công.


================================================================================
BƯỚC 2: CẤU HÌNH GITHUB ACTIONS LÀM NGUỒN XUẤT BẢN
================================================================================
Bước này chỉ cần thực hiện đúng một lần duy nhất trên trình duyệt web:

1. Mở trình duyệt web (Chrome, Edge...), truy cập vào trang chủ: https://github.com
2. Đăng nhập vào tài khoản của bạn, sau đó nhấp vào kho lưu trữ có tên **`vanban-nd30`** 
   (nằm ở danh sách bên cột trái).
3. Tại trang quản lý kho lưu trữ, nhìn lên **thanh menu ngang phía trên cùng**, nhấp vào mục 
   cuối cùng bên phải có hình bánh răng cưa mang tên **Settings**.
4. Ở **cột menu dọc bên trái** của trang Settings, cuộn chuột xuống nhóm "Code and automation", 
   nhấp chọn mục **Pages**.
5. Nhìn sang khung nội dung chính bên phải, tại mục **Build and deployment**:
   - Dòng **Source**: Nhấp vào ô đang chọn (thường mặc định là *Deploy from a branch*) 
     và đổi thành **GitHub Actions**.
   - Hệ thống sẽ tự động lưu lại thay đổi (hoặc bạn bấm nút Save nếu có).


================================================================================
BƯỚC 3: KIỂM TRA TÁC VỤ TỰ ĐỘNG (ACTIONS)
================================================================================
1. Ngay tại thanh menu ngang phía trên cùng của kho lưu trữ `vanban-nd30`, tìm và nhấp vào tab 
   **Actions** (có biểu tượng hình tam giác phát sáng / mũi tên chạy vòng).
2. Tại trang Actions, nhìn sang cột menu bên trái, tìm và nhấp chọn dòng workflow có tên 
   **“Dua len GitHub Pages”** (hoặc tên tương ứng do hệ thống cấu hình sẵn).
3. Kiểm tra trạng thái tự động:
   - Nếu workflow đã tự động chạy, bạn sẽ thấy một hàng đang chạy (hiện icon vòng tròn vàng hoặc cam). 
     Hãy bấm vào đó và đợi khoảng 1 đến 2 phút cho đến khi biểu tượng chuyển thành **dấu tích màu xanh lá (Success)**.
   - Nếu workflow **chưa** tự động chạy, hãy tìm nút màu vàng hoặc xanh dương có chữ **`Run workflow`** 
     (nằm ở phía trên bên phải danh sách), nhấp vào đó và bấm nút **`Run workflow`** màu xanh để kích hoạt thủ công.
   - Đợi cho đến khi quá trình chạy hoàn tất và xuất hiện dấu tích xanh.


================================================================================
BƯỚC 4: TRUY CẬP VÀ TRẢI NGHIỆM TRANG WEB
================================================================================
1. Mở một tab trình duyệt mới, nhập vào thanh địa chỉ đường dẫn chính xác của bạn:
   👉 **`https://[ten_tai_khoan_github_cua_ban].giHub.io/vanban-nd30/`**
   *(Thay thế `[ten_tai_khoan_github_cua_ban]` bằng tên tài khoản GitHub thực tế của bạn).*
2. **Lưu ý trong lần truy cập đầu tiên:**
   - Vì ứng dụng bên trong sử dụng công nghệ biên dịch mã nguồn trực tiếp trên trình duyệt, 
     ở lần truy cập đầu tiên trang web sẽ hiển thị màn hình thông báo **“Đang tải Python…”**.
   - Hãy kiên nhẫn giữ nguyên tab này khoảng **30 đến 60 giây** để hệ thống tải các thư viện cần thiết. 
     Sau thời gian này, giao diện trang web chính thức sẽ hiện ra hoàn chỉnh.
================================================================================