# Ứng dụng Tự Động Biên Mục Bản Tin Thế Giới 24H (TG24H)

Ứng dụng Desktop giúp tự động hóa quy trình biên mục bản tin **Thế Giới 24H** của HTV. Hỗ trợ hai nhà cung cấp **Gemini API (Provider 1)** và **Mistral AI API (Provider 2)** để bóc tách nội dung từ các file kịch bản `.rtf` và sinh ra 3 file Excel đầu ra theo chuẩn biên mục của Trung tâm Tư liệu HTV.

---

## Yêu cầu hệ thống

- **Python** 3.10 trở lên
- Cài đặt các thư viện Python cần thiết bằng lệnh:

```bash
pip install -r requirements.txt
```

---

## Cách sử dụng

### 1. Chạy ứng dụng

```bash
python app_bien_muc_tg24h.py
```

### 2. Cấu hình API Key & Chọn Provider

Nhấn nút **⚙ Cài đặt API** để mở hộp thoại cài đặt:
- **Chọn Provider**:
  - **Gemini (Mặc định)**: Sử dụng các mô hình của Google Gemini.
  - **Mistral**: Sử dụng các mô hình của Mistral AI.
- **Provider 1 – Gemini**:
  - **Gemini API Key**: Lấy tại [aistudio.google.com](https://aistudio.google.com/apikey).
  - **Models**: Thiết lập Model chính (`gemini-1.5-flash`) và 2 model dự phòng (`gemini-1.5-pro`, `gemini-2.0-flash`).
- **Provider 2 – Mistral**:
  - **Mistral API Key**: Nhập key thủ công hoặc khai báo biến môi trường `MISTRAL_API_KEY`.
  - **Models**: Thiết lập Model chính (`mistral-medium-latest`) và 2 model dự phòng (`mistral-small-latest`, `mistral-small-2409`).

> Cấu hình được lưu vào file `config.json` cục bộ (được tự động bỏ qua khi commit).

### 3. Chọn thư mục và chạy

1. Nhấn **Chọn...** bên cạnh "Thư mục Input" → chọn thư mục chứa các file `.xlsx` và `.rtf` của ngày phát sóng.
2. (Tùy chọn) Nhấn **Chọn...** bên cạnh "Thư mục Output" → nếu để trống, app tự tạo folder `output/` bên trong Input.
3. Nhập **Mã bản tin ($a090)** nếu cần thiết lập mã khác (mặc định hiển thị mờ gợi ý là `K303419`).
4. Nhấn **▶ BẮT ĐẦU BIÊN MỤC** và theo dõi tiến trình trong khung log.

---

## Quy định về thư mục Input & Định dạng

Để app có thể bóc tách chính xác, vui lòng chuẩn bị dữ liệu đầu vào đúng định dạng:

```
input/
├── BTTG24H_YYYYMMDD.xlsx        ← File danh sách bản tin (LIST) - BẮT BUỘC
├── NHUNG NGUOI THUC HIEN.rtf    ← File ê-kíp sản xuất (Bắt buộc hoặc dùng file tiền tố)
├── BGĐ MÃ CHÍ THÔNG.rtf         ← File ê-kíp bổ sung nếu NHUNG NGUOI THUC HIEN.rtf bị thiếu
├── BT KIM NGÂN, THẢO TRANG.rtf  ← File biên tập bổ sung (nếu có nhiều tên, cách bởi dấu phẩy)
├── 24H-TenTin.rtf               ← Kịch bản bản tin chính
├── GAT24H-TenTin.rtf            ← Kịch bản tin gạt
└── ...
```

- **File LIST (Excel)**: Tên file bắt đầu bằng `BTTG24H_YYYYMMDD.xlsx`. 
  - Dữ liệu được đọc từ **Active Sheet** (Sheet đang hoạt động).
  - Cột A phải chứa tên file bắt đầu bằng `24H-` hoặc `GAT24H-`.
  - Cột C là ID bản tin gồm **9 chữ số**.
  - Cột D ghi chữ `ONLINE`.
  - Cột F chứa thời lượng phát sóng.
- **File Ê-kíp**: File `NHUNG NGUOI THUC HIEN.rtf`. Nếu chức danh nào thiếu tên, app sẽ tự tìm các file RTF có tiền tố tương ứng (`BGĐ `, `BT `, `BD `, `MC `, `ĐD `, `KT `) trong folder để điền vào. Nếu một chức danh có từ 2 tên trở lên, các tên sẽ tự động được viết hoa và nối bằng dấu gạch ngang (` - `).
- **Quy tắc nhận diện Tiêu đề**: Tiêu đề chính là dòng chữ đầu tiên trong văn bản RTF thỏa mãn đồng thời:
  1. Được in đậm (`\b`).
  2. Được viết in HOA (`isupper()`).
  3. Có màu xanh lá cây (`\cf2` trong bảng màu).
  - *Nếu không tìm thấy dòng thỏa mãn định dạng trên, app sẽ tự động fallback sang tiêu đề do AI phân tích hoặc dòng viết hoa dài nhất trong văn bản.*
- **File Kịch bản tin (.rtf)**: Tên file cần chứa hoặc khớp với tên file định nghĩa trong cột A của file Excel LIST.

---

## File Output được tạo ra

Các file đầu ra luôn được định dạng với **font Times New Roman, cỡ chữ 11**. Ngoài ra, **cột B và C** của cả 3 file được tự động mở rộng gấp 3 lần chiều rộng mặc định (~30 ký tự) để hiển thị rõ nội dung:

| File | Mô tả |
|------|-------|
| `Import_SoLuoc_TG24H_YYYYMMDD.xlsx` | Danh sách sơ lược các bản tin, dùng để import vào hệ thống thư viện |
| `Map_BanTinTG_24G_YYYYMMDD.xlsx` | Bảng mapping thông tin ê-kíp và mục lục phát sóng (dùng mã bản tin $a090 được cung cấp) |
| `Map_ChiTiet_24G_YYYYMMDD.xlsx` | Nội dung chi tiết từng bản tin (người biên dịch + transcript) |

---

## Luồng xử lý & Cơ chế Dự phòng (Fallback)

```
File LIST (.xlsx)
   └── Đọc bằng openpyxl (không dùng AI)
       ├── Trích xuất ngày phát sóng
       ├── Lọc danh sách bản tin chính (24H-, GAT24H-, ONLINE, ID 9 số)
       └── Đọc thời lượng từ Cột F (format: 00:mm:ss)

NHUNG NGUOI THUC HIEN.rtf + Các file tiền tố
   └── Gọi AI (Gemini hoặc Mistral) / Quét file tiền tố → Trích xuất ê-kíp (JSON)

Mỗi file RTF tin chính (24H-*.rtf / GAT24H-*.rtf)
   ├── Tách và lấy tiêu đề trực tiếp bằng định dạng RTF (In đậm + In hoa + Màu xanh)
   ├── Cắt bỏ phần văn bản nhiễu (kể từ dòng chứa từ 3 ký tự '=' liên tiếp)
   └── Gửi văn bản đã xử lý qua AI (Gemini hoặc Mistral) theo cơ chế fallback:
        ├── 1. Gọi Model Chính (Gemini hoặc Mistral tương ứng) với timeout 4 phút.
        ├── 2. Dự phòng 1: Nếu lỗi, chuyển sang Model Dự phòng 1.
        ├── 3. Dự phòng 2: Nếu lỗi tiếp, chuyển sang Model Dự phòng 2.
        └── 4. Dự phòng không dùng AI: Nếu tất cả đều lỗi, tự động bóc tách bằng thuật toán
               (Lấy tất cả các đoạn văn bản phía dưới tiêu đề chính).

→ Tổng hợp → Sinh 3 file Excel output (Times New Roman, cỡ 11, tăng chiều rộng cột B và C)
```

---

## Quản lý Log & Gỡ lỗi

- **Log hiển thị trực tiếp**: Tiến trình chạy được hiển thị realtime trên giao diện app.
- **Tự động dọn dẹp log**: Toàn bộ nhật ký được ghi vào file `app_bien_muc_tg24h.log`. Để tránh phình to file log, app sẽ **tự động xóa log cũ và chỉ lưu lại dữ liệu của 3 ngày gần nhất** mỗi khi khởi động.
