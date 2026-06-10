# Ứng dụng Tự Động Biên Mục Bản Tin Thế Giới 24H (TG24H)

Ứng dụng Desktop giúp tự động hóa quy trình biên mục bản tin **Thế Giới 24H** của HTV. Sử dụng **Gemini API** để bóc tách nội dung từ các file kịch bản `.rtf` và sinh ra 3 file Excel đầu ra theo chuẩn biên mục của Trung tâm Tư liệu HTV.

---

## Yêu cầu hệ thống

- **Python** 3.10 trở lên
- Các thư viện Python (cài bằng lệnh bên dưới):

```bash
pip install openpyxl striprtf google-generativeai pydantic
```

---

## Cách sử dụng

### 1. Chạy ứng dụng

```bash
python app_bien_muc_tg24h.py
```

### 2. Cấu hình API Key & Models (lần đầu)

Nhấn nút **⚙ Cài đặt API** và điền:
- **Gemini API Key**: Lấy tại [aistudio.google.com](https://aistudio.google.com/apikey)
- **Model Chính**: Mặc định `gemini-1.5-flash`.
- **Model Dự phòng 1**: Mặc định `gemini-1.5-pro`.
- **Model Dự phòng 2**: Mặc định `gemini-2.0-flash`.

> Nếu bạn đã cài biến môi trường `GEMINI_API_KEY` trong hệ điều hành, app sẽ tự nhận diện, không cần nhập tay.
> Cấu hình được lưu vào file `config.json` (không commit lên git).

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
- **File Kịch bản tin (.rtf)**: Tên file cần chứa hoặc khớp với tên file định nghĩa trong cột A của file Excel LIST.

---

## File Output được tạo ra

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
   └── Gọi Gemini API / Quét file tiền tố → Trích xuất chức danh & tên ê-kíp (JSON)

Mỗi file RTF tin chính (24H-*.rtf / GAT24H-*.rtf)
   └── Phân tích văn bản: Cắt bỏ toàn bộ nội dung từ dòng chứa từ 3 ký tự '=' liên tiếp trở đi.
   └── Gọi Gemini API theo Cơ chế Primary Fallback:
       ├── 1. Thử bằng Model Chính (Ví dụ: gemini-1.5-flash) với timeout 4 phút.
       ├── 2. Nếu thất bại/timeout/rỗng, chuyển sang Model Dự phòng 1 (Ví dụ: gemini-1.5-pro).
       ├── 3. Nếu tiếp tục thất bại, chuyển sang Model Dự phòng 2 (Ví dụ: gemini-2.0-flash).
       └── 4. Nếu tất cả model đều thất bại, tự động chuyển sang Fallback không dùng AI:
              Trích xuất toàn bộ văn bản phía dưới tiêu đề (bỏ dòng trống và dòng quá ngắn).

→ Tổng hợp → Sinh 3 file Excel output (font Times New Roman, cỡ chữ 11)
```

---

## Quản lý Log & Gỡ lỗi

- **Log hiển thị trực tiếp**: Tiến trình chạy được hiển thị realtime trên giao diện app.
- **Tự động dọn dẹp log**: Toàn bộ nhật ký được ghi vào file `app_bien_muc_tg24h.log`. Để tránh phình to file log, app sẽ **tự động xóa log cũ và chỉ lưu lại dữ liệu của 3 ngày gần nhất** mỗi khi khởi động.

---

## Ghi chú kỹ thuật

- **Thời lượng**: Excel lưu giá trị dạng `HH:MM` nhưng thực chất là `mm:ss` (ví dụ `01:08` = 1 phút 8 giây). App xử lý chính xác bằng thuật toán nội bộ.
- **Giới hạn thời gian (Timeout)**: Mỗi lượt gọi API giới hạn tối đa **4 phút**. Quá 4 phút sẽ tự động ngắt và chuyển sang model dự phòng tiếp theo.
- **Cắt bỏ footer**: Gặp dòng chứa từ 3 dấu `=` liên tiếp trở lên (ví dụ `====`), app sẽ bỏ qua toàn bộ phần văn bản phía sau dòng này trước khi gửi cho AI/thuật toán để tránh nhiễu thông tin.
