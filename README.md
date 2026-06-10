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

### 2. Cấu hình API Key (lần đầu)

Nhấn nút **⚙ Cài đặt API** và điền:
- **Gemini API Key**: Lấy tại [aistudio.google.com](https://aistudio.google.com/apikey)
- **Model Name**: Mặc định `gemini-1.5-flash`. Có thể đổi sang `gemini-1.5-pro`, `gemma-4-26b-a4b-it`, v.v.

> Nếu bạn đã cài biến môi trường `GEMINI_API_KEY` trong hệ điều hành, app sẽ tự nhận diện, không cần nhập tay.

Cấu hình được lưu vào file `config.json` (không commit lên git).

### 3. Chọn thư mục và chạy

1. Nhấn **Chọn...** bên cạnh "Thư mục Input" → chọn thư mục chứa các file `.xlsx` và `.rtf` của ngày phát sóng.
2. (Tùy chọn) Nhấn **Chọn...** bên cạnh "Thư mục Output" → nếu để trống, app tự tạo folder `output/` bên trong Input.
3. Nhấn **▶ BẮT ĐẦU BIÊN MỤC** và theo dõi tiến trình trong khung log.

---

## Cấu trúc thư mục Input

```
input/
├── BTTG24H_YYYYMMDD.xlsx        ← File danh sách bản tin (LIST) - BẮT BUỘC
├── NHUNG NGUOI THUC HIEN.rtf    ← File ê-kíp sản xuất
├── 24H-TenTin.rtf               ← Kịch bản bản tin chính
├── GAT24H-TenTin.rtf            ← Kịch bản tin gạt
└── ...
```

> **Lưu ý:** File LIST phải đặt tên theo format `BTTG24H_YYYYMMDD.xlsx` (ví dụ: `BTTG24H_20260609.xlsx`).

---

## File Output được tạo ra

| File | Mô tả |
|------|-------|
| `Import_SoLuoc_TG24H_YYYYMMDD.xlsx` | Danh sách sơ lược các bản tin, dùng để import vào hệ thống thư viện |
| `Map_BanTinTG_24G_YYYYMMDD.xlsx` | Bảng mapping thông tin ê-kíp và mục lục phát sóng |
| `Map_ChiTiet_24G_YYYYMMDD.xlsx` | Nội dung chi tiết từng bản tin (người biên dịch + transcript) |

---

## Luồng xử lý

```
File LIST (.xlsx)
   └── Đọc bằng openpyxl (không dùng AI)
       ├── Trích xuất ngày phát sóng
       ├── Lọc danh sách bản tin chính (24H-, GAT24H-, ONLINE, ID 9 số)
       └── Đọc thời lượng từ Cột F (format: 00:mm:ss)

NHUNG NGUOI THUC HIEN.rtf
   └── Gọi Gemini API → Trích xuất chức danh & tên ê-kíp (JSON)

Mỗi file RTF tin chính (24H-*.rtf / GAT24H-*.rtf)
   └── Gọi Gemini API (cuốn chiếu từng file) → Trích xuất:
       ├── Tiêu đề (dòng HOA đầu tiên, có fallback tìm kiếm thuật toán)
       ├── Tên người biên dịch (nếu có, để trống nếu không rõ)
       └── Nội dung (phụ đề + lời đọc voiceover)

→ Tổng hợp → Sinh 3 file Excel output (font Times New Roman)
```

---

## Debug lỗi

- **Log trên màn hình**: Khung log trong app hiển thị tiến trình realtime.
- **File log**: Toàn bộ log được ghi vào `app_bien_muc_tg24h.log` cùng thư mục chạy app.
- **Gửi cho AI debug**: Copy nội dung file `.log` và gửi kèm mô tả lỗi cho agent AI.

---

## Ghi chú kỹ thuật

- **Thời lượng**: Excel lưu giá trị dạng `HH:MM` nhưng thực chất là `mm:ss` (ví dụ `01:08` = 1 phút 8 giây). App xử lý chính xác bằng thuật toán nội bộ.
- **Retry API**: Mỗi lần gọi API tự động retry tối đa 3 lần nếu bị lỗi mạng (503/504).
- **Fallback tiêu đề**: Nếu AI không bóc được tiêu đề, app tự tìm dòng VIẾT HOA đầu tiên trong văn bản.
- **Font output**: Tất cả file Excel đầu ra dùng font **Times New Roman** size 11.
