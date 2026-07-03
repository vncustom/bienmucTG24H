# Ứng dụng Tự Động Biên Mục Bản Tin Thế Giới 24H (TG24H)

Ứng dụng Desktop giúp tự động hóa quy trình biên mục bản tin **Thế Giới 24H** của HTV. Hỗ trợ hai nhà cung cấp dịch vụ AI (**Provider 1** và **Provider 2**) để bóc tách nội dung từ các file kịch bản `.rtf` và sinh ra 3 file Excel đầu ra theo chuẩn biên mục của Trung tâm Tư liệu HTV.

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
  - **Provider 1 (Mặc định)**: Sử dụng nhà cung cấp dịch vụ thứ nhất.
  - **Provider 2**: Sử dụng nhà cung cấp dịch vụ thứ hai.
- **Provider 1**:
  - **API Key Provider 1**: Điền API key tương ứng.
- **Provider 2**:
  - **API Key Provider 2**: Điền API key tương ứng (hoặc khai báo biến môi trường).

> Cấu hình được lưu vào file `config.json` cục bộ (được tự động bỏ qua khi commit).

### 3. Chọn thư mục và chạy

1. Nhấn **Chọn...** bên cạnh "Thư mục Input" → chọn thư mục chứa các file `.xlsx` và `.rtf` của ngày phát sóng.
2. (Tùy chọn) Nhấn **Chọn...** bên cạnh "Thư mục Output" → nếu để trống, app tự tạo folder `output/` bên trong Input.
3. Nhập **Mã bản tin ($a090)** nếu cần thiết lập mã khác (mặc định hiển thị mờ gợi ý là `K303419`).
4. Nhấn **▶ BẮT ĐẦU BIÊN MỤC** và theo dõi tiến trình trong khung log.
5. Khi hoàn tất, popup **Thành công** hiển thị đường dẫn output. Nhấn **Open output folder** để mở nhanh thư mục chứa file kết quả, hoặc nhấn **OK** để đóng popup.

---

## Quy định về thư mục Input & Định dạng

Để app có thể bóc tách chính xác, vui lòng chuẩn bị dữ liệu đầu vào đúng định dạng:

```
input/
├── BTTG24H_YYYYMMDD.xlsx        ← File danh sách bản tin (LIST) - BẮT BUỘC
├── NHUNG NGUOI THUC HIEN*.rtf   ← File ê-kíp sản xuất (Bắt buộc hoặc dùng file tiền tố)
├── BGĐ MÃ CHÍ THÔNG.rtf         ← File ê-kíp bổ sung nếu file NHUNG NGUOI THUC HIEN*.rtf bị thiếu
├── BT KIM NGÂN, THẢO TRANG.rtf  ← File biên tập bổ sung (nếu có nhiều tên, cách bởi dấu phẩy)
├── 24H-TenTin.rtf               ← Kịch bản bản tin chính
├── GAT24H-TenTin.rtf            ← Kịch bản tin gạt
├── GAT TenTin.rtf               ← Kịch bản tin gạt
└── ...
```

- **File LIST (Excel)**: Tên file bắt đầu bằng `BTTG24H_YYYYMMDD.xlsx`. 
  - Dữ liệu được đọc từ **Active Sheet** (Sheet đang hoạt động).
  - Cột A phải chứa tên file bắt đầu bằng `24H-`, `24h-`, `24 `, `GAT24H`, `GAT24h` hoặc `GAT `.
  - Cột C là ID bản tin và phải **bắt đầu bằng số**. Ví dụ `260611056` và `260611056a` hợp lệ; dòng trống hoặc `qc123` không hợp lệ.
  - Cột D không dùng để bắt tin.
  - Cột F chứa thời lượng phát sóng.
- **File Ê-kíp**: File `.rtf` có tên chứa `NHUNG NGUOI THUC HIEN`, ví dụ `NHUNG NGUOI THUC HIEN.rtf` hoặc `NHUNG NGUOI THUC HIEN abc.rtf`. Nếu chức danh nào thiếu tên, app sẽ tự tìm các file RTF có tiền tố tương ứng (`BGĐ `, `BT `, `BD `, `MC `, `ĐD `, `KT `) trong folder để điền vào. Nếu một chức danh có từ 2 tên trở lên, các tên sẽ tự động được viết hoa và nối bằng dấu gạch ngang (` - `).
- **Quy tắc nhận diện Tiêu đề**: Tiêu đề chính là dòng đầu tiên trong phần đầu kịch bản thỏa mãn đồng thời:
  1. Được viết IN HOA toàn bộ (`isupper()`).
  2. Được in đậm (BOLD).
  3. Dài hơn 16 ký tự.
  4. Không bắt đầu bằng `AFP`, `AP` hoặc `REUTERS`.
  5. Không phải nhãn phân loại như `GẠT TG24H`, `GAT24H` hoặc `HEADLINES`.
  - Quy tắc này được áp dụng cho cả kết quả AI, fallback và thuật toán nội bộ.
- **File Kịch bản tin (.rtf)**: Tên file cần chứa hoặc khớp với tên file định nghĩa trong cột A của file Excel LIST.

---

## File Output được tạo ra

Các file đầu ra luôn được định dạng với **font Times New Roman, cỡ chữ 11**. Ngoài ra, **cột B và C** của các file được tự động mở rộng gấp 3 lần chiều rộng mặc định (~30 ký tự) để hiển thị rõ nội dung:

### 1. Thư mục Output chính (Sử dụng AI kết hợp Định dạng/Thuật toán)
| File | Mô tả |
|------|-------|
| `Import_SoLuoc_TG24H_YYYYMMDD.xlsx` | Danh sách sơ lược các bản tin, dùng để import vào hệ thống thư viện |
| `Map_BanTinTG_24G_YYYYMMDD.xlsx` | Bảng mapping thông tin ê-kíp và mục lục phát sóng (dùng mã bản tin $a090) |
| `Map_ChiTiet_24G_YYYYMMDD.xlsx` | Nội dung chi tiết từng bản tin (người biên dịch + transcript). Các cụm `PB[số]` + 2 dòng IN HOA liên tiếp được gộp thành một dòng `DÒNG 1 - DÒNG 2`. |

### 2. Thư mục `tempbienmuc` (Sử dụng Thuật toán nội bộ - Không dùng AI)
Thư mục này nằm cùng cấp với thư mục `input` và được tự động **dọn dẹp sạch sẽ (xóa toàn bộ file cũ)** mỗi khi bắt đầu tiến trình biên mục. Các file output dự phòng tại đây bao gồm:
| File | Mô tả |
|------|-------|
| `Import_SoLuoc_TG24H_thuattoan_YYYYMMDD.xlsx` | Danh sách sơ lược các bản tin, tạo hoàn toàn bằng thuật toán bóc tách tiêu đề nội bộ |
| `Map_BanTinTG_24G_thuattoan_YYYYMMDD.xlsx` | Bảng mapping ê-kíp (bóc tách bằng regex) và mục lục tạo hoàn toàn bằng thuật toán |
| `Map_ChiTiet_ThuatToan_YYYYMMDD.xlsx` | Nội dung chi tiết từng bản tin bóc tách bằng thuật toán (người biên dịch nội bộ + transcript). Áp dụng cùng quy tắc gộp cụm `PB[số]` như file Map_ChiTiet AI. |

---

## Luồng xử lý & Thuật toán nội bộ (Không dùng AI)

Ứng dụng kết hợp bóc tách bằng Trí tuệ nhân tạo (AI) và Thuật toán nội bộ chạy độc lập để đối chiếu chéo kết quả:

### 1. Thuật toán bóc tách Ê-kíp nội bộ
- **Nguồn**: File `NHUNG NGUOI THUC HIEN.rtf`
- **Cơ chế**: Quét văn bản đã chuyển đổi sang plain text bằng các mẫu biểu thức chính quy (Regex) tương ứng với từng chức danh.
- **Fallback**: Nếu thiếu chức danh nào, tự động dò tìm các file RTF phụ có tiền tố tương ứng (`BGĐ `, `BT `, `BD `, `MC `, `ĐD `, `KT `) để lấy tên người thực hiện từ tên file.

### 2. Thuật toán bóc tách Tiêu đề, Biên dịch và Nội dung tin
- **Tiêu đề**: Dòng đầu tiên trong phần đầu kịch bản được viết IN HOA toàn bộ, in đậm (BOLD), dài hơn 16 ký tự, không bắt đầu bằng `AFP`, `AP`, `REUTERS`, và không phải nhãn gạt/headlines.
- **Biên dịch**: Dòng chữ không chứa chữ số hoặc các từ khóa của metadata nằm ngay trước dòng tiêu đề trong phạm vi 7 dòng.
- **Nội dung tin**: Lọc bỏ các dòng nhiễu (dòng separator `==`, tên tiếng Anh, link hình/video, ngày tháng).
  - Với tin thường: Lấy từ dưới tiêu đề đến chữ **màu đen cuối cùng**.
  - Với tin LIVE (tên file chứa chữ "LIVE"): Lấy từ dưới tiêu đề đến chữ **màu đỏ cuối cùng**, đồng thời tự động loại bỏ các dòng chữ màu xanh lá cây không in đậm.
  - Khi xuất `Map_ChiTiet`, nếu gặp cụm 3 dòng IN HOA liên tục theo mẫu `PB[số]`, `IN HOA 1`, `IN HOA 2`, app bỏ dòng `PB[số]` và gộp 2 dòng sau thành `IN HOA 1 - IN HOA 2`.

### 3. Cơ chế đối chiếu & Hiển thị tiến trình
- **Đối chiếu chéo**: Sau khi sinh các file, app so sánh số dòng của từng bản tin giữa `Map_ChiTiet` (AI) và `Map_ChiTiet_ThuatToan`.
- **Thông báo**: Nếu phát hiện sự chênh lệch, app ghi nhận nhẹ nhàng trong khung log tiến trình:
  `Phát hiện số dòng không trùng khớp giữa AI và ThuatToan:`
  `  ▸ Tin {ID}: AI={ai} dòng, ThuậtToán={tt} dòng`
  `Đề nghị kiểm tay những tin trên.`
  *(Không hiển thị popup cảnh báo làm phiền người dùng)*.
- **Thanh tiến độ**: Được thiết kế chạy realtime thread-safe mượt mà từ 0% đến 100% giúp giao diện không bị giật hoặc treo.
- **Popup hoàn tất**: Sau khi tạo file thành công, popup hiển thị đường dẫn output và có nút **Open output folder** để mở nhanh thư mục kết quả.

---

## Quản lý Log & Gỡ lỗi

- **Log hiển thị trực tiếp**: Tiến trình chạy được hiển thị realtime trên giao diện app.
- **Tự động dọn dẹp log**: Toàn bộ nhật ký được ghi vào file `app_bien_muc_tg24h.log`. Để tránh phình to file log, app sẽ **tự động xóa log cũ và chỉ lưu lại dữ liệu của 3 ngày gần nhất** mỗi khi khởi động.
