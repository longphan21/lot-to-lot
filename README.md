# 🔬 CLSI EP26-Ed2 — Lot-to-Lot Evaluation System

Ứng dụng web xây dựng bằng **Streamlit** hỗ trợ phòng xét nghiệm y khoa chuẩn hóa và tự động hóa quy trình **đánh giá độ tương đồng giữa các lô hóa chất (Reagent Lot-to-Lot Evaluation)** theo hướng dẫn **CLSI EP26-Ed2**, tích hợp cơ sở dữ liệu đám mây **Google Sheets**, xuất biên bản **Word (`.docx`) / Google Drive** và tối ưu hóa sử dụng trên cả **Máy tính (Desktop)** lẫn **Điện thoại (Mobile / PWA)**.

* **Ứng dụng chính (Streamlit Cloud):** [lot-to-lot-xn2o6yh7yyzpwnkaha63u6.streamlit.app](https://lot-to-lot-xn2o6yh7yyzpwnkaha63u6.streamlit.app/)
* **Cài đặt App trên điện thoại (PWA — Tên `Lot to Lot` & Icon riêng):** [longphan21.github.io/lot-to-lot-pwa](https://longphan21.github.io/lot-to-lot-pwa/)

---

## ✨ Tính năng nổi bật

### 1. Đánh giá Mẫu thực tế — Stage 2 (`1. Lot Evaluation`)
* **Khóa thông tin từng bước (Step-by-Step Locking):** Nút **Khóa / Mở khóa Thông tin** giúp cố định 6 trường đợt đánh giá (Xét nghiệm, Dòng máy, Ngày thực hiện, Loại hóa chất, LOT cũ, LOT mới) để tập trung nhập kết quả mẫu mà không lo chọn nhầm.
* **Giữ bền vững dữ liệu (State Persistence):** Toàn bộ thông tin đợt chạy và dữ liệu cặp mẫu đã gõ được lưu bền vững trong phiên làm việc — không bị mất khi bấm lưu, đóng thông báo hoặc chuyển sang tab khác để xuất biên bản. Đặc biệt tiện lợi cho **Huyết học / HbA1c** khi nhiều chỉ tiêu dùng chung một lô hóa chất.
* **Bảng nhập liệu tối ưu cho Mobile & Desktop:** Bỏ cột thừa, rút gọn tiêu đề (`SID`, `LOT cũ`, `LOT mới`), ẩn thanh công cụ nổi che chữ, không cần cuộn ngang màn hình; tích hợp nút **Thêm dòng**, **Bớt dòng** và **Làm trống bảng mẫu** cho đợt mới.
* **Tính toán tức thời:** Tự động tính độ chệch từng mẫu $|d|$, trung bình độ chệch $\overline{|d|}$, phần trăm độ chệch trung bình ($\overline{|d|}\%$) và đối chiếu với giới hạn từ chối ($RL\%$) để kết luận **Đạt (PASS)** hoặc **Không đạt (FAIL)**.

### 2. Lịch sử Đánh giá & Theo dõi Trình ký (`2. Evaluation History`)
* **Quản lý trạng thái trình ký hàng tuần:** Tích hợp **ô check Đã ký** nằm ngay bên trái mỗi đợt đánh giá, cho phép đánh dấu nhanh các biên bản đã đưa lãnh đạo ký (lưu trực tiếp lên cột `signed` trong Google Sheets và hiện thông báo Toast gọn nhẹ).
* **Bộ lọc & Tra cứu đa từ khóa:** Lọc nhanh theo trạng thái **Tất cả / Chưa ký / Đã ký** kết hợp thanh tìm kiếm thông minh theo Ngày, Tên XN, Máy, Hóa chất, số LOT, mã mẫu (SID) hoặc trạng thái (`đạt`, `chưa ký`, `đã ký`).
* **Xuất biên bản chuẩn & Đẩy Google Drive:**
  * Tự động tạo biên bản Word (`.docx`) theo chuẩn đặt tên: `Tên XN - Tên máy - Loại hóa chất - Ngày đánh giá.docx`.
  * Hỗ trợ **Tải về máy**, **Up trực tiếp lên thư mục Google Drive** (qua Google Apps Script) và chuyển đổi sang **PDF** (khi môi trường có LibreOffice).
* **Xóa an toàn (Soft Delete):** Xác nhận qua hộp thoại Pop-up và chuyển dữ liệu vào Thùng rác thay vì xóa mất ngay.

### 3. Tổng quan LOT đang sử dụng (`3. Lot Overview`)
* Tự động phân tách theo 2 phân hệ: **Sinh hóa - Miễn dịch** và **Huyết học - HbA1c**.
* Tra cứu tức thì **số LOT mới nhất đang chạy** trên từng dòng máy, ngày cập nhật gần nhất, trạng thái đạt/không đạt và bấm mở rộng để xem toàn bộ lịch sử thay LOT trước đó.

### 4. Cấu hình Thông số — Stage 1 (`4. Parameter Setup`)
* Nhập các thông số chất lượng cho từng xét nghiệm và thiết bị: nồng độ đích (*Target concentration*), $TEa\%$, $S_{wrl}\%$ ($CV_{wrl}\%$), $S_r\%$ ($CV_r\%$), mã thiết bị và đơn vị đo.
* Tự động tra bảng **CLSI EP26-Ed2** (Bảng A1, A2, A3 tùy theo số mức nồng độ được chọn) để tính toán tỷ số $CD / S_{wrl}$, $S_r / S_{wrl}$, lực thống kê ($Power$), hệ số $K$, số mẫu bệnh nhân tối thiểu ($n$) và giới hạn từ chối ($RL\%$).

### 5. Danh mục Xét nghiệm đã lưu (`5. Test Directory`)
* Xem tổng thể toàn bộ thông số Stage 1 đã cấu hình.
* **Bật / Tắt trạng thái sử dụng (`Active` / `Inactive`):** Chủ động ẩn tạm các mức nồng độ hoặc dòng máy không còn sử dụng mà không cần xóa mất dữ liệu lịch sử.
* **Sao lưu & Khôi phục:** Hỗ trợ **Export CSV** và **Import hàng loạt** từ file `.csv` / `.xlsx`.

### 6. Thùng rác độc lập (`6. Trash`)
* Tách riêng thành một trang trên **Left Sidebar** kèm bộ đếm số đợt đang lưu trong thùng rác.
* Hỗ trợ tìm kiếm nhanh, **Khôi phục** biên bản về lại mục `2. Evaluation History` hoặc **Xóa vĩnh viễn**.

### 7. Tự động nhận diện Giao diện Máy tính / Điện thoại & Cài đặt PWA
* **Chung 1 đường link duy nhất:** Tự động phát hiện thiết bị qua `User-Agent` (cùng tùy chọn chuyển tay trong Sidebar hoặc tham số `?view=mobile` / `?view=desktop`):
  * **Desktop:** Sidebar mở rộng, bố cục ngang nhiều cột và bảng chi tiết 10 cột.
  * **Mobile:** Thanh điều hướng nhanh trên đầu màn hình, thẻ dọc (*Card UI*) cho Lịch sử, Tổng quan LOT, Danh mục XN và Thùng rác.
* **Chống lỗi định dạng Google Sheets:** Toàn bộ thao tác đọc/ghi sử dụng chế độ `RAW`, ngăn Google Sheets tự biến đổi mã LOT chứa chữ `E` thành ký hiệu khoa học hoặc làm mất số `0` đầu mã SID/LOT.

---

## 🗂️ Cấu trúc thư mục dự án

```text
.
├── app.py                # Giao diện chính (6 trang), responsive Desktop/Mobile, Locking & Persistence
├── clsi_engine.py        # Động cơ tính toán & tra bảng CLSI EP26-Ed2 (Stage 1 & Stage 2)
├── database.py           # Thao tác Google Sheets (RAW mode): stage1_setup, stage2_evaluations, stage2_trash
├── connection.py         # Kết nối Google Sheets API thông qua Service Account
├── report_generator.py   # Tạo biên bản báo cáo định dạng Word (.docx)
├── icon.png              # Biểu tượng ứng dụng màu xanh lá tối giản
├── requirements.txt      # Danh sách thư viện Python phụ thuộc
└── README.md             # Tài liệu mô tả dự án
```

---

## 🗄️ Cấu trúc cơ sở dữ liệu (Google Sheets)

Ứng dụng tự động khởi tạo và quản lý 3 trang tính (worksheets) trong cùng một file Google Sheets:

1. **`stage1_setup`**: Lưu cấu hình thông số theo từng cặp `(test_name, machine)` gồm `target_conc`, `tea_pct`, `swrl_pct`, `sr_pct`, `cd_swrl`, `sr_swrl`, `power_achieve`, `n_samples`, `rejection_limit`, `k_factor`, `is_selected` (`1` = Active, `0` = Inactive).
2. **`stage2_evaluations`**: Lưu chi tiết các đợt đánh giá thực tế (`eval_date`, `lot_old`, `lot_new`, `reagent_name`, `sid`, `val_old`, `val_new`, `diff_d`, `mean_d`, `mean_d_pct`, `rl_pct`, `status`, `created_at`, `signed`).
3. **`stage2_trash`**: Lưu trữ tạm các đợt đánh giá đã xóa mềm để phục vụ khôi phục hoặc xóa vĩnh viễn.

---

## ⚙️ Hướng dẫn cài đặt & chạy cục bộ

### 1. Yêu cầu môi trường
* Python **3.9** trở lên.
* Google Cloud Project đã bật **Google Sheets API** và **Google Drive API**, có tài khoản **Service Account** được cấp quyền **Editor** trên file Google Sheets lưu dữ liệu.
* *(Tùy chọn)* Cài đặt **LibreOffice** nếu muốn sử dụng tính năng xuất trực tiếp sang PDF ở môi trường cục bộ.

### 2. Cài đặt thư viện
```bash
git clone <url-repository-cua-ban>
cd <thu-muc-du-an>
pip install -r requirements.txt
```

### 3. Cấu hình Secrets & Google Drive
Tạo file `.streamlit/secrets.toml` tại thư mục gốc dự án theo đúng khóa sử dụng trong `connection.py`:

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "your-private-key-id"
private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
client_email = "your-service-account@your-project.iam.gserviceaccount.com"
client_id = "your-client-id"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://accounts.google.com/o/oauth2/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "your-cert-url"
```

* Cập nhật `DRIVE_FOLDER_URL` và `APPS_SCRIPT_URL` ở đầu file `app.py` nếu muốn trỏ sang thư mục Google Drive lưu biên bản Word khác.

### 4. Khởi chạy ứng dụng
```bash
streamlit run app.py
```

---

## 📱 Cài đặt thành App trên điện thoại (PWA)

Để mở ứng dụng toàn màn hình với đúng tên **Lot to Lot** và **icon màu xanh lá** (không bị dính logo mặc định của Streamlit Cloud):

1. Mở link PWA trên trình duyệt **Chrome** (Android) hoặc **Safari** (iOS):
   `https://longphan21.github.io/lot-to-lot-pwa/`
2. Chọn menu trình duyệt $\rightarrow$ **Thêm vào màn hình chính (Add to Home screen)** $\rightarrow$ **Cài đặt ứng dụng (Install app)**.
3. Mọi cập nhật code tại repository chính này sẽ tự động hiển thị trên app điện thoại mà không cần cài lại.

---

## 📚 Tài liệu tham chiếu

* **CLSI EP26-Ed2:** *User Evaluation of Acceptability of a Reagent Lot Change; Approved Guideline — Second Edition*.
* **ISO 15189:** *Medical laboratories — Requirements for quality and competence*.