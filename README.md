# 🔬 Hệ thống Đánh giá Lot-to-Lot (CLSI EP26-Ed2)

Dự án ứng dụng web quản lý chất lượng phòng xét nghiệm Y khoa, được xây dựng trên nền tảng **Streamlit**, chuyên dùng để đánh giá độ tương đồng giữa hai lô hóa chất (Lot-to-Lot) tuân thủ theo tiêu chuẩn quốc tế **CLSI EP26-Ed2**.

---

## 🌟 Tính năng nổi bật

Ứng dụng được chia thành 5 phân hệ (Tabs) chính giúp tối ưu hóa luồng công việc của nhân viên quản lý chất lượng (QCM):

1. **🧪 Đánh giá Mẫu thực tế (Stage 2):** Nhập kết quả chạy mẫu trên 2 lô (LOT cũ & LOT mới), tự động tính toán độ chệch tuyệt đối ($d$), trung bình phần trăm độ chệch, và đối chiếu với Giới hạn từ chối (Rejection Limit - RL) để kết luận ĐẠT/TRƯỢT.
2. **🕒 Lịch sử Đánh giá:** Lưu trữ toàn bộ lịch sử các lần thay LOT, phân loại rành mạch theo Tháng/Năm. Hỗ trợ **xuất biên bản tự động ra file Word (.docx) và PDF** dựa trên biểu mẫu chuẩn của khoa.
3. **📊 Tổng quan LOT (Dashboard):** Quản lý tình trạng hóa chất hiện tại của từng hệ thống máy (Sinh hóa, Miễn dịch, Huyết học, HbA1c,...). Theo dõi mã LOT đang sử dụng và lịch sử thay đổi của riêng từng loại hóa chất.
4. **⚙️ Cấu hình Thông số (Stage 1):** Thiết lập thông số ban đầu (Target, TEa, CVwrl, CVr). Engine thuật toán sẽ tự động tra bảng xác định số lượng mẫu ($n$) cần chạy, hằng số $K$, mức công suất thống kê (Power) và Giới hạn từ chối (RL).
5. **📂 Quản lý Dữ liệu DB:** Sao lưu và phục hồi dữ liệu thông qua thao tác Import file Excel/CSV nhanh chóng.

---

## 🏗 Cấu trúc mã nguồn (Architecture)

Dự án được thiết kế theo mô hình phân tách Module (Modular Design) để dễ dàng bảo trì và mở rộng:

```text
lot-to-lot/
│
├── .streamlit/
│   └── config.toml          # Cấu hình giao diện (Light mode, màu sắc, font chữ)
│
├── app.py                   # File giao diện chính chạy Streamlit (Chứa 5 Tabs)
├── database.py              # Xử lý đọc/ghi cơ sở dữ liệu (Google Sheets) & Cache dữ liệu
├── connection.py            # Rẽ nhánh kết nối Google Cloud (Streamlit Secrets & Local JSON)
├── clsi_engine.py           # Core logic: Thuật toán thống kê y khoa chuẩn CLSI EP26
├── report_generator.py      # Module chèn dữ liệu tự động vào template Word (.docx)
├── template.docx            # File Word biểu mẫu gốc của Bệnh viện / Khoa Xét nghiệm
├── requirements.txt         # Danh sách các thư viện Python (Dependencies)
└── README.md                # Tài liệu mô tả dự án (File này)
