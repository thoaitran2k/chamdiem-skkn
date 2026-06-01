# 📋 Ứng dụng Chấm Điểm SKKN – UBND xã Lấp Vò

## Giới thiệu
Ứng dụng web (Streamlit) giúp chấm điểm hàng loạt file Sáng Kiến Kinh Nghiệm (SKKN) 
bằng Claude AI (Haiku) theo đúng barem của UBND xã Lấp Vò, tỉnh Đồng Tháp.

## Tính năng
- 📤 Upload 1 hoặc nhiều file `.docx` SKKN cùng lúc
- 🔍 Tự động trích xuất: Tên tác giả, đơn vị, chức danh, trình độ, tên đề tài, ngày áp dụng
- 💾 Lưu trữ vào SQLite database (không cần cài server)
- 🤖 Chấm điểm tự động bằng Claude Haiku theo 3 tiêu chí:
  - Tính mới, tính sáng tạo (30 điểm)
  - Khả năng áp dụng, nhân rộng (30 điểm)
  - Hiệu quả (40 điểm)
- 💬 Nhận xét từng tiêu chí + nhận xét chung + kiến nghị + gợi ý cải thiện
- 📊 Thống kê và xuất bảng điểm CSV

---

## Cài đặt

### 1. Yêu cầu
- Python 3.9+ (khuyến nghị 3.11)
- Anthropic API Key (có credit Claude)

### 2. Cài thư viện
```bash
pip install -r requirements.txt
```

### 3. Chạy ứng dụng
```bash
# Cách 1: Đặt API key vào biến môi trường (bảo mật hơn)
export ANTHROPIC_API_KEY="sk-ant-xxxxxxxx"
streamlit run skkn_app.py

# Cách 2: Nhập API key trực tiếp trên giao diện web
streamlit run skkn_app.py
```

Mở trình duyệt tại: http://localhost:8501

---

## Hướng dẫn sử dụng

### Bước 1: Nạp File SKKN
1. Vào tab **📤 Nạp File SKKN**
2. Kéo thả hoặc nhấn Browse để chọn file `.docx` (chọn nhiều file cùng lúc)
3. Nhấn **"Nạp vào hệ thống"**
4. Hệ thống tự động đọc và điền: tên tác giả, đơn vị, đề tài, ngày áp dụng...

### Bước 2: Chấm Điểm
1. Vào tab **🤖 Chấm Điểm AI**
2. Nhập **Anthropic API Key** (hoặc đã đặt biến môi trường)
3. Chọn:
   - **"Chấm theo lô"**: chấm tất cả SKKN chưa chấm 1 lần
   - **"Chấm từng bài"**: chọn từng SKKN, chấm riêng lẻ
4. Nhấn nút và đợi Claude phân tích (mỗi bài ~10-30 giây)

### Bước 3: Xem kết quả
1. Vào **📋 Danh sách SKKN** để xem tất cả
2. Lọc theo trạng thái (Đã chấm/Chưa chấm) hoặc kết quả (Đạt/Không đạt)
3. Vào **📊 Thống kê** để xem tổng quan và tải CSV

---

## Barem chấm điểm (theo UBND xã Lấp Vò)

| Tiêu chí | Điểm tối đa | Điểm tối thiểu |
|----------|-------------|----------------|
| Tính mới, sáng tạo | 30 | 21 |
| Khả năng áp dụng, nhân rộng | 30 | 21 |
| Hiệu quả | 40 | 25 |
| **Tổng** | **100** | **70** |

Sáng kiến **ĐẠT** khi: Tổng ≥ 70 điểm **VÀ** từng tiêu chí đạt điểm tối thiểu.

---

## Lưu ý kỹ thuật
- Database: `skkn_database.db` (SQLite, tự động tạo)
- Model AI: `claude-haiku-4-5-20251001` (nhanh, tiết kiệm credit)
- Mỗi lần chấm: ~1,000-3,000 tokens (~$0.001-0.003 USD/bài)
- Hỗ trợ file có STT hoặc không có STT trong bảng tác giả
- Tên đề tài được nhận dạng cả trong ngoặc kép, in đậm, in nghiêng, xuống hàng

---

## Cấu trúc file
```
skkn_app.py          ← File chính, chạy lệnh này
requirements.txt     ← Thư viện cần cài
skkn_database.db     ← Database (tự tạo khi chạy lần đầu)
```
