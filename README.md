# A4 – Tự động gửi email từ Excel/CSV

Pipeline: **Excel/CSV → Lọc người nhận → Sinh nội dung → Gửi email → Log kết quả**

## Cài đặt
```bash
pip install openpyxl          # chỉ cần nếu đọc file .xlsx (CSV không cần gì thêm)
cp .env.example .env          # điền thông tin SMTP (Gmail: dùng "Mật khẩu ứng dụng")
```

## Cấu trúc
```
email_sender.py        chương trình chính
data/                  dữ liệu mẫu (nhan_vien.csv/.xlsx, hoc_phi.csv)
templates/             sinh_nhat.txt/.html, nhac_han.txt
tests/fake_smtp.py     SMTP giả lập (pip install aiosmtpd) để thử không gửi thật
output/                (tự tạo) send_log.csv, run.log, preview/*.eml
```

## Template
Dòng đầu là `Subject: ...`, chèn dữ liệu bằng `{{ten_cot}}` (tên cột không phân biệt hoa/thường).
Biến tự tính: `{{ngay_hom_nay}}`, `{{tuoi}}` (với `--birthday`), `{{so_ngay_con_lai}}` (với `--due`).
Nếu một biến bị thiếu/trống thì người đó bị bỏ qua và ghi log, không gửi thư lỗi nội dung.

## Cách dùng – 3 chế độ, an toàn theo mặc định
```bash
# 1) DRY-RUN (mặc định): không gửi gì, chỉ xem trước + lưu .eml
python email_sender.py -d data/nhan_vien.csv -t templates/sinh_nhat.txt --birthday ngay_sinh

# 2) TEST: gửi thật nhưng tất cả về hộp thư của bạn, tối đa 3 thư, tiêu đề có [TEST]
python email_sender.py -d data/nhan_vien.csv -t templates/sinh_nhat.txt --birthday ngay_sinh --test-to ban@gmail.com

# 3) GỬI THẬT (hỏi xác nhận; thêm --yes khi chạy tự động/cron)
python email_sender.py -d data/nhan_vien.csv -t templates/sinh_nhat.txt --birthday ngay_sinh --send
```

Các bài toán trong đề:
```bash
# Nhắc hạn thanh toán (còn ≤3 ngày, chưa thanh toán, mỗi hoá đơn nhắc 1 lần/tháng)
python email_sender.py -d data/hoc_phi.csv -t templates/nhac_han.txt \
   --due han_thanh_toan --within 3 --where trang_thai=chua_thanh_toan \
   --key-cols ma_hoa_don --dedup-scope month --send

# Sinh nhật hằng năm (cho phép gửi lại năm sau)
... --birthday ngay_sinh --dedup-scope year --send

# Thông báo kết quả học tập / lịch học: dùng --where lop=CNTT1 và cột file_dinh_kem
# (bảng điểm PDF riêng của từng người, ngăn cách nhiều file bằng dấu ;)
```

Dùng `--today 18/11/2026` để giả lập ngày khi thử. Dùng `python email_sender.py -h` xem toàn bộ tuỳ chọn.

## Đối chiếu yêu cầu
| # | Yêu cầu | Cách thực hiện |
|---|---|---|
| 1 | Đọc Excel/CSV | `read_table` (CSV tự nhận dấu `,` `;` Tab, UTF-8 BOM; XLSX qua openpyxl, `--sheet`) |
| 2 | Kiểm tra email | `validate_email`: trống, sai định dạng, dấu chấm lạ, quá dài |
| 3 | Xác định đối tượng | `--where`, `--birthday`, `--due/--within/--include-overdue` |
| 4 | Nội dung riêng | `render` với `{{cot}}`, hỗ trợ text + HTML (tự escape) |
| 5 | Gửi tự động | `smtplib` (STARTTLS / SSL / không mã hoá), nghỉ `--delay` giữa thư |
| 6 | Đính kèm | cột `file_dinh_kem` (riêng từng người) + `--attach` (chung); kiểm tra tồn tại và ≤20 MB |
| 7 | Không gửi lặp | khoá = chiến dịch + email + `--key-cols` + kỳ `--dedup-scope`; kiểm tra cả trong file lẫn log cũ |
| 8 | Log | `output/send_log.csv` (thời gian, trạng thái, dòng, email, tiêu đề, file, lỗi), ghi ngay từng thư |
| 9 | Xử lý lỗi | lỗi tạm 4xx/mất kết nối → kết nối lại + thử lại tăng dần; lỗi 5xx → ghi `FAILED`, đi tiếp; sai mật khẩu → dừng ngay |
| 10 | Chế độ test | mặc định dry-run; `--test-to`; `--limit`; `--send` cần xác nhận |

## Trạng thái trong log
`SENT`, `TEST_SENT`, `DRYRUN`, `FAILED`, `SKIPPED_INVALID_EMAIL`, `SKIPPED_DUPLICATE`, `SKIPPED_PREP_ERROR`.
Chỉ `SENT` mới được tính là "đã gửi". Thư `FAILED` sẽ được thử lại ở lần chạy sau; test và dry-run không ảnh hưởng.
Muốn gửi lại cho một người: xoá dòng `SENT` của họ khỏi `send_log.csv` hoặc đổi `--campaign`.

## Lưu ý
- Không commit file `.env`. Gmail giới hạn khoảng 500 thư/ngày; giữ `--delay` ≥ 1 giây.
- Thử cục bộ không cần tài khoản thật: `python tests/fake_smtp.py` rồi đặt `SMTP_HOST=127.0.0.1 SMTP_PORT=2525 SMTP_SECURITY=none`.
