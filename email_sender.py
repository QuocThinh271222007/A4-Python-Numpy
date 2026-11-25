#!/usr/bin/env python3
"""
Bài A4 - Tự động gửi email từ dữ liệu Excel/CSV.

Pipeline:  Excel/CSV -> Lọc người nhận -> Sinh nội dung -> Gửi email -> Log kết quả

Chế độ chạy (an toàn theo mặc định):
  (mặc định)      DRY-RUN : không kết nối SMTP, chỉ xem trước + lưu file .eml
  --test-to MAIL  TEST    : gửi thật nhưng TẤT CẢ về MAIL, tối đa --limit (mặc định 3) thư
  --send          GỬI THẬT: gửi đến đúng người nhận (có bước xác nhận, bỏ qua bằng --yes)

Ví dụ:
  python email_sender.py -d data/nhan_vien.csv -t templates/sinh_nhat.txt --birthday ngay_sinh
  python email_sender.py -d data/nhan_vien.csv -t templates/sinh_nhat.txt --birthday ngay_sinh --test-to toi@gmail.com
  python email_sender.py -d data/nhan_vien.csv -t templates/sinh_nhat.txt --birthday ngay_sinh --send
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import logging
import mimetypes
import os
import re
import smtplib
import socket
import ssl
import sys
import time
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from email.utils import formataddr, formatdate, make_msgid
from html import escape as html_escape
from pathlib import Path

log = logging.getLogger("email_sender")

MAX_ATTACH_BYTES = 20 * 1024 * 1024  # 20 MB (Gmail giới hạn 25 MB cho cả thư)
LOG_FIELDS = ["timestamp", "campaign", "mode", "status", "row", "email",
              "name", "subject", "attachments", "key", "error"]

# Trạng thái trong file log
S_SENT, S_FAILED = "SENT", "FAILED"
S_TEST, S_DRY = "TEST_SENT", "DRYRUN"
S_BAD_EMAIL, S_DUP, S_PREP_ERR = "SKIPPED_INVALID_EMAIL", "SKIPPED_DUPLICATE", "SKIPPED_PREP_ERROR"


# --------------------------------------------------------------------------- #
# 1. Đọc dữ liệu Excel / CSV
# --------------------------------------------------------------------------- #
DATE_FORMATS = ("%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y")


def parse_date(value: str) -> date | None:
    value = (value or "").strip()
    for fmt in DATE_FORMATS:
        try:
            return datetime.strptime(value, fmt).date()
        except ValueError:
            continue
    return None


def _cell_to_str(v) -> str:
    if v is None:
        return ""
    if isinstance(v, datetime):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v).strip()


def read_table(path: Path, sheet: str | None = None) -> list[dict[str, str]]:
    """Trả về list dict; tên cột được chuẩn hoá về chữ thường, bỏ khoảng trắng."""
    ext = path.suffix.lower()
    if ext in (".csv", ".txt", ".tsv"):
        with open(path, newline="", encoding="utf-8-sig") as f:
            sample = f.read(4096)
            f.seek(0)
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\t")
            except csv.Error:
                dialect = csv.excel
            raw = [dict(r) for r in csv.DictReader(f, dialect=dialect)]
    elif ext in (".xlsx", ".xlsm"):
        try:
            from openpyxl import load_workbook
        except ImportError:
            sys.exit("Thiếu thư viện openpyxl: pip install openpyxl")
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb[sheet] if sheet else wb.active
        it = ws.iter_rows(values_only=True)
        header = [str(h).strip() if h is not None else "" for h in next(it, [])]
        raw = [dict(zip(header, r)) for r in it if any(c not in (None, "") for c in r)]
        wb.close()
    else:
        sys.exit(f"Định dạng không hỗ trợ: {ext} (dùng .csv hoặc .xlsx)")

    rows = []
    for r in raw:
        rows.append({(k or "").strip().lower(): _cell_to_str(v) for k, v in r.items() if k})
    return rows


# --------------------------------------------------------------------------- #
# 2. Kiểm tra email
# --------------------------------------------------------------------------- #
_EMAIL_RE = re.compile(
    r"^[A-Za-z0-9._%+\-]+@(?:[A-Za-z0-9](?:[A-Za-z0-9\-]*[A-Za-z0-9])?\.)+[A-Za-z]{2,}$")


def validate_email(addr: str) -> str | None:
    """Trả về None nếu hợp lệ, ngược lại là lý do lỗi."""
    addr = (addr or "").strip()
    if not addr:
        return "email trống"
    if len(addr) > 254:
        return "email quá dài"
    if ".." in addr or addr.startswith(".") or addr.split("@")[0].endswith("."):
        return "dấu chấm không hợp lệ"
    if not _EMAIL_RE.match(addr):
        return "sai định dạng"
    return None


# --------------------------------------------------------------------------- #
# 3. Lọc người nhận
# --------------------------------------------------------------------------- #
def _birthday_this_year(b: date, today: date) -> date:
    try:
        return b.replace(year=today.year)
    except ValueError:  # 29/02 vào năm không nhuận -> 28/02
        return date(today.year, 2, 28)


def select_recipients(rows, args, today: date):
    """Giữ lại những dòng thoả TẤT CẢ điều kiện lọc; bổ sung các trường tính toán."""
    selected = []
    for i, row in enumerate(rows, start=2):  # dòng 1 là header
        row = dict(row)
        row["ngay_hom_nay"] = today.strftime("%d/%m/%Y")
        ok = True

        for cond in args.where or []:
            m = re.match(r"^\s*([^=!]+?)\s*(!=|=)\s*(.*?)\s*$", cond)
            if not m:
                sys.exit(f"--where sai cú pháp: {cond!r} (dùng cot=gia_tri hoặc cot!=gia_tri)")
            col, op, val = m.group(1).lower(), m.group(2), m.group(3).lower()
            eq = row.get(col, "").strip().lower() == val
            if (op == "=" and not eq) or (op == "!=" and eq):
                ok = False

        if ok and args.birthday:
            b = parse_date(row.get(args.birthday.lower(), ""))
            if not b or _birthday_this_year(b, today) != today:
                ok = False
            else:
                row["tuoi"] = str(today.year - b.year)

        if ok and args.due:
            d = parse_date(row.get(args.due.lower(), ""))
            if not d:
                ok = False
            else:
                left = (d - today).days
                low = -10**6 if args.include_overdue else 0
                if not (low <= left <= args.within):
                    ok = False
                else:
                    row["so_ngay_con_lai"] = str(left)

        if ok:
            selected.append((i, row))
    return selected


# --------------------------------------------------------------------------- #
# 4. Sinh nội dung riêng cho từng người
# --------------------------------------------------------------------------- #
_PLACEHOLDER = re.compile(r"\{\{\s*(\w+)\s*\}\}")


class TemplateError(Exception):
    pass


def render(template: str, row: dict[str, str], html: bool = False) -> str:
    missing = []

    def repl(m):
        key = m.group(1).lower()
        val = row.get(key, "")
        if val == "":
            missing.append(key)
            return ""
        return html_escape(val) if html else val

    out = _PLACEHOLDER.sub(repl, template)
    if missing:
        raise TemplateError("thiếu/trống dữ liệu cho: " + ", ".join(sorted(set(missing))))
    return out


def load_template(path: Path) -> tuple[str, str]:
    """File template: dòng đầu 'Subject: ...' rồi dòng trống, sau đó là nội dung."""
    text = path.read_text(encoding="utf-8-sig")
    m = re.match(r"^\s*(?:Subject|Tiêu đề|Tieu de)\s*:\s*(.+?)\r?\n(.*)$", text, re.S | re.I)
    if not m:
        sys.exit(f"Template {path} phải bắt đầu bằng dòng 'Subject: ...'")
    return m.group(1).strip(), m.group(2).lstrip("\r\n")


# --------------------------------------------------------------------------- #
# 5. Dựng email + đính kèm file
# --------------------------------------------------------------------------- #
def resolve_attachments(row, args) -> list[Path]:
    names = []
    if args.attach_column:
        raw = row.get(args.attach_column.lower(), "")
        names += [n.strip() for n in re.split(r"[;|]", raw) if n.strip()]
    paths = [Path(args.attach_dir) / n if not Path(n).is_absolute() else Path(n) for n in names]
    paths += [Path(p) for p in (args.attach or [])]

    total = 0
    for p in paths:
        if not p.is_file():
            raise FileNotFoundError(f"không tìm thấy file đính kèm: {p}")
        total += p.stat().st_size
    if total > MAX_ATTACH_BYTES:
        raise ValueError(f"tổng dung lượng đính kèm {total/1e6:.1f} MB vượt giới hạn {MAX_ATTACH_BYTES/1e6:.0f} MB")
    return paths


def build_message(sender, sender_name, to_addr, subject, body_text, body_html, attachments, reply_to=None):
    msg = EmailMessage()
    msg["From"] = formataddr((sender_name, sender)) if sender_name else sender
    msg["To"] = to_addr
    msg["Subject"] = re.sub(r"[\r\n]+", " ", subject).strip()  # chống header injection
    msg["Date"] = formatdate(localtime=True)
    msg["Message-ID"] = make_msgid(domain=sender.split("@")[-1])
    if reply_to:
        msg["Reply-To"] = reply_to
    msg.set_content(body_text)
    if body_html:
        msg.add_alternative(body_html, subtype="html")
    for p in attachments:
        ctype, enc = mimetypes.guess_type(p.name)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        msg.add_attachment(p.read_bytes(), maintype=maintype, subtype=subtype, filename=p.name)
    return msg


# --------------------------------------------------------------------------- #
# 6. Gửi SMTP + xử lý lỗi (retry / reconnect / phân loại lỗi)
# --------------------------------------------------------------------------- #
class FatalSMTPError(Exception):
    """Lỗi làm hỏng toàn bộ phiên (sai tài khoản, không kết nối được...) -> dừng chạy."""


class PermanentSendError(Exception):
    """Lỗi riêng của một người nhận (hộp thư không tồn tại, bị từ chối...) -> không retry."""


@dataclass
class SMTPConfig:
    host: str = "localhost"
    port: int = 587
    security: str = "starttls"  # starttls | ssl | none
    user: str = ""
    password: str = ""
    sender: str = ""
    sender_name: str = ""
    timeout: int = 30


class Mailer:
    def __init__(self, cfg: SMTPConfig, retries: int = 3, backoff: float = 2.0):
        self.cfg, self.retries, self.backoff = cfg, retries, backoff
        self.server: smtplib.SMTP | None = None

    def connect(self):
        c = self.cfg
        try:
            if c.security == "ssl":
                self.server = smtplib.SMTP_SSL(c.host, c.port, timeout=c.timeout,
                                               context=ssl.create_default_context())
            else:
                self.server = smtplib.SMTP(c.host, c.port, timeout=c.timeout)
                self.server.ehlo()
                if c.security == "starttls":
                    self.server.starttls(context=ssl.create_default_context())
                    self.server.ehlo()
            if c.user:
                self.server.login(c.user, c.password)
        except smtplib.SMTPAuthenticationError as e:
            raise FatalSMTPError(f"đăng nhập SMTP thất bại ({e.smtp_code}): kiểm tra user/mật khẩu ứng dụng") from e
        except (OSError, smtplib.SMTPException) as e:
            raise ConnectionError(f"không kết nối được SMTP {c.host}:{c.port}: {e}") from e

    def close(self):
        if self.server:
            try:
                self.server.quit()
            except Exception:
                pass
            self.server = None

    @staticmethod
    def _classify(exc: Exception) -> str:
        """'permanent' (lỗi của người nhận) | 'transient' (thử lại được)."""
        if isinstance(exc, smtplib.SMTPRecipientsRefused):
            codes = [c for c, _ in exc.recipients.values()]
            return "permanent" if codes and all(c >= 500 for c in codes) else "transient"
        if isinstance(exc, smtplib.SMTPResponseException):  # có smtp_code
            return "permanent" if exc.smtp_code >= 500 else "transient"
        if isinstance(exc, smtplib.SMTPSenderRefused):
            return "permanent"
        return "transient"  # mất kết nối, timeout, OSError...

    def send(self, msg: EmailMessage, to_addr: str) -> int:
        """Gửi 1 thư, trả về số lần thử. Ném PermanentSendError / ConnectionError / FatalSMTPError."""
        last = None
        for attempt in range(1, self.retries + 2):
            try:
                if self.server is None:
                    self.connect()
                self.server.send_message(msg, to_addrs=[to_addr])
                return attempt
            except FatalSMTPError:
                raise
            except (smtplib.SMTPException, OSError, socket.timeout, ConnectionError) as e:
                last = e
                if isinstance(e, smtplib.SMTPAuthenticationError):
                    raise FatalSMTPError(str(e)) from e
                if self._classify(e) == "permanent":
                    # RSET để phiên còn dùng được cho thư kế tiếp
                    try:
                        self.server and self.server.rset()
                    except Exception:
                        self.close()
                    raise PermanentSendError(_short(e)) from e
                log.warning("Lỗi tạm thời (lần %d/%d): %s", attempt, self.retries + 1, _short(e))
                self.close()  # buộc kết nối lại ở lần thử sau
                if attempt <= self.retries:
                    time.sleep(self.backoff * 2 ** (attempt - 1))
        raise ConnectionError(f"hết {self.retries + 1} lần thử: {_short(last)}")


def _short(e) -> str:
    return re.sub(r"\s+", " ", f"{type(e).__name__}: {e}")[:300]


# --------------------------------------------------------------------------- #
# 7. Log trạng thái + chống gửi lặp
# --------------------------------------------------------------------------- #
class SendLog:
    def __init__(self, path: Path):
        self.path = path
        path.parent.mkdir(parents=True, exist_ok=True)
        self.new = not path.exists() or path.stat().st_size == 0

    def sent_keys(self, campaign: str) -> set[str]:
        """Các key đã gửi THÀNH CÔNG THẬT (không tính dry-run/test)."""
        if self.new:
            return set()
        with open(self.path, newline="", encoding="utf-8-sig") as f:
            return {r["key"] for r in csv.DictReader(f)
                    if r.get("campaign") == campaign and r.get("status") == S_SENT}

    def write(self, **kw):
        row = {k: kw.get(k, "") for k in LOG_FIELDS}
        row["timestamp"] = datetime.now().isoformat(timespec="seconds")
        with open(self.path, "a", newline="", encoding="utf-8-sig" if self.new else "utf-8") as f:
            w = csv.DictWriter(f, fieldnames=LOG_FIELDS)
            if self.new:
                w.writeheader()
                self.new = False
            w.writerow(row)
            f.flush()
            os.fsync(f.fileno())  # ghi ngay xuống đĩa: lỡ crash vẫn không gửi lặp


def make_key(campaign, email, row, key_cols, scope, today) -> str:
    period = {"forever": "", "year": str(today.year), "month": today.strftime("%Y-%m"),
              "day": today.isoformat()}[scope]
    parts = [campaign, email.lower(), period] + [row.get(c.lower(), "") for c in key_cols]
    return hashlib.sha1("|".join(parts).encode("utf-8")).hexdigest()[:16]


# --------------------------------------------------------------------------- #
# 8. Cấu hình (.env) và tham số dòng lệnh
# --------------------------------------------------------------------------- #
def load_env(path: Path):
    if not path.is_file():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def parse_args(argv=None):
    p = argparse.ArgumentParser(description="Tự động gửi email từ Excel/CSV",
                                formatter_class=argparse.RawDescriptionHelpFormatter,
                                epilog=__doc__)
    g = p.add_argument_group("dữ liệu & nội dung")
    g.add_argument("-d", "--data", required=True, type=Path, help="file .csv / .xlsx")
    g.add_argument("--sheet", help="tên sheet (xlsx), mặc định sheet đầu")
    g.add_argument("-t", "--template", required=True, type=Path, help="template văn bản (dòng đầu: Subject: ...)")
    g.add_argument("--html-template", type=Path, help="template HTML (tuỳ chọn), dùng chung Subject với --template")
    g.add_argument("--email-col", default="email", help="tên cột email (mặc định: email)")
    g.add_argument("--name-col", default="ten", help="tên cột họ tên, chỉ để ghi log (mặc định: ten)")
    g.add_argument("--campaign", help="tên chiến dịch, dùng để chống gửi lặp (mặc định: tên file template)")

    f = p.add_argument_group("lọc người nhận (tất cả điều kiện phải thoả)")
    f.add_argument("--where", action="append", metavar="COT=GIATRI", help="vd: trang_thai=chua_thanh_toan (lặp được; dùng != để loại trừ)")
    f.add_argument("--birthday", metavar="COT_NGAY_SINH", help="chỉ người có sinh nhật hôm nay (tạo thêm biến {{tuoi}})")
    f.add_argument("--due", metavar="COT_HAN", help="chỉ người có hạn trong --within ngày tới (tạo thêm biến {{so_ngay_con_lai}})")
    f.add_argument("--within", type=int, default=3, help="số ngày tới cho --due (mặc định 3)")
    f.add_argument("--include-overdue", action="store_true", help="--due: gồm cả người đã quá hạn")
    f.add_argument("--today", type=parse_date, help="giả lập ngày hôm nay (dd/mm/yyyy) để thử")

    a = p.add_argument_group("đính kèm")
    a.add_argument("--attach-column", default="file_dinh_kem", help="cột chứa tên file đính kèm riêng (ngăn cách bằng ;)")
    a.add_argument("--attach-dir", default=".", help="thư mục chứa file đính kèm (đường dẫn tương đối)")
    a.add_argument("--attach", action="append", help="file đính kèm chung cho mọi người")

    d = p.add_argument_group("chống gửi lặp & log")
    d.add_argument("--key-cols", nargs="*", default=[], metavar="COT", help="cột thêm vào khoá chống lặp (vd: ma_hoa_don)")
    d.add_argument("--dedup-scope", choices=["forever", "year", "month", "day"], default="forever",
                   help="trong bao lâu một người không bị gửi lại (sinh nhật hằng năm -> year)")
    d.add_argument("--outdir", type=Path, default=Path("output"), help="thư mục log/preview")
    d.add_argument("--log", type=Path, help="file log CSV (mặc định: <outdir>/send_log.csv)")

    m = p.add_argument_group("chế độ chạy")
    m.add_argument("--send", action="store_true", help="GỬI THẬT đến người nhận")
    m.add_argument("--test-to", metavar="EMAIL", help="chế độ test: gửi tất cả về địa chỉ này")
    m.add_argument("--limit", type=int, help="chỉ xử lý N người đầu (test mặc định 3)")
    m.add_argument("--yes", action="store_true", help="bỏ qua bước xác nhận khi --send")

    s = p.add_argument_group("SMTP")
    s.add_argument("--env", type=Path, default=Path(".env"), help="file cấu hình SMTP (mặc định .env)")
    s.add_argument("--delay", type=float, default=1.0, help="giây nghỉ giữa các thư (mặc định 1.0)")
    s.add_argument("--retries", type=int, default=3, help="số lần thử lại khi lỗi tạm thời")
    s.add_argument("--reply-to")
    return p.parse_args(argv)


# --------------------------------------------------------------------------- #
# 9. Pipeline chính
# --------------------------------------------------------------------------- #
@dataclass
class Job:
    row_no: int
    email: str
    name: str
    key: str
    subject: str
    msg: EmailMessage
    attachments: list[str] = field(default_factory=list)


def main(argv=None) -> int:
    args = parse_args(argv)
    args.outdir.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(
        level=logging.INFO, format="%(asctime)s %(levelname)-7s %(message)s", datefmt="%H:%M:%S",
        handlers=[logging.StreamHandler(sys.stdout),
                  logging.FileHandler(args.outdir / "run.log", encoding="utf-8")])

    if args.send and args.test_to:
        sys.exit("Chỉ chọn một: --send hoặc --test-to")
    mode = "SEND" if args.send else "TEST" if args.test_to else "DRYRUN"
    if args.test_to and (err := validate_email(args.test_to)):
        sys.exit(f"--test-to không hợp lệ: {err}")

    load_env(args.env)
    today = args.today or date.today()
    campaign = args.campaign or args.template.stem
    sendlog = SendLog(args.log or args.outdir / "send_log.csv")
    log.info("=== Chiến dịch '%s' | chế độ %s | ngày %s ===", campaign, mode, today.strftime("%d/%m/%Y"))

    # --- [1] Đọc
    rows = read_table(args.data, args.sheet)
    email_col, name_col = args.email_col.lower(), args.name_col.lower()
    if not rows:
        sys.exit("File dữ liệu trống")
    if email_col not in rows[0]:
        sys.exit(f"Không có cột '{args.email_col}'. Các cột hiện có: {', '.join(rows[0])}")
    log.info("[1] Đọc %d dòng từ %s", len(rows), args.data)

    # --- [2] Lọc
    selected = select_recipients(rows, args, today)
    log.info("[2] Lọc: %d/%d dòng thuộc đối tượng cần gửi", len(selected), len(rows))

    subj_tpl, body_tpl = load_template(args.template)
    html_tpl = args.html_template.read_text(encoding="utf-8-sig") if args.html_template else None

    sender = os.environ.get("SMTP_FROM") or os.environ.get("SMTP_USER") or "noreply@example.com"
    sender_name = os.environ.get("SMTP_FROM_NAME", "")

    already = sendlog.sent_keys(campaign)
    seen_in_run: set[str] = set()
    jobs: list[Job] = []
    stats = {"bad_email": 0, "dup": 0, "prep_err": 0}

    def skip(status, row_no, email, name, err, key=""):
        sendlog.write(campaign=campaign, mode=mode, status=status, row=row_no,
                      email=email, name=name, key=key, error=err)
        log.warning("  dòng %s <%s> bị bỏ qua: %s", row_no, email, err)

    # --- [3][4][5] Kiểm tra email, chống lặp, sinh nội dung, đính kèm
    for row_no, row in selected:
        email = row.get(email_col, "").strip()
        name = row.get(name_col, "")
        if err := validate_email(email):
            stats["bad_email"] += 1
            skip(S_BAD_EMAIL, row_no, email, name, err)
            continue
        key = make_key(campaign, email, row, args.key_cols, args.dedup_scope, today)
        if key in already or key in seen_in_run:
            stats["dup"] += 1
            skip(S_DUP, row_no, email, name,
                 "đã gửi trước đó" if key in already else "trùng trong file", key)
            continue
        seen_in_run.add(key)
        try:
            subject = render(subj_tpl, row)
            body = render(body_tpl, row)
            body_html = render(html_tpl, row, html=True) if html_tpl else None
            files = resolve_attachments(row, args)
        except (TemplateError, OSError, ValueError) as e:
            stats["prep_err"] += 1
            skip(S_PREP_ERR, row_no, email, name, str(e), key)
            continue
        to_addr = args.test_to or email
        if mode == "TEST":
            subject = "[TEST] " + subject
            body = f"[TEST - người nhận thật sẽ là: {email}]\n\n" + body
        msg = build_message(sender, sender_name, to_addr, subject, body, body_html, files, args.reply_to)
        jobs.append(Job(row_no, email, name, key, subject, msg, [p.name for p in files]))

    limit = args.limit if args.limit is not None else (3 if mode == "TEST" else None)
    if limit is not None and len(jobs) > limit:
        log.info("    Giới hạn --limit: chỉ xử lý %d/%d thư", limit, len(jobs))
        jobs = jobs[:limit]
    log.info("[3-5] Sẵn sàng %d thư | bỏ qua: %d email lỗi, %d trùng, %d lỗi dữ liệu/đính kèm",
             len(jobs), stats["bad_email"], stats["dup"], stats["prep_err"])

    # --- [6] Gửi
    ok = failed = 0
    if mode == "DRYRUN":
        prev = args.outdir / "preview"
        prev.mkdir(exist_ok=True)
        for j in jobs:
            (prev / f"{j.row_no:04d}_{re.sub(r'[^A-Za-z0-9]+', '_', j.email)}.eml").write_bytes(bytes(j.msg))
            sendlog.write(campaign=campaign, mode=mode, status=S_DRY, row=j.row_no, email=j.email,
                          name=j.name, subject=j.subject, attachments=";".join(j.attachments), key=j.key)
            log.info("  [DRY] dòng %-3s → %-30s | %s | đính kèm: %s", j.row_no, j.email, j.subject,
                     ", ".join(j.attachments) or "-")
            ok += 1
        log.info("DRY-RUN: KHÔNG gửi thư nào. File xem trước ở %s (mở .eml bằng Outlook/Thunderbird)", prev)
    elif jobs:
        if mode == "SEND" and not args.yes:
            if not sys.stdin.isatty():
                sys.exit("Chạy không tương tác: thêm --yes để xác nhận gửi thật.")
            if input(f"Sắp GỬI THẬT {len(jobs)} email. Gõ 'GUI' để xác nhận: ").strip().upper() != "GUI":
                sys.exit("Đã huỷ.")
        cfg = SMTPConfig(
            host=os.environ.get("SMTP_HOST", "smtp.gmail.com"),
            port=int(os.environ.get("SMTP_PORT", "587")),
            security=os.environ.get("SMTP_SECURITY", "starttls").lower(),
            user=os.environ.get("SMTP_USER", ""), password=os.environ.get("SMTP_PASSWORD", ""),
            sender=sender, sender_name=sender_name)
        mailer = Mailer(cfg, retries=args.retries)
        ok_status = S_SENT if mode == "SEND" else S_TEST
        try:
            mailer.connect()
            log.info("[6] Đã kết nối SMTP %s:%s", cfg.host, cfg.port)
            for n, j in enumerate(jobs, 1):
                to_addr = args.test_to or j.email
                try:
                    tries = mailer.send(j.msg, to_addr)
                    ok += 1
                    sendlog.write(campaign=campaign, mode=mode, status=ok_status, row=j.row_no,
                                  email=j.email, name=j.name, subject=j.subject,
                                  attachments=";".join(j.attachments), key=j.key,
                                  error=f"thành công sau {tries} lần thử" if tries > 1 else "")
                    log.info("  (%d/%d) ✓ %s", n, len(jobs), j.email)
                except (PermanentSendError, ConnectionError) as e:
                    failed += 1
                    sendlog.write(campaign=campaign, mode=mode, status=S_FAILED, row=j.row_no,
                                  email=j.email, name=j.name, subject=j.subject,
                                  attachments=";".join(j.attachments), key=j.key, error=str(e))
                    log.error("  (%d/%d) ✗ %s: %s", n, len(jobs), j.email, e)
                if n < len(jobs):
                    time.sleep(args.delay)
        except FatalSMTPError as e:
            log.critical("DỪNG: %s", e)
            failed += len(jobs) - ok - failed
            return 2
        except ConnectionError as e:
            log.critical("DỪNG: %s", e)
            return 2
        finally:
            mailer.close()

    log.info("=== KẾT QUẢ: thành công %d | thất bại %d | bỏ qua %d | log: %s ===",
             ok, failed, sum(stats.values()), sendlog.path)
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
