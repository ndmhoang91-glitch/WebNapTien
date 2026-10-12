import io
import os
import re
import sqlite3
import unicodedata
from datetime import datetime
from flask import Flask, jsonify, render_template, request, send_file
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

app = Flask(__name__)
app.config["TEMPLATES_AUTO_RELOAD"] = True
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "naptien.db")


def remove_accents(input_str: str) -> str:
    if not input_str:
        return ""
    s = str(input_str).replace("đ", "d").replace("Đ", "D")
    nfkd_form = unicodedata.normalize("NFKD", s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)]).lower().strip()


def normalize_phone(val) -> str:
    """Chuẩn hóa SĐT về dạng chuỗi có số 0 ở đầu (VD: 0932929119)."""
    if val is None:
        return ""
    s = str(val).strip()
    if not s:
        return ""
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"\D", "", s)
    if not digits:
        return ""
    if len(digits) == 9 and not digits.startswith("0"):
        digits = "0" + digits
    elif len(digits) == 11 and digits.startswith("84"):
        digits = "0" + digits[2:]
    return digits


def raw_phone_no_zero(phone: str) -> str:
    """Trả về SĐT bỏ số 0 ở đầu giống định dạng cột F trong file Excel."""
    p = normalize_phone(phone)
    return p[1:] if p.startswith("0") else p


def detect_telco(phone: str) -> str:
    p = normalize_phone(phone)
    if len(p) < 3:
        return "KHÁC"
    prefix = p[:3]
    viettel = {"086", "096", "097", "098", "032", "033", "034", "035", "036", "037", "038", "039"}
    mobi = {"089", "090", "093", "070", "079", "077", "076", "078"}
    vina = {"088", "091", "094", "083", "084", "085", "081", "082"}
    vnmobi = {"092", "052", "056", "058"}
    if prefix in viettel:
        return "VIETTEL"
    if prefix in mobi:
        return "MOBI"
    if prefix in vina:
        return "VINA"
    if prefix in vnmobi:
        return "VNMOBI"
    return "KHÁC"


def parse_amount(val) -> int:
    if val is None:
        return 0
    if isinstance(val, (int, float)):
        amt = int(round(val))
    else:
        digits = re.sub(r"[^\d]", "", str(val))
        amt = int(digits) if digits else 0
    # Nếu nhập tắt dạng 50, 100, 150, 200, 500 (< 10000) thì tự động nhân 1000 -> VNĐ
    if 0 < amt < 10000:
        amt *= 1000
    return amt


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS months (
            month_key TEXT PRIMARY KEY,
            label TEXT NOT NULL,
            note TEXT DEFAULT '',
            created_at TEXT NOT NULL
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS customers (
            phone TEXT PRIMARY KEY,
            name TEXT NOT NULL,
            name_search TEXT NOT NULL,
            telco TEXT NOT NULL,
            note TEXT DEFAULT '',
            updated_at TEXT NOT NULL
        )
        """
    )
    cur.execute(
        """
        CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            month_key TEXT NOT NULL,
            order_type TEXT NOT NULL, -- 'OTP' | 'NAP_HO' | 'CHUA_CHAY'
            phone TEXT NOT NULL,
            customer_name TEXT NOT NULL,
            name_search TEXT NOT NULL,
            telco TEXT NOT NULL,
            amount INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL, -- 'COMPLETED' | 'PENDING'
            otp_status TEXT NOT NULL DEFAULT 'NONE', -- 'NEED_OTP' | 'HAS_OTP' | 'DONE_OTP' | 'NONE'
            otp_code TEXT DEFAULT '',
            batch_code TEXT DEFAULT '', -- 'x', 'xx', 'xxx', ...
            note TEXT DEFAULT '',
            excel_row INTEGER DEFAULT 0,
            updated_at TEXT NOT NULL
        )
        """
    )
    cur.execute("CREATE INDEX IF NOT EXISTS idx_orders_month ON orders(month_key)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_orders_phone ON orders(phone)")
    cur.execute("CREATE INDEX IF NOT EXISTS idx_orders_type ON orders(month_key, order_type)")
    conn.commit()

    # Kiểm tra xem đã nạp dữ liệu khởi tạo từ file Excel mặc định chưa
    cur.execute("SELECT COUNT(*) AS cnt FROM months")
    if cur.fetchone()["cnt"] == 0:
        excel_files = [f for f in os.listdir(BASE_DIR) if f.lower().endswith(".xlsx") and not f.startswith("~$")]
        default_month = "2026-10"
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        cur.execute(
            "INSERT OR IGNORE INTO months (month_key, label, note, created_at) VALUES (?, ?, ?, ?)",
            (default_month, "Tháng 10/2026", "Dữ liệu gốc từ file Form Chuyển Đơn Nạp Tiền.xlsx", now_str),
        )
        conn.commit()
        if excel_files:
            try:
                import_workbook_into_month(os.path.join(BASE_DIR, excel_files[0]), default_month, mode="replace", conn=conn)
            except Exception as e:
                print("Initial Excel import warning:", e)
    conn.close()


def import_workbook_into_month(filepath_or_stream, month_key: str, mode: str = "append", default_order_type: str = "OTP", conn=None):
    close_conn = False
    if conn is None:
        conn = get_db()
        close_conn = True

    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Đảm bảo month_key tồn tại
    parts = month_key.split("-")
    label = f"Tháng {parts[1]}/{parts[0]}" if len(parts) == 2 else month_key
    cur.execute(
        "INSERT OR IGNORE INTO months (month_key, label, note, created_at) VALUES (?, ?, ?, ?)",
        (month_key, label, "", now_str),
    )

    wb = openpyxl.load_workbook(filepath_or_stream, data_only=True)
    sheet_names = wb.sheetnames

    # Lấy bảng ánh xạ khách hàng hiện có
    cur.execute("SELECT phone, name FROM customers")
    kh_map = {r["phone"]: r["name"] for r in cur.fetchall()}

    imported_customers = 0
    imported_orders = 0

    # 1. Kiểm tra nếu là file multi-sheet chuẩn "Form Chuyển Đơn Nạp Tiền.xlsx"
    has_standard_sheets = any(
        s in sheet_names for s in ["Danh Sách Tổng KH", "Trạng thái Đơn OTP", "Trạng thái Đơn Nạp hộ", "Danh Sách Đơn Chưa Cho Chạy"]
    )

    if has_standard_sheets:
        if mode == "replace":
            cur.execute("DELETE FROM orders WHERE month_key = ?", (month_key,))

        if "Danh Sách Tổng KH" in sheet_names:
            s_kh = wb["Danh Sách Tổng KH"]
            for r in range(1, s_kh.max_row + 1):
                p_raw = s_kh.cell(r, 1).value
                n_raw = s_kh.cell(r, 2).value
                phone = normalize_phone(p_raw)
                if not phone or len(phone) < 9:
                    continue
                name = str(n_raw or "").strip() or kh_map.get(phone, "Khách hàng")
                kh_map[phone] = name
                cur.execute(
                    """
                    INSERT INTO customers (phone, name, name_search, telco, updated_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(phone) DO UPDATE SET
                        name = excluded.name,
                        name_search = excluded.name_search,
                        telco = excluded.telco,
                        updated_at = excluded.updated_at
                    """,
                    (phone, name, remove_accents(name), detect_telco(phone), now_str),
                )
                imported_customers += 1

        # Sheet: Trạng thái Đơn OTP
        if "Trạng thái Đơn OTP" in sheet_names:
            s_otp = wb["Trạng thái Đơn OTP"]
            for r in range(2, s_otp.max_row + 1):
                p_raw = s_otp.cell(r, 1).value
                if p_raw is None or isinstance(p_raw, str) and not re.sub(r"\D", "", p_raw):
                    continue
                phone = normalize_phone(p_raw)
                if len(phone) < 9:
                    continue
                name_val = s_otp.cell(r, 2).value
                name = str(name_val).strip() if name_val and not str(name_val).startswith("=") else kh_map.get(phone, "")
                if not name:
                    name = kh_map.get(phone, "Chưa rõ tên")
                else:
                    kh_map[phone] = name

                status_raw = str(s_otp.cell(r, 3).value or "").strip()
                amt = parse_amount(s_otp.cell(r, 4).value)
                col_e = str(s_otp.cell(r, 5).value or "").strip() if s_otp.cell(r, 5).value is not None else ""

                # Xác định trạng thái nạp và trạng thái OTP
                is_done = (status_raw.lower() == "đã nạp") or (col_e.lower() in ("x", "thành công"))
                if is_done:
                    status = "COMPLETED"
                    otp_status = "DONE_OTP"
                else:
                    status = "PENDING"
                    otp_status = "NEED_OTP"

                batch_code = col_e if re.fullmatch(r"x+", col_e.lower()) else ""
                note = "" if batch_code or col_e.lower() == "x" else col_e

                cur.execute(
                    """
                    INSERT INTO orders (
                        month_key, order_type, phone, customer_name, name_search,
                        telco, amount, status, otp_status, otp_code, batch_code, note, excel_row, updated_at
                    ) VALUES (?, 'OTP', ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?)
                    """,
                    (
                        month_key,
                        phone,
                        name,
                        remove_accents(name),
                        detect_telco(phone),
                        amt,
                        status,
                        otp_status,
                        batch_code,
                        note,
                        r,
                        now_str,
                    ),
                )
                imported_orders += 1

        # Sheet: Trạng thái Đơn Nạp hộ
        if "Trạng thái Đơn Nạp hộ" in sheet_names:
            s_nh = wb["Trạng thái Đơn Nạp hộ"]
            for r in range(2, s_nh.max_row + 1):
                p_raw = s_nh.cell(r, 1).value
                if p_raw is None or (isinstance(p_raw, str) and not re.sub(r"\D", "", p_raw)):
                    continue
                phone = normalize_phone(p_raw)
                if len(phone) < 9:
                    continue
                name_val = s_nh.cell(r, 2).value
                name = str(name_val).strip() if name_val and not str(name_val).startswith("=") else kh_map.get(phone, "")
                if not name:
                    name = kh_map.get(phone, "Chưa rõ tên")
                else:
                    kh_map[phone] = name

                status_raw = str(s_nh.cell(r, 3).value or "").strip()
                amt = parse_amount(s_nh.cell(r, 4).value)
                col_e = str(s_nh.cell(r, 5).value or "").strip() if s_nh.cell(r, 5).value is not None else ""

                is_done = (status_raw.lower() == "đã nạp") or (col_e.lower() == "x")
                status = "COMPLETED" if is_done else "PENDING"
                batch_code = "x" if is_done else ""

                cur.execute(
                    """
                    INSERT INTO orders (
                        month_key, order_type, phone, customer_name, name_search,
                        telco, amount, status, otp_status, otp_code, batch_code, note, excel_row, updated_at
                    ) VALUES (?, 'NAP_HO', ?, ?, ?, ?, ?, ?, 'NONE', '', ?, '', ?, ?)
                    """,
                    (
                        month_key,
                        phone,
                        name,
                        remove_accents(name),
                        detect_telco(phone),
                        amt,
                        status,
                        batch_code,
                        r,
                        now_str,
                    ),
                )
                imported_orders += 1

        # Sheet: Danh Sách Đơn Chưa Cho Chạy
        if "Danh Sách Đơn Chưa Cho Chạy" in sheet_names:
            s_cc = wb["Danh Sách Đơn Chưa Cho Chạy"]
            for r in range(2, s_cc.max_row + 1):
                p_raw = s_cc.cell(r, 1).value
                if p_raw is None or (isinstance(p_raw, str) and not re.sub(r"\D", "", p_raw)):
                    continue
                phone = normalize_phone(p_raw)
                if len(phone) < 9:
                    continue
                name_val = s_cc.cell(r, 2).value
                name = str(name_val).strip() if name_val and not str(name_val).startswith("=") else kh_map.get(phone, "")
                if not name:
                    name = kh_map.get(phone, "Chưa rõ tên")
                else:
                    kh_map[phone] = name

                status_note = str(s_cc.cell(r, 3).value or "Chưa cho lên đơn").strip()
                amt = parse_amount(s_cc.cell(r, 4).value)
                col_e = str(s_cc.cell(r, 5).value or "").strip() if s_cc.cell(r, 5).value is not None else ""
                col_f = str(s_cc.cell(r, 6).value or "").strip() if s_cc.cell(r, 6).value is not None else ""

                is_scheduled = col_f.lower() in ("đã lên đơn", "x", "đã chạy")
                status = "COMPLETED" if is_scheduled else "PENDING"
                note_combined = f"{status_note}" + (f" | {col_f}" if col_f else "")

                cur.execute(
                    """
                    INSERT INTO orders (
                        month_key, order_type, phone, customer_name, name_search,
                        telco, amount, status, otp_status, otp_code, batch_code, note, excel_row, updated_at
                    ) VALUES (?, 'CHUA_CHAY', ?, ?, ?, ?, ?, ?, 'NONE', '', ?, ?, ?, ?)
                    """,
                    (
                        month_key,
                        phone,
                        name,
                        remove_accents(name),
                        detect_telco(phone),
                        amt,
                        status,
                        col_e,
                        note_combined,
                        r,
                        now_str,
                    ),
                )
                imported_orders += 1

    else:
        # 2. Hỗ trợ tải lên file Excel bất kỳ (1 hoặc nhiều sheet bảng danh sách KH / đơn nạp)
        if mode == "replace":
            cur.execute("DELETE FROM orders WHERE month_key = ?", (month_key,))

        for sname in sheet_names:
            s = wb[sname]
            # Nhận diện loại đơn dựa theo tên sheet hoặc default_order_type
            sname_lower = remove_accents(sname)
            if "nap ho" in sname_lower:
                otype = "NAP_HO"
            elif "chua" in sname_lower:
                otype = "CHUA_CHAY"
            elif "otp" in sname_lower:
                otype = "OTP"
            elif "tong kh" in sname_lower or "khach hang" in sname_lower:
                otype = "KH_ONLY"
            else:
                otype = default_order_type

            # Tự động nhận diện vị trí cột từ dòng tiêu đề (nếu có)
            col_phone, col_name, col_amount, col_status, col_type, col_otp, col_note = 1, 2, 3, 4, 5, 6, 7
            start_row = 1
            for r in range(1, min(5, s.max_row + 1)):
                row_headers = [remove_accents(str(s.cell(r, c).value or "")) for c in range(1, min(12, s.max_column + 1))]
                if any(h in ("sdt", "so dien thoai", "dien thoai", "mobile", "phone") for h in row_headers):
                    start_row = r + 1
                    for idx, h in enumerate(row_headers, start=1):
                        if any(k in h for k in ("sdt", "dien thoai", "mobile", "phone")):
                            col_phone = idx
                        elif any(k in h for k in ("ten", "khach", "ho ten", "name")):
                            col_name = idx
                        elif any(k in h for k in ("tien", "menh gia", "amount", "so tien")):
                            col_amount = idx
                        elif any(k in h for k in ("trang thai", "tinh trang", "status", "thanh cong")):
                            col_status = idx
                        elif any(k in h for k in ("loai", "hinh thuc")):
                            col_type = idx
                        elif "otp" in h:
                            col_otp = idx
                        elif any(k in h for k in ("ghi chu", "note", "dot")):
                            col_note = idx
                    break

            for r in range(start_row, s.max_row + 1):
                p_raw = s.cell(r, col_phone).value
                phone = normalize_phone(p_raw)
                if not phone or len(phone) < 9:
                    continue
                name_raw = s.cell(r, col_name).value if col_name <= s.max_column else None
                name = str(name_raw).strip() if name_raw and not str(name_raw).startswith("=") else kh_map.get(phone, "Khách hàng")
                kh_map[phone] = name

                cur.execute(
                    """
                    INSERT INTO customers (phone, name, name_search, telco, updated_at)
                    VALUES (?, ?, ?, ?, ?)
                    ON CONFLICT(phone) DO UPDATE SET
                        name = excluded.name,
                        name_search = excluded.name_search,
                        telco = excluded.telco,
                        updated_at = excluded.updated_at
                    """,
                    (phone, name, remove_accents(name), detect_telco(phone), now_str),
                )
                imported_customers += 1

                if otype == "KH_ONLY":
                    continue

                # Đọc số tiền
                amt_val = s.cell(r, col_amount).value if col_amount <= s.max_column else 0
                # Nếu cột 3 là trạng thái và cột 4 là số tiền (giống bố cục sheet OTP/Nạp hộ)
                if isinstance(amt_val, str) and not re.sub(r"\D", "", amt_val) and s.max_column >= 4:
                    alt_amt = s.cell(r, 4).value
                    if isinstance(alt_amt, (int, float)) or (alt_amt and re.sub(r"\D", "", str(alt_amt))):
                        col_status = 3
                        col_amount = 4
                        amt_val = alt_amt

                amt = parse_amount(amt_val)
                st_str = remove_accents(str(s.cell(r, col_status).value or "")) if col_status <= s.max_column else ""
                row_type_str = remove_accents(str(s.cell(r, col_type).value or "")) if col_type <= s.max_column else ""
                otp_str = remove_accents(str(s.cell(r, col_otp).value or "")) if col_otp <= s.max_column else ""
                note_str = str(s.cell(r, col_note).value or "").strip() if col_note <= s.max_column else ""

                actual_type = otype
                if "nap ho" in row_type_str:
                    actual_type = "NAP_HO"
                elif "otp" in row_type_str:
                    actual_type = "OTP"
                elif "chua" in row_type_str:
                    actual_type = "CHUA_CHAY"

                is_done = any(k in st_str for k in ("da nap", "hoan thanh", "thanh cong", "x"))
                status = "COMPLETED" if is_done else "PENDING"

                if actual_type == "OTP":
                    if is_done:
                        otp_status = "DONE_OTP"
                    elif any(k in st_str or k in otp_str for k in ("da co otp", "co otp", "has otp")):
                        otp_status = "HAS_OTP"
                    else:
                        otp_status = "NEED_OTP"
                else:
                    otp_status = "NONE"

                cur.execute(
                    """
                    INSERT INTO orders (
                        month_key, order_type, phone, customer_name, name_search,
                        telco, amount, status, otp_status, otp_code, batch_code, note, excel_row, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?)
                    """,
                    (
                        month_key,
                        actual_type,
                        phone,
                        name,
                        remove_accents(name),
                        detect_telco(phone),
                        amt,
                        status,
                        otp_status,
                        "x" if is_done else "",
                        note_str,
                        r,
                        now_str,
                    ),
                )
                imported_orders += 1

    # Đồng bộ lại toàn bộ khách hàng phát sinh trong orders vào bảng customers
    for phone, name in kh_map.items():
        if phone and name:
            cur.execute(
                """
                INSERT INTO customers (phone, name, name_search, telco, updated_at)
                VALUES (?, ?, ?, ?, ?)
                ON CONFLICT(phone) DO UPDATE SET
                    name = excluded.name,
                    name_search = excluded.name_search,
                    updated_at = excluded.updated_at
                """,
                (phone, name, remove_accents(name), detect_telco(phone), now_str),
            )

    conn.commit()
    if close_conn:
        conn.close()
    return {"imported_customers": imported_customers, "imported_orders": imported_orders}


@app.route("/")
def index():
    return render_template("index.html")


@app.route("/api/months", methods=["GET"])
def api_get_months():
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT m.month_key, m.label, m.note, m.created_at,
               COUNT(o.id) AS total_orders,
               COALESCE(SUM(CASE WHEN o.order_type IN ('OTP', 'NAP_HO') AND o.status = 'COMPLETED' THEN o.amount ELSE 0 END), 0) AS paid_amount,
               COALESCE(SUM(CASE WHEN o.order_type IN ('OTP', 'NAP_HO') AND o.status != 'COMPLETED' THEN o.amount ELSE 0 END), 0) AS unpaid_amount
        FROM months m
        LEFT JOIN orders o ON m.month_key = o.month_key
        GROUP BY m.month_key
        ORDER BY m.month_key DESC
        """
    )
    rows = [dict(r) for r in cur.fetchall()]
    conn.close()
    return jsonify({"months": rows})


@app.route("/api/months", methods=["POST"])
def api_create_month():
    data = request.json or {}
    month_key = str(data.get("month_key", "")).strip()
    copy_from = str(data.get("copy_from", "")).strip()
    reset_status = bool(data.get("reset_status", True))
    note = str(data.get("note", "")).strip()

    if not re.fullmatch(r"\d{4}-\d{2}", month_key):
        return jsonify({"error": "Định dạng tháng không hợp lệ (VD: 2026-11)"}), 400

    parts = month_key.split("-")
    label = f"Tháng {parts[1]}/{parts[0]}"
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT OR REPLACE INTO months (month_key, label, note, created_at) VALUES (?, ?, ?, ?)",
        (month_key, label, note, now_str),
    )

    copied = 0
    if copy_from:
        cur.execute("SELECT * FROM orders WHERE month_key = ? ORDER BY id ASC", (copy_from,))
        source_orders = cur.fetchall()
        for o in source_orders:
            new_status = "PENDING" if reset_status else o["status"]
            if o["order_type"] == "OTP":
                new_otp = "NEED_OTP" if reset_status else o["otp_status"]
            else:
                new_otp = "NONE"
            new_batch = "" if reset_status else o["batch_code"]
            cur.execute(
                """
                INSERT INTO orders (
                    month_key, order_type, phone, customer_name, name_search,
                    telco, amount, status, otp_status, otp_code, batch_code, note, excel_row, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '', ?, ?, ?, ?)
                """,
                (
                    month_key,
                    o["order_type"],
                    o["phone"],
                    o["customer_name"],
                    o["name_search"],
                    o["telco"],
                    o["amount"],
                    new_status,
                    new_otp,
                    new_batch,
                    o["note"],
                    o["excel_row"],
                    now_str,
                ),
            )
            copied += 1

    conn.commit()
    conn.close()
    return jsonify({"success": True, "month_key": month_key, "label": label, "copied_orders": copied})


@app.route("/api/dashboard", methods=["GET"])
def api_dashboard():
    month_key = request.args.get("month_key", "2026-10")
    conn = get_db()
    cur = conn.cursor()

    cur.execute(
        """
        SELECT
            order_type,
            status,
            otp_status,
            COUNT(*) AS cnt,
            COALESCE(SUM(amount), 0) AS total_amount
        FROM orders
        WHERE month_key = ?
        GROUP BY order_type, status, otp_status
        """,
        (month_key,),
    )
    stats_rows = cur.fetchall()

    summary = {
        "month_key": month_key,
        "total_active_orders": 0,
        "total_active_amount": 0,
        "paid_orders": 0,
        "paid_amount": 0,
        "unpaid_orders": 0,
        "unpaid_amount": 0,
        "otp": {
            "total_orders": 0,
            "total_amount": 0,
            "paid_orders": 0,
            "paid_amount": 0,
            "need_otp_orders": 0,
            "need_otp_amount": 0,
            "has_otp_orders": 0,
            "has_otp_amount": 0,
        },
        "nap_ho": {
            "total_orders": 0,
            "total_amount": 0,
            "paid_orders": 0,
            "paid_amount": 0,
            "unpaid_orders": 0,
            "unpaid_amount": 0,
        },
        "chua_chay": {
            "total_orders": 0,
            "total_amount": 0,
            "scheduled_orders": 0,
            "scheduled_amount": 0,
            "holding_orders": 0,
            "holding_amount": 0,
        },
    }

    for r in stats_rows:
        otype = r["order_type"]
        st = r["status"]
        otp_st = r["otp_status"]
        cnt = r["cnt"]
        amt = r["total_amount"]

        if otype == "OTP":
            summary["otp"]["total_orders"] += cnt
            summary["otp"]["total_amount"] += amt
            summary["total_active_orders"] += cnt
            summary["total_active_amount"] += amt
            if st == "COMPLETED":
                summary["otp"]["paid_orders"] += cnt
                summary["otp"]["paid_amount"] += amt
                summary["paid_orders"] += cnt
                summary["paid_amount"] += amt
            else:
                summary["unpaid_orders"] += cnt
                summary["unpaid_amount"] += amt
                if otp_st == "HAS_OTP":
                    summary["otp"]["has_otp_orders"] += cnt
                    summary["otp"]["has_otp_amount"] += amt
                else:
                    summary["otp"]["need_otp_orders"] += cnt
                    summary["otp"]["need_otp_amount"] += amt

        elif otype == "NAP_HO":
            summary["nap_ho"]["total_orders"] += cnt
            summary["nap_ho"]["total_amount"] += amt
            summary["total_active_orders"] += cnt
            summary["total_active_amount"] += amt
            if st == "COMPLETED":
                summary["nap_ho"]["paid_orders"] += cnt
                summary["nap_ho"]["paid_amount"] += amt
                summary["paid_orders"] += cnt
                summary["paid_amount"] += amt
            else:
                summary["nap_ho"]["unpaid_orders"] += cnt
                summary["nap_ho"]["unpaid_amount"] += amt
                summary["unpaid_orders"] += cnt
                summary["unpaid_amount"] += amt

        elif otype == "CHUA_CHAY":
            summary["chua_chay"]["total_orders"] += cnt
            summary["chua_chay"]["total_amount"] += amt
            if st == "COMPLETED":
                summary["chua_chay"]["scheduled_orders"] += cnt
                summary["chua_chay"]["scheduled_amount"] += amt
            else:
                summary["chua_chay"]["holding_orders"] += cnt
                summary["chua_chay"]["holding_amount"] += amt

    cur.execute("SELECT COUNT(*) AS cnt FROM customers")
    summary["total_customers"] = cur.fetchone()["cnt"]

    # Lấy danh sách các đợt (batch_code) có trong tháng
    cur.execute(
        "SELECT DISTINCT batch_code FROM orders WHERE month_key = ? AND batch_code != '' ORDER BY LENGTH(batch_code), batch_code",
        (month_key,),
    )
    summary["batches"] = [r["batch_code"] for r in cur.fetchall()]

    conn.close()
    return jsonify(summary)


@app.route("/api/orders", methods=["GET"])
def api_get_orders():
    month_key = request.args.get("month_key", "2026-10")
    order_type = request.args.get("order_type", "ALL")
    status = request.args.get("status", "ALL")
    otp_status = request.args.get("otp_status", "ALL")
    telco = request.args.get("telco", "ALL")
    batch_code = request.args.get("batch_code", "ALL")
    q = request.args.get("q", "").strip()
    dedup_latest = request.args.get("dedup_latest", "0") == "1"
    page = max(1, int(request.args.get("page", 1)))
    page_size = min(2000, max(10, int(request.args.get("page_size", 100))))

    conn = get_db()
    cur = conn.cursor()

    where_clauses = ["month_key = ?"]
    params = [month_key]

    if order_type != "ALL":
        where_clauses.append("order_type = ?")
        params.append(order_type)

    if status != "ALL":
        where_clauses.append("status = ?")
        params.append(status)

    if otp_status != "ALL":
        where_clauses.append("otp_status = ?")
        params.append(otp_status)

    if telco != "ALL":
        where_clauses.append("telco = ?")
        params.append(telco)

    if batch_code != "ALL":
        where_clauses.append("batch_code = ?")
        params.append(batch_code)

    if q:
        q_norm = remove_accents(q)
        q_digits = re.sub(r"\D", "", q)
        if q_digits:
            where_clauses.append("(phone LIKE ? OR name_search LIKE ? OR customer_name LIKE ?)")
            params.extend([f"%{q_digits}%", f"%{q_norm}%", f"%{q}%"])
        else:
            where_clauses.append("(name_search LIKE ? OR customer_name LIKE ? OR note LIKE ?)")
            params.extend([f"%{q_norm}%", f"%{q}%", f"%{q}%"])

    where_sql = " AND ".join(where_clauses)

    if dedup_latest:
        base_query = f"""
            SELECT * FROM orders
            WHERE id IN (
                SELECT MAX(id) FROM orders WHERE {where_sql} GROUP BY phone, order_type
            )
        """
    else:
        base_query = f"SELECT * FROM orders WHERE {where_sql}"

    cur.execute(
        f"""
        SELECT COUNT(*) AS cnt,
               COALESCE(SUM(amount), 0) AS sum_amount,
               COALESCE(SUM(CASE WHEN status = 'COMPLETED' THEN amount ELSE 0 END), 0) AS sum_paid,
               COALESCE(SUM(CASE WHEN status != 'COMPLETED' THEN amount ELSE 0 END), 0) AS sum_unpaid
        FROM ({base_query})
        """,
        params,
    )
    agg = dict(cur.fetchone())

    offset = (page - 1) * page_size
    # Ưu tiên đưa đơn HAS_OTP và NEED_OTP lên đầu hoặc sắp theo thứ tự gốc
    sort_by = request.args.get("sort_by", "default")
    if sort_by == "otp_priority":
        order_sql = """
            CASE otp_status
                WHEN 'HAS_OTP' THEN 1
                WHEN 'NEED_OTP' THEN 2
                ELSE 3
            END ASC, id ASC
        """
    elif sort_by == "amount_desc":
        order_sql = "amount DESC, id ASC"
    elif sort_by == "unpaid_first":
        order_sql = "CASE WHEN status = 'PENDING' THEN 0 ELSE 1 END ASC, id ASC"
    else:
        order_sql = "id ASC"

    cur.execute(
        f"{base_query} ORDER BY {order_sql} LIMIT ? OFFSET ?",
        params + [page_size, offset],
    )
    items = []
    for r in cur.fetchall():
        d = dict(r)
        d["phone_no_zero"] = raw_phone_no_zero(d["phone"])
        items.append(d)

    conn.close()
    return jsonify(
        {
            "orders": items,
            "total": agg["cnt"],
            "sum_amount": agg["sum_amount"],
            "sum_paid": agg["sum_paid"],
            "sum_unpaid": agg["sum_unpaid"],
            "page": page,
            "page_size": page_size,
        }
    )


@app.route("/api/orders", methods=["POST"])
def api_create_order():
    data = request.json or {}
    month_key = str(data.get("month_key", "2026-10")).strip()
    order_type = str(data.get("order_type", "OTP")).strip()
    phone = normalize_phone(data.get("phone", ""))
    customer_name = str(data.get("customer_name", "")).strip()
    amount = parse_amount(data.get("amount", 0))
    status = str(data.get("status", "PENDING")).strip()
    otp_status = str(data.get("otp_status", "NEED_OTP" if order_type == "OTP" else "NONE")).strip()
    otp_code = str(data.get("otp_code", "")).strip()
    batch_code = str(data.get("batch_code", "")).strip()
    note = str(data.get("note", "")).strip()

    if not phone or len(phone) < 9:
        return jsonify({"error": "Số điện thoại không hợp lệ"}), 400

    conn = get_db()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    if not customer_name:
        cur.execute("SELECT name FROM customers WHERE phone = ?", (phone,))
        row = cur.fetchone()
        customer_name = row["name"] if row else "Khách hàng mới"
    else:
        cur.execute(
            """
            INSERT INTO customers (phone, name, name_search, telco, updated_at)
            VALUES (?, ?, ?, ?, ?)
            ON CONFLICT(phone) DO UPDATE SET
                name = excluded.name,
                name_search = excluded.name_search,
                updated_at = excluded.updated_at
            """,
            (phone, customer_name, remove_accents(customer_name), detect_telco(phone), now_str),
        )

    if order_type == "OTP" and status == "COMPLETED":
        otp_status = "DONE_OTP"
        if not batch_code:
            batch_code = "x"
    elif order_type == "NAP_HO" and status == "COMPLETED":
        batch_code = "x"

    cur.execute(
        """
        INSERT INTO orders (
            month_key, order_type, phone, customer_name, name_search,
            telco, amount, status, otp_status, otp_code, batch_code, note, excel_row, updated_at
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0, ?)
        """,
        (
            month_key,
            order_type,
            phone,
            customer_name,
            remove_accents(customer_name),
            detect_telco(phone),
            amount,
            status,
            otp_status,
            otp_code,
            batch_code,
            note,
            now_str,
        ),
    )
    new_id = cur.lastrowid
    conn.commit()
    conn.close()
    return jsonify({"success": True, "id": new_id})


@app.route("/api/orders/<int:order_id>", methods=["PATCH"])
def api_update_order(order_id):
    data = request.json or {}
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM orders WHERE id = ?", (order_id,))
    order = cur.fetchone()
    if not order:
        conn.close()
        return jsonify({"error": "Không tìm thấy đơn"}), 404

    phone = normalize_phone(data.get("phone", order["phone"]))
    customer_name = str(data.get("customer_name", order["customer_name"])).strip()
    amount = parse_amount(data["amount"]) if "amount" in data else order["amount"]
    order_type = str(data.get("order_type", order["order_type"]))
    status = str(data.get("status", order["status"]))
    otp_status = str(data.get("otp_status", order["otp_status"]))
    otp_code = str(data.get("otp_code", order["otp_code"])).strip()
    batch_code = str(data.get("batch_code", order["batch_code"])).strip()
    note = str(data.get("note", order["note"])).strip()

    # Tự động đồng bộ status <-> otp_status cho Đơn OTP
    if order_type == "OTP":
        if "otp_status" in data and "status" not in data:
            if otp_status == "DONE_OTP":
                status = "COMPLETED"
                if not batch_code:
                    batch_code = "x"
            else:
                status = "PENDING"
                if batch_code == "x":
                    batch_code = ""
        elif "status" in data and "otp_status" not in data:
            if status == "COMPLETED":
                otp_status = "DONE_OTP"
                if not batch_code:
                    batch_code = "x"
            else:
                otp_status = "NEED_OTP"
                if batch_code == "x":
                    batch_code = ""
    elif order_type == "NAP_HO":
        otp_status = "NONE"
        if status == "COMPLETED":
            batch_code = "x"
        elif "status" in data and status == "PENDING" and batch_code == "x":
            batch_code = ""

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        UPDATE orders
        SET phone = ?, customer_name = ?, name_search = ?, telco = ?,
            amount = ?, order_type = ?, status = ?, otp_status = ?,
            otp_code = ?, batch_code = ?, note = ?, updated_at = ?
        WHERE id = ?
        """,
        (
            phone,
            customer_name,
            remove_accents(customer_name),
            detect_telco(phone),
            amount,
            order_type,
            status,
            otp_status,
            otp_code,
            batch_code,
            note,
            now_str,
            order_id,
        ),
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "status": status, "otp_status": otp_status, "amount": amount, "batch_code": batch_code})


@app.route("/api/orders/bulk", methods=["POST"])
def api_bulk_orders():
    data = request.json or {}
    ids = data.get("ids", [])
    action = data.get("action", "")
    if not ids:
        return jsonify({"error": "Chưa chọn đơn nào"}), 400

    conn = get_db()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    placeholders = ",".join(["?"] * len(ids))

    if action == "mark_paid":
        cur.execute(
            f"""
            UPDATE orders
            SET status = 'COMPLETED',
                otp_status = CASE WHEN order_type = 'OTP' THEN 'DONE_OTP' ELSE otp_status END,
                batch_code = CASE WHEN batch_code = '' THEN 'x' ELSE batch_code END,
                updated_at = ?
            WHERE id IN ({placeholders})
            """,
            [now_str] + ids,
        )
    elif action == "mark_unpaid":
        cur.execute(
            f"""
            UPDATE orders
            SET status = 'PENDING',
                otp_status = CASE WHEN order_type = 'OTP' THEN 'NEED_OTP' ELSE 'NONE' END,
                batch_code = CASE WHEN batch_code = 'x' THEN '' ELSE batch_code END,
                updated_at = ?
            WHERE id IN ({placeholders})
            """,
            [now_str] + ids,
        )
    elif action == "mark_has_otp":
        cur.execute(
            f"""
            UPDATE orders
            SET status = 'PENDING',
                otp_status = 'HAS_OTP',
                updated_at = ?
            WHERE id IN ({placeholders})
            """,
            [now_str] + ids,
        )
    elif action == "mark_need_otp":
        cur.execute(
            f"""
            UPDATE orders
            SET status = 'PENDING',
                otp_status = 'NEED_OTP',
                updated_at = ?
            WHERE id IN ({placeholders})
            """,
            [now_str] + ids,
        )
    elif action in ("move_to_otp", "move_to_nap_ho"):
        target_type = "OTP" if action == "move_to_otp" else "NAP_HO"
        cur.execute(f"SELECT * FROM orders WHERE id IN ({placeholders})", ids)
        rows = cur.fetchall()
        for r in rows:
            cur.execute(
                """
                INSERT INTO orders (
                    month_key, order_type, phone, customer_name, name_search,
                    telco, amount, status, otp_status, otp_code, batch_code, note, excel_row, updated_at
                ) VALUES (?, ?, ?, ?, ?, ?, ?, 'PENDING', ?, '', ?, ?, 0, ?)
                """,
                (
                    r["month_key"],
                    target_type,
                    r["phone"],
                    r["customer_name"],
                    r["name_search"],
                    r["telco"],
                    r["amount"],
                    "NEED_OTP" if target_type == "OTP" else "NONE",
                    r["batch_code"],
                    f"Chuyển từ Đơn Chưa Cho Chạy (#{r['id']})",
                    now_str,
                ),
            )
        cur.execute(
            f"UPDATE orders SET status = 'COMPLETED', note = 'Đã đẩy sang ' || ?, updated_at = ? WHERE id IN ({placeholders})",
            [("Đơn OTP" if target_type == "OTP" else "Đơn Nạp Hộ"), now_str] + ids,
        )
    elif action == "delete":
        cur.execute(f"DELETE FROM orders WHERE id IN ({placeholders})", ids)
    else:
        conn.close()
        return jsonify({"error": "Hành động không hợp lệ"}), 400

    conn.commit()
    conn.close()
    return jsonify({"success": True, "count": len(ids)})


@app.route("/api/customers/summary", methods=["GET"])
def api_customers_summary():
    """
    Tra cứu & tổng hợp nhanh khách hàng đã nạp bao nhiêu trong tháng.
    Hỗ trợ gom theo Tên Khách Hàng hoặc theo từng Số Điện Thoại.
    """
    month_key = request.args.get("month_key", "2026-10")
    q = request.args.get("q", "").strip()
    filter_mode = request.args.get("filter", "ALL")  # ALL | HAS_OTP | NEED_OTP | UNPAID | PAID

    conn = get_db()
    cur = conn.cursor()

    where_clauses = ["month_key = ?"]
    params = [month_key]

    if q:
        q_norm = remove_accents(q)
        q_digits = re.sub(r"\D", "", q)
        if q_digits:
            where_clauses.append("(phone LIKE ? OR name_search LIKE ? OR customer_name LIKE ?)")
            params.extend([f"%{q_digits}%", f"%{q_norm}%", f"%{q}%"])
        else:
            where_clauses.append("(name_search LIKE ? OR customer_name LIKE ?)")
            params.extend([f"%{q_norm}%", f"%{q}%"])

    where_sql = " AND ".join(where_clauses)

    cur.execute(
        f"""
        SELECT
            phone,
            customer_name,
            telco,
            COUNT(CASE WHEN order_type IN ('OTP', 'NAP_HO') THEN 1 END) AS total_orders,
            COALESCE(SUM(CASE WHEN order_type IN ('OTP', 'NAP_HO') AND status = 'COMPLETED' THEN amount ELSE 0 END), 0) AS total_paid,
            COALESCE(SUM(CASE WHEN order_type = 'OTP' AND status = 'COMPLETED' THEN amount ELSE 0 END), 0) AS otp_paid,
            COALESCE(SUM(CASE WHEN order_type = 'NAP_HO' AND status = 'COMPLETED' THEN amount ELSE 0 END), 0) AS napho_paid,
            COALESCE(SUM(CASE WHEN order_type IN ('OTP', 'NAP_HO') AND status != 'COMPLETED' THEN amount ELSE 0 END), 0) AS total_unpaid,
            COUNT(CASE WHEN order_type = 'OTP' AND status != 'COMPLETED' AND otp_status = 'NEED_OTP' THEN 1 END) AS need_otp_count,
            COALESCE(SUM(CASE WHEN order_type = 'OTP' AND status != 'COMPLETED' AND otp_status = 'NEED_OTP' THEN amount ELSE 0 END), 0) AS need_otp_amount,
            COUNT(CASE WHEN order_type = 'OTP' AND status != 'COMPLETED' AND otp_status = 'HAS_OTP' THEN 1 END) AS has_otp_count,
            COALESCE(SUM(CASE WHEN order_type = 'OTP' AND status != 'COMPLETED' AND otp_status = 'HAS_OTP' THEN amount ELSE 0 END), 0) AS has_otp_amount,
            COUNT(CASE WHEN order_type = 'NAP_HO' AND status != 'COMPLETED' THEN 1 END) AS napho_unpaid_count,
            COALESCE(SUM(CASE WHEN order_type = 'NAP_HO' AND status != 'COMPLETED' THEN amount ELSE 0 END), 0) AS napho_unpaid_amount,
            COUNT(CASE WHEN order_type = 'CHUA_CHAY' AND status != 'COMPLETED' THEN 1 END) AS holding_count,
            COALESCE(SUM(CASE WHEN order_type = 'CHUA_CHAY' AND status != 'COMPLETED' THEN amount ELSE 0 END), 0) AS holding_amount
        FROM orders
        WHERE {where_sql}
        GROUP BY phone
        ORDER BY total_paid DESC, total_unpaid DESC, customer_name ASC
        """,
        params,
    )

    rows = [dict(r) for r in cur.fetchall()]

    # Nếu người dùng tìm kiếm khách hàng nằm trong Danh Sách Tổng KH nhưng chưa có đơn trong tháng
    if q:
        existing_phones = {r["phone"] for r in rows}
        q_norm = remove_accents(q)
        q_digits = re.sub(r"\D", "", q)
        if q_digits:
            cur.execute("SELECT * FROM customers WHERE phone LIKE ? OR name_search LIKE ?", (f"%{q_digits}%", f"%{q_norm}%"))
        else:
            cur.execute("SELECT * FROM customers WHERE name_search LIKE ? OR name LIKE ?", (f"%{q_norm}%", f"%{q}%"))
        for c in cur.fetchall():
            if c["phone"] not in existing_phones:
                rows.append(
                    {
                        "phone": c["phone"],
                        "customer_name": c["name"],
                        "telco": c["telco"],
                        "total_orders": 0,
                        "total_paid": 0,
                        "otp_paid": 0,
                        "napho_paid": 0,
                        "total_unpaid": 0,
                        "need_otp_count": 0,
                        "need_otp_amount": 0,
                        "has_otp_count": 0,
                        "has_otp_amount": 0,
                        "napho_unpaid_count": 0,
                        "napho_unpaid_amount": 0,
                        "holding_count": 0,
                        "holding_amount": 0,
                    }
                )

    if filter_mode == "NEED_OTP":
        rows = [r for r in rows if r["need_otp_count"] > 0]
    elif filter_mode == "HAS_OTP":
        rows = [r for r in rows if r["has_otp_count"] > 0]
    elif filter_mode == "UNPAID":
        rows = [r for r in rows if r["total_unpaid"] > 0]
    elif filter_mode == "PAID":
        rows = [r for r in rows if r["total_paid"] > 0]

    conn.close()
    return jsonify({"customers": rows, "total": len(rows)})


@app.route("/api/customers/<phone>/detail", methods=["GET"])
def api_customer_detail(phone):
    month_key = request.args.get("month_key", "2026-10")
    norm_p = normalize_phone(phone)
    conn = get_db()
    cur = conn.cursor()

    cur.execute("SELECT * FROM customers WHERE phone = ?", (norm_p,))
    cust = cur.fetchone()
    customer_info = dict(cust) if cust else {"phone": norm_p, "name": "Khách hàng", "telco": detect_telco(norm_p)}

    # Các đơn của SĐT này trong tháng đang chọn
    cur.execute("SELECT * FROM orders WHERE phone = ? AND month_key = ? ORDER BY id ASC", (norm_p, month_key))
    month_orders = [dict(r) for r in cur.fetchall()]

    # Các SĐT khác có cùng Tên Khách Hàng (nếu khách có nhiều số)
    same_name_phones = []
    if customer_info.get("name"):
        cur.execute(
            """
            SELECT phone,
                   COALESCE(SUM(CASE WHEN order_type IN ('OTP', 'NAP_HO') AND status = 'COMPLETED' THEN amount ELSE 0 END), 0) AS paid,
                   COALESCE(SUM(CASE WHEN order_type IN ('OTP', 'NAP_HO') AND status != 'COMPLETED' THEN amount ELSE 0 END), 0) AS unpaid
            FROM orders
            WHERE customer_name = ? AND month_key = ? AND phone != ?
            GROUP BY phone
            """,
            (customer_info["name"], month_key, norm_p),
        )
        same_name_phones = [dict(r) for r in cur.fetchall()]

    # Lịch sử qua các tháng
    cur.execute(
        """
        SELECT o.month_key, m.label,
               COALESCE(SUM(CASE WHEN o.order_type IN ('OTP', 'NAP_HO') AND o.status = 'COMPLETED' THEN o.amount ELSE 0 END), 0) AS paid_amount,
               COALESCE(SUM(CASE WHEN o.order_type IN ('OTP', 'NAP_HO') AND o.status != 'COMPLETED' THEN o.amount ELSE 0 END), 0) AS unpaid_amount,
               COUNT(CASE WHEN o.order_type IN ('OTP', 'NAP_HO') THEN 1 END) AS order_count
        FROM orders o
        LEFT JOIN months m ON o.month_key = m.month_key
        WHERE o.phone = ?
        GROUP BY o.month_key
        ORDER BY o.month_key DESC
        """,
        (norm_p,),
    )
    history = [dict(r) for r in cur.fetchall()]

    conn.close()
    return jsonify(
        {
            "customer": customer_info,
            "month_orders": month_orders,
            "same_name_phones": same_name_phones,
            "history": history,
        }
    )


@app.route("/api/customers", methods=["POST"])
def api_upsert_customer():
    data = request.json or {}
    phone = normalize_phone(data.get("phone", ""))
    name = str(data.get("name", "")).strip()
    note = str(data.get("note", "")).strip()
    if not phone or not name:
        return jsonify({"error": "Vui lòng nhập đủ Số điện thoại và Tên khách hàng"}), 400

    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        """
        INSERT INTO customers (phone, name, name_search, telco, note, updated_at)
        VALUES (?, ?, ?, ?, ?, ?)
        ON CONFLICT(phone) DO UPDATE SET
            name = excluded.name,
            name_search = excluded.name_search,
            telco = excluded.telco,
            note = excluded.note,
            updated_at = excluded.updated_at
        """,
        (phone, name, remove_accents(name), detect_telco(phone), note, now_str),
    )
    # Cập nhật tên khách hàng trong bảng orders cho đồng bộ
    cur.execute(
        "UPDATE orders SET customer_name = ?, name_search = ?, updated_at = ? WHERE phone = ?",
        (name, remove_accents(name), now_str, phone),
    )
    conn.commit()
    conn.close()
    return jsonify({"success": True, "phone": phone, "name": name})


@app.route("/api/import-excel", methods=["POST"])
def api_import_excel():
    if "file" not in request.files:
        return jsonify({"error": "Vui lòng chọn file Excel (.xlsx)"}), 400
    f = request.files["file"]
    month_key = request.form.get("month_key", "2026-10").strip()
    mode = request.form.get("mode", "append").strip()  # 'append' | 'replace'
    default_order_type = request.form.get("default_order_type", "OTP").strip()

    try:
        stream = io.BytesIO(f.read())
        result = import_workbook_into_month(stream, month_key=month_key, mode=mode, default_order_type=default_order_type)
        return jsonify({"success": True, **result})
    except Exception as e:
        return jsonify({"error": f"Lỗi đọc file Excel: {str(e)}"}), 400


@app.route("/api/export-excel", methods=["GET"])
def api_export_excel():
    month_key = request.args.get("month_key", "2026-10")
    conn = get_db()
    cur = conn.cursor()

    wb = openpyxl.Workbook()
    # Xóa sheet mặc định
    wb.remove(wb.active)

    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    # 1. Sheet Danh Sách Tổng KH
    s_kh = wb.create_sheet(title="Danh Sách Tổng KH")
    cur.execute("SELECT phone, name FROM customers ORDER BY name ASC")
    for idx, r in enumerate(cur.fetchall(), start=1):
        p_num = int(raw_phone_no_zero(r["phone"])) if raw_phone_no_zero(r["phone"]).isdigit() else r["phone"]
        s_kh.cell(idx, 1, p_num)
        s_kh.cell(idx, 2, r["name"])
    s_kh.column_dimensions["A"].width = 16
    s_kh.column_dimensions["B"].width = 36

    # 2. Sheet Danh Sách Đơn Chưa Cho Chạy
    s_cc = wb.create_sheet(title="Danh Sách Đơn Chưa Cho Chạy")
    s_cc.cell(1, 1, f"ĐƠN CHƯA CHO CHẠY TRONG THÁNG {month_key}")
    headers_cc = ["SĐT", "Tên Khách Hàng", "Trạng Thái", "Số Tiền", "Cho Chạy", "Thành công"]
    for c_idx, h in enumerate(headers_cc, start=1):
        cell = s_cc.cell(2, c_idx, h)
        cell.font = header_font
        cell.fill = header_fill

    cur.execute("SELECT * FROM orders WHERE month_key = ? AND order_type = 'CHUA_CHAY' ORDER BY id ASC", (month_key,))
    for r_idx, r in enumerate(cur.fetchall(), start=3):
        p_num = int(raw_phone_no_zero(r["phone"])) if raw_phone_no_zero(r["phone"]).isdigit() else r["phone"]
        s_cc.cell(r_idx, 1, p_num)
        s_cc.cell(r_idx, 2, r["customer_name"])
        s_cc.cell(r_idx, 3, "Chưa cho lên đơn" if r["status"] != "COMPLETED" else "Đã lên đơn")
        s_cc.cell(r_idx, 4, r["amount"])
        s_cc.cell(r_idx, 5, r["batch_code"])
        s_cc.cell(r_idx, 6, "Đã lên đơn" if r["status"] == "COMPLETED" else r["note"])

    # 3. Sheet Trạng thái Đơn OTP
    s_otp = wb.create_sheet(title="Trạng thái Đơn OTP")
    s_otp.cell(1, 1, f"TRANG KIỂM TRA TRẠNG THÁI CHO CHẠY OTP TRONG THÁNG {month_key}")
    headers_otp = ["SĐT", "Tên Khách Hàng", "Trạng Thái", "Số Tiền", "Thành công", "Lệnh Nạp (SĐT,Tiền)", "Mã OTP / Ghi chú"]
    for c_idx, h in enumerate(headers_otp, start=1):
        cell = s_otp.cell(2, c_idx, h)
        cell.font = header_font
        cell.fill = header_fill

    cur.execute("SELECT * FROM orders WHERE month_key = ? AND order_type = 'OTP' ORDER BY id ASC", (month_key,))
    for r_idx, r in enumerate(cur.fetchall(), start=3):
        p_no_zero = raw_phone_no_zero(r["phone"])
        p_num = int(p_no_zero) if p_no_zero.isdigit() else r["phone"]
        if r["status"] == "COMPLETED":
            st_text = "Đã nạp"
            col_e = "x"
        elif r["otp_status"] == "HAS_OTP":
            st_text = "Đã có OTP"
            col_e = r["batch_code"] or "Đã có OTP"
        else:
            st_text = "Lấy OTP"
            col_e = r["batch_code"] or ""

        s_otp.cell(r_idx, 1, p_num)
        s_otp.cell(r_idx, 2, r["customer_name"])
        s_otp.cell(r_idx, 3, st_text)
        s_otp.cell(r_idx, 4, r["amount"])
        s_otp.cell(r_idx, 5, col_e)
        s_otp.cell(r_idx, 6, f"{p_no_zero},{r['amount']}")
        s_otp.cell(r_idx, 7, f"{r['otp_code']} {r['note']}".strip())

    # 4. Sheet Trạng thái Đơn Nạp hộ
    s_nh = wb.create_sheet(title="Trạng thái Đơn Nạp hộ")
    s_nh.cell(1, 1, f"TRANG KIỂM TRA TRẠNG THÁI CHO CHẠY NẠP HỘ TRONG THÁNG {month_key}")
    headers_nh = ["SĐT", "Tên Khách Hàng", "Trạng Thái", "Số Tiền", "Thành công", "Lệnh Nạp (SĐT,Tiền)"]
    for c_idx, h in enumerate(headers_nh, start=1):
        cell = s_nh.cell(2, c_idx, h)
        cell.font = header_font
        cell.fill = header_fill

    cur.execute("SELECT * FROM orders WHERE month_key = ? AND order_type = 'NAP_HO' ORDER BY id ASC", (month_key,))
    for r_idx, r in enumerate(cur.fetchall(), start=3):
        p_no_zero = raw_phone_no_zero(r["phone"])
        p_num = int(p_no_zero) if p_no_zero.isdigit() else r["phone"]
        st_text = "Đã nạp" if r["status"] == "COMPLETED" else "Chưa nạp"
        col_e = "x" if r["status"] == "COMPLETED" else ""
        s_nh.cell(r_idx, 1, p_num)
        s_nh.cell(r_idx, 2, r["customer_name"])
        s_nh.cell(r_idx, 3, st_text)
        s_nh.cell(r_idx, 4, r["amount"])
        s_nh.cell(r_idx, 5, col_e)
        s_nh.cell(r_idx, 6, f"{p_no_zero},{r['amount']}")

    for sheet in [s_cc, s_otp, s_nh]:
        sheet.column_dimensions["A"].width = 16
        sheet.column_dimensions["B"].width = 34
        sheet.column_dimensions["C"].width = 18
        sheet.column_dimensions["D"].width = 15
        sheet.column_dimensions["E"].width = 14
        sheet.column_dimensions["F"].width = 22

    conn.close()
    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    filename = f"Quan_Ly_Nap_Tien_{month_key}.xlsx"
    return send_file(
        output,
        as_attachment=True,
        download_name=filename,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.route("/api/template-excel", methods=["GET"])
def api_template_excel():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Mau_Nhap_Don_Thang"
    headers = ["SĐT", "Tên Khách Hàng", "Số Tiền", "Hình Thức (OTP / Nạp hộ)", "Trạng Thái (Lấy OTP / Đã có OTP / Đã nạp)", "Ghi chú"]
    ws.append(headers)
    ws.append(["0907602719", "Nguyễn Thị Thùy Loan", 160000, "OTP", "Lấy OTP", "Đợt 1"])
    ws.append(["0909320563", "Nguyễn Thị Ái Vi", 300000, "OTP", "Đã có OTP", "Khách đã gửi OTP"])
    ws.append(["0932929119", "Phạm Hoài Nam", 200000, "Nạp hộ", "Chưa nạp", ""])
    for col in ["A", "B", "C", "D", "E", "F"]:
        ws.column_dimensions[col].width = 25

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return send_file(
        output,
        as_attachment=True,
        download_name="File_Mau_Nhap_Don_Hang_Thang.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


BC26_LAST_SYNC = {
    "last_sync_at": None,
    "total_Read": 0,
    "updated_count": 0,
    "need_otp_count": 0,
    "has_otp_count": 0,
    "completed_count": 0,
    "sync_version": 0,
}


@app.after_request
def add_cors_headers(response):
    response.headers["Access-Control-Allow-Origin"] = "*"
    response.headers["Access-Control-Allow-Methods"] = "GET, POST, PATCH, DELETE, OPTIONS"
    response.headers["Access-Control-Allow-Headers"] = "Content-Type, Authorization, X-Requested-With"
    response.headers["Access-Control-Allow-Private-Network"] = "true"
    return response


@app.route("/api/bc26/sync", methods=["POST", "OPTIONS"])
def api_bc26_sync():
    if request.method == "OPTIONS":
        return jsonify({"ok": True})

    data = request.json or {}
    items = data.get("items", [])
    month_key = data.get("month_key") or "2026-10"

    if not items:
        return jsonify({"success": False, "error": "Không có dữ liệu dòng nào từ bc26.pro"}), 400

    conn = get_db()
    cur = conn.cursor()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # Lấy tháng mới nhất nếu không truyền month_key
    cur.execute("SELECT month_key FROM months ORDER BY month_key DESC LIMIT 1")
    m_row = cur.fetchone()
    if m_row and not data.get("month_key"):
        month_key = m_row["month_key"]

    updated_count = 0
    need_cnt = 0
    has_cnt = 0
    done_cnt = 0

    for item in items:
        phone = normalize_phone(item.get("phone", ""))
        if not phone or len(phone) < 9:
            continue

        bc_state = str(item.get("state", "")).strip()  # 'NEED_OTP' | 'HAS_OTP' | 'COMPLETED'
        paid_amt = parse_amount(item.get("paid_amount", 0))
        total_amt = parse_amount(item.get("amount", 0))
        bc_note = str(item.get("status_text", "")).strip()

        if bc_state == "COMPLETED":
            new_status = "COMPLETED"
            new_otp_status = "DONE_OTP"
            done_cnt += 1
        elif bc_state == "HAS_OTP":
            new_status = "PENDING"
            new_otp_status = "HAS_OTP"
            has_cnt += 1
        else:
            new_status = "PENDING"
            new_otp_status = "NEED_OTP"
            need_cnt += 1

        # Tìm đơn OTP (hoặc Nạp hộ) khớp SĐT trong tháng hiện tại
        cur.execute(
            "SELECT id, order_type, status, otp_status, note FROM orders WHERE month_key = ? AND phone = ? ORDER BY id DESC",
            (month_key, phone),
        )
        matched_rows = cur.fetchall()
        for r in matched_rows:
            if r["order_type"] == "OTP":
                if r["status"] != new_status or r["otp_status"] != new_otp_status or (bc_note and r["note"] != bc_note):
                    cur.execute(
                        "UPDATE orders SET status = ?, otp_status = ?, note = ?, updated_at = ? WHERE id = ?",
                        (new_status, new_otp_status, bc_note or r["note"], now_str, r["id"]),
                    )
                    updated_count += 1
            elif r["order_type"] == "NAP_HO" and new_status == "COMPLETED":
                if r["status"] != "COMPLETED":
                    cur.execute(
                        "UPDATE orders SET status = 'COMPLETED', batch_code = 'x', updated_at = ? WHERE id = ?",
                        (now_str, r["id"]),
                    )
                    updated_count += 1

    conn.commit()
    conn.close()

    BC26_LAST_SYNC["last_sync_at"] = datetime.now().strftime("%H:%M:%S %d/%m/%Y")
    BC26_LAST_SYNC["total_read"] = len(items)
    BC26_LAST_SYNC["updated_count"] = updated_count
    BC26_LAST_SYNC["need_otp_count"] = need_cnt
    BC26_LAST_SYNC["has_otp_count"] = has_cnt
    BC26_LAST_SYNC["completed_count"] = done_cnt
    BC26_LAST_SYNC["sync_version"] = BC26_LAST_SYNC.get("sync_version", 0) + 1

    return jsonify({"success": True, **BC26_LAST_SYNC})


@app.route("/api/bc26/status", methods=["GET"])
def api_bc26_status():
    return jsonify(BC26_LAST_SYNC)


# Khởi tạo database ngay khi ứng dụng được nạp (hỗ trợ cả Gunicorn trên Render và chạy trực tiếp)
try:
    init_db()
except Exception as _e:
    print("DB init warning:", _e)


if __name__ == "__main__":
    import sys
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass
    port = int(os.environ.get("PORT", 5000))
    print("====================================================")
    print(" HE THONG QUAN LY DON NAP TIEN & THEO DOI OTP")
    print(f" Truy cap tai: http://127.0.0.1:{port}")
    print("====================================================")
    app.run(host="0.0.0.0", port=port, debug=False)
