from flask import Flask, render_template, request, redirect, session, flash, url_for, jsonify, send_file
from itsdangerous import URLSafeTimedSerializer, BadSignature, SignatureExpired
from dotenv import load_dotenv
import sqlite3
from datetime import datetime, date, timedelta
import os
from io import BytesIO
import json
import uuid
import qrcode
import smtplib
from email.message import EmailMessage
from werkzeug.utils import secure_filename
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

load_dotenv()

app = Flask(__name__)
app.secret_key = "gojames-secret-key"

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, "database.db")
STATIC_DIR = os.path.join(BASE_DIR, "static")
UPLOAD_DIR = os.path.join(STATIC_DIR, "slips")

os.makedirs(STATIC_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)

ALLOWED_EXTENSIONS = {"png", "jpg", "jpeg", "webp"}

# ตั้งค่าอีเมลก่อนใช้งานจริง
# EMAIL_PASSWORD ต้องใช้ App Password ของ Gmail (ไม่ใช่รหัสผ่านปกติ)
EMAIL_SENDER = os.environ.get("EMAIL_SENDER", "")
EMAIL_PASSWORD = os.environ.get("EMAIL_PASSWORD", "")
BASE_URL = os.environ.get("BASE_URL", "http://127.0.0.1:5000").rstrip("/")

serializer = URLSafeTimedSerializer(app.secret_key)

# PromptPay สำหรับรับเงินจริง
# ถ้าเอาขึ้น GitHub/Render ภายหลัง แนะนำให้ย้ายไปใช้ Environment Variable แทน
PROMPTPAY_ID = os.environ.get("PROMPTPAY_ID", "0953087210")

BANKS = {
    "promptpay": {"name": "PromptPay", "short": "PP", "color": "#1f8bff", "prefixes": []},
    "kbank": {"name": "กสิกรไทย", "short": "KB", "color": "#0f9d58", "prefixes": ["004"]},
    "scb": {"name": "ไทยพาณิชย์", "short": "SCB", "color": "#6f42c1", "prefixes": ["014"]},
    "bbl": {"name": "กรุงเทพ", "short": "BBL", "color": "#1d4ed8", "prefixes": ["002"]},
    "ktb": {"name": "กรุงไทย", "short": "KTB", "color": "#00a2ff", "prefixes": ["006"]},
    "ttb": {"name": "ttb", "short": "TTB", "color": "#f97316", "prefixes": ["011"]},
    "bay": {"name": "กรุงศรีอยุธยา", "short": "BAY", "color": "#facc15", "prefixes": ["025"]},
    "gsb": {"name": "ออมสิน", "short": "GSB", "color": "#ec4899", "prefixes": ["030"]},
    "baac": {"name": "ธ.ก.ส.", "short": "BAAC", "color": "#22c55e", "prefixes": ["034"]},
    "cimb": {"name": "CIMB Thai", "short": "CIMB", "color": "#ef4444", "prefixes": ["018"]},
    "lhbank": {"name": "LH Bank", "short": "LH", "color": "#6b7280", "prefixes": ["073"]},
    "uob": {"name": "UOB", "short": "UOB", "color": "#1e3a8a", "prefixes": ["024"]},
    "icbc": {"name": "ICBC Thai", "short": "ICBC", "color": "#dc2626", "prefixes": ["070"]},
    "kkp": {"name": "เกียรตินาคินภัทร", "short": "KKP", "color": "#0f766e", "prefixes": ["069"]},
}



def _pp_tlv(tag, value):
    value = str(value)
    return f"{tag}{len(value):02d}{value}"


def _pp_crc16(payload):
    crc = 0xFFFF
    for ch in payload.encode("utf-8"):
        crc ^= ch << 8
        for _ in range(8):
            if crc & 0x8000:
                crc = ((crc << 1) ^ 0x1021) & 0xFFFF
            else:
                crc = (crc << 1) & 0xFFFF
    return f"{crc:04X}"


def _format_promptpay_target(target):
    digits = "".join(ch for ch in str(target) if ch.isdigit())
    if len(digits) == 10 and digits.startswith("0"):   # เบอร์มือถือไทย
        return "0066" + digits[1:]
    if len(digits) == 13:  # เลขบัตรประชาชน
        return digits
    raise ValueError("รูปแบบ PromptPay ID ไม่ถูกต้อง")


def generate_promptpay_payload(target, amount):
    formatted_target = _format_promptpay_target(target)

    # 01 = เบอร์โทรศัพท์, 02 = เลขบัตรประชาชน
    target_tag = "01" if formatted_target.startswith("0066") else "02"

    merchant_account_info = (
        _pp_tlv("00", "A000000677010111") +
        _pp_tlv(target_tag, formatted_target)
    )

    payload = (
        _pp_tlv("00", "01") +          # Payload Format Indicator
        _pp_tlv("01", "12") +          # Point of Initiation Method (Dynamic QR)
        _pp_tlv("29", merchant_account_info) +
        _pp_tlv("52", "0000") +        # Merchant Category Code
        _pp_tlv("53", "764") +         # Currency: THB
        _pp_tlv("54", f"{float(amount):.2f}") +
        _pp_tlv("58", "TH") +
        "6304"
    )
    return payload + _pp_crc16(payload)

ROOMS = [
    {
        "id": 1,
        "name": "ห้อง Lavender",
        "price": 10,
        "detail": "ห้องพักเตียงคู่สำหรับ 2 ท่าน บรรยากาศอบอุ่น",
        "cover": "room1_1.jpg",
        "images": ["room1_1.jpg", "room1_2.jpg", "room1_3.jpg"]
    },
    {
        "id": 2,
        "name": "ห้อง Lilac",
        "price": 800,
        "detail": "ห้องพักวิวสวน เงียบสงบ เหมาะกับการพักผ่อน",
        "cover": "room2_1.jpg",
        "images": ["room2_1.jpg", "room2_2.jpg", "room2_3.jpg"]
    },
    {
        "id": 3,
        "name": "ห้อง Orchid",
        "price": 800,
        "detail": "ห้องพักขนาดกลาง พร้อมพื้นที่นั่งเล่นเล็ก ๆ",
        "cover": "room3_1.jpg",
        "images": ["room3_1.jpg", "room3_2.jpg", "room3_3.jpg"]
    },
    {
        "id": 4,
        "name": "ห้อง Iris",
        "price": 800,
        "detail": "ห้องพักสำหรับครอบครัวขนาดเล็ก ดูเรียบง่ายสบายตา",
        "cover": "room4_real.jpg",
        "images": ["room4_1.svg", "room4_2.svg", "room4_3.svg"],
        "is_open": False,
        "status_text": "ยังไม่พร้อมให้บริการ",
        "maintenance_note": "ห้องนี้ยังไม่พร้อมให้บริการ และจะเปิดให้จองในภายหลัง"
    },
    {
        "id": 5,
        "name": "ห้อง Violet",
        "price": 800,
        "detail": "ห้องพรีเมียม พร้อมระเบียงและพื้นที่กว้างขึ้น",
        "cover": "room5_real.jpg",
        "images": ["room5_1.svg", "room5_2.svg", "room5_3.svg"],
        "is_open": False,
        "status_text": "ยังไม่พร้อมให้บริการ",
        "maintenance_note": "ห้องนี้ยังไม่พร้อมให้บริการ และจะเปิดให้จองในภายหลัง"
    }
]

BOOKED_STATUSES = ("approved", "waiting_slip_approval", "deposit_paid", "remaining_counter_pending", "fully_paid", "cancel_requested")


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def get_room_status_map():
    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT room_id, is_open FROM room_settings")
    rows = cur.fetchall()
    conn.close()
    return {row["room_id"]: bool(row["is_open"]) for row in rows}


def get_rooms_with_status():
    status_map = get_room_status_map()
    rooms = []
    for room in ROOMS:
        room_copy = dict(room)
        default_open = room_copy.get("is_open", True)
        room_copy["is_open"] = status_map.get(room_copy["id"], default_open)
        rooms.append(room_copy)
    return rooms


def get_now():
    return datetime.now()


def db_now_str():
    return get_now().strftime("%Y-%m-%d %H:%M:%S.%f")


def allowed_file(filename):
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def normalize_digits(value):
    return "".join(ch for ch in str(value or "") if ch.isdigit())


def get_bank_meta(bank_code):
    meta = BANKS.get((bank_code or "").lower())
    if meta:
        return meta
    return {"name": bank_code or "-", "short": (bank_code or "-")[:4].upper(), "color": "#6b7280", "prefixes": []}


def detect_bank_from_account(account_number):
    digits = normalize_digits(account_number)
    if len(digits) in (10, 13) and digits.startswith("0"):
        return "promptpay"
    prefix = digits[:3]
    for code, meta in BANKS.items():
        if code == "promptpay":
            continue
        if prefix in meta.get("prefixes", []):
            return code
    return ""


def validate_refund_destination(refund_method, refund_account, refund_bank=""):
    method = (refund_method or "").lower().strip()
    digits = normalize_digits(refund_account)

    if method == "promptpay":
        if (len(digits) == 10 and digits.startswith("0")) or len(digits) == 13:
            return True, "promptpay", digits, ""
        return False, "", digits, "PromptPay ต้องเป็นเบอร์โทร 10 หลักหรือเลขบัตรประชาชน 13 หลัก"

    if method == "bank":
        if not digits.isdigit() or len(digits) < 10 or len(digits) > 12:
            return False, "", digits, "เลขบัญชีธนาคารต้องเป็นตัวเลข 10-12 หลัก"
        detected_bank = detect_bank_from_account(digits)
        selected_bank = (refund_bank or "").lower().strip()
        final_bank = selected_bank or detected_bank
        if not final_bank:
            return False, "", digits, "ไม่สามารถตรวจจับธนาคารจากเลขบัญชีนี้ได้ กรุณาเลือกธนาคาร"
        if selected_bank and detected_bank and selected_bank != detected_bank:
            return False, "", digits, "เลขบัญชีนี้ดูเหมือนจะไม่ตรงกับธนาคารที่เลือก กรุณาตรวจสอบอีกครั้ง"
        return True, final_bank, digits, ""

    return False, "", digits, "กรุณาเลือกวิธีรับเงินคืน"


app.jinja_env.globals["get_bank_meta"] = get_bank_meta


def send_email_notification(to_email, subject, body):
    if not EMAIL_SENDER or not EMAIL_PASSWORD:
        return

    try:
        msg = EmailMessage()
        msg["Subject"] = subject
        msg["From"] = EMAIL_SENDER
        msg["To"] = to_email
        msg.set_content(body)

        with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
            smtp.login(EMAIL_SENDER, EMAIL_PASSWORD)
            smtp.send_message(msg)
    except Exception as e:
        print("ส่งอีเมลไม่สำเร็จ:", e)


def send_reset_password_email(user_row):
    token = serializer.dumps({"user_id": user_row["id"], "email": user_row["email"]}, salt="reset-password")
    reset_link = f"{BASE_URL}{url_for('reset_password', token=token)}"

    subject = "รีเซ็ตรหัสผ่าน - GoJames Vacation Home"
    body = (
        f"สวัสดี {user_row['username']}\n\n"
        "มีการขอรีเซ็ตรหัสผ่านสำหรับบัญชี GoJames Vacation Home ของคุณ\n"
        "กดลิงก์ด้านล่างเพื่อตั้งรหัสผ่านใหม่\n\n"
        f"{reset_link}\n\n"
        "ลิงก์นี้จะหมดอายุภายใน 30 นาที\n"
        "หากคุณไม่ได้เป็นผู้ขอรีเซ็ต สามารถละเว้นอีเมลฉบับนี้ได้"
    )
    send_email_notification(user_row["email"], subject, body)


def create_notification(username, title, message):
    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO notifications(username, title, message, is_read, created_at)
        VALUES(?,?,?,?,?)
    """, (username, title, message, 0, db_now_str()))
    conn.commit()
    conn.close()


def notify_user(username, title, message, email_subject=None, email_body=None):
    create_notification(username, title, message)

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT email FROM users WHERE username = ?", (username,))
    user = cur.fetchone()
    conn.close()

    if user and user["email"]:
        send_email_notification(
            user["email"],
            email_subject if email_subject else title,
            email_body if email_body else message
        )


def notify_admin(title, message, email_subject=None, email_body=None):
    notify_user("admin", title, message, email_subject=email_subject, email_body=email_body)


@app.context_processor
def inject_notification_count():
    unread_count = 0
    latest_notifications = []
    user_nav_badge_count = 0
    admin_nav_badge_count = 0

    if session.get("user"):
        conn = get_db()
        cur = conn.cursor()

        # badge ของ "แจ้งเตือน" ใช้จาก unread ปกติ
        cur.execute("""
            SELECT COUNT(*) AS total FROM notifications
            WHERE username = ? AND is_read = 0
        """, (session["user"],))
        row = cur.fetchone()
        unread_count = row["total"] if row else 0

        cur.execute("""
            SELECT * FROM notifications
            WHERE username = ?
            ORDER BY id DESC
            LIMIT 5
        """, (session["user"],))
        latest_notifications = cur.fetchall()

        # badge ของเมนู "การจองของฉัน" / "จัดการการจอง"
        cur.execute("""
            SELECT last_seen_user_notification_id, last_seen_admin_notification_id
            FROM users
            WHERE username = ?
        """, (session["user"],))
        seen_row = cur.fetchone()

        if session.get("role") == "admin":
            last_seen_admin_id = seen_row["last_seen_admin_notification_id"] if seen_row else 0
            cur.execute("""
                SELECT COUNT(*) AS total FROM notifications
                WHERE username = 'admin' AND id > ?
            """, (last_seen_admin_id or 0,))
            row = cur.fetchone()
            admin_nav_badge_count = row["total"] if row else 0
        else:
            last_seen_user_id = seen_row["last_seen_user_notification_id"] if seen_row else 0
            cur.execute("""
                SELECT COUNT(*) AS total FROM notifications
                WHERE username = ? AND id > ?
            """, (session["user"], last_seen_user_id or 0))
            row = cur.fetchone()
            user_nav_badge_count = row["total"] if row else 0
        conn.close()

    return {
        "unread_notification_count": unread_count,
        "latest_notifications": latest_notifications,
        "user_nav_badge_count": user_nav_badge_count,
        "admin_nav_badge_count": admin_nav_badge_count
    }

def get_live_badge_payload():
    unread_count = 0
    user_nav_badge_count = 0
    admin_nav_badge_count = 0

    if session.get("user"):
        conn = get_db()
        cur = conn.cursor()

        cur.execute("""
            SELECT COUNT(*) AS total FROM notifications
            WHERE username = ? AND is_read = 0
        """, (session["user"],))
        row = cur.fetchone()
        unread_count = row["total"] if row else 0

        cur.execute("""
            SELECT last_seen_user_bookings, last_seen_admin_bookings
            FROM users
            WHERE username = ?
        """, (session["user"],))
        seen_row = cur.fetchone()

        if session.get("role") == "admin":
            last_seen_admin = seen_row["last_seen_admin_bookings"] if seen_row else None
            if last_seen_admin:
                cur.execute("""
                    SELECT COUNT(*) AS total FROM notifications
                    WHERE username = 'admin' AND created_at > ?
                """, (last_seen_admin,))
            else:
                cur.execute("""
                    SELECT COUNT(*) AS total FROM notifications
                    WHERE username = 'admin'
                """)
            row = cur.fetchone()
            admin_nav_badge_count = row["total"] if row else 0
        else:
            last_seen_user = seen_row["last_seen_user_bookings"] if seen_row else None
            if last_seen_user:
                cur.execute("""
                    SELECT COUNT(*) AS total FROM notifications
                    WHERE username = ? AND created_at > ?
                """, (session["user"], last_seen_user))
            else:
                cur.execute("""
                    SELECT COUNT(*) AS total FROM notifications
                    WHERE username = ?
                """, (session["user"],))
            row = cur.fetchone()
            user_nav_badge_count = row["total"] if row else 0

        conn.close()

    return {
        "unread_notification_count": unread_count,
        "user_nav_badge_count": user_nav_badge_count,
        "admin_nav_badge_count": admin_nav_badge_count
    }


@app.route("/api/live-badges")
def api_live_badges():
    if not session.get("user"):
        return jsonify({"ok": False, "logged_in": False})
    payload = get_live_badge_payload()
    payload["ok"] = True
    payload["logged_in"] = True
    return jsonify(payload)

@app.before_request
def run_auto_cancel_before_each_request():
    auto_cancel_expired_bookings()
    remind_before_checkin()



def remind_before_checkin():
    conn = get_db()
    cur = conn.cursor()
    now_dt = get_now()
    upcoming_limit = now_dt + timedelta(days=1)

    cur.execute("""
        SELECT * FROM bookings
        WHERE status IN ('deposit_paid', 'remaining_counter_pending')
          AND checkin_reminder_sent_at IS NULL
    """)
    rows = cur.fetchall()

    changed_ids = []
    for booking in rows:
        try:
            checkin_dt = datetime.strptime(f"{booking['checkin']} {booking['checkin_time']}", "%Y-%m-%d %H:%M")
        except ValueError:
            continue

        if now_dt <= checkin_dt <= upcoming_limit:
            notify_user(
                booking["username"],
                "แจ้งเตือนก่อนวันเข้าพัก",
                f"รายการจองห้อง {booking['room_name']} ของคุณใกล้ถึงวันเข้าพักแล้ว และยังคงเหลือยอดชำระ {booking['remaining_due']} บาท กรุณาชำระส่วนที่เหลือก่อนเข้าพัก"
            )
            changed_ids.append(booking["id"])

    for booking_id in changed_ids:
        cur.execute("""
            UPDATE bookings
            SET checkin_reminder_sent_at = ?
            WHERE id = ?
        """, (now_dt.strftime("%Y-%m-%d %H:%M:%S"), booking_id))

    conn.commit()
    conn.close()


def auto_cancel_expired_bookings():
    conn = get_db()
    cur = conn.cursor()
    now_str = get_now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute("""
        SELECT * FROM bookings
        WHERE status = 'approved'
          AND payment_deadline IS NOT NULL
          AND payment_deadline < ?
    """, (now_str,))
    expired_bookings = cur.fetchall()

    if not expired_bookings:
        conn.close()
        return

    for booking in expired_bookings:
        cur.execute("UPDATE bookings SET status = 'cancelled' WHERE id = ?", (booking["id"],))

    conn.commit()
    conn.close()

    for booking in expired_bookings:
        notify_user(
            booking["username"],
            "การจองถูกยกเลิกอัตโนมัติ",
            f"รายการจองห้อง {booking['room_name']} ของคุณถูกยกเลิกอัตโนมัติ เพราะไม่ได้ชำระเงินภายในเวลาที่กำหนด",
            email_subject="การจองถูกยกเลิกอัตโนมัติ - โกเจมส์ ที่พักตากอากาศ",
            email_body=(
                f"รายการจองห้อง {booking['room_name']} ของคุณถูกยกเลิกอัตโนมัติ\n"
                f"เนื่องจากไม่ได้ชำระเงินภายในเวลาที่กำหนด ({booking['payment_deadline']})"
            )
        )


def init_db():
    conn = get_db()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        first_name TEXT NOT NULL,
        last_name TEXT NOT NULL,
        nickname TEXT NOT NULL,
        gender TEXT NOT NULL,
        phone TEXT NOT NULL,
        email TEXT NOT NULL,
        username TEXT UNIQUE NOT NULL,
        password TEXT NOT NULL,
        role TEXT NOT NULL DEFAULT 'user',
        last_seen_user_bookings TEXT,
        last_seen_admin_bookings TEXT
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS room_settings(
        room_id INTEGER PRIMARY KEY,
        is_open INTEGER NOT NULL DEFAULT 1
    )
    """)

    for room in ROOMS:
        cur.execute(
            "INSERT OR IGNORE INTO room_settings(room_id, is_open) VALUES (?, ?)",
            (room["id"], 1 if room.get("is_open", True) else 0)
        )

    cur.execute("""
    CREATE TABLE IF NOT EXISTS bookings(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL,
        room_id INTEGER NOT NULL,
        room_name TEXT NOT NULL,
        checkin TEXT NOT NULL,
        checkout TEXT NOT NULL,
        checkin_time TEXT NOT NULL,
        checkout_time TEXT NOT NULL,
        nights INTEGER NOT NULL,
        total REAL NOT NULL,
        paytype TEXT NOT NULL,
        paid REAL NOT NULL,
        status TEXT NOT NULL DEFAULT 'waiting_approval',
        created_at TEXT,
        approved_at TEXT,
        payment_deadline TEXT,
        slip_filename TEXT,
        batch_id TEXT,
        cancel_requested_from TEXT,
        cancel_requested_at TEXT,
        refund_completed_at TEXT,
        remaining_due REAL,
        checkin_reminder_sent_at TEXT,
        remaining_payment_method TEXT DEFAULT 'qr'
    )
    """)

    booking_columns = {row["name"] for row in cur.execute("PRAGMA table_info(bookings)").fetchall()}
    if "batch_id" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN batch_id TEXT")
    if "cancel_requested_from" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN cancel_requested_from TEXT")
    if "cancel_requested_at" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN cancel_requested_at TEXT")
    if "refund_completed_at" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN refund_completed_at TEXT")
    if "refund_method" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN refund_method TEXT")
    if "refund_bank" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN refund_bank TEXT")
    if "refund_account" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN refund_account TEXT")
    if "refund_amount" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN refund_amount REAL DEFAULT 0")
    if "refund_note" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN refund_note TEXT")
    if "cancelled_by" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN cancelled_by TEXT")
    if "remaining_due" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN remaining_due REAL DEFAULT 0")
    if "checkin_reminder_sent_at" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN checkin_reminder_sent_at TEXT")
    if "remaining_payment_method" not in booking_columns:
        cur.execute("ALTER TABLE bookings ADD COLUMN remaining_payment_method TEXT DEFAULT 'qr'")

    cur.execute("""
    CREATE TABLE IF NOT EXISTS room_blocks(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        room_id INTEGER NOT NULL,
        start_date TEXT NOT NULL,
        end_date TEXT NOT NULL,
        note TEXT,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS notifications(
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT NOT NULL,
        title TEXT NOT NULL,
        message TEXT NOT NULL,
        is_read INTEGER NOT NULL DEFAULT 0,
        created_at TEXT NOT NULL
    )
    """)

    user_columns = {row["name"] for row in cur.execute("PRAGMA table_info(users)").fetchall()}
    if "last_seen_user_bookings" not in user_columns:
        cur.execute("ALTER TABLE users ADD COLUMN last_seen_user_bookings TEXT")
    if "last_seen_admin_bookings" not in user_columns:
        cur.execute("ALTER TABLE users ADD COLUMN last_seen_admin_bookings TEXT")
    if "last_seen_user_notification_id" not in user_columns:
        cur.execute("ALTER TABLE users ADD COLUMN last_seen_user_notification_id INTEGER DEFAULT 0")
    if "last_seen_admin_notification_id" not in user_columns:
        cur.execute("ALTER TABLE users ADD COLUMN last_seen_admin_notification_id INTEGER DEFAULT 0")

    cur.execute("SELECT id FROM users WHERE username = ?", ("admin",))
    if cur.fetchone() is None:
        cur.execute("""
            INSERT INTO users(
                first_name, last_name, nickname, gender, phone, email,
                username, password, role
            )
            VALUES(?,?,?,?,?,?,?,?,?)
        """, (
            "Admin", "System", "admin", "อื่นๆ", "0000000000",
            "admin@email.com", "admin", "admin123", "admin"
        ))

    conn.commit()
    conn.close()


def get_room(room_id):
    for room in get_rooms_with_status():
        if room["id"] == room_id:
            return room
    return None


def create_qr(booking_id, amount):
    payload = generate_promptpay_payload(PROMPTPAY_ID, amount)
    img = qrcode.make(payload)
    filename = f"qr_{booking_id}.png"
    path = os.path.join(STATIC_DIR, filename)
    img.save(path)
    return filename


def is_overlap(start1, end1, start2, end2):
    return start1 < end2 and start2 < end1


def block_overlaps_booking(booking_start, booking_end, block_start, block_end):
    block_start_dt = datetime.combine(block_start, datetime.min.time())
    block_end_dt = datetime.combine(block_end + timedelta(days=1), datetime.min.time())
    return is_overlap(booking_start, booking_end, block_start_dt, block_end_dt)


def calculate_charge_nights(checkin_dt, checkout_dt):
    day_diff = (checkout_dt.date() - checkin_dt.date()).days
    return max(1, day_diff)


def calculate_due_now(total_amount, paytype):
    return float(total_amount) if paytype == "full" else round(float(total_amount) * 0.25, 2)


def calculate_remaining_due(total_amount, paytype):
    due_now = calculate_due_now(total_amount, paytype)
    return round(float(total_amount) - float(due_now), 2)


def calculate_refund_amount(booking, cancelled_by="user"):
    paid = float(booking["paid"] or 0)
    paytype = booking["paytype"]
    if cancelled_by == "admin":
        return round(paid, 2)
    if paytype == "full":
        return round(paid * 0.5, 2)
    return 0.0




def choose_group_status(rows):
    statuses = [row["status"] for row in rows if row["status"]]
    if not statuses:
        return ""
    if len(set(statuses)) == 1:
        return statuses[0]

    priority = [
        "cancel_requested",
        "refund_pending",
        "refunded",
        "waiting_remaining_slip_approval",
        "remaining_counter_pending",
        "deposit_paid",
        "waiting_slip_approval",
        "approved",
        "waiting_approval",
        "fully_paid",
        "rejected",
        "cancelled",
    ]
    for status in priority:
        if status in statuses:
            return status
    return statuses[0]


def build_booking_groups(rows):
    grouped = {}
    for row in rows:
        key = row["batch_id"] if row["batch_id"] else f"single-{row['id']}"
        grouped.setdefault(key, []).append(row)

    countable_statuses = {
        "waiting_approval",
        "approved",
        "waiting_slip_approval",
        "deposit_paid",
        "waiting_remaining_slip_approval",
        "remaining_counter_pending",
        "fully_paid",
    }

    groups = []
    for key, group_rows in grouped.items():
        ordered = sorted(
            group_rows,
            key=lambda r: datetime.strptime(f"{r['checkin']} {r['checkin_time']}", "%Y-%m-%d %H:%M")
        )
        representative = max(group_rows, key=lambda r: r["id"])
        active_rows = [r for r in ordered if r["status"] in countable_statuses]
        rows_for_totals = active_rows if active_rows else ordered
        first_row = rows_for_totals[0]
        last_row = max(
            rows_for_totals,
            key=lambda r: datetime.strptime(f"{r['checkout']} {r['checkout_time']}", "%Y-%m-%d %H:%M")
        )
        deadlines = [r["payment_deadline"] for r in rows_for_totals if r["payment_deadline"]]
        slip_row = next((r for r in ordered if r["slip_filename"]), representative)

        segments = []
        for r in ordered:
            segments.append({
                "id": r["id"],
                "range_text": f"{r['checkin']} {r['checkin_time']} → {r['checkout']} {r['checkout_time']}",
                "status": r["status"],
                "nights": int(float(r["nights"] or 0)),
                "total": round(float(r["total"] or 0), 2),
                "remaining_due": round(float(r["remaining_due"] or 0), 2),
                "cancel_allowed": r["status"] in {
                    "waiting_approval", "approved", "waiting_slip_approval",
                    "deposit_paid", "remaining_counter_pending", "fully_paid"
                }
            })

        group = dict(representative)
        group.update({
            "id": representative["id"],
            "display_id": representative["id"],
            "action_id": representative["id"],
            "room_name": representative["room_name"],
            "checkin": first_row["checkin"],
            "checkin_time": first_row["checkin_time"],
            "checkout": last_row["checkout"],
            "checkout_time": last_row["checkout_time"],
            "nights": sum(float(r["nights"] or 0) for r in rows_for_totals),
            "total": round(sum(float(r["total"] or 0) for r in rows_for_totals), 2),
            "paid": round(sum(float(r["paid"] or 0) for r in rows_for_totals), 2),
            "remaining_due": round(sum(float(r["remaining_due"] or 0) for r in rows_for_totals), 2),
            "status": choose_group_status(rows_for_totals if rows_for_totals else group_rows),
            "payment_deadline": min(deadlines) if deadlines else None,
            "slip_filename": slip_row["slip_filename"],
            "ranges_text": [segment["range_text"] for segment in segments],
            "segments": segments,
            "segment_count": len(group_rows),
            "active_segment_count": len(active_rows),
            "is_grouped": len(group_rows) > 1,
            "sort_id": representative["id"],
        })
        groups.append(group)

    groups.sort(key=lambda r: r["sort_id"], reverse=True)
    return groups


def get_today_available_time_options():
    now = datetime.now()
    options = []
    for hour in range(24):
        slot_time = now.replace(hour=hour, minute=0, second=0, microsecond=0)
        if slot_time > now:
            options.append(f"{hour:02d}:00")
    return options


def get_all_time_options():
    return [f"{hour:02d}:00" for hour in range(24)]


def get_calendar_data(days=14):
    conn = get_db()
    cur = conn.cursor()

    start_day = date.today()
    day_list = [start_day + timedelta(days=i) for i in range(days)]

    placeholders = ",".join("?" for _ in BOOKED_STATUSES)
    cur.execute(f"""
        SELECT bookings.*, users.first_name, users.last_name
        FROM bookings
        LEFT JOIN users ON bookings.username = users.username
        WHERE bookings.status IN ({placeholders})
    """, BOOKED_STATUSES)
    active_bookings = cur.fetchall()

    cur.execute("SELECT * FROM room_blocks ORDER BY start_date ASC")
    blocks = cur.fetchall()
    conn.close()

    today = date.today()

    room_rows = []
    for room in get_rooms_with_status():
        cells = []
        room_closed = not room.get("is_open", True)

        for day in day_list:
            state = "free"
            detail = None

            if room_closed:
                state = "blocked"
                detail = {
                    "guest_name": "-",
                    "room_name": room["name"],
                    "checkin": day.strftime("%Y-%m-%d"),
                    "checkin_time": "-",
                    "checkout": day.strftime("%Y-%m-%d"),
                    "checkout_time": "-",
                    "status_text": "ปิดปรับปรุง",
                    "note": room.get("maintenance_note") or "-"
                }
            else:
                for block in blocks:
                    if block["room_id"] == room["id"]:
                        block_start = datetime.strptime(block["start_date"], "%Y-%m-%d").date()
                        block_end = datetime.strptime(block["end_date"], "%Y-%m-%d").date()
                        if block_start <= day <= block_end:
                            state = "blocked"
                            detail = {
                                "guest_name": "-",
                                "room_name": room["name"],
                                "checkin": block["start_date"],
                                "checkin_time": "-",
                                "checkout": block["end_date"],
                                "checkout_time": "-",
                                "status_text": "แอดมินบล็อกห้อง",
                                "note": block["note"] if block["note"] else "-"
                            }
                            break

                if state == "free":
                    for booking in active_bookings:
                        if booking["room_id"] == room["id"]:
                            booking_start = datetime.strptime(f"{booking['checkin']} {booking['checkin_time']}", "%Y-%m-%d %H:%M")
                            booking_end = datetime.strptime(f"{booking['checkout']} {booking['checkout_time']}", "%Y-%m-%d %H:%M")
                            day_start = datetime.combine(day, datetime.min.time())
                            day_end = day_start + timedelta(days=1)
                            if is_overlap(booking_start, booking_end, day_start, day_end):
                                if day == today and booking_start <= datetime.now() < booking_end:
                                    state = "occupied"
                                    status_text = "มีผู้พักอยู่"
                                else:
                                    state = "booked"
                                    status_text = "มีผู้จอง"

                                full_name = " ".join(
                                    x for x in [booking["first_name"], booking["last_name"]] if x
                                ).strip()
                                detail = {
                                    "guest_name": full_name if full_name else booking["username"],
                                    "room_name": booking["room_name"],
                                    "checkin": booking["checkin"],
                                    "checkin_time": booking["checkin_time"],
                                    "checkout": booking["checkout"],
                                    "checkout_time": booking["checkout_time"],
                                    "status_text": status_text,
                                    "note": f"รหัสการจอง #{booking['id']}"
                                }
                                break

            cells.append({"date": day, "state": state, "detail": detail})
        room_rows.append({"room": room, "cells": cells})

    return day_list, room_rows


@app.route("/")
def index():
    auto_cancel_expired_bookings()
    return render_template("index.html", rooms=get_rooms_with_status())


@app.route("/register", methods=["GET", "POST"])
def register():
    if request.method == "POST":
        first_name = request.form["first_name"].strip()
        last_name = request.form["last_name"].strip()
        nickname = request.form["nickname"].strip()
        gender = request.form["gender"].strip()
        phone = request.form["phone"].strip()
        email = request.form["email"].strip()
        username = request.form["username"].strip()
        password = request.form["password"].strip()
        confirm_password = request.form["confirm_password"].strip()

        if not all([first_name, last_name, nickname, gender, phone, email, username, password, confirm_password]):
            flash("กรอกข้อมูลให้ครบ")
            return redirect("/register")

        if password != confirm_password:
            flash("รหัสผ่านและยืนยันรหัสผ่านไม่ตรงกัน")
            return redirect("/register")

        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT id FROM users WHERE username = ?", (username,))
        if cur.fetchone():
            conn.close()
            flash("ชื่อผู้ใช้นี้มีอยู่แล้ว")
            return redirect("/register")

        cur.execute("""
            INSERT INTO users(first_name, last_name, nickname, gender, phone, email, username, password, role)
            VALUES(?,?,?,?,?,?,?,?,?)
        """, (first_name, last_name, nickname, gender, phone, email, username, password, "user"))
        conn.commit()
        conn.close()

        flash("สมัครสมาชิกสำเร็จ กรุณาเข้าสู่ระบบ")
        return redirect("/login")

    return render_template("register.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form["username"].strip()
        password = request.form["password"].strip()

        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE username = ? AND password = ?", (username, password))
        user = cur.fetchone()
        conn.close()

        if user:
            session["user"] = user["username"]
            session["role"] = user["role"]
            flash("เข้าสู่ระบบสำเร็จ")
            return redirect("/")
        else:
            flash("ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง")

    return render_template("login.html")


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password():
    if request.method == "POST":
        email = request.form["email"].strip()

        conn = get_db()
        cur = conn.cursor()
        cur.execute("SELECT * FROM users WHERE email = ? ORDER BY id DESC LIMIT 1", (email,))
        user = cur.fetchone()
        conn.close()

        if user:
            send_reset_password_email(user)

        flash("หากอีเมลนี้มีอยู่ในระบบ เราได้ส่งลิงก์รีเซ็ตรหัสผ่านไปแล้ว")
        return redirect("/login")

    return render_template("forgot_password.html")


@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password(token):
    try:
        payload = serializer.loads(token, salt="reset-password", max_age=1800)
    except SignatureExpired:
        flash("ลิงก์รีเซ็ตรหัสผ่านหมดอายุแล้ว กรุณาขอใหม่อีกครั้ง")
        return redirect("/forgot-password")
    except BadSignature:
        flash("ลิงก์รีเซ็ตรหัสผ่านไม่ถูกต้อง")
        return redirect("/forgot-password")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM users WHERE id = ? AND email = ?", (payload["user_id"], payload["email"]))
    user = cur.fetchone()

    if not user:
        conn.close()
        flash("ไม่พบบัญชีผู้ใช้งาน")
        return redirect("/forgot-password")

    if request.method == "POST":
        password = request.form["password"].strip()
        confirm_password = request.form["confirm_password"].strip()

        if not password or not confirm_password:
            conn.close()
            flash("กรุณากรอกรหัสผ่านให้ครบ")
            return redirect(request.url)

        if password != confirm_password:
            conn.close()
            flash("รหัสผ่านและยืนยันรหัสผ่านไม่ตรงกัน")
            return redirect(request.url)

        cur.execute("UPDATE users SET password = ? WHERE id = ?", (password, user["id"]))
        conn.commit()
        conn.close()

        flash("เปลี่ยนรหัสผ่านสำเร็จ กรุณาเข้าสู่ระบบ")
        return redirect("/login")

    conn.close()
    return render_template("reset_password.html", token=token)


@app.route("/logout")
def logout():
    session.clear()
    flash("ออกจากระบบแล้ว")
    return redirect("/")


@app.route("/booking/<int:room_id>", methods=["GET", "POST"])
def booking(room_id):
    auto_cancel_expired_bookings()

    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อนจอง")
        return redirect("/login")

    room = get_room(room_id)
    if room is None:
        return "ไม่พบห้องพัก", 404
    if room.get("is_open", True) is False:
        flash(room.get("maintenance_note", "ห้องนี้ยังไม่พร้อมให้บริการ"))
        return redirect("/")

    all_time_options = get_all_time_options()
    today_time_options = get_today_available_time_options()

    if request.method == "POST":
        paytype = request.form["paytype"]
        raw_extra_segments = request.form.get("extra_segments", "[]")

        try:
            extra_segments = json.loads(raw_extra_segments) if raw_extra_segments else []
        except json.JSONDecodeError:
            extra_segments = []

        segment_payloads = [{
            "checkin": request.form["checkin"],
            "checkout": request.form["checkout"],
            "checkin_time": request.form["checkin_time"],
            "checkout_time": request.form["checkout_time"]
        }]

        for seg in extra_segments:
            if isinstance(seg, dict):
                segment_payloads.append({
                    "checkin": str(seg.get("checkin", "")).strip(),
                    "checkout": str(seg.get("checkout", "")).strip(),
                    "checkin_time": str(seg.get("checkin_time", "")).strip(),
                    "checkout_time": str(seg.get("checkout_time", "")).strip()
                })

        validated_segments = []
        now = get_now()

        for idx, seg in enumerate(segment_payloads, start=1):
            checkin = seg["checkin"]
            checkout = seg["checkout"]
            checkin_time = seg["checkin_time"]
            checkout_time = seg["checkout_time"]

            try:
                d1 = datetime.strptime(checkin, "%Y-%m-%d").date()
                d2 = datetime.strptime(checkout, "%Y-%m-%d").date()
                dt1 = datetime.strptime(f"{checkin} {checkin_time}", "%Y-%m-%d %H:%M")
                dt2 = datetime.strptime(f"{checkout} {checkout_time}", "%Y-%m-%d %H:%M")
            except ValueError:
                flash(f"ช่วงที่ {idx}: รูปแบบวันที่หรือเวลาไม่ถูกต้อง")
                return redirect(url_for("booking", room_id=room_id))

            if d1 < date.today():
                flash(f"ช่วงที่ {idx}: ไม่สามารถจองย้อนหลังได้")
                return redirect(url_for("booking", room_id=room_id))

            if d1 == date.today() and dt1 <= now:
                flash(f"ช่วงที่ {idx}: เวลาเข้าพักต้องมากกว่าเวลาปัจจุบัน")
                return redirect(url_for("booking", room_id=room_id))

            if dt2 <= dt1:
                flash(f"ช่วงที่ {idx}: วันและเวลาออกต้องมากกว่าวันและเวลาเข้า")
                return redirect(url_for("booking", room_id=room_id))

            validated_segments.append({
                "checkin": checkin,
                "checkout": checkout,
                "checkin_time": checkin_time,
                "checkout_time": checkout_time,
                "start_dt": dt1,
                "end_dt": dt2,
                "nights": calculate_charge_nights(dt1, dt2)
            })

        for i in range(len(validated_segments)):
            for j in range(i + 1, len(validated_segments)):
                if is_overlap(
                    validated_segments[i]["start_dt"],
                    validated_segments[i]["end_dt"],
                    validated_segments[j]["start_dt"],
                    validated_segments[j]["end_dt"]
                ):
                    flash("ช่วงวันเข้าพักที่เพิ่มมาซ้อนกันเอง กรุณาตรวจสอบอีกครั้ง")
                    return redirect(url_for("booking", room_id=room_id))

        conn = get_db()
        cur = conn.cursor()

        cur.execute("""
            SELECT * FROM bookings
            WHERE username = ?
              AND status NOT IN ('cancelled', 'rejected', 'refunded')
        """, (session["user"],))
        existing_user_bookings = cur.fetchall()

        for seg in validated_segments:
            for existing in existing_user_bookings:
                exist_in = datetime.strptime(f"{existing['checkin']} {existing['checkin_time']}", "%Y-%m-%d %H:%M")
                exist_out = datetime.strptime(f"{existing['checkout']} {existing['checkout_time']}", "%Y-%m-%d %H:%M")
                if is_overlap(seg["start_dt"], seg["end_dt"], exist_in, exist_out):
                    conn.close()
                    flash(
                        f"ช่วงวันที่ {seg['checkin']} {seg['checkin_time']} ถึง {seg['checkout']} {seg['checkout_time']} "
                        f"ซ้ำกับรายการเดิมของคุณ (ห้อง {existing['room_name']})"
                    )
                    return redirect(url_for("booking", room_id=room_id))

        batch_id = uuid.uuid4().hex
        created_count = 0

        for seg in validated_segments:
            total = seg["nights"] * room["price"]
            paid = calculate_due_now(total, paytype)
            remaining_due = calculate_remaining_due(total, paytype)

            cur.execute("""
                INSERT INTO bookings(
                    username, room_id, room_name, checkin, checkout,
                    checkin_time, checkout_time, nights, total, paytype,
                    paid, status, created_at, batch_id, remaining_due
                )
                VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
            """, (
                session["user"], room_id, room["name"], seg["checkin"], seg["checkout"],
                seg["checkin_time"], seg["checkout_time"], seg["nights"], total, paytype,
                paid, "waiting_approval", get_now().strftime("%Y-%m-%d %H:%M:%S"), batch_id,
                remaining_due
            ))
            created_count += 1

        conn.commit()
        conn.close()

        if created_count > 1:
            flash(f"ส่งคำขอจองเรียบร้อยแล้ว จำนวน {created_count} ช่วงวัน กรุณารอแอดมินอนุมัติ")
        else:
            flash("ส่งคำขอจองเรียบร้อยแล้ว กรุณารอแอดมินอนุมัติ")
        return redirect("/my-bookings")

    return render_template(
        "booking.html",
        room=room,
        all_time_options=all_time_options,
        today_time_options=today_time_options,
        today_str=date.today().strftime("%Y-%m-%d")
    )


@app.route("/my-bookings")
def my_bookings():
    auto_cancel_expired_bookings()

    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect("/login")

    conn = get_db()
    cur = conn.cursor()

    # เปิดหน้า "การจองของฉัน" แล้วให้ badge ของเมนูนี้หาย
    cur.execute("SELECT COALESCE(MAX(id), 0) AS max_id FROM notifications WHERE username = ?", (session["user"],))
    latest_notification = cur.fetchone()
    latest_notification_id = latest_notification["max_id"] if latest_notification else 0

    cur.execute("""
        UPDATE users
        SET last_seen_user_bookings = ?, last_seen_user_notification_id = ?
        WHERE username = ?
    """, (db_now_str(), latest_notification_id, session["user"]))

    cur.execute("SELECT * FROM bookings WHERE username = ? ORDER BY id DESC", (session["user"],))
    bookings = cur.fetchall()
    grouped_bookings = build_booking_groups(bookings)
    conn.commit()
    conn.close()

    return render_template("my_bookings.html", bookings=grouped_bookings)


@app.route("/cancel/<int:booking_id>", methods=["GET", "POST"])
def cancel_booking(booking_id):
    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect("/login")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/my-bookings")

    is_admin = session.get("role") == "admin"
    is_owner = booking["username"] == session.get("user")
    if not is_admin and not is_owner:
        conn.close()
        flash("คุณไม่มีสิทธิ์ยกเลิกรายการนี้")
        return redirect("/my-bookings")

    current_status = booking["status"]
    now_str = get_now().strftime("%Y-%m-%d %H:%M:%S")
    immediate_cancel_statuses = {"waiting_approval", "approved", "waiting_slip_approval"}
    paid_statuses = {"deposit_paid", "remaining_counter_pending", "fully_paid"}

    if current_status in immediate_cancel_statuses:
        cur.execute(
            """
            UPDATE bookings
            SET status = 'cancelled',
                cancel_requested_from = NULL,
                cancel_requested_at = NULL,
                refund_completed_at = NULL,
                refund_amount = 0,
                refund_method = NULL,
                refund_bank = NULL,
                refund_account = NULL,
                refund_note = NULL,
                cancelled_by = ?
            WHERE id = ?
            """,
            ("admin" if is_admin else "user", booking_id)
        )
        conn.commit()
        conn.close()
        flash("ยกเลิกการจองเรียบร้อย")
        return redirect("/admin" if is_admin else "/my-bookings")

    if current_status in paid_statuses:
        if is_admin:
            refund_amount = calculate_refund_amount(booking, "admin")
            cur.execute(
                """
                UPDATE bookings
                SET status = 'refund_pending',
                    cancel_requested_from = ?,
                    cancel_requested_at = ?,
                    refund_amount = ?,
                    refund_note = ?,
                    cancelled_by = 'admin',
                    refund_method = CASE WHEN paytype = 'deposit' THEN 'refund' ELSE refund_method END
                WHERE id = ?
                """,
                (current_status, now_str, refund_amount, "แอดมินเป็นผู้ยกเลิกรายการ คืนเงิน 100%", booking_id)
            )
            conn.commit()
            conn.close()
            flash(f"ยกเลิกรายการแล้ว และต้องคืนเงิน {refund_amount:.2f} บาท")
            return redirect("/admin")

        if request.method == "POST":
            refund_method = request.form.get("refund_method", "").strip()
            refund_account = request.form.get("refund_account", "").strip()
            refund_bank = request.form.get("refund_bank", "").strip()

            valid, final_bank, normalized_account, error_message = validate_refund_destination(
                refund_method, refund_account, refund_bank
            )
            if not valid:
                conn.close()
                flash(error_message)
                return redirect(request.path)

            refund_amount = calculate_refund_amount(booking, "user")
            note = "ลูกค้ายกเลิกหลังชำระเงิน: จ่ายเต็มคืน 50% / มัดจำไม่คืน"
            cur.execute(
                """
                UPDATE bookings
                SET status = 'cancel_requested',
                    cancel_requested_from = ?,
                    cancel_requested_at = ?,
                    refund_method = ?,
                    refund_bank = ?,
                    refund_account = ?,
                    refund_amount = ?,
                    refund_note = ?,
                    cancelled_by = 'user'
                WHERE id = ?
                """,
                (
                    current_status,
                    now_str,
                    refund_method.lower(),
                    final_bank,
                    normalized_account,
                    refund_amount,
                    note,
                    booking_id
                )
            )
            conn.commit()
            conn.close()
            flash(f"ส่งคำขอยกเลิกแล้ว ยอดคืนตามนโยบายเบื้องต้น {refund_amount:.2f} บาท")
            return redirect("/my-bookings")

        conn.close()
        refund_preview = calculate_refund_amount(booking, "user")
        return render_template("cancel_request.html", booking=booking, refund_preview=refund_preview, banks=BANKS)

    conn.close()
    flash("ไม่สามารถยกเลิกรายการนี้ได้")
    return redirect("/admin" if is_admin else "/my-bookings")


@app.route("/admin/approve-cancel/<int:booking_id>")
def admin_approve_cancel_request(booking_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/admin")

    if booking["status"] != "cancel_requested":
        conn.close()
        flash("รายการนี้ไม่ได้อยู่ในสถานะรออนุมัติการยกเลิก")
        return redirect("/admin")

    refund_amount = calculate_refund_amount(booking, booking["cancelled_by"] if booking["cancelled_by"] else "user")
    cur.execute(
        """
        UPDATE bookings
        SET status = 'refund_pending',
            refund_amount = ?,
            refund_note = COALESCE(refund_note, ?)
        WHERE id = ?
        """,
        (refund_amount, "รอแอดมินโอนคืนเงินให้ลูกค้า", booking_id)
    )
    conn.commit()
    conn.close()

    notify_user(
        booking["username"],
        "อนุมัติการยกเลิกแล้ว",
        f"คำขอยกเลิกการจองห้อง {booking['room_name']} ของคุณได้รับการอนุมัติแล้ว ระบบจะคืนเงิน {refund_amount:.2f} บาท ภายใน 2-4 วันทำการ"
    )
    flash("อนุมัติคำขอยกเลิกแล้ว ระบบจะรอคืนเงิน")
    return redirect("/admin")


@app.route("/admin/reject-cancel/<int:booking_id>")
def admin_reject_cancel_request(booking_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/admin")

    if booking["status"] != "cancel_requested":
        conn.close()
        flash("รายการนี้ไม่ได้อยู่ในสถานะรออนุมัติการยกเลิก")
        return redirect("/admin")

    rollback_status = booking["cancel_requested_from"] if booking["cancel_requested_from"] else "fully_paid"
    cur.execute(
        """
        UPDATE bookings
        SET status = ?,
            cancel_requested_from = NULL,
            cancel_requested_at = NULL
        WHERE id = ?
        """,
        (rollback_status, booking_id)
    )
    conn.commit()
    conn.close()

    notify_user(
        booking["username"],
        "ไม่อนุมัติการยกเลิก",
        f"คำขอยกเลิกการจองห้อง {booking['room_name']} ของคุณไม่ได้รับการอนุมัติ รายการจองยังคงเดิม"
    )
    flash("ปฏิเสธคำขอยกเลิกแล้ว")
    return redirect("/admin")


@app.route("/admin/refund-done/<int:booking_id>")
def admin_refund_done(booking_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/admin")

    if booking["status"] != "refund_pending":
        conn.close()
        flash("รายการนี้ไม่ได้อยู่ในสถานะรอคืนเงิน")
        return redirect("/admin")

    now_str = get_now().strftime("%Y-%m-%d %H:%M:%S")
    cur.execute(
        """
        UPDATE bookings
        SET status = 'refunded',
            refund_completed_at = ?,
            cancel_requested_from = NULL,
            cancel_requested_at = NULL
        WHERE id = ?
        """,
        (now_str, booking_id)
    )
    conn.commit()
    conn.close()

    notify_user(
        booking["username"],
        "คืนเงินเรียบร้อยแล้ว",
        f"รายการจองห้อง {booking['room_name']} ของคุณได้รับการคืนเงินเรียบร้อยแล้ว จำนวน {float(booking['refund_amount'] or 0):.2f} บาท"
    )
    flash("อัปเดตเป็นคืนเงินแล้ว")
    return redirect("/admin")



@app.route("/admin/export-refunds")
def admin_export_refunds():
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT id, username, room_name, paytype, paid, status, refund_method, refund_bank, refund_account,
               refund_amount, refund_note, cancelled_by, cancel_requested_at, refund_completed_at
        FROM bookings
        WHERE status IN ('cancel_requested', 'refund_pending', 'refunded')
        ORDER BY id DESC
    """)
    rows = cur.fetchall()
    conn.close()

    wb = Workbook()
    ws = wb.active
    ws.title = "Refunds"

    headers = [
        "Booking ID", "Username", "Room", "Payment Type", "Paid (THB)", "Refund Status",
        "Refund Amount (THB)", "Refund Method", "Bank", "Account", "Cancelled By",
        "Requested At", "Completed At", "Note"
    ]
    ws.append(headers)

    header_fill = PatternFill("solid", fgColor="5B3DB5")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="DDDDDD")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = Border(bottom=thin)

    for row in rows:
        bank_meta = get_bank_meta(row["refund_bank"])
        ws.append([
            row["id"],
            row["username"],
            row["room_name"],
            "จ่ายเต็ม" if row["paytype"] == "full" else "มัดจำ 25%",
            float(row["paid"] or 0),
            row["status"],
            float(row["refund_amount"] or 0),
            "PromptPay" if row["refund_method"] == "promptpay" else ("บัญชีธนาคาร" if row["refund_method"] == "bank" else row["refund_method"] or "-"),
            bank_meta["name"] if row["refund_bank"] else "-",
            row["refund_account"] or "-",
            row["cancelled_by"] or "-",
            row["cancel_requested_at"] or "-",
            row["refund_completed_at"] or "-",
            row["refund_note"] or "-"
        ])

    widths = {"A":12, "B":18, "C":18, "D":15, "E":14, "F":18, "G":18, "H":16, "I":18, "J":20, "K":14, "L":20, "M":20, "N":40}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    xlsx_data = BytesIO()
    wb.save(xlsx_data)
    xlsx_data.seek(0)
    return send_file(
        xlsx_data,
        as_attachment=True,
        download_name="refunds_gojames.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/approve/<int:booking_id>")
def approve_booking(booking_id):
    auto_cancel_expired_bookings()

    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/admin")

    if booking["status"] != "waiting_approval":
        conn.close()
        flash("รายการนี้ไม่ได้อยู่ในสถานะรออนุมัติ")
        return redirect("/admin")

    room_id = booking["room_id"]
    new_in = datetime.strptime(f"{booking['checkin']} {booking['checkin_time']}", "%Y-%m-%d %H:%M")
    new_out = datetime.strptime(f"{booking['checkout']} {booking['checkout_time']}", "%Y-%m-%d %H:%M")

    placeholders = ",".join("?" for _ in BOOKED_STATUSES)
    cur.execute(f"""
        SELECT * FROM bookings
        WHERE room_id = ? AND id != ? AND status IN ({placeholders})
    """, (room_id, booking_id, *BOOKED_STATUSES))
    existing_bookings = cur.fetchall()

    for row in existing_bookings:
        exist_in = datetime.strptime(f"{row['checkin']} {row['checkin_time']}", "%Y-%m-%d %H:%M")
        exist_out = datetime.strptime(f"{row['checkout']} {row['checkout_time']}", "%Y-%m-%d %H:%M")
        if is_overlap(new_in, new_out, exist_in, exist_out):
            conn.close()
            flash("อนุมัติไม่ได้ เพราะช่วงวันและเวลาซ้อนกับรายการที่ได้รับอนุมัติแล้ว")
            return redirect("/admin")

    cur.execute("SELECT * FROM room_blocks WHERE room_id = ?", (room_id,))
    blocks = cur.fetchall()

    for block in blocks:
        block_start = datetime.strptime(block["start_date"], "%Y-%m-%d").date()
        block_end = datetime.strptime(block["end_date"], "%Y-%m-%d").date()
        if block_overlaps_booking(new_in, new_out, block_start, block_end):
            conn.close()
            flash("อนุมัติไม่ได้ เพราะห้องนี้ถูกแอดมินตั้งเป็นไม่ว่างในช่วงดังกล่าว")
            return redirect("/admin")

    approved_at = get_now()
    deadline = approved_at + timedelta(minutes=15)

    cur.execute("""
        UPDATE bookings
        SET status = 'approved',
            approved_at = ?,
            payment_deadline = ?
        WHERE id = ?
    """, (
        approved_at.strftime("%Y-%m-%d %H:%M:%S"),
        deadline.strftime("%Y-%m-%d %H:%M:%S"),
        booking_id
    ))
    conn.commit()
    conn.close()

    notify_user(
        booking["username"],
        "จองสำเร็จ - รอชำระเงิน",
        f"รายการจองห้อง {booking['room_name']} ของคุณได้รับการอนุมัติแล้ว กรุณาชำระเงินภายใน 15 นาที",
        email_subject="จองสำเร็จ - กรุณาชำระเงินภายในเวลาที่กำหนด",
        email_body=(
            f"รายการจองห้อง {booking['room_name']} ของคุณได้รับการอนุมัติแล้ว\n"
            f"กรุณาชำระเงินภายใน 15 นาที\n"
            f"กำหนดชำระ: {deadline.strftime('%Y-%m-%d %H:%M:%S')}"
        )
    )
    flash("อนุมัติการจองแล้ว ลูกค้าต้องชำระภายใน 15 นาที")
    return redirect("/admin")


@app.route("/reject/<int:booking_id>")
def reject_booking(booking_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/admin")

    cur.execute("UPDATE bookings SET status = 'rejected' WHERE id = ?", (booking_id,))
    conn.commit()
    conn.close()

    notify_user(booking["username"], "การจองถูกปฏิเสธ", f"รายการจองห้อง {booking['room_name']} ของคุณถูกปฏิเสธ")
    flash("ปฏิเสธการจองแล้ว")
    return redirect("/admin")


@app.route("/payment/<int:booking_id>")
def payment(booking_id):
    auto_cancel_expired_bookings()

    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect("/login")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/my-bookings")

    if session.get("role") != "admin" and booking["username"] != session["user"]:
        conn.close()
        flash("คุณไม่มีสิทธิ์ดูรายการนี้")
        return redirect("/my-bookings")

    if booking["status"] != "approved":
        conn.close()
        flash("รายการนี้ยังไม่อยู่ในสถานะชำระเงิน")
        return redirect("/my-bookings")

    if booking["batch_id"]:
        cur.execute("""
            SELECT * FROM bookings
            WHERE batch_id = ? AND username = ?
            ORDER BY checkin ASC, checkin_time ASC
        """, (booking["batch_id"], booking["username"]))
        batch_bookings = cur.fetchall()
    else:
        batch_bookings = [booking]
    conn.close()

    amount = round(sum(row["paid"] for row in batch_bookings), 2)
    total = round(sum(row["total"] for row in batch_bookings), 2)
    remaining_due = round(sum(row["remaining_due"] for row in batch_bookings), 2)
    qr_filename = create_qr(booking["id"], amount)

    return render_template(
        "payment.html",
        booking=booking,
        booking_id=booking["id"],
        batch_bookings=batch_bookings,
        amount=amount,
        remaining_due=remaining_due,
        qr_filename=qr_filename,
        paytype=booking["paytype"],
        total=total,
        payment_mode="initial"
    )



@app.route("/remaining-payment/<int:booking_id>")
def remaining_payment(booking_id):
    auto_cancel_expired_bookings()

    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect("/login")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/my-bookings")

    if booking["username"] != session["user"]:
        conn.close()
        flash("คุณไม่มีสิทธิ์ดูรายการนี้")
        return redirect("/my-bookings")

    allowed_remaining_statuses = {"deposit_paid", "remaining_counter_pending"}
    if booking["status"] not in allowed_remaining_statuses:
        conn.close()
        flash("รายการนี้ยังไม่อยู่ในสถานะชำระส่วนที่เหลือ")
        return redirect("/my-bookings")

    if booking["batch_id"]:
        cur.execute("""
            SELECT * FROM bookings
            WHERE batch_id = ? AND username = ?
            ORDER BY checkin ASC, checkin_time ASC
        """, (booking["batch_id"], booking["username"]))
        batch_bookings = [row for row in cur.fetchall() if row["status"] in allowed_remaining_statuses]
    else:
        batch_bookings = [booking]
    conn.close()

    amount = round(sum(row["remaining_due"] for row in batch_bookings), 2)
    total = round(sum(row["total"] for row in batch_bookings), 2)
    paid_so_far = round(sum(row["paid"] for row in batch_bookings), 2)
    qr_filename = create_qr(booking["id"], amount)

    return render_template(
        "payment.html",
        booking=booking,
        booking_id=booking["id"],
        batch_bookings=batch_bookings,
        amount=amount,
        paid_so_far=paid_so_far,
        remaining_due=0,
        qr_filename=qr_filename,
        paytype=booking["paytype"],
        total=total,
        payment_mode="remaining"
    )


@app.route("/upload_slip/<int:booking_id>", methods=["POST"])
def upload_slip(booking_id):
    auto_cancel_expired_bookings()

    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect("/login")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/my-bookings")

    if booking["username"] != session["user"]:
        conn.close()
        flash("คุณไม่มีสิทธิ์อัปโหลดสลิปของรายการนี้")
        return redirect("/my-bookings")

    valid_statuses = {"approved", "deposit_paid", "remaining_counter_pending"}
    if booking["status"] not in valid_statuses:
        conn.close()
        flash("รายการนี้ยังไม่สามารถอัปโหลดสลิปได้")
        return redirect("/my-bookings")

    is_remaining_flow = booking["status"] in {"deposit_paid", "remaining_counter_pending"}
    remain_method = request.form.get("remain_method", "qr") if is_remaining_flow else "qr"

    if is_remaining_flow and remain_method == "counter":
        if booking["batch_id"]:
            cur.execute("""
                UPDATE bookings
                SET status = 'remaining_counter_pending', remaining_payment_method = 'counter'
                WHERE batch_id = ? AND username = ? AND status IN ('deposit_paid', 'remaining_counter_pending')
            """, (booking["batch_id"], booking["username"]))
        else:
            cur.execute("""
                UPDATE bookings
                SET status = 'remaining_counter_pending', remaining_payment_method = 'counter'
                WHERE id = ?
            """, (booking_id,))
        conn.commit()
        conn.close()
        flash("บันทึกแล้ว: ยอดส่วนที่เหลือจะชำระหน้าเคาน์เตอร์ที่พัก")
        return redirect("/my-bookings")

    slip = request.files.get("slip")
    if not slip or slip.filename == "":
        conn.close()
        flash("กรุณาเลือกไฟล์สลิป")
        return redirect(url_for("payment", booking_id=booking_id) if booking["status"] == "approved" else url_for("remaining_payment", booking_id=booking_id))

    if not allowed_file(slip.filename):
        conn.close()
        flash("รองรับเฉพาะไฟล์ png, jpg, jpeg, webp")
        return redirect(url_for("payment", booking_id=booking_id) if booking["status"] == "approved" else url_for("remaining_payment", booking_id=booking_id))

    ext = slip.filename.rsplit(".", 1)[1].lower()
    filename = secure_filename(f"slip_{booking_id}.{ext}")
    filepath = os.path.join(UPLOAD_DIR, filename)
    slip.save(filepath)

    target_status = "waiting_slip_approval" if booking["status"] == "approved" else "waiting_remaining_slip_approval"

    if booking["batch_id"]:
        if is_remaining_flow:
            cur.execute(f"""
                UPDATE bookings
                SET slip_filename = ?, status = ?, remaining_payment_method = 'qr'
                WHERE batch_id = ? AND username = ? AND status IN ('deposit_paid', 'remaining_counter_pending')
            """, (filename, target_status, booking["batch_id"], booking["username"]))
        else:
            cur.execute(f"""
                UPDATE bookings
                SET slip_filename = ?, status = ?
                WHERE batch_id = ? AND username = ? AND status = ?
            """, (filename, target_status, booking["batch_id"], booking["username"], booking["status"]))
    else:
        if is_remaining_flow:
            cur.execute("""
                UPDATE bookings
                SET slip_filename = ?, status = ?, remaining_payment_method = 'qr'
                WHERE id = ?
            """, (filename, target_status, booking_id))
        else:
            cur.execute("""
                UPDATE bookings
                SET slip_filename = ?, status = ?
                WHERE id = ?
            """, (filename, target_status, booking_id))
    conn.commit()
    conn.close()

    flash("อัปโหลดสลิปเรียบร้อยแล้ว กรุณารอแอดมินตรวจสอบ")
    return redirect("/my-bookings")

@app.route("/approve_slip/<int:booking_id>")
def approve_slip(booking_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/admin")

    if booking["status"] not in ("waiting_slip_approval", "waiting_remaining_slip_approval"):
        conn.close()
        flash("รายการนี้ไม่ได้อยู่ในสถานะรอตรวจสลิป")
        return redirect("/admin")

    if booking["status"] == "waiting_slip_approval":
        new_status = "fully_paid" if booking["paytype"] == "full" else "deposit_paid"
        update_paid_sql = ""
        params = ()
        notify_msg = "รายการจองห้อง {room} ของคุณได้รับการยืนยันการชำระเงินเรียบร้อยแล้ว"
    else:
        new_status = "fully_paid"
        notify_msg = "รายการจองห้อง {room} ของคุณได้รับการยืนยันการชำระส่วนที่เหลือเรียบร้อยแล้ว"

    if booking["batch_id"]:
        if booking["status"] == "waiting_remaining_slip_approval":
            cur.execute("""
                UPDATE bookings
                SET status = ?, paid = total, remaining_due = 0
                WHERE batch_id = ? AND username = ? AND status = 'waiting_remaining_slip_approval'
            """, (new_status, booking["batch_id"], booking["username"]))
        else:
            cur.execute("""
                UPDATE bookings
                SET status = ?
                WHERE batch_id = ? AND username = ? AND status = 'waiting_slip_approval'
            """, (new_status, booking["batch_id"], booking["username"]))
    else:
        if booking["status"] == "waiting_remaining_slip_approval":
            cur.execute("UPDATE bookings SET status = ?, paid = total, remaining_due = 0 WHERE id = ?", (new_status, booking_id))
        else:
            cur.execute("UPDATE bookings SET status = ? WHERE id = ?", (new_status, booking_id))

    conn.commit()
    conn.close()

    notify_user(booking["username"], "ชำระเงินสำเร็จ", notify_msg.format(room=booking['room_name']))
    flash("อนุมัติสลิปเรียบร้อยแล้ว")
    return redirect("/admin")


@app.route("/reject_slip/<int:booking_id>")
def reject_slip(booking_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/admin")

    if booking["status"] not in ("waiting_slip_approval", "waiting_remaining_slip_approval"):
        conn.close()
        flash("รายการนี้ไม่ได้อยู่ในสถานะรอตรวจสลิป")
        return redirect("/admin")

    rollback_status = "approved" if booking["status"] == "waiting_slip_approval" else "deposit_paid"

    if booking["batch_id"]:
        cur.execute("""
            UPDATE bookings
            SET status = ?, slip_filename = NULL
            WHERE batch_id = ? AND username = ? AND status = ?
        """, (rollback_status, booking["batch_id"], booking["username"], booking["status"]))
    else:
        cur.execute("""
            UPDATE bookings
            SET status = ?, slip_filename = NULL
            WHERE id = ?
        """, (rollback_status, booking_id))
    conn.commit()
    conn.close()

    msg = "สลิปการชำระเงิน" if rollback_status == "approved" else "สลิปชำระส่วนที่เหลือ"
    notify_user(booking["username"], "สลิปถูกปฏิเสธ", f"{msg} สำหรับห้อง {booking['room_name']} ของคุณถูกปฏิเสธ กรุณาอัปโหลดใหม่")
    flash("ปฏิเสธสลิปแล้ว ให้ลูกค้าอัปโหลดใหม่")
    return redirect("/admin")


@app.route("/block-room", methods=["POST"])
def block_room():
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    room_id = request.form["room_id"]
    start_date = request.form["start_date"]
    end_date = request.form["end_date"]
    note = request.form.get("note", "").strip()

    try:
        room_id = int(room_id)
        start_day = datetime.strptime(start_date, "%Y-%m-%d").date()
        end_day = datetime.strptime(end_date, "%Y-%m-%d").date()
    except ValueError:
        flash("ข้อมูลวันที่ไม่ถูกต้อง")
        return redirect("/admin")

    if end_day < start_day:
        flash("วันสิ้นสุดต้องไม่ก่อนวันเริ่มต้น")
        return redirect("/admin")

    conn = get_db()
    cur = conn.cursor()

    placeholders = ",".join("?" for _ in BOOKED_STATUSES)
    cur.execute(f"""
        SELECT * FROM bookings
        WHERE room_id = ? AND status IN ({placeholders})
    """, (room_id, *BOOKED_STATUSES))
    bookings = cur.fetchall()

    for booking in bookings:
        checkin = datetime.strptime(f"{booking['checkin']} {booking['checkin_time']}", "%Y-%m-%d %H:%M")
        checkout = datetime.strptime(f"{booking['checkout']} {booking['checkout_time']}", "%Y-%m-%d %H:%M")
        if block_overlaps_booking(checkin, checkout, start_day, end_day):
            conn.close()
            flash("บล็อกไม่ได้ เพราะมีรายการจองหรือชำระแล้วในช่วงดังกล่าว")
            return redirect("/admin")

    cur.execute("""
        INSERT INTO room_blocks(room_id, start_date, end_date, note, created_at)
        VALUES(?,?,?,?,?)
    """, (room_id, start_date, end_date, note, get_now().strftime("%Y-%m-%d %H:%M:%S")))
    conn.commit()
    conn.close()

    flash("บล็อกห้องเรียบร้อยแล้ว")
    return redirect("/admin")


@app.route("/unblock-room/<int:block_id>")
def unblock_room(block_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("DELETE FROM room_blocks WHERE id = ?", (block_id,))
    conn.commit()
    conn.close()

    flash("ยกเลิกการบล็อกห้องแล้ว")
    return redirect("/admin")


@app.route("/notifications")
def notifications():
    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect("/login")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        SELECT * FROM notifications
        WHERE username = ?
        ORDER BY id DESC
    """, (session["user"],))
    items = cur.fetchall()
    conn.close()

    return render_template("notifications.html", notifications=items)


@app.route("/notifications/read/<int:notification_id>")
def read_notification(notification_id):
    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect("/login")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE notifications
        SET is_read = 1
        WHERE id = ? AND username = ?
    """, (notification_id, session["user"]))
    conn.commit()
    conn.close()

    return redirect("/notifications")


@app.route("/notifications/read-all")
def read_all_notifications():
    if "user" not in session:
        flash("กรุณาเข้าสู่ระบบก่อน")
        return redirect("/login")

    conn = get_db()
    cur = conn.cursor()
    cur.execute("""
        UPDATE notifications
        SET is_read = 1
        WHERE username = ?
    """, (session["user"],))
    conn.commit()
    conn.close()

    return redirect("/notifications")


@app.route("/admin/mark-counter-paid/<int:booking_id>")
def admin_mark_counter_paid(booking_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()
    cur.execute("SELECT * FROM bookings WHERE id = ?", (booking_id,))
    booking = cur.fetchone()

    if booking is None:
        conn.close()
        flash("ไม่พบรายการจอง")
        return redirect("/admin")

    if booking["status"] != "remaining_counter_pending":
        conn.close()
        flash("รายการนี้ยังไม่ได้เลือกชำระยอดคงเหลือหน้าเคาน์เตอร์")
        return redirect("/admin")

    if booking["batch_id"]:
        cur.execute("""
            UPDATE bookings
            SET status = 'fully_paid', paid = total, remaining_due = 0
            WHERE batch_id = ? AND username = ? AND status = 'remaining_counter_pending'
        """, (booking["batch_id"], booking["username"]))
    else:
        cur.execute("""
            UPDATE bookings
            SET status = 'fully_paid', paid = total, remaining_due = 0
            WHERE id = ?
        """, (booking_id,))

    conn.commit()
    conn.close()

    notify_user(booking["username"], "ชำระเงินครบแล้ว", f"รายการจองห้อง {booking['room_name']} ของคุณได้รับการยืนยันว่าชำระยอดคงเหลือหน้าเคาน์เตอร์เรียบร้อยแล้ว")
    flash("บันทึกการชำระหน้าเคาน์เตอร์เรียบร้อยแล้ว")
    return redirect("/admin")


@app.route("/admin")
def admin():
    auto_cancel_expired_bookings()

    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    conn = get_db()
    cur = conn.cursor()

    # เปิดหน้า "จัดการการจอง" แล้วให้ badge ของเมนูนี้หาย
    cur.execute("SELECT COALESCE(MAX(id), 0) AS max_id FROM notifications WHERE username = 'admin'")
    latest_notification = cur.fetchone()
    latest_notification_id = latest_notification["max_id"] if latest_notification else 0

    cur.execute("""
        UPDATE users
        SET last_seen_admin_bookings = ?, last_seen_admin_notification_id = ?
        WHERE username = ?
    """, (db_now_str(), latest_notification_id, session["user"]))

    cur.execute("""
        SELECT bookings.*, users.first_name, users.last_name
        FROM bookings
        LEFT JOIN users ON bookings.username = users.username
        ORDER BY bookings.id DESC
    """)
    bookings = cur.fetchall()

    cur.execute("SELECT * FROM room_blocks ORDER BY start_date ASC")
    blocks = cur.fetchall()
    conn.commit()
    conn.close()

    return render_template("admin.html", bookings=bookings, blocks=blocks, rooms=get_rooms_with_status())


@app.route("/admin/toggle-room/<int:room_id>")
def admin_toggle_room(room_id):
    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    room = None
    for r in get_rooms_with_status():
        if r["id"] == room_id:
            room = r
            break

    if room is None:
        flash("ไม่พบห้องพัก")
        return redirect("/admin")

    new_is_open = 0 if room.get("is_open", True) else 1

    conn = get_db()
    cur = conn.cursor()
    cur.execute(
        "INSERT INTO room_settings(room_id, is_open) VALUES (?, ?) ON CONFLICT(room_id) DO UPDATE SET is_open = excluded.is_open",
        (room_id, new_is_open)
    )
    conn.commit()
    conn.close()

    if new_is_open:
        flash(f"เปิดห้อง {room['name']} ให้จองได้แล้ว")
    else:
        flash(f"ปิดห้อง {room['name']} ชั่วคราวแล้ว")

    return redirect("/admin")


@app.route("/admin/calendar")
def admin_calendar():
    auto_cancel_expired_bookings()

    if session.get("role") != "admin":
        return "สำหรับผู้ดูแลระบบเท่านั้น"

    day_list, room_rows = get_calendar_data(days=14)
    return render_template("admin_calendar.html", day_list=day_list, room_rows=room_rows)


init_db()

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)