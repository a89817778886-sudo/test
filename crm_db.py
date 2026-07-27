# -*- coding: utf-8 -*-
"""
crm_db.py — полнофункциональный CRM-слой для приложения КП по кранам/траверсам ЛКС.
Версия 2.0 — реализует 20-пунктовый план доработки (напоминания, история изменений,
причины отказа, заметки, дашборд-агрегаты, отчёт по менеджерам, прогноз выручки,
поддержка ИНН/риски, роли пользователей). Обратная совместимость с версией 1.0 сохранена.
"""

from __future__ import annotations

import sqlite3
import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional, List, Dict, Any

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "crm.db"
KP_FILES_DIR = APP_DIR / "crm_files"
KP_FILES_DIR.mkdir(exist_ok=True)

DISCOUNT_MIN_PCT = 0
DISCOUNT_MAX_PCT = 30
DISCOUNT_RECOMMENDED = (5, 10)

STATUS_DRAFT = "draft"
STATUS_SENT = "sent"
STATUS_WON = "won"
STATUS_LOST = "lost"

STATUS_LABELS = {
    STATUS_DRAFT: "Черновик",
    STATUS_SENT: "Отправлен клиенту",
    STATUS_WON: "Продано",
    STATUS_LOST: "Отказ",
}
STATUS_ORDER = [STATUS_DRAFT, STATUS_SENT, STATUS_WON, STATUS_LOST]

PRODUCT_TYPES = ["Кран", "Траверса", "Кран + траверса", "Тали", "Прочее"]

LOSS_REASONS = [
    "Дорого / нашли дешевле",
    "Выбрали конкурента",
    "Отложили закупку",
    "Проект не состоялся",
    "Долгий срок поставки",
    "Не устроили условия монтажа",
    "Не отвечает / потеряна связь",
    "Другое",
]

ROLE_ADMIN = "admin"
ROLE_MANAGER = "manager"
ROLE_LABELS = {ROLE_ADMIN: "Администратор", ROLE_MANAGER: "Менеджер"}


# --- Адаптер PostgreSQL (Supabase) ---
# Если в secrets задан SUPABASE_DB_URL — всё пишется в облако, иначе в SQLite файл.
try:
    import db_adapter as _db_adapter
    _USE_PG = _db_adapter.is_pg_enabled()
except Exception:
    _USE_PG = False
    _db_adapter = None


def get_conn():
    if _USE_PG:
        conn = _db_adapter.get_pg_connection()
        conn.row_factory = True  # признак что нужно возвращать dict-подобные ряды
        return conn
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def _column_exists(conn, table: str, col: str) -> bool:
    if _USE_PG:
        # В PostgreSQL — через information_schema
        cur = conn.execute(
            "SELECT column_name FROM information_schema.columns "
            "WHERE table_schema='public' AND table_name=? AND column_name=?",
            (table, col))
        return cur.fetchone() is not None
    rows = conn.execute(f"PRAGMA table_info({table})").fetchall()
    return any(r["name"] == col for r in rows)


def _add_column_if_missing(conn, table: str, col: str, coltype: str, default_sql: str = ""):
    if not _column_exists(conn, table, col):
        default_clause = f" DEFAULT {default_sql}" if default_sql else ""
        conn.execute(f"ALTER TABLE {table} ADD COLUMN {col} {coltype}{default_clause}")


def init_db() -> None:
    # В режиме PostgreSQL — схема создаётся отдельно через supabase_schema.sql
    # Здесь — no-op, чтобы не падать на SQLite-специфичном CREATE TABLE
    if _USE_PG:
        return
    conn = get_conn()
    cur = conn.cursor()

    cur.execute("""
    CREATE TABLE IF NOT EXISTS customers (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        inn TEXT UNIQUE,
        kpp TEXT, ogrn TEXT, name_full TEXT, name_short TEXT,
        phone TEXT, email TEXT, address TEXT,
        director_position TEXT, director_fio TEXT,
        bank TEXT, bik TEXT, rs TEXT, ks TEXT,
        created_at TEXT NOT NULL, updated_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS quotes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        kp_number TEXT NOT NULL,
        customer_id INTEGER,
        product_type TEXT NOT NULL,
        product_model TEXT,
        include_montage INTEGER NOT NULL DEFAULT 0,
        delivery_city TEXT,
        base_total REAL NOT NULL,
        discount_pct REAL NOT NULL DEFAULT 0,
        discount_amount REAL NOT NULL DEFAULT 0,
        final_total REAL NOT NULL,
        status TEXT NOT NULL DEFAULT 'draft',
        currency TEXT NOT NULL DEFAULT 'RUB',
        created_at TEXT NOT NULL,
        updated_at TEXT NOT NULL,
        sold_at TEXT,
        notes TEXT,
        contact_fio TEXT,
        last_contact_at TEXT,
        request_summary TEXT,
        FOREIGN KEY (customer_id) REFERENCES customers(id)
    )
    """)

    # Миграция: добавляем колонки в существующую таблицу если их нет
    _existing_cols = {row[1] for row in cur.execute("PRAGMA table_info(quotes)").fetchall()}
    for _col_name, _col_type in [
        ("contact_fio", "TEXT"),
        ("last_contact_at", "TEXT"),
        ("request_summary", "TEXT"),
    ]:
        if _col_name not in _existing_cols:
            try:
                cur.execute(f"ALTER TABLE quotes ADD COLUMN {_col_name} {_col_type}")
            except Exception:
                pass

    cur.execute("""
    CREATE TABLE IF NOT EXISTS quote_items (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        quote_id INTEGER NOT NULL,
        code TEXT, name TEXT NOT NULL, unit TEXT DEFAULT 'шт',
        qty REAL NOT NULL DEFAULT 1, price REAL NOT NULL DEFAULT 0, total REAL NOT NULL DEFAULT 0,
        FOREIGN KEY (quote_id) REFERENCES quotes(id) ON DELETE CASCADE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS sales (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        quote_id INTEGER NOT NULL UNIQUE,
        customer_id INTEGER,
        product_type TEXT, product_model TEXT,
        sale_date TEXT NOT NULL, price REAL NOT NULL,
        discount_pct REAL NOT NULL DEFAULT 0, delivery_city TEXT,
        FOREIGN KEY (quote_id) REFERENCES quotes(id),
        FOREIGN KEY (customer_id) REFERENCES customers(id)
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS users (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        username TEXT UNIQUE NOT NULL,
        password_hash TEXT NOT NULL,
        full_name TEXT, role TEXT NOT NULL DEFAULT 'manager',
        telegram_chat_id TEXT, is_active INTEGER NOT NULL DEFAULT 1,
        created_at TEXT NOT NULL
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS quote_history (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        quote_id INTEGER NOT NULL, user_id INTEGER, username TEXT,
        field_changed TEXT NOT NULL, old_value TEXT, new_value TEXT, changed_at TEXT NOT NULL,
        FOREIGN KEY (quote_id) REFERENCES quotes(id) ON DELETE CASCADE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS customer_notes (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        customer_id INTEGER NOT NULL, user_id INTEGER, username TEXT,
        note TEXT NOT NULL, created_at TEXT NOT NULL,
        FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE CASCADE
    )
    """)

    cur.execute("""
    CREATE TABLE IF NOT EXISTS reminders (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        quote_id INTEGER, customer_id INTEGER, user_id INTEGER,
        due_at TEXT NOT NULL, message TEXT NOT NULL,
        is_done INTEGER NOT NULL DEFAULT 0, created_at TEXT NOT NULL,
        FOREIGN KEY (quote_id) REFERENCES quotes(id) ON DELETE CASCADE,
        FOREIGN KEY (customer_id) REFERENCES customers(id) ON DELETE CASCADE
    )
    """)

    conn.commit()

    _add_column_if_missing(conn, "quotes", "loss_reason", "TEXT")
    _add_column_if_missing(conn, "quotes", "owner_id", "INTEGER")
    _add_column_if_missing(conn, "quotes", "owner_name", "TEXT")
    _add_column_if_missing(conn, "quotes", "probability_pct", "REAL", "50")
    _add_column_if_missing(conn, "customers", "risk_flag", "TEXT")
    _add_column_if_missing(conn, "customers", "risk_checked_at", "TEXT")
    # Миграция: файлы PDF/DOCX КП в базе (пути или байты)
    _add_column_if_missing(conn, "quotes", "pdf_path", "TEXT")
    _add_column_if_missing(conn, "quotes", "docx_path", "TEXT")

    # Миграция: пустые ИНН → NULL (избавляемся от дубля UNIQUE=''),
    # чтобы PostgreSQL не баговал customers_inn_key
    try:
        conn.execute("UPDATE customers SET inn=NULL WHERE inn=''")
    except Exception:
        # старые БД — могут не быть колонки inn, не критично
        pass

    conn.commit()
    conn.close()


@dataclass
class Customer:
    id: Optional[int] = None
    inn: str = ""
    kpp: str = ""
    ogrn: str = ""
    name_full: str = ""
    name_short: str = ""
    phone: str = ""
    email: str = ""
    address: str = ""
    director_position: str = ""
    director_fio: str = ""
    bank: str = ""
    bik: str = ""
    rs: str = ""
    ks: str = ""
    risk_flag: str = ""
    risk_checked_at: str = ""


def upsert_customer(c: Customer) -> int:
    conn = get_conn()
    cur = conn.cursor()
    now = dt.datetime.now().isoformat(timespec="seconds")

    # Нормализация: пустой/whitespace ИНН → None (NULL в БД),
    # чтобы UNIQUE-constraint не ловил дубли пустых строк (PostgreSQL).
    _inn_norm = (c.inn or "").strip() or None

    existing = None
    if _inn_norm:
        # Ищем по ИНН
        row = cur.execute("SELECT id FROM customers WHERE inn = ?", (_inn_norm,)).fetchone()
        if row:
            existing = row["id"]
    else:
        # Без ИНН — ищем по короткому названию (чтобы не плодить дубли «без ИНН»)
        _name = (c.name_short or c.name_full or "").strip()
        if _name:
            row = cur.execute(
                "SELECT id FROM customers WHERE (inn IS NULL OR inn = '') "
                "AND (name_short = ? OR name_full = ?) LIMIT 1",
                (_name, _name)
            ).fetchone()
            if row:
                existing = row["id"]

    if existing:
        cur.execute("""
            UPDATE customers SET inn=?, kpp=?, ogrn=?, name_full=?, name_short=?, phone=?, email=?,
                address=?, director_position=?, director_fio=?, bank=?, bik=?, rs=?, ks=?, updated_at=?
            WHERE id=?
        """, (_inn_norm, c.kpp, c.ogrn, c.name_full, c.name_short, c.phone, c.email, c.address,
              c.director_position, c.director_fio, c.bank, c.bik, c.rs, c.ks, now, existing))
        conn.commit()
        conn.close()
        return existing

    cur.execute("""
        INSERT INTO customers (inn, kpp, ogrn, name_full, name_short, phone, email, address,
            director_position, director_fio, bank, bik, rs, ks, created_at, updated_at)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (_inn_norm, c.kpp, c.ogrn, c.name_full, c.name_short, c.phone, c.email, c.address,
          c.director_position, c.director_fio, c.bank, c.bik, c.rs, c.ks, now, now))
    conn.commit()
    new_id = cur.lastrowid
    conn.close()
    return new_id


def find_customer_by_inn(inn: str) -> Optional[Customer]:
    conn = get_conn()
    row = conn.execute("SELECT * FROM customers WHERE inn = ?", (inn,)).fetchone()
    conn.close()
    if not row:
        return None
    return Customer(**{k: row[k] for k in row.keys() if k in Customer.__dataclass_fields__})


def get_customer(customer_id: int) -> Optional[Customer]:
    conn = get_conn()
    row = conn.execute("SELECT * FROM customers WHERE id = ?", (customer_id,)).fetchone()
    conn.close()
    if not row:
        return None
    return Customer(**{k: row[k] for k in row.keys() if k in Customer.__dataclass_fields__})


def search_customers(query: str) -> List[Customer]:
    conn = get_conn()
    like = f"%{query}%"
    rows = conn.execute("""
        SELECT * FROM customers WHERE inn LIKE ? OR name_full LIKE ? OR name_short LIKE ?
           OR phone LIKE ? OR email LIKE ? ORDER BY updated_at DESC LIMIT 50
    """, (like, like, like, like, like)).fetchall()
    conn.close()
    return [Customer(**{k: r[k] for k in r.keys() if k in Customer.__dataclass_fields__}) for r in rows]


def update_customer_risk(customer_id: int, risk_flag: str) -> None:
    conn = get_conn()
    now = dt.datetime.now().isoformat(timespec="seconds")
    conn.execute("UPDATE customers SET risk_flag=?, risk_checked_at=? WHERE id=?", (risk_flag, now, customer_id))
    conn.commit()
    conn.close()


def autofill_customer_by_inn(inn: str, requisites_text: str = "") -> Customer:
    existing = find_customer_by_inn(inn)
    if existing:
        return existing
    c = Customer(inn=inn)
    if requisites_text:
        try:
            import req_parser
            parsed = req_parser.parse_requisites(requisites_text)
            c.kpp = parsed.get("kpp", "")
            c.ogrn = parsed.get("ogrn", "")
            c.name_full = parsed.get("name_full", "") or parsed.get("full", "")
            c.name_short = parsed.get("name_short", "") or parsed.get("short", "")
            c.phone = parsed.get("phone", "")
            c.email = parsed.get("email", "")
            c.address = parsed.get("address", "")
            c.director_position = parsed.get("director_position", "")
            c.director_fio = parsed.get("director_fio", "")
            c.bank = parsed.get("bank", "")
            c.bik = parsed.get("bik", "")
            c.rs = parsed.get("rs", "")
            c.ks = parsed.get("ks", "")
        except Exception:
            pass
    return c


def get_customer_full(customer_id: int) -> Dict[str, Any]:
    customer = get_customer(customer_id)
    if not customer:
        return {}
    conn = get_conn()
    quotes = conn.execute("SELECT * FROM quotes WHERE customer_id=? ORDER BY created_at DESC", (customer_id,)).fetchall()
    sales = conn.execute("SELECT * FROM sales WHERE customer_id=? ORDER BY sale_date DESC", (customer_id,)).fetchall()
    notes = conn.execute("SELECT * FROM customer_notes WHERE customer_id=? ORDER BY created_at DESC", (customer_id,)).fetchall()
    reminders = conn.execute("SELECT * FROM reminders WHERE customer_id=? ORDER BY due_at ASC", (customer_id,)).fetchall()
    conn.close()
    return {
        "customer": customer,
        "quotes": [dict(q) for q in quotes],
        "sales": [dict(s) for s in sales],
        "notes": [dict(n) for n in notes],
        "reminders": [dict(r) for r in reminders],
    }


def add_customer_note(customer_id: int, note: str, user_id: Optional[int] = None, username: str = "") -> int:
    conn = get_conn()
    now = dt.datetime.now().isoformat(timespec="seconds")
    cur = conn.execute("""
        INSERT INTO customer_notes (customer_id, user_id, username, note, created_at) VALUES (?,?,?,?,?)
    """, (customer_id, user_id, username, note, now))
    conn.commit()
    note_id = cur.lastrowid
    conn.close()
    return note_id


def list_customer_notes(customer_id: int) -> List[Dict[str, Any]]:
    conn = get_conn()
    rows = conn.execute("SELECT * FROM customer_notes WHERE customer_id=? ORDER BY created_at DESC", (customer_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def add_reminder(due_at: str, message: str, quote_id: Optional[int] = None,
                  customer_id: Optional[int] = None, user_id: Optional[int] = None) -> int:
    conn = get_conn()
    now = dt.datetime.now().isoformat(timespec="seconds")
    cur = conn.execute("""
        INSERT INTO reminders (quote_id, customer_id, user_id, due_at, message, created_at) VALUES (?,?,?,?,?,?)
    """, (quote_id, customer_id, user_id, due_at, message, now))
    conn.commit()
    rid = cur.lastrowid
    conn.close()
    return rid


def list_due_reminders(as_of: Optional[str] = None, only_pending: bool = True) -> List[Dict[str, Any]]:
    as_of = as_of or dt.datetime.now().isoformat(timespec="seconds")
    conn = get_conn()
    q = """
        SELECT r.*, c.name_short, c.phone, q.kp_number FROM reminders r
        LEFT JOIN customers c ON c.id = r.customer_id LEFT JOIN quotes q ON q.id = r.quote_id
        WHERE r.due_at <= ?
    """
    if only_pending:
        q += " AND r.is_done = 0"
    q += " ORDER BY r.due_at ASC"
    rows = conn.execute(q, (as_of,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def list_all_reminders(only_pending: bool = True) -> List[Dict[str, Any]]:
    conn = get_conn()
    q = """
        SELECT r.*, c.name_short, c.phone, q.kp_number FROM reminders r
        LEFT JOIN customers c ON c.id = r.customer_id LEFT JOIN quotes q ON q.id = r.quote_id WHERE 1=1
    """
    if only_pending:
        q += " AND r.is_done = 0"
    q += " ORDER BY r.due_at ASC"
    rows = conn.execute(q).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def mark_reminder_done(reminder_id: int) -> None:
    conn = get_conn()
    conn.execute("UPDATE reminders SET is_done=1 WHERE id=?", (reminder_id,))
    conn.commit()
    conn.close()


def apply_discount(base_total: float, discount_pct: float) -> Dict[str, float]:
    discount_pct = max(DISCOUNT_MIN_PCT, min(DISCOUNT_MAX_PCT, float(discount_pct)))
    discount_amount = round(base_total * discount_pct / 100, 2)
    final_total = round(base_total - discount_amount, 2)
    return {"discount_pct": discount_pct, "discount_amount": discount_amount, "final_total": final_total}


@dataclass
class QuoteItem:
    code: str
    name: str
    unit: str = "шт"
    qty: float = 1
    price: float = 0.0

    @property
    def total(self) -> float:
        return round(self.price * self.qty, 2)


@dataclass
class QuoteRecord:
    id: Optional[int] = None
    kp_number: str = ""
    customer_id: Optional[int] = None
    product_type: str = "Кран"
    product_model: str = ""
    include_montage: bool = False
    delivery_city: str = ""
    base_total: float = 0.0
    discount_pct: float = 0.0
    discount_amount: float = 0.0
    final_total: float = 0.0
    status: str = STATUS_DRAFT
    notes: str = ""
    loss_reason: str = ""
    owner_id: Optional[int] = None
    owner_name: str = ""
    probability_pct: float = 50.0
    items: List[QuoteItem] = field(default_factory=list)
    pdf_bytes: Optional[bytes] = None   # байты PDF-файла (сохраняем в crm_files/)
    docx_bytes: Optional[bytes] = None  # байты DOCX-файла


def _log_history(conn, quote_id: int, field_changed: str, old_value: Any, new_value: Any,
                  user_id: Optional[int] = None, username: str = ""):
    now = dt.datetime.now().isoformat(timespec="seconds")
    conn.execute("""
        INSERT INTO quote_history (quote_id, user_id, username, field_changed, old_value, new_value, changed_at)
        VALUES (?,?,?,?,?,?,?)
    """, (quote_id, user_id, username, field_changed, str(old_value), str(new_value), now))


def save_quote(q: QuoteRecord, user_id: Optional[int] = None, username: str = "") -> int:
    disc = apply_discount(q.base_total, q.discount_pct)
    q.discount_pct = disc["discount_pct"]
    q.discount_amount = disc["discount_amount"]
    q.final_total = disc["final_total"]

    conn = get_conn()
    cur = conn.cursor()
    now = dt.datetime.now().isoformat(timespec="seconds")

    cur.execute("""
        INSERT INTO quotes (kp_number, customer_id, product_type, product_model, include_montage,
            delivery_city, base_total, discount_pct, discount_amount, final_total, status,
            created_at, updated_at, notes, loss_reason, owner_id, owner_name, probability_pct)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, (q.kp_number, q.customer_id, q.product_type, q.product_model, int(q.include_montage),
          q.delivery_city, q.base_total, q.discount_pct, q.discount_amount, q.final_total, q.status,
          now, now, q.notes, q.loss_reason, user_id or q.owner_id, username or q.owner_name, q.probability_pct))
    quote_id = cur.lastrowid

    for it in q.items:
        cur.execute("""
            INSERT INTO quote_items (quote_id, code, name, unit, qty, price, total) VALUES (?,?,?,?,?,?,?)
        """, (quote_id, it.code, it.name, it.unit, it.qty, it.price, it.total))

    # Сохраняем файлы PDF/DOCX на диск и прописываем пути в базу
    pdf_path = None
    docx_path = None
    if q.pdf_bytes:
        pdf_path = str(KP_FILES_DIR / f"kp_{quote_id}.pdf")
        with open(pdf_path, "wb") as f:
            f.write(q.pdf_bytes)
    if q.docx_bytes:
        docx_path = str(KP_FILES_DIR / f"kp_{quote_id}.docx")
        with open(docx_path, "wb") as f:
            f.write(q.docx_bytes)
    if pdf_path or docx_path:
        cur.execute("UPDATE quotes SET pdf_path=?, docx_path=? WHERE id=?",
                    (pdf_path, docx_path, quote_id))

    _log_history(conn, quote_id, "created", "", f"КП {q.kp_number} создан", user_id, username)
    conn.commit()
    conn.close()
    return quote_id


def delete_customer(customer_id: int) -> None:
    """Удаляет клиента вместе с его КП, позициями, продажами, заметками и напоминаниями."""
    conn = get_conn()
    cur = conn.cursor()
    quote_ids = [r["id"] for r in cur.execute(
        "SELECT id FROM quotes WHERE customer_id=?", (customer_id,)).fetchall()]
    for qid in quote_ids:
        cur.execute("DELETE FROM quote_items WHERE quote_id=?", (qid,))
        cur.execute("DELETE FROM sales WHERE quote_id=?", (qid,))
        cur.execute("DELETE FROM quote_history WHERE quote_id=?", (qid,))
        cur.execute("DELETE FROM reminders WHERE quote_id=?", (qid,))
    cur.execute("DELETE FROM quotes WHERE customer_id=?", (customer_id,))
    cur.execute("DELETE FROM customer_notes WHERE customer_id=?", (customer_id,))
    cur.execute("DELETE FROM reminders WHERE customer_id=?", (customer_id,))
    cur.execute("DELETE FROM customers WHERE id=?", (customer_id,))
    conn.commit()
    conn.close()


def delete_customer_note(note_id: int) -> None:
    """Удаляет одну заметку клиента."""
    conn = get_conn()
    conn.execute("DELETE FROM customer_notes WHERE id=?", (int(note_id),))
    conn.commit()
    conn.close()


def delete_reminder(reminder_id: int) -> None:
    """Удаляет одно напоминание."""
    conn = get_conn()
    conn.execute("DELETE FROM reminders WHERE id=?", (int(reminder_id),))
    conn.commit()
    conn.close()


def delete_quote(quote_id: int) -> None:
    """Удаляет КП с позициями, продажей, историей и файлами."""
    conn = get_conn()
    cur = conn.cursor()
    row = cur.execute("SELECT pdf_path, docx_path FROM quotes WHERE id=?",
                       (quote_id,)).fetchone()
    if row:
        for path_key in ("pdf_path", "docx_path"):
            p = row[path_key]
            if p and Path(p).exists():
                try:
                    Path(p).unlink()
                except Exception:
                    pass
    cur.execute("DELETE FROM quote_items WHERE quote_id=?", (quote_id,))
    cur.execute("DELETE FROM sales WHERE quote_id=?", (quote_id,))
    cur.execute("DELETE FROM quote_history WHERE quote_id=?", (quote_id,))
    cur.execute("DELETE FROM reminders WHERE quote_id=?", (quote_id,))
    cur.execute("DELETE FROM quotes WHERE id=?", (quote_id,))
    conn.commit()
    conn.close()


def get_kp_file_bytes(quote_id: int, file_type: str) -> Optional[bytes]:
    """Читает сохранённый файл (file_type = 'pdf' или 'docx')."""
    conn = get_conn()
    row = conn.execute(
        f"SELECT {file_type}_path FROM quotes WHERE id=?", (quote_id,)
    ).fetchone()
    conn.close()
    if not row or not row[f"{file_type}_path"]:
        return None
    path = Path(row[f"{file_type}_path"])
    if not path.exists():
        return None
    return path.read_bytes()


def _get_dadata_token() -> str:
    """Ищет токен DaData в нескольких местах (по приоритету)."""
    import os
    # 1. Переменная окружения
    t = os.environ.get("DADATA_TOKEN", "").strip()
    if t:
        return t
    # 2. Streamlit secrets (на Cloud)
    try:
        import streamlit as _st
        t = str(_st.secrets.get("DADATA_TOKEN", "") or "").strip()
        if t:
            return t
        t = str(_st.session_state.get("dadata_token", "") or "").strip()
        if t:
            return t
    except Exception:
        pass
    # 3. Файл .dadata_token в корне проекта
    try:
        from pathlib import Path
        f = Path(__file__).parent / ".dadata_token"
        if f.exists():
            return f.read_text(encoding="utf-8").strip()
    except Exception:
        pass
    return ""


class DaDataError(Exception):
    """Ошибка DaData с конкретным сообщением для UI."""
    pass


def fetch_by_inn_dadata(inn: str, token: str = "") -> Optional[Customer]:
    """Подгрузка реквизитов через DaData Suggestions API.

    Токен берётся: аргумент → DADATA_TOKEN env → st.secrets → st.session_state
    → файл .dadata_token. Если нигде нет — выбрасывает DaDataError.
    """
    import json
    import urllib.request
    import urllib.error

    inn = "".join(ch for ch in (inn or "") if ch.isdigit())
    if len(inn) not in (10, 12):
        raise DaDataError("ИНН должен содержать 10 цифр (юр.лицо) или 12 (ИП).")

    if not token:
        token = _get_dadata_token()
    if not token:
        raise DaDataError(
            "Токен DaData не найден. Зарегистрируйтесь на dadata.ru — "
            "бесплатно 10 000 запросов/сутки. Затем введите токен в Настройках "
            "или заполните поля вручную.")

    url = ("https://suggestions.dadata.ru/suggestions/api/4_1/rs/"
           "findById/party")
    headers = {"Content-Type": "application/json",
               "Accept": "application/json",
               "Authorization": f"Token {token}"}
    body = json.dumps({"query": inn, "count": 1}).encode("utf-8")
    req = urllib.request.Request(url, data=body, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=8) as resp:
            data = json.loads(resp.read().decode("utf-8"))
    except urllib.error.HTTPError as e:
        if e.code == 401:
            raise DaDataError(
                "Токен DaData отклонён (401 Unauthorized). "
                "Проверьте корректность токена в кабинете dadata.ru → АПИ.")
        elif e.code == 403:
            raise DaDataError("DaData: доступ запрещён (403). Токен верный?")
        elif e.code == 429:
            raise DaDataError("DaData: исчерпан лимит запросов на сегодня.")
        else:
            raise DaDataError(f"DaData HTTP {e.code}: {e.reason}")
    except Exception as e:
        raise DaDataError(f"Ошибка сети: {type(e).__name__}: {e}")
    suggestions = data.get("suggestions") or []
    if not suggestions:
        return None
    party = suggestions[0].get("data") or {}
    name_obj = party.get("name") or {}
    addr = party.get("address") or {}
    mgmt = party.get("management") or {}
    return Customer(
        inn=inn,
        kpp=(party.get("kpp") or "") or "",
        ogrn=(party.get("ogrn") or "") or "",
        name_full=(name_obj.get("full_with_opf")
                    or name_obj.get("full") or ""),
        name_short=(name_obj.get("short_with_opf")
                     or name_obj.get("short") or ""),
        address=(addr.get("unrestricted_value") or "") or "",
        director_position=(mgmt.get("post") or "") or "",
        director_fio=(mgmt.get("name") or "") or "",
    )


def update_quote_discount(quote_id: int, discount_pct: float, user_id: Optional[int] = None, username: str = "") -> None:
    conn = get_conn()
    row = conn.execute("SELECT base_total, discount_pct FROM quotes WHERE id=?", (quote_id,)).fetchone()
    if not row:
        conn.close()
        raise ValueError(f"КП id={quote_id} не найден")
    old_pct = row["discount_pct"]
    disc = apply_discount(row["base_total"], discount_pct)
    now = dt.datetime.now().isoformat(timespec="seconds")
    conn.execute("""
        UPDATE quotes SET discount_pct=?, discount_amount=?, final_total=?, updated_at=? WHERE id=?
    """, (disc["discount_pct"], disc["discount_amount"], disc["final_total"], now, quote_id))
    _log_history(conn, quote_id, "discount_pct", old_pct, disc["discount_pct"], user_id, username)
    conn.commit()
    conn.close()


def set_quote_status(quote_id: int, status: str, loss_reason: str = "",
                      user_id: Optional[int] = None, username: str = "") -> None:
    if status not in STATUS_LABELS:
        raise ValueError(f"Неизвестный статус: {status}")
    conn = get_conn()
    cur = conn.cursor()
    now = dt.datetime.now().isoformat(timespec="seconds")
    row = cur.execute("SELECT * FROM quotes WHERE id=?", (quote_id,)).fetchone()
    if not row:
        conn.close()
        raise ValueError(f"КП id={quote_id} не найден")

    old_status = row["status"]
    sold_at = now if status == STATUS_WON else row["sold_at"]
    final_loss_reason = loss_reason if status == STATUS_LOST else row["loss_reason"]

    cur.execute("UPDATE quotes SET status=?, updated_at=?, sold_at=?, loss_reason=? WHERE id=?",
                (status, now, sold_at, final_loss_reason, quote_id))
    _log_history(conn, quote_id, "status", old_status, status, user_id, username)
    if status == STATUS_LOST and loss_reason:
        _log_history(conn, quote_id, "loss_reason", row["loss_reason"], loss_reason, user_id, username)

    if status == STATUS_WON:
        exists = cur.execute("SELECT id FROM sales WHERE quote_id=?", (quote_id,)).fetchone()
        if not exists:
            cur.execute("""
                INSERT INTO sales (quote_id, customer_id, product_type, product_model, sale_date, price, discount_pct, delivery_city)
                VALUES (?,?,?,?,?,?,?,?)
            """, (quote_id, row["customer_id"], row["product_type"], row["product_model"], now,
                  row["final_total"], row["discount_pct"], row["delivery_city"]))

    conn.commit()
    conn.close()


def get_quote(quote_id: int) -> Optional[Dict[str, Any]]:
    conn = get_conn()
    row = conn.execute("SELECT * FROM quotes WHERE id=?", (quote_id,)).fetchone()
    if not row:
        conn.close()
        return None
    items = conn.execute("SELECT * FROM quote_items WHERE quote_id=?", (quote_id,)).fetchall()
    conn.close()
    result = dict(row)
    result["items"] = [dict(i) for i in items]
    return result


def generate_unique_kp_number(base_number: str) -> str:
    """Генерирует уникальный номер КП.

    Если base_number уже есть в базе — добавляет суффикс -2, -3, …
    Пример: "22072026/ЛКС" → "22072026-2/ЛКС" → "22072026-3/ЛКС".
    """
    conn = get_conn()
    cur = conn.cursor()
    existing = {r[0] for r in cur.execute(
        "SELECT kp_number FROM quotes").fetchall()}
    conn.close()

    if base_number not in existing:
        return base_number

    # Разбиваем на части: «date» + «/ПРЕФ»
    if "/" in base_number:
        date_part, suffix = base_number.split("/", 1)
        prefix_slash = "/" + suffix
    else:
        date_part = base_number
        prefix_slash = ""

    n = 2
    while True:
        candidate = f"{date_part}-{n}{prefix_slash}"
        if candidate not in existing:
            return candidate
        n += 1
        if n > 999:
            # аварийный выход
            import time
            return f"{date_part}-{int(time.time())}{prefix_slash}"


def update_quote_items(quote_id: int, items: List[Dict[str, Any]],
                       user_id: Optional[int] = None, username: str = "") -> Dict[str, float]:
    """Полностью перезаписывает позиции КП + пересчитывает base_total и final_total.

    items: [{code, name, unit, qty, price}, ...]
    Возвращает {'base_total', 'final_total'}.
    """
    conn = get_conn()
    cur = conn.cursor()
    # Старый base_total для истории
    row = cur.execute("SELECT base_total, discount_pct FROM quotes WHERE id=?",
                     (quote_id,)).fetchone()
    if not row:
        conn.close()
        raise ValueError(f"quote_id={quote_id} не найден")
    old_base = float(row[0] or 0.0)
    discount_pct = float(row[1] or 0.0)

    # Перезаписываем позиции
    cur.execute("DELETE FROM quote_items WHERE quote_id=?", (quote_id,))
    new_base = 0.0
    _inserted = 0
    for it in items:
        # Безопасное преобразование (None → 0, строка → float, NaN → 0)
        def _num(v):
            if v is None:
                return 0.0
            try:
                f = float(v)
                if f != f:  # NaN
                    return 0.0
                return f
            except (TypeError, ValueError):
                return 0.0

        qty = _num(it.get("qty"))
        price = _num(it.get("price"))
        code = str(it.get("code") or "").strip()
        name = str(it.get("name") or "").strip()
        unit = str(it.get("unit") or "шт").strip() or "шт"

        # Пропускаем полностью пустые строки (нет наименования и нет кода)
        if not name and not code:
            continue

        total = qty * price
        new_base += total
        _inserted += 1
        cur.execute(
            "INSERT INTO quote_items (quote_id, code, name, unit, qty, price, total) "
            "VALUES (?,?,?,?,?,?,?)",
            (quote_id, code, name, unit, qty, price, total),
        )
    # Пересчёт final_total с той же скидкой
    disc = apply_discount(new_base, discount_pct)
    cur.execute(
        "UPDATE quotes SET base_total=?, discount_amount=?, final_total=? WHERE id=?",
        (new_base, disc["discount_amount"], disc["final_total"], quote_id),
    )
    _log_history(conn, quote_id, "items_edited",
                 f"base_total {old_base:.2f}",
                 f"base_total {new_base:.2f} ({_inserted} позиций)",
                 user_id, username)
    conn.commit()
    conn.close()
    return {"base_total": new_base, "final_total": disc["final_total"]}


def get_quote_history(quote_id: int) -> List[Dict[str, Any]]:
    conn = get_conn()
    rows = conn.execute("SELECT * FROM quote_history WHERE quote_id=? ORDER BY changed_at DESC", (quote_id,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def update_quote_contact_info(quote_id: int, *,
                              contact_fio: Optional[str] = None,
                              last_contact_at: Optional[str] = None,
                              request_summary: Optional[str] = None,
                              customer_phone: Optional[str] = None,
                              customer_email: Optional[str] = None) -> None:
    """Обновляет контактные поля КП и/или клиента. None — поле не меняется."""
    conn = get_conn()
    cur = conn.cursor()
    updates = []
    params: List[Any] = []
    if contact_fio is not None:
        updates.append("contact_fio=?"); params.append(contact_fio)
    if last_contact_at is not None:
        updates.append("last_contact_at=?"); params.append(last_contact_at)
    if request_summary is not None:
        updates.append("request_summary=?"); params.append(request_summary)
    if updates:
        params.append(dt.datetime.now().isoformat(timespec="seconds"))
        updates.append("updated_at=?")
        params.append(quote_id)
        cur.execute(f"UPDATE quotes SET {', '.join(updates)} WHERE id=?", params)
    # Клиентские поля — в таблицу customers
    row = cur.execute("SELECT customer_id FROM quotes WHERE id=?", (quote_id,)).fetchone()
    if row and row[0]:
        cu_updates = []
        cu_params: List[Any] = []
        if customer_phone is not None:
            cu_updates.append("phone=?"); cu_params.append(customer_phone)
        if customer_email is not None:
            cu_updates.append("email=?"); cu_params.append(customer_email)
        if cu_updates:
            cu_params.append(dt.datetime.now().isoformat(timespec="seconds"))
            cu_updates.append("updated_at=?")
            cu_params.append(row[0])
            cur.execute(f"UPDATE customers SET {', '.join(cu_updates)} WHERE id=?", cu_params)
    conn.commit()
    conn.close()


def delete_quote(quote_id: int, user_id: Optional[int] = None, username: str = "") -> bool:
    """Полное удаление КП (вместе с позициями, историей, напоминаниями).
    Файлы в crm_files/ тоже удаляем."""
    conn = get_conn()
    cur = conn.cursor()
    # Удаляем файлы
    try:
        from pathlib import Path
        for row in cur.execute(
            "SELECT file_path FROM quote_files WHERE quote_id=?", (quote_id,)
        ).fetchall():
            p = Path(row[0])
            if p.exists():
                p.unlink()
    except Exception:
        pass
    cur.execute("DELETE FROM quote_items WHERE quote_id=?", (quote_id,))
    cur.execute("DELETE FROM quote_history WHERE quote_id=?", (quote_id,))
    cur.execute("DELETE FROM reminders WHERE quote_id=?", (quote_id,))
    try:
        cur.execute("DELETE FROM quote_files WHERE quote_id=?", (quote_id,))
    except Exception:
        pass
    cur.execute("DELETE FROM sales WHERE quote_id=?", (quote_id,))
    cur.execute("DELETE FROM quotes WHERE id=?", (quote_id,))
    conn.commit()
    conn.close()
    return True


def list_quotes(customer_id: Optional[int] = None, status: Optional[str] = None, owner_id: Optional[int] = None) -> List[Dict[str, Any]]:
    conn = get_conn()
    q = """
        SELECT quotes.*, customers.name_short, customers.phone, customers.email
        FROM quotes LEFT JOIN customers ON customers.id = quotes.customer_id WHERE 1=1
    """
    params: List[Any] = []
    if customer_id is not None:
        q += " AND quotes.customer_id = ?"
        params.append(customer_id)
    if status is not None:
        q += " AND quotes.status = ?"
        params.append(status)
    if owner_id is not None:
        q += " AND quotes.owner_id = ?"
        params.append(owner_id)
    q += " ORDER BY quotes.created_at DESC"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def list_sales(date_from: Optional[str] = None, date_to: Optional[str] = None) -> List[Dict[str, Any]]:
    conn = get_conn()
    q = """
        SELECT sales.*, customers.name_short, customers.phone, customers.email
        FROM sales LEFT JOIN customers ON customers.id = sales.customer_id WHERE 1=1
    """
    params: List[Any] = []
    if date_from:
        q += " AND sales.sale_date >= ?"
        params.append(date_from)
    if date_to:
        q += " AND sales.sale_date <= ?"
        params.append(date_to)
    q += " ORDER BY sales.sale_date DESC"
    rows = conn.execute(q, params).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def sales_summary() -> Dict[str, Any]:
    conn = get_conn()
    row = conn.execute("SELECT COUNT(*) cnt, COALESCE(SUM(price),0) total, COALESCE(AVG(discount_pct),0) avg_discount FROM sales").fetchone()
    conn.close()
    return {"cnt": row["cnt"], "total": row["total"], "avg_discount": row["avg_discount"]}


def sales_by_month() -> List[Dict[str, Any]]:
    conn = get_conn()
    rows = conn.execute("""
        SELECT strftime('%Y-%m', sale_date) AS month, SUM(price) AS total, COUNT(*) AS cnt
        FROM sales GROUP BY month ORDER BY month
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def sales_by_product_type() -> List[Dict[str, Any]]:
    conn = get_conn()
    rows = conn.execute("""
        SELECT product_type, COUNT(*) cnt, SUM(price) total FROM sales GROUP BY product_type ORDER BY total DESC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def conversion_funnel() -> Dict[str, int]:
    conn = get_conn()
    rows = conn.execute("SELECT status, COUNT(*) cnt FROM quotes GROUP BY status").fetchall()
    conn.close()
    result = {s: 0 for s in STATUS_ORDER}
    for r in rows:
        result[r["status"]] = r["cnt"]
    return result


def loss_reason_breakdown() -> List[Dict[str, Any]]:
    conn = get_conn()
    rows = conn.execute("""
        SELECT loss_reason, COUNT(*) cnt FROM quotes WHERE status=? AND loss_reason != '' AND loss_reason IS NOT NULL
        GROUP BY loss_reason ORDER BY cnt DESC
    """, (STATUS_LOST,)).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def manager_report() -> List[Dict[str, Any]]:
    conn = get_conn()
    rows = conn.execute("""
        SELECT owner_name, COUNT(*) cnt, COALESCE(SUM(CASE WHEN status='won' THEN final_total ELSE 0 END),0) won_total,
               SUM(CASE WHEN status='won' THEN 1 ELSE 0 END) won_cnt
        FROM quotes WHERE owner_name IS NOT NULL AND owner_name != '' GROUP BY owner_name ORDER BY won_total DESC
    """).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def revenue_forecast() -> Dict[str, Any]:
    conn = get_conn()
    rows = conn.execute("""
        SELECT final_total, probability_pct FROM quotes WHERE status IN ('draft','sent')
    """).fetchall()
    conn.close()
    total_pipeline = sum(r["final_total"] for r in rows)
    weighted = sum(r["final_total"] * (r["probability_pct"] or 50) / 100 for r in rows)
    return {"pipeline_total": total_pipeline, "weighted_forecast": weighted, "cnt": len(rows)}


def update_quote_probability(quote_id: int, probability_pct: float) -> None:
    conn = get_conn()
    conn.execute("UPDATE quotes SET probability_pct=? WHERE id=?", (probability_pct, quote_id))
    conn.commit()
    conn.close()


def export_sales_to_dataframe():
    import pandas as pd
    rows = list_sales()
    if not rows:
        return pd.DataFrame()
    return pd.DataFrame(rows)


# --- Автоматическая инициализация БД при импорте модуля ---
# Гарантирует что все таблицы созданы до любой операции (save_quote,
# generate_unique_kp_number и т.д.). Идемпотентно (CREATE IF NOT EXISTS).
try:
    init_db()
except Exception as _e:
    import sys as _sys
    print(f"[crm_db] Warning: init_db failed on import: {_e}", file=_sys.stderr)
