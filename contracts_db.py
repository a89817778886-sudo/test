"""
База данных договоров.

Каждый сгенерированный договор автоматически сохраняется в таблицу contracts
с параметрами и DOCX-содержимым, чтобы можно было:
- посмотреть список
- открыть карточку, отредактировать параметры
- перегенерировать DOCX
- скачать
- удалить
"""
from __future__ import annotations
import json
import sqlite3
import datetime as dt
from typing import Optional
from pathlib import Path


DB_PATH = Path(__file__).parent / "crm.db"


def _get_conn():
    """Возвращает соединение к БД (SQLite локально или PostgreSQL если задан SUPABASE_DB_URL)."""
    try:
        from crm_db import get_conn as _get_shared_conn
        return _get_shared_conn()
    except Exception:
        c = sqlite3.connect(str(DB_PATH))
        c.row_factory = sqlite3.Row
        return c


# ==================== СХЕМА ====================

CONTRACTS_SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS contracts (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    contract_number TEXT NOT NULL,
    contract_date TEXT NOT NULL,
    buyer_short TEXT,
    buyer_full TEXT,
    buyer_inn TEXT,
    buyer_kpp TEXT,
    supplier_key TEXT,
    supplier_short TEXT,
    total_amount REAL,
    has_vat INTEGER DEFAULT 1,
    vat_amount REAL DEFAULT 0,
    prepay_pct INTEGER DEFAULT 100,
    shipment_days INTEGER DEFAULT 20,
    warranty_months INTEGER DEFAULT 12,
    delivery_terms TEXT,
    delivery_address TEXT,
    kp_id INTEGER,
    kp_number TEXT,
    contract_type TEXT DEFAULT 'КП',
    params_json TEXT,
    docx_blob BLOB,
    created_at TEXT NOT NULL,
    updated_at TEXT
);
"""

CONTRACTS_SCHEMA_PG = """
CREATE TABLE IF NOT EXISTS contracts (
    id SERIAL PRIMARY KEY,
    contract_number TEXT NOT NULL,
    contract_date TEXT NOT NULL,
    buyer_short TEXT,
    buyer_full TEXT,
    buyer_inn TEXT,
    buyer_kpp TEXT,
    supplier_key TEXT,
    supplier_short TEXT,
    total_amount DOUBLE PRECISION,
    has_vat INTEGER DEFAULT 1,
    vat_amount DOUBLE PRECISION DEFAULT 0,
    prepay_pct INTEGER DEFAULT 100,
    shipment_days INTEGER DEFAULT 20,
    warranty_months INTEGER DEFAULT 12,
    delivery_terms TEXT,
    delivery_address TEXT,
    kp_id INTEGER,
    kp_number TEXT,
    contract_type TEXT DEFAULT 'КП',
    params_json TEXT,
    docx_blob BYTEA,
    created_at TEXT NOT NULL,
    updated_at TEXT
);
"""


def init_schema():
    """Создаёт таблицу contracts если не существует."""
    conn = _get_conn()
    try:
        cur = conn.cursor()
        # Определяем тип БД (SQLite или PostgreSQL)
        try:
            # PostgreSQL psycopg2 конкретный тип
            import psycopg2.extensions
            if isinstance(conn, psycopg2.extensions.connection):
                cur.execute(CONTRACTS_SCHEMA_PG)
            else:
                cur.executescript(CONTRACTS_SCHEMA_SQLITE)
        except Exception:
            try:
                cur.executescript(CONTRACTS_SCHEMA_SQLITE)
            except Exception:
                cur.execute(CONTRACTS_SCHEMA_PG)
        conn.commit()
    finally:
        conn.close()


# ==================== СОХРАНЕНИЕ ====================

def save_contract(
    *,
    contract_number: str,
    contract_date: str,
    buyer: dict,
    supplier: dict,
    lines: list,
    total_amount: float,
    has_vat: bool,
    vat_amount: float,
    prepay_pct: int = 100,
    shipment_days: int = 20,
    warranty_months: int = 12,
    delivery_terms: str = "",
    delivery_address: str = "",
    kp_id: Optional[int] = None,
    kp_number: str = "",
    contract_type: str = "КП",
    docx_bytes: bytes = None,
    extra_params: dict = None,
) -> int:
    """Сохраняет договор в БД, возвращает id."""
    init_schema()

    params = {
        "buyer": buyer,
        "supplier_key": supplier.get("label", ""),
        "lines": lines,
        "extra": extra_params or {},
    }
    params_json = json.dumps(params, ensure_ascii=False, default=str)

    now = dt.datetime.now().isoformat(timespec="seconds")

    conn = _get_conn()
    try:
        cur = conn.cursor()
        # Определяем поставщика по метке
        supplier_key = "LKS"
        label = supplier.get("label", "")
        if "МОДЕРНИЗАЦ" in label.upper(): supplier_key = "MODERNIZATSIYA"
        elif "КИНЕМАТ" in label.upper(): supplier_key = "KINEMATIKA"

        # Проверяем есть ли уже такой договор (по contract_number)
        try:
            cur.execute("SELECT id FROM contracts WHERE contract_number = %s LIMIT 1",
                        (contract_number,))
        except Exception:
            cur.execute("SELECT id FROM contracts WHERE contract_number = ? LIMIT 1",
                        (contract_number,))
        existing = cur.fetchone()

        if existing:
            # Обновляем существующий
            eid = existing[0] if not hasattr(existing, "keys") else existing["id"]
            try:
                cur.execute("""UPDATE contracts SET
                    contract_date=%s, buyer_short=%s, buyer_full=%s, buyer_inn=%s, buyer_kpp=%s,
                    supplier_key=%s, supplier_short=%s, total_amount=%s, has_vat=%s, vat_amount=%s,
                    prepay_pct=%s, shipment_days=%s, warranty_months=%s,
                    delivery_terms=%s, delivery_address=%s,
                    kp_id=%s, kp_number=%s, contract_type=%s,
                    params_json=%s, docx_blob=%s, updated_at=%s
                    WHERE id=%s""",
                    (contract_date, buyer.get("short",""), buyer.get("full",""),
                     buyer.get("inn",""), buyer.get("kpp",""),
                     supplier_key, supplier.get("short",""), float(total_amount),
                     1 if has_vat else 0, float(vat_amount),
                     int(prepay_pct), int(shipment_days), int(warranty_months),
                     delivery_terms, delivery_address,
                     kp_id, kp_number, contract_type,
                     params_json, docx_bytes, now, eid))
            except Exception:
                cur.execute("""UPDATE contracts SET
                    contract_date=?, buyer_short=?, buyer_full=?, buyer_inn=?, buyer_kpp=?,
                    supplier_key=?, supplier_short=?, total_amount=?, has_vat=?, vat_amount=?,
                    prepay_pct=?, shipment_days=?, warranty_months=?,
                    delivery_terms=?, delivery_address=?,
                    kp_id=?, kp_number=?, contract_type=?,
                    params_json=?, docx_blob=?, updated_at=?
                    WHERE id=?""",
                    (contract_date, buyer.get("short",""), buyer.get("full",""),
                     buyer.get("inn",""), buyer.get("kpp",""),
                     supplier_key, supplier.get("short",""), float(total_amount),
                     1 if has_vat else 0, float(vat_amount),
                     int(prepay_pct), int(shipment_days), int(warranty_months),
                     delivery_terms, delivery_address,
                     kp_id, kp_number, contract_type,
                     params_json, docx_bytes, now, eid))
            conn.commit()
            return int(eid)
        else:
            try:
                cur.execute("""INSERT INTO contracts
                    (contract_number, contract_date, buyer_short, buyer_full, buyer_inn, buyer_kpp,
                     supplier_key, supplier_short, total_amount, has_vat, vat_amount,
                     prepay_pct, shipment_days, warranty_months,
                     delivery_terms, delivery_address, kp_id, kp_number, contract_type,
                     params_json, docx_blob, created_at)
                    VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)
                    RETURNING id""",
                    (contract_number, contract_date, buyer.get("short",""), buyer.get("full",""),
                     buyer.get("inn",""), buyer.get("kpp",""),
                     supplier_key, supplier.get("short",""), float(total_amount),
                     1 if has_vat else 0, float(vat_amount),
                     int(prepay_pct), int(shipment_days), int(warranty_months),
                     delivery_terms, delivery_address, kp_id, kp_number, contract_type,
                     params_json, docx_bytes, now))
                new_id = cur.fetchone()[0]
            except Exception:
                cur.execute("""INSERT INTO contracts
                    (contract_number, contract_date, buyer_short, buyer_full, buyer_inn, buyer_kpp,
                     supplier_key, supplier_short, total_amount, has_vat, vat_amount,
                     prepay_pct, shipment_days, warranty_months,
                     delivery_terms, delivery_address, kp_id, kp_number, contract_type,
                     params_json, docx_blob, created_at)
                    VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)""",
                    (contract_number, contract_date, buyer.get("short",""), buyer.get("full",""),
                     buyer.get("inn",""), buyer.get("kpp",""),
                     supplier_key, supplier.get("short",""), float(total_amount),
                     1 if has_vat else 0, float(vat_amount),
                     int(prepay_pct), int(shipment_days), int(warranty_months),
                     delivery_terms, delivery_address, kp_id, kp_number, contract_type,
                     params_json, docx_bytes, now))
                new_id = cur.lastrowid
            conn.commit()
            return int(new_id)
    finally:
        conn.close()


# ==================== ЧТЕНИЕ ====================

def list_contracts(limit: int = 500) -> list[dict]:
    """Список всех договоров, новые сверху."""
    init_schema()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        cur.execute("""SELECT id, contract_number, contract_date, buyer_short, buyer_inn,
                              supplier_short, supplier_key, total_amount, has_vat, vat_amount,
                              prepay_pct, shipment_days, warranty_months,
                              delivery_terms, delivery_address, kp_id, kp_number, contract_type,
                              created_at, updated_at
                       FROM contracts ORDER BY created_at DESC, id DESC""")
        rows = cur.fetchall()
        result = []
        for r in rows:
            if hasattr(r, "keys"):
                result.append(dict(r))
            else:
                result.append({
                    "id": r[0], "contract_number": r[1], "contract_date": r[2],
                    "buyer_short": r[3], "buyer_inn": r[4],
                    "supplier_short": r[5], "supplier_key": r[6],
                    "total_amount": r[7], "has_vat": r[8], "vat_amount": r[9],
                    "prepay_pct": r[10], "shipment_days": r[11], "warranty_months": r[12],
                    "delivery_terms": r[13], "delivery_address": r[14],
                    "kp_id": r[15], "kp_number": r[16], "contract_type": r[17],
                    "created_at": r[18], "updated_at": r[19],
                })
        return result[:limit]
    finally:
        conn.close()


def get_contract(contract_id: int) -> Optional[dict]:
    """Полная карточка договора (с params_json и docx_blob)."""
    init_schema()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        try:
            cur.execute("SELECT * FROM contracts WHERE id = %s", (contract_id,))
        except Exception:
            cur.execute("SELECT * FROM contracts WHERE id = ?", (contract_id,))
        row = cur.fetchone()
        if not row:
            return None
        if hasattr(row, "keys"):
            d = dict(row)
        else:
            # Собираем по названиям колонок
            cols = [c[0] for c in cur.description]
            d = dict(zip(cols, row))
        # Разбираем params_json
        try:
            d["params"] = json.loads(d.get("params_json") or "{}")
        except Exception:
            d["params"] = {}
        return d
    finally:
        conn.close()


def delete_contract(contract_id: int) -> bool:
    """Удаляет договор."""
    conn = _get_conn()
    try:
        cur = conn.cursor()
        try:
            cur.execute("DELETE FROM contracts WHERE id = %s", (contract_id,))
        except Exception:
            cur.execute("DELETE FROM contracts WHERE id = ?", (contract_id,))
        conn.commit()
        return True
    finally:
        conn.close()


def update_contract_docx(contract_id: int, docx_bytes: bytes,
                         total_amount: float = None, vat_amount: float = None) -> bool:
    """Обновить DOCX (после перегенерации) и суммы."""
    conn = _get_conn()
    now = dt.datetime.now().isoformat(timespec="seconds")
    try:
        cur = conn.cursor()
        if total_amount is not None and vat_amount is not None:
            try:
                cur.execute("UPDATE contracts SET docx_blob=%s, total_amount=%s, vat_amount=%s, updated_at=%s WHERE id=%s",
                            (docx_bytes, float(total_amount), float(vat_amount), now, contract_id))
            except Exception:
                cur.execute("UPDATE contracts SET docx_blob=?, total_amount=?, vat_amount=?, updated_at=? WHERE id=?",
                            (docx_bytes, float(total_amount), float(vat_amount), now, contract_id))
        else:
            try:
                cur.execute("UPDATE contracts SET docx_blob=%s, updated_at=%s WHERE id=%s",
                            (docx_bytes, now, contract_id))
            except Exception:
                cur.execute("UPDATE contracts SET docx_blob=?, updated_at=? WHERE id=?",
                            (docx_bytes, now, contract_id))
        conn.commit()
        return True
    finally:
        conn.close()


def update_contract_params(contract_id: int, params_json: str,
                           lines_summary: dict = None) -> bool:
    """Обновить параметры договора (без regen DOCX)."""
    conn = _get_conn()
    now = dt.datetime.now().isoformat(timespec="seconds")
    try:
        cur = conn.cursor()
        if lines_summary:
            try:
                cur.execute("""UPDATE contracts SET params_json=%s, total_amount=%s, vat_amount=%s,
                    prepay_pct=%s, shipment_days=%s, warranty_months=%s,
                    delivery_terms=%s, delivery_address=%s, updated_at=%s WHERE id=%s""",
                    (params_json, float(lines_summary.get("total", 0)), float(lines_summary.get("vat", 0)),
                     int(lines_summary.get("prepay_pct", 100)),
                     int(lines_summary.get("shipment_days", 20)),
                     int(lines_summary.get("warranty_months", 12)),
                     lines_summary.get("delivery_terms", ""),
                     lines_summary.get("delivery_address", ""), now, contract_id))
            except Exception:
                cur.execute("""UPDATE contracts SET params_json=?, total_amount=?, vat_amount=?,
                    prepay_pct=?, shipment_days=?, warranty_months=?,
                    delivery_terms=?, delivery_address=?, updated_at=? WHERE id=?""",
                    (params_json, float(lines_summary.get("total", 0)), float(lines_summary.get("vat", 0)),
                     int(lines_summary.get("prepay_pct", 100)),
                     int(lines_summary.get("shipment_days", 20)),
                     int(lines_summary.get("warranty_months", 12)),
                     lines_summary.get("delivery_terms", ""),
                     lines_summary.get("delivery_address", ""), now, contract_id))
        else:
            try:
                cur.execute("UPDATE contracts SET params_json=%s, updated_at=%s WHERE id=%s",
                            (params_json, now, contract_id))
            except Exception:
                cur.execute("UPDATE contracts SET params_json=?, updated_at=? WHERE id=?",
                            (params_json, now, contract_id))
        conn.commit()
        return True
    finally:
        conn.close()
