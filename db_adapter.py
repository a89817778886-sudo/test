# -*- coding: utf-8 -*-
"""
db_adapter.py — тонкий адаптер: даёт SQLite-совместимый интерфейс над PostgreSQL.

Используется в crm_db.py через мокирование sqlite3.connect(): если задан
секрет SUPABASE_DB_URL (или переменная окружения) — используем PostgreSQL,
иначе — обычный SQLite файл.

Что перехватывается:
- Cursor.execute(sql, params) — заменяет '?' на '%s', ловит SQLite-специфику
  (PRAGMA, AUTOINCREMENT, ON CONFLICT, INSERT OR IGNORE и т.п.)
- lastrowid — возвращается через RETURNING id
- row_factory=Row — эмулируется через RealDictCursor
- executemany — работает как в sqlite3

Совместимо с использованием conn.execute(...), conn.row_factory, cur.execute(...),
cur.executemany(...), cur.fetchone(), cur.fetchall(), cur.lastrowid.
"""
from __future__ import annotations
import os
import re
import threading
from typing import Any, Optional, Sequence

try:
    import streamlit as st
    _has_streamlit = True
except Exception:
    _has_streamlit = False

_POOL = None
_POOL_LOCK = threading.Lock()


def _get_dsn() -> Optional[str]:
    """Возвращает DSN для PostgreSQL из secrets или env, или None если не задан."""
    # 1) Streamlit secrets
    if _has_streamlit:
        try:
            if "SUPABASE_DB_URL" in st.secrets:
                return str(st.secrets["SUPABASE_DB_URL"]).strip()
        except Exception:
            pass
    # 2) Переменная окружения
    dsn = os.environ.get("SUPABASE_DB_URL", "").strip()
    if dsn:
        return dsn
    # 3) Локальный файл .supabase_dsn (для тестов)
    from pathlib import Path
    p = Path(__file__).resolve().parent / ".supabase_dsn"
    if p.exists():
        try:
            return p.read_text(encoding="utf-8").strip()
        except Exception:
            pass
    return None


def is_pg_enabled() -> bool:
    return _get_dsn() is not None


# ============================================================
# Трансляция SQLite → PostgreSQL
# ============================================================

# Регекс для замены ? на %s (не трогая ? внутри строковых литералов)
def _sqlite_to_pg(sql: str) -> str:
    # 1) PRAGMA — превращаем в безобидный no-op
    if sql.strip().upper().startswith("PRAGMA"):
        return "SELECT 1 WHERE FALSE"
    # 2) INSERT OR IGNORE → INSERT ... ON CONFLICT DO NOTHING
    sql = re.sub(r"\bINSERT\s+OR\s+IGNORE\b", "INSERT", sql, flags=re.IGNORECASE)
    _has_or_ignore = bool(re.search(r"INSERT\s+INTO\s+\w+", sql, re.IGNORECASE))
    # 3) INSERT OR REPLACE → эмулируется через ON CONFLICT DO UPDATE (упрощённо)
    sql = re.sub(r"\bINSERT\s+OR\s+REPLACE\b", "INSERT", sql, flags=re.IGNORECASE)
    # 4) AUTOINCREMENT INTEGER PRIMARY KEY уже не встречается в runtime-запросах
    # 5) datetime('now') → NOW()
    sql = re.sub(r"datetime\('now'\)", "NOW()", sql, flags=re.IGNORECASE)
    # 5b) strftime('%Y-%m', col) → TO_CHAR(col::timestamp, 'YYYY-MM')
    def _strftime_to_pg(m):
        fmt = m.group(1)
        col = m.group(2).strip()
        # Маппинг SQLite → PostgreSQL форматов
        _map = {"%Y": "YYYY", "%m": "MM", "%d": "DD",
                "%H": "HH24", "%M": "MI", "%S": "SS"}
        pg_fmt = fmt
        for k, v in _map.items():
            pg_fmt = pg_fmt.replace(k, v)
        return f"TO_CHAR({col}::timestamp, '{pg_fmt}')"
    sql = re.sub(r"strftime\(\s*'([^']+)'\s*,\s*([^)]+)\)", _strftime_to_pg, sql, flags=re.IGNORECASE)
    # 5c) julianday(a) - julianday(b) → EXTRACT(EPOCH FROM (a::timestamp - b::timestamp))/86400
    sql = re.sub(
        r"julianday\(\s*([^)]+?)\s*\)\s*-\s*julianday\(\s*([^)]+?)\s*\)",
        r"EXTRACT(EPOCH FROM (\1::timestamp - \2::timestamp))/86400",
        sql, flags=re.IGNORECASE)
    # 5d) date('now') → CURRENT_DATE
    sql = re.sub(r"date\('now'\)", "CURRENT_DATE", sql, flags=re.IGNORECASE)
    # 6) Замена ? на %s (учитываем строки и комментарии)
    result = []
    i = 0
    in_str = None
    while i < len(sql):
        ch = sql[i]
        if in_str:
            result.append(ch)
            if ch == in_str and (i + 1 >= len(sql) or sql[i+1] != in_str):
                in_str = None
            elif ch == in_str and i + 1 < len(sql) and sql[i+1] == in_str:
                # экранированная кавычка
                result.append(sql[i+1])
                i += 1
            i += 1
            continue
        if ch in ("'", '"'):
            in_str = ch
            result.append(ch)
            i += 1
            continue
        if ch == '?':
            result.append('%s')
            i += 1
            continue
        result.append(ch)
        i += 1
    return "".join(result)


# ============================================================
# Wrapper для psycopg2
# ============================================================

class _RowShim(dict):
    """Эмулирует sqlite3.Row: доступ и по индексу, и по имени."""
    __slots__ = ('_keys',)

    def __init__(self, keys, values):
        super().__init__(zip(keys, values))
        object.__setattr__(self, '_keys', tuple(keys))

    def __getitem__(self, key):
        if isinstance(key, int):
            return super().__getitem__(self._keys[key])
        return super().__getitem__(key)

    def keys(self):
        return self._keys


class PgCursor:
    """SQLite-совместимый курсор для PostgreSQL."""

    def __init__(self, conn: 'PgConnection'):
        self._conn = conn
        self._raw = conn._raw.cursor()
        self._description = None
        self._last_query_is_insert = False
        self._lastrowid = None
        self._pending_rows = None

    def execute(self, sql: str, params: Sequence = ()):
        pg_sql = _sqlite_to_pg(sql)
        # Для INSERT — добавляем RETURNING id (для lastrowid)
        m = re.match(r"^\s*INSERT\s+INTO\s+", pg_sql, re.IGNORECASE)
        self._last_query_is_insert = bool(m)
        if self._last_query_is_insert and "RETURNING" not in pg_sql.upper():
            # Ищем ON CONFLICT — если есть DO NOTHING, RETURNING может вернуть 0 строк
            pg_sql = pg_sql.rstrip().rstrip(";") + " RETURNING id"

        try:
            self._raw.execute(pg_sql, tuple(params) if params else None)
        except Exception:
            # Если ошибка с RETURNING (например, таблица без id) — повторим без RETURNING
            if self._last_query_is_insert and pg_sql.endswith(" RETURNING id"):
                self._conn._raw.rollback()
                pg_sql2 = pg_sql[: -len(" RETURNING id")]
                self._raw.execute(pg_sql2, tuple(params) if params else None)
            else:
                raise

        # lastrowid
        self._lastrowid = None
        if self._last_query_is_insert:
            try:
                if self._raw.description:
                    _row = self._raw.fetchone()
                    if _row:
                        self._lastrowid = _row[0] if not isinstance(_row, dict) else _row.get("id")
            except Exception:
                pass

        # description для итерации
        self._description = self._raw.description
        return self

    def executemany(self, sql: str, seq_of_params):
        pg_sql = _sqlite_to_pg(sql)
        # RETURNING в executemany не годится
        self._raw.executemany(pg_sql, [tuple(p) for p in seq_of_params])
        self._description = self._raw.description
        return self

    def executescript(self, script: str):
        """SQLite executescript: несколько statements через ;"""
        # Разбиваем на statements
        statements = [s.strip() for s in script.split(";") if s.strip()]
        for stmt in statements:
            self.execute(stmt)
        return self

    def fetchone(self):
        try:
            row = self._raw.fetchone()
        except Exception:
            return None
        if row is None:
            return None
        if self._conn._row_factory:
            keys = [d[0] for d in (self._raw.description or [])]
            return _RowShim(keys, row)
        return row

    def fetchall(self):
        try:
            rows = self._raw.fetchall()
        except Exception:
            return []
        if self._conn._row_factory:
            keys = [d[0] for d in (self._raw.description or [])]
            return [_RowShim(keys, r) for r in rows]
        return rows

    def fetchmany(self, size=None):
        try:
            rows = self._raw.fetchmany(size) if size else self._raw.fetchmany()
        except Exception:
            return []
        if self._conn._row_factory:
            keys = [d[0] for d in (self._raw.description or [])]
            return [_RowShim(keys, r) for r in rows]
        return rows

    def close(self):
        try:
            self._raw.close()
        except Exception:
            pass

    def __iter__(self):
        return iter(self.fetchall())

    @property
    def lastrowid(self):
        return self._lastrowid

    @property
    def description(self):
        return self._description

    @property
    def rowcount(self):
        return self._raw.rowcount


class PgConnection:
    """SQLite-совместимый Connection для PostgreSQL."""

    def __init__(self, dsn: str):
        import psycopg2
        self._raw = psycopg2.connect(dsn, connect_timeout=15)
        self._raw.autocommit = False
        self._row_factory = None

    @property
    def row_factory(self):
        return self._row_factory

    @row_factory.setter
    def row_factory(self, factory):
        # sqlite3.Row → эмулируем через _RowShim
        self._row_factory = factory

    def cursor(self):
        return PgCursor(self)

    def execute(self, sql: str, params: Sequence = ()):
        cur = self.cursor()
        cur.execute(sql, params)
        return cur

    def executemany(self, sql: str, seq):
        cur = self.cursor()
        cur.executemany(sql, seq)
        return cur

    def executescript(self, script: str):
        cur = self.cursor()
        cur.executescript(script)
        return cur

    def commit(self):
        try:
            self._raw.commit()
        except Exception:
            pass

    def rollback(self):
        try:
            self._raw.rollback()
        except Exception:
            pass

    def close(self):
        try:
            self._raw.close()
        except Exception:
            pass

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if exc_type is None:
            self.commit()
        else:
            self.rollback()
        self.close()


def get_pg_connection() -> PgConnection:
    """Открывает новое подключение к Supabase."""
    dsn = _get_dsn()
    if not dsn:
        raise RuntimeError("SUPABASE_DB_URL не задан. Настройте в .streamlit/secrets.toml или переменной окружения.")
    return PgConnection(dsn)
