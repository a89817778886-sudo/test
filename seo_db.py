"""
БД SEO-модуля.

Использует shared connection из crm_db._get_conn() — работает с SQLite локально
и с PostgreSQL (Supabase) в проде. Не хардкодит sqlite3.
"""
from __future__ import annotations
import datetime as dt
import json
from dataclasses import dataclass
from typing import Optional, List, Dict, Any


def _get_conn():
    """Единая точка подключения — переиспользуем инфру crm_db."""
    from crm_db import get_conn as _shared
    return _shared()


def _is_postgres(conn) -> bool:
    """Определить тип БД по классу соединения."""
    try:
        import psycopg2.extensions
        return isinstance(conn, psycopg2.extensions.connection)
    except Exception:
        return False


# ==================== ИНИЦИАЛИЗАЦИЯ СХЕМЫ ====================

SCHEMA_SQLITE = """
CREATE TABLE IF NOT EXISTS seo_leads (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source TEXT,
    name TEXT, phone TEXT, email TEXT, company TEXT,
    site_url TEXT, comment TEXT,
    utm_source TEXT, utm_medium TEXT, utm_campaign TEXT,
    utm_term TEXT, utm_content TEXT,
    ip TEXT, user_agent TEXT,
    referrer TEXT,
    created_at TEXT NOT NULL,
    is_processed INTEGER NOT NULL DEFAULT 0,
    customer_id INTEGER,
    project_id INTEGER
);

CREATE TABLE IF NOT EXISTS seo_projects (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_id INTEGER,
    site_url TEXT NOT NULL,
    domain TEXT,
    region TEXT,
    service_type TEXT,
    monthly_fee REAL NOT NULL DEFAULT 0,
    status TEXT NOT NULL DEFAULT 'new',
    manager_id INTEGER,
    start_date TEXT,
    end_date TEXT,
    notes TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS seo_tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    title TEXT NOT NULL,
    description TEXT,
    priority TEXT NOT NULL DEFAULT 'medium',
    due_at TEXT,
    status TEXT NOT NULL DEFAULT 'open',
    assigned_to INTEGER,
    created_at TEXT NOT NULL,
    updated_at TEXT
);

CREATE TABLE IF NOT EXISTS seo_reports (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    report_period TEXT NOT NULL,
    visibility_score REAL NOT NULL DEFAULT 0,
    indexed_pages INTEGER NOT NULL DEFAULT 0,
    ai_mentions INTEGER NOT NULL DEFAULT 0,
    leads_count INTEGER NOT NULL DEFAULT 0,
    conversions INTEGER NOT NULL DEFAULT 0,
    notes TEXT,
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS seo_keywords (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    project_id INTEGER NOT NULL,
    keyword TEXT NOT NULL,
    target_page TEXT,
    frequency INTEGER DEFAULT 0,
    priority TEXT DEFAULT 'medium',
    created_at TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS seo_positions_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    keyword_id INTEGER NOT NULL,
    position INTEGER,
    search_engine TEXT DEFAULT 'yandex',
    check_date TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS seo_integrations (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    api_url TEXT,
    api_key_masked TEXT,
    is_enabled INTEGER NOT NULL DEFAULT 0,
    config_json TEXT,
    created_at TEXT NOT NULL,
    updated_at TEXT NOT NULL
);
"""

SCHEMA_PG = SCHEMA_SQLITE.replace(
    "INTEGER PRIMARY KEY AUTOINCREMENT", "SERIAL PRIMARY KEY"
).replace(" REAL ", " DOUBLE PRECISION ")


def init_seo_db() -> None:
    """Создаёт таблицы если их нет. Идемпотентно."""
    conn = _get_conn()
    try:
        cur = conn.cursor()
        if _is_postgres(conn):
            for stmt in SCHEMA_PG.split(";"):
                if stmt.strip():
                    cur.execute(stmt)
        else:
            cur.executescript(SCHEMA_SQLITE)
        conn.commit()
    finally:
        conn.close()


# ==================== ХЕЛПЕРЫ ====================

def _exec(cur, sql_pg: str, sql_sqlite: str, params: tuple = ()):
    """Универсальный execute — пробуем PG-синтаксис, при ошибке SQLite."""
    try:
        cur.execute(sql_pg, params)
    except Exception:
        cur.execute(sql_sqlite, params)


def _row_to_dict(row, cur) -> dict:
    """Конвертирует row в dict, работает и для sqlite Row и для psycopg2."""
    if row is None:
        return {}
    if hasattr(row, "keys"):
        return dict(row)
    cols = [c[0] for c in cur.description] if cur.description else []
    return dict(zip(cols, row))


# ==================== ЛИДЫ ====================

@dataclass
class SEOLead:
    source: str = 'site'
    name: str = ''
    phone: str = ''
    email: str = ''
    company: str = ''
    site_url: str = ''
    comment: str = ''
    utm_source: str = ''
    utm_medium: str = ''
    utm_campaign: str = ''
    utm_term: str = ''
    utm_content: str = ''
    ip: str = ''
    user_agent: str = ''
    referrer: str = ''


def create_lead(lead: SEOLead) -> int:
    init_seo_db()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        now = dt.datetime.now().isoformat(timespec='seconds')
        if _is_postgres(conn):
            cur.execute("""INSERT INTO seo_leads
                (source, name, phone, email, company, site_url, comment,
                 utm_source, utm_medium, utm_campaign, utm_term, utm_content,
                 ip, user_agent, referrer, created_at, is_processed)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,0)
                RETURNING id""",
                (lead.source, lead.name, lead.phone, lead.email, lead.company,
                 lead.site_url, lead.comment, lead.utm_source, lead.utm_medium,
                 lead.utm_campaign, lead.utm_term, lead.utm_content,
                 lead.ip, lead.user_agent, lead.referrer, now))
            new_id = cur.fetchone()[0]
        else:
            cur.execute("""INSERT INTO seo_leads
                (source, name, phone, email, company, site_url, comment,
                 utm_source, utm_medium, utm_campaign, utm_term, utm_content,
                 ip, user_agent, referrer, created_at, is_processed)
                VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,0)""",
                (lead.source, lead.name, lead.phone, lead.email, lead.company,
                 lead.site_url, lead.comment, lead.utm_source, lead.utm_medium,
                 lead.utm_campaign, lead.utm_term, lead.utm_content,
                 lead.ip, lead.user_agent, lead.referrer, now))
            new_id = cur.lastrowid
        conn.commit()
        return int(new_id)
    finally:
        conn.close()


def list_leads(limit: int = 500, only_unprocessed: bool = False) -> List[Dict[str, Any]]:
    init_seo_db()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        if only_unprocessed:
            try:
                cur.execute("SELECT * FROM seo_leads WHERE is_processed = 0 ORDER BY created_at DESC LIMIT %s", (limit,))
            except Exception:
                cur.execute("SELECT * FROM seo_leads WHERE is_processed = 0 ORDER BY created_at DESC LIMIT ?", (limit,))
        else:
            try:
                cur.execute("SELECT * FROM seo_leads ORDER BY created_at DESC LIMIT %s", (limit,))
            except Exception:
                cur.execute("SELECT * FROM seo_leads ORDER BY created_at DESC LIMIT ?", (limit,))
        rows = cur.fetchall()
        return [_row_to_dict(r, cur) for r in rows]
    finally:
        conn.close()


def get_lead(lead_id: int) -> Optional[Dict[str, Any]]:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        try:
            cur.execute("SELECT * FROM seo_leads WHERE id = %s", (lead_id,))
        except Exception:
            cur.execute("SELECT * FROM seo_leads WHERE id = ?", (lead_id,))
        row = cur.fetchone()
        return _row_to_dict(row, cur) if row else None
    finally:
        conn.close()


def mark_lead_processed(lead_id: int, customer_id: int = None, project_id: int = None) -> None:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        try:
            cur.execute("""UPDATE seo_leads SET is_processed = 1,
                         customer_id = COALESCE(%s, customer_id),
                         project_id = COALESCE(%s, project_id)
                         WHERE id = %s""",
                        (customer_id, project_id, lead_id))
        except Exception:
            cur.execute("""UPDATE seo_leads SET is_processed = 1,
                         customer_id = COALESCE(?, customer_id),
                         project_id = COALESCE(?, project_id)
                         WHERE id = ?""",
                        (customer_id, project_id, lead_id))
        conn.commit()
    finally:
        conn.close()


def delete_lead(lead_id: int) -> None:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        try:
            cur.execute("DELETE FROM seo_leads WHERE id = %s", (lead_id,))
        except Exception:
            cur.execute("DELETE FROM seo_leads WHERE id = ?", (lead_id,))
        conn.commit()
    finally:
        conn.close()


# ==================== ПРОЕКТЫ ====================

def create_project(customer_id: Optional[int], site_url: str, domain: str = '',
                   region: str = '', service_type: str = 'SEO автопилот',
                   monthly_fee: float = 0, status: str = 'new',
                   manager_id: Optional[int] = None, notes: str = '') -> int:
    init_seo_db()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        now = dt.datetime.now().isoformat(timespec='seconds')
        if _is_postgres(conn):
            cur.execute("""INSERT INTO seo_projects
                (customer_id, site_url, domain, region, service_type, monthly_fee,
                 status, manager_id, notes, created_at, updated_at)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                (customer_id, site_url, domain, region, service_type,
                 float(monthly_fee), status, manager_id, notes, now, now))
            new_id = cur.fetchone()[0]
        else:
            cur.execute("""INSERT INTO seo_projects
                (customer_id, site_url, domain, region, service_type, monthly_fee,
                 status, manager_id, notes, created_at, updated_at)
                VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
                (customer_id, site_url, domain, region, service_type,
                 float(monthly_fee), status, manager_id, notes, now, now))
            new_id = cur.lastrowid
        conn.commit()
        return int(new_id)
    finally:
        conn.close()


def list_projects(customer_id: int = None) -> List[Dict[str, Any]]:
    init_seo_db()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        if customer_id is not None:
            try:
                cur.execute("""SELECT p.*, c.name_short AS customer_name
                    FROM seo_projects p LEFT JOIN customers c ON c.id = p.customer_id
                    WHERE p.customer_id = %s ORDER BY p.updated_at DESC""",
                    (customer_id,))
            except Exception:
                cur.execute("""SELECT p.*, c.name_short AS customer_name
                    FROM seo_projects p LEFT JOIN customers c ON c.id = p.customer_id
                    WHERE p.customer_id = ? ORDER BY p.updated_at DESC""",
                    (customer_id,))
        else:
            cur.execute("""SELECT p.*, c.name_short AS customer_name
                FROM seo_projects p LEFT JOIN customers c ON c.id = p.customer_id
                ORDER BY p.updated_at DESC""")
        rows = cur.fetchall()
        return [_row_to_dict(r, cur) for r in rows]
    finally:
        conn.close()


def get_project(project_id: int) -> Optional[Dict[str, Any]]:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        try:
            cur.execute("""SELECT p.*, c.name_short AS customer_name, c.inn AS customer_inn
                FROM seo_projects p LEFT JOIN customers c ON c.id = p.customer_id
                WHERE p.id = %s""", (project_id,))
        except Exception:
            cur.execute("""SELECT p.*, c.name_short AS customer_name, c.inn AS customer_inn
                FROM seo_projects p LEFT JOIN customers c ON c.id = p.customer_id
                WHERE p.id = ?""", (project_id,))
        row = cur.fetchone()
        return _row_to_dict(row, cur) if row else None
    finally:
        conn.close()


def update_project(project_id: int, **fields) -> None:
    if not fields: return
    conn = _get_conn()
    try:
        cur = conn.cursor()
        now = dt.datetime.now().isoformat(timespec='seconds')
        fields['updated_at'] = now
        # Строим SET-часть
        set_pg = ", ".join(f"{k} = %s" for k in fields)
        set_sq = ", ".join(f"{k} = ?" for k in fields)
        params = tuple(fields.values()) + (project_id,)
        try:
            cur.execute(f"UPDATE seo_projects SET {set_pg} WHERE id = %s", params)
        except Exception:
            cur.execute(f"UPDATE seo_projects SET {set_sq} WHERE id = ?", params)
        conn.commit()
    finally:
        conn.close()


def delete_project(project_id: int) -> None:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        # Каскадно удаляем связанные задачи и отчёты
        for tbl in ("seo_tasks", "seo_reports", "seo_keywords"):
            try:
                cur.execute(f"DELETE FROM {tbl} WHERE project_id = %s", (project_id,))
            except Exception:
                cur.execute(f"DELETE FROM {tbl} WHERE project_id = ?", (project_id,))
        try:
            cur.execute("DELETE FROM seo_projects WHERE id = %s", (project_id,))
        except Exception:
            cur.execute("DELETE FROM seo_projects WHERE id = ?", (project_id,))
        conn.commit()
    finally:
        conn.close()


# ==================== ЗАДАЧИ ====================

def create_task(project_id: int, title: str, description: str = '',
                priority: str = 'medium', due_at: Optional[str] = None,
                status: str = 'open', assigned_to: Optional[int] = None) -> int:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        now = dt.datetime.now().isoformat(timespec='seconds')
        if _is_postgres(conn):
            cur.execute("""INSERT INTO seo_tasks
                (project_id, title, description, priority, due_at, status, assigned_to, created_at)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                (project_id, title, description, priority, due_at, status, assigned_to, now))
            new_id = cur.fetchone()[0]
        else:
            cur.execute("""INSERT INTO seo_tasks
                (project_id, title, description, priority, due_at, status, assigned_to, created_at)
                VALUES (?,?,?,?,?,?,?,?)""",
                (project_id, title, description, priority, due_at, status, assigned_to, now))
            new_id = cur.lastrowid
        conn.commit()
        return int(new_id)
    finally:
        conn.close()


def list_tasks(project_id: int = None, status: str = None) -> List[Dict[str, Any]]:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        base = """SELECT t.*, p.domain, p.site_url, p.customer_id
                  FROM seo_tasks t LEFT JOIN seo_projects p ON p.id = t.project_id
                  WHERE 1=1"""
        params = []
        if project_id:
            base += " AND t.project_id = ?"
            params.append(project_id)
        if status:
            base += " AND t.status = ?"
            params.append(status)
        base += " ORDER BY t.created_at DESC"
        try:
            cur.execute(base.replace("?", "%s"), tuple(params))
        except Exception:
            cur.execute(base, tuple(params))
        rows = cur.fetchall()
        return [_row_to_dict(r, cur) for r in rows]
    finally:
        conn.close()


def update_task_status(task_id: int, status: str) -> None:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        now = dt.datetime.now().isoformat(timespec='seconds')
        try:
            cur.execute("UPDATE seo_tasks SET status = %s, updated_at = %s WHERE id = %s",
                        (status, now, task_id))
        except Exception:
            cur.execute("UPDATE seo_tasks SET status = ?, updated_at = ? WHERE id = ?",
                        (status, now, task_id))
        conn.commit()
    finally:
        conn.close()


def delete_task(task_id: int) -> None:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        try:
            cur.execute("DELETE FROM seo_tasks WHERE id = %s", (task_id,))
        except Exception:
            cur.execute("DELETE FROM seo_tasks WHERE id = ?", (task_id,))
        conn.commit()
    finally:
        conn.close()


# ==================== ОТЧЁТЫ ====================

def create_report(project_id: int, report_period: str, visibility_score: float = 0,
                  indexed_pages: int = 0, ai_mentions: int = 0, leads_count: int = 0,
                  conversions: int = 0, notes: str = '') -> int:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        now = dt.datetime.now().isoformat(timespec='seconds')
        if _is_postgres(conn):
            cur.execute("""INSERT INTO seo_reports
                (project_id, report_period, visibility_score, indexed_pages,
                 ai_mentions, leads_count, conversions, notes, created_at)
                VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                (project_id, report_period, float(visibility_score), int(indexed_pages),
                 int(ai_mentions), int(leads_count), int(conversions), notes, now))
            new_id = cur.fetchone()[0]
        else:
            cur.execute("""INSERT INTO seo_reports
                (project_id, report_period, visibility_score, indexed_pages,
                 ai_mentions, leads_count, conversions, notes, created_at)
                VALUES (?,?,?,?,?,?,?,?,?)""",
                (project_id, report_period, float(visibility_score), int(indexed_pages),
                 int(ai_mentions), int(leads_count), int(conversions), notes, now))
            new_id = cur.lastrowid
        conn.commit()
        return int(new_id)
    finally:
        conn.close()


def list_reports(project_id: int = None) -> List[Dict[str, Any]]:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        if project_id:
            try:
                cur.execute("""SELECT r.*, p.domain, p.site_url
                    FROM seo_reports r LEFT JOIN seo_projects p ON p.id = r.project_id
                    WHERE r.project_id = %s ORDER BY r.report_period DESC""",
                    (project_id,))
            except Exception:
                cur.execute("""SELECT r.*, p.domain, p.site_url
                    FROM seo_reports r LEFT JOIN seo_projects p ON p.id = r.project_id
                    WHERE r.project_id = ? ORDER BY r.report_period DESC""",
                    (project_id,))
        else:
            cur.execute("""SELECT r.*, p.domain, p.site_url
                FROM seo_reports r LEFT JOIN seo_projects p ON p.id = r.project_id
                ORDER BY r.created_at DESC""")
        rows = cur.fetchall()
        return [_row_to_dict(r, cur) for r in rows]
    finally:
        conn.close()


def delete_report(report_id: int) -> None:
    conn = _get_conn()
    try:
        cur = conn.cursor()
        try:
            cur.execute("DELETE FROM seo_reports WHERE id = %s", (report_id,))
        except Exception:
            cur.execute("DELETE FROM seo_reports WHERE id = ?", (report_id,))
        conn.commit()
    finally:
        conn.close()


# ==================== ИНТЕГРАЦИИ ====================

def save_integration(name: str, api_url: str = '', api_key_masked: str = '',
                     is_enabled: bool = False, config: dict = None) -> int:
    init_seo_db()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        now = dt.datetime.now().isoformat(timespec='seconds')
        cfg_json = json.dumps(config or {}, ensure_ascii=False)
        try:
            cur.execute("SELECT id FROM seo_integrations WHERE name = %s", (name,))
        except Exception:
            cur.execute("SELECT id FROM seo_integrations WHERE name = ?", (name,))
        row = cur.fetchone()
        if row:
            eid = row[0] if not hasattr(row, "keys") else row["id"]
            try:
                cur.execute("""UPDATE seo_integrations SET api_url=%s, api_key_masked=%s,
                    is_enabled=%s, config_json=%s, updated_at=%s WHERE id=%s""",
                    (api_url, api_key_masked, 1 if is_enabled else 0, cfg_json, now, eid))
            except Exception:
                cur.execute("""UPDATE seo_integrations SET api_url=?, api_key_masked=?,
                    is_enabled=?, config_json=?, updated_at=? WHERE id=?""",
                    (api_url, api_key_masked, 1 if is_enabled else 0, cfg_json, now, eid))
            conn.commit()
            return int(eid)
        else:
            if _is_postgres(conn):
                cur.execute("""INSERT INTO seo_integrations
                    (name, api_url, api_key_masked, is_enabled, config_json, created_at, updated_at)
                    VALUES (%s,%s,%s,%s,%s,%s,%s) RETURNING id""",
                    (name, api_url, api_key_masked, 1 if is_enabled else 0, cfg_json, now, now))
                new_id = cur.fetchone()[0]
            else:
                cur.execute("""INSERT INTO seo_integrations
                    (name, api_url, api_key_masked, is_enabled, config_json, created_at, updated_at)
                    VALUES (?,?,?,?,?,?,?)""",
                    (name, api_url, api_key_masked, 1 if is_enabled else 0, cfg_json, now, now))
                new_id = cur.lastrowid
            conn.commit()
            return int(new_id)
    finally:
        conn.close()


def list_integrations() -> List[Dict[str, Any]]:
    init_seo_db()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        cur.execute("SELECT * FROM seo_integrations ORDER BY name")
        rows = cur.fetchall()
        return [_row_to_dict(r, cur) for r in rows]
    finally:
        conn.close()


# ==================== АНАЛИТИКА / DASHBOARD ====================

def dashboard_metrics() -> Dict[str, Any]:
    init_seo_db()
    conn = _get_conn()
    try:
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM seo_projects")
        projects = cur.fetchone()[0] or 0

        cur.execute("SELECT COUNT(*) FROM seo_projects WHERE status = 'active'")
        active_projects = cur.fetchone()[0] or 0

        cur.execute("SELECT COUNT(*) FROM seo_leads WHERE is_processed = 0")
        active_leads = cur.fetchone()[0] or 0

        month = dt.datetime.now().strftime('%Y-%m')
        try:
            cur.execute("SELECT COUNT(*) FROM seo_leads WHERE substr(created_at,1,7) = %s", (month,))
        except Exception:
            cur.execute("SELECT COUNT(*) FROM seo_leads WHERE substr(created_at,1,7) = ?", (month,))
        leads_month = cur.fetchone()[0] or 0

        try:
            cur.execute("""SELECT COALESCE(SUM(monthly_fee),0) FROM seo_projects
                           WHERE status IN ('active','new','in_work')""")
        except Exception: pass
        revenue = cur.fetchone()[0] or 0

        # Распределение по статусам
        cur.execute("SELECT status, COUNT(*) FROM seo_projects GROUP BY status")
        status_rows = cur.fetchall()
        status_dist = {}
        for r in status_rows:
            if hasattr(r, "keys"):
                status_dist[r[0] or "unknown"] = r[1] or 0
            else:
                status_dist[r[0] or "unknown"] = r[1] or 0

        return {
            'projects': int(projects),
            'active_projects': int(active_projects),
            'active_leads': int(active_leads),
            'leads_month': int(leads_month),
            'forecast_revenue': float(revenue),
            'status_dist': status_dist,
        }
    finally:
        conn.close()
