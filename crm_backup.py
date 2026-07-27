# -*- coding: utf-8 -*-
"""crm_backup.py — автоматический бэкап базы crm.db: ротация локальных копий + опциональный WebDAV Яндекс.Диск."""

from __future__ import annotations

import shutil
import datetime as dt
from pathlib import Path
from typing import List, Dict, Any

APP_DIR = Path(__file__).resolve().parent
DB_PATH = APP_DIR / "crm.db"
BACKUP_DIR = APP_DIR / "backups"
MAX_BACKUPS = 20


def create_local_backup() -> Path:
    BACKUP_DIR.mkdir(exist_ok=True)
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    dest = BACKUP_DIR / f"crm_{ts}.db"
    if DB_PATH.exists():
        shutil.copy2(DB_PATH, dest)
    _rotate_backups()
    return dest


def _rotate_backups() -> None:
    if not BACKUP_DIR.exists():
        return
    backups = sorted(BACKUP_DIR.glob("crm_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)
    for old in backups[MAX_BACKUPS:]:
        old.unlink(missing_ok=True)


def list_backups() -> List[Dict[str, Any]]:
    if not BACKUP_DIR.exists():
        return []
    backups = sorted(BACKUP_DIR.glob("crm_*.db"), key=lambda p: p.stat().st_mtime, reverse=True)
    return [{"name": p.name, "path": str(p), "size_kb": round(p.stat().st_size / 1024, 1),
             "created": dt.datetime.fromtimestamp(p.stat().st_mtime).strftime("%d.%m.%Y %H:%M")} for p in backups]


def upload_to_yandex_disk(local_path: Path) -> Dict[str, Any]:
    import requests
    try:
        import streamlit as st
        cfg = st.secrets.get("yadisk", {})
    except Exception:
        cfg = {}

    webdav_url = cfg.get("webdav_url", "")
    login = cfg.get("login", "")
    password = cfg.get("password", "")

    if not (webdav_url and login and password):
        return {"ok": False, "error": "Яндекс.Диск не настроен — заполните .streamlit/secrets.toml [yadisk]"}

    try:
        with open(local_path, "rb") as f:
            resp = requests.put(f"{webdav_url}/{local_path.name}", data=f, auth=(login, password), timeout=30)
        if resp.status_code in (200, 201, 204):
            return {"ok": True, "error": ""}
        return {"ok": False, "error": f"HTTP {resp.status_code}"}
    except Exception as e:
        return {"ok": False, "error": str(e)}
