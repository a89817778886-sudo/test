# -*- coding: utf-8 -*-
"""crm_inn.py — проверка ИНН/контрагента через открытый сервис ЕГРЮЛ ФНС (без API-ключа)."""

from __future__ import annotations

import requests
import time
from typing import Optional, Dict, Any

EGRUL_SEARCH_URL = "https://egrul.nalog.ru/"
EGRUL_QUERY_URL = "https://egrul.nalog.ru/search-result/"
TIMEOUT = 8


def lookup_inn(inn: str) -> Optional[Dict[str, Any]]:
    inn = (inn or "").strip()
    if not inn or not inn.isdigit():
        return None
    try:
        session = requests.Session()
        resp = session.post(EGRUL_SEARCH_URL, data={"query": inn}, timeout=TIMEOUT)
        resp.raise_for_status()
        data = resp.json()
        search_id = data.get("t")
        if not search_id:
            return None
        for _ in range(5):
            time.sleep(0.7)
            result = session.get(f"{EGRUL_QUERY_URL}{search_id}/", timeout=TIMEOUT)
            result.raise_for_status()
            rows = result.json().get("rows", [])
            if rows:
                row = rows[0]
                return {
                    "name_full": row.get("n") or row.get("c") or "",
                    "name_short": row.get("c") or row.get("n") or "",
                    "address": row.get("ad") or "",
                    "ogrn": row.get("o") or "",
                    "status": "действующая" if not row.get("l") else "ликвидирована",
                }
        return None
    except Exception:
        return None


def check_counterparty_risk(inn: str) -> str:
    info = lookup_inn(inn)
    if not info:
        return "Не удалось проверить"
    if info["status"] == "ликвидирована":
        return "⚠️ Организация ликвидирована"
    return "✅ Действующая организация"
