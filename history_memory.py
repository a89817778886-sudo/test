# -*- coding: utf-8 -*-
"""
history_memory.py — простая память ранее введённых значений.

Хранит списки уникальных значений и «карточки товаров» (name → code, price)
в файле history_memory.json рядом с приложением.

Используется для автодополнения полей, заполняемых вручную:
- ручные позиции спецификации (наименование + код + цена);
- реквизиты покупателя (наименование, ИНН);
- реквизиты для внешнего договора (весь блок по ИНН).
"""

from __future__ import annotations
import json
import os
from pathlib import Path
from typing import Any, Dict, List, Optional

_HERE = Path(__file__).parent
HISTORY_PATH = _HERE / "history_memory.json"

_CACHE: Optional[Dict[str, Any]] = None


def _load() -> Dict[str, Any]:
    global _CACHE
    if _CACHE is not None:
        return _CACHE
    if HISTORY_PATH.exists():
        try:
            with open(HISTORY_PATH, encoding="utf-8") as f:
                _CACHE = json.load(f)
        except Exception:
            _CACHE = {}
    else:
        _CACHE = {}
    return _CACHE


def _save() -> None:
    if _CACHE is None:
        return
    try:
        with open(HISTORY_PATH, "w", encoding="utf-8") as f:
            json.dump(_CACHE, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def get_list(key: str) -> List[str]:
    """Вернуть список значений (в порядке от последнего к первому)."""
    data = _load()
    return list(data.get(key, []))


def add_value(key: str, value: str, max_items: int = 200) -> None:
    """Добавить значение в историю (уникально, последнее — вверху)."""
    if not value or not str(value).strip():
        return
    value = str(value).strip()
    data = _load()
    lst = data.get(key, [])
    if value in lst:
        lst.remove(value)
    lst.insert(0, value)
    if len(lst) > max_items:
        lst = lst[:max_items]
    data[key] = lst
    _save()


def remove_value(key: str, value: str) -> None:
    data = _load()
    lst = data.get(key, [])
    if value in lst:
        lst.remove(value)
        data[key] = lst
        _save()


def get_dict(key: str) -> Dict[str, Any]:
    """Вернуть словарь-справочник (например, name → {code, price})."""
    data = _load()
    return dict(data.get(key, {}))


def set_dict_entry(key: str, entry_key: str, entry_value: Any) -> None:
    """Записать одну запись в словарь-справочник."""
    if not entry_key:
        return
    data = _load()
    d = data.get(key, {})
    if not isinstance(d, dict):
        d = {}
    d[str(entry_key).strip()] = entry_value
    data[key] = d
    _save()


def remove_dict_entry(key: str, entry_key: str) -> None:
    data = _load()
    d = data.get(key, {})
    if isinstance(d, dict) and entry_key in d:
        del d[entry_key]
        data[key] = d
        _save()


# --- Специализированные помощники для товаров ---

def remember_product(name: str, code: str = "", price: float = 0.0,
                     unit: str = "шт") -> None:
    """Запомнить товар: связка наименование ↔ код + последняя цена."""
    if not name or not str(name).strip():
        return
    add_value("hist_product_names", name)
    if code and str(code).strip():
        add_value("hist_product_codes", code)
    set_dict_entry(
        "hist_products_map",
        str(name).strip(),
        {"code": str(code or "").strip(),
         "price": float(price or 0),
         "unit": str(unit or "шт").strip()},
    )


def get_product(name: str) -> Optional[Dict[str, Any]]:
    """Вернуть карточку товара по наименованию (если раньше встречался)."""
    data = get_dict("hist_products_map")
    return data.get(str(name).strip())


def get_product_names() -> List[str]:
    return get_list("hist_product_names")


def get_product_codes() -> List[str]:
    return get_list("hist_product_codes")


# --- Специализированные помощники для покупателей ---

def remember_buyer_short(name: str) -> None:
    add_value("hist_buyer_names", name)


def get_buyer_names() -> List[str]:
    return get_list("hist_buyer_names")


def remember_ext_buyer(inn: str, data_block: Dict[str, Any]) -> None:
    """Полный блок реквизитов внешнего договора — по ИНН."""
    if not inn or not str(inn).strip():
        return
    set_dict_entry("hist_ext_buyers_by_inn", str(inn).strip(), data_block)
    if data_block.get("name_short"):
        add_value("hist_buyer_names", data_block["name_short"])


def get_ext_buyer_by_inn(inn: str) -> Optional[Dict[str, Any]]:
    return get_dict("hist_ext_buyers_by_inn").get(str(inn).strip())


def get_ext_buyers() -> Dict[str, Any]:
    return get_dict("hist_ext_buyers_by_inn")
