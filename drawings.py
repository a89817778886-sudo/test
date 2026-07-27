"""Поиск габаритных чертежей крана / траверсы / аксессуаров.

Ищет файл в media/crane_drawings/ media/traverse_drawings/ media/accessory_drawings/
по имени. Возвращает Path или None.

Поддерживаемые расширения: .jpg .jpeg .png .pdf
"""
from pathlib import Path
from typing import Optional

BASE = Path(__file__).parent / "media"
CRANE_DIR = BASE / "crane_drawings"
TRAVERSE_DIR = BASE / "traverse_drawings"
ACCESSORY_DIR = BASE / "accessory_drawings"

_EXTS = (".jpg", ".jpeg", ".png", ".pdf")


def _find_file(directory: Path, basename: str) -> Optional[Path]:
    """Ищет файл basename.{ext} в directory (регистр не важен)."""
    if not directory.exists():
        return None
    basename_low = basename.lower()
    for f in directory.iterdir():
        if not f.is_file():
            continue
        stem = f.stem.lower()
        ext = f.suffix.lower()
        if ext not in _EXTS:
            continue
        if stem == basename_low:
            return f
    return None


def find_crane_drawing(series: str, capacity: int = 0, boom: float = 0,
                       height_to_arm: float = 0,
                       use_lllm: bool = False, lllm_code: str = "") -> Optional[Path]:
    """Ищет чертёж крана. Идёт от точного к общему:
    lks73m-1000-6-6 → lks73m-1000 → lks73m → lks73 → None.
    """
    # Определяем "ключ серии" — lks71/lks73/lks73m/lks77/lks78
    s = series.lower().replace("лкс", "lks")
    if use_lllm and lllm_code:
        s_key = "lks73m"
    else:
        s_key = s

    # Разбираем boom/height для LLL (могут быть в lllm_code типа "ЛКС73М.1000-6-6.LLL")
    if use_lllm and lllm_code:
        import re
        m = re.search(r"(\d+)-(\d+)-(\d+)", lllm_code)
        if m:
            capacity = int(m.group(1))
            boom = int(m.group(2))
            height_to_arm = int(m.group(3))

    cap = int(capacity or 0)
    b = int(boom or 0)
    h = int(height_to_arm or 0)

    # Приоритет поиска
    candidates = []
    if cap and b and h:
        candidates.append(f"{s_key}-{cap}-{b}-{h}")
    if cap:
        candidates.append(f"{s_key}-{cap}")
    candidates.append(s_key)
    # Fallback без "м" (для LLL: lks73m → lks73)
    if s_key.endswith("m"):
        candidates.append(s_key[:-1])

    for name in candidates:
        f = _find_file(CRANE_DIR, name)
        if f:
            return f
    return None


def find_traverse_drawing(model_code: str) -> Optional[Path]:
    """Ищет чертёж траверсы. model_code = «6P-500», «10P-900» и т.д."""
    if not model_code:
        return None
    return _find_file(TRAVERSE_DIR, model_code.lower())


def find_accessory_drawing(name_or_code: str) -> Optional[Path]:
    """Ищет чертёж аксессуара по нормализованному имени."""
    if not name_or_code:
        return None
    # Нормализуем: «Спиральный кабель» → «spiral-cable»
    key = name_or_code.lower().strip()
    # Часто встречающиеся сопоставления
    aliases = {
        "спиральный кабель": "spiral-cable",
        "спиральный кабель для vacutec": "spiral-cable",
        "опоры для хранения": "storage-supports",
        "опоры для хранения траверсы": "storage-supports",
        "наклонная ручка": "inclined-handle",
        "сменный аккумулятор": "battery",
    }
    for alias, target in aliases.items():
        if alias in key:
            f = _find_file(ACCESSORY_DIR, target)
            if f:
                return f
    # Прямой поиск по коду
    return _find_file(ACCESSORY_DIR, key.replace(" ", "-"))


def list_all_crane_drawings() -> list:
    """Возвращает список всех чертежей кранов (для отладки/UI)."""
    if not CRANE_DIR.exists():
        return []
    return sorted([f for f in CRANE_DIR.iterdir()
                   if f.is_file() and f.suffix.lower() in _EXTS])


def list_all_traverse_drawings() -> list:
    if not TRAVERSE_DIR.exists():
        return []
    return sorted([f for f in TRAVERSE_DIR.iterdir()
                   if f.is_file() and f.suffix.lower() in _EXTS])
