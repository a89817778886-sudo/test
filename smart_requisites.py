# -*- coding: utf-8 -*-
"""
smart_requisites.py — «умное» досконально-точное распознавание реквизитов.

Двухслойная схема:
  1. OCR-слой: любой файл → текст
     • DOCX/DOC → python-docx + XML fallback
     • PDF с текстом → pdfplumber / PyMuPDF
     • PDF-скан или картинка (JPG/PNG/скриншот) → Tesseract OCR (rus + eng)
  2. ИИ-слой: текст → структурированный JSON с реквизитами через GPT-4
     Использует OPENAI_API_KEY из secrets/переменных окружения.
     Если ИИ недоступен — fallback на регексы (external_kp_parser).

Возвращает единый формат dict, совместимый с buyer_data / ext_buyer_data.
"""
from __future__ import annotations

import io
import json
import os
from pathlib import Path
from typing import Any, Dict, Optional


# ============================================================
#  1. OCR — извлечение текста из любого файла
# ============================================================

def _extract_text_smart(file_bytes: bytes, filename: str) -> str:
    """Универсальный экстрактор текста. Пробует все методы по очереди."""
    ext = Path(filename).suffix.lower()
    text = ""

    if ext == ".docx":
        text = _from_docx(file_bytes)
    elif ext == ".doc":
        text = _from_doc(file_bytes)
    elif ext == ".pdf":
        text = _from_pdf_text(file_bytes)
        # Критерий «поломанного» PDF: мало текста ЛИБО мало кириллицы (как в EN-TD)
        _cyr = sum(1 for c in text if "А" <= c <= "я")
        if len(text.strip()) < 50 or _cyr < 30:
            # PDF-скан или битая кодировка — прогоняем через OCR
            ocr_text = _from_pdf_ocr(file_bytes)
            if len(ocr_text.strip()) > len(text.strip()):
                text = ocr_text
    elif ext in (".jpg", ".jpeg", ".png", ".webp", ".bmp", ".tiff", ".tif"):
        text = _from_image_ocr(file_bytes)
    elif ext == ".txt":
        try:
            text = file_bytes.decode("utf-8")
        except Exception:
            text = file_bytes.decode("cp1251", errors="ignore")

    return text


def _from_docx(file_bytes: bytes) -> str:
    """Извлечение из docx: параграфы + таблицы + textbox через XML."""
    try:
        from external_kp_parser import _extract_text_from_docx
        return _extract_text_from_docx(file_bytes)
    except Exception:
        return ""


def _from_doc(file_bytes: bytes) -> str:
    """Извлечение из старого .doc (Word 97-2003).
    Использует antiword или textract если доступны, иначе OCR через LibreOffice."""
    import tempfile
    import subprocess

    # Пробуем antiword
    try:
        with tempfile.NamedTemporaryFile(suffix=".doc", delete=False) as tmp:
            tmp.write(file_bytes)
            tmp_path = tmp.name
        try:
            result = subprocess.run(
                ["antiword", tmp_path],
                capture_output=True, timeout=30
            )
            if result.returncode == 0:
                return result.stdout.decode("utf-8", errors="ignore")
        except (FileNotFoundError, subprocess.TimeoutExpired):
            pass
        finally:
            try:
                os.unlink(tmp_path)
            except Exception:
                pass
    except Exception:
        pass

    # Fallback: конвертим через libreoffice → docx → извлекаем
    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            src = Path(tmpdir) / "src.doc"
            src.write_bytes(file_bytes)
            subprocess.run(
                ["libreoffice", "--headless", "--convert-to", "docx",
                 "--outdir", tmpdir, str(src)],
                capture_output=True, timeout=60
            )
            docx_path = Path(tmpdir) / "src.docx"
            if docx_path.exists():
                return _from_docx(docx_path.read_bytes())
    except Exception:
        pass

    return ""


def _from_pdf_text(file_bytes: bytes) -> str:
    """PDF с текстовым слоем."""
    try:
        from external_kp_parser import _extract_text_from_pdf
        return _extract_text_from_pdf(file_bytes)
    except Exception:
        return ""


def _from_pdf_ocr(file_bytes: bytes) -> str:
    """PDF-скан: рендерим страницы в изображения и распознаём OCR."""
    try:
        import fitz  # PyMuPDF
        import pytesseract
        from PIL import Image

        doc = fitz.open(stream=file_bytes, filetype="pdf")
        parts = []
        for page in doc:
            # 300 DPI для качественного распознавания
            pix = page.get_pixmap(dpi=300)
            img = Image.open(io.BytesIO(pix.tobytes("png")))
            page_text = pytesseract.image_to_string(img, lang="rus+eng")
            if page_text.strip():
                parts.append(page_text)
        doc.close()
        return "\n".join(parts)
    except Exception as e:
        return f""


def _from_image_ocr(file_bytes: bytes) -> str:
    """Распознавание текста с картинки (jpg/png/скриншот)."""
    try:
        import pytesseract
        from PIL import Image
        img = Image.open(io.BytesIO(file_bytes))
        # Если картинка большая — уменьшаем не сильно (OCR любит крупный текст)
        if img.width > 3000:
            ratio = 3000 / img.width
            img = img.resize((3000, int(img.height * ratio)))
        return pytesseract.image_to_string(img, lang="rus+eng")
    except Exception:
        return ""


# ============================================================
#  2. ИИ-слой — GPT разбирает текст в структурированный JSON
# ============================================================

_SYSTEM_PROMPT = """Ты — эксперт по извлечению реквизитов российских юридических лиц из документов.

Тебе дадут текст (карточка организации, реквизиты из письма, договор и т.п.), возможно с ошибками OCR (пропущенные символы, лишние пробелы, склеенные слова).

Твоя задача — извлечь все поля и вернуть СТРОГО валидный JSON без markdown-обёртки, без комментариев, ровно с такими ключами:

{
  "short": "краткое наименование (напр. ООО «Ромашка»)",
  "full": "полное наименование (напр. Общество с ограниченной ответственностью «Ромашка»)",
  "inn": "10 или 12 цифр",
  "kpp": "9 цифр",
  "ogrn": "13 или 15 цифр (ОГРН/ОГРНИП)",
  "address": "полный юридический адрес одной строкой",
  "post_address": "почтовый адрес если отличается, иначе пусто",
  "bank": "полное название банка (напр. Филиал «Санкт-Петербургский» АО «АЛЬФА-БАНК»)",
  "bik": "9 цифр",
  "rs": "20 цифр расчётного счёта",
  "ks": "20 цифр корреспондентского счёта",
  "phone": "телефон в формате +7 (999) 111-22-33",
  "email": "email",
  "director_position": "должность руководителя (Генеральный директор / Директор / ИП)",
  "director_fio_short": "ФИО в сокращённой форме, напр. Иванов И.И.",
  "director_fio_gen": "ФИО в родительном падеже, напр. Иванова Ивана Ивановича",
  "basis": "основание действий (обычно 'Устава' или 'Свидетельства о регистрации ИП')"
}

Правила:
- Если поле не найдено — верни пустую строку "".
- НЕ придумывай данные, которых нет в тексте.
- Директор в родительном падеже: правильно склоняй фамилию, имя и отчество. «Иванов Иван Иванович» → «Иванова Ивана Ивановича». «Сурина Анастасия Витальевна» → «Суриной Анастасии Витальевны». Женские фамилии на «-ина/-ова» → «-иной/-овой».
- Банк — полное название вместе с формой (Филиал «...» АО «...», ПАО, АО, и т.п.). Не обрывай на первом слове.
- Адрес — полный, вместе с индексом, городом, улицей, домом, помещением.
- Только JSON, ничего вокруг."""


def _extract_with_ai(text: str) -> Optional[Dict[str, Any]]:
    """Отправляем текст в GPT-4 и получаем JSON. None если API недоступен."""
    api_key = _get_openai_key()
    if not api_key:
        return None

    try:
        from openai import OpenAI
        client = OpenAI(api_key=api_key)

        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": _SYSTEM_PROMPT},
                {"role": "user", "content": text[:15000]},
            ],
            temperature=0,
            response_format={"type": "json_object"},
            max_tokens=1500,
        )
        content = response.choices[0].message.content
        return json.loads(content)
    except Exception as e:
        print(f"AI extraction failed: {e}")
        return None


def _get_openai_key() -> str:
    """Ищем ключ в secrets Streamlit, в env, или в файле .openai_key."""
    # 1. Streamlit secrets
    try:
        import streamlit as st
        key = st.secrets.get("OPENAI_API_KEY", "")
        if key:
            return key
    except Exception:
        pass
    # 2. Переменная окружения
    key = os.environ.get("OPENAI_API_KEY", "")
    if key:
        return key
    # 3. Локальный файл
    try:
        p = Path(__file__).parent / ".openai_key"
        if p.exists():
            return p.read_text().strip()
    except Exception:
        pass
    return ""


# ============================================================
#  3. Fallback — старый парсер регексов
# ============================================================

def _extract_with_regex(text: str) -> Dict[str, Any]:
    """Fallback через регексы (external_kp_parser)."""
    try:
        from external_kp_parser import extract_requisites_from_text
        req = extract_requisites_from_text(text)
        return {
            "short": req.company_short or "",
            "full": req.company_full or req.company_short or "",
            "inn": req.inn or "",
            "kpp": req.kpp or "",
            "ogrn": req.ogrn or "",
            "address": req.address or "",
            "post_address": req.address or "",
            "bank": req.bank_name or "",
            "bik": req.bank_bik or "",
            "rs": req.bank_account or "",
            "ks": req.corr_account or "",
            "phone": req.phone or "",
            "email": req.email or "",
            "director_position": req.director_title or "Генеральный директор",
            "director_fio_short": req.director_short or "",
            "director_fio_gen": req.director_gen or "",
            "basis": "Устава",
        }
    except Exception:
        return {}


# ============================================================
#  4. Публичное API
# ============================================================

def smart_extract_from_file(file_bytes: bytes, filename: str) -> Dict[str, Any]:
    """Извлечь реквизиты из файла (docx/doc/pdf/jpg/png/txt).

    Пробует ИИ, при неудаче — регексы. Возвращает dict с ключами:
    short, full, inn, kpp, ogrn, address, post_address, bank, bik, rs, ks,
    phone, email, director_position, director_fio_short, director_fio_gen, basis.
    """
    text = _extract_text_smart(file_bytes, filename)
    if not text or not text.strip():
        return {}
    return smart_extract_from_text(text)


def smart_extract_from_text(text: str) -> Dict[str, Any]:
    """Извлечь реквизиты из готового текста (copy-paste из письма/сайта)."""
    if not text or not text.strip():
        return {}
    # Только регексы (ИИ выключен)
    return _extract_with_regex(text)


def has_ai_available() -> bool:
    """Есть ли доступ к OpenAI API."""
    return bool(_get_openai_key())


def get_extracted_text_preview(file_bytes: bytes, filename: str,
                                max_chars: int = 3000) -> str:
    """Показать первые 3000 символов извлечённого текста — для отладки."""
    text = _extract_text_smart(file_bytes, filename)
    return text[:max_chars]
