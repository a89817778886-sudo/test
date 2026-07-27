"""SEO модуль — вспомогательные утилиты."""
from urllib.parse import urlparse
import re


def extract_domain(url: str) -> str:
    """Вытащить домен из URL. 'https://www.example.com/page' → 'example.com'"""
    if not url:
        return ''
    raw = url.strip()
    if '://' not in raw:
        raw = 'https://' + raw
    try:
        return (urlparse(raw).netloc or '').replace('www.', '')
    except Exception:
        return ''


def validate_lead_payload(payload: dict) -> tuple[bool, str]:
    """Валидация входного пакета от сайта."""
    if not payload:
        return False, 'Пустой payload'
    if not (payload.get('phone') or payload.get('email')):
        return False, 'Нужен телефон или email'
    site = str(payload.get('site_url', '')).strip()
    if not site:
        return False, 'Нужен site_url'
    return True, ''


def normalize_phone(phone: str) -> str:
    """Нормализация телефона до +7XXXXXXXXXX."""
    if not phone:
        return ''
    digits = re.sub(r'[^\d]', '', phone)
    if len(digits) == 11 and digits.startswith('8'):
        digits = '7' + digits[1:]
    if len(digits) == 10:
        digits = '7' + digits
    if len(digits) == 11 and digits.startswith('7'):
        return '+' + digits
    return phone.strip()


def fmt_money(v) -> str:
    """1000000 → '1 000 000'"""
    if v is None: return '0'
    try:
        return f"{int(float(v)):,}".replace(',', ' ')
    except Exception:
        return str(v)


# Статусы SEO-проекта (для селектов и метрик)
PROJECT_STATUSES = ['new', 'in_work', 'active', 'paused', 'done', 'lost']
PROJECT_STATUS_LABELS = {
    'new': '🆕 Новый',
    'in_work': '⚙️ В работе',
    'active': '✅ Активный',
    'paused': '⏸ Пауза',
    'done': '🏁 Завершён',
    'lost': '❌ Отказ',
}

TASK_STATUSES = ['open', 'in_progress', 'done', 'cancelled']
TASK_STATUS_LABELS = {
    'open': '📋 Открыта',
    'in_progress': '⚙️ В работе',
    'done': '✅ Готово',
    'cancelled': '❌ Отменена',
}

TASK_PRIORITIES = ['low', 'medium', 'high', 'urgent']
TASK_PRIORITY_LABELS = {
    'low': '🔵 Низкий',
    'medium': '🟡 Средний',
    'high': '🟠 Высокий',
    'urgent': '🔴 Срочно',
}
