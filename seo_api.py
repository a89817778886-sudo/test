"""
SEO API — функции приёма заявок с внешнего сайта (SEO Автопилот).

Использование:
1) Внутри Streamlit: from seo_api import receive_site_lead → receive_site_lead(payload)
2) Как HTTP: запусти этот файл — uvicorn seo_api:app --host 0.0.0.0 --port 8100
   Endpoint: POST /api/seo-lead  (JSON)

Токен приёма: SEO_LEAD_TOKEN (env variable или st.secrets)
"""
from __future__ import annotations
import os
from typing import Dict, Any, Optional

from seo_db import SEOLead, create_lead, create_project, mark_lead_processed
from seo_utils import extract_domain, validate_lead_payload, normalize_phone


def _get_secret(name: str, default: str = '') -> str:
    """Читает секрет из st.secrets или env."""
    try:
        import streamlit as st
        v = st.secrets.get(name, None)
        if v: return str(v)
    except Exception:
        pass
    return os.environ.get(name, default)


# ==================== ГЛАВНАЯ ФУНКЦИЯ ПРИЁМА ====================

def receive_site_lead(payload: Dict[str, Any], token: str = None) -> Dict[str, Any]:
    """
    Принимает лид с внешнего сайта. Возвращает dict со статусом.

    payload = {
      "source": "seo-autopilot-site",   # опционально
      "name": "Иван Иванов",
      "phone": "+7 999 123-45-67",       # обязательно (или email)
      "email": "ivan@example.com",       # обязательно (или phone)
      "company": "ООО Ромашка",
      "site_url": "https://client-site.ru",  # обязательно
      "comment": "Хочу продвижение",
      "utm_source": "yandex",
      "utm_medium": "cpc",
      "utm_campaign": "seo-autopilot",
      "utm_term": "seo продвижение",
      "utm_content": "banner1",
      "referrer": "https://...",
      "ip": "...", "user_agent": "..."
    }

    Возвращает: {'ok': True, 'lead_id': 42, 'domain': 'client-site.ru'}
                {'ok': False, 'error': '...'}
    """
    # Проверка токена (если задан в secrets)
    expected = _get_secret('SEO_LEAD_TOKEN', '')
    if expected and token != expected:
        return {'ok': False, 'error': 'Invalid or missing token'}

    ok, err = validate_lead_payload(payload)
    if not ok:
        return {'ok': False, 'error': err}

    lead = SEOLead(
        source=str(payload.get('source', 'seo-autopilot-site')).strip(),
        name=str(payload.get('name', '')).strip()[:200],
        phone=normalize_phone(str(payload.get('phone', '')).strip()),
        email=str(payload.get('email', '')).strip()[:200],
        company=str(payload.get('company', '')).strip()[:250],
        site_url=str(payload.get('site_url', '')).strip()[:500],
        comment=str(payload.get('comment', '')).strip()[:2000],
        utm_source=str(payload.get('utm_source', '')).strip()[:100],
        utm_medium=str(payload.get('utm_medium', '')).strip()[:100],
        utm_campaign=str(payload.get('utm_campaign', '')).strip()[:200],
        utm_term=str(payload.get('utm_term', '')).strip()[:200],
        utm_content=str(payload.get('utm_content', '')).strip()[:200],
        ip=str(payload.get('ip', '')).strip()[:64],
        user_agent=str(payload.get('user_agent', '')).strip()[:500],
        referrer=str(payload.get('referrer', '')).strip()[:500],
    )
    try:
        lead_id = create_lead(lead)
        return {
            'ok': True,
            'lead_id': lead_id,
            'domain': extract_domain(lead.site_url),
        }
    except Exception as e:
        return {'ok': False, 'error': f'DB error: {e}'}


def create_project_from_lead_id(lead_id: int, customer_id: Optional[int] = None,
                                 monthly_fee: float = 0, region: str = '') -> Dict[str, Any]:
    """Создать SEO-проект из существующего лида."""
    from seo_db import get_lead
    lead = get_lead(lead_id)
    if not lead:
        return {'ok': False, 'error': f'Lead #{lead_id} not found'}
    site_url = lead.get('site_url', '')
    project_id = create_project(
        customer_id=customer_id,
        site_url=site_url,
        domain=extract_domain(site_url),
        region=region or '',
        monthly_fee=monthly_fee,
        status='new',
    )
    mark_lead_processed(lead_id, customer_id=customer_id, project_id=project_id)
    return {'ok': True, 'project_id': project_id, 'domain': extract_domain(site_url)}


# ==================== ОПЦИОНАЛЬНЫЙ FASTAPI ENDPOINT ====================
# Не мешает Streamlit — запускается отдельным процессом
# Команда: uvicorn seo_api:app --host 0.0.0.0 --port 8100

try:
    from fastapi import FastAPI, Request, Header, HTTPException
    from fastapi.middleware.cors import CORSMiddleware

    app = FastAPI(title="SEO Autopilot Lead API", version="1.0")

    # CORS — чтобы можно было отправлять формы с любого сайта клиента
    app.add_middleware(
        CORSMiddleware,
        allow_origins=["*"],
        allow_credentials=False,
        allow_methods=["POST", "OPTIONS", "GET"],
        allow_headers=["*"],
    )

    @app.get("/api/health")
    def health():
        return {"status": "ok", "service": "seo-lead-api"}

    @app.post("/api/seo-lead")
    async def api_seo_lead(request: Request,
                            x_token: Optional[str] = Header(None)):
        """Принимает JSON payload от формы сайта. Опциональный X-Token для защиты."""
        try:
            payload = await request.json()
        except Exception:
            payload = {}

        # Добавляем IP и user-agent из headers
        client_ip = request.headers.get("x-forwarded-for", "") or (
            request.client.host if request.client else "")
        user_agent = request.headers.get("user-agent", "")
        referrer = request.headers.get("referer", "")
        payload.setdefault('ip', client_ip)
        payload.setdefault('user_agent', user_agent)
        payload.setdefault('referrer', referrer)

        result = receive_site_lead(payload, token=x_token)
        if not result.get('ok'):
            raise HTTPException(status_code=400, detail=result.get('error', 'Invalid payload'))
        return result

except ImportError:
    # FastAPI не установлен — работаем только как модуль
    app = None
