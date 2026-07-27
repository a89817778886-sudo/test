# -*- coding: utf-8 -*-
"""crm_notify.py — email-отправка КП, Telegram-уведомления, напоминания."""

from __future__ import annotations

import smtplib
import ssl
import os
from email.message import EmailMessage
from typing import Dict, Any

try:
    import streamlit as st
    _SECRETS = st.secrets
except Exception:
    _SECRETS = {}


def _get_secret(section: str, key: str, env_fallback: str = "") -> str:
    try:
        return _SECRETS[section][key]
    except Exception:
        return os.environ.get(env_fallback, "")


def send_kp_email(to_email: str, subject: str, body_text: str,
                   attachment_bytes: bytes, attachment_filename: str) -> Dict[str, Any]:
    host = _get_secret("smtp", "host", "SMTP_HOST")
    port = int(_get_secret("smtp", "port", "SMTP_PORT") or 465)
    user = _get_secret("smtp", "user", "SMTP_USER")
    password = _get_secret("smtp", "password", "SMTP_PASSWORD")
    from_name = _get_secret("smtp", "from_name", "SMTP_FROM_NAME") or user

    if not (host and user and password):
        return {"ok": False, "error": "SMTP не настроен — заполните .streamlit/secrets.toml [smtp]"}

    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{from_name} <{user}>"
    msg["To"] = to_email
    msg.set_content(body_text)
    msg.add_attachment(attachment_bytes, maintype="application", subtype="octet-stream", filename=attachment_filename)

    try:
        context = ssl.create_default_context()
        with smtplib.SMTP_SSL(host, port, context=context, timeout=15) as server:
            server.login(user, password)
            server.send_message(msg)
        return {"ok": True, "error": ""}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def send_telegram_message(chat_id: str, text: str) -> Dict[str, Any]:
    import requests
    bot_token = _get_secret("telegram", "bot_token", "TELEGRAM_BOT_TOKEN")
    if not bot_token:
        return {"ok": False, "error": "Telegram bot_token не настроен — заполните .streamlit/secrets.toml [telegram]"}
    if not chat_id:
        return {"ok": False, "error": "chat_id не указан"}
    url = f"https://api.telegram.org/bot{bot_token}/sendMessage"
    try:
        resp = requests.post(url, data={"chat_id": chat_id, "text": text}, timeout=10)
        resp.raise_for_status()
        return {"ok": True, "error": ""}
    except Exception as e:
        return {"ok": False, "error": str(e)}


def notify_new_quote(user_chat_id: str, kp_number: str, customer_name: str, total: float) -> None:
    if not user_chat_id:
        return
    text = f"📄 Новый КП {kp_number}\nКлиент: {customer_name}\nСумма: {total:,.0f} ₽".replace(",", " ")
    send_telegram_message(user_chat_id, text)


def notify_status_change(user_chat_id: str, kp_number: str, new_status_label: str) -> None:
    if not user_chat_id:
        return
    text = f"🔄 КП {kp_number} → статус «{new_status_label}»"
    send_telegram_message(user_chat_id, text)


def notify_stale_quote(user_chat_id: str, kp_number: str, days: int) -> None:
    if not user_chat_id:
        return
    text = f"⏰ КП {kp_number} без движения уже {days} дн. — возможно, стоит связаться с клиентом."
    send_telegram_message(user_chat_id, text)
