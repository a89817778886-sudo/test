# -*- coding: utf-8 -*-
"""crm_auth.py — аутентификация и роли для Streamlit (логин/пароль, admin/manager)."""

from __future__ import annotations

import hashlib
import secrets
import datetime as dt
from typing import Optional, Dict, Any

import streamlit as st

import crm_db
from crm_db import ROLE_ADMIN, ROLE_MANAGER, ROLE_LABELS


def _hash_password(password: str, salt: str) -> str:
    return hashlib.sha256((salt + password).encode("utf-8")).hexdigest()


def _make_password_hash(password: str) -> str:
    salt = secrets.token_hex(8)
    return f"{salt}${_hash_password(password, salt)}"


def _verify_password(password: str, stored_hash: str) -> bool:
    try:
        salt, digest = stored_hash.split("$", 1)
    except ValueError:
        return False
    return _hash_password(password, salt) == digest


def ensure_default_admin() -> None:
    conn = crm_db.get_conn()
    row = conn.execute("SELECT COUNT(*) AS cnt FROM users").fetchone()
    if row["cnt"] == 0:
        now = dt.datetime.now().isoformat(timespec="seconds")
        conn.execute("""
            INSERT INTO users (username, password_hash, full_name, role, is_active, created_at)
            VALUES (?,?,?,?,1,?)
        """, ("admin", _make_password_hash("admin"), "Администратор", ROLE_ADMIN, now))
        conn.commit()
    conn.close()


def create_user(username: str, password: str, full_name: str = "",
                 role: str = ROLE_MANAGER, telegram_chat_id: str = "") -> int:
    conn = crm_db.get_conn()
    now = dt.datetime.now().isoformat(timespec="seconds")
    cur = conn.execute("""
        INSERT INTO users (username, password_hash, full_name, role, telegram_chat_id, is_active, created_at)
        VALUES (?,?,?,?,?,1,?)
    """, (username, _make_password_hash(password), full_name, role, telegram_chat_id, now))
    conn.commit()
    uid = cur.lastrowid
    conn.close()
    return uid


def list_users(active_only: bool = False):
    conn = crm_db.get_conn()
    q = "SELECT * FROM users"
    if active_only:
        q += " WHERE is_active=1"
    q += " ORDER BY username"
    rows = conn.execute(q).fetchall()
    conn.close()
    return [dict(r) for r in rows]


def set_user_active(user_id: int, is_active: bool) -> None:
    conn = crm_db.get_conn()
    conn.execute("UPDATE users SET is_active=? WHERE id=?", (int(is_active), user_id))
    conn.commit()
    conn.close()


def change_password(user_id: int, new_password: str) -> None:
    conn = crm_db.get_conn()
    conn.execute("UPDATE users SET password_hash=? WHERE id=?", (_make_password_hash(new_password), user_id))
    conn.commit()
    conn.close()


def authenticate(username: str, password: str) -> Optional[Dict[str, Any]]:
    conn = crm_db.get_conn()
    row = conn.execute("SELECT * FROM users WHERE username=? AND is_active=1", (username,)).fetchone()
    conn.close()
    if not row:
        return None
    if not _verify_password(password, row["password_hash"]):
        return None
    return {"id": row["id"], "username": row["username"], "full_name": row["full_name"],
            "role": row["role"], "telegram_chat_id": row["telegram_chat_id"]}


def require_login() -> Optional[Dict[str, Any]]:
    ensure_default_admin()
    if "crm_user" in st.session_state:
        return st.session_state["crm_user"]

    st.title("Вход в систему")
    st.caption("Приложение КП по кранам ЛКС — доступ только для авторизованных сотрудников")

    with st.form("crm_login_form"):
        username = st.text_input("Логин")
        password = st.text_input("Пароль", type="password")
        submitted = st.form_submit_button("Войти", type="primary")

    if submitted:
        user = authenticate(username.strip(), password)
        if user:
            st.session_state["crm_user"] = user
            st.rerun()
        else:
            st.error("Неверный логин или пароль.")

    st.info("Первый вход: логин **admin**, пароль **admin** — обязательно смените пароль после входа во вкладке «Настройки».")
    return None


def logout_button():
    user = st.session_state.get("crm_user")
    if not user:
        return
    with st.sidebar:
        st.markdown(f"**{user.get('full_name') or user['username']}** ({ROLE_LABELS.get(user['role'], user['role'])})")
        if st.button("Выйти", key="crm_logout_btn"):
            st.session_state.pop("crm_user", None)
            st.rerun()


def is_admin(user: Dict[str, Any]) -> bool:
    return bool(user) and user.get("role") == ROLE_ADMIN


def render_users_admin_tab():
    user = st.session_state.get("crm_user")
    if not is_admin(user):
        st.warning("Доступно только администратору.")
        return

    st.subheader("Пользователи и роли")
    users = list_users()
    for u in users:
        c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
        with c1:
            st.write(f"**{u['username']}** — {u['full_name'] or '—'}")
        with c2:
            st.write(ROLE_LABELS.get(u["role"], u["role"]))
        with c3:
            st.write("Активен" if u["is_active"] else "Отключён")
        with c4:
            if st.button("Отключить" if u["is_active"] else "Включить", key=f"toggle_user_{u['id']}"):
                set_user_active(u["id"], not u["is_active"])
                st.rerun()

    st.markdown("---")
    st.markdown("**Добавить пользователя**")
    with st.form("add_user_form"):
        new_username = st.text_input("Логин")
        new_full_name = st.text_input("Имя сотрудника")
        new_password = st.text_input("Пароль", type="password")
        new_role = st.selectbox("Роль", list(ROLE_LABELS.keys()), format_func=lambda k: ROLE_LABELS[k])
        new_tg = st.text_input("Telegram chat_id (для уведомлений, необязательно)")
        submitted = st.form_submit_button("Создать пользователя")
    if submitted:
        if not new_username or not new_password:
            st.error("Логин и пароль обязательны.")
        else:
            try:
                create_user(new_username.strip(), new_password, new_full_name.strip(), new_role, new_tg.strip())
                st.success(f"Пользователь {new_username} создан.")
                st.rerun()
            except Exception as e:
                st.error(f"Не удалось создать пользователя: {e}")

    st.markdown("---")
    st.markdown("**Сменить свой пароль**")
    with st.form("change_own_password_form"):
        cur_pass_new = st.text_input("Новый пароль", type="password", key="own_new_pass")
        submitted2 = st.form_submit_button("Сменить пароль")
    if submitted2 and cur_pass_new:
        change_password(user["id"], cur_pass_new)
        st.success("Пароль изменён.")
