# -*- coding: utf-8 -*-
"""
crm_ui.py — главный UI-слой CRM для приложения КП по кранам/траверсам ЛКС.
Версия 2.0. Объединяет все 20 пунктов доработки в готовые Streamlit-вкладки.

Подключение в app.py:

    import crm_db, crm_auth, crm_ui

    user = crm_auth.require_login()
    if not user:
        st.stop()
    crm_auth.logout_button()

    tab_calc, tab_dash, tab_funnel, tab_clients, tab_quotes, \
        tab_reminders, tab_sales, tab_settings = st.tabs([
            "Расчёт КП", "Дашборд", "Воронка", "Клиенты",
            "КП", "Напоминания", "Продажи", "Настройки",
        ])

    with tab_calc:
        ... существующий расчёт КП ...
        crm_ui.render_save_quote_block(
            base_total=total, product_type="Кран",
            product_model=f"{series} {capacity}кг {boom}м",
            include_montage=include_montage,
            spec_lines=spec_df.to_dict("records"),
            pdf_bytes=pdf_bytes, pdf_filename=f"КП_{kp_number}.pdf",
        )

    with tab_dash:
        crm_ui.render_dashboard_tab()
    with tab_funnel:
        crm_ui.render_funnel_tab()
    with tab_clients:
        crm_ui.render_customers_tab()
    with tab_quotes:
        crm_ui.render_quotes_tab()
    with tab_reminders:
        crm_ui.render_reminders_tab()
    with tab_sales:
        crm_ui.render_sales_tab()
    with tab_settings:
        crm_ui.render_settings_tab()
"""

from __future__ import annotations

import datetime as dt
import os as _os
import streamlit as st
import pandas as pd

import crm_db
import crm_export
import crm_notify
import crm_backup
import crm_inn
from crm_db import (
    Customer, QuoteRecord, QuoteItem,
    STATUS_DRAFT, STATUS_SENT, STATUS_WON, STATUS_LOST, STATUS_LABELS, STATUS_ORDER,
    PRODUCT_TYPES, LOSS_REASONS, DISCOUNT_MIN_PCT, DISCOUNT_MAX_PCT, DISCOUNT_RECOMMENDED,
)


# Цвета градиентных карточек — 4 основных цвета
_GRADIENT_CLASSES = [
    "gradient-card-purple",
    "gradient-card-teal",
    "gradient-card-blue",
    "gradient-card-orange",
]


def _gradient_metric(label: str, value, icon: str = "", color_idx: int = 0):
    """Рендер градиентной карточки-метрики.
    color_idx: 0=фиолет, 1=бирюз, 2=синий, 3=оранжевый."""
    _cls = _GRADIENT_CLASSES[color_idx % 4]
    _icon = f"{icon} " if icon else ""
    st.markdown(f'''<div class="{_cls}">
        <div class="gradient-card-label">{_icon}{label}</div>
        <div class="gradient-card-value">{value}</div>
    </div>''', unsafe_allow_html=True)


def _money(v: float) -> str:
    return f"{v:,.0f} ₽".replace(",", " ")


def _current_user():
    return st.session_state.get("crm_user", {})


# ============================================================
# 1. Блок сохранения КП (из вкладки "Расчёт КП")
# ============================================================

def render_save_quote_block(
    base_total: float,
    product_type: str,
    product_model: str,
    include_montage: bool,
    spec_lines: list[dict] | None = None,
    kp_number: str | None = None,
    pdf_bytes: bytes | None = None,
    pdf_filename: str = "kp.pdf",
):
    st.subheader("Сохранить КП в базу")
    user = _current_user()

    if base_total <= 0:
        st.info("Сначала сформируйте расчёт КП — сумма пока равна 0.")
        return

    col1, col2 = st.columns(2)
    with col1:
        inn = st.text_input("ИНН клиента", key="crm_inn", max_chars=12)
        phone = st.text_input("Телефон клиента", key="crm_phone", placeholder="+7 900 000-00-00")
        email = st.text_input("Email клиента", key="crm_email", placeholder="client@example.com")
    with col2:
        name_short = st.text_input("Наименование организации (краткое)", key="crm_name_short")
        delivery_city = st.text_input("Город доставки", key="crm_delivery_city")
        requisites_raw = st.text_area("Реквизиты (для автоподгрузки)", key="crm_req_raw", height=80)

    btn_col1, btn_col2 = st.columns(2)
    with btn_col1:
        if inn and st.button("Подгрузить данные клиента по ИНН", key="crm_autofill_btn"):
            autofilled = crm_db.autofill_customer_by_inn(inn, requisites_raw)
            st.session_state["crm_autofilled"] = autofilled
            if autofilled.name_short or autofilled.phone:
                st.success("Данные клиента найдены в базе и подгружены.")
            else:
                st.warning("Клиент не найден в локальной базе — заполните карточку вручную.")
    with btn_col2:
        if inn and st.button("🔎 Проверить ИНН в ЕГРЮЛ (онлайн)", key="crm_egrul_btn"):
            info = crm_inn.lookup_inn(inn)
            if info:
                st.session_state["crm_egrul_info"] = info
                st.success(f"Найдено: {info['name_short']} — {info['status']}")
            else:
                st.warning("Не удалось найти организацию по ИНН в ЕГРЮЛ (сервис недоступен или ИНН не найден).")

    autofilled = st.session_state.get("crm_autofilled")
    egrul_info = st.session_state.get("crm_egrul_info")

    st.markdown("**Состав КП**")
    montage_label = "с монтажом" if include_montage else "без монтажа"
    st.write(f"Товар: **{product_type}** ({product_model}) — {montage_label}")

    st.markdown("**Скидка**")
    disc_col1, disc_col2 = st.columns([2, 1])
    with disc_col1:
        discount_pct = st.slider(
            "Скидка, %", float(DISCOUNT_MIN_PCT), float(DISCOUNT_MAX_PCT), 0.0, 1.0,
            key="crm_discount_pct",
            help=f"Рекомендуемый диапазон {DISCOUNT_RECOMMENDED[0]}–{DISCOUNT_RECOMMENDED[1]}%",
        )
    disc_preview = crm_db.apply_discount(base_total, discount_pct)
    with disc_col2:
        st.metric("Итого со скидкой", _money(disc_preview["final_total"]))
    st.caption(
        f"Сумма без скидки: {_money(base_total)} · Скидка: {discount_pct:.0f}% "
        f"(−{_money(disc_preview['discount_amount'])})"
    )

    probability_pct = st.slider(
        "Вероятность закрытия сделки, % (для прогноза выручки)",
        0, 100, 50, 5, key="crm_probability_pct",
    )

    if st.button("💾 Сохранить КП в базу", key="crm_save_btn", type="primary"):
        customer = Customer(
            inn=inn.strip(), phone=phone.strip(), email=email.strip(), name_short=name_short.strip(),
        )
        if autofilled:
            for f in ("kpp", "ogrn", "name_full", "address", "director_position",
                      "director_fio", "bank", "bik", "rs", "ks"):
                val = getattr(autofilled, f, "")
                if val:
                    setattr(customer, f, val)
            if not customer.name_short and autofilled.name_short:
                customer.name_short = autofilled.name_short
        if egrul_info:
            if not customer.name_short:
                customer.name_short = egrul_info.get("name_short", "")
            if not customer.name_full:
                customer.name_full = egrul_info.get("name_full", "")
            if not customer.address:
                customer.address = egrul_info.get("address", "")
            if not customer.ogrn:
                customer.ogrn = egrul_info.get("ogrn", "")

        customer_id = crm_db.upsert_customer(customer)
        if egrul_info:
            crm_db.update_customer_risk(customer_id, crm_inn.check_counterparty_risk(inn))

        items = []
        if spec_lines:
            for ln in spec_lines:
                items.append(QuoteItem(
                    code=str(ln.get("Код", ln.get("code", ""))),
                    name=str(ln.get("Наименование", ln.get("name", ""))),
                    unit=str(ln.get("Ед.", ln.get("unit", "шт"))),
                    qty=float(ln.get("Кол-во", ln.get("qty", 1)) or 1),
                    price=float(ln.get("Цена", ln.get("price", 0)) or 0),
                ))

        record = QuoteRecord(
            kp_number=kp_number or f"КП-{pd.Timestamp.now().strftime('%Y%m%d-%H%M%S')}",
            customer_id=customer_id, product_type=product_type, product_model=product_model,
            include_montage=include_montage, delivery_city=delivery_city.strip(),
            base_total=base_total, discount_pct=discount_pct, status=STATUS_DRAFT,
            probability_pct=probability_pct, items=items,
        )
        quote_id = crm_db.save_quote(record, user_id=user.get("id"), username=user.get("username", ""))
        st.success(f"КП сохранён в базу, id={quote_id}. Статус: «{STATUS_LABELS[STATUS_DRAFT]}».")

        if user.get("telegram_chat_id"):
            crm_notify.notify_new_quote(user["telegram_chat_id"], record.kp_number,
                                         customer.name_short or customer.inn, disc_preview["final_total"])

        if pdf_bytes and email:
            st.session_state[f"crm_pdf_{quote_id}"] = (pdf_bytes, pdf_filename)
            if st.button("📧 Отправить КП клиенту на email", key=f"crm_send_email_{quote_id}"):
                res = crm_notify.send_kp_email(
                    email, f"Коммерческое предложение {record.kp_number}",
                    f"Добрый день!\n\nВысылаем коммерческое предложение {record.kp_number}.\n\nС уважением.",
                    pdf_bytes, pdf_filename,
                )
                if res["ok"]:
                    st.success("КП отправлен клиенту на email.")
                else:
                    st.error(f"Не удалось отправить: {res['error']}")

        for k in ("crm_autofilled", "crm_egrul_info"):
            st.session_state.pop(k, None)


# ============================================================
# 2. Дашборд
# ============================================================

def render_dashboard_tab():
    st.subheader("📊 Дашборд продаж")

    summary = crm_db.sales_summary()
    forecast = crm_db.revenue_forecast()
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        _gradient_metric("Продаж всего", int(summary["cnt"]), icon="📄", color_idx=0)
    with c2:
        _gradient_metric("Сумма продаж", _money(summary["total"]), icon="💰", color_idx=1)
    with c3:
        _gradient_metric("Средняя скидка", f"{summary['avg_discount']:.1f}%", icon="🏷️", color_idx=2)
    with c4:
        _gradient_metric("Прогноз выручки", _money(forecast["weighted_forecast"]), icon="📈", color_idx=3)

    st.markdown("**Продажи по месяцам**")
    by_month = crm_db.sales_by_month()
    if by_month:
        df_month = pd.DataFrame(by_month)
        st.bar_chart(df_month.set_index("month")["total"])
    else:
        st.info("Пока нет данных о продажах.")

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("**Продажи по типу товара**")
        by_type = crm_db.sales_by_product_type()
        if by_type:
            df_type = pd.DataFrame(by_type)
            st.dataframe(df_type, use_container_width=True, hide_index=True)
        else:
            st.info("Нет данных.")
    with col2:
        st.markdown("**Причины отказов**")
        losses = crm_db.loss_reason_breakdown()
        if losses:
            df_loss = pd.DataFrame(losses)
            st.dataframe(df_loss, use_container_width=True, hide_index=True)
        else:
            st.info("Отказов не зафиксировано.")

    st.markdown("**Отчёт по менеджерам**")
    managers = crm_db.manager_report()
    if managers:
        st.dataframe(pd.DataFrame(managers), use_container_width=True, hide_index=True)
    else:
        st.info("Нет данных по менеджерам.")


# ============================================================
# 3. Воронка продаж (Kanban)
# ============================================================

def render_funnel_tab():
    st.subheader("🎯 Воронка продаж")
    funnel = crm_db.conversion_funnel()

    # Градиентные шапки колонок воронки
    _funnel_colors = [0, 2, 1, 3]  # Каждому статусу — свой цвет
    _funnel_icons = {"draft": "📝", "sent": "📤", "won": "🏆", "lost": "❌"}

    cols = st.columns(len(STATUS_ORDER))
    for i, status in enumerate(STATUS_ORDER):
        with cols[i]:
            _icon = _funnel_icons.get(status, "")
            _gradient_metric(STATUS_LABELS[status], funnel.get(status, 0),
                             icon=_icon, color_idx=_funnel_colors[i % 4])
            quotes = crm_db.list_quotes(status=status)
            for q in quotes[:20]:
                with st.container(border=True):
                    st.write(f"**{q['kp_number']}**")
                    st.caption(f"{q.get('name_short') or 'без клиента'}")
                    st.write(_money(q["final_total"]))
                    next_options = [s for s in STATUS_ORDER if s != status]
                    new_status = st.selectbox(
                        "Переместить в", next_options,
                        format_func=lambda s: STATUS_LABELS[s],
                        key=f"funnel_move_{q['id']}",
                    )
                    loss_reason = ""
                    if new_status == STATUS_LOST:
                        loss_reason = st.selectbox(
                            "Причина отказа", LOSS_REASONS, key=f"funnel_loss_{q['id']}",
                        )
                    if st.button("Переместить", key=f"funnel_btn_{q['id']}"):
                        user = _current_user()
                        crm_db.set_quote_status(
                            q["id"], new_status, loss_reason=loss_reason,
                            user_id=user.get("id"), username=user.get("username", ""),
                        )
                        if user.get("telegram_chat_id"):
                            crm_notify.notify_status_change(
                                user["telegram_chat_id"], q["kp_number"], STATUS_LABELS[new_status],
                            )
                        st.rerun()


# ============================================================
# 4. Клиенты (список + карточка)
# ============================================================

def render_customers_tab():
    st.subheader("👥 Клиенты")

    # --- ➕ Новый клиент ---
    with st.expander("➕ Добавить нового клиента", expanded=False):
        st.caption("💡 Заполните вручную или вставьте текст с реквизитами в поле ниже — будет автоматически разобран.")

        # Copy-paste через текст
        _paste_txt_cust = st.text_area(
            "Или вставьте текст с реквизитами (для автораспознавания)",
            key="crm_new_cust_paste",
            height=100,
            placeholder="ООО «Ромашка»\nИНН 7712345678, КПП 771201001\nАдрес: ...")

        if st.button("🔍 Распознать из текста", key="crm_new_cust_parse", disabled=not _paste_txt_cust.strip()):
            try:
                from external_kp_parser import extract_requisites_from_text
                _r = extract_requisites_from_text(_paste_txt_cust)
                if _r.is_empty():
                    st.warning("Не удалось распознать. Заполните поля вручную.")
                else:
                    st.session_state["crm_new_cust_name"] = _r.company_short or ""
                    st.session_state["crm_new_cust_full"] = _r.company_full or ""
                    st.session_state["crm_new_cust_inn"] = _r.inn or ""
                    st.session_state["crm_new_cust_kpp"] = _r.kpp or ""
                    st.session_state["crm_new_cust_addr"] = _r.address or ""
                    st.session_state["crm_new_cust_phone"] = _r.phone or ""
                    st.session_state["crm_new_cust_email"] = _r.email or ""
                    st.session_state["crm_new_cust_bank"] = _r.bank_name or ""
                    st.session_state["crm_new_cust_bik"] = _r.bank_bik or ""
                    st.session_state["crm_new_cust_rs"] = _r.bank_account or ""
                    st.session_state["crm_new_cust_ks"] = _r.corr_account or ""
                    st.session_state["crm_new_cust_dir"] = _r.director_short or ""
                    st.success("✅ Распознано. Проверьте поля ниже и сохраните.")
                    st.rerun()
            except Exception as e:
                st.error(f"Ошибка: {e}")

        st.markdown("---")
        st.markdown("**Заполните поля (или отредактируйте распознанное):**")
        _nc_c1, _nc_c2 = st.columns(2)
        with _nc_c1:
            _new_name = st.text_input("Краткое название *", key="crm_new_cust_name")
            _new_full = st.text_input("Полное название", key="crm_new_cust_full")
            _new_inn = st.text_input("ИНН *", key="crm_new_cust_inn", help="Служит уникальным ключом — дубли будут объединены")
            _new_kpp = st.text_input("КПП", key="crm_new_cust_kpp")
            _new_addr = st.text_area("Юр. адрес", key="crm_new_cust_addr", height=60)
            _new_phone = st.text_input("Телефон", key="crm_new_cust_phone")
            _new_email = st.text_input("E-mail", key="crm_new_cust_email")
        with _nc_c2:
            _new_bank = st.text_input("Банк", key="crm_new_cust_bank")
            _new_bik = st.text_input("БИК", key="crm_new_cust_bik")
            _new_rs = st.text_input("Р/с", key="crm_new_cust_rs")
            _new_ks = st.text_input("К/с", key="crm_new_cust_ks")
            _new_dir = st.text_input("ФИО директора (Фамилия И.О.)", key="crm_new_cust_dir")
            _new_dir_title = st.text_input("Должность", value="Генеральный директор", key="crm_new_cust_dir_title")
            _new_notes = st.text_area("Заметки (опционально)", key="crm_new_cust_notes", height=60)

        if st.button("💾 Создать клиента в CRM", type="primary", key="crm_new_cust_save",
                     use_container_width=True):
            if not (_new_name.strip() and _new_inn.strip()):
                st.error("✖ Обязательные поля: **Краткое название** и **ИНН**")
            else:
                try:
                    _cust = crm_db.Customer(
                        name_short=_new_name.strip(),
                        name_full=_new_full.strip() or _new_name.strip(),
                        inn=_new_inn.strip(),
                        kpp=_new_kpp.strip(),
                        address=_new_addr.strip(),
                        phone=_new_phone.strip(),
                        email=_new_email.strip(),
                        bank=_new_bank.strip(),
                        bik=_new_bik.strip(),
                        rs=_new_rs.strip(),
                        ks=_new_ks.strip(),
                        director_fio=_new_dir.strip(),
                        director_position=_new_dir_title.strip(),
                    )
                    _cust_id = crm_db.upsert_customer(_cust)
                    # Заметка
                    if _new_notes.strip():
                        user = _current_user()
                        crm_db.add_customer_note(_cust_id, _new_notes.strip(),
                                                 user.get("id"), user.get("username", ""))
                    st.success(f"✅ Клиент **{_new_name}** сохранён (id={_cust_id}). "
                              "Теперь его можно выбрать при выставлении КП.")
                    # Очищаем поля
                    for k in list(st.session_state.keys()):
                        if k.startswith("crm_new_cust_"):
                            del st.session_state[k]
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка сохранения: {e}")

    st.markdown("---")

    query = st.text_input("🔍 Поиск по ИНН / названию / телефону / email", key="crm_customer_search")
    customers = crm_db.search_customers(query or "")
    if not customers:
        if query.strip():
            st.info("По вашему запросу клиенты не найдены.")
        else:
            st.info("База клиентов пуста. Нажмите «➕ Добавить нового клиента» вверху, чтобы добавить первого.")
        return

    st.caption(f"Найдено клиентов: **{len(customers)}**")
    for c in customers:
        risk_badge = f" · {c.risk_flag}" if c.risk_flag else ""
        with st.expander(f"{c.name_short or c.inn or 'Без названия'} · ИНН {c.inn}{risk_badge}"):
            _render_customer_card(c)


def _render_customer_card(c: Customer):
    st.write(f"**Телефон:** {c.phone or '—'}  ·  **Email:** {c.email or '—'}")
    st.write(f"**Адрес:** {c.address or '—'}")
    if c.risk_flag:
        st.write(f"**Проверка контрагента:** {c.risk_flag} (на {c.risk_checked_at})")

    full = crm_db.get_customer_full(c.id)

    st.markdown("**История КП**")
    if full["quotes"]:
        df = pd.DataFrame(full["quotes"])[["kp_number", "product_type", "final_total", "status", "created_at"]]
        df["status"] = df["status"].map(STATUS_LABELS)
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.caption("КП не выставлялись.")

    st.markdown("**История продаж**")
    if full["sales"]:
        st.dataframe(pd.DataFrame(full["sales"]), use_container_width=True, hide_index=True)
    else:
        st.caption("Продаж не было.")

    st.markdown("**Заметки / история общения**")
    for note in full["notes"]:
        _ncol1, _ncol2 = st.columns([10, 1])
        with _ncol1:
            st.write(f"🗒️ {note['created_at']} ({note.get('username') or '—'}): {note['note']}")
        with _ncol2:
            if st.button("🗑", key=f"crm_del_note_{note['id']}",
                         help="Удалить заметку"):
                try:
                    crm_db.delete_customer_note(note['id'])
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка: {e}")
    new_note = st.text_area("Новая заметка", key=f"crm_new_note_{c.id}", height=60)
    if st.button("Добавить заметку", key=f"crm_add_note_btn_{c.id}"):
        user = _current_user()
        if new_note.strip():
            crm_db.add_customer_note(c.id, new_note.strip(), user.get("id"), user.get("username", ""))
            st.rerun()

    st.markdown("**Напоминания по клиенту**")
    for r in full["reminders"]:
        status_mark = "✅" if r["is_done"] else "⏰"
        _rcol1, _rcol2 = st.columns([10, 1])
        with _rcol1:
            st.write(f"{status_mark} {r['due_at']} — {r['message']}")
        with _rcol2:
            if st.button("🗑", key=f"crm_del_rem_{r['id']}",
                         help="Удалить напоминание"):
                try:
                    crm_db.delete_reminder(r['id'])
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка: {e}")
    rc1, rc2, rc3 = st.columns([2, 3, 1])
    with rc1:
        due_date = st.date_input("Дата напоминания", key=f"crm_rem_date_{c.id}")
    with rc2:
        rem_msg = st.text_input("Текст напоминания", key=f"crm_rem_msg_{c.id}", placeholder="Позвонить клиенту")
    with rc3:
        if st.button("Добавить", key=f"crm_add_rem_btn_{c.id}"):
            user = _current_user()
            crm_db.add_reminder(
                due_at=dt.datetime.combine(due_date, dt.time(9, 0)).isoformat(timespec="seconds"),
                message=rem_msg or "Связаться с клиентом", customer_id=c.id, user_id=user.get("id"),
            )
            st.rerun()

    # --- 🗑 Удаление клиента ---
    st.markdown("---")
    _del_confirm_key = f"crm_del_cust_confirm_{c.id}"
    _del_flag = st.session_state.get(_del_confirm_key, False)

    _delc1, _delc2 = st.columns([1, 3])
    with _delc1:
        if not _del_flag:
            if st.button("🗑 Удалить клиента", key=f"crm_del_cust_btn_{c.id}", type="secondary"):
                st.session_state[_del_confirm_key] = True
                st.rerun()
        else:
            if st.button("✖ Да, удалить навсегда", key=f"crm_del_cust_confirm_{c.id}_go", type="primary"):
                try:
                    crm_db.delete_customer(c.id)
                    st.session_state.pop(_del_confirm_key, None)
                    st.success(f"Клиент «{c.name_short}» и все его КП/заметки удалены.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка: {e}")
    with _delc2:
        if _del_flag:
            st.warning(f"⚠️ **Внимание:** будет удалён клиент «{c.name_short}» и ВСЕ связанные с ним КП, заметки, напоминания. Отменить нельзя.")
            if st.button("↶ Отмена", key=f"crm_del_cust_cancel_{c.id}"):
                st.session_state.pop(_del_confirm_key, None)
                st.rerun()


# ============================================================
# 5. КП (список, история изменений, статус/скидка)
# ============================================================

def _render_kp_view_card(quote_id: int):
    """Карточка одного КП: PDF-просмотр, скачать, редактировать, вернуться."""
    q = crm_db.get_quote(quote_id)
    if not q:
        st.error(f"КП id={quote_id} не найдено в базе. Возможно удалено.")
        if st.button("← Назад к списку", key="crm_view_back_notfound"):
            st.session_state.pop("crm_viewing_kp_id", None)
            st.rerun()
        return

    # Шапка
    _hc1, _hc2 = st.columns([5, 1])
    with _hc1:
        st.markdown(f"### 📄 КП **{q.get('kp_number', '')}**")
        _cust_name = q.get("name_short") or "—"
        _stat = STATUS_LABELS.get(q.get("status", "draft"), "Черновик")
        _created = str(q.get("created_at", "") or "")[:10]
        st.caption(f"Заказчик: **{_cust_name}** · Статус: **{_stat}** · Создано: {_created}")
    with _hc2:
        if st.button("← Назад к списку", key="crm_view_back_top",
                     use_container_width=True, type="secondary"):
            st.session_state.pop("crm_viewing_kp_id", None)
            st.session_state.pop("crm_edit_kp_id", None)
            st.rerun()

    st.markdown("---")

    # Ключевые метрики
    _m1, _m2, _m3, _m4 = st.columns(4)
    _m1.metric("Сумма без скидки", _money(float(q.get("base_total") or 0)))
    _m2.metric("Скидка", f"{float(q.get('discount_pct') or 0):.0f} %")
    _m3.metric("Итого", _money(float(q.get("final_total") or 0)))
    _m4.metric("Товар", f"{q.get('product_type', '')}")

    st.markdown(f"**Модель:** {q.get('product_model', '—')}")
    if q.get("delivery_city"):
        st.markdown(f"**Город доставки:** {q.get('delivery_city')}")
    if q.get("contact_fio") or q.get("phone") or q.get("email"):
        st.markdown(f"**Контакт:** {q.get('contact_fio', '')} · {q.get('phone', '')} · {q.get('email', '')}")
    if q.get("request_summary"):
        st.markdown(f"**Что запрашивал:** {q.get('request_summary')}")

    # Спецификация
    if q.get("items"):
        st.markdown("---")
        st.markdown("**📋 Спецификация:**")
        import pandas as _pd_view
        _spec = _pd_view.DataFrame([{
            "Код": _it.get("code", ""),
            "Наименование": _it.get("name", ""),
            "Ед.": _it.get("unit", "шт"),
            "Кол-во": float(_it.get("qty") or 0),
            "Цена, ₽": float(_it.get("price") or 0),
            "Сумма, ₽": float(_it.get("total") or (float(_it.get("qty") or 0) * float(_it.get("price") or 0))),
        } for _it in q["items"]])
        st.dataframe(_spec, use_container_width=True, hide_index=True)

    st.markdown("---")
    st.markdown("### 🛠️ Действия")

    # Проверяем есть ли сохранённые PDF/DOCX файлы (в crm_files/ через pdf_path/docx_path)
    from pathlib import Path as _P
    _pdf_path = q.get("pdf_path")
    _docx_path = q.get("docx_path")
    _pdf_bytes = None
    _docx_bytes = None
    if _pdf_path and _P(_pdf_path).exists():
        try:
            _pdf_bytes = _P(_pdf_path).read_bytes()
        except Exception:
            pass
    if _docx_path and _P(_docx_path).exists():
        try:
            _docx_bytes = _P(_docx_path).read_bytes()
        except Exception:
            pass

    _ac1, _ac2, _ac3, _ac4 = st.columns(4)

    with _ac1:
        if _pdf_bytes:
            st.download_button(
                "📄 Скачать PDF",
                data=_pdf_bytes,
                file_name=f"КП_{q.get('kp_number', quote_id)}.pdf",
                mime="application/pdf",
                use_container_width=True,
                key="crm_view_dl_pdf")
        else:
            st.button("📄 PDF не сохранён", disabled=True,
                     use_container_width=True, key="crm_view_no_pdf",
                     help="Сгенерируйте PDF заново — нажмите «Открыть в редакторе»")

    with _ac2:
        if _docx_bytes:
            st.download_button(
                "📝 Скачать DOCX",
                data=_docx_bytes,
                file_name=f"КП_{q.get('kp_number', quote_id)}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                key="crm_view_dl_docx")
        else:
            st.button("📝 DOCX не сохранён", disabled=True,
                     use_container_width=True, key="crm_view_no_docx")

    with _ac3:
        if st.button("📝 Открыть в редакторе", type="primary",
                     use_container_width=True, key="crm_view_open_editor",
                     help="Загрузит данные КП в вкладку «📋 КП» для редактирования"):
            # Сохраняем в session_state флаг что открываем это КП в редакторе
            st.session_state["crm_edit_kp_id"] = quote_id
            # Подставляем основные поля в сайдбар КП
            st.session_state["kp_buyer_name"] = q.get("name_short", "")
            if q.get("phone"):
                st.session_state["kp_buyer_phone"] = q["phone"]
            if q.get("email"):
                st.session_state["kp_buyer_email"] = q["email"]
            st.success("✅ Данные КП загружены. Перейдите во вкладку «📋 КП» чтобы отредактировать.")
            st.info("💡 После правки нажмите кнопку «💾 Сохранить КП» — она обновит текущее КП в CRM.")

    with _ac4:
        # Удаление КП
        _del_key = f"crm_view_del_confirm_{quote_id}"
        _del_flag = st.session_state.get(_del_key, False)
        if not _del_flag:
            if st.button("🗑 Удалить КП", use_container_width=True,
                         key="crm_view_del_btn"):
                st.session_state[_del_key] = True
                st.rerun()
        else:
            if st.button("✖ Да, удалить!", type="primary",
                         use_container_width=True, key="crm_view_del_go"):
                try:
                    crm_db.delete_quote(quote_id)
                    st.session_state.pop("crm_viewing_kp_id", None)
                    st.session_state.pop(_del_key, None)
                    st.success(f"КП удалено.")
                    st.rerun()
                except Exception as e:
                    st.error(f"Ошибка: {e}")

    # Если в режиме подтверждения удаления — показываем предупреждение
    if st.session_state.get(f"crm_view_del_confirm_{quote_id}", False):
        st.warning("⚠️ КП будет удалено навсегда (вместе со всеми позициями). Отменить нельзя.")
        if st.button("↶ Отмена удаления", key="crm_view_del_cancel"):
            st.session_state.pop(f"crm_view_del_confirm_{quote_id}", None)
            st.rerun()

    # === РЕДАКТИРОВАНИЕ ПОЗИЦИЙ (товаров) КП ===
    st.markdown("---")
    with st.expander("✏️ Редактировать позиции КП (добавить/удалить товары)", expanded=False):
        st.caption("💡 **Редактирование:** кликни по ячейке → введи значение → Enter. "
                   "**Удаление строки:** отметь ☑️ в колонке «🗑 Уд.» и нажми  **🗑 Удалить отмеченные**. "
                   "**Добавление:** кнопка  **➕ Добавить строку**  ниже таблицы. "
                   "Все изменения сохраняются одной кнопкой «💾 Сохранить» внизу.")

        import pandas as _pd
        _items = q.get("items") or []
        # Буфер в session_state — чтобы «Добавить строку» не теряло текущие правки data_editor
        _buffer_key = f"crm_edit_items_buffer_{quote_id}"
        if _buffer_key not in st.session_state:
            st.session_state[_buffer_key] = [
                {
                    "🗑 Уд.": False,
                    "Код": str(it.get("code") or ""),
                    "Наименование": str(it.get("name") or ""),
                    "Ед.": str(it.get("unit") or "шт"),
                    "Кол-во": float(it.get("qty") or 0),
                    "Цена, ₽": float(it.get("price") or 0),
                }
                for it in _items
            ]

        _edit_df = _pd.DataFrame(st.session_state[_buffer_key])
        if _edit_df.empty:
            _edit_df = _pd.DataFrame(
                columns=["🗑 Уд.", "Код", "Наименование", "Ед.", "Кол-во", "Цена, ₽"]
            )

        _edited = st.data_editor(
            _edit_df,
            num_rows="fixed",  # Не даём streamlit добавлять/удалять — делаем это явными кнопками
            use_container_width=True,
            hide_index=True,
            key=f"crm_edit_items_editor_{quote_id}",
            column_config={
                "🗑 Уд.": st.column_config.CheckboxColumn(
                    width="small",
                    help="Отметь ☑️ чтобы удалить строку при следующем клике по «Удалить отмеченные»"),
                "Код": st.column_config.TextColumn(width="small"),
                "Наименование": st.column_config.TextColumn(width="large"),
                "Ед.": st.column_config.TextColumn(width="small"),
                "Кол-во": st.column_config.NumberColumn(
                    min_value=0, step=1, format="%d"),
                "Цена, ₽": st.column_config.NumberColumn(
                    min_value=0, step=100, format="%.2f"),
            },
        )
        # Синхронизируем буфер с текущей таблице
        st.session_state[_buffer_key] = _edited.to_dict("records")

        # === КНОПКИ: добавить строку / удалить отмеченные ===
        _btn_add, _btn_del = st.columns(2)
        with _btn_add:
            if st.button("➕ Добавить строку",
                         use_container_width=True,
                         key=f"crm_edit_add_row_{quote_id}"):
                _cur = st.session_state[_buffer_key]
                _cur.append({
                    "🗑 Уд.": False,
                    "Код": "",
                    "Наименование": "",
                    "Ед.": "шт",
                    "Кол-во": 1.0,
                    "Цена, ₽": 0.0,
                })
                st.session_state[_buffer_key] = _cur
                st.rerun()
        with _btn_del:
            _to_delete = sum(1 for r in _edited.to_dict("records") if r.get("🗑 Уд."))
            if st.button(f"🗑 Удалить отмеченные ({_to_delete})",
                         use_container_width=True,
                         disabled=(_to_delete == 0),
                         key=f"crm_edit_del_marked_{quote_id}"):
                _cur = st.session_state[_buffer_key]
                st.session_state[_buffer_key] = [r for r in _cur if not r.get("🗑 Уд.")]
                st.rerun()

        # Обновляем _edited для последующего расчёта суммы (без отмеченных к удалению)
        _edited = _pd.DataFrame(st.session_state[_buffer_key])

        # Отображаем предварительную сумму (без скидки)
        try:
            _preview_base = float((_edited["Кол-во"] * _edited["Цена, ₽"]).sum())
        except Exception:
            _preview_base = 0.0

        # === СКИДКА (редактируемая) ===
        _dc1, _dc2 = st.columns([1, 2])
        with _dc1:
            _current_discount = float(q.get("discount_pct") or 0)
            _new_discount = st.number_input(
                "Скидка, %",
                min_value=0.0, max_value=50.0,
                value=_current_discount,
                step=0.5,
                key=f"crm_edit_discount_{quote_id}",
                help="Применяется к base_total. Сохраняется в БД вместе с позициями.")
        with _dc2:
            _preview_final = _preview_base * (1 - _new_discount / 100)
            _disc_amount = _preview_base - _preview_final
            if _new_discount > 0:
                st.info(
                    f"💰 **Сумма позиций:** {_preview_base:,.2f} ₽  \n"
                    f"**Скидка −{_new_discount}%:** −{_disc_amount:,.2f} ₽  \n"
                    f"**Итого к оплате:** {_preview_final:,.2f} ₽".replace(",", " "))
            else:
                st.info(f"💰 **Сумма КП:** {_preview_base:,.2f} ₽".replace(",", " "))

        _save_col, _contract_col = st.columns([1, 1])
        with _save_col:
            if st.button("💾 Сохранить все изменения",
                         type="primary", use_container_width=True,
                         key=f"crm_edit_items_save_{quote_id}"):
                # Конвертируем df в items для API
                _new_items = []
                for _, _row in _edited.iterrows():
                    _name = str(_row.get("Наименование", "")).strip()
                    if not _name:
                        continue
                    _new_items.append({
                        "code": str(_row.get("Код", "") or "").strip(),
                        "name": _name,
                        "unit": str(_row.get("Ед.", "шт") or "шт").strip() or "шт",
                        "qty": float(_row.get("Кол-во", 0) or 0),
                        "price": float(_row.get("Цена, ₽", 0) or 0),
                    })
                try:
                    # 1. Сохраняем позиции (обновит base_total и final_total по текущей скидке)
                    _res = crm_db.update_quote_items(quote_id, _new_items)
                    # 2. Если скидка поменялась — обновляем
                    if abs(_new_discount - _current_discount) > 1e-6:
                        crm_db.update_quote_discount(quote_id, float(_new_discount))
                    # Сбрасываем буфер — чтобы при rerun подтянулись свежие данные из БД
                    st.session_state.pop(_buffer_key, None)
                    # Получаем актуальный final_total после всех операций
                    _q_fresh = crm_db.get_quote(quote_id)
                    st.success(
                        f"✅ Сохранено! {len(_new_items)} позиций, сумма КП: "
                        f"{float(_q_fresh.get('final_total', 0)):,.2f} ₽"
                        f"{' (со скидкой −' + str(_new_discount) + '%)' if _new_discount > 0 else ''}"
                        .replace(",", " "))
                    st.rerun()
                except Exception as _e:
                    st.error(f"Ошибка сохранения: {_e}")
        with _contract_col:
            # Кнопка перехода к формированию договора по актуальным данным КП
            if st.button("📝 Сформировать договор по этому КП",
                         use_container_width=True,
                         key=f"crm_edit_gen_contract_{quote_id}",
                         help="Перейти к вкладке договора с подгруженными текущими позициями (можно откатить к Excel в Word)"):
                # Передаём в session_state актуальные данные для вкладки договора
                st.session_state["crm_contract_from_quote_id"] = quote_id
                st.info(
                    "⚠️ Сначала сохрани изменения, затем перейди на вкладку «💼 Расчёт КП» → сверху «📄 Открыть в редакторе», "
                    "затем внизу — кнопка «📄 Сформировать Договор поставки». Файл берётся по актуальным данным из БД.")

    # Превью PDF внутри страницы (если есть)
    if _pdf_bytes:
        st.markdown("---")
        st.markdown("### 👁 Просмотр PDF (превью)")
        import base64 as _b64
        _b64_pdf = _b64.b64encode(_pdf_bytes).decode("utf-8")
        _pdf_display = (
            f'<iframe src="data:application/pdf;base64,{_b64_pdf}" '
            f'width="100%" height="800" type="application/pdf"></iframe>'
        )
        st.markdown(_pdf_display, unsafe_allow_html=True)

    # Кнопка назад внизу тоже
    st.markdown("---")
    if st.button("← Назад к списку КП", key="crm_view_back_bottom",
                 use_container_width=True):
        st.session_state.pop("crm_viewing_kp_id", None)
        st.rerun()


def render_quotes_tab():
    """Таблица КП с редактируемыми колонками для быстрой навигации CRM."""
    st.subheader("📄 Выставленные КП")

    # === РЕЖИМ ПРОСМОТРА ОДНОГО КП ===
    _viewing_id = st.session_state.get("crm_viewing_kp_id")
    if _viewing_id:
        _render_kp_view_card(int(_viewing_id))
        return

    # --- Фильтры ---
    _fc1, _fc2, _fc3 = st.columns([1, 1, 2])
    with _fc1:
        status_filter = st.selectbox(
            "Статус",
            ["Все"] + list(STATUS_LABELS.values()),
            key="crm_quotes_status_filter")
    with _fc2:
        product_filter = st.selectbox(
            "Тип товара",
            ["Все", "Кран", "Траверса", "Кран + траверса", "Внешний договор"],
            key="crm_quotes_product_filter")
    with _fc3:
        search_query = st.text_input(
            "🔍 Поиск (заказчик · ФИО · телефон · № КП · товар)",
            key="crm_quotes_search",
            placeholder="Ромашка / +7999 / ЛКС73 / 22072026")

    status_key = None
    if status_filter != "Все":
        status_key = next(k for k, v in STATUS_LABELS.items() if v == status_filter)

    quotes = crm_db.list_quotes(status=status_key)

    if product_filter != "Все":
        quotes = [q for q in quotes if q.get("product_type") == product_filter]

    if search_query and search_query.strip():
        import re as _re_ph
        _sq = search_query.strip().lower()
        _sq_digits = _re_ph.sub(r"\D", "", _sq)
        def _match(q_row):
            fields = [
                str(q_row.get("kp_number", "") or ""),
                str(q_row.get("name_short", "") or ""),
                str(q_row.get("product_model", "") or ""),
                str(q_row.get("product_type", "") or ""),
                str(q_row.get("contact_fio", "") or ""),
                str(q_row.get("email", "") or ""),
                str(q_row.get("request_summary", "") or ""),
            ]
            if any(_sq in f.lower() for f in fields):
                return True
            phone_digits = _re_ph.sub(r"\D", "", str(q_row.get("phone", "") or ""))
            if _sq_digits and len(_sq_digits) >= 3 and _sq_digits in phone_digits:
                return True
            return False
        quotes = [q for q in quotes if _match(q)]

    if not quotes:
        st.info("КП не найдены. Попробуйте убрать фильтры или создать первое КП.")
        return

    # --- Сводные метрики (градиентные карточки) ---
    _total_sum = sum(float(q.get("final_total") or 0) for q in quotes)
    _sold = [q for q in quotes if q.get("status") == "sold"]
    _sold_sum = sum(float(q.get("final_total") or 0) for q in _sold)
    _rate = (len(_sold) / len(quotes) * 100) if quotes else 0

    _mc1, _mc2, _mc3, _mc4 = st.columns(4)
    with _mc1:
        st.markdown(f'''<div class="gradient-card-purple">
            <div class="gradient-card-label">📄 Всего КП</div>
            <div class="gradient-card-value">{len(quotes)}</div>
        </div>''', unsafe_allow_html=True)
    with _mc2:
        st.markdown(f'''<div class="gradient-card-teal">
            <div class="gradient-card-label">💰 Общая сумма</div>
            <div class="gradient-card-value">{_money(_total_sum)}</div>
        </div>''', unsafe_allow_html=True)
    with _mc3:
        st.markdown(f'''<div class="gradient-card-blue">
            <div class="gradient-card-label">✅ Продано</div>
            <div class="gradient-card-value">{len(_sold)} <span style="font-size:1rem;opacity:0.85;">({_rate:.0f}%)</span></div>
        </div>''', unsafe_allow_html=True)
    with _mc4:
        st.markdown(f'''<div class="gradient-card-orange">
            <div class="gradient-card-label">💵 Сумма продаж</div>
            <div class="gradient-card-value">{_money(_sold_sum)}</div>
        </div>''', unsafe_allow_html=True)

    # --- Экспорт ---
    _ec1, _ec2 = st.columns([1, 5])
    with _ec1:
        xlsx_bytes = crm_export.export_quotes_xlsx(quotes)
        st.download_button(
            "⬇️ Excel", data=xlsx_bytes, file_name="quotes_export.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True, key="crm_quotes_xlsx_btn")

    st.markdown("---")

    # --- Редактируемая таблица ---
    st.caption("💡 Отредактируйте ФИО, телефон, email, статус, дату обратной связи прямо в таблице. "
               "Отметьте 🗑️ и нажмите «Сохранить» чтобы удалить КП.")

    import pandas as _pd
    _status_options = list(STATUS_LABELS.values())
    _rows = []
    for _q in quotes:
        _status_label = STATUS_LABELS.get(_q.get("status", "draft"), "Черновик")
        _created = str(_q.get("created_at", "") or "")[:10]
        _last_contact = str(_q.get("last_contact_at", "") or "")[:10]
        _rows.append({
            "👁": False,
            "🗑": False,
            "id": int(_q["id"]),
            "№ КП": _q.get("kp_number", ""),
            "Заказчик": _q.get("name_short") or "—",
            "ФИО контакта": _q.get("contact_fio") or "",
            "Телефон": _q.get("phone") or "",
            "Email": _q.get("email") or "",
            "Что запрашивал": (_q.get("request_summary") or
                              f"{_q.get('product_type','')} · {_q.get('product_model','')}").strip(" ·"),
            "Сумма ₽": float(_q.get("final_total") or 0),
            "Статус": _status_label,
            "Дата КП": _created,
            "Обратная связь": _last_contact,
        })
    _df = _pd.DataFrame(_rows)

    _edited = st.data_editor(
        _df,
        use_container_width=True,
        hide_index=True,
        key="crm_quotes_table",
        column_config={
            "👁": st.column_config.CheckboxColumn("👁", help="Открыть КП (поставьте галочку и нажмите «👁 Открыть» ниже)",
                                                 default=False, width="small"),
            "🗑": st.column_config.CheckboxColumn("🗑", help="Удалить", default=False, width="small"),
            "id": None,   # скрыть
            "№ КП": st.column_config.TextColumn("№ КП", disabled=True, width="small"),
            "Заказчик": st.column_config.TextColumn("Заказчик", disabled=True, width="medium"),
            "ФИО контакта": st.column_config.TextColumn("ФИО контакта", width="medium"),
            "Телефон": st.column_config.TextColumn("Телефон", width="small"),
            "Email": st.column_config.TextColumn("Email", width="medium"),
            "Что запрашивал": st.column_config.TextColumn("Что запрашивал", width="large"),
            "Сумма ₽": st.column_config.NumberColumn("Сумма ₽", format="%.2f ₽", disabled=True),
            "Статус": st.column_config.SelectboxColumn("Статус", options=_status_options),
            "Дата КП": st.column_config.TextColumn("Дата КП", disabled=True, width="small"),
            "Обратная связь": st.column_config.TextColumn("Обратная связь", help="Дата ГГГГ-ММ-ДД", width="small"),
        },
        disabled=["id", "№ КП", "Заказчик", "Сумма ₽", "Дата КП"],
    )

    # Подсчёт отмеченных для удаления
    _to_del_ids = [int(_row["id"]) for _, _row in _edited.iterrows() if bool(_row.get("🗑"))]
    _to_open_ids = [int(_row["id"]) for _, _row in _edited.iterrows() if bool(_row.get("👁"))]

    st.caption(f"💡 Отмечено к открытию: **{len(_to_open_ids)}** · отмечено к удалению: **{len(_to_del_ids)}**")

    _bc0, _bc1, _bc_del, _bc2, _ = st.columns([1, 1, 1.2, 1, 1.8])
    with _bc0:
        if st.button("👁 Открыть", type="secondary", use_container_width=True,
                     key="crm_quotes_open_btn",
                     help="Поставьте галочку в колонке 👁 напротив КП и нажмите — откроется карточка с PDF/редактором"):
            if _to_open_ids:
                st.session_state["crm_viewing_kp_id"] = _to_open_ids[0]
                st.rerun()
            else:
                st.warning("Отметьте галочкой 👁 хотя бы одно КП в таблице выше, потом нажмите «Открыть».")
    with _bc_del:
        # Отдельная красная кнопка удаления пачкой
        _del_confirm_key = "crm_quotes_del_pack_confirm"
        _del_flag = st.session_state.get(_del_confirm_key, False)
        if not _del_flag:
            _btn_label = f"🗑 Удалить ({len(_to_del_ids)})" if _to_del_ids else "🗑 Удалить"
            if st.button(_btn_label, type="secondary", use_container_width=True,
                         key="crm_quotes_del_pack_btn",
                         help="Отметьте галочками 🗑 одно или несколько КП и нажмите — удалятся все отмеченные"):
                if _to_del_ids:
                    st.session_state[_del_confirm_key] = True
                    st.rerun()
                else:
                    st.warning("Отметьте галочкой 🗑 хотя бы одно КП в таблице выше.")
        else:
            if st.button(f"✖ Да, удалить {len(_to_del_ids)} КП", type="primary",
                         use_container_width=True, key="crm_quotes_del_pack_go"):
                user = _current_user()
                _n = 0
                for _qid in _to_del_ids:
                    try:
                        crm_db.delete_quote(_qid, user.get("id"), user.get("username", ""))
                        _n += 1
                    except Exception:
                        pass
                st.session_state.pop(_del_confirm_key, None)
                st.success(f"✅ Удалено КП: **{_n}**")
                st.rerun()
    with _bc1:
        if st.button("💾 Сохранить изменения", type="primary", use_container_width=True,
                     key="crm_quotes_save_all"):
            user = _current_user()
            _updated_count = 0
            _deleted_count = 0
            _status_key_by_label = {v: k for k, v in STATUS_LABELS.items()}
            for _, _row in _edited.iterrows():
                _qid = int(_row["id"])
                if bool(_row.get("🗑")):
                    crm_db.delete_quote(_qid, user.get("id"), user.get("username", ""))
                    _deleted_count += 1
                    continue
                # Сравниваем с исходным — только реально изменившиеся поля
                _orig = next((r for r in _rows if r["id"] == _qid), None)
                if not _orig:
                    continue
                # Контактная информация
                _cf = str(_row.get("ФИО контакта") or "").strip() or None
                _ph = str(_row.get("Телефон") or "").strip() or None
                _em = str(_row.get("Email") or "").strip() or None
                _rs = str(_row.get("Что запрашивал") or "").strip() or None
                _lc = str(_row.get("Обратная связь") or "").strip() or None
                _cf_arg = _cf if _cf != (_orig["ФИО контакта"] or None) else None
                _ph_arg = _ph if _ph != (_orig["Телефон"] or None) else None
                _em_arg = _em if _em != (_orig["Email"] or None) else None
                _rs_arg = _rs if _rs != (_orig["Что запрашивал"] or None) else None
                _lc_arg = _lc if _lc != (_orig["Обратная связь"] or None) else None
                if any(x is not None for x in (_cf_arg, _lc_arg, _rs_arg, _ph_arg, _em_arg)):
                    crm_db.update_quote_contact_info(
                        _qid,
                        contact_fio=_cf_arg,
                        last_contact_at=_lc_arg,
                        request_summary=_rs_arg,
                        customer_phone=_ph_arg,
                        customer_email=_em_arg,
                    )
                    _updated_count += 1
                # Статус
                _new_status = str(_row.get("Статус") or "").strip()
                _orig_status = _orig["Статус"]
                if _new_status and _new_status != _orig_status:
                    _sk = _status_key_by_label.get(_new_status)
                    if _sk:
                        crm_db.set_quote_status(
                            _qid, _sk, "", user.get("id"), user.get("username", ""))
                        _updated_count += 1
            _msg_parts = []
            if _updated_count:
                _msg_parts.append(f"Обновлено: {_updated_count}")
            if _deleted_count:
                _msg_parts.append(f"Удалено: {_deleted_count}")
            if _msg_parts:
                st.success(" · ".join(_msg_parts))
                st.rerun()
            else:
                st.info("Изменений нет.")
    with _bc2:
        if st.button("🔄 Обновить", use_container_width=True, key="crm_quotes_refresh_all"):
            st.rerun()

    # Предупреждение о подтверждении удаления пачкой (после всех колонок)
    if st.session_state.get("crm_quotes_del_pack_confirm", False):
        st.warning(f"⚠️ **Внимание:** будет удалено **{len(_to_del_ids)} КП** из CRM. Отменить нельзя!")
        if st.button("↶ Отмена удаления", key="crm_quotes_del_pack_cancel"):
            st.session_state.pop("crm_quotes_del_pack_confirm", None)
            st.rerun()

    st.markdown("---")
    st.markdown("### 🔎 Детальный просмотр / редактирование одного КП")
    _kp_options = {f"{_q['kp_number']} · {_q.get('name_short') or '—'} · {_money(_q.get('final_total') or 0)}": _q for _q in quotes}
    _selected_label = st.selectbox(
        "Выберите КП для детального просмотра", ["—"] + list(_kp_options.keys()),
        key="crm_quotes_detail_select")
    if _selected_label != "—":
        _q_selected = _kp_options[_selected_label]
        _render_single_quote_details(_q_selected)


def _render_single_quote_details(q: dict):
    """Детальная карточка одного КП: редактор позиций, история, договор, файлы, статус."""
    st.markdown(f"#### {q.get('kp_number')} · {q.get('name_short') or 'без клиента'}")
    _c1, _c2, _c3 = st.columns(3)
    with _c1:
        st.write(f"**Товар:** {q.get('product_model') or '—'}")
        st.write(f"**Сумма:** {_money(q.get('final_total') or 0)}")
        st.write(f"**Статус:** {STATUS_LABELS.get(q.get('status', 'draft'), 'Черновик')}")
    with _c2:
        st.write(f"**Телефон:** {q.get('phone') or '—'}")
        st.write(f"**Email:** {q.get('email') or '—'}")
        st.write(f"**ФИО контакта:** {q.get('contact_fio') or '—'}")
    with _c3:
        st.write(f"**Дата КП:** {str(q.get('created_at') or '')[:10]}")
        st.write(f"**Обратная связь:** {str(q.get('last_contact_at') or '')[:10] or '—'}")
        st.write(f"**Скидка:** {float(q.get('discount_pct') or 0):.1f}%")

    # Вызываем legacy-рендер (редактор позиций, договор, история, скачивание PDF/DOCX)
    _render_single_quote_full_editor(q)


def _render_single_quote_full_editor(q: dict):
    """Развёрнутая карточка (использует старый код render_quotes_tab тело after expander)."""
    c1, c2, c3 = st.columns(3)
    with c1:
        st.write(f"Товар: **{q['product_model']}**")
        st.write(f"Монтаж: {'да' if q['include_montage'] else 'нет'}")
        st.write(f"Город доставки: {q['delivery_city'] or '—'}")
        st.write(f"Менеджер: {q.get('owner_name') or '—'}")
    with c2:
        st.write(f"Телефон: {q.get('phone') or '—'}")
        st.write(f"Email: {q.get('email') or '—'}")
        st.write(f"Создан: {q['created_at']}")
        if q["status"] == STATUS_LOST and q.get("loss_reason"):
            st.write(f"Причина отказа: {q['loss_reason']}")
    with c3:
        st.write(f"Сумма без скидки: {_money(q['base_total'])}")
        st.write(f"Скидка: {q['discount_pct']:.0f}% (−{_money(q['discount_amount'])})")
        st.write(f"**Итого: {_money(q['final_total'])}**")

    new_discount = st.slider(
        "Изменить скидку, %", float(DISCOUNT_MIN_PCT), float(DISCOUNT_MAX_PCT),
        float(q["discount_pct"]), 1.0, key=f"crm_edit_discount_{q['id']}",
    )
    colA, colB, colC = st.columns(3)
    with colA:
        if st.button("Обновить скидку", key=f"crm_update_discount_btn_{q['id']}"):
            user = _current_user()
            crm_db.update_quote_discount(q["id"], new_discount, user.get("id"), user.get("username", ""))
            st.success("Скидка обновлена.")
            st.rerun()
    with colB:
        new_status = st.selectbox(
            "Статус", list(STATUS_LABELS.values()),
            index=list(STATUS_LABELS.keys()).index(q["status"]),
            key=f"crm_status_select_{q['id']}",
        )
    with colC:
        loss_reason_input = ""
        status_key2 = next(k for k, v in STATUS_LABELS.items() if v == new_status)
        if status_key2 == STATUS_LOST:
            loss_reason_input = st.selectbox("Причина отказа", LOSS_REASONS, key=f"crm_loss_reason_{q['id']}")
        if st.button("Сохранить статус", key=f"crm_save_status_btn_{q['id']}"):
            user = _current_user()
            crm_db.set_quote_status(q["id"], status_key2, loss_reason_input, user.get("id"), user.get("username", ""))
            if user.get("telegram_chat_id"):
                crm_notify.notify_status_change(user["telegram_chat_id"], q["kp_number"], new_status)
            st.success(f"Статус обновлён: {new_status}")
            st.rerun()

    # --- Редактор позиций КП ---
    with st.expander("✏️ Редактировать позиции КП", expanded=False):
        _full = crm_db.get_quote(q["id"])
        _items = _full.get("items", []) if _full else []
        if _items:
            import pandas as _pd
            st.caption("💡 Чтобы **удалить** позицию — поставьте галочку в колонке «Удалить» и нажмите «Сохранить». "
                      "Чтобы добавить — введите данные в последнюю пустую строку.")
            _df_items = _pd.DataFrame([
                {"Удалить": False,
                 "Код": i.get("code", "") or "",
                 "Наименование": i.get("name", "") or "",
                 "Ед.": i.get("unit", "шт") or "шт",
                 "Кол-во": float(i.get("qty") or 0),
                 "Цена": float(i.get("price") or 0)}
                for i in _items
            ])
            # Нормализуем код (убираем переносы)
            if "Код" in _df_items.columns:
                _df_items["Код"] = _df_items["Код"].astype(str).str.replace(
                    r"[\s\n\r]+", " ", regex=True).str.strip()

            # Превью с полными артикулами (если есть длинные)
            _long_codes_crm = [str(i.get("code") or "") for i in _items
                               if len(str(i.get("code") or "")) > 20]
            if _long_codes_crm:
                with st.expander("👁 Полный список артикулов (без обрезки)", expanded=False):
                    for i in _items:
                        _c = str(i.get("code") or "").strip()
                        _n = str(i.get("name") or "").strip()
                        if not _c and not _n:
                            continue
                        st.markdown(f"**код:** `{_c}`  —  **{_n}**")

            _edited = st.data_editor(
                _df_items, num_rows="dynamic",
                use_container_width=True, hide_index=True,
                key=f"crm_items_editor_{q['id']}",
                column_config={
                    "Удалить": st.column_config.CheckboxColumn(
                        "Удалить", help="Отметьте чтобы удалить эту строку", default=False, width="small"),
                    "Код": st.column_config.TextColumn("Код / Артикул", width="large",
                        help="Полный артикул (если длинный — расширьте колонку мышами)"),
                    "Наименование": st.column_config.TextColumn("Наименование", width="large"),
                    "Ед.": st.column_config.TextColumn("Ед.", width="small"),
                    "Кол-во": st.column_config.NumberColumn(format="%.1f", min_value=0.0, width="small"),
                    "Цена": st.column_config.NumberColumn(format="%.2f", min_value=0.0, width="medium"),
                },
                height=min(600, 60 + len(_df_items) * 45),
            )
            if st.button("💾 Сохранить изменения позиций",
                        key=f"crm_save_items_{q['id']}", type="primary"):
                # Безопасное преобразование (переживает None/NaN/строки)
                def _safe_num(v):
                    if v is None: return 0.0
                    try:
                        f = float(v)
                        return 0.0 if f != f else f
                    except Exception:
                        return 0.0

                _new_items = []
                _removed_count = 0
                for _, r in _edited.iterrows():
                    _code = str(r.get("Код") or "").strip()
                    _name = str(r.get("Наименование") or "").strip()
                    # Пропускаем отмеченные к удалению
                    if bool(r.get("Удалить")):
                        _removed_count += 1
                        continue
                    # Пропускаем пустые строки
                    if not _name and not _code:
                        continue
                    _new_items.append({
                        "code": _code,
                        "name": _name,
                        "unit": str(r.get("Ед.") or "шт").strip() or "шт",
                        "qty": _safe_num(r.get("Кол-во")),
                        "price": _safe_num(r.get("Цена")),
                    })
                user = _current_user()
                try:
                    res = crm_db.update_quote_items(
                        q["id"], _new_items,
                        user.get("id"), user.get("username", ""))
                    _msg = f"Сохранено позиций: {len(_new_items)}"
                    if _removed_count:
                        _msg += f" (удалено: {_removed_count})"
                    _msg += f". Сумма: {_money(res['final_total'])}"
                    st.success(_msg)
                    st.rerun()
                except Exception as _e:
                    st.error(f"Не удалось сохранить: {_e}")
        else:
            st.info("У этого КП нет сохранённых позиций.")

    # --- 📄 Файлы КП: скачать / перегенерировать ---
    st.markdown("#### 📄 Файлы КП")
    _pdf_path = q.get("pdf_path")
    _docx_path = q.get("docx_path")
    _files_col1, _files_col2, _files_col3 = st.columns(3)

    with _files_col1:
        if _pdf_path and _os.path.exists(_pdf_path):
            try:
                with open(_pdf_path, "rb") as _f:
                    _pdf_data = _f.read()
                st.download_button(
                    "⬇️ Скачать PDF", data=_pdf_data,
                    file_name=f"КП_{q['kp_number'].replace('/','_')}.pdf",
                    mime="application/pdf",
                    key=f"dl_saved_pdf_{q['id']}",
                    use_container_width=True)
            except Exception:
                st.caption("❌ PDF-файл повреждён")
        else:
            st.caption("📭PDF не сохранён")

    with _files_col2:
        if _docx_path and _os.path.exists(_docx_path):
            try:
                with open(_docx_path, "rb") as _f:
                    _docx_data = _f.read()
                st.download_button(
                    "⬇️ Скачать DOCX", data=_docx_data,
                    file_name=f"КП_{q['kp_number'].replace('/','_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dl_saved_docx_{q['id']}",
                    use_container_width=True)
            except Exception:
                st.caption("❌ DOCX-файл повреждён")
        else:
            st.caption("📭DOCX не сохранён")

    with _files_col3:
        if st.button("🔄 Перегенерировать DOCX",
                     key=f"regen_files_{q['id']}",
                     use_container_width=True,
                     help="Пересобрать DOCX-файл КП по текущим позициям. "
                          "Работает даже для старых КП, у которых файлы не были сохранены."):
            try:
                import sys as _sys
                _app_mod = _sys.modules.get("__main__") or _sys.modules.get("app")
                if _app_mod is None:
                    st.error("Модуль приложения не найден.")
                    st.stop()
                _full = crm_db.get_quote(q["id"])
                _items = _full.get("items", []) if _full else []
                if not _items:
                    st.error("Нет позиций — нечего генерировать.")
                else:
                    # Собираем DOCX через build_simple_kp_docx — работает для любого списка позиций
                    _regen_docx = _app_mod.build_simple_kp_docx(
                        _items,
                        kp_number=q.get("kp_number", ""),
                        buyer_name=q.get("name_short") or "",
                        kp_date=str(q.get("created_at") or "")[:10],
                        comment=str(q.get("notes") or "").strip(),
                    )
                    # Если пути не было — создаём
                    _new_docx_path = _docx_path
                    if not _new_docx_path:
                        _new_docx_path = str(crm_db.KP_FILES_DIR / f"kp_{q['id']}.docx")
                    with open(_new_docx_path, "wb") as _f:
                        _f.write(_regen_docx)
                    # Обновляем путь в БД
                    if _new_docx_path != _docx_path:
                        _conn = crm_db.get_conn()
                        _conn.execute(
                            "UPDATE quotes SET docx_path=? WHERE id=?",
                            (_new_docx_path, q["id"]))
                        _conn.commit()
                        _conn.close()
                    st.success(f"✅ DOCX перегенерирован ({len(_regen_docx)} байт). "
                              "Обновите страницу, чтобы появилась кнопка скачивания.")
                    st.rerun()
            except Exception as _e:
                import traceback as _tb
                st.error(f"Ошибка: {_e}")
                st.code(_tb.format_exc())

    # --- Создать ещё КП для этого клиента ---
    if q.get("customer_id"):
        _newkp_col1, _newkp_col2 = st.columns([1, 2])
        with _newkp_col1:
            if st.button("➕ Ещё КП для этого клиента",
                        key=f"crm_new_kp_for_cust_{q['id']}", type="secondary"):
                # Подгружаем клиента в crm_autofilled — поля КП будут заполнены автоматом
                _cust = crm_db.get_customer(q["customer_id"])
                if _cust:
                    st.session_state["crm_autofilled"] = _cust
                    st.session_state["kp_buyer_name"] = _cust.name_short
                    st.session_state["kp_buyer_phone"] = _cust.phone or ""
                    # Переключаемся на вкладку «Расчёт КП»
                    st.session_state["_switch_to_kp_calc"] = True
                    st.success(
                        f"Клиент «{_cust.name_short}» подгружен. "
                        f"Перейдите на вкладку «💼 Расчёт КП» и выберите оборудование. "
                        f"Номер КП присвоится новый автоматически."
                    )
        with _newkp_col2:
            # Список всех КП для этого клиента
            _cust_quotes = crm_db.list_quotes(customer_id=q["customer_id"])
            if len(_cust_quotes) > 1:
                st.caption(
                    f"📄 Всего КП для этого клиента: **{len(_cust_quotes)}** — "
                    + ", ".join(f"{_qq['kp_number']} ({_money(_qq['final_total'])})"
                                for _qq in _cust_quotes[:5])
                )

    # --- Сформировать договор по этому КП ---
    with st.expander("📄 Сформировать договор по этому КП", expanded=False):
        _dog_c1, _dog_c2, _dog_c3 = st.columns([1, 1, 1])
        with _dog_c1:
            _dog_prepay = st.selectbox(
                "% предоплаты",
                [30, 50, 70, 80, 100],
                index=2, key=f"dog_prepay_{q['id']}")
        with _dog_c2:
            _dog_vat = st.selectbox(
                "НДС",
                ["С НДС 22 %", "Без НДС"],
                index=0, key=f"dog_vat_{q['id']}")
        with _dog_c3:
            _dog_delivery = st.selectbox(
                "Доставка",
                ["Самовывоз", "Доставка"],
                index=0, key=f"dog_delivery_{q['id']}")

        _dog_c4, _dog_c5, _dog_c6 = st.columns([1, 1, 1])
        with _dog_c4:
            _dog_shipterm = st.number_input(
                "Срок изготовления, дней",
                min_value=1, max_value=180, value=20,
                key=f"dog_shipterm_{q['id']}")
        with _dog_c5:
            _dog_warr = st.number_input(
                "Гарантия, мес.",
                min_value=1, max_value=60, value=12,
                key=f"dog_warr_{q['id']}")
        with _dog_c6:
            _dog_stamp = st.checkbox(
                "Печать и подпись",
                value=True, key=f"dog_stamp_{q['id']}")

        _dog_addr = st.text_input(
            "Адрес доставки / самовывоза",
            value=q.get("delivery_city", "") or "",
            key=f"dog_addr_{q['id']}")
        _dog_delcost = st.number_input(
            "Стоимость доставки, ₽ (0 = бесплатно/самовывоз)",
            min_value=0.0, value=0.0, step=1000.0,
            key=f"dog_delcost_{q['id']}",
            disabled=(_dog_delivery == "Самовывоз"))

        # --- Редактор спецификации договора ---
        st.markdown("---")
        st.markdown("**📋 Спецификация договора** — можно поменять цену/количество и добавить позиции")
        _dog_full = crm_db.get_quote(q["id"])
        _base_items = _dog_full.get("items", []) if _dog_full else []
        # Исходный список + пара пустых строк для добавления новых
        import pandas as _pd_d
        _spec_key = f"dog_spec_editor_{q['id']}"
        _spec_init = st.session_state.get(_spec_key + "_data")
        if _spec_init is None:
            _spec_init = [
                {"Код": i.get("code", ""),
                 "Наименование": i.get("name", ""),
                 "Ед.": i.get("unit", "шт"),
                 "Кол-во": float(i.get("qty") or 0),
                 "Цена, ₽": float(i.get("price") or 0)}
                for i in _base_items
            ]
        _df_spec = _pd_d.DataFrame(_spec_init) if _spec_init else _pd_d.DataFrame(
            columns=["Код", "Наименование", "Ед.", "Кол-во", "Цена, ₽"])
        _spec_edited = st.data_editor(
            _df_spec,
            num_rows="dynamic",  # ← можно добавлять строки
            use_container_width=True, hide_index=True,
            key=_spec_key,
            column_config={
                "Код": st.column_config.TextColumn("Код", width="small"),
                "Наименование": st.column_config.TextColumn("Наименование", width="large"),
                "Ед.": st.column_config.TextColumn("Ед.", width="small"),
                "Кол-во": st.column_config.NumberColumn(
                    "Кол-во", min_value=0.0, step=1.0, format="%.2f"),
                "Цена, ₽": st.column_config.NumberColumn(
                    "Цена, ₽", min_value=0.0, step=100.0, format="%.2f"),
            },
        )
        # Подсчёт итого
        _spec_total = 0.0
        for _, _r in _spec_edited.iterrows():
            try:
                _spec_total += float(_r.get("Кол-во") or 0) * float(_r.get("Цена, ₽") or 0)
            except Exception:
                pass
        st.caption(f"💰 Сумма спецификации: **{_money(_spec_total)}**")
        if st.button("🔄 Сбросить к позициям КП",
                    key=f"dog_spec_reset_{q['id']}"):
            st.session_state.pop(_spec_key + "_data", None)
            st.session_state.pop(_spec_key, None)
            st.rerun()
        # Сохраняем список для генерации договора
        _dog_edited_items = [
            {"code": str(_r.get("Код", "") or ""),
             "name": str(_r.get("Наименование", "") or ""),
             "unit": str(_r.get("Ед.", "шт") or "шт"),
             "qty": float(_r.get("Кол-во") or 0),
             "price": float(_r.get("Цена, ₽") or 0)}
            for _, _r in _spec_edited.iterrows()
            if str(_r.get("Наименование", "") or "").strip()
        ]

        # --- Приложить чертёж в Приложение № 1 ---
        st.markdown("---")
        _dog_drawing = st.file_uploader(
            "Приложение № 1 — чертёж (PDF / JPG / PNG). Необязательно.",
            type=["pdf", "jpg", "jpeg", "png"],
            key=f"dog_draw_{q['id']}")
        _dog_drawing_caption = st.text_input(
            "Подпись под чертежом",
            value="Габаритный чертёж оборудования",
            key=f"dog_draw_cap_{q['id']}")

        if st.button("📄 Сформировать договор поставки",
                    key=f"dog_gen_{q['id']}",
                    type="primary", use_container_width=True):
            try:
                import external_contract as _ext_dog
                import suppliers as _suppliers
                # Полные данные КП
                _full = crm_db.get_quote(q["id"])
                _items_db = _full.get("items", []) if _full else []
                # Покупатель
                _cust = None
                if q.get("customer_id"):
                    _cust = crm_db.get_customer(q["customer_id"])
                # Поставщик — такой же как в КП (парсим из номера по суффиксу)
                _kp_num = q.get("kp_number", "")
                _sup_key = "LKS"
                if "/МОД" in _kp_num:
                    _sup_key = "MODERNIZATSIYA"
                elif "/КИН" in _kp_num:
                    _sup_key = "KINEMATIKA"
                _supplier = _suppliers.SUPPLIERS.get(
                    _sup_key, _suppliers.SUPPLIERS["LKS"])

                # === ЕДИНАЯ болванка: всегда build_dogovor_docx из app.py ===
                # Строим QuoteData для договора (тот же шаблон что в КП)
                import sys as _sys
                _app_mod = _sys.modules.get("__main__") or _sys.modules.get("app")
                if _app_mod is None:
                    import app as _app_mod

                _lines = [
                    _app_mod.SpecLine(
                        code=str(i.get("code") or ""),
                        name=str(i.get("name") or ""),
                        unit=str(i.get("unit") or "шт"),
                        qty=float(i.get("qty") or 0),
                        price=float(i.get("price") or 0),
                    )
                    for i in _items_db
                ]
                _q_from_kp = _app_mod.QuoteData(
                    series=str(q.get("product_model") or "Товар"),
                    capacity=0, boom=0, height_to_arm=0,
                    hoist_brand="—", hoist_mode="—", hoist_height=0,
                    include_electrification=False,
                    include_montage=bool(q.get("include_montage", False)),
                    montage_price=0.0,
                    include_vat=(_dog_vat == "С НДС 22 %"),
                )
                _q_from_kp.lines = _lines

                _buyer_dict = {
                    "short": (getattr(_cust, "name_short", "") if _cust else "") or "",
                    "full": (getattr(_cust, "name_full", "") if _cust else "") or
                            (getattr(_cust, "name_short", "") if _cust else "") or "",
                    "address": (getattr(_cust, "address", "") if _cust else "") or "",
                    "post_address": (getattr(_cust, "address", "") if _cust else "") or "",
                    "inn": (getattr(_cust, "inn", "") if _cust else "") or "",
                    "kpp": (getattr(_cust, "kpp", "") if _cust else "") or "",
                    "ogrn": (getattr(_cust, "ogrn", "") if _cust else "") or "",
                    "phone": (getattr(_cust, "phone", "") if _cust else "") or "",
                    "email": (getattr(_cust, "email", "") if _cust else "") or "",
                    "bank": (getattr(_cust, "bank", "") if _cust else "") or "",
                    "bik": (getattr(_cust, "bik", "") if _cust else "") or "",
                    "rs": (getattr(_cust, "rs", "") if _cust else "") or "",
                    "ks": (getattr(_cust, "ks", "") if _cust else "") or "",
                    "director_position": (getattr(_cust, "director_position", "") if _cust else "") or "Генеральный директор",
                    "director_fio_gen": (getattr(_cust, "director_fio", "") if _cust else "") or "",
                    "director_fio_short": (getattr(_cust, "director_fio", "") if _cust else "") or "",
                    "basis": "Устава",
                }
                # Ставим выбранного поставщика глобально
                _app_mod.SUPPLIER = _supplier

                _delivery_terms_str = _dog_delivery
                if _dog_addr:
                    _delivery_terms_str = f"{_dog_delivery}, адрес доставки: {_dog_addr}"

                with st.spinner("Формирую договор…"):
                    _dog_docx = _app_mod.build_dogovor_docx(
                        _q_from_kp, _buyer_dict,
                        f"Д-{q.get('kp_number','')}",
                        dt.date.today().strftime("%d.%m.%Y"),
                        prepay_pct=int(_dog_prepay),
                        delivery_terms=_delivery_terms_str,
                        include_stamp=bool(_dog_stamp),
                        shipment_term=f"{int(_dog_shipterm)} рабочих дней со дня поступления оплаты на расчётный счёт Поставщика.",
                        warranty_text=f"Гарантия на все товары — {int(_dog_warr)} месяцев со дня получения Покупателем.",
                    )
                    # PDF — для единого шаблона генерация PDF не реализована — не генерируем
                    _dog_pdf = None
                st.success(f"✅ Договор Н-{q.get('kp_number','')} сформирован")
                st.download_button(
                    "⬇️ Скачать Договор поставки товаров (DOCX)", data=_dog_docx,
                    file_name=f"Договор_КП_{q.get('kp_number','').replace('/','_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    key=f"dog_dl_docx_{q['id']}", use_container_width=True)
                st.info("💡 Для PDF: открой DOCX в Word/Pages → Файл → Сохранить как PDF")
                if not _cust or not (_cust.inn or _cust.name_short):
                    st.warning("⚠️ У КП нет привязанного клиента в CRM — "
                             "в договоре реквизиты Покупателя пусты. Привяжите клиента в вкладке «Клиенты».")
            except Exception as _e:
                import traceback
                st.error(f"Ошибка: {_e}")
                st.code(traceback.format_exc())

    st.markdown("**История изменений**")
    history = crm_db.get_quote_history(q["id"])
    if history:
        for h in history:
            st.caption(f"{h['changed_at']} · {h.get('username') or '—'}: {h['field_changed']} "
                       f"«{h['old_value']}» → «{h['new_value']}»")
    else:
        st.caption("Изменений не зафиксировано.")


# ============================================================
# 6. Напоминания (сводная вкладка по всем клиентам)
# ============================================================

def render_reminders_tab():
    st.subheader("⏰ Напоминания")
    due_now = crm_db.list_due_reminders()
    if due_now:
        st.warning(f"⏰ {len(due_now)} напоминаний требуют внимания сейчас")
        for r in due_now:
            c1, c2 = st.columns([4, 1])
            with c1:
                st.write(f"**{r['due_at']}** — {r['message']} ({r.get('name_short') or '—'}, {r.get('kp_number') or ''})")
            with c2:
                if st.button("Выполнено", key=f"rem_done_{r['id']}"):
                    crm_db.mark_reminder_done(r["id"])
                    st.rerun()

    st.markdown("**Все ожидающие напоминания**")
    all_pending = crm_db.list_all_reminders()
    if all_pending:
        df = pd.DataFrame(all_pending)[["due_at", "message", "name_short", "kp_number"]]
        st.dataframe(df, use_container_width=True, hide_index=True)
    else:
        st.info("Нет активных напоминаний.")


# ============================================================
# 7. Продажи (аналитика + экспорт)
# ============================================================

def render_sales_tab():
    st.subheader("💵 База проданных товаров")
    summary = crm_db.sales_summary()
    c1, c2, c3 = st.columns(3)
    with c1:
        _gradient_metric("Продаж всего", int(summary["cnt"]), icon="📄", color_idx=0)
    with c2:
        _gradient_metric("Сумма продаж", _money(summary["total"]), icon="💰", color_idx=1)
    with c3:
        _gradient_metric("Средняя скидка", f"{summary['avg_discount']:.1f}%", icon="🏷️", color_idx=2)

    df = crm_db.export_sales_to_dataframe()
    if df.empty:
        st.info("Продаж пока нет.")
        return
    st.dataframe(df, use_container_width=True, hide_index=True)

    rows = crm_db.list_sales()
    xlsx_bytes = crm_export.export_sales_xlsx(rows)
    st.download_button(
        "⬇️ Скачать продажи (Excel)", data=xlsx_bytes,
        file_name="sales_export.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


# ============================================================
# 8. Настройки: пользователи, бэкапы, интеграции
# ============================================================

def render_settings_tab():
    import crm_auth
    st.subheader("⚙️ Настройки")

    # --- 🔌 Статус подключения базы данных ---
    st.markdown("### 🔌 Режим базы данных")
    try:
        import db_adapter as _da_status
        _pg_on = _da_status.is_pg_enabled()
        _dsn = _da_status._get_dsn() if _pg_on else None
    except Exception as _e:
        _pg_on = False
        _dsn = None
        st.error(f"Не удалось загрузить db_adapter: {_e}")

    if _pg_on:
        # Пытаемся реально подключиться
        try:
            _test_conn = _da_status.get_pg_connection()
            _cur = _test_conn.execute("SELECT COUNT(*) FROM customers")
            _n_c = _cur.fetchone()[0]
            _cur2 = _test_conn.execute("SELECT COUNT(*) FROM quotes")
            _n_q = _cur2.fetchone()[0]
            _test_conn.close()
            st.success(f"✅ Режим: ☁️ **Supabase PostgreSQL** — данные в облаке, не теряются при деплоях. "
                       f"В базе: **{_n_c}** клиентов, **{_n_q}** КП.")
            # Показываем часть DSN без пароля
            _dsn_safe = _dsn or ""
            if "@" in _dsn_safe:
                _left, _right = _dsn_safe.rsplit("@", 1)
                _dsn_safe = _left.split(":")[0] + ":***@" + _right
            st.caption(f"Подключено: `{_dsn_safe}`")
        except Exception as _e_pg:
            st.error(f"⚠️ **Secret SUPABASE_DB_URL есть, но подключиться не удалось:** {_e_pg}")
            st.warning("💡 Сейчас данные пишутся в локальный SQLite и будут теряться. Проверьте пароль в секрете.")
    else:
        st.error("⚠️ **Режим: 💾 SQLite (локальный)** — данные будут теряться при каждом Reboot Streamlit!")
        st.info(
            "Чтобы данные не терялись:\n"
            "1. Откройте [share.streamlit.io](https://share.streamlit.io) → ваше приложение → ⋮ → **Settings**\n"
            "2. Вкладка **Secrets**\n"
            "3. Добавьте строку:\n"
            "```toml\n"
            'SUPABASE_DB_URL = "postgresql://postgres.xxxxx:PASSWORD@aws-0-eu-west-1.pooler.supabase.com:5432/postgres"\n'
            "```\n"
            "4. **Save** → **Manage app → Reboot app**\n"
            "5. Обновите эту страницу — должно стать ☁️ Supabase."
        )

    st.markdown("---")

    # --- 💾 Бэкап CRM (критично перед каждым коммитом!) ---
    st.markdown("### 💾 Бэкап CRM (клиенты и КП)")
    st.warning("⚠️ **ВАЖНО:** перед каждым коммитом в GitHub скачайте бэкап! "
               "Streamlit Cloud удаляет базу данных при каждом деплое. "
               "Рекомендуется скачивать бэкап каждые несколько дней.")

    _bkp_col1, _bkp_col2 = st.columns([1, 1])

    with _bkp_col1:
        st.markdown("**⬇️ Скачать бэкап**")
        # Собираем все данные CRM в zip: crm.db + crm_files/
        try:
            import io as _io
            import zipfile as _zf
            from pathlib import Path as _Path
            _app_dir = _Path(crm_db.__file__).parent
            _db_path = _app_dir / "crm.db"
            _files_dir = _app_dir / "crm_files"

            _buf = _io.BytesIO()
            with _zf.ZipFile(_buf, "w", _zf.ZIP_DEFLATED) as _zp:
                if _db_path.exists():
                    _zp.write(_db_path, arcname="crm.db")
                if _files_dir.exists() and _files_dir.is_dir():
                    for _f in _files_dir.rglob("*"):
                        if _f.is_file():
                            _zp.write(_f, arcname=f"crm_files/{_f.name}")
                # История реквизитов
                _hist = _app_dir / "history_memory.json"
                if _hist.exists():
                    _zp.write(_hist, arcname="history_memory.json")

            _stamp = dt.datetime.now().strftime("%Y-%m-%d_%H%M")
            st.download_button(
                f"💾 Скачать crm_backup_{_stamp}.zip",
                data=_buf.getvalue(),
                file_name=f"crm_backup_{_stamp}.zip",
                mime="application/zip",
                use_container_width=True,
                key="crm_backup_download_btn",
            )
            # Смотрим что внутри
            _n_quotes = 0
            _n_customers = 0
            try:
                _conn = crm_db.get_conn()
                _n_quotes = _conn.execute("SELECT COUNT(*) FROM quotes").fetchone()[0]
                _n_customers = _conn.execute("SELECT COUNT(*) FROM customers").fetchone()[0]
                _conn.close()
            except Exception:
                pass
            st.caption(f"В бэкапе: **{_n_customers}** клиентов · **{_n_quotes}** КП · размер {len(_buf.getvalue())/1024:.1f} КБ")
        except Exception as e:
            st.error(f"Ошибка сборки бэкапа: {e}")

    with _bkp_col2:
        st.markdown("**⬆️ Восстановить из файла**")
        _restore_file = st.file_uploader(
            "Загрузите бэкап (от любой даты)",
            type=["zip", "db"],
            key="crm_restore_upload",
            help="Заменит текущую базу на данные из бэкапа")
        if _restore_file is not None:
            _restore_col1, _restore_col2 = st.columns([1, 1])
            with _restore_col1:
                if st.button("🗘️ Восстановить", type="primary",
                             key="crm_restore_confirm_btn",
                             use_container_width=True):
                    try:
                        from pathlib import Path as _Path
                        import io as _io
                        import zipfile as _zf
                        _app_dir = _Path(crm_db.__file__).parent
                        _data = _restore_file.read()
                        # Fallback: голый .db-файл (без zip)
                        if _restore_file.name.lower().endswith(".db"):
                            (_app_dir / "crm.db").write_bytes(_data)
                            st.success("✅ База восстановлена из .db-файла.")
                        else:
                            with _zf.ZipFile(_io.BytesIO(_data)) as _zp:
                                for _name in _zp.namelist():
                                    if _name == "crm.db":
                                        (_app_dir / "crm.db").write_bytes(_zp.read(_name))
                                    elif _name.startswith("crm_files/"):
                                        _target = _app_dir / _name
                                        _target.parent.mkdir(parents=True, exist_ok=True)
                                        _target.write_bytes(_zp.read(_name))
                                    elif _name == "history_memory.json":
                                        (_app_dir / "history_memory.json").write_bytes(_zp.read(_name))
                            st.success("✅ CRM восстановлена из бэкапа.")
                        st.info("Обновите страницу или нажмите F5 чтобы увидеть восстановленные КП.")
                        st.balloons()
                    except Exception as e:
                        import traceback
                        st.error(f"Ошибка восстановления: {e}")
                        st.code(traceback.format_exc())

    st.markdown("---")

    # --- Токен DaData ---
    st.markdown("### 🔑 API-ключи")
    _current_token = crm_db._get_dadata_token()
    _masked = (_current_token[:6] + "…" + _current_token[-4:]) if len(_current_token) >= 10 else (_current_token or "не задан")
    st.caption(f"Токен DaData: **{_masked}**")
    _new_tok = st.text_input(
        "Токен DaData (для автоподгрузки реквизитов по ИНН)",
        type="password", key="dadata_token_settings",
        help="Бесплатная регистрация на dadata.ru → Личный кабинет → АПИ → API-ключ. "
             "Лимит 10 000 запросов/сутки бесплатно.")
    _c1, _c2 = st.columns([1, 1])
    with _c1:
        if _new_tok and st.button("💾 Сохранить токен", key="dadata_token_save_settings"):
            st.session_state["dadata_token"] = _new_tok.strip()
            try:
                from pathlib import Path
                Path(__file__).parent.joinpath(".dadata_token").write_text(
                    _new_tok.strip(), encoding="utf-8")
                st.success("Токен сохранён.")
                st.rerun()
            except Exception as e:
                st.error(f"Ошибка сохранения: {e}")
    with _c2:
        if _current_token and st.button("🧪 Проверить токен", key="dadata_token_test"):
            try:
                # Тестовый запрос по ИНН Сбербанка
                r = crm_db.fetch_by_inn_dadata("7707083893")
                if r:
                    st.success(f"✅ Токен рабочий. Нашлось: {r.name_short}")
                else:
                    st.warning("Ответ пуст (но ошибки нет) — токен валиден")
            except crm_db.DaDataError as e:
                st.error(f"✖ {e}")

    st.markdown("---")
    st.markdown("### Пользователи и роли")
    crm_auth.render_users_admin_tab()

    st.markdown("---")
    st.markdown("### Резервное копирование базы")
    if st.button("Создать бэкап сейчас", key="crm_backup_btn"):
        path = crm_backup.create_local_backup()
        st.success(f"Бэкап создан: {path.name}")
    backups = crm_backup.list_backups()
    if backups:
        st.dataframe(pd.DataFrame(backups), use_container_width=True, hide_index=True)
    else:
        st.caption("Бэкапов пока нет.")

    if st.button("Загрузить последний бэкап на Яндекс.Диск", key="crm_yadisk_btn"):
        if backups:
            from pathlib import Path
            res = crm_backup.upload_to_yandex_disk(Path(backups[0]["path"]))
            if res["ok"]:
                st.success("Бэкап загружен на Яндекс.Диск.")
            else:
                st.error(res["error"])
        else:
            st.warning("Сначала создайте локальный бэкап.")

    st.caption(
        "Для email-рассылки и Telegram-уведомлений заполните файл "
        "`.streamlit/secrets.toml` секциями [smtp], [telegram], [yadisk] — см. README_CRM.md."
    )
