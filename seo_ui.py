"""
SEO модуль — Streamlit UI.

Все функции render_*() ожидают запуска в Streamlit-контексте.
Придерживаются стиля crm_ui.py: карточки, метрики, dataframes.
"""
from __future__ import annotations
import datetime as dt
import pandas as pd
import streamlit as st

import seo_db
import crm_db
from seo_utils import (
    extract_domain, fmt_money, normalize_phone,
    PROJECT_STATUSES, PROJECT_STATUS_LABELS,
    TASK_STATUSES, TASK_STATUS_LABELS,
    TASK_PRIORITIES, TASK_PRIORITY_LABELS,
)


# ==================== DASHBOARD ====================

def render_seo_dashboard_tab():
    st.markdown("### 📊 SEO Dashboard")
    seo_db.init_seo_db()
    m = seo_db.dashboard_metrics()

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric('Проекты', m['projects'])
    c2.metric('Активные', m['active_projects'])
    c3.metric('Новые лиды', m['active_leads'])
    c4.metric('Лиды за месяц', m['leads_month'])
    c5.metric('Прогноз выручки/мес', f"{fmt_money(m['forecast_revenue'])} ₽")

    st.divider()

    col_left, col_right = st.columns(2)

    with col_left:
        st.markdown("**📈 Распределение проектов по статусам**")
        if m.get('status_dist'):
            data = [{"Статус": PROJECT_STATUS_LABELS.get(s, s), "Кол-во": n}
                    for s, n in m['status_dist'].items()]
            df = pd.DataFrame(data)
            st.dataframe(df, hide_index=True, use_container_width=True)
        else:
            st.info("Проектов ещё нет")

    with col_right:
        st.markdown("**📬 Последние лиды**")
        leads = seo_db.list_leads(limit=8)
        if leads:
            for lead in leads:
                emoji = "🆕" if not lead.get("is_processed") else "✅"
                name = lead.get('company') or lead.get('name') or 'Без имени'
                st.markdown(f"{emoji} **{name}** · {lead.get('site_url') or '—'}")
                st.caption(f"{lead.get('created_at','')} · {lead.get('phone') or lead.get('email') or ''}")
        else:
            st.info("Лидов пока нет")


# ==================== ЛИДЫ ====================

def render_seo_leads_tab():
    st.markdown("### 📬 SEO Leads")
    seo_db.init_seo_db()

    with st.expander("➕ Добавить лид вручную", expanded=False):
        _render_add_lead_form()

    # Фильтры
    col1, col2, col3 = st.columns(3)
    with col1:
        f_search = st.text_input("🔍 Поиск", key="seo_lead_search",
                                 placeholder="имя / компания / телефон / сайт")
    with col2:
        f_source = st.text_input("Источник", key="seo_lead_filter_source",
                                 placeholder="site, manual, yandex...")
    with col3:
        f_status = st.selectbox("Статус", ["Все", "Только новые", "Только обработанные"],
                                key="seo_lead_filter_status")

    leads = seo_db.list_leads(limit=500)
    if f_search:
        s = f_search.lower()
        leads = [l for l in leads if any(s in str(l.get(k,'') or '').lower()
                 for k in ('name','company','phone','email','site_url'))]
    if f_source:
        leads = [l for l in leads if f_source.lower() in str(l.get('source','') or '').lower()]
    if f_status == "Только новые":
        leads = [l for l in leads if not l.get('is_processed')]
    elif f_status == "Только обработанные":
        leads = [l for l in leads if l.get('is_processed')]

    st.caption(f"Показано лидов: {len(leads)}")
    st.divider()

    if not leads:
        st.info("Лидов по фильтру нет")
        return

    for lead in leads:
        _render_lead_card(lead)


def _render_add_lead_form():
    c1, c2 = st.columns(2)
    with c1:
        name = st.text_input('Имя', key='seo_lead_name_new')
        phone = st.text_input('Телефон', key='seo_lead_phone_new')
        email = st.text_input('Email', key='seo_lead_email_new')
        company = st.text_input('Компания', key='seo_lead_company_new')
    with c2:
        site_url = st.text_input('Сайт', key='seo_lead_site_new')
        source = st.text_input('Источник', value='manual', key='seo_lead_source_new')
        utm_source = st.text_input('utm_source', key='seo_lead_utm_source_new')
        utm_campaign = st.text_input('utm_campaign', key='seo_lead_utm_campaign_new')
    comment = st.text_area('Комментарий', key='seo_lead_comment_new')

    if st.button('💾 Сохранить лид', type='primary', key='seo_save_lead_btn'):
        if not (phone or email):
            st.error("Нужен телефон или email")
        elif not site_url:
            st.error("Нужен сайт")
        else:
            lead_id = seo_db.create_lead(seo_db.SEOLead(
                source=source, name=name, phone=normalize_phone(phone), email=email,
                company=company, site_url=site_url, comment=comment,
                utm_source=utm_source, utm_campaign=utm_campaign,
            ))
            st.success(f'✅ Лид сохранён, ID {lead_id}')
            st.rerun()


def _render_lead_card(lead: dict):
    with st.container(border=True):
        emoji = "🆕" if not lead.get("is_processed") else "✅"
        header = f"{emoji} **{lead.get('company') or lead.get('name') or 'Без имени'}**"
        if lead.get("site_url"):
            header += f" · {lead['site_url']}"
        st.markdown(header)
        st.caption(
            f"ID {lead['id']} · {lead.get('created_at','')} · "
            f"📞 {lead.get('phone','') or '—'} · 📧 {lead.get('email','') or '—'} · "
            f"Источник: {lead.get('source','—')}"
        )
        if lead.get("comment"):
            st.write(lead["comment"])
        # UTM
        utm_parts = [f"{k}: {lead.get(k)}" for k in ('utm_source','utm_medium','utm_campaign') if lead.get(k)]
        if utm_parts:
            st.caption("📊 " + " · ".join(utm_parts))

        c1, c2, c3, c4 = st.columns(4)
        with c1:
            if not lead.get("is_processed"):
                if st.button("✅ Обработано", key=f"seo_done_{lead['id']}", use_container_width=True):
                    seo_db.mark_lead_processed(int(lead['id']))
                    st.rerun()
        with c2:
            if st.button("👤 Создать клиента", key=f"seo_mkcust_{lead['id']}", use_container_width=True):
                _create_customer_from_lead(lead)
                st.rerun()
        with c3:
            if st.button("🎯 Создать SEO-проект", key=f"seo_mkproj_{lead['id']}", use_container_width=True):
                _create_project_from_lead(lead)
                st.rerun()
        with c4:
            if st.button("🗑 Удалить", key=f"seo_del_lead_{lead['id']}", use_container_width=True):
                seo_db.delete_lead(int(lead['id']))
                st.rerun()


def _create_customer_from_lead(lead: dict) -> int:
    """Создать customer из лида и привязать к лиду."""
    from crm_db import upsert_customer, Customer
    cust = Customer(
        name_short=(lead.get('company') or lead.get('name') or '')[:120],
        name_full=(lead.get('company') or lead.get('name') or '')[:250],
        phone=lead.get('phone','') or '',
        email=lead.get('email','') or '',
    )
    cust_id = upsert_customer(cust)
    seo_db.mark_lead_processed(int(lead['id']), customer_id=cust_id)
    st.success(f"✅ Клиент создан (ID {cust_id}) и привязан к лиду")
    return cust_id


def _create_project_from_lead(lead: dict) -> int:
    """Создать SEO-проект из лида (и клиента, если ещё нет)."""
    customer_id = lead.get('customer_id')
    if not customer_id:
        customer_id = _create_customer_from_lead(lead)
    site_url = lead.get('site_url', '')
    project_id = seo_db.create_project(
        customer_id=customer_id, site_url=site_url,
        domain=extract_domain(site_url), status='new',
    )
    seo_db.mark_lead_processed(int(lead['id']), customer_id=customer_id, project_id=project_id)
    st.success(f"✅ SEO-проект создан (ID {project_id})")
    return project_id


# ==================== ПРОЕКТЫ ====================

def render_seo_projects_tab():
    st.markdown("### 🎯 SEO Projects")
    seo_db.init_seo_db()

    # Если открыт проект — показываем карточку
    open_id = st.session_state.get("_seo_project_open_id")
    if open_id:
        _render_project_card(open_id)
        return

    with st.expander("➕ Создать проект", expanded=False):
        _render_add_project_form()

    projects = seo_db.list_projects()
    if not projects:
        st.info("SEO-проектов пока нет. Создайте вручную или из лида.")
        return

    # Фильтры
    c1, c2 = st.columns(2)
    with c1:
        f_search = st.text_input("🔍 Поиск", key="seo_proj_search")
    with c2:
        f_status = st.selectbox("Статус", ["Все"] + PROJECT_STATUSES, key="seo_proj_status_f",
                                format_func=lambda x: x if x == "Все" else PROJECT_STATUS_LABELS.get(x, x))

    filtered = projects
    if f_search:
        s = f_search.lower()
        filtered = [p for p in filtered if any(s in str(p.get(k,'') or '').lower()
                    for k in ('site_url','domain','customer_name','region'))]
    if f_status != "Все":
        filtered = [p for p in filtered if p.get('status') == f_status]

    st.caption(f"Показано проектов: {len(filtered)}")

    for p in filtered:
        with st.container(border=True):
            cols = st.columns([0.5, 2.5, 2, 1.5, 1.2, 1, 0.8])
            cols[0].caption(f"#{p['id']}")
            cols[1].markdown(f"**{p.get('domain') or p.get('site_url','—')}**")
            cols[1].caption(p.get('site_url') or '')
            cols[2].markdown(p.get('customer_name') or '_без клиента_')
            cols[2].caption(f"Регион: {p.get('region') or '—'}")
            cols[3].markdown(PROJECT_STATUS_LABELS.get(p.get('status'), p.get('status','—')))
            cols[4].markdown(f"**{fmt_money(p.get('monthly_fee',0))} ₽/мес**")
            with cols[5]:
                if st.button("📂 Открыть", key=f"open_proj_{p['id']}", use_container_width=True):
                    st.session_state["_seo_project_open_id"] = p['id']
                    st.rerun()
            with cols[6]:
                if st.button("🗑", key=f"del_proj_{p['id']}", use_container_width=True):
                    if st.session_state.get(f"_confirm_del_proj_{p['id']}"):
                        seo_db.delete_project(int(p['id']))
                        st.rerun()
                    else:
                        st.session_state[f"_confirm_del_proj_{p['id']}"] = True
                        st.warning("Ещё раз — удалить")


def _render_add_project_form():
    # Список клиентов для селекта
    from crm_db import list_customers
    try:
        customers = list_customers()
    except Exception:
        customers = []

    c1, c2 = st.columns(2)
    with c1:
        cust_options = ["— без клиента —"] + [f"#{c.id} {c.name_short or c.name_full or c.inn}"
                                              for c in customers]
        cust_choice = st.selectbox("Клиент", cust_options, key="seo_new_proj_cust")
        cust_id = None
        if cust_choice != "— без клиента —":
            try:
                cust_id = int(cust_choice.split()[0][1:])
            except Exception:
                cust_id = None

        site_url = st.text_input('Сайт (URL)', key='seo_new_proj_site',
                                 placeholder="https://example.com")
        domain = st.text_input('Домен', value=extract_domain(site_url) if site_url else '',
                                key='seo_new_proj_domain')
    with c2:
        region = st.text_input('Регион', key='seo_new_proj_region',
                              placeholder="Москва, Санкт-Петербург, РФ...")
        service_type = st.text_input('Услуга', value='SEO автопилот', key='seo_new_proj_service')
        monthly_fee = st.number_input('Абонплата, ₽/мес', min_value=0.0, step=1000.0,
                                      value=0.0, key='seo_new_proj_fee')
    status = st.selectbox('Статус', PROJECT_STATUSES, key='seo_new_proj_status',
                          format_func=lambda x: PROJECT_STATUS_LABELS.get(x,x))
    notes = st.text_area("Заметки", key='seo_new_proj_notes')

    if st.button('💾 Создать проект', type='primary', key='seo_create_project_btn'):
        if not site_url:
            st.error("Нужно указать сайт")
        else:
            pid = seo_db.create_project(
                customer_id=cust_id, site_url=site_url,
                domain=domain or extract_domain(site_url),
                region=region, service_type=service_type,
                monthly_fee=monthly_fee, status=status, notes=notes,
            )
            st.success(f"✅ Проект создан, ID {pid}")
            st.rerun()


def _render_project_card(project_id: int):
    proj = seo_db.get_project(project_id)
    if not proj:
        st.error("Проект не найден")
        if st.button("← К списку"):
            st.session_state["_seo_project_open_id"] = None
            st.rerun()
        return

    col_back, col_title = st.columns([1, 5])
    with col_back:
        if st.button("← К списку", use_container_width=True):
            st.session_state["_seo_project_open_id"] = None
            st.rerun()
    with col_title:
        st.markdown(f"### 🎯 {proj.get('domain') or proj.get('site_url')}")
        st.caption(f"ID {proj['id']} · Клиент: {proj.get('customer_name') or '—'} · "
                   f"ИНН {proj.get('customer_inn') or '—'}")

    st.divider()

    tabs = st.tabs(["📊 Обзор", "✏️ Редактирование", "📋 Задачи", "📈 Отчёты"])

    with tabs[0]:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Статус", PROJECT_STATUS_LABELS.get(proj.get('status'), proj.get('status','—')))
        c2.metric("Абонплата", f"{fmt_money(proj.get('monthly_fee',0))} ₽")
        c3.metric("Регион", proj.get('region') or '—')
        c4.metric("Услуга", proj.get('service_type') or '—')

        if proj.get('notes'):
            st.markdown("**Заметки**")
            st.info(proj['notes'])

        st.caption(f"Создан: {proj.get('created_at','')} · Обновлён: {proj.get('updated_at','')}")

    with tabs[1]:
        c1, c2 = st.columns(2)
        with c1:
            new_site = st.text_input("Сайт", value=proj.get('site_url','') or '', key=f"ep_site_{project_id}")
            new_domain = st.text_input("Домен", value=proj.get('domain','') or '', key=f"ep_domain_{project_id}")
            new_region = st.text_input("Регион", value=proj.get('region','') or '', key=f"ep_region_{project_id}")
        with c2:
            new_service = st.text_input("Услуга", value=proj.get('service_type','') or '', key=f"ep_service_{project_id}")
            new_fee = st.number_input("Абонплата, ₽/мес", value=float(proj.get('monthly_fee',0) or 0),
                                       min_value=0.0, step=1000.0, key=f"ep_fee_{project_id}")
            new_status = st.selectbox("Статус", PROJECT_STATUSES,
                                       index=PROJECT_STATUSES.index(proj.get('status','new'))
                                             if proj.get('status') in PROJECT_STATUSES else 0,
                                       format_func=lambda x: PROJECT_STATUS_LABELS.get(x,x),
                                       key=f"ep_status_{project_id}")
        new_notes = st.text_area("Заметки", value=proj.get('notes','') or '', key=f"ep_notes_{project_id}")

        if st.button("💾 Сохранить", type='primary', key=f"ep_save_{project_id}"):
            seo_db.update_project(project_id,
                site_url=new_site, domain=new_domain, region=new_region,
                service_type=new_service, monthly_fee=float(new_fee),
                status=new_status, notes=new_notes)
            st.success("✅ Сохранено")
            st.rerun()

    with tabs[2]:
        _render_project_tasks(project_id)

    with tabs[3]:
        _render_project_reports(project_id)


# ==================== ЗАДАЧИ ====================

def _render_project_tasks(project_id: int):
    with st.expander("➕ Новая задача", expanded=False):
        title = st.text_input("Название", key=f"nt_title_{project_id}")
        desc = st.text_area("Описание", key=f"nt_desc_{project_id}")
        c1, c2 = st.columns(2)
        with c1:
            priority = st.selectbox("Приоритет", TASK_PRIORITIES,
                index=1, format_func=lambda x: TASK_PRIORITY_LABELS.get(x,x),
                key=f"nt_prio_{project_id}")
        with c2:
            due = st.date_input("Срок", value=None, key=f"nt_due_{project_id}")
        if st.button("💾 Создать задачу", type='primary', key=f"nt_save_{project_id}"):
            if not title:
                st.error("Название обязательно")
            else:
                seo_db.create_task(project_id, title, desc, priority,
                                    str(due) if due else None)
                st.rerun()

    tasks = seo_db.list_tasks(project_id=project_id)
    if not tasks:
        st.info("Задач нет")
        return

    for t in tasks:
        with st.container(border=True):
            c1, c2, c3, c4 = st.columns([3, 1.2, 1.2, 1])
            c1.markdown(f"**{t.get('title')}**")
            if t.get('description'):
                c1.caption(t['description'])
            c2.markdown(TASK_PRIORITY_LABELS.get(t.get('priority','medium'), t.get('priority')))
            with c3:
                new_st = st.selectbox("", TASK_STATUSES,
                    index=TASK_STATUSES.index(t.get('status','open')) if t.get('status') in TASK_STATUSES else 0,
                    format_func=lambda x: TASK_STATUS_LABELS.get(x,x),
                    key=f"tst_{t['id']}", label_visibility="collapsed")
                if new_st != t.get('status'):
                    seo_db.update_task_status(int(t['id']), new_st)
                    st.rerun()
            with c4:
                if st.button("🗑", key=f"del_task_{t['id']}"):
                    seo_db.delete_task(int(t['id']))
                    st.rerun()
            if t.get('due_at'):
                st.caption(f"⏰ Срок: {t['due_at']}")


def render_seo_tasks_tab():
    """Общий раздел «Все задачи»."""
    st.markdown("### 📋 SEO Tasks — все задачи")
    seo_db.init_seo_db()

    c1, c2 = st.columns(2)
    with c1:
        f_status = st.selectbox("Статус", ["Все"] + TASK_STATUSES,
            format_func=lambda x: x if x=="Все" else TASK_STATUS_LABELS.get(x,x),
            key="seo_alltasks_status")
    with c2:
        f_prio = st.selectbox("Приоритет", ["Все"] + TASK_PRIORITIES,
            format_func=lambda x: x if x=="Все" else TASK_PRIORITY_LABELS.get(x,x),
            key="seo_alltasks_prio")

    tasks = seo_db.list_tasks()
    if f_status != "Все":
        tasks = [t for t in tasks if t.get('status') == f_status]
    if f_prio != "Все":
        tasks = [t for t in tasks if t.get('priority') == f_prio]

    st.caption(f"Показано: {len(tasks)} задач")

    if not tasks:
        st.info("Задач нет")
        return

    for t in tasks:
        with st.container(border=True):
            cols = st.columns([2.5, 2, 1.2, 1.2, 1.2, 0.8])
            cols[0].markdown(f"**{t.get('title')}**")
            cols[0].caption(t.get('description','')[:80] if t.get('description') else '')
            cols[1].markdown(f"🌐 {t.get('domain') or t.get('site_url','—')}")
            cols[1].caption(f"Проект #{t.get('project_id')}")
            cols[2].markdown(TASK_PRIORITY_LABELS.get(t.get('priority','medium')))
            cols[3].markdown(TASK_STATUS_LABELS.get(t.get('status','open')))
            cols[4].caption(t.get('due_at','—') or '—')
            with cols[5]:
                if st.button("🗑", key=f"del_atask_{t['id']}"):
                    seo_db.delete_task(int(t['id']))
                    st.rerun()


# ==================== ОТЧЁТЫ ====================

def _render_project_reports(project_id: int):
    with st.expander("➕ Новый отчёт", expanded=False):
        period = st.text_input("Период (YYYY-MM)",
                                value=dt.datetime.now().strftime('%Y-%m'),
                                key=f"nr_period_{project_id}")
        c1, c2, c3 = st.columns(3)
        with c1:
            vis = st.number_input("Видимость (%)", 0.0, 100.0, 0.0, step=1.0, key=f"nr_vis_{project_id}")
            idx = st.number_input("Стр. в индексе", 0, step=1, value=0, key=f"nr_idx_{project_id}")
        with c2:
            ai = st.number_input("Упоминания в AI", 0, step=1, value=0, key=f"nr_ai_{project_id}")
            leads = st.number_input("Лиды", 0, step=1, value=0, key=f"nr_leads_{project_id}")
        with c3:
            conv = st.number_input("Конверсии", 0, step=1, value=0, key=f"nr_conv_{project_id}")
        notes = st.text_area("Заметки", key=f"nr_notes_{project_id}")

        if st.button("💾 Сохранить отчёт", type='primary', key=f"nr_save_{project_id}"):
            seo_db.create_report(project_id, period, vis, idx, ai, leads, conv, notes)
            st.rerun()

    reports = seo_db.list_reports(project_id=project_id)
    if not reports:
        st.info("Отчётов нет")
        return

    df = pd.DataFrame([{
        "Период": r.get('report_period'),
        "Видимость %": r.get('visibility_score', 0),
        "В индексе": r.get('indexed_pages', 0),
        "AI-упом.": r.get('ai_mentions', 0),
        "Лидов": r.get('leads_count', 0),
        "Конверсий": r.get('conversions', 0),
        "Создан": r.get('created_at','')[:10],
    } for r in reports])
    st.dataframe(df, hide_index=True, use_container_width=True)


def render_seo_reports_tab():
    """Общая вкладка «Все отчёты»."""
    st.markdown("### 📈 SEO Reports — все отчёты")
    seo_db.init_seo_db()

    reports = seo_db.list_reports()
    if not reports:
        st.info("Отчётов пока нет. Создайте отчёт в карточке SEO-проекта.")
        return

    df = pd.DataFrame([{
        "ID": r.get('id'),
        "Домен": r.get('domain','—'),
        "Проект #": r.get('project_id'),
        "Период": r.get('report_period'),
        "Видимость %": r.get('visibility_score', 0),
        "В индексе": r.get('indexed_pages', 0),
        "AI": r.get('ai_mentions', 0),
        "Лидов": r.get('leads_count', 0),
        "Конверсий": r.get('conversions', 0),
        "Создан": r.get('created_at','')[:10],
    } for r in reports])
    st.dataframe(df, hide_index=True, use_container_width=True)


# ==================== ИНТЕГРАЦИИ ====================

def render_seo_integrations_tab():
    st.markdown("### 🔌 SEO Integrations")
    seo_db.init_seo_db()

    st.info("Секреты храните через `.streamlit/secrets.toml` или переменные окружения. "
            "В базе — только маскированные значения.")

    with st.expander("Webhook приёма лидов с сайта", expanded=True):
        st.markdown("""
        **URL для формы на сайте SEO Автопилот:**

        ```python
        # Внутренняя функция (Streamlit):
        from seo_api import receive_site_lead
        receive_site_lead({
            "source": "seo-autopilot-site",
            "name": "...", "phone": "...", "email": "...",
            "company": "...", "site_url": "https://client-site.ru",
            "comment": "...", "utm_source": "...", "utm_campaign": "..."
        })
        ```

        **Если нужен HTTP-endpoint** — запусти `seo_api.py` как FastAPI-приложение
        (см. подсказки в файле). Он не мешает Streamlit — работает отдельным процессом.
        """)

    st.divider()

    with st.expander("➕ Добавить/обновить интеграцию", expanded=False):
        name = st.text_input("Название", key="int_name", placeholder="Bitrix24 webhook")
        api_url = st.text_input("URL", key="int_url")
        key_mask = st.text_input("Ключ (маскированный)", key="int_key_mask",
                                  placeholder="****abcd")
        enabled = st.checkbox("Включена", value=False, key="int_enabled")
        if st.button("💾 Сохранить", type='primary', key="int_save"):
            seo_db.save_integration(name, api_url, key_mask, enabled)
            st.rerun()

    integrations = seo_db.list_integrations()
    if not integrations:
        st.info("Интеграций пока нет")
        return

    for i in integrations:
        with st.container(border=True):
            icon = "✅" if i.get("is_enabled") else "⏸"
            st.markdown(f"{icon} **{i.get('name')}**")
            st.caption(f"URL: {i.get('api_url') or '—'} · Ключ: {i.get('api_key_masked') or '—'}")


# ==================== ИНТЕГРАЦИЯ С КАРТОЧКОЙ КЛИЕНТА ====================

def render_customer_seo_projects(customer_id: int):
    """
    Отдельный виджет — вставляется в карточку клиента (crm_ui.py).
    Показывает SEO-проекты клиента и позволяет создать новый.
    """
    if not customer_id:
        return
    seo_db.init_seo_db()

    st.markdown("#### 🎯 SEO-проекты клиента")
    projects = seo_db.list_projects(customer_id=customer_id)

    if projects:
        for p in projects:
            with st.container(border=True):
                cols = st.columns([3, 1.5, 1.5, 1])
                cols[0].markdown(f"**{p.get('domain') or p.get('site_url')}**")
                cols[0].caption(f"ID {p['id']} · {p.get('service_type','—')}")
                cols[1].markdown(PROJECT_STATUS_LABELS.get(p.get('status'), p.get('status')))
                cols[2].markdown(f"{fmt_money(p.get('monthly_fee',0))} ₽/мес")
                with cols[3]:
                    if st.button("📂", key=f"cust_open_seo_{p['id']}"):
                        st.session_state["_seo_project_open_id"] = p['id']
                        st.session_state["_pending_nav"] = "CRM: SEO Projects"
                        st.rerun()
    else:
        st.info("У клиента ещё нет SEO-проектов")

    with st.expander("➕ Создать SEO-проект для этого клиента", expanded=False):
        site = st.text_input("Сайт", key=f"cust_new_seo_site_{customer_id}")
        fee = st.number_input("Абонплата, ₽/мес", 0.0, step=1000.0, value=0.0,
                              key=f"cust_new_seo_fee_{customer_id}")
        region = st.text_input("Регион", key=f"cust_new_seo_region_{customer_id}")
        if st.button("💾 Создать", key=f"cust_new_seo_btn_{customer_id}", type='primary'):
            if not site:
                st.error("Нужен сайт")
            else:
                pid = seo_db.create_project(
                    customer_id=customer_id, site_url=site,
                    domain=extract_domain(site), region=region,
                    monthly_fee=fee, status='new',
                )
                st.success(f"✅ SEO-проект #{pid} создан")
                st.rerun()
