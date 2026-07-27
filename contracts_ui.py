"""
Вкладка «Договоры» — список, просмотр, редактирование, перегенерация, скачивание.
"""
from __future__ import annotations
import json
import io
import datetime as dt
import streamlit as st

import contracts_db
import suppliers as _suppliers
import dogovor_traversa as _dt


def _fmt_money_ui(v):
    """Формат для UI без ₽."""
    if v is None or v == 0: return "0"
    return _dt._fmt_money(v)


def render_contracts_tab():
    """Основная вкладка «Договоры»."""
    st.markdown("### 📑 Договоры")

    contracts_db.init_schema()

    # Если выбран конкретный договор — показываем карточку
    open_id = st.session_state.get("_contract_open_id")
    if open_id:
        _render_contract_card(open_id)
        return

    # Иначе — список
    _render_contracts_list()


def _render_contracts_list():
    contracts = contracts_db.list_contracts()

    # Метрики
    total = sum(float(c.get("total_amount") or 0) for c in contracts)
    with_vat = sum(1 for c in contracts if c.get("has_vat"))
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Всего договоров", len(contracts))
    col2.metric("Сумма", _fmt_money_ui(total) + " ₽")
    col3.metric("С НДС", with_vat)
    col4.metric("Без НДС", len(contracts) - with_vat)

    st.divider()

    if not contracts:
        st.info("Пока нет сохранённых договоров. Сформируйте договор в любом режиме — он автоматически появится здесь.")
        return

    # Фильтр
    with st.expander("🔍 Фильтры", expanded=False):
        col1, col2, col3 = st.columns(3)
        with col1:
            q_search = st.text_input("Поиск по № или покупателю", key="c_search")
        with col2:
            f_supplier = st.selectbox("Поставщик",
                ["Все", "ЛКС", "МОДЕРНИЗАЦИЯ", "КИНЕМАТИКА"], key="c_sup_filter")
        with col3:
            f_vat = st.selectbox("НДС", ["Все", "С НДС 22%", "Без НДС (УСН)"], key="c_vat_filter")

    filtered = contracts
    if q_search:
        q_lower = q_search.lower()
        filtered = [c for c in filtered if
                    q_lower in str(c.get("contract_number","")).lower() or
                    q_lower in str(c.get("buyer_short","")).lower() or
                    q_lower in str(c.get("buyer_inn","")).lower()]
    if f_supplier != "Все":
        supkey_map = {"ЛКС": "LKS", "МОДЕРНИЗАЦИЯ": "MODERNIZATSIYA", "КИНЕМАТИКА": "KINEMATIKA"}
        filtered = [c for c in filtered if c.get("supplier_key") == supkey_map[f_supplier]]
    if f_vat != "Все":
        want_vat = 1 if f_vat == "С НДС 22%" else 0
        filtered = [c for c in filtered if (c.get("has_vat") or 0) == want_vat]

    st.caption(f"Показано: {len(filtered)} из {len(contracts)}")

    # Таблица-список
    for c in filtered:
        with st.container():
            cols = st.columns([0.8, 2.2, 2.8, 1.5, 1.2, 1.0, 1.0, 1.0])
            cols[0].caption(f"#{c['id']}")
            cols[1].markdown(f"**{c.get('contract_number','—')}**")
            cols[1].caption(c.get("contract_date","") or "")
            cols[2].markdown(c.get("buyer_short","—") or "—")
            cols[2].caption(f"ИНН {c.get('buyer_inn','')}" if c.get("buyer_inn") else "")
            cols[3].markdown(c.get("supplier_short","—") or "—")
            cols[4].markdown(f"**{_fmt_money_ui(c.get('total_amount',0))} ₽**")
            cols[4].caption("с НДС 22%" if c.get("has_vat") else "без НДС")
            with cols[5]:
                if st.button("📂 Открыть", key=f"open_c_{c['id']}", use_container_width=True):
                    st.session_state["_contract_open_id"] = c["id"]
                    st.rerun()
            with cols[6]:
                # Скачивание — берём blob
                if st.button("⬇️", key=f"dl_c_{c['id']}", help="Скачать DOCX", use_container_width=True):
                    full = contracts_db.get_contract(c["id"])
                    if full and full.get("docx_blob"):
                        st.session_state[f"_dl_c_{c['id']}"] = bytes(full["docx_blob"])
                if st.session_state.get(f"_dl_c_{c['id']}"):
                    st.download_button("↓",
                        data=st.session_state[f"_dl_c_{c['id']}"],
                        file_name=f"Договор_{c.get('contract_number','').replace('/','_')}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dlbtn_c_{c['id']}", use_container_width=True)
            with cols[7]:
                if st.button("🗑", key=f"del_c_{c['id']}", help="Удалить", use_container_width=True):
                    st.session_state[f"_confirm_del_c_{c['id']}"] = True
                if st.session_state.get(f"_confirm_del_c_{c['id']}"):
                    if st.button("Точно?", key=f"del_ok_c_{c['id']}"):
                        contracts_db.delete_contract(c["id"])
                        st.session_state[f"_confirm_del_c_{c['id']}"] = False
                        st.success(f"Договор #{c['id']} удалён")
                        st.rerun()
        st.divider()


def _render_contract_card(contract_id: int):
    contract = contracts_db.get_contract(contract_id)
    if not contract:
        st.error(f"Договор #{contract_id} не найден")
        if st.button("← К списку"):
            st.session_state["_contract_open_id"] = None
            st.rerun()
        return

    col_back, col_title = st.columns([1, 5])
    with col_back:
        if st.button("← К списку", use_container_width=True):
            st.session_state["_contract_open_id"] = None
            st.rerun()
    with col_title:
        st.markdown(f"### 📝 Договор № **{contract.get('contract_number','—')}** от {contract.get('contract_date','—')}")

    st.divider()

    params = contract.get("params", {}) or {}
    buyer = params.get("buyer", {}) or {}
    lines = params.get("lines", []) or []

    tabs = st.tabs(["📊 Обзор", "✏️ Редактирование", "📥 Скачать / Перегенерировать"])

    # ============ ОБЗОР ============
    with tabs[0]:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("**Покупатель**")
            st.markdown(f"- {buyer.get('short','—')}")
            st.markdown(f"- ИНН {buyer.get('inn','—')} / КПП {buyer.get('kpp','—')}")
            st.markdown(f"- Адрес: {buyer.get('address','—')}")
            st.markdown(f"- Директор: {buyer.get('director_fio_short','—')}")
        with col2:
            st.markdown("**Поставщик**")
            st.markdown(f"- {contract.get('supplier_short','—')} ({contract.get('supplier_key','—')})")
            st.markdown(f"- НДС: {'22% (включён в стоимость)' if contract.get('has_vat') else 'без НДС (УСН)'}")
            st.markdown(f"- Итого: **{_fmt_money_ui(contract.get('total_amount',0))} ₽**")
            if contract.get("has_vat") and contract.get("vat_amount"):
                st.markdown(f"- в т.ч. НДС: {_fmt_money_ui(contract.get('vat_amount',0))} ₽")

        st.divider()
        st.markdown("**Условия**")
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Предоплата", f"{contract.get('prepay_pct',100)} %")
        col2.metric("Срок отгрузки", f"{contract.get('shipment_days',20)} раб.дн.")
        col3.metric("Гарантия", f"{contract.get('warranty_months',12)} мес.")
        col4.metric("Тип", contract.get("contract_type","КП"))
        if contract.get("delivery_terms"):
            st.info(f"🚚 Доставка: {contract['delivery_terms']}" +
                   (f" — {contract['delivery_address']}" if contract.get("delivery_address") else ""))

        st.divider()
        st.markdown(f"**Спецификация ({len(lines)} позиций)**")
        if lines:
            import pandas as pd
            df = pd.DataFrame([{
                "№": i+1,
                "Код": l.get("code",""),
                "Наименование": l.get("name",""),
                "Ед.": l.get("unit","шт"),
                "Кол-во": l.get("qty",0),
                "Цена, ₽": _fmt_money_ui(l.get("price",0)),
                "Сумма, ₽": _fmt_money_ui(float(l.get("qty",0)) * float(l.get("price",0))),
            } for i, l in enumerate(lines)])
            st.dataframe(df, hide_index=True, use_container_width=True)

    # ============ РЕДАКТИРОВАНИЕ ============
    with tabs[1]:
        st.markdown("**Параметры договора**")

        col1, col2, col3 = st.columns(3)
        with col1:
            edit_prepay = st.number_input("Предоплата, %",
                min_value=0, max_value=100, value=int(contract.get("prepay_pct",100)),
                key=f"edit_prepay_{contract_id}")
        with col2:
            edit_ship = st.number_input("Срок отгрузки, раб.дн.",
                min_value=1, max_value=365, value=int(contract.get("shipment_days",20)),
                key=f"edit_ship_{contract_id}")
        with col3:
            edit_warr = st.number_input("Гарантия, мес.",
                min_value=1, max_value=120, value=int(contract.get("warranty_months",12)),
                key=f"edit_warr_{contract_id}")

        edit_delivery = st.text_input("Условия доставки",
            value=contract.get("delivery_terms","") or "",
            key=f"edit_delivery_{contract_id}")
        edit_addr = st.text_input("Адрес доставки",
            value=contract.get("delivery_address","") or "",
            key=f"edit_addr_{contract_id}")

        st.markdown("**Позиции**")
        st.caption("Редактируй ниже — цены/кол-во сохранятся в базе и попадут в перегенерированный DOCX")

        import pandas as pd
        df_edit = pd.DataFrame([{
            "Код": l.get("code",""),
            "Наименование": l.get("name",""),
            "Ед.": l.get("unit","шт"),
            "Кол-во": float(l.get("qty",0)),
            "Цена, ₽": float(l.get("price",0)),
        } for l in lines])
        edited = st.data_editor(df_edit, num_rows="dynamic", hide_index=True,
                                use_container_width=True,
                                key=f"edit_lines_{contract_id}")

        # Пересчёт итога
        new_lines = []
        for _, r in edited.iterrows():
            if not r.get("Наименование"): continue
            new_lines.append({
                "code": str(r.get("Код","") or ""),
                "name": str(r.get("Наименование","")),
                "unit": str(r.get("Ед.","шт") or "шт"),
                "qty": float(r.get("Кол-во",0) or 0),
                "price": float(r.get("Цена, ₽",0) or 0),
            })

        new_total, new_vat = _dt._calc_totals_vat(new_lines, bool(contract.get("has_vat")))

        col1, col2 = st.columns(2)
        col1.metric("Итог", f"{_fmt_money_ui(new_total)} ₽")
        if contract.get("has_vat"):
            col2.metric("в т.ч. НДС 22%", f"{_fmt_money_ui(new_vat)} ₽")

        st.divider()
        col_a, col_b = st.columns(2)
        with col_a:
            if st.button("💾 Сохранить изменения (без перегенерации)",
                        use_container_width=True, key=f"save_edit_{contract_id}"):
                # Обновляем params_json + суммы
                new_params = dict(params)
                new_params["lines"] = new_lines
                contracts_db.update_contract_params(
                    contract_id, json.dumps(new_params, ensure_ascii=False),
                    lines_summary={
                        "total": float(new_total), "vat": float(new_vat),
                        "prepay_pct": edit_prepay, "shipment_days": edit_ship,
                        "warranty_months": edit_warr,
                        "delivery_terms": edit_delivery, "delivery_address": edit_addr,
                    })
                st.success("✅ Сохранено. Для получения нового DOCX нажмите «Перегенерировать»")
                st.rerun()
        with col_b:
            if st.button("🔄 Перегенерировать DOCX", type="primary",
                        use_container_width=True, key=f"regen_{contract_id}"):
                # Строим docx через единую болванку
                sup_key = contract.get("supplier_key","LKS")
                supplier = _suppliers.get_supplier(sup_key)

                docx_bytes = _dt.build_dogovor_traversa_docx(
                    lines=new_lines, buyer=buyer, supplier=supplier,
                    contract_number=contract["contract_number"],
                    contract_date_str=contract["contract_date"],
                    has_vat=bool(contract.get("has_vat")),
                    prepay_pct=int(edit_prepay),
                    shipment_days=int(edit_ship),
                    warranty_months=int(edit_warr),
                    delivery_terms=edit_delivery or "Самовывоз со склада Поставщика",
                    delivery_address=edit_addr or "",
                )
                # Сохраняем + обновляем params
                new_params = dict(params)
                new_params["lines"] = new_lines
                contracts_db.update_contract_params(
                    contract_id, json.dumps(new_params, ensure_ascii=False),
                    lines_summary={
                        "total": float(new_total), "vat": float(new_vat),
                        "prepay_pct": edit_prepay, "shipment_days": edit_ship,
                        "warranty_months": edit_warr,
                        "delivery_terms": edit_delivery, "delivery_address": edit_addr,
                    })
                contracts_db.update_contract_docx(
                    contract_id, docx_bytes,
                    total_amount=float(new_total), vat_amount=float(new_vat))
                st.success("✅ Договор перегенерирован")
                st.session_state[f"_regen_bytes_{contract_id}"] = docx_bytes
                st.rerun()

    # ============ СКАЧАТЬ / ПЕРЕГЕНЕРИРОВАТЬ ============
    with tabs[2]:
        if contract.get("docx_blob"):
            docx = bytes(contract["docx_blob"])
            st.download_button(
                "⬇️ Скачать текущий DOCX",
                data=docx,
                file_name=f"Договор_{contract.get('contract_number','').replace('/','_')}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True, type="primary")
        else:
            st.warning("DOCX не сохранён в базе. Перегенерируйте на вкладке «Редактирование».")

        st.caption(f"Создан: {contract.get('created_at','—')}")
        if contract.get("updated_at"):
            st.caption(f"Обновлён: {contract.get('updated_at','—')}")
