# -*- coding: utf-8 -*-
"""UI-вкладка «Внешний договор» — договор по стороннему КП."""
from __future__ import annotations
import io
from datetime import date

import streamlit as st

import external_kp_parser as _ext_parser
import external_contract as _ext_dog
import suppliers as _suppliers
import crm_db as _crm_db


def _fmt_money(x: float) -> str:
    return f"{x:,.2f}".replace(",", " ").replace(".", ",")


def _get_next_ext_contract_number(prefix: str) -> str:
    """Уникальный № договора: 22072026-ВД/ЛКС → -2 -ВД/ЛКС и т.д."""
    base = date.today().strftime("%d%m%Y") + f"-ВД/{prefix}"
    return _crm_db.generate_unique_kp_number(base)


def render_external_contract_tab():
    """Главная страница вкладки «📎 Внешний договор»."""
    st.markdown(
        "<div style='background:#FFF3E0;border-left:4px solid #F97316;"
        "padding:10px 14px;border-radius:6px;margin-bottom:12px;'>"
        "<b>📎 Внешний договор</b> — сгенерировать договор поставки по стороннему КП. "
        "Загрузите PDF или DOCX внешнего КП — спецификация распарсится автоматически. "
        "Договор автоматически сохранится в CRM."
        "</div>", unsafe_allow_html=True)

    # ============= 1. ПОСТАВЩИК =============
    st.subheader("1. Поставщик (от кого договор)")
    _sup_keys = list(_suppliers.SUPPLIERS.keys())
    _sup_names = [_suppliers.SUPPLIERS[k].get("short") or _suppliers.SUPPLIERS[k].get("label", k)
                  for k in _sup_keys]
    _sup_pick = st.selectbox(
        "От какой компании выставляем договор",
        range(len(_sup_keys)),
        format_func=lambda i: _sup_names[i],
        index=_sup_keys.index(st.session_state.get("ext_supplier_key",
                              _suppliers.DEFAULT_SUPPLIER_KEY))
        if st.session_state.get("ext_supplier_key",
                                _suppliers.DEFAULT_SUPPLIER_KEY) in _sup_keys else 0,
        key="ext_supplier_pick",
    )
    supplier_key = _sup_keys[_sup_pick]
    st.session_state["ext_supplier_key"] = supplier_key
    supplier = _suppliers.SUPPLIERS[supplier_key]
    _sup_prefix = {"LKS": "ЛКС", "MODERNIZATSIYA": "МОД",
                   "KINEMATIKA": "КИН"}.get(supplier_key, "ЛКС")

    st.markdown("---")

    # ============= 2. ПОКУПАТЕЛЬ =============
    st.subheader("2. Покупатель")
    _mode_col1, _mode_col2 = st.columns([1, 3])
    with _mode_col1:
        buyer_mode = st.radio(
            "Способ ввода",
            ["Вручную", "Скопировать текст", "Из файла (автораспознавание)"],
            key="ext_buyer_mode",
            label_visibility="collapsed",
        )

    if buyer_mode == "Скопировать текст":
        st.markdown(
            "<div style='background:#FEF3E7;border-left:4px solid #F97316;"
            "padding:8px 12px;border-radius:6px;font-size:13px;'>"
            "📋 Вставьте в поле любой текст с реквизитами: из письма, с сайта, "
            "скопированный из чужого договора или выписки. Приложение автоматически "
            "разберёт ИНН, КПП, ОГРН, банк, БИК, счета, адрес, телефон, email, ФИО директора "
            "и раскладет по полям ниже."
            "</div>", unsafe_allow_html=True)
        _paste_txt = st.text_area(
            "Вставьте текст с реквизитами",
            height=250,
            key="ext_buyer_paste_txt",
            placeholder="Пример:\n\n"
                       "ООО «Ромашка»\n"
                       "ИНН 7712345678, КПП 771201001\n"
                       "ОГРН 1234567890123\n"
                       "Юр. адрес: г. Москва, ул. Тестовая, д. 1\n"
                       "Банк: АО «Тинькофф Банк»\n"
                       "БИК 044525974\n"
                       "Р/с 40702810100000000001\n"
                       "Кор/с 30101810400000000225\n"
                       "Тел.: +7 (999) 111-22-33, info@romashka.ru\n"
                       "Генеральный директор Иванов Иван Иванович")
        if st.button("🔍 Разобрать и разложить по полям",
                     key="ext_buyer_paste_extract_btn",
                     type="primary",
                     use_container_width=True,
                     disabled=not _paste_txt.strip()):
            try:
                req = _ext_parser.extract_requisites_from_text(_paste_txt)
                if req.is_empty():
                    st.warning("⚠️ Не удалось распознать ни одного поля. "
                              "Проверьте что в тексте есть ИНН и остальные реквизиты, или заполните вручную.")
                else:
                    _new_data = {
                        "full": req.company_full or "",
                        "short": req.company_short or "",
                        "inn": req.inn or "",
                        "kpp": req.kpp or "",
                        "ogrn": req.ogrn or "",
                        "address": req.address or "",
                        "bank": req.bank_name or "",
                        "bik": req.bank_bik or "",
                        "account": req.bank_account or "",
                        "corr_account": req.corr_account or "",
                        "director_short": req.director_short or "",
                        "director_gen": req.director_gen or "",
                        "director_title": req.director_title or "Генеральный директор",
                        "basis": "Устава",
                        "phone": req.phone or "",
                        "email": req.email or "",
                    }
                    st.session_state["ext_buyer_data"] = _new_data
                    # Автосохранение в историю по ИНН (так же как в КП)
                    try:
                        import history_memory as _hm_paste
                        if req.inn:
                            _hm_paste.remember_ext_buyer(req.inn, _new_data)
                    except Exception:
                        pass
                    _found = []
                    if req.company_short: _found.append(f"Компания: {req.company_short}")
                    if req.inn: _found.append(f"ИНН {req.inn}")
                    if req.kpp: _found.append(f"КПП {req.kpp}")
                    if req.ogrn: _found.append(f"ОГРН {req.ogrn}")
                    if req.bank_name: _found.append(f"Банк: {req.bank_name}")
                    if req.bank_bik: _found.append(f"БИК {req.bank_bik}")
                    if req.bank_account: _found.append(f"р/с {req.bank_account[:8]}…")
                    if req.director_short: _found.append(f"Директор: {req.director_short}")
                    if req.phone: _found.append(f"Тел.: {req.phone}")
                    if req.email: _found.append(f"E-mail: {req.email}")
                    st.success("✅ Распознано: " + " · ".join(_found))
                    st.info("Поля ниже заполнены — проверьте и при необходимости отредактируйте.")
                    st.rerun()
            except Exception as e:
                st.error(f"Ошибка распознавания: {e}")

    elif buyer_mode == "Из файла (автораспознавание)":
        st.markdown(
            "<div style='background:#EFF6FF;border-left:4px solid #3B82F6;"
            "padding:8px 12px;border-radius:6px;font-size:13px;'>"
            "📄 Загрузите документ покупателя (карточка организации / выписка / "
            "существующий договор в PDF или DOCX). Приложение автоматически "
            "распознает ИНН, КПП, ОГРН, банк, БИК, счета, адрес, телефон, "
            "email и ФИО директора."
            "</div>", unsafe_allow_html=True)
        buyer_upl = st.file_uploader(
            "Файл покупателя (DOCX / DOC / PDF / JPG / PNG)",
            type=["pdf", "docx", "doc", "jpg", "jpeg", "png", "webp", "txt"],
            key="ext_buyer_file_uploader",
            help="Поддерживаются: DOCX, DOC (старый Word), PDF (текст или скан), JPG/PNG (фото/скриншот)")
        if buyer_upl and st.button("🔍 Распознать реквизиты из файла",
                                    key="ext_buyer_extract_btn",
                                    type="primary",
                                    use_container_width=True):
            # ---------------------------------------------------------------
            # ТОЧНО ТАКАЯ ЖЕ ЛОГИКА КАК В КП (app.py, стр. ~4353):
            # 1) вытягиваем текст из файла (пробуем smart, если нет — просто docx)
            # 2) прогоняем через extract_requisites_from_text
            # 3) раскладываем по ключам внешнего договора
            # ---------------------------------------------------------------
            _buyer_bytes = buyer_upl.read()
            _fname = buyer_upl.name
            _raw_text = ""

            # Шаг 1: текст из файла
            try:
                from smart_requisites import _extract_text_smart
                _raw_text = _extract_text_smart(_buyer_bytes, _fname)
            except Exception:
                _raw_text = ""

            # Fallback: простой чтение docx через req_parser (гарантированно есть на Cloud)
            if not _raw_text.strip():
                if _fname.lower().endswith(".docx"):
                    try:
                        from req_parser import extract_text_from_docx
                        _raw_text = extract_text_from_docx(_buyer_bytes)
                    except Exception:
                        pass
                elif _fname.lower().endswith(".pdf"):
                    try:
                        from external_kp_parser import _extract_text_from_pdf
                        _raw_text = _extract_text_from_pdf(_buyer_bytes)
                    except Exception:
                        pass
                elif _fname.lower().endswith(".txt"):
                    try:
                        _raw_text = _buyer_bytes.decode("utf-8", errors="ignore")
                    except Exception:
                        _raw_text = _buyer_bytes.decode("cp1251", errors="ignore")

            # Шаг 2: парсим текст
            req = None
            if _raw_text.strip():
                try:
                    req = _ext_parser.extract_requisites_from_text(_raw_text)
                except Exception:
                    req = None
            # Финальный fallback — прямой extract_requisites(файл)
            if req is None or req.is_empty():
                try:
                    req = _ext_parser.extract_requisites(_buyer_bytes, _fname)
                except Exception:
                    req = None

            if req is None or req.is_empty():
                st.warning("⚠️ Не удалось распознать реквизиты. "
                          "Проверьте формат файла или заполните вручную.")
            else:
                # Шаг 3: раскладываем по ключам внешнего договора
                _new_data = {
                    "full": req.company_full or "",
                    "short": req.company_short or "",
                    "inn": req.inn or "",
                    "kpp": req.kpp or "",
                    "ogrn": req.ogrn or "",
                    "address": req.address or "",
                    "bank": req.bank_name or "",
                    "bik": req.bank_bik or "",
                    "account": req.bank_account or "",
                    "corr_account": req.corr_account or "",
                    "director_short": req.director_short or "",
                    "director_gen": req.director_gen or "",
                    "director_title": req.director_title or "Генеральный директор",
                    "basis": "Устава",
                    "phone": req.phone or "",
                    "email": req.email or "",
                }
                st.session_state["ext_buyer_data"] = _new_data
                # КРИТИЧНО: записываем НАПРЯМУЮ в ключи виджетов — иначе поля не обновятся
                st.session_state["ext_bf"] = _new_data["full"]
                st.session_state["ext_bs"] = _new_data["short"]
                st.session_state["ext_bin"] = _new_data["inn"]
                st.session_state["ext_bkp"] = _new_data["kpp"]
                st.session_state["ext_bog"] = _new_data["ogrn"]
                st.session_state["ext_bad"] = _new_data["address"]
                st.session_state["ext_bds"] = _new_data["director_short"]
                st.session_state["ext_bdg"] = _new_data["director_gen"]
                st.session_state["ext_bdt"] = _new_data["director_title"]
                st.session_state["ext_bbs"] = _new_data["basis"]
                st.session_state["ext_bbk"] = _new_data["bank"]
                st.session_state["ext_bbik"] = _new_data["bik"]
                st.session_state["ext_bac"] = _new_data["account"]
                st.session_state["ext_bcs"] = _new_data["corr_account"]
                st.session_state["ext_bph"] = _new_data["phone"]
                st.session_state["ext_bem"] = _new_data["email"]
                # Автосохранение в историю по ИНН
                try:
                    import history_memory as _hm_paste
                    if req.inn:
                        _hm_paste.remember_ext_buyer(req.inn, _new_data)
                except Exception:
                    pass
                _found = []
                if req.company_short: _found.append(f"Компания: {req.company_short}")
                if req.inn: _found.append(f"ИНН {req.inn}")
                if req.kpp: _found.append(f"КПП {req.kpp}")
                if req.ogrn: _found.append(f"ОГРН {req.ogrn}")
                if req.bank_name: _found.append(f"Банк: {req.bank_name}")
                if req.bank_account: _found.append(f"р/с {req.bank_account[:8]}…")
                if req.director_short: _found.append(f"Директор: {req.director_short}")
                if req.phone: _found.append(f"Тел.: {req.phone}")
                if req.email: _found.append(f"E-mail: {req.email}")
                st.success("✅ Распознано: " + " · ".join(_found))
                st.info("Поля заполнены ниже — проверьте и при необходимости отредактируйте.")
                st.rerun()

    # Данные покупателя — редактируемые поля
    _bd = st.session_state.get("ext_buyer_data", {})
    st.markdown("**Данные покупателя (можно отредактировать):**")

    # === Тип покупателя (Пункт 4) ===
    _ext_buyer_type_ui = st.radio(
        "Тип покупателя",
        options=["ООО / Юр. лицо", "ИП (Индивидуальный предприниматель)"],
        index=0 if _bd.get("buyer_type", "OOO") == "OOO" else 1,
        key="ext_buyer_type_ui",
        horizontal=True,
        help="Для ИП преамбула автоматически адаптируется: «Свидетельство о гос. регистрации», «именуемый».")
    ext_buyer_type = "IP" if _ext_buyer_type_ui.startswith("ИП") else "OOO"

    bc1, bc2 = st.columns(2)
    with bc1:
        buyer_full = st.text_input("Полное наименование",
                                    value=_bd.get("full", ""), key="ext_bf")
        buyer_short = st.text_input("Краткое наименование",
                                     value=_bd.get("short", ""), key="ext_bs")
        buyer_inn = st.text_input("ИНН", value=_bd.get("inn", ""), key="ext_bin")
        buyer_kpp = st.text_input("КПП", value=_bd.get("kpp", ""), key="ext_bkp")
        buyer_ogrn = st.text_input("ОГРН", value=_bd.get("ogrn", ""), key="ext_bog")
        buyer_address = st.text_area("Юридический адрес",
                                      value=_bd.get("address", ""),
                                      key="ext_bad", height=60)
    with bc2:
        buyer_director_short = st.text_input(
            "Директор (Фамилия И.О.)",
            value=_bd.get("director_short", ""), key="ext_bds")
        buyer_director_gen = st.text_input(
            "Директор в родительном падеже (кого)",
            value=_bd.get("director_gen", ""), key="ext_bdg",
            help="Пример: Иванова Ивана Ивановича")
        buyer_director_title = st.text_input(
            "Должность директора",
            value=_bd.get("director_title", "Генеральный директор"),
            key="ext_bdt")
        buyer_basis = st.text_input(
            "Действует на основании",
            value=_bd.get("basis", "Устава"), key="ext_bbs")
        buyer_bank = st.text_input("Банк",
                                    value=_bd.get("bank", ""), key="ext_bbk")
        buyer_bik = st.text_input("БИК",
                                   value=_bd.get("bik", ""), key="ext_bbik")

    bc3, bc4 = st.columns(2)
    with bc3:
        buyer_account = st.text_input("Расчётный счёт",
                                       value=_bd.get("account", ""), key="ext_bac")
        buyer_corr_account = st.text_input("Корр. счёт",
                                            value=_bd.get("corr_account", ""),
                                            key="ext_bca")
    with bc4:
        buyer_phone = st.text_input("Телефон",
                                     value=_bd.get("phone", "+7 "), key="ext_bph")
        buyer_email = st.text_input("E-mail",
                                     value=_bd.get("email", ""), key="ext_bem")

    st.markdown("---")

    # ============= 3. СПЕЦИФИКАЦИЯ =============
    st.subheader("3. Спецификация (PDF или DOCX внешнего КП)")

    upl = st.file_uploader(
        "Загрузите PDF или DOCX внешнего КП",
        type=["pdf", "docx"],
        key="ext_kp_uploader",
        help="Позиции распарсятся автоматически. Если формат нестандартный — "
             "нажмите «Ввести спецификацию вручную» ниже.")

    parsed_items = []
    if upl:
        try:
            file_bytes = upl.read()
            parsed = _ext_parser.parse_external_kp(file_bytes, upl.name)
            parsed_items = [
                {"code": p.code, "name": p.name, "unit": p.unit,
                 "qty": p.qty, "price": p.price}
                for p in parsed
            ]
            if parsed_items:
                st.success(f"✅ Распарсено позиций: {len(parsed_items)}")
                st.session_state["ext_parsed_items"] = parsed_items
                st.session_state["ext_kp_source_bytes"] = file_bytes
                st.session_state["ext_kp_source_name"] = upl.name
            else:
                st.warning("⚠️ Не удалось распарсить позиции автоматически. "
                          "Проверьте формат КП или введите спецификацию вручную.")
        except Exception as e:
            st.error(f"Ошибка парсинга: {e}")

    # Показываем распарсенное или даём ввести вручную
    if not st.session_state.get("ext_parsed_items"):
        st.session_state["ext_parsed_items"] = []

    manual_mode = st.checkbox(
        "✍️ Ввести спецификацию вручную (без загрузки файла)",
        value=len(st.session_state.get("ext_parsed_items", [])) == 0,
        key="ext_manual_mode")

    if manual_mode or st.session_state.get("ext_parsed_items"):
        import pandas as pd
        _items_for_editor = st.session_state.get("ext_parsed_items", [])
        if not _items_for_editor:
            _items_for_editor = [
                {"code": "", "name": "", "unit": "шт", "qty": 1.0, "price": 0.0}
                for _ in range(3)
            ]
        _df_items = pd.DataFrame(_items_for_editor)
        # Нормализуем многострочные артикулы (переносы → пробел)
        if "code" in _df_items.columns:
            _df_items["code"] = _df_items["code"].astype(str).str.replace(
                r"[\s\n\r]+", " ", regex=True).str.strip()

        # Превью — списком с полным артикулом (т.к. Streamlit data_editor рендерит canvas без переноса)
        _long_codes = [str(r.get("code") or "") for r in _items_for_editor 
                       if len(str(r.get("code") or "")) > 20]
        if _long_codes:
            with st.expander("👁 Полный список артикулов (если в таблице обрезаны)", expanded=False):
                for r in _items_for_editor:
                    _code = str(r.get("code") or "").strip()
                    _name = str(r.get("name") or "").strip()
                    if not _code and not _name:
                        continue
                    st.markdown(f"**код:** `{_code}`  —  **{_name}**")

        _df_edited = st.data_editor(
            _df_items,
            num_rows="dynamic",
            use_container_width=True,
            hide_index=True,
            key="ext_items_editor",
            column_config={
                "code": st.column_config.TextColumn(
                    "Код / Артикул",
                    width="large",
                    help="Полный артикул — если он длинный, расширьте колонку"),
                "name": st.column_config.TextColumn("Наименование", width="large"),
                "unit": st.column_config.TextColumn("Ед.", width="small"),
                "qty": st.column_config.NumberColumn("Кол-во", format="%.1f", min_value=0.0, width="small"),
                "price": st.column_config.NumberColumn("Цена, ₽", format="%.2f", min_value=0.0, width="medium"),
            },
            height=min(600, 60 + len(_df_items) * 45),  # высокие строки под длинные артикулы
        )
        # Финальный список позиций из редактора
        final_items = [
            {"code": str(r["code"] or ""), "name": str(r["name"] or ""),
             "unit": str(r["unit"] or "шт"), "qty": float(r["qty"] or 0),
             "price": float(r["price"] or 0)}
            for _, r in _df_edited.iterrows()
            if str(r["name"] or "").strip()
        ]
    else:
        final_items = st.session_state.get("ext_parsed_items", [])

    _spec_total_raw = sum(it["qty"] * it["price"] for it in final_items)

    # === СКИДКА (Пункт 2) ===
    _dc1, _dc2 = st.columns([1, 2])
    with _dc1:
        ext_discount_pct = st.number_input(
            "Скидка, %",
            min_value=0.0, max_value=50.0,
            value=float(st.session_state.get("ext_discount_pct", 0.0)),
            step=0.5,
            key="ext_discount_pct",
            help="Применяется пропорционально ко всем ценам позиций перед генерацией договора.")
    with _dc2:
        if ext_discount_pct > 0:
            _spec_total = _spec_total_raw * (1 - ext_discount_pct / 100)
            _disc_amount = _spec_total_raw - _spec_total
            st.info(
                f"💰 Сумма позиций: ~~{_fmt_money(_spec_total_raw)}~~ → "
                f"**{_fmt_money(_spec_total)} ₽** (−{_fmt_money(_disc_amount)} = −{ext_discount_pct}%)")
            # Применяем скидку к каждой позиции в final_items (пропорционально)
            _coef = (1 - ext_discount_pct / 100)
            final_items = [
                {**it, "price": round(float(it["price"]) * _coef, 2)}
                for it in final_items
            ]
        else:
            _spec_total = _spec_total_raw
            st.info(f"💰 Сумма позиций: **{_fmt_money(_spec_total)} ₽**")

    st.markdown("---")

    # ============= 4. ЧЕРТЁЖ =============
    st.subheader("4. Чертёж (необязательно)")
    draw_upl = st.file_uploader(
        "Габаритный чертёж — PDF (первая страница), JPG или PNG",
        type=["pdf", "jpg", "jpeg", "png"],
        key="ext_drawing_uploader")
    drawing_bytes = None
    drawing_caption = st.text_input(
        "Подпись под чертежом",
        value="Габаритный чертёж оборудования",
        key="ext_drawing_caption")
    if draw_upl:
        try:
            _db = draw_upl.read()
            drawing_bytes = _ext_parser.extract_drawing_image(_db, draw_upl.name)
            st.success(f"✅ Чертёж загружен ({len(drawing_bytes) // 1024} КБ, "
                      f"будет в Приложении № 1)")
            st.image(drawing_bytes, caption=drawing_caption, use_column_width=True)
        except Exception as e:
            st.error(f"Не удал��сь обработать чертёж: {e}")

    st.markdown("---")

    # ============= 5. ПАРАМЕТРЫ ДОГОВОРА =============
    st.subheader("5. Параметры договора")
    pc1, pc2, pc3 = st.columns(3)
    with pc1:
        contract_number = st.text_input(
            "№ договора",
            value=st.session_state.get("ext_contract_number",
                                       _get_next_ext_contract_number(_sup_prefix)),
            key="ext_contract_number",
            help="Автоматически подставляется следующий свободный номер")
    with pc2:
        contract_date_str = st.text_input(
            "Дата договора",
            value=date.today().strftime("%d.%m.%Y"),
            key="ext_contract_date")
    with pc3:
        vat_mode = st.selectbox(
            "НДС",
            ["С НДС 22 %", "Без НДС"],
            key="ext_vat_mode",
            help="С НДС — цены в спецификации ВКЛЮЧАЮТ 22 % НДС. "
                 "Без НДС — цены без НДС.")

    pc4, pc5, pc6 = st.columns(3)
    with pc4:
        prepay_pct = st.selectbox(
            "% предоплаты",
            [30, 50, 70, 80, 100],
            index=2,
            key="ext_prepay_pct")
    with pc5:
        # Пункт 7: radio-кнопки без дублирования самовывоза
        delivery_type = st.radio(
            "Условия поставки",
            ["Самовывоз", "Доставка"],
            key="ext_delivery_type",
            horizontal=True,
            help="Самовывоз — Покупатель забирает сам со склада; Доставка — силами Поставщика.")
    with pc6:
        shipment_term = st.number_input(
            "Срок изготовления, дней",
            min_value=1, max_value=180, value=20, step=1,
            key="ext_shipment_term")

    # Адрес доставки/самовывоза + стоимость
    dc1, dc2 = st.columns([3, 1])
    with dc1:
        # Дефолтные адреса как в обычном КП
        _pickup_defaults = {
            "Самовывоз": "г. Санкт-Петербург, склад Поставщика "
                         "(ул. Революции, 114, лит. А)",
            "Доставка": "",
        }
        delivery_address = st.text_area(
            "Адрес " + ("доставки" if delivery_type == "Доставка" else "самовывоза"),
            value=st.session_state.get(
                "ext_delivery_address",
                _pickup_defaults.get(delivery_type, "")),
            key="ext_delivery_address", height=60)
    with dc2:
        delivery_cost = st.number_input(
            "Стоимость доставки, ₽",
            min_value=0.0, value=0.0, step=1000.0,
            key="ext_delivery_cost",
            help="0 — если самовывоз или доставка бесплатна",
            disabled=(delivery_type == "Самовывоз"))
        if delivery_type == "Самовывоз":
            delivery_cost = 0.0

    warranty_months = st.number_input(
        "Гарантия, мес.", min_value=1, max_value=60, value=12, step=1,
        key="ext_warranty_months")

    include_stamp = st.checkbox("✍️ Печать и подпись в договоре",
                                value=True, key="ext_include_stamp")

    # Комментарий к договору
    st.markdown(
        "<div style='background:#FFF3E0;border-left:4px solid #F97316;"
        "padding:6px 10px;border-radius:6px;margin-top:8px;font-size:13px;'>"
        "💬 <b>Дополнительные условия</b> — попадут в раздел 7 договора (если заполнены)"
        "</div>", unsafe_allow_html=True)
    comment = st.text_area(" ", key="ext_contract_comment",
                          value=st.session_state.get("ext_contract_comment", ""),
                          height=80, label_visibility="collapsed")

    st.markdown("---")

    # ============= 6. ИТОГИ + ГЕНЕРАЦИЯ =============
    total = _spec_total + float(delivery_cost)
    from money_words import money_to_words
    st.markdown(f"### 💵 Итого договора")
    _im1, _im2 = st.columns([1, 2])
    with _im1:
        st.metric("Общая сумма",
                 f"{_fmt_money(total)} ₽",
                 "с НДС 22 %" if vat_mode == "С НДС 22 %" else "без НДС")
        _prepay_sum = total * prepay_pct / 100
        _remainder = total - _prepay_sum
        st.metric(f"Предоплата {prepay_pct} %", f"{_fmt_money(_prepay_sum)} ₽")
        _rem_pct = round(100 - prepay_pct)
        if _rem_pct > 0:
            st.metric(f"Остаток {_rem_pct} %", f"{_fmt_money(_remainder)} ₽")
    with _im2:
        st.markdown(f"**Прописью:** {money_to_words(total)}")
        if prepay_pct < 100:
            st.markdown(f"**Предоплата прописью:** {money_to_words(_prepay_sum)}")
            st.markdown(f"**Остаток прописью:** {money_to_words(_remainder)}")

    # Ошибки перед генерацией
    _errors = []
    if not buyer_full and not buyer_short:
        _errors.append("Не указано наименование покупателя.")
    if not final_items:
        _errors.append("Не указано ни одной позиции в спецификации.")
    if total <= 0:
        _errors.append("Общая сумма договора должна быть больше нуля.")

    if _errors:
        st.error("⚠️ " + " ".join(_errors))

    if st.button("📥 Сформировать договор",
                type="primary", use_container_width=True,
                disabled=bool(_errors),
                key="ext_generate_btn"):
        try:
            data = _ext_dog.ExternalContractData(
                buyer_full=buyer_full or buyer_short,
                buyer_short=buyer_short or buyer_full,
                buyer_inn=buyer_inn,
                buyer_kpp=buyer_kpp,
                buyer_ogrn=buyer_ogrn,
                buyer_address=buyer_address,
                buyer_bank=buyer_bank,
                buyer_bik=buyer_bik,
                buyer_account=buyer_account,
                buyer_corr_account=buyer_corr_account,
                buyer_director_short=buyer_director_short,
                buyer_director_gen=buyer_director_gen,
                buyer_director_title=buyer_director_title,
                buyer_basis=buyer_basis,
                buyer_phone=buyer_phone,
                buyer_email=buyer_email,
                items=final_items,
                has_vat=(vat_mode == "С НДС 22 %"),
                prepay_pct=int(prepay_pct),
                delivery_type=delivery_type,
                delivery_address=delivery_address,
                delivery_cost=float(delivery_cost or 0),
                contract_number=contract_number,
                contract_date=contract_date_str,
                supplier_key=supplier_key,
                include_stamp=include_stamp,
                drawing_bytes=drawing_bytes,
                drawing_caption=drawing_caption,
                comment=comment,
                shipment_term_days=int(shipment_term),
                warranty_months=int(warranty_months),
                buyer_type=ext_buyer_type,
            )

            with st.spinner("Формирую договор DOCX…"):
                # Стандартный договор build_dogovor_docx — тот же что в КП.
                # ВАЖНО: НЕ делать `import app` — это перевыполнит весь app.py и создаст дубликаты Streamlit-ключей.
                # Мы уже внутри app.py-сессии — берём глобальные из __main__.
                import sys
                _app_mod = sys.modules.get("__main__") or sys.modules.get("app")
                if _app_mod is None:
                    # Fallback — если вдруг нет, всё-таки импортируем (но это будет больно)
                    import app as _app_mod
                # Собираем QuoteData с lines из final_items
                _lines = [
                    _app_mod.SpecLine(
                        code=str(it.get("code") or ""),
                        name=str(it.get("name") or ""),
                        unit=str(it.get("unit") or "шт"),
                        qty=float(it.get("qty") or 0),
                        price=float(it.get("price") or 0),
                    )
                    for it in final_items
                ]
                _q_ext = _app_mod.QuoteData(
                    series="EXT", capacity=0, boom=0, height_to_arm=0,
                    hoist_brand="", hoist_mode="", hoist_height=0,
                    include_electrification=False, include_montage=False,
                    montage_price=0.0,
                )
                _q_ext.lines = _lines
                # buyer dict в формате build_dogovor_docx (комбинируем все ключи)
                _buyer = {
                    "full": buyer_full or buyer_short,
                    "short": buyer_short or buyer_full,
                    "inn": buyer_inn,
                    "kpp": buyer_kpp,
                    "ogrn": buyer_ogrn,
                    "address": buyer_address,
                    "phone": buyer_phone,
                    "email": buyer_email,
                    "bank": buyer_bank,
                    "bik": buyer_bik,
                    "rs": buyer_account,
                    "ks": buyer_corr_account,
                    "director_position": buyer_director_title,
                    "director_fio_gen": buyer_director_gen,
                    "director_fio_short": buyer_director_short,
                    "director_short": buyer_director_short,
                    "director_gen": buyer_director_gen,
                    "basis": buyer_basis,
                }
                # Устанавливаем поставщика глобально для build_dogovor_docx
                _app_mod.SUPPLIER = supplier
                _delivery_terms = delivery_type
                if delivery_address:
                    _delivery_terms = f"{delivery_type}, адрес доставки: {delivery_address}"
                docx_bytes = _app_mod.build_dogovor_docx(
                    _q_ext, _buyer, contract_number, contract_date_str,
                    prepay_pct=int(prepay_pct),
                    delivery_terms=_delivery_terms,
                    include_stamp=bool(include_stamp),
                    shipment_term=str(shipment_term),
                    warranty_text=f"{int(warranty_months)} месяцев" if warranty_months else None,
                    kp_comment=str(comment or "").strip(),
                    drawing_bytes=drawing_bytes,
                    drawing_caption=drawing_caption or "",
                    buyer_type=ext_buyer_type,
                )
                # PDF во внешнем договоре не генерируем — так же как в КП (только DOCX-договор).
                # Если нужно PDF — открой DOCX в Word → Файл → Сохранить как PDF.
                pdf_bytes = None

            st.success("✅ Договор сформирован")

            # --- Сохранение в CRM ---
            try:
                # Уникализация номера
                unique_num = _crm_db.generate_unique_kp_number(contract_number)
                if unique_num != contract_number:
                    st.info(f"⚠️ Номер {contract_number} уже занят — присвоен: **{unique_num}**")
                    data.contract_number = unique_num
                    # Пересобираем с новым номером — стандартная форма
                    docx_bytes = _app_mod.build_dogovor_docx(
                        _q_ext, _buyer, unique_num, contract_date_str,
                        prepay_pct=int(prepay_pct),
                        delivery_terms=_delivery_terms,
                        include_stamp=bool(include_stamp),
                        shipment_term=str(shipment_term),
                        warranty_text=f"{int(warranty_months)} месяцев" if warranty_months else None,
                        kp_comment=str(comment or "").strip(),
                        drawing_bytes=drawing_bytes,
                        drawing_caption=drawing_caption or "",
                        buyer_type=ext_buyer_type,
                    )
                    pdf_bytes = None

                # Клиент — upsert
                cust_id = None
                if buyer_inn:
                    _cust = _crm_db.Customer(
                        name_short=buyer_short or buyer_full,
                        name_full=buyer_full,
                        inn=buyer_inn, kpp=buyer_kpp, ogrn=buyer_ogrn,
                        address=buyer_address,
                        bank=buyer_bank, bik=buyer_bik,
                        rs=buyer_account,
                        ks=buyer_corr_account,
                        director_fio=buyer_director_short,
                        director_position=buyer_director_title,
                        phone=buyer_phone, email=buyer_email,
                    )
                    cust_id = _crm_db.upsert_customer(_cust)

                # Формируем QuoteItem-ы
                q_items = [
                    _crm_db.QuoteItem(
                        code=it["code"], name=it["name"], unit=it["unit"],
                        qty=it["qty"], price=it["price"])
                    for it in final_items
                ]
                if delivery_cost > 0:
                    q_items.append(_crm_db.QuoteItem(
                        code="—", name=f"Доставка ({delivery_type})",
                        unit="усл.", qty=1, price=float(delivery_cost)))
                rec = _crm_db.QuoteRecord(
                    kp_number=unique_num, customer_id=cust_id,
                    product_type="Внешний договор",
                    product_model="—",
                    include_montage=False,
                    delivery_city=delivery_address,
                    base_total=total, discount_pct=float(ext_discount_pct or 0),
                    status=_crm_db.STATUS_DRAFT,
                    items=q_items,
                )
                qid = _crm_db.save_quote(rec)
                st.success(f"💾 Договор {unique_num} сохранён в CRM — id={qid}")
            except Exception as _ce:
                st.warning(f"CRM сохранение: {_ce}")

            # --- Скачивание ---
            dc1, dc2 = st.columns(2)
            with dc1:
                st.download_button(
                    "⬇️ Скачать Договор поставки (DOCX)",
                    data=docx_bytes,
                    file_name=f"Договор_{data.contract_number.replace('/','_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True,
                    key="ext_dl_docx")
            with dc2:
                st.info("💡 Для PDF: открой DOCX в Word/Pages → Файл → Сохранить как PDF")
        except Exception as e:
            import traceback
            st.error(f"Ошибка формирования договора: {e}")
            st.code(traceback.format_exc())
