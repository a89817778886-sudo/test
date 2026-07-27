# -*- coding: utf-8 -*-
"""
Красивое КП в формате PDF на reportlab.
Работает без LibreOffice — подходит для Streamlit Cloud.
Кириллица через DejaVu Sans (лежит в fonts/).
"""
from __future__ import annotations

import io
from datetime import date
from pathlib import Path
from typing import Optional

from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    BaseDocTemplate,
    Frame,
    Image,
    KeepTogether,
    PageBreak,
    PageTemplate,
    Paragraph,
    Spacer,
    Table,
    TableStyle,
)

APP_DIR = Path(__file__).resolve().parent
MEDIA_DIR = APP_DIR / "media"
FONTS_DIR = APP_DIR / "fonts"

# Фирменная палитра rolls-kran.ru — белый/оранжевый/чёрный
BRAND = colors.HexColor("#FBBF87")          # светлый персиковый (мягкий)
BRAND_DARK = colors.HexColor("#111111")     # чёрный
BRAND_ON_YELLOW = colors.HexColor("#111111")  # текст на оранжевом (чёрный)
DARK = colors.HexColor("#111111")
GRAY = colors.HexColor("#666666")
LIGHT_GRAY = colors.HexColor("#F5F5F5")
SOFT_YELLOW = colors.HexColor("#FFF1E5")     # мягкая оранжевая заливка для итогов
BORDER = colors.HexColor("#DDDDDD")
ROSE = SOFT_YELLOW  # backward compatibility


# ---------------------------------------------------------
# Регистрация шрифтов
# ---------------------------------------------------------
_FONT_REGISTERED = False


def _find_matplotlib_dejavu() -> tuple[Optional[Path], Optional[Path]]:
    """Matplotlib всегда везёт DejaVu Sans с собой (в mpl-data/fonts/ttf/)."""
    try:
        import matplotlib
        mpl_dir = Path(matplotlib.get_data_path()) / "fonts" / "ttf"
        reg = mpl_dir / "DejaVuSans.ttf"
        bold = mpl_dir / "DejaVuSans-Bold.ttf"
        if reg.exists() and bold.exists():
            return reg, bold
    except Exception:
        pass
    return None, None


def _register_fonts() -> tuple[str, str]:
    """Регистрируем кириллический шрифт близкий к SF UI Display Light.

    Приоритет выбора:
    1. Inter Light (CDN Google Fonts) — максимально близко к SF UI Display Light.
    2. Manrope Light (CDN).
    3. Noto Sans / Roboto системные.
    4. Matplotlib DejaVu — fallback.
    5. fonts/ в репозитории.
    6. Helvetica — крайний случай.
    """
    global _FONT_REGISTERED
    reg_name = "AppSans"
    bold_name = "AppSans-Bold"

    if _FONT_REGISTERED:
        return reg_name, bold_name

    reg_path, bold_path = None, None

    # 1. Локальный SF UI Display Light/Medium — основной шрифт
    sf_light = FONTS_DIR / "SFUIDisplay-Light.ttf"
    sf_medium = FONTS_DIR / "SFUIDisplay-Medium.ttf"
    if sf_light.exists() and sf_light.stat().st_size > 20_000:
        reg_path = sf_light
        bold_path = sf_medium if sf_medium.exists() else sf_light

    # 2. Inter/Manrope — fallback через CDN
    if reg_path is None:
     try:
        import urllib.request
        tmp_dir = Path("/tmp")
        candidates_cdn = [
            ("https://cdn.jsdelivr.net/fontsource/fonts/inter@latest/cyrillic-300-normal.ttf",
             "https://cdn.jsdelivr.net/fontsource/fonts/inter@latest/cyrillic-500-normal.ttf",
             "Inter-Light.ttf", "Inter-Medium.ttf"),
            ("https://cdn.jsdelivr.net/fontsource/fonts/manrope@latest/cyrillic-300-normal.ttf",
             "https://cdn.jsdelivr.net/fontsource/fonts/manrope@latest/cyrillic-500-normal.ttf",
             "Manrope-Light.ttf", "Manrope-Medium.ttf"),
            ("https://cdn.jsdelivr.net/fontsource/fonts/roboto@latest/cyrillic-300-normal.ttf",
             "https://cdn.jsdelivr.net/fontsource/fonts/roboto@latest/cyrillic-500-normal.ttf",
             "Roboto-Light.ttf", "Roboto-Medium.ttf"),
        ]
        for url_r, url_b, fname_r, fname_b in candidates_cdn:
            dst_r = tmp_dir / fname_r
            dst_b = tmp_dir / fname_b
            try:
                if not dst_r.exists() or dst_r.stat().st_size < 20_000:
                    urllib.request.urlretrieve(url_r, str(dst_r))
                if not dst_b.exists() or dst_b.stat().st_size < 20_000:
                    urllib.request.urlretrieve(url_b, str(dst_b))
                if dst_r.stat().st_size > 20_000:
                    reg_path = dst_r
                    bold_path = dst_b if dst_b.stat().st_size > 20_000 else dst_r
                    break
            except Exception:
                continue
     except Exception:
        pass

    # 2. Локальные fonts/ (SF UI Display Light/Medium — приоритет)
    if reg_path is None:
        candidates = [
            (FONTS_DIR / "SFUIDisplay-Light.ttf", FONTS_DIR / "SFUIDisplay-Medium.ttf"),
            (FONTS_DIR / "Inter-Light.ttf", FONTS_DIR / "Inter-Medium.ttf"),
            (FONTS_DIR / "Manrope-Light.ttf", FONTS_DIR / "Manrope-Medium.ttf"),
            (Path("/usr/share/fonts/truetype/noto/NotoSans-Light.ttf"),
             Path("/usr/share/fonts/truetype/noto/NotoSans-Medium.ttf")),
            (Path("/usr/share/fonts/truetype/roboto/Roboto-Light.ttf"),
             Path("/usr/share/fonts/truetype/roboto/Roboto-Medium.ttf")),
            (Path("/System/Library/Fonts/SFNS.ttf"),
             Path("/System/Library/Fonts/SFNS.ttf")),  # macOS SF
            (FONTS_DIR / "DejaVuSans.ttf", FONTS_DIR / "DejaVuSans-Bold.ttf"),
            (Path("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"),
             Path("/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf")),
            (Path("/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf"),
             Path("/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf")),
        ]
        for reg_p, bold_p in candidates:
            if reg_p.exists() and reg_p.stat().st_size > 20_000:
                reg_path = reg_p
                bold_path = bold_p if bold_p.exists() else reg_p
                break

    # 3. Matplotlib DejaVu — крайний fallback
    if reg_path is None:
        reg_path, bold_path = _find_matplotlib_dejavu()

    if reg_path is None:
        return "Helvetica", "Helvetica-Bold"

    try:
        pdfmetrics.registerFont(TTFont(reg_name, str(reg_path)))
        pdfmetrics.registerFont(TTFont(bold_name, str(bold_path)))
        # Связываем варианты, чтобы <b>...</b> в Paragraph переключал шрифт автоматически
        pdfmetrics.registerFontFamily(
            reg_name, normal=reg_name, bold=bold_name,
            italic=reg_name, boldItalic=bold_name)
        # Fallback-шрифт для глифов, которых нет в SF UI Display (например №)
        _fb_path = FONTS_DIR / "DejaVuSans-Fallback.ttf"
        if _fb_path.exists():
            try:
                pdfmetrics.registerFont(TTFont("Fallback", str(_fb_path)))
            except Exception:
                pass
        _FONT_REGISTERED = True
        return reg_name, bold_name
    except Exception:
        return "Helvetica", "Helvetica-Bold"


# ---------------------------------------------------------
# Утилиты
# ---------------------------------------------------------
def fmt_money(v: float) -> str:
    return f"{v:,.2f}".replace(",", " ").replace(".", ",") + " ₽"


def fmt_money_plain(v: float) -> str:
    return f"{v:,.2f}".replace(",", " ").replace(".", ",")


# ---------------------------------------------------------
# Шаблон страницы (колонтитулы)
# ---------------------------------------------------------
def _on_page(canvas, doc, supplier: dict):
    canvas.saveState()
    reg, bold = _register_fonts()

    # Логотип слева вверху
    logo_path = MEDIA_DIR / "logo.jpg"
    if logo_path.exists():
        # Логотип с пропорцией 800:149 → рисуем шириной 75мм
        logo_w = 75 * mm
        logo_h = logo_w * 149 / 800
        canvas.drawImage(str(logo_path), 20 * mm, A4[1] - 15 * mm - logo_h / 2,
                        width=logo_w, height=logo_h,
                        preserveAspectRatio=True, mask="auto")

    # Контакты справа
    canvas.setFont(reg, 8)
    canvas.setFillColor(DARK)
    canvas.drawRightString(A4[0] - 20 * mm, A4[1] - 12 * mm,
                          f"{supplier['phone']}")
    canvas.setFillColor(GRAY)
    canvas.drawRightString(A4[0] - 20 * mm, A4[1] - 16 * mm,
                          f"{supplier['email']}")
    canvas.drawRightString(A4[0] - 20 * mm, A4[1] - 20 * mm,
                          "rolls-kran.ru")

    # Жёлтая фирменная линия под шапкой — опущена ниже логотипа
    canvas.setStrokeColor(BRAND)
    canvas.setLineWidth(2.5)
    canvas.line(20 * mm, A4[1] - 32 * mm, A4[0] - 20 * mm, A4[1] - 32 * mm)

    # footer
    _foot_parts = [supplier['short'], f"ИНН {supplier['inn']}",
                   f"КПП {supplier['kpp']}"]
    if supplier.get('ogrn'):
        _foot_parts.append(f"ОГРН {supplier['ogrn']}")
    footer_text = "  ·  ".join(_foot_parts)
    canvas.drawCentredString(A4[0] / 2, 12 * mm, footer_text)
    canvas.setFillColor(GRAY)
    canvas.setFont(reg, 7)
    canvas.drawCentredString(A4[0] / 2, 8 * mm, supplier["address"])

    # номер страницы
    canvas.setFont(reg, 8)
    canvas.setFillColor(GRAY)
    canvas.drawRightString(A4[0] - 20 * mm, 8 * mm, f"Стр. {doc.page}")

    canvas.restoreState()


# ---------------------------------------------------------
# Основная функция
# ---------------------------------------------------------
def build_kp_pdf(
    q,  # QuoteData
    kp_number: str,
    buyer_name: str,
    supplier: dict,
    series_descriptions: dict,
    get_crane_image_fn,
    get_hoist_image_fn,
    crane_characteristics_fn,
    hoist_characteristics_fn,
    prepayment_rate: float = 0.70,
    kp_date: str = None,
    **kwargs,  # толерантность к лишним кейвордам
) -> bytes:
    """Строит PDF КП. Логика построения — общая с DOCX."""
    reg, bold = _register_fonts()

    styles = getSampleStyleSheet()

    # Шрифты — в стиле Apple SF Pro: тонкие, воздушные, без чрезмерной жирности.
    # Заголовки — bold, остальное в обычном весе с широким leading.
    # Шрифты: без жирности; таблицы тело 12 pt, заголовки 14 pt.
    style_title = ParagraphStyle(
        "Title", parent=styles["Title"], fontName=reg, fontSize=24,
        textColor=DARK, alignment=TA_CENTER, spaceAfter=8, leading=30,
    )
    style_subtitle = ParagraphStyle(
        "Subtitle", fontName=reg, fontSize=11, textColor=GRAY,
        alignment=TA_CENTER, spaceAfter=4, leading=15,
    )
    style_h1 = ParagraphStyle(
        "H1", fontName=reg, fontSize=20, textColor=DARK,
        alignment=TA_CENTER, spaceAfter=16, spaceBefore=4, leading=26,
    )
    style_h2 = ParagraphStyle(
        "H2", fontName=reg, fontSize=14, textColor=DARK,
        alignment=TA_CENTER, spaceAfter=10, spaceBefore=12, leading=18,
    )
    style_body = ParagraphStyle(
        "Body", fontName=reg, fontSize=12, textColor=DARK,
        alignment=TA_JUSTIFY, spaceAfter=6, leading=17,
    )
    style_body_center = ParagraphStyle(
        "BodyC", fontName=reg, fontSize=12, textColor=DARK,
        alignment=TA_CENTER, spaceAfter=6, leading=17,
    )
    style_small_gray = ParagraphStyle(
        "SmallGray", fontName=reg, fontSize=11, textColor=GRAY,
        alignment=TA_CENTER, spaceAfter=2, leading=16,
    )
    style_hero_series = ParagraphStyle(
        "Hero", fontName=reg, fontSize=16, textColor=DARK,
        alignment=TA_CENTER, spaceAfter=4, leading=22,
    )
    style_hero_sub = ParagraphStyle(
        "HeroSub", fontName=reg, fontSize=12, textColor=GRAY,
        alignment=TA_CENTER, spaceAfter=12, leading=16,
    )
    style_kp_meta = ParagraphStyle(
        "KpMeta", fontName=reg, fontSize=11, textColor=DARK,
        alignment=TA_CENTER, spaceAfter=6, leading=16,
    )

    # buffer
    buf = io.BytesIO()
    doc = BaseDocTemplate(
        buf, pagesize=A4,
        leftMargin=20 * mm, rightMargin=20 * mm,
        topMargin=38 * mm, bottomMargin=25 * mm,  # контент под линией (32мм)
    )
    frame = Frame(doc.leftMargin, doc.bottomMargin,
                  doc.width, doc.height, id="normal")
    doc.addPageTemplates([
        PageTemplate(id="main", frames=frame,
                     onPage=lambda c, d: _on_page(c, d, supplier)),
    ])

    story = []

    # -------- Титул --------
    story.append(Paragraph("КОММЕРЧЕСКОЕ ПРЕДЛОЖЕНИЕ", style_title))
    kp_date_str = kp_date or date.today().strftime('%d.%m.%Y')
    story.append(Paragraph(
        f"КП {kp_number}  ·  от {kp_date_str}",
        style_subtitle))
    if buyer_name:
        story.append(Paragraph(f"Кому:  <b>{buyer_name}</b>",
                              style_kp_meta))
    # Отступ перед блоком модели — чтобы он сместился чуть ниже
    story.append(Spacer(1, 28))

    # Для LLL-кранов (ЛКС73М) показываем правильное имя серии и корректные стрелу/высоту из кода
    _use_lll_title = bool(getattr(q, "use_lllm", False) and getattr(q, "lllm_code", None))
    if _use_lll_title:
        import re as _re_ttl
        _m_ttl = _re_ttl.match(r"ЛКС73М\.\d+-(\d+)-(\d+)\.LLL", q.lllm_code)
        _title_boom = int(_m_ttl.group(1)) if _m_ttl else q.boom
        _title_height = int(_m_ttl.group(2)) if _m_ttl else q.height_to_arm
        _title_series = "ЛКС73М"
        # Описание типа: как у ЛКС73 + пометка про электроповорот
        _hero_series_txt = (series_descriptions.get("ЛКС73", "")
                            + " (с электрическим поворотом стрелы)")
    else:
        _title_boom = q.boom
        _title_height = q.height_to_arm
        _title_series = q.series
        _hero_series_txt = series_descriptions[q.series]

    story.append(Paragraph(_hero_series_txt, style_hero_series))
    extra = ""
    if _title_series in ("ЛКС73", "ЛКС73М", "ЛКС78"):
        extra = f" · высота {_title_height} м"
    if q.series == "VACUTEC":
        # Если в КП несколько траверс (варианты) — перечисляем их г/п
        _extras_for_sub = kwargs.get("traverse_extras") or []
        _caps = [q.capacity]
        for _ex in _extras_for_sub:
            _sel = _ex.get("selection") if isinstance(_ex, dict) else None
            if _sel is not None:
                _cap = getattr(_sel, "capacity", None)
                if _cap is None:
                    # Парсим из base_code вида "VacuTec 8P-760"
                    import re as _re
                    _m = _re.search(r"-(\d+)", getattr(_sel, "base_code", ""))
                    if _m:
                        _cap = int(_m.group(1))
                if _cap:
                    _caps.append(_cap)
        # Убираем дубли, сохраняем порядок
        _seen = set(); _caps_unique = []
        for _c in _caps:
            if _c not in _seen:
                _seen.add(_c); _caps_unique.append(_c)
        if len(_caps_unique) > 1:
            _caps_str = " и ".join(f"г/п {int(_c)} кг" for _c in _caps_unique)
            sub = f"серии VacuTec · {_caps_str}"
        else:
            sub = f"серии VacuTec · г/п {q.capacity} кг"
    else:
        sub = f"серии {_title_series} · г/п {q.capacity} кг · стрела {_title_boom} м{extra}"
    story.append(Paragraph(sub, style_hero_sub))
    story.append(Spacer(1, 10))

    # Фото крана на титульной странице — крупно, на всю ширину
    img_path = get_crane_image_fn(q.series)
    if img_path:
        p = Path(str(img_path))
        if p.exists() and p.stat().st_size > 500:
            try:
                img = Image(str(p))
                iw, ih = img.wrap(0, 0)
                target_w = 140 * mm
                # Для траверсы главное фото меньше — чтобы карточки уместились ниже
                target_h = 100 * mm if q.series == "VACUTEC" else 100 * mm
                scale = min(target_w / iw, target_h / ih)
                img.drawWidth = iw * scale
                img.drawHeight = ih * scale
                img.hAlign = "CENTER"
                story.append(Spacer(1, 4))
                story.append(img)
                story.append(Spacer(1, 8))
            except Exception as e:
                print(f"[KP PDF] Не удалось вставить фото крана: {e}")

    story.append(Spacer(1, 6))
    # Карточки узлов — только в КП на саму траверсу (в комбо с краном не выводим)
    _show_traverse_cards = (q.series == "VACUTEC")
    # В комбо (кран + траверса) на титуле — текст про S355 выводится ниже (как у обычного КП на кран);
    # карточки узлов в комбо не нужны.
    if _show_traverse_cards:
        story.append(Spacer(1, 34))
        # Блок 6 карточек с узлами траверсы
        features_dir = MEDIA_DIR / "traverse" / "features"
        cards = [
            ("filter.jpg", "механический вакуумметр в поле зрения оператора с красно-зелёной шкалой"),
            ("klapan.jpg", "акустическая и оптическая система оповещения при снижении вакуума"),
            ("nasos.jpg", "укомплектованы вакуумными насосами необходимой производительности"),
            ("manometr.jpg", "фильтры-влагоотделители коалесцентного типа"),
            ("beam.jpg", "несущая балка выполняет функцию ресивера для обеспечения безопасности"),
            ("alert.jpg", "обратный клапан предотвращает утечки вакуума"),
        ]
        card_style = ParagraphStyle(
            "FeatCard", fontName=reg, fontSize=6.5,
            textColor=DARK, alignment=TA_CENTER, leading=8)
        # Строим 3×2 грид — ширина селлы в A4 минус отступы: 55мм
        col_w = 28 * mm
        img_row_h = 20 * mm  # фиксированная высота строки с фото
        def _make_img_cell(fn):
            p = features_dir / fn
            if p.exists():
                img = Image(str(p))
                iw, ih = img.wrap(0, 0)
                w = col_w - 4 * mm
                scale = min(w / iw, img_row_h / ih)
                img.drawWidth = iw * scale
                img.drawHeight = ih * scale
                img.hAlign = "CENTER"
                return img
            return ""
        img_row = [_make_img_cell(fn) for fn, _ in cards]
        txt_row = [Paragraph(cap, card_style) for _, cap in cards]
        feat_table = Table([img_row, txt_row], colWidths=[col_w] * 6,
                           rowHeights=[img_row_h + 4 * mm, None])
        feat_table.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 0.4, BORDER),
            ("INNERGRID", (0, 0), (-1, -1), 0.4, BORDER),
            # Фото — по верхнему уровню; текст — по верху (все чашки на одной линии)
            ("VALIGN", (0, 0), (-1, 0), "TOP"),
            ("VALIGN", (0, 1), (-1, 1), "TOP"),
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("LEFTPADDING", (0, 0), (-1, -1), 2),
            ("RIGHTPADDING", (0, 0), (-1, -1), 2),
            ("TOPPADDING", (0, 0), (-1, 0), 4),
            ("BOTTOMPADDING", (0, 0), (-1, 0), 2),
            ("TOPPADDING", (0, 1), (-1, 1), 3),
            ("BOTTOMPADDING", (0, 1), (-1, 1), 4),
            # Граница между фото и текстом внутри ячейки — убираем верт. границу между строками
            ("LINEBELOW", (0, 0), (-1, 0), 0.4, BORDER),
        ]))
        feat_table.hAlign = "CENTER"
        story.append(feat_table)
    else:
        intro = (
            "Легкие крановые системы созданы на базе нового уникального сверхпрочного "
            "профиля. Холоднокатаные профили разработаны и выполнены из специальной "
            "стали S355. Использование специального профиля и высококачественных "
            "тележек обеспечивает лёгкость хода при ручном перемещении весь срок "
            "службы, усилие перемещения не превышает 1% переносимого груза."
        )
        story.append(Paragraph(intro, style_small_gray))

    story.append(PageBreak())

    # -------- Характеристики --------
    story.append(Paragraph("Технические характеристики", style_h1))

    # Блок крана — в KeepTogether, чтобы не разбился.
    # Картинка крана — шире (ландшафтная), тали — вертикальная.
    crane_block = [
        Paragraph(
            "Вакуумная траверса VacuTec" if q.series == "VACUTEC"
            else f"Кран {q.series}", style_h2),
        _char_block(crane_characteristics_fn(q),
                   get_crane_image_fn(q.series), reg, bold,
                   image_w=60 * mm, image_max_h=90 * mm),
    ]
    story.append(KeepTogether(crane_block))

    # --- Габаритный чертёж крана отдельной страницей после ТХ ---
    try:
        import drawings as _drw
        _drawing_path = _drw.find_crane_drawing(
            series=q.series, capacity=q.capacity,
            boom=q.boom, height_to_arm=q.height_to_arm,
            use_lllm=bool(getattr(q, "use_lllm", False)),
            lllm_code=str(getattr(q, "lllm_code", "") or ""),
        )
        if _drawing_path and _drawing_path.suffix.lower() != ".pdf":
            story.append(PageBreak())
            story.append(Paragraph(
                f"Габаритный чертёж крана {q.series} "
                f"{int(q.capacity)} кг · {q.boom:g}×{q.height_to_arm:g} м",
                style_h1))
            story.append(Spacer(1, 8))
            _dr_img = Image(str(_drawing_path), width=170*mm, height=200*mm,
                            kind="proportional")
            _dr_img.hAlign = "CENTER"
            story.append(_dr_img)
    except Exception:
        pass

    story.append(Spacer(1, 8))

    # Предупреждение по фундаменту — для ЛКС73 и ЛКС73М LLL (колонные краны).
    # Показываем ТОЛЬКО при включённом монтаже, и НЕ в режиме «Кран + траверса».
    _is_column_crane = (q.series == "ЛКС73"
                        or bool(getattr(q, "use_lllm", False) and getattr(q, "lllm_code", None)))
    _is_combo_mode = bool(getattr(q, "with_traverse", False))
    _needs_foundation_warn = _is_column_crane and q.include_montage and not _is_combo_mode
    if _needs_foundation_warn:
        _warn_title_style = ParagraphStyle(
            "WarnTitle", fontName=bold, fontSize=8, leading=10,
            textColor=colors.HexColor("#B85C0F"),
            spaceAfter=2, alignment=TA_LEFT)
        _warn_body_style = ParagraphStyle(
            "WarnBody", fontName=reg, fontSize=6.8, leading=8.4,
            textColor=DARK, spaceAfter=2, alignment=TA_JUSTIFY)
        _warn_cell = [
            Paragraph(
                '<font color="#F97316">!</font>&nbsp;&nbsp;'
                'ПРИМЕЧАНИЕ! Ответственность за фундамент',
                _warn_title_style),
            Paragraph(
                "Заказчик несёт полную ответственность за состояние, несущую способность, "
                "геометрические размеры и соответствие проектным нагрузкам того фундамента, "
                "на который будет производиться монтаж консольного крана.",
                _warn_body_style),
            Paragraph(
                "В случае, если фундамент требуется изготовить заново или усилить — все работы "
                "по его устройству, включая разработку проекта, закупку материалов, бетонирование "
                "и контроль качества, выполняются силами и за счёт Заказчика. Все работы по "
                "устройству, ремонту или усилению фундамента (земляные, арматурные, бетонные, "
                "гидроизоляционные) должны производиться специализированными организациями, имеющими "
                "допуски к данным видам работ, в строгом соответствии с действующими строительными "
                "нормами (СП 63.13330, СП 22.13330 и др.). Заказчик обязан обеспечить надлежащий "
                "технический надзор за выполнением этих работ.",
                _warn_body_style),
        ]
        _warn_table = Table([[_warn_cell]], colWidths=[170 * mm])
        _warn_table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1),
             colors.HexColor("#FFF4E6")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#F4A65A")),
            ("LEFTPADDING", (0, 0), (-1, -1), 8),
            ("RIGHTPADDING", (0, 0), (-1, -1), 8),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
        ]))
        story.append(KeepTogether(_warn_table))
        story.append(Spacer(1, 8))

    # Подробное описание траверсы: либо для КП-траверсы, либо в комбо (кран + траверса)
    description_fn = kwargs.get("description_fn")
    _show_traverse_desc = (
        q.series == "VACUTEC" or kwargs.get("attach_traverse_desc", False))
    _combo_mode = kwargs.get("attach_traverse_desc", False) and q.series != "VACUTEC"
    if _show_traverse_desc and description_fn:
        try:
            desc = description_fn(q)
        except Exception:
            desc = None
        if desc:
            # В комбо-режиме — отдельная страница для траверсы (свой заголовок и описание)
            if _combo_mode:
                story.append(PageBreak())
                story.append(Paragraph("Технические характеристики", style_h1))
                story.append(Paragraph("Вакуумная траверса VacuTec", style_h2))
                # Карточка траверсы с фото и характеристиками
                tv_chars_fn = kwargs.get("traverse_characteristics_fn")
                tv_img_fn = kwargs.get("traverse_image_fn")
                tv_chars = tv_chars_fn(q) if tv_chars_fn else None
                tv_img = tv_img_fn(q) if tv_img_fn else None
                if tv_chars:
                    story.append(_char_block(tv_chars, tv_img, reg, bold,
                                              image_w=60 * mm, image_max_h=90 * mm))
                    story.append(Spacer(1, 8))
            desc_intro_style = ParagraphStyle(
                "DescIntro", fontName=reg, fontSize=9.5, leading=12,
                textColor=DARK, spaceAfter=4, alignment=TA_JUSTIFY)
            desc_bullet_style = ParagraphStyle(
                "DescBullet", fontName=reg, fontSize=9, leading=11.5,
                textColor=DARK, spaceAfter=1,
                leftIndent=10, bulletIndent=0, alignment=TA_JUSTIFY)
            intro_text, bullets = desc[0], desc[1:]
            # Первая буква — заглавная
            if intro_text and intro_text[:1].islower():
                intro_text = intro_text[:1].upper() + intro_text[1:]
            story.append(Paragraph(intro_text, desc_intro_style))
            for b in bullets:
                story.append(Paragraph(b, desc_bullet_style, bulletText="•"))

    # --- Альтернативные траверсы (сравнение) — каждая на своей странице ---
    tv_extras = kwargs.get("traverse_extras") or []
    # Совместимость: если передан traverse2_selection — добавляем к extras
    tv2_sel_legacy = kwargs.get("traverse2_selection")
    if tv2_sel_legacy and not tv_extras:
        tv_extras = [{
            "selection": tv2_sel_legacy,
            "image_fn": kwargs.get("traverse2_image_fn"),
            "characteristics_fn": kwargs.get("traverse2_characteristics_fn"),
            "description_fn": kwargs.get("traverse2_description_fn"),
        }]
    desc_intro_style2 = ParagraphStyle(
        "DescIntro2", fontName=reg, fontSize=9.5, leading=12,
        textColor=DARK, spaceAfter=4, alignment=TA_JUSTIFY)
    desc_bullet_style2 = ParagraphStyle(
        "DescBullet2", fontName=reg, fontSize=9, leading=11.5,
        textColor=DARK, spaceAfter=1,
        leftIndent=10, bulletIndent=0, alignment=TA_JUSTIFY)
    for _idx, _extra in enumerate(tv_extras, start=1):
        _e_sel = _extra.get("selection")
        _e_chars_fn = _extra.get("characteristics_fn")
        _e_img_fn = _extra.get("image_fn")
        _e_desc_fn = _extra.get("description_fn")
        if not (_e_sel and _e_chars_fn):
            continue
        story.append(PageBreak())
        story.append(Paragraph(
            f'Вариант <font name="Fallback">№</font> {_idx + 1}',
            style_h1))
        story.append(Paragraph("Вакуумная траверса VacuTec", style_h2))
        _e_chars = _e_chars_fn(q) or []
        _e_img = _e_img_fn(q) if _e_img_fn else None
        if _e_chars:
            story.append(_char_block(
                _e_chars, _e_img, reg, bold,
                image_w=60 * mm, image_max_h=90 * mm))
            story.append(Spacer(1, 8))
        _e_desc = _e_desc_fn(q) if _e_desc_fn else None
        if _e_desc:
            _intro2 = _e_desc[0]
            if _intro2 and _intro2[:1].islower():
                _intro2 = _intro2[:1].upper() + _intro2[1:]
            story.append(Paragraph(_intro2, desc_intro_style2))
            for _b in _e_desc[1:]:
                story.append(Paragraph(_b, desc_bullet_style2, bulletText="•"))

    # Блок тали — пропускаем если нет характеристик (напр. КП на траверсу)
    _hoist_chars = hoist_characteristics_fn(q)
    if _hoist_chars:
        hoist_block = [
            Paragraph(f"Электротельфер (таль) {q.hoist_brand}", style_h2),
            _char_block(_hoist_chars,
                       get_hoist_image_fn(q.hoist_brand, getattr(q, "hoist_exec", "") or ""), reg, bold,
                       image_w=45 * mm, image_max_h=110 * mm),
        ]
        story.append(KeepTogether(hoist_block))

    story.append(PageBreak())

    # -------- Спецификация --------
    story.append(Paragraph("Спецификация и расчёт стоимости", style_h1))
    _vat_label_main = "с НДС 22 %" if getattr(q, "include_vat", True) else "без НДС"
    story.append(_spec_table(q.lines, reg, bold, vat_label=_vat_label_main))

    # Электрификация — заголовок + описание + таблица держать вместе
    if q.electrification_lines:
        elec_block = [
            Spacer(1, 10),
            Paragraph(
                "Дополнительная комплектация — электрификация", style_h2),
            Paragraph(
                f"Опциональный пакет полной электрификации перемещения тали по стреле "
                f"для {q.series} г/п {q.capacity} кг.",
                ParagraphStyle("Italic", fontName=reg, fontSize=9,
                              textColor=GRAY, alignment=TA_LEFT, spaceAfter=6),
            ),
            _spec_table(q.electrification_lines, reg, bold, vat_label=_vat_label_main),
        ]
        story.append(KeepTogether(elec_block))

    # Стоимость поставки — сразу после спецификации/электрификации,
    # МОНТАЖ в неё не входит (поставка с НДС отдельно).
    supply_total = sum(ln.total for ln in q.lines) + \
                   sum(ln.total for ln in q.electrification_lines)
    story.append(Spacer(1, 8))
    totals = [
        ["Стоимость поставки с НДС 22 %:", fmt_money(supply_total)],
        [f"Предоплата {round(prepayment_rate*100)} %:",
         fmt_money(supply_total * prepayment_rate)],
        [f"Остаток {round((1-prepayment_rate)*100)} % перед отгрузкой:",
         fmt_money(supply_total * (1 - prepayment_rate))],
    ]
    # Та же ширина (170мм), что и у спецификации,
    # чтобы выравнивалось в одну линию.
    # Левый край выравнен с таблицей спецификации (та же 170мм)
    t = Table(totals, colWidths=[130 * mm, 40 * mm])
    t.hAlign = "LEFT"
    t.setStyle(TableStyle([
        ("FONTNAME", (0, 0), (-1, -1), reg),
        ("FONTNAME", (0, 0), (-1, 0), bold),
        ("FONTSIZE", (0, 0), (-1, -1), 10),
        ("FONTSIZE", (0, 0), (-1, 0), 11),
        ("TEXTCOLOR", (1, 0), (1, 0), DARK),
        ("ALIGN", (0, 0), (0, -1), "RIGHT"),
        ("ALIGN", (1, 0), (1, -1), "RIGHT"),
        ("BACKGROUND", (0, 0), (-1, 0), SOFT_YELLOW),
        ("LINEABOVE", (0, 0), (-1, 0), 1.2, BRAND_DARK),
        ("BOX", (0, 0), (-1, -1), 0.5, BORDER),
        ("INNERGRID", (0, 0), (-1, -1), 0.25, BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
    ]))
    story.append(t)

    # Подпись о скидке — если discount_pct > 0
    _disc_pct = float(kwargs.get("discount_pct") or 0)
    if _disc_pct > 0:
        _disc_style = ParagraphStyle(
            "DiscNote", fontName=bold, fontSize=11, leading=14,
            textColor=BRAND_DARK, alignment=TA_LEFT,
            spaceBefore=8, spaceAfter=4)
        story.append(Paragraph(
            f"Вам предоставлена скидка {_disc_pct:.0f} % на все позиции КП.",
            _disc_style))

    # -------- МОНТАЖ — ВСЕГДА ОТДЕЛЬНОЙ СТРАНИЦЕЙ --------
    if q.include_montage and q.montage_price > 0:
        story.append(PageBreak())
        montage_vat = getattr(q, "montage_vat", False)
        h2_montage = ("Монтаж и пусконаладка с НДС 22 %"
                      if montage_vat else "Монтаж и пусконаладка без НДС")
        story.append(Paragraph(h2_montage, style_h1))

        class _Line:
            def __init__(self, c, n, u, qty, p, t):
                self.code, self.name, self.unit = c, n, u
                self.qty, self.price, self.total = qty, p, t

        m_price = float(q.montage_price)
        ml = _Line("—", "Монтаж и пусконаладочные работы консольного крана",
                   "усл.", 1, m_price, m_price)
        _mnt_vat_label = "с НДС 22 %" if montage_vat else "без НДС"
        story.append(_spec_table([ml], reg, bold, vat_label=_mnt_vat_label))

        if montage_vat:
            note_text = ("Стоимость услуг монтажа включает НДС 22 %. "
                         "Гарантия на монтажные работы — 12 месяцев. "
                         "При заказе монтажа в нашей компании вы получаете "
                         "расширенную гарантию 24 месяца на комплектующие крановой системы "
                         "(электрическая таль, вакуумная траверса и пакет электрификации "
                         "в расширенную гарантию не входят).")
        else:
            note_text = ("Стоимость услуг монтажа (предварительная): "
                         "рассчитывается без учёта НДС по отдельному договору с нашей компанией. "
                         "При заказе монтажа в нашей компании вы получаете расширенную "
                         "гарантию 24 месяца на комплектующие крановой системы "
                         "(электрическая таль, вакуумная траверса и пакет электрификации "
                         "в расширенную гарантию не входят). Гарантия на монтажные работы — 12 месяцев.")
        story.append(Paragraph(
            note_text,
            ParagraphStyle("MontageNote", fontName=reg, fontSize=10,
                          textColor=DARK, alignment=TA_JUSTIFY,
                          spaceBefore=6, spaceAfter=6, leading=14),
        ))

        # Состав работ по монтажу и ПНР
        story.append(Spacer(1, 8))
        story.append(Paragraph(
            "<b>Состав работ по монтажу и ПНР консольного крана</b>",
            ParagraphStyle("MontageWorksH", fontName=reg, fontSize=11,
                          textColor=DARK, alignment=TA_LEFT,
                          spaceBefore=4, spaceAfter=4, leading=14),
        ))
        story.append(Paragraph(
            "Мы выполняем полный цикл установки на подготовленный фундамент "
            "(армированный бетонный пол от 300 мм). В фиксированную цену "
            "включены все сопутствующие затраты:",
            ParagraphStyle("MontageWorks", fontName=reg, fontSize=10,
                          textColor=DARK, alignment=TA_JUSTIFY,
                          spaceBefore=2, spaceAfter=4, leading=14),
        ))
        bullets = [
            "Монтаж металлоконструкций крана;",
            "Подключение электротали (тельфер) к сети 380 В;",
            "Крепёж (химические анкера + шпильки класса прочности не менее 8.8) в комплекте;",
            "Пусконаладка и комплексное опробование оборудования;",
            "Официальные испытания с записью результатов в паспорт крана;",
            "Логистика и доставка специалистов.",
        ]
        for b in bullets:
            story.append(Paragraph(
                f'<font color="#F97316">■</font>&nbsp;&nbsp;{b}',
                ParagraphStyle("MntBul", fontName=reg, fontSize=10,
                              textColor=DARK, alignment=TA_LEFT,
                              leftIndent=8, leading=14, spaceAfter=2),
            ))
        story.append(Spacer(1, 4))
        story.append(Paragraph(
            "<b>За счёт Заказчика:</b> предоставление грузоподъёмного механизма "
            "(погрузчик или существующая кран-балка на несколько часов) для перемещения "
            "элементов крана в зоне работ. Контрольный груз для проведения испытаний крана предоставляет заказчик.",
            ParagraphStyle("MontageWorksEnd", fontName=reg, fontSize=10,
                          textColor=DARK, alignment=TA_JUSTIFY,
                          spaceBefore=4, spaceAfter=6, leading=14),
        ))

        # --- Примечание о фундаменте на странице монтажа (для колонных кранов) ---
        if _is_column_crane:
            _mnt_warn_title = ParagraphStyle(
                "MntWarnTitle", fontName=bold, fontSize=9, leading=11,
                textColor=colors.HexColor("#B85C0F"),
                spaceAfter=3, alignment=TA_LEFT)
            _mnt_warn_body = ParagraphStyle(
                "MntWarnBody", fontName=reg, fontSize=7.5, leading=10,
                textColor=DARK, spaceAfter=2, alignment=TA_JUSTIFY)
            _mnt_warn_cell = [
                Paragraph(
                    '<font color="#F97316">!</font>&nbsp;&nbsp;'
                    'ПРИМЕЧАНИЕ! Ответственность за фундамент',
                    _mnt_warn_title),
                Paragraph(
                    "Заказчик несёт полную ответственность за состояние, несущую "
                    "способность, геометрические размеры и соответствие проектным "
                    "нагрузкам того фундамента, на который будет производиться монтаж "
                    "консольного крана. В случае, если фундамент требуется изготовить "
                    "заново или усилить — все работы по его устройству, включая разработку "
                    "проекта, закупку материалов, бетонирование и контроль качества, "
                    "выполняются силами и за счёт Заказчика. Работы должны производиться "
                    "специализированными организациями в соответствии с СП 63.13330 и СП 22.13330.",
                    _mnt_warn_body),
            ]
            _mnt_warn_tbl = Table([[_mnt_warn_cell]], colWidths=[170*mm])
            _mnt_warn_tbl.setStyle(TableStyle([
                ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF4E6")),
                ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#F4A65A")),
                ("LEFTPADDING", (0, 0), (-1, -1), 8),
                ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                ("TOPPADDING", (0, 0), (-1, -1), 6),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ]))
            story.append(Spacer(1, 8))
            story.append(_mnt_warn_tbl)

    # -------- Условия --------
    story.append(PageBreak())
    story.append(Paragraph("Условия поставки", style_h1))

    if q.series == "VACUTEC":
        _pp = int(kwargs.get("prepay_pct", 100))
        _addr = kwargs.get("delivery_address") or "указывается покупателем"
        _delivery_price = float(kwargs.get("delivery_price") or 0)
        # Способ доставки — текст выбранный пользователем в UI; если пусто — дефолт
        _delivery_terms = (kwargs.get("delivery_terms") or
            "Доставка до ТК «Деловые линии» в г. Орёл и отправка с учётом обрешётки за счёт покупателя")
        # Адрес доставки (если указан) — в текст условий
        _delivery_target = (kwargs.get("delivery_target") or "").strip()
        if _delivery_target:
            _delivery_terms += f", адрес доставки: {_delivery_target}"
        # Если в условиях указана цена доставки — добавляем её в конец
        if _delivery_price > 0:
            _delivery_terms += (f" — стоимость доставки "
                                 f"{fmt_money(_delivery_price)}.")
        conds = [
            ("Порядок оплаты",
             f"Покупатель вносит предоплату на расчётный счёт поставщика, указанный в счёте, в размере {_pp} %."),
            ("Срок отгрузки продукции",
             "3–5 рабочих дней со дня поступления оплаты на расчётный счёт Поставщика."),
            ("Условия доставки", _delivery_terms),
            ("Гарантия", "12 месяцев."),
            ("Срок действия КП", "14 календарных дней с даты выставления."),
        ]
    else:
        # Условия оплаты и доставки — из UI (передаются через kwargs)
        _pp_kran = int(round(float(prepayment_rate) * 100))
        _rem_kran = 100 - _pp_kran
        _delivery_terms_kr = (kwargs.get("delivery_terms")
                              or "Бесплатная доставка до ТК «Деловые линии» в Санкт-Петербурге. Далее — за счёт Покупателя.")
        _delivery_target_kr = (kwargs.get("delivery_target") or "").strip()
        if _delivery_target_kr and _delivery_target_kr not in _delivery_terms_kr:
            _delivery_terms_kr += f", адрес доставки: {_delivery_target_kr}"
        _delivery_price_kr = float(kwargs.get("delivery_price") or 0)
        if _delivery_price_kr > 0:
            _delivery_terms_kr += f" — стоимость доставки {fmt_money(_delivery_price_kr)}."
        elif not _delivery_terms_kr.endswith("."):
            _delivery_terms_kr += "."
        conds = [
            ("Оплата",
             f"{_pp_kran} % предоплата после подписания договора и спецификации, "
             f"{_rem_kran} % — по уведомлению о готовности к отгрузке."),
            ("Срок изготовления", "до 20 рабочих дней после поступления предоплаты. "
                                  "Возможна досрочная поставка."),
            ("Условия доставки", _delivery_terms_kr),
            ("Гарантия", "12 месяцев на кран, 12 месяцев на таль. "
                         "24 месяца при монтаже нашей компанией."),
            ("Документы", "Паспорт крана, паспорт тали, руководство "
                          "по эксплуатации, сертификат соответствия."),
            ("Срок действия КП", "14 календарных дней с даты выставления."),
        ]
    for k, v in conds:
        # если в валюе нет точки в конце и это плейсхолдер адреса — без точки после ключа
        sep = "" if k == "Адрес доставки" else "."
        p = Paragraph(
            f'<font color="#F97316">•</font>  <font color="#111111">{k}{sep} </font>{v}',
            style_body,
        )
        story.append(p)

    story.append(Spacer(1, 12))

    # --- Комментарий к КП на оранжевом фоне ---
    _kp_comment_text = str(kwargs.get("kp_comment", "") or "").strip()
    if _kp_comment_text:
        _cmt_lbl = Paragraph(
            '<b>Комментарий</b>',
            ParagraphStyle("KPCmtLbl", fontName=bold, fontSize=11,
                          textColor=colors.HexColor("#F97316"),
                          spaceAfter=4))
        _cmt_body = Paragraph(
            _kp_comment_text.replace("\n", "<br/>"),
            ParagraphStyle("KPCmtBody", fontName=reg, fontSize=10,
                          textColor=DARK, leading=13, alignment=0))
        _cmt_tbl = Table([[_cmt_lbl], [_cmt_body]], colWidths=[170*mm])
        _cmt_tbl.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#FFF3E0")),
            ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#F97316")),
            ("LINEBEFORE", (0, 0), (0, -1), 3, colors.HexColor("#F97316")),
            ("TOPPADDING", (0, 0), (-1, -1), 6),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 6),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ]))
        story.append(_cmt_tbl)
        story.append(Spacer(1, 12))

    story.append(Spacer(1, 12))

    # Блок подписи — таблица 2×2: слева текст, справа подпись+печать
    stamp_path = MEDIA_DIR / "stamp.jpg"
    sig_path = MEDIA_DIR / "signature.jpg"

    left_cell = []
    left_cell.append(Paragraph("С уважением,",
                              ParagraphStyle("SigLbl", fontName=reg,
                                            fontSize=11, textColor=DARK,
                                            spaceAfter=6)))
    left_cell.append(Paragraph(
        f"Исполнительный директор {supplier['short']}",
        ParagraphStyle("SigPos", fontName=reg, fontSize=11,
                      textColor=DARK, spaceAfter=6)))
    left_cell.append(Paragraph(
        "Букреев Антон",
        ParagraphStyle("SigFio", fontName=reg, fontSize=13,
                      textColor=DARK, spaceAfter=12)))
    left_cell.append(Paragraph(
        f"тел. {supplier['phone']}",
        ParagraphStyle("SigC1", fontName=reg, fontSize=9,
                      textColor=GRAY, spaceAfter=2)))
    left_cell.append(Paragraph(
        supplier['email'],
        ParagraphStyle("SigC2", fontName=reg, fontSize=9,
                      textColor=GRAY)))

    # Правая ячейка — слоение подписи и печати
    right_cell = []
    if sig_path.exists():
        sig_img = Image(str(sig_path))
        iw, ih = sig_img.wrap(0, 0)
        target_w = 45 * mm
        sig_img.drawWidth = target_w
        sig_img.drawHeight = ih * (target_w / iw)
        sig_img.hAlign = "CENTER"
        right_cell.append(sig_img)
    if stamp_path.exists():
        stamp_img = Image(str(stamp_path))
        iw, ih = stamp_img.wrap(0, 0)
        target_w = 32 * mm
        stamp_img.drawWidth = target_w
        stamp_img.drawHeight = ih * (target_w / iw)
        stamp_img.hAlign = "CENTER"
        # Печать слегка наложена на подпись — отрицательный spacer
        right_cell.append(Spacer(1, -20))
        right_cell.append(stamp_img)

    sig_table = Table(
        [[left_cell, right_cell]],
        colWidths=[100 * mm, 65 * mm],
    )
    sig_table.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    story.append(sig_table)

    # ---- QR-код на Rutube-канал ----
    qr_path = MEDIA_DIR / "qr_rutube.jpg"
    if qr_path.exists():
        story.append(Spacer(1, 18))
        qr_img = Image(str(qr_path))
        iw, ih = qr_img.wrap(0, 0)
        target_w = 35 * mm
        qr_img.drawWidth = target_w
        qr_img.drawHeight = ih * (target_w / iw)
        qr_img.hAlign = "CENTER"
        caption = Paragraph(
            "<b>Наши проекты на Rutube</b><br/>"
            "Наведите камеру телефона, чтобы увидеть видео-библиотеку "
            "внедрённых крановых систем ЛКС.",
            ParagraphStyle("QR", fontName=reg, fontSize=10, leading=13,
                          textColor=DARK, alignment=TA_CENTER, spaceBefore=6),
        )
        qr_table = Table([[qr_img], [caption]], colWidths=[170 * mm])
        qr_table.setStyle(TableStyle([
            ("ALIGN", (0, 0), (-1, -1), "CENTER"),
            ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("LEFTPADDING", (0, 0), (-1, -1), 0),
            ("RIGHTPADDING", (0, 0), (-1, -1), 0),
            ("TOPPADDING", (0, 0), (-1, -1), 4),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ]))
        story.append(qr_table)

    doc.build(story)
    return buf.getvalue()


# ---------------------------------------------------------
# Вспомогательные табличные конструкторы
# ---------------------------------------------------------
def _char_block(rows: list[tuple[str, str]], img_path,
               reg: str, bold: str, image_w: float,
               image_max_h: float = 100 * mm) -> Table:
    """Блок характеристик: слева KV-таблица, справа картинка."""
    kv = _kv_table(rows, reg, bold)

    img_flow = None
    if img_path:
        p = Path(str(img_path))
        if p.exists() and p.stat().st_size > 500:
            try:
                img = Image(str(p))
                iw, ih = img.wrap(0, 0)
                # Масштабируем по меньшей из величин, чтобы сохранить пропорции
                scale = min(image_w / iw, image_max_h / ih)
                img.drawWidth = iw * scale
                img.drawHeight = ih * scale
                img.hAlign = "CENTER"
                img_flow = img
            except Exception as e:
                print(f"[KP PDF] Не удалось вставить фото {p}: {e}")

    if img_flow is None:
        return kv

    # Композит: колонки под таблицу и картинку
    left_w = 105 * mm    # как в kv (42+63)
    right_w = 165 * mm - left_w
    wrapper = Table(
        [[kv, img_flow]],
        colWidths=[left_w, right_w],
    )
    wrapper.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("ALIGN", (1, 0), (1, 0), "CENTER"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 0),
    ]))
    return wrapper


def _kv_table(rows: list[tuple[str, str]], reg: str, bold: str) -> Table:
    # Оборачиваем в Paragraph, чтобы длинные значения переносились аккуратно
    # Тело таблиц характеристик — 12 pt.
    ps_key = ParagraphStyle("KVKey", fontName=reg, fontSize=11,
                            textColor=GRAY, leading=15)
    ps_val = ParagraphStyle("KVVal", fontName=reg, fontSize=12,
                            textColor=DARK, leading=15)
    data = [[Paragraph(str(k), ps_key), Paragraph(str(v), ps_val)]
            for k, v in rows]
    t = Table(data, colWidths=[42 * mm, 63 * mm])
    style_cmds = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
        # Только тонкие линии между строками — Apple-стиль
        ("LINEBELOW", (0, 0), (-1, -2), 0.25, colors.HexColor("#EEEEEE")),
    ]
    t.setStyle(TableStyle(style_cmds))
    return t


def _spec_table(lines: list, reg: str, bold: str, vat_label: str = "с НДС 22 %"):
    """
    Красивая таблица спецификации: все числа по центру,
    одинаковая высота строк, без пустых ячеек. Итого — отдельной
    таблицей, чтобы при разрыве страницы не возникали гигантские пустые ячейки.
    Возвращает KeepTogether([body, total]) для story.
    """
    # Стили: текст (артикул, наименование) по левому, цена по правому, кол-во по центру.
    ps_name = ParagraphStyle("SpecName", fontName=reg, fontSize=10,
                             textColor=DARK, leading=13, alignment=TA_LEFT)
    ps_code = ParagraphStyle("SpecCode", fontName=reg, fontSize=10,
                             textColor=DARK, leading=13, alignment=TA_LEFT)
    ps_num = ParagraphStyle("SpecNum", fontName=reg, fontSize=10,
                            textColor=DARK, leading=13, alignment=TA_CENTER)
    ps_money = ParagraphStyle("SpecMoney", fontName=reg, fontSize=10,
                              textColor=DARK, leading=13, alignment=TA_RIGHT)
    ps_hdr = ParagraphStyle("SpecHdr", fontName=reg, fontSize=11,
                            textColor=BRAND_ON_YELLOW, leading=13,
                            alignment=TA_CENTER)

    header = [
        Paragraph("n/n", ps_hdr),
        Paragraph("Код", ps_hdr),
        Paragraph("Наименование", ps_hdr),
        Paragraph("Ед.", ps_hdr),
        Paragraph("Кол-во", ps_hdr),
        Paragraph(f"Сумма, ₽<br/>{vat_label}", ps_hdr),
    ]
    data = [header]
    for idx, ln in enumerate(lines, start=1):
        data.append([
            Paragraph(str(idx), ps_num),
            Paragraph(str(ln.code), ps_code),
            Paragraph(str(ln.name), ps_name),
            Paragraph(ln.unit, ps_num),
            Paragraph(f"{ln.qty:g}", ps_num),
            Paragraph(fmt_money_plain(ln.total) if ln.total > 0 else "—",
                      ps_money),
        ])

    # Суммарная ширина 170мм — лучше распределение:
    # кол-во и сумма шире, чтобы шапка влезала в одну строку.
    col_widths = [10 * mm, 28 * mm, 60 * mm, 14 * mm, 18 * mm, 40 * mm]

    body = Table(data, colWidths=col_widths, repeatRows=1)
    style_cmds = [
        # Шапка — оранжевая
        ("BACKGROUND", (0, 0), (-1, 0), BRAND),
        ("TEXTCOLOR", (0, 0), (-1, 0), BRAND_ON_YELLOW),
        ("FONTNAME", (0, 0), (-1, 0), reg),
        ("ALIGN", (0, 0), (-1, 0), "CENTER"),
        ("VALIGN", (0, 0), (-1, 0), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 0), 8),
        ("BOTTOMPADDING", (0, 0), (-1, 0), 8),
        # Тело — всё по центру (горизонталь в Paragraph, вертикаль — тут)
        ("VALIGN", (0, 1), (-1, -1), "MIDDLE"),
        ("GRID", (0, 0), (-1, -1), 0.4, BORDER),
        ("LEFTPADDING", (0, 0), (-1, -1), 5),
        ("RIGHTPADDING", (0, 0), (-1, -1), 5),
        ("TOPPADDING", (0, 1), (-1, -1), 6),
        ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
    ]
    # чередование — чётные строки тела сероватые
    for i in range(2, len(data), 2):
        style_cmds.append(("BACKGROUND", (0, i), (-1, i),
                          colors.HexColor("#FDF6EF")))
    body.setStyle(TableStyle(style_cmds))

    # Итого — отдельной полоской во всю ширину
    total = sum(ln.total for ln in lines)
    ps_total_label = ParagraphStyle("TotalLabel", fontName=reg, fontSize=12,
                                    textColor=DARK, leading=14,
                                    alignment=TA_RIGHT)
    ps_total_value = ParagraphStyle("TotalValue", fontName=reg, fontSize=13,
                                    textColor=DARK, leading=15,
                                    alignment=TA_CENTER)
    total_row = Table(
        [[Paragraph(f"Итого {vat_label}", ps_total_label),
          Paragraph(fmt_money_plain(total), ps_total_value)]],
        colWidths=[130 * mm, 40 * mm],
    )
    total_row.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), SOFT_YELLOW),
        ("LINEABOVE", (0, 0), (-1, 0), 1.2, BRAND_DARK),
        ("LINEBELOW", (0, 0), (-1, 0), 0.4, BORDER),
        ("LINEBEFORE", (0, 0), (0, 0), 0.4, BORDER),
        ("LINEAFTER", (-1, 0), (-1, 0), 0.4, BORDER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
    ]))

    return KeepTogether([body, total_row])
