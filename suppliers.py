"""Каталог поставщиков (ЛКС, Модернизация, Кинематика).

Каждый поставщик содержит полные реквизиты для КП/договоров + пути к печати/подписи.
"""
from pathlib import Path

STAMPS_DIR = Path(__file__).parent / "media" / "stamps"

SUPPLIERS: dict[str, dict] = {
    "LKS": {
        "label": "ООО «ЛКС»",
        "short": "ООО «ЛКС»",
        "full": "Общество с ограниченной ответственностью",  # «ЛКС» в шаблоне
        "address": "195030, г. Санкт-Петербург, вн.тер.г. муниципальный округ Ржевка, "
                   "ш. Революции, д. 114, Литера А, помещ. 111",
        "inn": "7806586802",
        "kpp": "780601001",
        "ogrn": "12178000979903",
        "rs": "40702810110000819349",
        "bank": "АО «ТИНЬКОФФ БАНК»",
        "ks": "30101810145250000974",
        "bik": "044525974",
        "phone": "8 (800) 302-73-10 — бесплатный звонок по всей РФ",
        "phone_short": "8 (800) 302-73-10",
        "email": "zakaz@rolls-kran.ru",
        "director_position": "Генеральный директор",
        "director_fio_gen": "Щербакова Ильи Витальевича",
        "director_fio_short": "Щербаков И.В.",
        "director_basis": "Устава",
        "stamp_path": str(STAMPS_DIR / "lks.jpg"),
        "edo_provider": "СБИС",
        "edo_id": "2BE2fc30491f9e94f8ebce17c64df74224b",
    },
    "MODERNIZATSIYA": {
        "label": "ООО «МОДЕРНИЗАЦИЯ»",
        "short": "ООО «МОДЕРНИЗАЦИЯ»",
        "full": 'Общество с ограниченной ответственностью "МОДЕРНИЗАЦИЯ"',
        "address": "195030, г. Санкт-Петербург, вн.тер.г. муниципальный округ Ржевка, "
                   "ш. Революции, д. 114, Литера А, помещ. 111",
        "inn": "7806601419",
        "kpp": "780601001",
        "ogrn": "1227800103706",
        "rs": "40702810501500143127",
        "bank": 'ООО "Банк Точка"',
        "ks": "30101810745374525104",
        "bik": "044525104",
        "phone": "8 (800) 302-73-10 — бесплатный звонок по всей РФ",
        "phone_short": "8 (800) 302-73-10",
        "phone_direct": "+7 (999) 069-77-07",
        "email": "zakaz@rolls-kran.ru",
        "director_position": "Генеральный директор",
        "director_fio_gen": "Букреева Антона Сергеевича",
        "director_fio_short": "Букреев А.С.",
        "director_basis": "Устава",
        "vat_note": "УСН, НДС не облагается на основании статьи 346.11 главы 26.2 НК РФ",
        "stamp_path": str(STAMPS_DIR / "modernizatsiya.jpg"),
        "edo_provider": "Компания «Тензор»",
        "edo_id": "2BE8c7a0930b20e404081ef788f122f8680",
    },
    "KINEMATIKA": {
        "label": "ООО «КИНЕМАТИКА»",
        "short": "ООО «КИНЕМАТИКА»",
        "full": 'Общество с ограниченной ответственностью "КИНЕМАТИКА"',
        "address": "195030, г. Санкт-Петербург, вн.тер.г. муниципальный округ Ржевка, "
                   "ш. Революции, д. 114, Литера А, помещ. 111",
        "inn": "6000001911",
        "kpp": "600001001",
        "ogrn": "",
        "rs": "40702810110002081379",
        "bank": "АО «Тинькофф Банк»",
        "ks": "30101810145250000974",
        "bik": "044525974",
        "phone": "8 (800) 302-73-10 — бесплатный звонок по всей РФ",
        "phone_short": "8 (800) 302-73-10",
        "phone_direct": "+7 (931) 621-70-20",
        "email": "zakaz@rolls-kran.ru",
        "director_position": "Генеральный директор",
        "director_fio_gen": "Ворсина Артёма Константиновича",
        "director_fio_short": "Ворсин А.К.",
        "director_basis": "Устава",
        "stamp_path": str(STAMPS_DIR / "kinematika.jpg"),
        "edo_provider": "Компания «Тензор»",
        "edo_id": "2BEb74705d710674cd7a46c03c77abd7b07",
    },
}

DEFAULT_SUPPLIER_KEY = "LKS"


def get_supplier(key: str) -> dict:
    """Вернуть словарь реквизитов поставщика по ключу."""
    return SUPPLIERS.get(key, SUPPLIERS[DEFAULT_SUPPLIER_KEY])


def supplier_labels() -> list[tuple[str, str]]:
    """Список [(key, label)] для селектбокса."""
    return [(k, v["label"]) for k, v in SUPPLIERS.items()]
