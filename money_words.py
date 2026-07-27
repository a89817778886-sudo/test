# -*- coding: utf-8 -*-
"""Число прописью на русском: сумма в рублях и копейках.

Использование:
    money_to_words(1234567.89)
    → "Один миллион двести тридцать четыре тысячи пятьсот шестьдесят семь рублей 89 копеек"

    money_to_words_short(1234567.89)
    → "1 234 567,89 руб. (Один миллион двести тридцать четыре тысячи пятьсот шестьдесят семь рублей 89 копеек)"
"""
from __future__ import annotations
from decimal import Decimal, ROUND_HALF_UP


_ONES_MALE = ["", "один", "два", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_ONES_FEMALE = ["", "одна", "две", "три", "четыре", "пять", "шесть", "семь", "восемь", "девять"]
_TEENS = ["десять", "одиннадцать", "двенадцать", "тринадцать", "четырнадцать",
          "пятнадцать", "шестнадцать", "семнадцать", "восемнадцать", "девятнадцать"]
_TENS = ["", "", "двадцать", "тридцать", "сорок", "пятьдесят", "шестьдесят",
         "семьдесят", "восемьдесят", "девяносто"]
_HUNDREDS = ["", "сто", "двести", "триста", "четыреста", "пятьсот",
             "шестьсот", "семьсот", "восемьсот", "девятьсот"]


def _triplet(n: int, female: bool = False) -> str:
    """Разряд от 0 до 999 прописью."""
    if n == 0:
        return ""
    words = []
    h = n // 100
    t = (n % 100) // 10
    o = n % 10
    if h:
        words.append(_HUNDREDS[h])
    if t == 1:
        words.append(_TEENS[o])
    else:
        if t:
            words.append(_TENS[t])
        if o:
            words.append(_ONES_FEMALE[o] if female else _ONES_MALE[o])
    return " ".join(words)


def _plural(n: int, forms: tuple) -> str:
    """forms = (для 1, для 2-4, для 5+)"""
    n = abs(n) % 100
    if 11 <= n <= 19:
        return forms[2]
    n = n % 10
    if n == 1:
        return forms[0]
    if 2 <= n <= 4:
        return forms[1]
    return forms[2]


def int_to_words_ru(number: int, currency_gender: str = "male") -> str:
    """Целое число прописью. Пустая строка для 0."""
    if number == 0:
        return "ноль"
    if number < 0:
        return "минус " + int_to_words_ru(-number, currency_gender)

    parts = []
    billions = number // 1_000_000_000
    millions = (number % 1_000_000_000) // 1_000_000
    thousands = (number % 1_000_000) // 1_000
    units = number % 1_000

    if billions:
        parts.append(_triplet(billions))
        parts.append(_plural(billions, ("миллиард", "миллиарда", "миллиардов")))
    if millions:
        parts.append(_triplet(millions))
        parts.append(_plural(millions, ("миллион", "миллиона", "миллионов")))
    if thousands:
        parts.append(_triplet(thousands, female=True))
        parts.append(_plural(thousands, ("тысяча", "тысячи", "тысяч")))
    if units:
        parts.append(_triplet(units, female=(currency_gender == "female")))

    return " ".join(parts)


def money_to_words(amount: float) -> str:
    """Сумма прописью: '1234567.89' → 'Один миллион ... рублей 89 копеек'.

    Копейки: цифрами (по стандарту делового оборота).
    Рубли: прописью с большой буквы, склонение (рубль/рубля/рублей).
    """
    d = Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    integer_rub = int(d)
    kop = int((d - integer_rub) * 100)

    words = int_to_words_ru(integer_rub, currency_gender="male")
    rub_word = _plural(integer_rub, ("рубль", "рубля", "рублей"))
    # Первая буква с большой
    words_cap = words[0].upper() + words[1:] if words else ""
    kop_word = _plural(kop, ("копейка", "копейки", "копеек"))
    return f"{words_cap} {rub_word} {kop:02d} {kop_word}"


def money_full(amount: float) -> str:
    """Число и цифрами, и прописью: '1 234 567,89 руб. (Один миллион ... 89 копеек)'."""
    d = Decimal(str(amount)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    integer_rub = int(d)
    kop = int((d - integer_rub) * 100)
    num_str = f"{integer_rub:,}".replace(",", " ") + f",{kop:02d}"
    return f"{num_str} руб. ({money_to_words(amount)})"


if __name__ == "__main__":
    for x in [1234567.89, 100.00, 1.01, 500000, 0.55, 21.00, 2222222.22, 1_000_000_000.99]:
        print(f"{x:>15,.2f}  →  {money_to_words(x)}")
