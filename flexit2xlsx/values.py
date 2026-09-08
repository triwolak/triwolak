# -*- coding: utf-8 -*-
"""Normalizacja wartości z XML i przygotowanie ich do zapisu w XLSX.

Moduł jest celowo **OSTROŻNY**. W plikach aukcyjnych (sprzęt, pakiety, lot-y)
znacznie częściej spotyka się numery katalogowe, wersje i oznaczenia niż liczby,
dlatego w razie jakiejkolwiek wątpliwości wartość pozostaje TEKSTEM.

Podjęte decyzje projektowe (przypadki niejednoznaczne)
-----------------------------------------------------
* ``"1,234"`` -> ``1.234`` (float).  Pojedynczy przecinek jest ZAWSZE traktowany
  jako separator dziesiętny.  Portal i użytkownik są polskojęzyczni, a w polskiej
  konwencji separatorem tysięcy jest spacja (albo kropka), nigdy przecinek.
  Przecinek jako separator tysięcy jest rozpoznawany dopiero, gdy występuje
  wielokrotnie i tworzy poprawne grupy: ``"1,234,567"`` -> ``1234567``.
* Pojedyncza kropka jest zawsze separatorem dziesiętnym: ``"1.234"`` -> ``1.234``.
  Kropka jako separator tysięcy wymaga poprawnego grupowania i co najmniej dwóch
  wystąpień: ``"1.234.567"`` -> ``1234567``.
* Wiodące zero blokuje konwersję: ``"007"``, ``"0012"``, ``"00"`` zostają tekstem
  (numery katalogowe, kody, numery kierunkowe).  ``"0"`` i ``"0,5"`` są liczbami.
* Wiodący plus NIE jest konwertowany (``"+48 123 456 789"`` to numer telefonu).
* Liczby całkowite dłuższe niż :data:`MAX_INT_DIGITS` cyfr zostają tekstem —
  Excel przechowuje liczby jako float64 i psuje takie identyfikatory
  (kody kreskowe, IMEI, numery seryjne).
* Ciągi wyglądające jak adres IPv4 (``"10.0.0.1"``) zostają tekstem, mimo że
  formalnie przechodzą test grupowania tysięcy.
* Data w formacie ``DD/MM/YYYY`` NIE jest konwertowana — nie da się jej odróżnić
  od amerykańskiego ``MM/DD/YYYY``.
* ``"tak"``/``"nie"``/``"yes"``/``"no"``/``"1"``/``"0"`` NIE są konwertowane na bool.
  Tylko dosłowne ``"true"``/``"false"`` (dowolna wielkość liter).
* Notacja wykładnicza (``"1e3"``) NIE jest obsługiwana — zbyt łatwo pomylić ją
  z kodem katalogowym typu ``"1E5"``.
* :func:`coerce_value` zachowuje offset strefy czasowej (obiekt "aware").
  Excel nie zna stref, więc dopiero :func:`sanitize_cell` sprowadza taką datę
  do UTC i zdejmuje ``tzinfo`` — moduł zapisujący MUSI wołać ``sanitize_cell``.
"""

from __future__ import annotations

import datetime as _dt
import math
import re
from decimal import Decimal, InvalidOperation
from typing import Any

__all__ = [
    "normalize_ws",
    "coerce_value",
    "sanitize_cell",
    "looks_like_formula",
    "MAX_CELL_CHARS",
    "MAX_INT_DIGITS",
    "MAX_SCALAR_LEN",
]

#: Twardy limit długości tekstu w komórce arkusza (limit Excela).
MAX_CELL_CHARS = 32767

#: Maksymalna liczba cyfr liczby całkowitej, którą jeszcze konwertujemy na ``int``.
#: Powyżej tego Excel (float64) traci precyzję, więc zostawiamy tekst.
MAX_INT_DIGITS = 15

#: Powyżej tylu znaków nie próbujemy już rozpoznawać liczby/daty/bool — żadna
#: sensowna liczba ani data nie jest tak długa, a to chroni przed dziwnymi danymi.
MAX_SCALAR_LEN = 64

#: Znaki spacji, które w zapisie liczb pełnią rolę separatora tysięcy.
_GROUP_SPACES = "\u00a0\u2007\u2009\u202f\u2008\u2002\u2003\u205f"

#: Znaki niedozwolone w pliku XLSX (poza \t \n \r) + surogaty + nie-znaki.
_ILLEGAL_XLSX_RE = re.compile(
    "[\x00-\x08\x0b\x0c\x0e-\x1f\ud800-\udfff\ufffe\uffff]"
)

_DIGITS_RE = re.compile(r"[0-9]+")
_NUM_BODY_RE = re.compile(r"[0-9., ]+")
_IPV4_RE = re.compile(r"(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})")

# Data "YYYY-M-D" (miesiąc/dzień mogą być jedno- lub dwucyfrowe).
_YMD_RE = re.compile(r"(\d{4})-(\d{1,2})-(\d{1,2})")
# Data "DD.MM.YYYY" albo "DD-MM-YYYY" (separator musi być ten sam po obu stronach).
_DMY_RE = re.compile(r"(\d{1,2})([.-])(\d{1,2})\2(\d{4})")
# ISO-8601 z czasem: "T" albo spacja, sekundy i ułamek opcjonalne, strefa opcjonalna.
_ISO_DT_RE = re.compile(
    r"(\d{4})-(\d{2})-(\d{2})[T ](\d{2}):(\d{2})"
    r"(?::(\d{2}))?"
    r"(?:[.,](\d{1,6})\d*)?"
    r"\s*(Z|z|[+-]\d{2}:?\d{2}|[+-]\d{2})?"
)

#: Największa liczba, jaką Excel jest w stanie zapisać.
_EXCEL_MAX_NUMBER = 1e308


# --------------------------------------------------------------------------- #
# Białe znaki
# --------------------------------------------------------------------------- #
def normalize_ws(text: str) -> str:
    """Zwija ciągi białych znaków do pojedynczej spacji i przycina brzegi.

    Twarda spacja (NBSP), tabulatory i znaki nowej linii są traktowane jak
    zwykłe białe znaki.  Polskie znaki diakrytyczne i emoji pozostają nietknięte.
    ``None`` (i inne typy) są obsłużone defensywnie — zwracany jest tekst.
    """
    if text is None:
        return ""
    if not isinstance(text, str):
        text = str(text)
    # str.split() bez argumentów dzieli po dowolnym białym znaku Unicode (w tym NBSP).
    return " ".join(text.split())


# --------------------------------------------------------------------------- #
# Pomocnicze — liczby
# --------------------------------------------------------------------------- #
def _all_digits(text: str) -> bool:
    """True, gdy tekst składa się wyłącznie z cyfr ASCII (i jest niepusty)."""
    return bool(text) and _DIGITS_RE.fullmatch(text) is not None


def _has_valid_grouping(text: str, sep: str) -> bool:
    """Sprawdza poprawne grupowanie tysięcy, np. ``"1 234 567"``.

    Pierwsza grupa ma 1-3 cyfry, każda kolejna dokładnie 3.
    Gdy separator nie występuje, wystarczy, że tekst to same cyfry.
    """
    parts = text.split(sep)
    if len(parts) == 1:
        return _all_digits(text)
    head, tail = parts[0], parts[1:]
    if not _all_digits(head) or len(head) > 3:
        return False
    return all(len(part) == 3 and _all_digits(part) for part in tail)


def _looks_like_ipv4(text: str) -> bool:
    """True dla ciągów wyglądających jak adres IPv4 (nie konwertujemy ich)."""
    match = _IPV4_RE.fullmatch(text)
    if match is None:
        return False
    return all(int(group) <= 255 for group in match.groups())


def _parse_number(text: str):
    """Próbuje zinterpretować tekst jako ``int``/``float``; ``None`` = to nie liczba."""
    body = text.replace("−", "-")  # matematyczny minus
    for space in _GROUP_SPACES:
        body = body.replace(space, " ")

    if not body:
        return None

    negative = False
    if body[0] in "+-":
        if body[0] == "+":
            # "+48 123 456 789" i podobne — to niemal zawsze numer telefonu.
            return None
        negative = True
        body = body[1:]

    if not body or _NUM_BODY_RE.fullmatch(body) is None:
        return None

    last_dot = body.rfind(".")
    last_comma = body.rfind(",")

    if last_dot >= 0 and last_comma >= 0:
        # Separatorem dziesiętnym jest ten znak, który stoi bliżej końca.
        decimal_sep = "." if last_dot > last_comma else ","
        if body.count(decimal_sep) != 1:
            return None
        int_part, frac_part = body.split(decimal_sep)
    elif last_comma >= 0:
        if body.count(",") == 1:
            int_part, frac_part = body.split(",")
        else:
            int_part, frac_part = body, None
    elif last_dot >= 0:
        if body.count(".") == 1:
            int_part, frac_part = body.split(".")
        else:
            int_part, frac_part = body, None
    else:
        int_part, frac_part = body, None

    # W części całkowitej dopuszczamy tylko JEDEN rodzaj separatora grup.
    separators = {char for char in int_part if char in ". ,"}
    if len(separators) > 1:
        return None
    group_sep = separators.pop() if separators else None

    if group_sep is not None:
        if not _has_valid_grouping(int_part, group_sep):
            return None
        digits = int_part.replace(group_sep, "")
    else:
        digits = int_part

    if not _all_digits(digits):
        return None
    if len(digits) > 1 and digits[0] == "0":
        return None  # numer katalogowy / kod, np. "007"

    if frac_part is None:
        if len(digits) > MAX_INT_DIGITS:
            return None
        value = int(digits)
        return -value if negative else value

    if not _all_digits(frac_part):
        return None
    try:
        value = float(digits + "." + frac_part)
    except (ValueError, OverflowError):
        return None
    if not math.isfinite(value):
        return None
    return -value if negative else value


# --------------------------------------------------------------------------- #
# Pomocnicze — daty
# --------------------------------------------------------------------------- #
def _timezone_from(text):
    """Buduje ``timezone`` z sufiksu ISO (``Z``, ``+02:00``, ``-0500``, ``+02``)."""
    if not text:
        return None
    if text in ("Z", "z"):
        return _dt.timezone.utc
    sign = 1 if text[0] == "+" else -1
    rest = text[1:].replace(":", "")
    hours = int(rest[0:2])
    minutes = int(rest[2:4]) if len(rest) >= 4 else 0
    if minutes > 59:
        raise ValueError("nieprawidłowe minuty offsetu")
    return _dt.timezone(sign * _dt.timedelta(hours=hours, minutes=minutes))


def _parse_date(text: str):
    """Próbuje zinterpretować tekst jako datę/datę z czasem; ``None`` gdy się nie da."""
    match = _ISO_DT_RE.fullmatch(text)
    if match is not None:
        year, month, day, hour, minute = (int(g) for g in match.group(1, 2, 3, 4, 5))
        second = int(match.group(6) or 0)
        micro = int((match.group(7) or "").ljust(6, "0")) if match.group(7) else 0
        try:
            tzinfo = _timezone_from(match.group(8))
            return _dt.datetime(
                year, month, day, hour, minute, second, micro, tzinfo=tzinfo
            )
        except ValueError:
            return None

    match = _YMD_RE.fullmatch(text)
    if match is not None:
        try:
            return _dt.date(int(match.group(1)), int(match.group(2)), int(match.group(3)))
        except ValueError:
            return None

    match = _DMY_RE.fullmatch(text)
    if match is not None:
        try:
            return _dt.date(int(match.group(4)), int(match.group(3)), int(match.group(1)))
        except ValueError:
            return None

    return None


def _clean_multiline(text: str) -> str:
    """Porządkuje tekst wielolinijkowy: ujednolica końce linii, przycina wcięcia.

    Wcięcia pochodzące z formatowania XML znikają, puste linie na początku i końcu
    są usuwane, a ciągi pustych linii zwijane do jednej.  Podział na linie
    pozostaje zachowany (Excel pokaże go jako łamanie wiersza w komórce).
    """
    normalized = text.replace("\r\n", "\n").replace("\r", "\n")
    lines = [line.strip() for line in normalized.split("\n")]
    result = []
    for line in lines:
        if not line and result and not result[-1]:
            continue  # kolejna pusta linia z rzędu — pomijamy
        result.append(line)
    while result and not result[0]:
        result.pop(0)
    while result and not result[-1]:
        result.pop()
    return "\n".join(result)


# --------------------------------------------------------------------------- #
# Konwersja typów
# --------------------------------------------------------------------------- #
def coerce_value(text: "str | None") -> Any:
    """Zamienia tekst z XML na najbardziej sensowny typ Pythona.

    Kolejność rozpoznawania: pusto -> tekst wielolinijkowy -> bool -> data ->
    liczba -> tekst.  Gdy nic nie pasuje, zwracany jest tekst z przyciętymi
    białymi znakami z brzegów (treść wewnętrzna pozostaje nietknięta).

    Szczegóły reguł i przypadki niejednoznaczne opisuje docstring modułu.
    """
    if text is None:
        return None
    if not isinstance(text, str):
        return text  # defensywnie: gotowe typy przepuszczamy bez zmian

    stripped = text.strip()  # str.strip() usuwa też NBSP
    if not stripped:
        return None

    if "\n" in stripped or "\r" in stripped:
        return _clean_multiline(stripped)

    if len(stripped) > MAX_SCALAR_LEN:
        return stripped

    lowered = stripped.lower()
    if lowered == "true":
        return True
    if lowered == "false":
        return False

    parsed_date = _parse_date(stripped)
    if parsed_date is not None:
        return parsed_date

    if _looks_like_ipv4(stripped):
        return stripped

    number = _parse_number(stripped)
    if number is not None:
        return number

    return stripped


# --------------------------------------------------------------------------- #
# Bezpieczeństwo komórki
# --------------------------------------------------------------------------- #
def _sanitize_text(text: str) -> str:
    """Usuwa znaki niedozwolone w XLSX i przycina do limitu Excela."""
    cleaned = _ILLEGAL_XLSX_RE.sub("", text)
    if len(cleaned) > MAX_CELL_CHARS:
        cleaned = cleaned[:MAX_CELL_CHARS]
    return cleaned


def sanitize_cell(value: Any) -> Any:
    """Przygotowuje wartość do zapisu w arkuszu.

    * ``None`` zostaje ``None`` (pusta komórka),
    * ``bool``/``int``/``float`` przechodzą bez zmian (NaN/Inf -> tekst),
    * data z offsetem strefy jest sprowadzana do UTC i pozbawiana ``tzinfo``
      (XLSX nie zna stref czasowych),
    * tekst traci znaki sterujące (poza ``\\t``, ``\\n``, ``\\r``), samotne surogaty
      i nie-znaki Unicode, a następnie jest przycinany do 32767 znaków,
    * pozostałe typy są zamieniane na tekst.

    Funkcja **nie** dodaje apostrofu — ochroną przed wstrzyknięciem formuły
    zajmuje się moduł zapisujący (patrz :func:`looks_like_formula`).
    """
    if value is None:
        return None

    if isinstance(value, bool):
        return value

    if isinstance(value, int):
        if abs(value) > _EXCEL_MAX_NUMBER:
            return _sanitize_text(str(value))
        return value

    if isinstance(value, float):
        if math.isnan(value) or math.isinf(value):
            return str(value)
        return value

    if isinstance(value, Decimal):
        try:
            as_float = float(value)
        except (ValueError, OverflowError, InvalidOperation):
            return _sanitize_text(str(value))
        if math.isnan(as_float) or math.isinf(as_float):
            return str(as_float)
        return as_float

    if isinstance(value, _dt.datetime):
        if value.tzinfo is not None:
            value = value.astimezone(_dt.timezone.utc).replace(tzinfo=None)
        return value

    if isinstance(value, _dt.date):  # sam date (datetime obsłużony wyżej)
        return value

    if isinstance(value, _dt.time):
        return value.replace(tzinfo=None) if value.tzinfo is not None else value

    if isinstance(value, _dt.timedelta):
        return value

    if isinstance(value, (bytes, bytearray)):
        return _sanitize_text(bytes(value).decode("utf-8", "replace"))

    if isinstance(value, str):
        return _sanitize_text(value)

    return _sanitize_text(str(value))


def looks_like_formula(value: Any) -> bool:
    """True, gdy tekst mógłby zostać uznany przez Excela za formułę.

    Dotyczy ciągów zaczynających się (po odcięciu białych znaków) od
    ``=``, ``+``, ``-``, ``@`` oraz ciągów zaczynających się od tabulatora
    lub powrotu karetki.  Wartości nietekstowe (liczby, daty) nie są formułami.
    """
    if not isinstance(value, str) or not value:
        return False
    if value[0] in "\t\r":
        return True
    stripped = value.lstrip()
    return bool(stripped) and stripped[0] in "=+-@"
