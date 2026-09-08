# -*- coding: utf-8 -*-
"""Zapis danych do pliku XLSX — dwie niezależne ścieżki zapisu.

Moduł udostępnia JEDNO API (:func:`write_workbook`) i dwie implementacje pod spodem:

``openpyxl``
    używana, gdy biblioteka jest zainstalowana; tryb ``write_only`` (strumieniowy),
``stdlib``
    WBUDOWANY generator OOXML oparty wyłącznie na :mod:`zipfile` i sklejaniu XML-a;
    działa bez żadnych zależności zewnętrznych.

Wybór ścieżki steruje zmienna środowiskowa :data:`ENV_BACKEND`
(``FLEXIT_XLSX_BACKEND``) o wartościach ``stdlib`` | ``openpyxl`` | ``auto``
(domyślnie ``auto``).  Obie ścieżki produkują plik o tej samej semantyce:
te same typy komórek, ten sam nagłówek, ta sama ochrona przed formułami.

Podjęte decyzje projektowe (przypadki niejednoznaczne)
-----------------------------------------------------
* **Ochrona przed wstrzyknięciem formuły.**  Tekst NIGDY nie trafia do pliku jako
  element ``<f>``.  Ścieżka stdlib zapisuje każdy napis jako ``t="inlineStr"``,
  więc ``"=A1+1"`` jest tekstem z definicji formatu.  openpyxl samo zamieniłoby
  taki napis na formułę, więc wymuszamy ``cell.data_type = "s"`` wszędzie tam,
  gdzie mógłby to zrobić: dla napisów uznanych przez
  :func:`~flexit2xlsx.values.looks_like_formula` (``= + - @`` oraz wiodący
  tabulator/CR) i dla kodów błędów Excela (patrz niżej).  Pozostałe napisy
  openpyxl i tak zapisuje jako tekst, a owijanie każdej komórki w obiekt
  ``WriteOnlyCell`` kosztowałoby ~40% czasu zapisu.  NIE dodajemy apostrofu —
  wartość odczytana z pliku jest identyczna z tą, którą podał użytkownik.
* **Napisy zapisujemy „inline”, nie przez ``sharedStrings.xml``.**  Tablica
  napisów wspólnych wymagałaby trzymania w pamięci wszystkich unikalnych tekstów
  (przy 200 tys. wierszy to setki MB), a zysk zjada kompresja ZIP.  Kontrakt
  dopuszcza oba warianty.
* **Strumieniowość vs. szerokości kolumn.**  Element ``<cols>`` musi znaleźć się
  w XML-u PRZED danymi, a wierszy nie chcemy trzymać w pamięci.  Dlatego
  szerokości liczymy z nagłówka i pierwszych :data:`WIDTH_SAMPLE_ROWS` wierszy
  (jedyna część danych trzymana w RAM).  To samo ograniczenie ma tryb
  ``write_only`` openpyxl-a, więc obie ścieżki działają identycznie.
* **Element ``<dimension>`` pomijamy** — jego wyliczenie wymagałoby znajomości
  liczby wierszy przed ich zapisaniem.  Jest opcjonalny; Excel i LibreOffice
  wyznaczają zakres same (tak samo postępuje tryb ``write_only`` openpyxl-a).
  Skutek uboczny: ``openpyxl.load_workbook(..., read_only=True)`` poda dla
  takiego arkusza ``max_row is None``, dopóki nie zawoła się
  ``worksheet.reset_dimensions()``.  ``<autoFilter>`` stoi w XML-u PO
  ``<sheetData>``, więc jego zakres znamy już dokładnie.
* **Limity arkusza wstrzykiwane są przez stałe modułu** :data:`MAX_ROWS_PER_SHEET`
  i :data:`MAX_COLS_PER_SHEET`, a nie przez parametry — sygnatura
  :func:`write_workbook` jest częścią wiążącego kontraktu i nie wolno jej zmieniać.
  Stałe są odczytywane przy KAŻDYM wywołaniu, więc test może je podmienić
  (``unittest.mock.patch.object``) i sprawdzić dzielenie arkusza na kilkunastu
  wierszach zamiast na milionie.
* **Dzielenie po wierszach jest strumieniowe**, po kolumnach — nie.  Nadmiarowe
  bloki kolumn wymagają drugiego przejścia po tych samych wierszach, a źródło
  bywa jednorazowym generatorem; w tym (skrajnie rzadkim, >16 384 kolumn)
  przypadku wiersze lądują w pliku tymczasowym (``pickle``), nie w pamięci.
* **Arkusz-kontynuacja powtarza kolumny kluczowe.**  Po podziale na bloki kolumn
  drugi (trzeci, …) arkusz zaczynałby się od przypadkowej kolumny danych i nie
  dałoby się przypisać wiersza do aukcji inaczej niż po numerze wiersza.  Dlatego
  na początku KAŻDEGO arkusza-kontynuacji powtarzamy :data:`KEY_COLUMNS_ON_SPLIT`
  pierwszych kolumn arkusza (CLI trzyma tam ``Aukcja``, ``Plik``, ``Nr pozycji``).
  Liczbę można zmienić polem ``Sheet.key_columns`` (``0`` wyłącza powtarzanie).
* **Nazwy kolejnych części** powstają przez zwykłe :func:`safe_sheet_name`,
  które dokleja ``" (2)"``, ``" (3)"``…  Numeracja jest wspólna dla podziału po
  wierszach i po kolumnach — prostsza i zawsze unikalna.  Co dokładnie się stało,
  mówi ostrzeżenie na stderr.
* **``bool`` zapisujemy jako natywny typ logiczny** (``t="b"``), a nie jako tekst
  „PRAWDA”/„FAŁSZ”.  Napis „PRAWDA” w pliku byłby poprawny tylko dla polskiego
  Excela; typ natywny wyświetla się poprawnie w KAŻDEJ wersji językowej
  (i tak samo w LibreOffice).
* **Daty sprzed 1900-01-01** zapisujemy jako tekst ISO — Excel nie ma dla nich
  numeru seryjnego.  Dotyczy to także 1899-12-31: numer seryjny wyszedłby ``0``,
  a Excel pokazałby wtedy godzinę ``00:00`` zamiast daty.  Dla pozostałych dat
  stosujemy system 1900 z celowym błędem przestępności roku 1900 (zgodność
  z Excelem).
* **``None`` to komórka pusta** — nie zapisujemy dla niej elementu ``<c>``.
* Każda wartość **oraz każdy nagłówek kolumny** przechodzi przez
  :func:`~flexit2xlsx.values.sanitize_cell` (znaki sterujące, limit 32767 znaków,
  NaN/Inf, strefy czasowe).  Nagłówek to zwykła komórka arkusza — nazwa kolumny
  pochodzi z danych XML, więc bez odkażania potrafiła uszkodzić plik.
  Gdy cokolwiek zostało skrócone do limitu Excela, mówimy o tym na stderr.
* **Każdy napis jest w pliku typem tekstowym** (``t="s"`` / ``inlineStr``), także
  taki, który wygląda na kod błędu Excela (``#N/A``, ``#REF!``, ``#DIV/0!`` …).
  Bez tego openpyxl zapisałby go jako komórkę BŁĘDU (``t="e"``), która zatruwa
  wyniki formuł (``SUMA`` po kolumnie zwraca ``#N/A``) — i obie ścieżki zapisu
  dawałyby z tego samego XML-a różne dane.
* **Nazwa arkusza jest odkażana tak samo jak treść komórki** — znaki zabronione
  w XML 1.0 (sterujące, samotne surogaty z ``surrogateescape`` w nazwach plików,
  ``U+FFFE``/``U+FFFF``) są usuwane, zanim trafią do ``xl/workbook.xml``.
* **Uprawnienia pliku wynikowego** biorą się z ``umask`` użytkownika (typowo
  ``0644``), a przy nadpisywaniu — z uprawnień istniejącego pliku.  Dokładnie tak
  zachowuje się zwykłe ``open(path, "wb")``; plik tymczasowy narzucał ``0600``,
  przez co wynik bywał nieczytelny dla innych kont (Samba, serwer WWW).
* **Zapis jest atomowy**: plik powstaje obok celu i dopiero ``os.replace``
  podmienia go na miejscu, więc przerwany zapis nie zostawia uszkodzonego XLSX-a.
* **Wynik jest deterministyczny**: we wbudowanym zapisie co do bajtu (znaczniki
  czasu w archiwum ZIP są ustalone na stałe, 1980-01-01, a metadane nie zawierają
  dat).  W ścieżce openpyxl identyczne są wszystkie części z DANYMI; sama
  biblioteka stempluje bieżącą godziną wpisy ZIP oraz pole ``dcterms:modified``
  w ``docProps/core.xml`` (robi to bezwarunkowo w ``save_workbook`` — nie da się
  tego wyłączyć), więc tam powtarzalność kończy się na metadanych.
"""

from __future__ import annotations

import datetime as _dt
import itertools
import os
import pickle
import re
import stat
import sys
import tempfile
import zipfile
from dataclasses import dataclass, field
from typing import Any, Iterable, Iterator, Sequence

from .values import MAX_CELL_CHARS, looks_like_formula, sanitize_cell

#: Napisy, które openpyxl sam zamieniłby na komórkę BŁĘDU (``t="e"``).  Lista jest
#: zaszyta na wypadek starszej/nowszej wersji biblioteki — my zapisujemy je jako
#: zwykły tekst, tak samo jak robi to wbudowany zapis.
_ERROR_CODES = frozenset(
    ("#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A")
)

try:  # jedyny dozwolony import spoza biblioteki standardowej — w pełni opcjonalny
    import openpyxl as _openpyxl
    from openpyxl.cell import WriteOnlyCell as _WriteOnlyCell
    from openpyxl.styles import Font as _Font
except ImportError:  # pragma: no cover - zależy od środowiska
    _openpyxl = None
    _WriteOnlyCell = None
    _Font = None

try:  # lista kodów błędów tej konkretnej wersji openpyxl-a (osobno, bo jej brak
    # nie może wyłączyć całej ścieżki zapisu — mamy własną kopię wyżej)
    from openpyxl.cell.cell import ERROR_CODES as _OPENPYXL_ERROR_CODES

    _ERROR_CODES = _ERROR_CODES | frozenset(_OPENPYXL_ERROR_CODES)
except ImportError:  # pragma: no cover - zależy od środowiska
    pass

__all__ = [
    "Sheet",
    "backend_name",
    "safe_sheet_name",
    "write_workbook",
    "ENV_BACKEND",
    "MAX_ROWS_PER_SHEET",
    "MAX_COLS_PER_SHEET",
    "SHEET_NAME_MAX",
    "WIDTH_SAMPLE_ROWS",
    "KEY_COLUMNS_ON_SPLIT",
]

#: Zmienna środowiskowa wymuszająca ścieżkę zapisu: ``stdlib`` | ``openpyxl`` | ``auto``.
ENV_BACKEND = "FLEXIT_XLSX_BACKEND"

#: Twardy limit Excela: liczba wierszy w arkuszu (razem z nagłówkiem).
MAX_ROWS_PER_SHEET = 1_048_576

#: Twardy limit Excela: liczba kolumn w arkuszu.
MAX_COLS_PER_SHEET = 16_384

#: Maksymalna długość nazwy arkusza (limit Excela).
SHEET_NAME_MAX = 31

#: Ile pierwszych wierszy części arkusza trafia do pamięci, by wyliczyć szerokości kolumn.
WIDTH_SAMPLE_ROWS = 200

#: Ile pierwszych kolumn arkusza powtarzamy na początku arkusza-kontynuacji, gdy
#: kolumny nie mieszczą się w jednym arkuszu (CLI trzyma tam kolumny techniczne:
#: ``Aukcja``, ``Plik``, ``Nr pozycji``).  Nadpisuje to pole ``Sheet.key_columns``.
KEY_COLUMNS_ON_SPLIT = 3

#: Powyżej tylu zapisanych wierszy podpowiadamy, że wbudowany zapis jest szybszy
#: od openpyxl-a (podpowiedź pokazujemy tylko w trybie automatycznym).
SLOW_BACKEND_HINT_ROWS = 100_000

#: Zakres dopuszczalnych szerokości kolumn (w „znakach” Excela).
MIN_COL_WIDTH = 8.0
MAX_COL_WIDTH = 60.0

#: Nazwa arkusza używana, gdy po oczyszczeniu nic nie zostanie.
DEFAULT_SHEET_NAME = "Arkusz"

#: Co ile fragmentów XML-a opróżniamy bufor do strumienia ZIP.
_FLUSH_EVERY = 512

#: Znaki zakazane w nazwie arkusza przez Excela.
_FORBIDDEN_SHEET_CHARS = "[]:*?/\\"

#: Znaki, których NIE WOLNO umieścić w dokumencie XML 1.0 — ten sam zbiór, który
#: :func:`~flexit2xlsx.values.sanitize_cell` usuwa z treści komórek (plus ``\x7f``).
#: Samotne surogaty biorą się z ``surrogateescape`` przy nazwach plików spoza UTF-8,
#: a ``U+FFFE``/``U+FFFF`` są w nazwie pliku legalne — w XML-u już nie.
_ILLEGAL_XML_RE = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\x7f\ud800-\udfff\ufffe\uffff]")

#: Nazwy zarezerwowane przez Excela (porównanie bez uwzględniania wielkości liter).
_RESERVED_SHEET_NAMES = {"history"}

#: Epoka Excela w systemie 1900 (z uwzględnieniem błędu roku przestępnego).
_EXCEL_EPOCH = _dt.datetime(1899, 12, 30)

#: Stały znacznik czasu wpisów ZIP — gwarantuje powtarzalny (deterministyczny) wynik.
_ZIP_DATE = (1980, 1, 1, 0, 0, 0)

#: Ta sama data wstawiana w ``docProps/core.xml`` przez ścieżkę openpyxl
#: (domyślnie openpyxl wpisałby tam bieżący czas i wynik nie byłby powtarzalny).
_DOC_DATE = _dt.datetime(*_ZIP_DATE)

_NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
_NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
_XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'

# Indeksy stylów w ``xl/styles.xml`` (patrz _STYLES_XML).
_S_DEFAULT = 0
_S_HEADER = 1
_S_DATE = 2
_S_DATETIME = 3
_S_TIME = 4
_S_DURATION = 5


# --------------------------------------------------------------------------- #
# Typy publiczne
# --------------------------------------------------------------------------- #


@dataclass
class Sheet:
    """Opis jednego arkusza do zapisania.

    :param name: nazwa arkusza (zostanie oczyszczona przez :func:`safe_sheet_name`),
    :param columns: nagłówki kolumn w kolejności zapisu,
    :param rows: źródło wierszy — lista, krotka albo **generator**; każdy wiersz to
        sekwencja wartości.  Wiersz krótszy niż ``columns`` zostanie uzupełniony
        pustymi komórkami, dłuższy — przycięty.
    :param key_columns: ile pierwszych kolumn powtórzyć na początku każdego
        arkusza-kontynuacji, gdy kolumny nie mieszczą się w jednym arkuszu
        (``None`` = :data:`KEY_COLUMNS_ON_SPLIT`, ``0`` = nie powtarzaj).
        Pole jest opcjonalne — kontraktowe ``Sheet(name, columns, rows)`` działa
        bez zmian.
    """

    name: str
    columns: list[str]
    rows: Iterable[Sequence[Any]] = field(default_factory=list)
    key_columns: int | None = None


@dataclass(frozen=True)
class _Options:
    """Opcje formatowania wspólne dla obu ścieżek zapisu."""

    freeze_header: bool = True
    autofilter: bool = True
    auto_width: bool = True


# --------------------------------------------------------------------------- #
# Wybór ścieżki zapisu
# --------------------------------------------------------------------------- #


def _plural(count: int, one: str, few: str, many: str) -> str:
    """Wstawia liczbę do wzorca w poprawnej polskiej formie gramatycznej.

    Wzorce zawierają ``%d``, np. ``("%d arkusz", "%d arkusze", "%d arkuszy")``.
    """
    if count == 1:
        pattern = one
    elif 2 <= count % 10 <= 4 and not 12 <= count % 100 <= 14:
        pattern = few
    else:
        pattern = many
    return pattern % count


def _warn(message: str) -> None:
    """Wypisuje ostrzeżenie na stderr (nigdy nie przerywa zapisu)."""
    print("[flexit2xlsx] " + message, file=sys.stderr)


def backend_name() -> str:
    """Zwraca nazwę ścieżki zapisu, której użyje :func:`write_workbook`.

    Wartość: ``"openpyxl"`` albo ``"stdlib"``.  Zmienna środowiskowa
    ``FLEXIT_XLSX_BACKEND`` pozwala wymusić wybór; jest odczytywana przy każdym
    wywołaniu, więc można ją zmieniać w trakcie działania programu (i w testach).

    Żądanie ``openpyxl`` przy braku biblioteki NIE jest błędem — wypisujemy
    ostrzeżenie i schodzimy na wbudowany zapis, bo lepiej oddać użytkownikowi
    plik niż wyjątek.
    """
    choice = (os.environ.get(ENV_BACKEND) or "auto").strip().lower()
    available = "openpyxl" if _openpyxl is not None else "stdlib"

    if choice in ("", "auto"):
        return available
    if choice == "stdlib":
        return "stdlib"
    if choice == "openpyxl":
        if _openpyxl is None:
            _warn(
                "zażądano %s=openpyxl, ale biblioteka nie jest zainstalowana; "
                "używam wbudowanego zapisu (stdlib)" % ENV_BACKEND
            )
            return "stdlib"
        return "openpyxl"

    _warn(
        "nieznana wartość %s=%r (dozwolone: stdlib, openpyxl, auto); "
        "używam trybu automatycznego" % (ENV_BACKEND, choice)
    )
    return available


# --------------------------------------------------------------------------- #
# Nazwy arkuszy
# --------------------------------------------------------------------------- #


def _trim_sheet_name(text: str) -> str:
    """Obcina z brzegów białe znaki i apostrofy (Excel nie przyjmuje takich nazw).

    Pętla jest potrzebna, bo po odcięciu apostrofu może wyjść spacja i odwrotnie
    (``"' a '"``).  Wołamy to także PO skróceniu do 31 znaków — inaczej apostrof
    ukryty w środku nazwy wylądowałby na jej końcu.
    """
    previous = None
    while previous != text:
        previous = text
        text = text.strip().strip("'")
    return text


def safe_sheet_name(name: str, used: set[str]) -> str:
    """Zamienia dowolny napis na poprawną, unikalną nazwę arkusza Excela.

    Reguły Excela: maksymalnie 31 znaków, bez ``[ ] : * ? / \\``, bez znaków
    sterujących, nazwa nie może być pusta ani zaczynać/kończyć się apostrofem.
    Unikalność Excel sprawdza **bez uwzględniania wielkości liter**, więc
    ``"Dane"`` i ``"dane"`` to dla niego kolizja.

    Nazwa trafia do atrybutu ``name`` w ``xl/workbook.xml``, więc usuwamy z niej
    także wszystko, czego nie wolno umieścić w dokumencie XML 1.0: znaki sterujące,
    samotne surogaty (nazwa pliku spoza UTF-8 dekodowana przez ``surrogateescape``)
    oraz nie-znaki ``U+FFFE``/``U+FFFF``.  Bez tego powstawał plik `.xlsx`,
    którego Excel ani openpyxl nie potrafiły otworzyć — albo goły
    ``UnicodeEncodeError`` z wnętrza ``zipfile``.

    Znaki zakazane zamieniamy na ``_`` (zachowuje czytelność: ``"a/b"`` -> ``"a_b"``),
    białe znaki zwijamy do pojedynczych spacji.  Duplikaty rozróżniamy przyrostkiem
    ``" (2)"``, ``" (3)"`` …, skracając podstawę tak, by zmieścić się w 31 znakach.

    **Efekt uboczny (świadomy):** zwrócona nazwa jest dopisywana do ``used``.
    Dzięki temu wołanie w pętli po prostu działa, a wywołujący nie musi pamiętać
    o ręcznej aktualizacji zbioru.
    """
    if used is None:  # defensywnie — kontrakt mówi o zbiorze, ale None nie ma boleć
        used = set()

    raw = "" if name is None else str(name)
    # Najpierw znaki nielegalne w XML-u (surogaty, U+FFFE/U+FFFF, sterujące) —
    # dokładnie ten sam zbiór, który values.sanitize_cell usuwa z treści komórek.
    raw = _ILLEGAL_XML_RE.sub(" ", raw)
    chars = []
    for ch in raw:
        if ch in _FORBIDDEN_SHEET_CHARS:
            chars.append("_")
        elif ch < " " or ch == "\x7f":  # znaki sterujące -> spacja (potem zwinięta)
            chars.append(" ")
        else:
            chars.append(ch)
    text = _trim_sheet_name(" ".join("".join(chars).split()))

    if not text:
        text = DEFAULT_SHEET_NAME
    if text.lower() in _RESERVED_SHEET_NAMES:
        text = text + "_"
    if len(text) > SHEET_NAME_MAX:
        # Skracamy, a DOPIERO POTEM czyścimy brzegi — po cięciu na 31 znaku
        # nazwa mogłaby kończyć się apostrofem albo spacją.
        text = _trim_sheet_name(text[:SHEET_NAME_MAX]) or DEFAULT_SHEET_NAME

    lowered = {str(u).lower() for u in used}
    if text.lower() not in lowered:
        used.add(text)
        return text

    counter = 2
    while True:
        suffix = " (%d)" % counter
        base = _trim_sheet_name(text[: SHEET_NAME_MAX - len(suffix)])
        if not base:
            base = DEFAULT_SHEET_NAME[: SHEET_NAME_MAX - len(suffix)]
        candidate = base + suffix
        if candidate.lower() not in lowered:
            used.add(candidate)
            return candidate
        counter += 1


def _xml_safe_name(name: str) -> str:
    """Ostatnia linia obrony przed uszkodzonym ``xl/workbook.xml``.

    Nazwy arkuszy przechodzą przez :func:`safe_sheet_name`, więc normalnie nie ma
    tu nic do roboty.  Gdyby jednak do zapisu trafiła nazwa z innego źródła,
    wolimy podmienić znak niż zapisać plik, którego nie da się otworzyć (albo
    wywalić się ``UnicodeEncodeError`` w środku ``zipfile``).
    """
    cleaned = " ".join(_ILLEGAL_XML_RE.sub(" ", name).split()) or DEFAULT_SHEET_NAME
    try:
        cleaned.encode("utf-8")
    except UnicodeEncodeError:  # pragma: no cover - po podstawieniu wyżej nie zachodzi
        cleaned = cleaned.encode("utf-8", "replace").decode("utf-8")
    return cleaned


# --------------------------------------------------------------------------- #
# Pomocnicze: kolumny, daty, escaping
# --------------------------------------------------------------------------- #


def _col_letters(count: int) -> list[str]:
    """Zwraca listę oznaczeń kolumn: ``["A", "B", ..., "Z", "AA", ...]``."""
    letters: list[str] = []
    for index in range(1, count + 1):
        n = index
        buf = []
        while n:
            n, rest = divmod(n - 1, 26)
            buf.append(chr(65 + rest))
        letters.append("".join(reversed(buf)))
    return letters


def _excel_serial(value: Any) -> float | None:
    """Zamienia datę/czas na numer seryjny Excela (system 1900).

    Zwraca ``None``, gdy wartości nie da się w Excelu przedstawić (data przed
    1900-01-01) — wtedy zapisujemy ją jako tekst ISO.

    Granicą jest ``days < 2``, a nie ``days < 1``: dla 1899-12-31 wychodzi
    ``days == 1``, po korekcie roku przestępnego numer seryjny ``0``, a Excel
    wyświetla ``0`` jako godzinę ``00:00``, nie jako datę (openpyxl odczytuje
    wtedy ``datetime.time(0, 0)``).  Data po prostu by zniknęła.
    """
    if isinstance(value, _dt.datetime):
        moment = value
    elif isinstance(value, _dt.date):
        moment = _dt.datetime(value.year, value.month, value.day)
    elif isinstance(value, _dt.time):
        return (
            value.hour * 3600 + value.minute * 60 + value.second
        ) / 86400.0 + value.microsecond / 86400000000.0
    else:  # pragma: no cover - dobór typu po stronie wywołującego
        return None

    days = (moment - _EXCEL_EPOCH).days
    if days < 2:
        return None  # przed 1900-01-01 Excel nie ma numeru seryjnego
    if days < 61:
        # Excel uważa 1900 za rok przestępny; do 1900-02-28 numery są o 1 mniejsze.
        days -= 1
    fraction = (
        moment.hour * 3600 + moment.minute * 60 + moment.second
    ) / 86400.0 + moment.microsecond / 86400000000.0
    return days + fraction


def _esc(text: str) -> str:
    """Escapuje tekst na potrzeby zawartości elementu XML.

    ``\\r`` zamieniamy na encję, bo parser XML normalizuje surowy CR do LF —
    bez tego znak nie przetrwałby zapisu i odczytu.
    """
    if "&" in text:
        text = text.replace("&", "&amp;")
    if "<" in text:
        text = text.replace("<", "&lt;")
    if ">" in text:
        text = text.replace(">", "&gt;")
    if "\r" in text:
        text = text.replace("\r", "&#13;")
    return text


def _esc_attr(text: str) -> str:
    """Escapuje tekst na potrzeby wartości atrybutu XML."""
    return _esc(text).replace('"', "&quot;").replace("\t", "&#9;").replace("\n", "&#10;")


class _Stats:
    """Liczniki zbierane w trakcie zapisu — na ich podstawie ostrzegamy użytkownika."""

    __slots__ = ("truncated", "rows")

    def __init__(self) -> None:
        self.truncated = 0  # komórki (i nagłówki) skrócone do limitu Excela
        self.rows = 0  # zapisane wiersze danych


def _header_text(column: Any, stats: _Stats | None = None) -> str:
    """Zwraca nagłówek kolumny odkażony DOKŁADNIE tak jak wartość komórki.

    Nazwa kolumny pochodzi z danych (np. z atrybutu ``<spec name="...">``), więc
    równie dobrze jak wartość może zawierać znak sterujący, samotny surogat albo
    40 000 znaków.  Bez :func:`~flexit2xlsx.values.sanitize_cell` ścieżka stdlib
    produkowała niepoprawny XML, a openpyxl rzucał ``IllegalCharacterError``.
    """
    text = "" if column is None else str(column)
    if stats is not None and len(text) > MAX_CELL_CHARS:
        stats.truncated += 1
    cleaned = sanitize_cell(text)
    return cleaned if isinstance(cleaned, str) else str(cleaned)


def _display_len(value: Any) -> int:
    """Szacuje szerokość wyświetlania wartości (w znakach) — do doboru szerokości kolumn."""
    if value is None:
        return 0
    kind = type(value)
    if kind is str:
        if "\n" in value:
            return max(len(part) for part in value.split("\n"))
        return len(value)
    if kind is bool:
        return 6  # PRAWDA / FAŁSZ
    if kind is _dt.datetime:
        return 19
    if kind is _dt.date:
        return 10
    if kind is _dt.time:
        return 8
    return len(str(value))


def _part_values(row: Sequence[Any], part: "_Part") -> Sequence[Any]:
    """Wycina z wiersza wartości należące do danej części arkusza.

    Dla arkusza-kontynuacji sklejamy powtórzone kolumny kluczowe (``row[:key]``)
    z właściwym blokiem kolumn (``row[start:stop]``).
    """
    key = part.key
    if key:
        return list(row[:key]) + list(row[part.start : part.stop])
    if part.start != 0 or part.stop is not None:
        return row[part.start : part.stop]
    return row


def _estimate_widths(part: "_Part") -> list[float]:
    """Wylicza szerokości kolumn z nagłówka i próbki wierszy."""
    columns = part.columns
    count = len(columns)
    best = [_display_len(_header_text(col)) for col in columns]
    for row in part.sample:
        values = _part_values(row, part)
        for index, value in enumerate(values):
            if index >= count:
                break
            length = _display_len(value)
            if length > best[index]:
                best[index] = length
    return [min(MAX_COL_WIDTH, max(MIN_COL_WIDTH, width + 2.0)) for width in best]


# --------------------------------------------------------------------------- #
# Podział na części (limity Excela)
# --------------------------------------------------------------------------- #


class _RowFeeder:
    """Iterator wierszy z podglądem jednego wiersza do przodu.

    Podgląd jest potrzebny, żeby wiedzieć, czy trzeba otworzyć kolejną część
    arkusza, ZANIM zacznie się ją zapisywać (a więc bez gubienia wiersza).
    """

    __slots__ = ("_iterator", "_pending", "_empty")

    def __init__(self, rows: Iterable[Sequence[Any]]):
        self._iterator = iter(rows)
        self._pending: Any = None
        self._empty = False
        self._advance()

    def _advance(self) -> None:
        try:
            self._pending = next(self._iterator)
        except StopIteration:
            self._pending = None
            self._empty = True

    @property
    def empty(self) -> bool:
        """True, gdy źródło nie ma już żadnego wiersza."""
        return self._empty

    def pop(self) -> Sequence[Any]:
        """Zwraca kolejny wiersz i przesuwa podgląd."""
        row = self._pending
        self._advance()
        return row


class _RowSpool:
    """Bufor wierszy w pliku tymczasowym — pozwala przejść po nich wielokrotnie.

    Używany WYŁĄCZNIE przy podziale na bloki kolumn (>16 384 kolumn), gdzie te
    same wiersze trzeba zapisać do kilku arkuszy, a źródłem bywa jednorazowy
    generator.  Dane idą na dysk, nie do pamięci.
    """

    def __init__(self, rows: Iterable[Sequence[Any]]):
        self._file = tempfile.TemporaryFile()
        self._count = 0
        dump = pickle.dump
        handle = self._file
        for row in rows:
            dump(list(row), handle, protocol=pickle.HIGHEST_PROTOCOL)
            self._count += 1

    def replay(self) -> Iterator[list[Any]]:
        """Zwraca świeży iterator po zbuforowanych wierszach."""
        self._file.seek(0)
        load = pickle.load
        handle = self._file
        for _ in range(self._count):
            yield load(handle)

    def close(self) -> None:
        """Zamyka i kasuje plik tymczasowy."""
        try:
            self._file.close()
        except OSError:  # pragma: no cover - zależne od systemu plików
            pass


@dataclass
class _Part:
    """Jedna fizyczna część arkusza (po podziale na limity Excela).

    ``start``/``stop`` wyznaczają blok kolumn danych, a ``key`` mówi, ile
    pierwszych kolumn arkusza doklejamy PRZED tym blokiem (kolumny kluczowe
    powtarzane w arkuszu-kontynuacji).  Dla pierwszej części ``key`` to zawsze 0,
    bo kolumny kluczowe i tak są na jej początku.
    """

    name: str
    columns: list[str]
    sample: list[Sequence[Any]]
    rows: Iterator[Sequence[Any]]
    start: int
    stop: int | None
    key: int = 0


def _row_limit() -> int:
    """Maksymalna liczba wierszy DANYCH w jednej części (nagłówek zajmuje jeden wiersz)."""
    return max(1, int(MAX_ROWS_PER_SHEET) - 1)


def _col_limit() -> int:
    """Maksymalna liczba kolumn w jednej części."""
    return max(1, int(MAX_COLS_PER_SHEET))


def _key_columns(sheet: Sheet, total: int, col_limit: int) -> int:
    """Ile pierwszych kolumn powtórzyć w arkuszu-kontynuacji.

    Wartość bierzemy z ``Sheet.key_columns`` (``None`` = :data:`KEY_COLUMNS_ON_SPLIT`)
    i przycinamy do ćwierci bloku kolumn — powtórka nie może zjeść miejsca na dane
    (przy sztucznie małym limicie z testów wychodzi wtedy po prostu 0).
    """
    requested = sheet.key_columns
    if requested is None:
        requested = KEY_COLUMNS_ON_SPLIT
    try:
        requested = int(requested)
    except (TypeError, ValueError):  # pragma: no cover - defensywnie
        requested = 0
    return max(0, min(requested, total, col_limit // 4))


def _drain(feeder: _RowFeeder, count: int) -> Iterator[Sequence[Any]]:
    """Wydaje co najwyżej ``count`` kolejnych wierszy ze źródła."""
    while count > 0 and not feeder.empty:
        yield feeder.pop()
        count -= 1


def _iter_parts(sheet: Sheet, used: set[str]) -> Iterator[_Part]:
    """Dzieli arkusz na części mieszczące się w limitach Excela.

    Generator jest **leniwy**: kolejna część powstaje dopiero wtedy, gdy poprzednia
    została w całości zapisana.  Konsument MUSI wyczerpać ``part.rows``, zanim
    poprosi o następną część — inaczej zgubi wiersze.
    """
    columns = list(sheet.columns or [])
    col_limit = _col_limit()
    key = _key_columns(sheet, len(columns), col_limit)
    if len(columns) <= col_limit:
        blocks = [(columns, 0, None, 0)]
    else:
        # Pierwszy blok to zwykły początek arkusza (kolumny kluczowe już w nim są),
        # każdy kolejny dostaje je z powrotem doklejone z przodu.
        blocks = [(columns[:col_limit], 0, col_limit, 0)]
        step = col_limit - key
        for index in range(col_limit, len(columns), step):
            stop = min(index + step, len(columns))
            blocks.append((columns[:key] + columns[index:stop], index, stop, key))
        _warn(
            "arkusz %r ma %d kolumn, a limit Excela to %d — kolumny trafią do %s%s"
            % (
                sheet.name,
                len(columns),
                col_limit,
                _plural(len(blocks), "%d arkusza", "%d arkuszy", "%d arkuszy"),
                (
                    ""
                    if not key
                    else "; %s powtórzono na początku każdego arkusza-kontynuacji"
                    % _plural(key, "%d pierwszą kolumnę", "%d pierwsze kolumny",
                              "%d pierwszych kolumn")
                ),
            )
        )

    spool = _RowSpool(sheet.rows) if len(blocks) > 1 else None
    produced: list[str] = []
    try:
        for block_columns, start, stop, block_key in blocks:
            source = sheet.rows if spool is None else spool.replay()
            feeder = _RowFeeder(source)
            first_part = True
            while first_part or not feeder.empty:
                first_part = False
                limit = _row_limit()
                sample: list[Sequence[Any]] = []
                cap = min(WIDTH_SAMPLE_ROWS, limit)
                while len(sample) < cap and not feeder.empty:
                    row = feeder.pop()
                    # Wiersz może być jednorazowym iteratorem — materializujemy go
                    # OD RAZU, bo próbkę czyta jeszcze _estimate_widths, a wyczerpany
                    # iterator zapisałby się potem jako pusty wiersz.
                    sample.append(row if isinstance(row, (list, tuple)) else list(row))
                name = safe_sheet_name(sheet.name, used)
                produced.append(name)
                yield _Part(
                    name=name,
                    columns=block_columns,
                    sample=sample,
                    rows=_drain(feeder, limit - len(sample)),
                    start=start,
                    stop=stop,
                    key=block_key,
                )
        if len(produced) > len(blocks):
            _warn(
                "arkusz %r przekracza limit %d wierszy na arkusz — nadmiar trafia "
                "do kolejnych arkuszy" % (sheet.name, int(MAX_ROWS_PER_SHEET))
            )
        if len(produced) > 1:
            _warn(
                "arkusz %r zapisany jako %s: %s"
                % (
                    sheet.name,
                    _plural(len(produced), "%d arkusz", "%d arkusze", "%d arkuszy"),
                    ", ".join(repr(item) for item in produced),
                )
            )
    finally:
        if spool is not None:
            spool.close()


# --------------------------------------------------------------------------- #
# Ścieżka „stdlib” — własny generator OOXML
# --------------------------------------------------------------------------- #


def _cell_xml_fallback(ref: str, value: Any) -> str:
    """Zapis komórki dla typów, które nie trafiły w szybką ścieżkę (podklasy itp.)."""
    if isinstance(value, bool):
        return '<c r="%s" t="b"><v>%d</v></c>' % (ref, 1 if value else 0)
    if isinstance(value, int):
        return '<c r="%s"><v>%d</v></c>' % (ref, value)
    if isinstance(value, float):
        return '<c r="%s"><v>%r</v></c>' % (ref, value)
    if isinstance(value, (_dt.datetime, _dt.date, _dt.time, _dt.timedelta)):
        return _date_cell_xml(ref, value)
    return _text_cell_xml(ref, str(value), _S_DEFAULT)


def _text_cell_xml(ref: str, text: str, style: int) -> str:
    """Komórka tekstowa — zawsze ``inlineStr``, więc nigdy nie jest formułą."""
    style_attr = "" if style == _S_DEFAULT else ' s="%d"' % style
    if text != text.strip():
        return '<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % (
            ref,
            style_attr,
            _esc(text),
        )
    return '<c r="%s"%s t="inlineStr"><is><t>%s</t></is></c>' % (ref, style_attr, _esc(text))


def _date_cell_xml(ref: str, value: Any) -> str:
    """Komórka z datą/czasem — liczba plus format liczbowy z ``xl/styles.xml``."""
    if isinstance(value, _dt.timedelta):
        return '<c r="%s" s="%d"><v>%r</v></c>' % (
            ref,
            _S_DURATION,
            value.total_seconds() / 86400.0,
        )
    serial = _excel_serial(value)
    if serial is None:
        return _text_cell_xml(ref, value.isoformat(), _S_DEFAULT)
    if isinstance(value, _dt.datetime):
        style = _S_DATETIME
    elif isinstance(value, _dt.time):
        style = _S_TIME
    else:
        style = _S_DATE
    return '<c r="%s" s="%d"><v>%r</v></c>' % (ref, style, serial)


def _row_xml(
    values: Sequence[Any],
    row_number: int,
    letters: Sequence[str],
    count: int,
    part: _Part,
    stats: _Stats,
) -> str:
    """Buduje XML jednego wiersza danych.  Gorąca pętla — stąd ręczne sklejanie."""
    if part.key or part.start != 0 or part.stop is not None:
        values = _part_values(values, part)
    parts = ['<row r="', str(row_number), '">']
    append = parts.append
    row_text = str(row_number)
    index = 0
    for value in values:
        if index >= count:
            break
        ref = letters[index] + row_text
        index += 1
        if type(value) is str and len(value) > MAX_CELL_CHARS:
            stats.truncated += 1  # sanitize_cell zaraz przytnie ten tekst
        value = sanitize_cell(value)
        if value is None:
            continue
        kind = type(value)
        if kind is str:
            if not value:
                continue
            if value != value.strip():
                append(
                    '<c r="%s" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
                    % (ref, _esc(value))
                )
            else:
                append('<c r="%s" t="inlineStr"><is><t>%s</t></is></c>' % (ref, _esc(value)))
        elif kind is int:
            append('<c r="%s"><v>%d</v></c>' % (ref, value))
        elif kind is float:
            append('<c r="%s"><v>%r</v></c>' % (ref, value))
        elif kind is bool:
            append('<c r="%s" t="b"><v>%d</v></c>' % (ref, 1 if value else 0))
        elif kind is _dt.datetime or kind is _dt.date or kind is _dt.time or kind is _dt.timedelta:
            append(_date_cell_xml(ref, value))
        else:
            append(_cell_xml_fallback(ref, value))
    append("</row>")
    return "".join(parts)


def _cols_xml(widths: Sequence[float]) -> str:
    """Buduje element ``<cols>``, sklejając sąsiednie kolumny o tej samej szerokości."""
    if not widths:
        return ""
    chunks = ["<cols>"]
    first = 0
    for index in range(1, len(widths) + 1):
        if index < len(widths) and widths[index] == widths[first]:
            continue
        chunks.append(
            '<col min="%d" max="%d" width="%.2f" customWidth="1"/>' % (first + 1, index, widths[first])
        )
        first = index
    chunks.append("</cols>")
    return "".join(chunks)


def _write_sheet_stdlib(sink: Any, part: _Part, options: _Options, stats: _Stats) -> int:
    """Zapisuje XML jednego arkusza do strumienia.  Zwraca liczbę wierszy danych."""
    columns = part.columns
    count = len(columns)
    letters = _col_letters(count)

    buffer: list[str] = []
    append = buffer.append
    append(_XML_DECL)
    append('<worksheet xmlns="%s" xmlns:r="%s">' % (_NS_MAIN, _NS_REL))
    if options.freeze_header and count:
        append(
            '<sheetViews><sheetView workbookViewId="0">'
            '<pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>'
            '<selection pane="bottomLeft" activeCell="A2" sqref="A2"/>'
            "</sheetView></sheetViews>"
        )
    else:
        append('<sheetViews><sheetView workbookViewId="0"/></sheetViews>')
    append('<sheetFormatPr defaultRowHeight="15"/>')
    if options.auto_width and count:
        append(_cols_xml(_estimate_widths(part)))
    append("<sheetData>")

    if count:
        header = ['<row r="1">']
        for index, column in enumerate(columns):
            # Nagłówek jest zwykłą komórką — musi przejść przez sanitize_cell.
            header.append(
                _text_cell_xml(letters[index] + "1", _header_text(column, stats), _S_HEADER)
            )
        header.append("</row>")
        append("".join(header))

    written = 0
    write = sink.write
    for row in itertools.chain(part.sample, part.rows):
        if not isinstance(row, (list, tuple)):
            row = list(row)
        written += 1
        append(_row_xml(row, written + 1, letters, count, part, stats))
        if len(buffer) >= _FLUSH_EVERY:
            write("".join(buffer).encode("utf-8"))
            del buffer[:]

    append("</sheetData>")
    if options.autofilter and count:
        append('<autoFilter ref="A1:%s%d"/>' % (letters[count - 1], written + 1))
    append('<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>')
    append("</worksheet>")
    write("".join(buffer).encode("utf-8"))
    stats.rows += written
    return written


_STYLES_XML = (
    _XML_DECL
    + '<styleSheet xmlns="%s">' % _NS_MAIN
    + '<numFmts count="4">'
    + '<numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>'
    + '<numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd\\ hh:mm:ss"/>'
    + '<numFmt numFmtId="166" formatCode="hh:mm:ss"/>'
    + '<numFmt numFmtId="167" formatCode="[h]:mm:ss"/>'
    + "</numFmts>"
    + '<fonts count="2">'
    + '<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>'
    + '<font><b/><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>'
    + "</fonts>"
    + '<fills count="2">'
    + '<fill><patternFill patternType="none"/></fill>'
    + '<fill><patternFill patternType="gray125"/></fill>'
    + "</fills>"
    + '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="6">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    + '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="166" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + '<xf numFmtId="167" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
    + "</cellXfs>"
    + '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
    + "</styleSheet>"
)

_ROOT_RELS_XML = (
    _XML_DECL
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/officeDocument" Target="xl/workbook.xml"/>'
    '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/'
    'relationships/metadata/core-properties" Target="docProps/core.xml"/>'
    '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/'
    'relationships/extended-properties" Target="docProps/app.xml"/>'
    "</Relationships>"
)

_CORE_XML = (
    _XML_DECL
    + '<cp:coreProperties '
    'xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" '
    'xmlns:dc="http://purl.org/dc/elements/1.1/" '
    'xmlns:dcterms="http://purl.org/dc/terms/" '
    'xmlns:dcmitype="http://purl.org/dc/dcmitype/" '
    'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
    "<dc:creator>flexit2xlsx</dc:creator>"
    "<cp:lastModifiedBy>flexit2xlsx</cp:lastModifiedBy>"
    "</cp:coreProperties>"
)

_APP_XML = (
    _XML_DECL
    + '<Properties '
    'xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" '
    'xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">'
    "<Application>flexit2xlsx</Application>"
    "</Properties>"
)


def _zinfo(name: str) -> zipfile.ZipInfo:
    """Tworzy wpis ZIP ze stałą datą (powtarzalny wynik) i kompresją deflate."""
    info = zipfile.ZipInfo(name, date_time=_ZIP_DATE)
    info.compress_type = zipfile.ZIP_DEFLATED
    info.external_attr = 0o644 << 16
    return info


def _workbook_xml(names: Sequence[str]) -> str:
    """Buduje ``xl/workbook.xml`` z listą arkuszy."""
    sheets = "".join(
        '<sheet name="%s" sheetId="%d" r:id="rId%d"/>'
        % (_esc_attr(_xml_safe_name(name)), index + 1, index + 1)
        for index, name in enumerate(names)
    )
    return (
        _XML_DECL
        + '<workbook xmlns="%s" xmlns:r="%s">' % (_NS_MAIN, _NS_REL)
        + '<workbookPr/><bookViews><workbookView/></bookViews>'
        + "<sheets>"
        + sheets
        + "</sheets></workbook>"
    )


def _workbook_rels_xml(count: int) -> str:
    """Buduje ``xl/_rels/workbook.xml.rels`` (arkusze + style)."""
    base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    items = [
        '<Relationship Id="rId%d" Type="%sworksheet" Target="worksheets/sheet%d.xml"/>'
        % (index + 1, base, index + 1)
        for index in range(count)
    ]
    items.append(
        '<Relationship Id="rId%d" Type="%sstyles" Target="styles.xml"/>' % (count + 1, base)
    )
    return (
        _XML_DECL
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        + "".join(items)
        + "</Relationships>"
    )


def _content_types_xml(count: int) -> str:
    """Buduje ``[Content_Types].xml`` dla wygenerowanego pakietu."""
    doc = "application/vnd.openxmlformats-officedocument.spreadsheetml."
    overrides = "".join(
        '<Override PartName="/xl/worksheets/sheet%d.xml" ContentType="%sworksheet+xml"/>'
        % (index + 1, doc)
        for index in range(count)
    )
    return (
        _XML_DECL
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" '
        'ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/xl/workbook.xml" ContentType="%ssheet.main+xml"/>' % doc
        + overrides
        + '<Override PartName="/xl/styles.xml" ContentType="%sstyles+xml"/>' % doc
        + '<Override PartName="/docProps/core.xml" '
        'ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
        '<Override PartName="/docProps/app.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
        "</Types>"
    )


def _write_stdlib(
    path: str, sheets: Sequence[Sheet], options: _Options, stats: _Stats
) -> None:
    """Zapisuje skoroszyt wyłącznie biblioteką standardową (zipfile + własny XML)."""
    used: set[str] = set()
    names: list[str] = []
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED, allowZip64=True) as archive:
        for sheet in sheets:
            for part in _iter_parts(sheet, used):
                names.append(part.name)
                entry = "xl/worksheets/sheet%d.xml" % len(names)
                with archive.open(_zinfo(entry), "w") as stream:
                    _write_sheet_stdlib(stream, part, options, stats)

        workbook_xml = _workbook_xml(names)
        # Asercja zamiast wyjątku z głębi zipfile: nazwa arkusza MUSI dać się
        # zapisać w UTF-8 (samotny surogat z nazwy pliku spoza UTF-8 by tu wybuchł).
        try:
            workbook_xml.encode("utf-8")
        except UnicodeEncodeError as exc:  # pragma: no cover - _xml_safe_name to usuwa
            raise ValueError(
                "Nazwa arkusza zawiera znak, którego nie da się zapisać w pliku XLSX "
                "— popraw nazwę pliku źródłowego albo użyj innej nazwy aukcji."
            ) from exc

        archive.writestr(_zinfo("xl/styles.xml"), _STYLES_XML)
        archive.writestr(_zinfo("xl/workbook.xml"), workbook_xml)
        archive.writestr(_zinfo("xl/_rels/workbook.xml.rels"), _workbook_rels_xml(len(names)))
        archive.writestr(_zinfo("docProps/core.xml"), _CORE_XML)
        archive.writestr(_zinfo("docProps/app.xml"), _APP_XML)
        archive.writestr(_zinfo("_rels/.rels"), _ROOT_RELS_XML)
        archive.writestr(_zinfo("[Content_Types].xml"), _content_types_xml(len(names)))


# --------------------------------------------------------------------------- #
# Ścieżka „openpyxl”
# --------------------------------------------------------------------------- #


def _for_openpyxl(value: Any) -> Any:
    """Sprowadza wartość do postaci, którą openpyxl zapisze tak jak zapis wbudowany.

    Jedyna rozbieżność, jaką trzeba wyrównać: openpyxl zapisuje daty sprzed
    1900-01-01 jako UJEMNY numer seryjny, którego Excel nie potrafi wyświetlić.
    Zamieniamy je — tak jak robi to ścieżka stdlib — na tekst ISO.
    """
    if isinstance(value, _dt.date) and _excel_serial(value) is None:
        return value.isoformat()
    return value


def _write_openpyxl(
    path: str, sheets: Sequence[Sheet], options: _Options, stats: _Stats
) -> None:
    """Zapisuje skoroszyt przez openpyxl w trybie ``write_only`` (strumieniowym)."""
    workbook = _openpyxl.Workbook(write_only=True)
    # Powtarzalny wynik: bez tego openpyxl wpisuje do docProps/core.xml bieżący
    # czas.  Pole `modified` i tak nadpisze sobie w save_workbook — dlatego
    # pełnej powtarzalności co do bajtu gwarantuje tylko wbudowany zapis.
    try:
        workbook.properties.created = _DOC_DATE
        workbook.properties.modified = _DOC_DATE
    except (AttributeError, TypeError):  # pragma: no cover - zależy od wersji openpyxl
        pass
    bold = _Font(bold=True)
    used: set[str] = set()

    for sheet in sheets:
        for part in _iter_parts(sheet, used):
            columns = part.columns
            count = len(columns)
            worksheet = workbook.create_sheet(title=_xml_safe_name(part.name))

            # Szerokości i zamrożenie nagłówka MUSZĄ być ustawione przed pierwszym
            # `append` — w trybie write_only openpyxl zapisuje nagłówek arkusza
            # (elementy <sheetViews> i <cols>) przy pierwszym dopisanym wierszu.
            if options.freeze_header and count:
                worksheet.freeze_panes = "A2"
            if options.auto_width and count:
                letters = _col_letters(count)
                widths = _estimate_widths(part)
                for index, width in enumerate(widths):
                    worksheet.column_dimensions[letters[index]].width = width

            if count:
                header = []
                for column in columns:
                    # Ten sam odkażony nagłówek co w ścieżce stdlib — inaczej
                    # openpyxl rzuca IllegalCharacterError, a limit 32767 znaków
                    # łamał się tylko w jednej ze ścieżek.
                    cell = _WriteOnlyCell(worksheet, value=_header_text(column, stats))
                    cell.font = bold
                    cell.data_type = "s"  # nagłówek nigdy nie jest formułą
                    header.append(cell)
                worksheet.append(header)

            sliced = part.key or part.start != 0 or part.stop is not None
            written = 0
            for row in itertools.chain(part.sample, part.rows):
                if not isinstance(row, (list, tuple)):
                    row = list(row)
                if sliced:
                    row = _part_values(row, part)
                cells: list[Any] = []
                for index, value in enumerate(row):
                    if index >= count:
                        break
                    if type(value) is str and len(value) > MAX_CELL_CHARS:
                        stats.truncated += 1
                    value = _for_openpyxl(sanitize_cell(value))
                    if isinstance(value, str) and (
                        looks_like_formula(value) or value in _ERROR_CODES
                    ):
                        # Napis zapisujemy jako tekst: openpyxl zrobiłby z "=..."
                        # formułę, a z "#N/A" komórkę BŁĘDU (t="e"), przez co ten
                        # sam XML dawałby inne dane niż wbudowany zapis.  Pozostałe
                        # napisy openpyxl i tak zapisuje jako tekst — nie owijamy
                        # ich, bo _WriteOnlyCell dla każdej komórki kosztuje.
                        cell = _WriteOnlyCell(worksheet, value=value)
                        cell.data_type = "s"
                        cells.append(cell)
                    else:
                        cells.append(value)
                worksheet.append(cells)
                written += 1
            stats.rows += written

            if options.autofilter and count:
                last = _col_letters(count)[-1]
                worksheet.auto_filter.ref = "A1:%s%d" % (last, written + 1)

    if not workbook.worksheets:  # pusty skoroszyt nie jest poprawnym plikiem XLSX
        workbook.create_sheet(title=DEFAULT_SHEET_NAME)
    workbook.save(path)


# --------------------------------------------------------------------------- #
# API publiczne
# --------------------------------------------------------------------------- #


def _apply_default_mode(temporary: str, target: str) -> None:
    """Nadaje plikowi tymczasowemu uprawnienia, jakie miałby zwykły ``open(..., "wb")``.

    ``tempfile.mkstemp`` tworzy plik z trybem ``0600``, a ``os.replace`` ten tryb
    zachowuje — użytkownik dostawał plik nie do odczytania przez inne konto
    (Samba, serwer WWW), choć niczego takiego nie ustawiał.

    Nadpisując istniejący plik, zachowujemy JEGO uprawnienia (tak samo jak
    ``open``); dla nowego pliku bierzemy je z ``umask`` użytkownika.
    """
    try:
        try:
            mode = stat.S_IMODE(os.stat(target).st_mode)
        except OSError:  # pliku jeszcze nie ma — decyduje umask
            mask = os.umask(0)
            os.umask(mask)
            mode = 0o666 & ~mask
        os.chmod(temporary, mode)
    except OSError:  # pragma: no cover - np. system plików bez uprawnień POSIX
        pass


def _report(stats: _Stats, backend: str) -> None:
    """Wypisuje na stderr to, o czym użytkownik powinien wiedzieć po zapisie."""
    if stats.truncated:
        _warn(
            "UWAGA: skrócono %s do %d znaków (limit Excela) — pełna treść "
            "została w plikach XML"
            % (
                _plural(stats.truncated, "%d komórkę", "%d komórki", "%d komórek"),
                MAX_CELL_CHARS,
            )
        )
    if (
        backend == "openpyxl"
        and stats.rows > SLOW_BACKEND_HINT_ROWS
        and (os.environ.get(ENV_BACKEND) or "auto").strip().lower() in ("", "auto")
    ):
        _warn(
            "zapisano %d wierszy przez openpyxl; przy takich rozmiarach wbudowany "
            "zapis jest około dwa razy szybszy — ustaw %s=stdlib"
            % (stats.rows, ENV_BACKEND)
        )


def write_workbook(
    path,
    sheets: list[Sheet],
    *,
    freeze_header: bool = True,
    autofilter: bool = True,
    auto_width: bool = True,
) -> None:
    """Zapisuje listę arkuszy do pliku ``.xlsx``.

    :param path: ścieżka pliku wynikowego (``str`` albo ``os.PathLike``),
    :param sheets: lista obiektów :class:`Sheet`; pusta lista da skoroszyt
        z jednym pustym arkuszem (plik XLSX bez arkuszy jest niepoprawny),
    :param freeze_header: zamrożenie wiersza nagłówka,
    :param autofilter: autofiltr na zakresie danych,
    :param auto_width: automatyczny dobór szerokości kolumn.

    Ścieżkę zapisu wybiera :func:`backend_name`.  Zapis jest atomowy — plik
    docelowy powstaje dopiero po pomyślnym zakończeniu całości, a jego uprawnienia
    wynikają z ``umask`` użytkownika.

    Po zapisie na stderr trafia informacja o komórkach skróconych do limitu
    Excela (32767 znaków) — utrata treści nigdy nie jest cicha.
    """
    target = os.fspath(path)
    plan = list(sheets) if sheets else [Sheet(name=DEFAULT_SHEET_NAME, columns=[], rows=[])]
    options = _Options(
        freeze_header=bool(freeze_header),
        autofilter=bool(autofilter),
        auto_width=bool(auto_width),
    )

    stats = _Stats()
    backend = backend_name()

    directory = os.path.dirname(os.path.abspath(target))
    handle, temporary = tempfile.mkstemp(prefix=".flexit2xlsx-", suffix=".tmp", dir=directory)
    os.close(handle)
    try:
        if backend == "openpyxl":
            _write_openpyxl(temporary, plan, options, stats)
        else:
            _write_stdlib(temporary, plan, options, stats)
        _apply_default_mode(temporary, target)
        os.replace(temporary, target)
    except BaseException:
        try:
            os.unlink(temporary)
        except OSError:  # pragma: no cover - plik mógł już zniknąć
            pass
        raise

    _report(stats, backend)
