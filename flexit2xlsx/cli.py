"""Interfejs wiersza poleceń: spina scrape + xmlflatten + values + xlsxwrite.

Podkomendy
----------

``download``
    Pobiera pliki XML "zawartość pakietu" z portalu do katalogu.
``build``
    Buduje JEDEN plik ``.xlsx`` z katalogu (albo listy) plików XML.
``all``
    ``download`` + ``build`` w jednym uruchomieniu.

Arkusze w pliku wynikowym
-------------------------

1. ``Wszystkie aukcje`` — wszystkie wiersze ze wszystkich plików; unia kolumn.
   Na początku kolumny techniczne: ``Aukcja``, ``Plik``, ``Nr pozycji``.
2. ``Podsumowanie`` — po jednym wierszu na plik: aukcja, liczba pozycji,
   wykryta ścieżka rekordu, status (``OK`` albo treść błędu).
3. Opcjonalnie (``--per-auction``) — osobny arkusz na każdą aukcję.

Kody wyjścia
------------

======  =====================================================================
   0    wszystko się udało
   1    błąd ogólny (np. brak uprawnień do zapisu, plik wyjściowy już istnieje)
   2    błąd składni polecenia (kod argparse)
   3    zrobione częściowo — wynik powstał, ale część plików się nie wczytała
   4    nie znaleziono danych (żadnej aukcji, żadnego XML-a, żadnego pliku)
   5    błąd komunikacji z portalem (sieć, HTTP, blokada)
 130    przerwane przez użytkownika (Ctrl+C)
======  =====================================================================

Decyzje projektowe (świadome, bo kontrakt ich nie rozstrzyga)
-------------------------------------------------------------

* Kolumny techniczne mają STAŁE nazwy; gdy XML ma własne pole o takiej samej
  nazwie, kolumna z XML-a dostaje przyrostek ``(XML)`` — nic nie ginie.
* Wejściem może być katalog, POJEDYNCZY PLIK albo ARCHIWUM ``.zip``/``.gz``.
  Pliki z archiwum są wypakowywane do katalogu tymczasowego (sprzątanego przy
  wyjściu) pod nazwą niosącą ślad pochodzenia — kolumna ``Plik`` dalej mówi,
  skąd wziął się wiersz.
* Pliki o IDENTYCZNEJ treści są wykrywane (rozmiar, potem suma kontrolna).
  Domyślnie tylko ostrzegamy, bo to użytkownik wie, czy druga kopia jest
  pomyłką; ``--skip-duplicates`` każe je pominąć.
* ``--csv`` zapisuje arkusz zbiorczy DODATKOWO jako CSV, a gdy zapis XLSX się
  nie powiedzie, CSV powstaje automatycznie (zapis ratunkowy) — nieudany zapis
  nie może kosztować godzin pobierania.
* ``build`` NIE nadpisuje istniejącego pliku ``.xlsx`` bez ``--overwrite``
  (łatwo pomylić katalogi; zniszczony wynik pracy boli bardziej niż komunikat).
* Błąd pojedynczego pliku XML nigdy nie przerywa całości — trafia do arkusza
  ``Podsumowanie``, na ``stderr`` i do kodu wyjścia 3.
* Kolejność plików jest deterministyczna (sortowanie po ścieżce), żeby dwa
  uruchomienia dawały identyczny wynik.
* **Potok jest strumieniowy.**  ``load_documents`` zwraca lekkie uchwyty
  (:class:`DocHandle`); rekordy zostają w pamięci tylko do wysokości
  :data:`MEMORY_ROW_BUDGET`, a powyżej — plik jest parsowany ponownie dopiero
  wtedy, gdy generator wierszy do niego dojdzie.  Kosztem jest drugi odczyt
  z dysku, zyskiem — szczyt pamięci ``O(największy plik)`` zamiast
  ``O(cały korpus)``.
* ``all`` sprawdza plik z ``--out`` PRZED pierwszym żądaniem HTTP — inaczej
  kilkanaście minut uprzejmego pobierania kończyłoby się komunikatem
  "dodaj --overwrite".
* Żaden błąd nie wychodzi z ``main`` tracebackiem: ``MemoryError`` i wszystko
  inne kończy się zdaniem po polsku i kodem z tabeli powyżej.
* Pomoc i komunikaty składniowe argparse są tłumaczone na polski
  (:class:`PolishArgumentParser`), bo to one witają użytkownika przy pomyłce.
"""

from __future__ import annotations

import argparse
import atexit
import csv
import gzip
import hashlib
import os
import re
import shutil
import sys
import tempfile
import zipfile
from typing import Any, Dict, Iterable, Iterator, List, Optional, Sequence, Set, Tuple

from . import values, xlsxwrite, xmlflatten
from .xlsxwrite import Sheet

__all__ = [
    "main",
    "build_parser",
    "collect_xml_files",
    "expand_archive",
    "find_duplicate_files",
    "load_documents",
    "build_sheets",
    "collective_table",
    "write_csv",
    "TECH_COLUMNS",
    "SHEET_ALL",
    "SHEET_SUMMARY",
    "SUMMARY_COLUMNS",
    "EXIT_OK",
    "EXIT_ERROR",
    "EXIT_USAGE",
    "EXIT_PARTIAL",
    "EXIT_NO_DATA",
    "EXIT_NETWORK",
    "EXIT_INTERRUPTED",
]

# --------------------------------------------------------------------------- #
# Stałe
# --------------------------------------------------------------------------- #

PROG = "flexit2xlsx"

#: Kolumny techniczne dopisywane przez CLI na początku arkusza zbiorczego.
TECH_COLUMNS = ["Aukcja", "Plik", "Nr pozycji"]

#: Kolumny techniczne arkusza pojedynczej aukcji (``Aukcja`` byłaby stała).
TECH_COLUMNS_PER_AUCTION = ["Plik", "Nr pozycji"]

SHEET_ALL = "Wszystkie aukcje"
SHEET_SUMMARY = "Podsumowanie"

SUMMARY_COLUMNS = [
    "Plik",
    "Aukcja",
    "Liczba pozycji",
    "Ścieżka rekordu",
    "Liczba kolumn",
    "Status",
]

#: Domyślny katalog na pobrane XML-e.
DEFAULT_XML_DIR = "xml_flexit"

#: Domyślna nazwa pliku wynikowego.
DEFAULT_XLSX = "aukcje.xlsx"

#: Domyślny adres portalu.
DEFAULT_BASE_URL = "https://flexitauctions.com/"

#: Rozszerzenia uznawane za pliki XML przy przeszukiwaniu katalogu.
XML_SUFFIXES = (".xml",)

#: Archiwa, z których CLI samo wyjmuje pliki XML.  Portal (i przeglądarka przy
#: "pobierz wszystko") potrafi oddać paczkę ZIP, a użytkownik pakuje katalog,
#: zanim go przeniesie na inny komputer — bez tego dostawał "nie znalazłem
#: żadnego pliku .xml" nad katalogiem pełnym danych.
ZIP_SUFFIXES = (".zip",)

#: Pojedyncze pliki spakowane gzipem (``batch.xml.gz``).
GZIP_SUFFIXES = (".gz",)

#: Górny limit sumy bajtów wypakowanych z JEDNEGO archiwum — ochrona przed
#: "bombą zip" (kilkadziesiąt kB potrafi rozwinąć się w gigabajty).
MAX_ARCHIVE_BYTES = 512 * 1024 * 1024

#: Ile poziomów archiwum w archiwum rozpakowywać (paczka paczek — po jednej na
#: lot — zdarza się przy "pobierz wszystko"; głębiej to już tylko ryzyko).
MAX_ARCHIVE_DEPTH = 2

#: Co ile wierszy pokazywać postęp przy zapisie (najdłuższa i dotąd całkiem
#: niema faza pracy: przy 300 plikach potrafi trwać minuty).
PROGRESS_ROWS = 25_000

#: Domyślny separator pola w eksporcie CSV.  Średnik, bo taki jest domyślny
#: separator listy w polskim Excelu — plik otwiera się dwuklikiem, bez kreatora.
DEFAULT_CSV_SEP = ";"

#: Ile wierszy wolno trzymać naraz w pamięci po wczytaniu plików.
#:
#: Poniżej tej granicy dokumenty zostają w pamięci (jedno parsowanie, szybciej),
#: powyżej — CLI je zwalnia i parsuje pliki ponownie dopiero w chwili zapisu.
#: Dzięki temu szczyt pamięci zależy od NAJWIĘKSZEGO pliku, a nie od sumy
#: wszystkich (``xlsxwrite`` i tak zapisuje strumieniowo).
MEMORY_ROW_BUDGET = 100_000

#: Ile kandydatów na ścieżkę rekordu wolno wypróbować w ``unify_record_paths``.
#: Dalsze i tak prawie nigdy nie pasują, a każdy kosztuje jedno parsowanie.
UNIFY_MAX_CANDIDATES = 3

#: Ile wyników ``coerce_value`` pamiętać (wartości w kolumnach powtarzają się
#: masowo: grade, jednostki, pola kontekstu w każdym wierszu).  Rozmiar dobrany
#: tak, żeby bufor sam nie zjadł zysku z pracy strumieniowej.
TYPE_CACHE_MAX = 50_000

#: Dłuższych napisów nie cache'ujemy — pamięć ważniejsza niż te kilka trafień.
TYPE_CACHE_MAX_LEN = 200

EXIT_OK = 0
EXIT_ERROR = 1
EXIT_USAGE = 2
EXIT_PARTIAL = 3
EXIT_NO_DATA = 4
EXIT_NETWORK = 5
EXIT_INTERRUPTED = 130


# --------------------------------------------------------------------------- #
# Komunikaty
# --------------------------------------------------------------------------- #


class Reporter:
    """Wypisywanie komunikatów do użytkownika (po polsku).

    Informacje idą na ``stdout`` i milkną przy ``--quiet``; ostrzeżenia i błędy
    idą na ``stderr`` ZAWSZE — nawet w trybie cichym, bo inaczej użytkownik
    nigdy by się nie dowiedział, że coś poszło nie tak.
    """

    def __init__(self, quiet: bool = False, out=None, err=None) -> None:
        self.quiet = bool(quiet)
        self._out = out
        self._err = err

    # Strumienie rozwiązujemy przy każdym wypisaniu, żeby działały podmiany
    # ``contextlib.redirect_stdout`` w testach.
    @property
    def out(self):
        return self._out if self._out is not None else sys.stdout

    @property
    def err(self):
        return self._err if self._err is not None else sys.stderr

    def info(self, message: str = "") -> None:
        """Zwykły komunikat postępu."""
        if not self.quiet:
            print(message, file=self.out)

    def step(self, message: str) -> None:
        """Wyróżniony krok scenariusza."""
        self.info("==> " + message)

    def warn(self, message: str) -> None:
        """Ostrzeżenie — praca trwa dalej."""
        print("UWAGA: " + message, file=self.err)

    def error(self, message: str) -> None:
        """Błąd — coś się nie udało."""
        print("BŁĄD: " + message, file=self.err)


def _plural(count: int, one: str, few: str, many: str) -> str:
    """Polska odmiana liczebnika: ``1 plik``, ``2 pliki``, ``5 plików``."""
    if count == 1:
        return "%d %s" % (count, one)
    if 2 <= count % 10 <= 4 and not 12 <= count % 100 <= 14:
        return "%d %s" % (count, few)
    return "%d %s" % (count, many)


def _files(count: int) -> str:
    return _plural(count, "plik", "pliki", "plików")


def _rows(count: int) -> str:
    return _plural(count, "wiersz", "wiersze", "wierszy")


def _short(text: str, limit: int = 300) -> str:
    """Skraca długi komunikat błędu do jednej, czytelnej linijki."""
    flat = " ".join(str(text).split())
    return flat if len(flat) <= limit else flat[: limit - 1] + "…"


# --------------------------------------------------------------------------- #
# Wejście: archiwa ZIP / GZIP
# --------------------------------------------------------------------------- #

#: Katalog tymczasowy na pliki wyjęte z archiwów (tworzony leniwie, jeden na
#: uruchomienie, sprzątany przy wyjściu z programu).
_SCRATCH: List[str] = []

#: Nazwy plików już zajęte w katalogu tymczasowym (bez rozróżniania wielkości).
_USED_NAMES: Set[str] = set()

#: Zapamiętane wyniki wypakowania: ``(ścieżka, rozmiar, mtime) -> lista plików``.
#: Bez tego ``all`` wypakowywałoby to samo archiwum dwa razy (raz przy
#: sprawdzeniu katalogu, raz przy budowaniu) i dublowało wiersze.
_ARCHIVE_CACHE: Dict[Tuple[str, int, int], List[str]] = {}

#: Wszystko poza literami, cyframi, ``.``, ``-`` i ``_`` zamieniamy na ``_``.
_SAFE_NAME_RE = re.compile(r"[^\w.-]+", re.UNICODE)


def cleanup_archives() -> None:
    """Usuwa pliki wypakowane z archiwów (wołane automatycznie przy wyjściu)."""
    while _SCRATCH:
        shutil.rmtree(_SCRATCH.pop(), ignore_errors=True)
    _USED_NAMES.clear()
    _ARCHIVE_CACHE.clear()


def _scratch_dir() -> str:
    """Katalog tymczasowy na wypakowane pliki (tworzony przy pierwszym użyciu)."""
    if not _SCRATCH:
        _SCRATCH.append(tempfile.mkdtemp(prefix="flexit2xlsx-"))
        atexit.register(cleanup_archives)
    return _SCRATCH[0]


def _safe_base(name: str) -> str:
    """Bezpieczna nazwa pliku z (niezaufanej) nazwy wpisu w archiwum.

    Archiwum może zawierać wpisy typu ``../../.bashrc`` albo ``C:\\Windows\\x``.
    Struktury katalogów NIE odtwarzamy: cała nazwa staje się JEDNYM członem,
    więc nie da się wyjść poza katalog tymczasowy (path traversal).
    """
    flat = _SAFE_NAME_RE.sub("_", str(name).replace("\\", "/").strip("/")).strip("_")
    if not flat or flat in (".", ".."):
        flat = "plik.xml"
    return flat[-120:]


def _scratch_path(name: str) -> str:
    """Wolna ścieżka w katalogu tymczasowym dla pliku o (mniej więcej) tej nazwie."""
    base = _safe_base(name)
    root, ext = os.path.splitext(base)
    candidate = base
    counter = 2
    while candidate.lower() in _USED_NAMES:
        candidate = "%s_%d%s" % (root, counter, ext)
        counter += 1
    _USED_NAMES.add(candidate.lower())
    return os.path.join(_scratch_dir(), candidate)


def _copy_limited(source: Any, sink: Any, budget: int) -> int:
    """Przepisuje strumień, pilnując limitu bajtów (ochrona przed bombą zip)."""
    written = 0
    while True:
        chunk = source.read(256 * 1024)
        if not chunk:
            return written
        written += len(chunk)
        if written > budget:
            raise ValueError(
                "po rozpakowaniu przekracza limit %d MB" % (MAX_ARCHIVE_BYTES // 1048576)
            )
        sink.write(chunk)


def _extract_zip(path: str, reporter: Reporter, depth: int = 0) -> List[str]:
    """Wyjmuje pliki ``*.xml`` z archiwum ZIP do katalogu tymczasowego.

    Archiwum w archiwum (portal potrafi oddać paczkę paczek — jedną na lot)
    jest rozpakowywane rekurencyjnie do :data:`MAX_ARCHIVE_DEPTH` poziomów.
    """
    label = os.path.basename(path)
    stem = os.path.splitext(label)[0]
    found: List[str] = []
    nested: List[str] = []
    budget = MAX_ARCHIVE_BYTES
    try:
        with zipfile.ZipFile(path) as archive:
            members = [info for info in archive.infolist() if not info.is_dir()]
            members.sort(key=lambda info: info.filename)
            wanted = [
                info for info in members
                if (info.filename.lower().endswith(XML_SUFFIXES)
                    or (depth < MAX_ARCHIVE_DEPTH and _is_archive(info.filename)))
                and not info.filename.startswith("__MACOSX/")
                and not os.path.basename(info.filename).startswith(".")
            ]
            if not wanted:
                reporter.warn(
                    "Archiwum %s nie zawiera plików .xml (wpisów w środku: %d)"
                    % (label, len(members))
                )
                return []
            for info in wanted:
                try:
                    destination = _scratch_path("%s__%s" % (stem, info.filename))
                    with archive.open(info) as source:
                        with open(destination, "wb") as sink:
                            budget -= _copy_limited(source, sink, budget)
                except RuntimeError as exc:      # archiwum zabezpieczone hasłem
                    reporter.warn(
                        "Nie wyjmę %s z %s — %s. Rozpakuj archiwum ręcznie "
                        "(program nie zna hasła) i wskaż katalog przez --in."
                        % (info.filename, label, _short(exc))
                    )
                    continue
                except (zipfile.BadZipFile, OSError, EOFError, ValueError) as exc:
                    reporter.warn(
                        "Nie wyjmę %s z %s — %s" % (info.filename, label, _short(exc))
                    )
                    continue
                if _is_archive(destination):
                    nested.append(destination)
                else:
                    found.append(destination)
    except (zipfile.BadZipFile, OSError, EOFError, ValueError, RuntimeError) as exc:
        reporter.warn("Pomijam archiwum %s — %s" % (label, _short(exc)))
        return []
    if found:
        reporter.info("  archiwum %s: wyjęto %s" % (label, _files(len(found))))
    for inner in nested:
        if inner.lower().endswith(ZIP_SUFFIXES):
            found.extend(_extract_zip(inner, reporter, depth + 1))
        else:
            found.extend(_extract_gzip(inner, reporter))
    return found


def _extract_gzip(path: str, reporter: Reporter) -> List[str]:
    """Rozpakowuje pojedynczy plik ``*.gz`` (typowo ``batch.xml.gz``)."""
    label = os.path.basename(path)
    inner = label[:-3] if label.lower().endswith(".gz") else label
    if not inner.lower().endswith(XML_SUFFIXES):
        inner += ".xml"
    destination = _scratch_path(inner)
    try:
        with gzip.open(path, "rb") as source:
            with open(destination, "wb") as sink:
                _copy_limited(source, sink, MAX_ARCHIVE_BYTES)
    except (OSError, EOFError, ValueError) as exc:
        reporter.warn("Pomijam %s — nie mogę rozpakować (%s)" % (label, _short(exc)))
        return []
    reporter.info("  rozpakowano %s" % label)
    return [destination]


def _is_archive(name: str) -> bool:
    """Czy nazwa wygląda na archiwum, które umiemy otworzyć?"""
    lower = name.lower()
    return lower.endswith(ZIP_SUFFIXES) or lower.endswith(GZIP_SUFFIXES)


def expand_archive(path: str, reporter: Optional[Reporter] = None) -> List[str]:
    """Zwraca pliki XML wyjęte z archiwum ``path`` (ZIP albo GZIP).

    Wynik jest zapamiętywany po ``(ścieżka, rozmiar, czas modyfikacji)``, więc
    dwa wywołania dla tego samego archiwum dają te SAME ścieżki — inaczej
    podkomenda ``all`` (która ogląda katalog dwa razy) zdublowałaby wiersze.
    Wypakowane pliki znikają razem z końcem programu.
    """
    reporter = reporter or Reporter(quiet=True)
    try:
        info = os.stat(path)
        key: Optional[Tuple[str, int, int]] = (
            os.path.normcase(os.path.abspath(path)), info.st_size, int(info.st_mtime)
        )
    except OSError:
        key = None
    if key is not None and key in _ARCHIVE_CACHE:
        return list(_ARCHIVE_CACHE[key])
    if path.lower().endswith(ZIP_SUFFIXES):
        found = _extract_zip(path, reporter)
    else:
        found = _extract_gzip(path, reporter)
    if key is not None:
        _ARCHIVE_CACHE[key] = list(found)
    return found


# --------------------------------------------------------------------------- #
# Wejście: zbieranie plików XML
# --------------------------------------------------------------------------- #


def collect_xml_files(paths: Sequence[str], reporter: Optional[Reporter] = None) -> List[str]:
    """Zamienia listę ścieżek (katalogi, pliki, archiwa) na listę plików XML.

    * katalog — przeszukiwany rekurencyjnie, brane pliki ``*.xml`` oraz archiwa
      ``*.zip`` / ``*.gz`` (z pominięciem katalogów ukrytych i plików ``.part``
      po przerwanym pobieraniu),
    * archiwum — pliki ``*.xml`` z jego wnętrza są wypakowywane do katalogu
      tymczasowego i traktowane jak zwykłe pliki wejściowe,
    * plik — brany dosłownie, niezależnie od rozszerzenia (użytkownik wie, co robi),
    * ścieżka nieistniejąca — ostrzeżenie, praca trwa dalej.

    Kolejność jest deterministyczna: pliki podane wprost zachowują kolejność
    z linii poleceń, zawartość katalogu jest posortowana alfabetycznie.
    """
    reporter = reporter or Reporter(quiet=True)
    found: List[str] = []
    seen = set()

    def remember(path: str) -> None:
        key = os.path.normcase(os.path.abspath(path))
        if key not in seen:
            seen.add(key)
            found.append(path)

    for raw in paths:
        path = os.path.expanduser(str(raw))
        if os.path.isdir(path):
            batch: List[str] = []
            for root, dirs, names in os.walk(path):
                dirs[:] = sorted(d for d in dirs if not d.startswith("."))
                for name in sorted(names):
                    if name.startswith("."):
                        continue
                    if name.lower().endswith(XML_SUFFIXES):
                        batch.append(os.path.join(root, name))
                    elif _is_archive(name):
                        batch.extend(expand_archive(os.path.join(root, name), reporter))
            if not batch:
                reporter.warn(
                    "Katalog %s nie zawiera plików .xml ani archiwów .zip/.gz" % path
                )
            for item in batch:
                remember(item)
        elif os.path.isfile(path):
            if _is_archive(path):
                for item in expand_archive(path, reporter):
                    remember(item)
            else:
                remember(path)
        else:
            reporter.warn("Pomijam %s — nie ma takiego pliku ani katalogu" % path)
    return found


# --------------------------------------------------------------------------- #
# Wejście: powtórzone pliki
# --------------------------------------------------------------------------- #


def _digest(path: str) -> Optional[str]:
    """Suma kontrolna zawartości pliku (``None``, gdy pliku nie da się czytać)."""
    checksum = hashlib.sha256()
    try:
        with open(path, "rb") as handle:
            for chunk in iter(lambda: handle.read(256 * 1024), b""):
                checksum.update(chunk)
    except OSError:
        return None
    return checksum.hexdigest()


def find_duplicate_files(files: Sequence[str]) -> Dict[str, str]:
    """Wskazuje pliki o treści IDENTYCZNEJ z innym plikiem z listy.

    Ten sam pakiet bywa na portalu pod dwoma adresami (``/lot/…`` i
    ``/auction/…/…``), a użytkownik pobiera go raz z przeglądarki i raz
    programem — w arkuszu te same sztuki policzyłyby się dwa razy, cicho
    zawyżając stan magazynu.  Zwraca odwzorowanie
    ``duplikat -> pierwszy plik o tej samej treści``.

    Najpierw grupujemy po rozmiarze (identyczne pliki MUSZĄ mieć ten sam
    rozmiar), więc w katalogu bez powtórzeń nie liczymy żadnej sumy kontrolnej.
    """
    by_size: Dict[int, List[str]] = {}
    for path in files:
        try:
            by_size.setdefault(os.path.getsize(path), []).append(path)
        except OSError:
            continue
    duplicates: Dict[str, str] = {}
    for group in by_size.values():
        if len(group) < 2:
            continue
        first_with: Dict[str, str] = {}
        for path in group:
            digest = _digest(path)
            if digest is None:
                continue
            if digest in first_with:
                duplicates[path] = first_with[digest]
            else:
                first_with[digest] = path
    return duplicates


# --------------------------------------------------------------------------- #
# Wczytywanie dokumentów
# --------------------------------------------------------------------------- #


class DocHandle:
    """Wczytany plik XML: metadane zawsze w pamięci, rekordy — na żądanie.

    Po co: ``load_documents`` trzymało wszystkie ``ParsedDoc`` (z kompletem
    rekordów) aż do końca zapisu, więc szczyt pamięci rósł LINIOWO z rozmiarem
    całego korpusu (~715 B na wiersz; milion wierszy = ponad 800 MB).  Tymczasem
    ``xlsxwrite`` zapisuje strumieniowo i potrzebuje tylko jednego wiersza naraz.

    Uchwyt pamięta to, czego potrzebuje arkusz ``Podsumowanie`` i unia kolumn
    (źródło, aukcja, ścieżka rekordu, kolumny, liczba pozycji), a sam dokument
    trzyma tylko dopóki mieści się w budżecie :data:`MEMORY_ROW_BUDGET`.
    Powyżej budżetu plik jest parsowany ponownie dopiero w chwili, gdy przychodzi
    jego kolej w generatorze wierszy — szczyt pamięci spada z ``O(cały korpus)``
    do ``O(największy plik)``.
    """

    __slots__ = ("source", "auction", "record_path", "columns", "count", "_kw", "_doc")

    def __init__(self, doc: xmlflatten.ParsedDoc, parse_kw: Dict[str, Any]) -> None:
        self._kw = dict(parse_kw)
        self._doc = doc
        self._absorb(doc)

    def _absorb(self, doc: xmlflatten.ParsedDoc) -> None:
        self.source = doc.source
        self.auction = doc.auction
        self.record_path = doc.record_path
        self.columns = list(doc.columns)
        self.count = len(doc.records)

    @property
    def cached(self) -> bool:
        """Czy dokument jest jeszcze w pamięci (bez ponownego parsowania)?"""
        return self._doc is not None

    def release(self) -> None:
        """Zwalnia rekordy z pamięci — zostaje sama metryczka."""
        self._doc = None

    def load(self) -> xmlflatten.ParsedDoc:
        """Zwraca pełny dokument (parsując plik ponownie, gdy trzeba)."""
        if self._doc is not None:
            return self._doc
        return xmlflatten.parse_file(self.source, **self._kw)

    def adopt(self, doc: xmlflatten.ParsedDoc, record_path: Optional[str]) -> None:
        """Podmienia dokument po ujednoliceniu ścieżki rekordu."""
        self._kw["record_path"] = record_path
        if self._doc is not None:
            self._doc = doc
        self._absorb(doc)


def _doc_source(doc: Any) -> xmlflatten.ParsedDoc:
    """Pełny ``ParsedDoc`` — z uchwytu albo wprost (gdy ktoś podał dokument)."""
    loader = getattr(doc, "load", None)
    return loader() if callable(loader) else doc


def _doc_count(doc: Any) -> int:
    """Liczba pozycji bez wciągania rekordów do pamięci."""
    count = getattr(doc, "count", None)
    if count is None:
        return len(doc.records)
    return int(count)


def _looks_like_html(path: str) -> bool:
    """Czy plik zaczyna się jak strona HTML (zapisana ściana logowania)?"""
    try:
        with open(path, "rb") as handle:
            head = handle.read(2048)
    except OSError:
        return False
    head = head.lstrip(b"\xef\xbb\xbf").lstrip().lower()
    if head.startswith((b"<!doctype html", b"<html")):
        return True
    # bywa i tak: <?xml ...?> a zaraz potem XHTML-owa strona logowania
    return head.startswith(b"<?xml") and b"<html" in head[:1024]


def _archive_kind(path: str) -> Optional[str]:
    """Czy plik jest archiwum udającym XML? (``PK`` = ZIP, ``\\x1f\\x8b`` = GZIP)

    Zdarza się to często: przeglądarka zapisuje paczkę pod nazwą ``batch.xml``,
    a użytkownik widzi tylko "uszkodzony XML" i nie wie, że wystarczy zmienić
    rozszerzenie.
    """
    try:
        with open(path, "rb") as handle:
            head = handle.read(4)
    except OSError:
        return None
    if head[:4] == b"PK\x03\x04":
        return ".zip"
    if head[:2] == b"\x1f\x8b":
        return ".gz"
    return None


def _parse_failure_message(path: str, exc: BaseException) -> str:
    """Zamienia wyjątek parsera na komunikat, z którym użytkownik coś zrobi."""
    if isinstance(exc, MemoryError):
        return (
            "za mało pamięci na wczytanie tego pliku — podziel katalog na części, "
            "użyj --limit albo wskaż --record-path, żeby ograniczyć liczbę kolumn"
        )
    if _looks_like_html(path):
        return (
            "to strona HTML, a nie XML — prawdopodobnie zapisała się strona "
            "logowania portalu (wygasła sesja). Zaloguj się w przeglądarce i "
            "pobierz plik jeszcze raz albo użyj: download --cookie \"...\""
        )
    kind = _archive_kind(path)
    if kind:
        return (
            "to archiwum %s zapisane pod nazwą .xml — zmień rozszerzenie na %s, "
            "a program sam wyjmie z niego pliki XML"
            % (kind.lstrip(".").upper(), kind)
        )
    return _short(exc)


def load_documents(
    files: Sequence[str],
    *,
    repeat: str = "join",
    join_sep: str = " | ",
    record_path: Optional[str] = None,
    strip_ns: bool = True,
    reporter: Optional[Reporter] = None,
) -> Tuple[List[Any], List[Tuple[str, str]]]:
    """Parsuje pliki XML; błąd JEDNEGO pliku nie przerywa całości.

    Zwraca ``(uchwyty, błędy)``, gdzie uchwyt to :class:`DocHandle` (udostępnia
    ``source``, ``auction``, ``record_path``, ``columns`` i ``count``), a błąd —
    para ``(ścieżka, komunikat)``.  Rekordy są trzymane w pamięci tylko do
    wysokości :data:`MEMORY_ROW_BUDGET`; powyżej niej pliki są parsowane
    ponownie dopiero przy generowaniu wierszy.
    """
    reporter = reporter or Reporter(quiet=True)
    parse_kw = {
        "repeat": repeat,
        "join_sep": join_sep,
        "record_path": record_path,
        "strip_ns": strip_ns,
    }
    docs: List[Any] = []
    errors: List[Tuple[str, str]] = []
    kept_rows = 0
    streaming = False
    total = len(files)
    for number, path in enumerate(files, 1):
        try:
            doc = xmlflatten.parse_file(path, **parse_kw)
        # MemoryError NIE dziedziczy po ValueError/OSError, a bez tej gałęzi
        # brak pamięci kończył się surowym tracebackiem zamiast komunikatem
        except (ValueError, OSError, MemoryError) as exc:
            message = _parse_failure_message(path, exc)
            errors.append((path, message))
            reporter.warn("Pomijam %s — %s" % (os.path.basename(path), message))
            continue
        handle = DocHandle(doc, parse_kw)
        del doc
        docs.append(handle)
        if streaming:
            handle.release()
        else:
            kept_rows += handle.count
            if kept_rows > MEMORY_ROW_BUDGET:
                # przekroczyliśmy budżet — od tej chwili pracujemy strumieniowo
                streaming = True
                for earlier in docs:
                    earlier.release()
        # Licznik [k/N] jest tu po to, żeby przy 300 plikach było widać, że coś
        # się dzieje i ile jeszcze zostało — bez niego praca wygląda na zawieszoną.
        reporter.info(
            "  [%d/%d] %-40s aukcja=%s, %s"
            % (
                number,
                total,
                _short(os.path.basename(path), 40),
                handle.auction,
                _plural(handle.count, "pozycja", "pozycje", "pozycji"),
            )
        )
    return docs, errors


def _tag_of(record_path: str) -> str:
    """Ostatni segment ścieżki rekordu (``batch/lot/items/item`` -> ``item``)."""
    return (record_path or "").rstrip("/").rsplit("/", 1)[-1]


def _tag_in_bytes(data: bytes, tag: str) -> bool:
    """Czy w surowych bajtach pliku w ogóle występuje znacznik o tej nazwie?

    Tani przedfiltr: bez niego ``unify_record_paths`` próbował KAŻDEJ ścieżki
    w KAŻDYM pliku (koszt iloczynowy — 400 plików x 400 ścieżek to 160 000
    parsowań i 22 MB odczytu dla katalogu ważącego 0,5 MB).  Gdy nazwy
    znacznika w pliku nie ma, parsowanie na pewno nic nie da.

    Nazwy spoza ASCII przepuszczamy bez sprawdzania (plik może być w innym
    kodowaniu niż UTF-8 — lepiej spróbować, niż zgubić dopasowanie).
    """
    if not tag:
        return False
    try:
        raw = tag.encode("ascii")
    except UnicodeEncodeError:
        return True
    pattern = rb"<\s*(?:[A-Za-z0-9_.\-]+:)?" + re.escape(raw) + rb"(?=[\s/>])"
    return re.search(pattern, data) is not None


def unify_record_paths(
    docs: List[Any],
    *,
    repeat: str = "join",
    join_sep: str = " | ",
    strip_ns: bool = True,
    reporter: Optional[Reporter] = None,
) -> set:
    """Ujednolica ścieżkę rekordu między plikami o TYM SAMYM schemacie.

    Problem: plik z jedną pozycją (``<items><item/></items>``) nie ma nic
    powtarzalnego, więc ``xmlflatten`` traktuje go jako JEDEN wiersz całego
    dokumentu — a wtedy jego kolumny nazywają się inaczej (``lot/items/item/model``)
    niż w pliku z trzema pozycjami (``model``).  Dla użytkownika, który chce
    JEDNEJ tabeli, to katastrofa: te same dane trafiają do dwóch kolumn.

    Rozwiązanie: bierzemy ścieżki rekordu wykryte w pozostałych plikach
    (najczęstsza pierwsza) i próbujemy wczytać ponownie te pliki, w których nic
    się nie powtarzało.  Gdy ścieżka pasuje — plik dostaje takie same kolumny
    jak reszta.  Gdy nie pasuje — zostaje bez zmian.

    Koszt jest LINIOWY względem liczby plików: bajty czytamy raz, kandydatów
    odsiewamy tanim przedfiltrem po nazwie znacznika, a i tak próbujemy najwyżej
    :data:`UNIFY_MAX_CANDIDATES` najczęstszych ścieżek.

    Modyfikuje listę ``docs`` w miejscu; zwraca zbiór ścieżek plików, które
    zostały ponownie wczytane.
    """
    reporter = reporter or Reporter(quiet=True)
    counts: Dict[str, int] = {}
    for doc in docs:
        if doc.record_path:
            counts[doc.record_path] = counts.get(doc.record_path, 0) + 1
    if not counts:
        return set()
    candidates = sorted(counts, key=lambda path: (-counts[path], path))

    unified = set()
    for index, doc in enumerate(docs):
        if doc.record_path or not os.path.exists(doc.source):
            continue
        try:
            with open(doc.source, "rb") as handle:
                data = handle.read()          # JEDEN odczyt na plik, nie N
        except (OSError, MemoryError):
            continue
        tried = 0
        for candidate in candidates:
            if not _tag_in_bytes(data, _tag_of(candidate)):
                continue                      # znacznika nie ma — nie ma czego parsować
            if tried >= UNIFY_MAX_CANDIDATES:
                break
            tried += 1
            try:
                fresh = xmlflatten.parse_bytes(
                    data,
                    doc.source,
                    repeat=repeat,
                    join_sep=join_sep,
                    record_path=candidate,
                    strip_ns=strip_ns,
                )
            except (ValueError, OSError, MemoryError):
                continue
            if fresh.record_path and fresh.records:
                adopt = getattr(doc, "adopt", None)
                if callable(adopt):
                    adopt(fresh, candidate)
                else:                         # ktoś podał gołe ParsedDoc
                    docs[index] = fresh
                unified.add(fresh.source)
                reporter.info(
                    "  ujednolicam kolumny: %s -> ścieżka rekordu %s"
                    % (os.path.basename(doc.source), candidate)
                )
                break
        del data
    return unified


# --------------------------------------------------------------------------- #
# Budowanie arkuszy
# --------------------------------------------------------------------------- #


def _convert(value: Any, use_typing: bool, cache: Optional[dict] = None) -> Any:
    """Zamienia napis na liczbę/datę/bool, gdy włączone jest rozpoznawanie typów.

    ``cache`` (opcjonalny) pamięta wyniki dla powtarzających się napisów.
    Rozpoznawanie typów było najdroższym elementem składania wierszy (5 mln
    wywołań ``coerce_value`` na 150 tys. wierszy), a wartości w kolumnach
    powtarzają się masowo: grade, stan, jednostki, pola kontekstu powtórzone
    w każdym wierszu.  Wynik ``coerce_value`` jest niezmienny, więc dzielenie go
    między komórki jest bezpieczne.
    """
    if not (use_typing and isinstance(value, str)):
        return value
    if cache is None or len(value) > TYPE_CACHE_MAX_LEN:
        return values.coerce_value(value)
    try:
        return cache[value]
    except KeyError:
        pass
    result = values.coerce_value(value)
    if len(cache) < TYPE_CACHE_MAX:
        cache[value] = result
    return result


def _unique_headers(columns: Sequence[str], reserved: Sequence[str]) -> List[str]:
    """Nadaje kolumnom z XML-a nagłówki nie kolidujące z kolumnami technicznymi.

    ``Aukcja`` z XML-a stanie się ``Aukcja (XML)``; dane nie giną, a użytkownik
    od razu widzi, która kolumna jest nasza, a która z pliku.
    """
    used = {str(name).lower() for name in reserved}
    headers: List[str] = []
    for column in columns:
        label = str(column)
        if label.lower() in used:
            candidate = label + " (XML)"
            counter = 2
            while candidate.lower() in used:
                candidate = "%s (XML %d)" % (label, counter)
                counter += 1
            label = candidate
        used.add(label.lower())
        headers.append(label)
    return headers


def _iter_rows(
    docs: Sequence[Any],
    columns: Sequence[str],
    *,
    with_auction: bool,
    use_typing: bool,
    reporter: Optional[Reporter] = None,
    failures: Optional[List[str]] = None,
) -> Iterator[List[Any]]:
    """Generuje wiersze arkusza: kolumny techniczne + wartości z XML-a.

    Świadomie jest to GENERATOR — arkusz może mieć setki tysięcy wierszy,
    a ``xlsxwrite`` zapisuje strumieniowo.  Dokument jest wczytywany dopiero
    wtedy, gdy przychodzi jego kolej, i porzucany zaraz po oddaniu wierszy —
    dzięki temu naraz w pamięci jest najwyżej JEDEN plik.

    ``failures`` (opcjonalna lista) zbiera pliki, których NIE udało się wczytać
    ponownie w chwili zapisu.  Ich wiersze nie trafią do arkusza, a arkusz
    ``Podsumowanie`` — zbudowany wcześniej — nadal liczy je jako wczytane;
    bez tej listy taka rozbieżność byłaby cichą utratą danych.
    """
    cache: Dict[str, Any] = {}
    done = 0
    next_report = PROGRESS_ROWS
    for handle in docs:
        filename = os.path.basename(handle.source) or handle.source
        try:
            doc = _doc_source(handle)
        except (ValueError, OSError, MemoryError) as exc:
            # plik zniknął albo zmienił się w trakcie pracy — lepiej głośno
            # pominąć jeden plik niż wywrócić cały zapis
            if failures is not None:
                failures.append(handle.source)
            if reporter is not None:
                reporter.warn(
                    "Nie udało się ponownie wczytać %s — pomijam jego wiersze (%s)"
                    % (filename, _short(exc))
                )
            continue
        auction = handle.auction
        for index, raw in enumerate(xmlflatten.rows_for(doc, columns), 1):
            cells = [_convert(value, use_typing, cache) for value in raw]
            done += 1
            if reporter is not None and done >= next_report:
                # zapis dużego korpusu to najdłuższa faza pracy — pokazujemy,
                # że postępuje, zamiast milczeć przez kilka minut
                next_report += PROGRESS_ROWS
                reporter.info("  ... zapisano %s" % _rows(done))
            if with_auction:
                yield [auction, filename, index] + cells
            else:
                yield [filename, index] + cells
        doc = None                     # zwolnij rekordy przed następnym plikiem


def _summary_rows(
    docs: Sequence[Any],
    errors: Sequence[Tuple[str, str]],
    column_count: int,
    unified: Sequence[str] = (),
) -> List[List[Any]]:
    """Buduje wiersze arkusza ``Podsumowanie`` (razem z błędami i sumą)."""
    unified = set(unified)
    rows: List[List[Any]] = []
    for doc in docs:
        rows.append(
            [
                os.path.basename(doc.source) or doc.source,
                doc.auction,
                _doc_count(doc),
                doc.record_path or "(brak — cały plik jako 1 wiersz)",
                len(doc.columns),
                "OK (ścieżka ujednolicona)" if doc.source in unified else "OK",
            ]
        )
    for path, message in errors:
        rows.append(
            [
                os.path.basename(path) or path,
                "",
                0,
                "",
                0,
                "BŁĄD: " + message,
            ]
        )
    total = sum(_doc_count(doc) for doc in docs)
    rows.append(
        [
            "RAZEM",
            "%s" % _plural(len({doc.auction for doc in docs}), "aukcja", "aukcje", "aukcji"),
            total,
            "",
            column_count,
            "wczytane: %s; z błędem: %s" % (_files(len(docs)), _files(len(errors))),
        ]
    )
    return rows


def _make_sheet(name: str, columns: List[str], rows: Any, key_columns: int) -> Sheet:
    """Buduje ``Sheet``, prosząc o powtórzenie kolumn kluczowych po podziale.

    Gdy kolumn jest więcej niż mieści arkusz Excela, ``xlsxwrite`` dzieli je na
    bloki.  Bez powtórzenia ``Aukcja``/``Plik``/``Nr pozycji`` wiersza w
    arkuszu-kontynuacji nie dałoby się przypisać do aukcji inaczej niż po
    numerze wiersza.  Parametr jest opcjonalny — gdy zapis go nie zna,
    budujemy arkusz po staremu (kontraktowe ``Sheet(name, columns, rows)``).
    """
    try:
        return Sheet(name=name, columns=columns, rows=rows, key_columns=key_columns)
    except TypeError:  # pragma: no cover - starsza wersja xlsxwrite
        return Sheet(name=name, columns=columns, rows=rows)


def _auction_sheet_name(auction: str, position: int, used: set) -> str:
    """Nazwa zakładki dla arkusza jednej aukcji — z ZACHOWANIEM końcówki.

    Identyfikatory aukcji (``flexit-auctions-18-06-2026-1103``) różnią się na
    KOŃCU, a limit Excela to 31 znaków.  Zwykłe obcięcie od prawej zostawiało
    same identyczne początki, rozróżniane potem automatycznym ``(2)``, ``(3)``.
    Dlatego numerujemy arkusze (zgodnie z kolejnością w ``Podsumowaniu``)
    i skracamy nazwę w ŚRODKU, zostawiając rozpoznawalny ogon.
    """
    prefix = "%02d " % position
    room = getattr(xlsxwrite, "SHEET_NAME_MAX", 31) - len(prefix)
    label = str(auction or "").strip() or SHEET_ALL
    if len(label) > room:
        head = max(1, (room - 1) * 2 // 5)
        tail = room - 1 - head
        label = label[:head] + "…" + label[-tail:]
    return xlsxwrite.safe_sheet_name(prefix + label, used)


def collective_table(
    docs: Sequence[Any],
    *,
    use_typing: bool = True,
    reporter: Optional[Reporter] = None,
    failures: Optional[List[str]] = None,
) -> Tuple[List[str], Iterator[List[Any]]]:
    """Nagłówki i wiersze arkusza zbiorczego — dokładnie to, co arkusz nr 1.

    Wydzielone, bo tę samą tabelę zapisuje też eksport CSV (i ratunkowy zapis,
    gdy XLSX się nie uda).  Wiersze są GENERATOREM — nie materializujemy ich.
    """
    data_columns = xmlflatten.merge_columns(list(docs))
    headers = TECH_COLUMNS + _unique_headers(data_columns, TECH_COLUMNS)
    rows = _iter_rows(docs, data_columns, with_auction=True, use_typing=use_typing,
                      reporter=reporter, failures=failures)
    return headers, rows


def build_sheets(
    docs: Sequence[Any],
    errors: Sequence[Tuple[str, str]] = (),
    *,
    per_auction: bool = False,
    use_typing: bool = True,
    unified: Sequence[str] = (),
    reporter: Optional[Reporter] = None,
    failures: Optional[List[str]] = None,
) -> List[Sheet]:
    """Składa listę arkuszy do zapisania przez :func:`xlsxwrite.write_workbook`."""
    data_columns = xmlflatten.merge_columns(list(docs))
    headers = _unique_headers(data_columns, TECH_COLUMNS)

    used_names: set = set()
    sheets: List[Sheet] = [
        _make_sheet(
            xlsxwrite.safe_sheet_name(SHEET_ALL, used_names),
            TECH_COLUMNS + headers,
            _iter_rows(docs, data_columns, with_auction=True, use_typing=use_typing,
                       reporter=reporter, failures=failures),
            len(TECH_COLUMNS),
        ),
        _make_sheet(
            xlsxwrite.safe_sheet_name(SHEET_SUMMARY, used_names),
            list(SUMMARY_COLUMNS),
            _summary_rows(docs, errors, len(data_columns), unified),
            0,
        ),
    ]

    if per_auction:
        groups: Dict[str, List[Any]] = {}
        for doc in docs:
            groups.setdefault(doc.auction, []).append(doc)
        for position, (auction, group) in enumerate(groups.items(), 1):
            own_columns = xmlflatten.merge_columns(group)
            own_headers = _unique_headers(own_columns, TECH_COLUMNS_PER_AUCTION)
            sheets.append(
                _make_sheet(
                    _auction_sheet_name(auction, position, used_names),
                    TECH_COLUMNS_PER_AUCTION + own_headers,
                    _iter_rows(group, own_columns, with_auction=False,
                               use_typing=use_typing, reporter=reporter,
                               failures=failures),
                    len(TECH_COLUMNS_PER_AUCTION),
                )
            )
    return sheets


# --------------------------------------------------------------------------- #
# Zapis awaryjny: CSV
# --------------------------------------------------------------------------- #


def _csv_text(value: str) -> str:
    """Czyści tekst jak :func:`values.sanitize_cell`, ale BEZ obcinania.

    Limit 32767 znaków jest limitem KOMÓRKI Excela — w pliku CSV nie
    obowiązuje, więc obcinanie byłoby tu bezcelową utratą treści.  Czyszczenie
    jest znak po znaku, więc dzielenie tekstu na kawałki niczego nie zmienia.
    """
    if len(value) <= values.MAX_CELL_CHARS:
        return values.sanitize_cell(value)
    step = values.MAX_CELL_CHARS
    return "".join(values.sanitize_cell(value[i:i + step])
                   for i in range(0, len(value), step))


def _csv_cell(value: Any, decimal_comma: bool) -> str:
    """Zamienia wartość na tekst do CSV: bez utraty treści i bez formuł.

    * ``None`` -> pusta komórka, ``bool`` -> ``PRAWDA``/``FAŁSZ`` (jak w XLSX),
    * daty w formacie ISO (jednoznacznym w każdym ustawieniu regionalnym),
    * przy separatorze ``;`` liczby dostają przecinek dziesiętny — tak Excel
      w polskiej wersji rozpozna je jako liczby, a nie tekst,
    * tekst zaczynający się od ``=`` ``+`` ``-`` ``@`` dostaje apostrof.  W CSV
      nie ma "typu komórki", więc bez tego Excel policzyłby taką wartość jak
      formułę (klasyczne wstrzyknięcie formuły z pliku z sieci).
    """
    value = value if isinstance(value, str) else values.sanitize_cell(value)
    if value is None:
        return ""
    if isinstance(value, bool):
        return "PRAWDA" if value else "FAŁSZ"
    if isinstance(value, (int, float)):
        text = repr(value) if isinstance(value, float) else str(value)
        return text.replace(".", ",") if decimal_comma else text
    if isinstance(value, str):
        text = _csv_text(value)
        return "'" + text if values.looks_like_formula(text) else text
    text = values.sanitize_cell(str(value))
    return "'" + text if values.looks_like_formula(text) else text


def write_csv(
    path,
    columns: Sequence[str],
    rows: Iterable[Sequence[Any]],
    *,
    sep: str = DEFAULT_CSV_SEP,
) -> int:
    """Zapisuje JEDNĄ tabelę do pliku CSV; zwraca liczbę wierszy danych.

    Kodowanie to UTF-8 **z BOM** — dzięki temu Excel otwiera plik z polskimi
    znakami po dwukliku, bez kreatora importu.  Wiersze mogą być generatorem.
    """
    target = os.path.expanduser(os.fspath(path))
    decimal_comma = sep == ";"
    width = len(columns)
    written = 0
    with open(target, "w", encoding="utf-8-sig", newline="") as handle:
        writer = csv.writer(handle, delimiter=sep, quoting=csv.QUOTE_MINIMAL,
                            lineterminator="\r\n")
        writer.writerow([_csv_cell(name, decimal_comma) for name in columns])
        for row in rows:
            cells = [_csv_cell(value, decimal_comma) for value in row]
            if len(cells) < width:
                cells.extend([""] * (width - len(cells)))
            writer.writerow(cells)
            written += 1
    return written


def _free_path(path: str) -> str:
    """Pierwsza wolna nazwa: ``a.csv``, ``a-2.csv``, ``a-3.csv``…

    Używane TYLKO przy zapisie ratunkowym — tam odmowa ("plik już istnieje")
    oznaczałaby utratę wszystkiego, co udało się zebrać.
    """
    if not os.path.exists(path):
        return path
    root, ext = os.path.splitext(path)
    for counter in range(2, 1000):
        candidate = "%s-%d%s" % (root, counter, ext)
        if not os.path.exists(candidate):
            return candidate
    return path


# --------------------------------------------------------------------------- #
# Podkomenda: build
# --------------------------------------------------------------------------- #


def _output_is_writable(path: str, overwrite: bool, reporter: Reporter) -> bool:
    """Czy wolno zapisać wynik pod tą ścieżką? (sam warunek, bez tworzenia katalogów)

    Wydzielone, bo podkomenda ``all`` musi sprawdzić to PRZED pobieraniem —
    inaczej użytkownik ściąga kilkanaście minut z portalu, żeby na koniec
    usłyszeć, że plik .xlsx już istnieje.
    """
    target = os.path.expanduser(path)
    if os.path.isdir(target):
        reporter.error(
            "%s to katalog, a potrzebna jest nazwa pliku (np. %s)"
            % (target, os.path.join(target, DEFAULT_XLSX))
        )
        return False
    if os.path.exists(target) and not overwrite:
        reporter.error(
            "Plik %s już istnieje. Dodaj opcję --overwrite, żeby go nadpisać, "
            "albo podaj inną nazwę w --out." % target
        )
        return False
    return True


def _prepare_output(path: str, overwrite: bool, reporter: Reporter,
                    suffix: str = ".xlsx") -> Optional[str]:
    """Sprawdza ścieżkę wyniku i tworzy brakujące katalogi.

    Zwraca ścieżkę albo ``None``, gdy zapisu nie wolno wykonać.
    """
    target = os.path.expanduser(path)
    if not _output_is_writable(path, overwrite, reporter):
        return None
    parent = os.path.dirname(os.path.abspath(target))
    try:
        os.makedirs(parent, exist_ok=True)
    except OSError as exc:
        reporter.error("Nie mogę utworzyć katalogu %s (%s)" % (parent, exc))
        return None
    if not target.lower().endswith(suffix):
        reporter.warn(
            "Nazwa %s nie kończy się na %s — Excel może nie skojarzyć pliku."
            % (os.path.basename(target), suffix)
        )
    return target


def _handle_duplicates(files: List[str], skip: bool, reporter: Reporter) -> List[str]:
    """Ostrzega o plikach o identycznej treści (albo je pomija przy ``--skip-duplicates``).

    Zdublowany plik to zdublowane wiersze — a więc zawyżony stan magazynu bez
    ŻADNEGO widocznego objawu.  Domyślnie tylko mówimy o tym głośno; decyzję
    zostawiamy użytkownikowi, bo to on wie, czy dwie kopie są przypadkiem.
    """
    duplicates = find_duplicate_files(files)
    if not duplicates:
        return files
    if skip:
        reporter.info(
            "Pomijam %s o treści identycznej z innymi (opcja --skip-duplicates)."
            % _files(len(duplicates))
        )
        for path in sorted(duplicates):
            reporter.info("  duplikat: %s = %s"
                          % (os.path.basename(path),
                             os.path.basename(duplicates[path])))
        return [path for path in files if path not in duplicates]
    reporter.warn(
        "%s treść identyczną z innym plikiem — te same pozycje policzą się "
        "DWA razy. Dodaj --skip-duplicates, żeby je pominąć."
        % _plural(len(duplicates), "plik ma", "pliki mają", "plików ma")
    )
    for number, path in enumerate(sorted(duplicates)):
        if number == 5:
            reporter.warn("  ... oraz %d dalszych" % (len(duplicates) - 5))
            break
        reporter.warn("  %s = %s" % (os.path.basename(path),
                                     os.path.basename(duplicates[path])))
    return files


def cmd_build(args: argparse.Namespace, reporter: Reporter) -> int:
    """Buduje jeden plik XLSX z plików XML."""
    inputs = list(args.inputs or [])
    if not inputs:
        reporter.error(
            "Nie podano danych wejściowych. Użyj --in <katalog_z_plikami_xml> "
            "(albo wskaż konkretne pliki)."
        )
        return EXIT_NO_DATA

    files = collect_xml_files(inputs, reporter)
    if args.limit and args.limit > 0:
        if len(files) > args.limit:
            reporter.info("Biorę tylko %s (opcja --limit)." % _files(args.limit))
        files = files[: args.limit]
    if not files:
        reporter.error(
            "Nie znalazłem żadnego pliku .xml w: %s\n"
            "       Sprawdź ścieżkę albo pobierz pliki podkomendą 'download'."
            % ", ".join(inputs)
        )
        return EXIT_NO_DATA

    files = _handle_duplicates(files, getattr(args, "skip_duplicates", False), reporter)

    reporter.step("Wczytuję %s XML" % _files(len(files)))
    docs, errors = load_documents(
        files,
        repeat=args.repeat,
        join_sep=args.join_sep,
        record_path=args.record_path,
        strip_ns=not args.keep_ns,
        reporter=reporter,
    )
    if not docs:
        reporter.error(
            "Żadnego pliku nie udało się wczytać (%s). "
            "Sprawdź, czy to na pewno pliki XML." % _files(len(errors))
        )
        return EXIT_NO_DATA

    unified: set = set()
    if not args.record_path and not args.no_unify:
        unified = unify_record_paths(
            docs,
            repeat=args.repeat,
            join_sep=args.join_sep,
            strip_ns=not args.keep_ns,
            reporter=reporter,
        )

    total_rows = sum(_doc_count(doc) for doc in docs)
    data_columns = xmlflatten.merge_columns(docs)
    auctions = sorted({doc.auction for doc in docs})

    reporter.step(
        "Wynik: %s, %s, %s"
        % (
            _rows(total_rows),
            _plural(len(data_columns) + len(TECH_COLUMNS), "kolumna", "kolumny", "kolumn"),
            _plural(len(auctions), "aukcja", "aukcje", "aukcji"),
        )
    )

    if args.dry_run:
        reporter.info("Tryb --dry-run: nie zapisuję pliku. Plan arkuszy:")
        reporter.info("  1. %s — %s" % (SHEET_ALL, _rows(total_rows)))
        reporter.info("  2. %s — %s" % (SHEET_SUMMARY, _rows(len(docs) + len(errors) + 1)))
        if args.per_auction:
            for number, auction in enumerate(auctions, 3):
                reporter.info("  %d. %s" % (number, auction))
        if getattr(args, "csv", None):
            reporter.info("Dodatkowo plik CSV: %s" % args.csv)
        reporter.info("Kolumny: " + ", ".join(TECH_COLUMNS + list(data_columns[:20])))
        if len(data_columns) > 20:
            reporter.info("  ... oraz %d dalszych" % (len(data_columns) - 20))
        _report_errors(errors, reporter)
        return EXIT_PARTIAL if errors else EXIT_OK

    if args.per_auction and len(auctions) > 50:
        reporter.warn(
            "Aukcji jest %d, więc powstanie tyle samo dodatkowych arkuszy — "
            "plik może otwierać się bardzo wolno. Rozważ pominięcie --per-auction."
            % len(auctions)
        )

    target = _prepare_output(args.out, args.overwrite, reporter)
    if target is None:
        return EXIT_ERROR

    # Ścieżkę CSV sprawdzamy PRZED zapisem XLSX — żeby nie okazało się po
    # wszystkim, że dodatkowego pliku i tak nie wolno zapisać.
    csv_target: Optional[str] = None
    csv_sep = getattr(args, "csv_sep", DEFAULT_CSV_SEP) or DEFAULT_CSV_SEP
    if getattr(args, "csv", None):
        csv_target = _prepare_output(args.csv, args.overwrite, reporter, ".csv")
        if csv_target is None:
            return EXIT_ERROR

    lost: List[str] = []
    sheets = build_sheets(
        docs,
        errors,
        per_auction=args.per_auction,
        use_typing=not args.no_typing,
        unified=unified,
        reporter=reporter,
        failures=lost,
    )
    reporter.step("Zapisuję %s" % os.path.abspath(target))
    try:
        xlsxwrite.write_workbook(target, sheets)
    except (OSError, MemoryError, ValueError) as exc:
        if isinstance(exc, MemoryError):
            reporter.error(
                "Za mało pamięci przy zapisie %s. Spróbuj podzielić katalog na "
                "części (--limit), pominąć --per-auction albo ograniczyć kolumny "
                "opcją --record-path." % target
            )
        else:
            reporter.error("Nie udało się zapisać %s — %s" % (target, _short(exc)))
        # Ratunek: dane są już wczytane, więc zamiast odejść z pustymi rękami
        # zapisujemy je w formacie, który poradzi sobie zawsze.
        _rescue_to_csv(docs, target, csv_sep, not args.no_typing, reporter)
        return EXIT_ERROR

    reporter.step(
        "Zapisano %s (zapis: %s)" % (os.path.abspath(target), xlsxwrite.backend_name())
    )

    if csv_target is not None:
        columns, rows = collective_table(
            docs, use_typing=not args.no_typing, reporter=reporter)
        try:
            count = write_csv(csv_target, columns, rows, sep=csv_sep)
        except (OSError, ValueError, MemoryError) as exc:
            reporter.error("Nie udało się zapisać %s — %s" % (csv_target, _short(exc)))
            return EXIT_ERROR
        reporter.step("Zapisano %s (%s, separator '%s')"
                      % (os.path.abspath(csv_target), _rows(count), csv_sep))

    _report_errors(errors, reporter)
    if lost:
        reporter.warn(
            "UWAGA na dane: %s nie dało się wczytać ponownie w chwili zapisu, "
            "więc ich wiersze NIE trafiły do arkusza (arkusz %s liczy je jako "
            "wczytane). Uruchom program ponownie na nieruszanym katalogu."
            % (_files(len(set(lost))), SHEET_SUMMARY)
        )
        return EXIT_PARTIAL
    return EXIT_PARTIAL if errors else EXIT_OK


def _rescue_to_csv(docs: Sequence[Any], target: str, sep: str,
                   use_typing: bool, reporter: Reporter) -> None:
    """Ostatnia deska ratunku: zapisuje zebrane dane do CSV, gdy XLSX padł.

    Bez tego nieudany zapis (brak miejsca, plik zablokowany przez otwarty
    Excel, błąd w generatorze OOXML) oznaczał dla użytkownika stratę CAŁEJ
    pracy — łącznie z godzinami pobierania.  CSV nie ma limitów Excela i
    powstaje z tych samych wierszy, co arkusz zbiorczy.
    """
    fallback = _free_path(os.path.splitext(target)[0] + ".csv")
    try:
        columns, rows = collective_table(docs, use_typing=use_typing, reporter=None)
        count = write_csv(fallback, columns, rows, sep=sep)
    except (OSError, ValueError, MemoryError) as exc:
        reporter.error("Nie udało się nawet ratunkowe CSV (%s)" % _short(exc))
        return
    reporter.warn(
        "Dane uratowane do %s (%s). Excel otworzy ten plik dwuklikiem; "
        "w razie potrzeby: Dane -> Z pliku tekstowego/CSV, separator '%s'."
        % (os.path.abspath(fallback), _rows(count), sep)
    )


def _report_errors(errors: Sequence[Tuple[str, str]], reporter: Reporter) -> None:
    """Wypisuje zbiorczą listę problemów na koniec pracy."""
    if not errors:
        return
    reporter.warn("Pominięto %s (szczegóły w arkuszu %s):" % (_files(len(errors)), SHEET_SUMMARY))
    for path, message in errors:
        reporter.warn("  %s — %s" % (os.path.basename(path) or path, message))


# --------------------------------------------------------------------------- #
# Podkomenda: download
# --------------------------------------------------------------------------- #


def _load_scrape():
    """Leniwy import modułu pobierania (``build`` ma działać bez sieci)."""
    from . import scrape  # import lokalny: świadomy, patrz docstring modułu

    return scrape


def _filter_auctions(auctions: Sequence[Any], pattern: Optional[str],
                     limit: Optional[int], reporter: Reporter) -> List[Any]:
    """Zawęża listę aukcji wzorcem ``--match`` i liczbą ``--limit``."""
    result = list(auctions)
    if pattern:
        regex = re.compile(pattern, re.IGNORECASE)
        result = [a for a in result if regex.search(a.id) or regex.search(a.title or "")]
        reporter.info(
            "Wzorzec --match %r zostawia %s."
            % (pattern, _plural(len(result), "aukcję", "aukcje", "aukcji"))
        )
    if limit and limit > 0 and len(result) > limit:
        reporter.info("Ograniczam do %d aukcji (opcja --limit)." % limit)
        result = result[:limit]
    return result


def cmd_download(args: argparse.Namespace, reporter: Reporter) -> int:
    """Pobiera pliki XML z portalu do katalogu."""
    scrape = _load_scrape()

    if args.auction_re:
        try:
            re.compile(args.auction_re)
        except re.error as exc:
            reporter.error("Niepoprawne wyrażenie regularne w --auction-re: %s" % exc)
            return EXIT_ERROR
    if args.match:
        try:
            re.compile(args.match)
        except re.error as exc:
            reporter.error("Niepoprawne wyrażenie regularne w --match: %s" % exc)
            return EXIT_ERROR

    out_dir = os.path.expanduser(args.out)
    if not args.dry_run:
        try:
            os.makedirs(out_dir, exist_ok=True)
        except OSError as exc:
            reporter.error("Nie mogę utworzyć katalogu %s (%s)" % (out_dir, exc))
            return EXIT_ERROR

    try:
        opener = scrape.make_opener(
            cookie=args.cookie,
            user_agent=args.user_agent or scrape.DEFAULT_USER_AGENT,
            timeout=args.timeout,
            retries=args.retries,
            delay=args.delay,
            diagnose=args.diagnose,
        )
    except scrape.ScrapeError as exc:
        # np. ciasteczko wklejone razem ze znakiem końca linii
        reporter.error(_short(exc))
        return EXIT_ERROR

    reporter.step("Szukam aukcji na %s" % args.base_url)
    try:
        auctions = scrape.discover_auctions(
            opener, args.base_url, auction_re=args.auction_re, max_pages=args.max_pages
        )
    except scrape.ScrapeError as exc:
        reporter.error("Nie udało się pobrać listy aukcji: %s" % _short(exc))
        reporter.error(
            "Podpowiedzi: sprawdź adres w --base-url, połączenie z internetem, "
            "a gdy portal wymaga zalogowania — podaj --cookie."
        )
        return EXIT_NETWORK
    except OSError as exc:  # pragma: no cover - zależne od środowiska sieciowego
        reporter.error("Błąd sieci: %s" % _short(exc))
        return EXIT_NETWORK

    if not auctions:
        reporter.error(
            "Nie znalazłem żadnej aukcji pod adresem %s.\n"
            "       Spróbuj: --diagnose (pokaże, co jest na stronie), własny wzorzec "
            "--auction-re, albo pobierz pliki ręcznie i użyj podkomendy 'build'."
            % args.base_url
        )
        return EXIT_NO_DATA

    reporter.step("Znaleziono %s" % _plural(len(auctions), "aukcję", "aukcje", "aukcji"))
    auctions = _filter_auctions(auctions, args.match, args.limit, reporter)
    if not auctions:
        reporter.error("Po zastosowaniu filtrów nie została żadna aukcja.")
        return EXIT_NO_DATA

    existing = set(os.listdir(out_dir)) if os.path.isdir(out_dir) else set()
    downloaded: List[str] = []
    skipped: List[str] = []
    found_any = False
    errors: List[Tuple[str, str]] = []

    for number, auction in enumerate(auctions, 1):
        reporter.info(
            "[%d/%d] %s — %s" % (number, len(auctions), auction.id, auction.title or "")
        )
        try:
            refs = scrape.find_xml_links(
                opener,
                auction,
                descend=args.descend,
                max_lots=args.max_lots,
                probe_limit=args.probe_limit,
            )
        except scrape.ScrapeError as exc:
            errors.append((auction.id, _short(exc)))
            reporter.warn("Pomijam aukcję %s — %s" % (auction.id, _short(exc)))
            continue
        if not refs:
            reporter.info("      (brak plików XML)")
            continue
        found_any = True
        for ref in refs:
            if args.dry_run:
                reporter.info("      %s  ->  %s" % (ref.url, ref.filename))
                continue
            try:
                path = scrape.download_xml(opener, ref, out_dir, overwrite=args.overwrite)
            except scrape.ScrapeError as exc:
                errors.append((ref.url, _short(exc)))
                reporter.warn("Nie pobrałem %s — %s" % (ref.url, _short(exc)))
                continue
            except OSError as exc:
                errors.append((ref.url, _short(exc)))
                reporter.warn("Nie zapisałem %s — %s" % (ref.filename, _short(exc)))
                continue
            name = os.path.basename(path)
            if name in existing and not args.overwrite:
                skipped.append(path)
                reporter.info("      pomijam (jest już na dysku): %s" % name)
            else:
                existing.add(name)
                downloaded.append(path)
                reporter.info("      pobrano: %s" % name)

    if args.dry_run:
        reporter.step("Tryb --dry-run: nic nie pobrano.")
        return EXIT_OK if found_any else EXIT_NO_DATA

    reporter.step(
        "Pobrano %s, pominięto %s (już były), błędów: %d. Katalog: %s"
        % (_files(len(downloaded)), _files(len(skipped)), len(errors), os.path.abspath(out_dir))
    )
    if errors:
        for what, message in errors:
            reporter.warn("  %s — %s" % (what, message))
    if not downloaded and not skipped:
        reporter.error(
            "Nie pobrano żadnego pliku XML.\n"
            "       Portal mógł zmienić układ stron albo wymaga zalogowania — "
            "zobacz sekcję o ciasteczkach w README.md, spróbuj --diagnose."
        )
        return EXIT_NO_DATA
    return EXIT_PARTIAL if errors else EXIT_OK


# --------------------------------------------------------------------------- #
# Podkomenda: all
# --------------------------------------------------------------------------- #


def cmd_all(args: argparse.Namespace, reporter: Reporter) -> int:
    """Pobiera XML-e, a potem od razu buduje z nich plik XLSX."""
    # Warunek "plik wynikowy już istnieje" znamy PRZED pierwszym żądaniem HTTP.
    # Bez tej kontroli użytkownik pobierał setki lotów (z uprzejmym --delay),
    # żeby na końcu usłyszeć "dodaj --overwrite" i zacząć od nowa.
    if not args.dry_run and not _output_is_writable(args.out, args.overwrite, reporter):
        reporter.error("Nie zaczynam pobierania — najpierw popraw --out.")
        return EXIT_ERROR
    if (not args.dry_run and getattr(args, "csv", None)
            and not _output_is_writable(args.csv, args.overwrite, reporter)):
        reporter.error("Nie zaczynam pobierania — najpierw popraw --csv.")
        return EXIT_ERROR

    download_args = argparse.Namespace(**vars(args))
    download_args.out = args.xml_dir
    code_download = cmd_download(download_args, reporter)
    if code_download == EXIT_INTERRUPTED:  # pragma: no cover - obsługa Ctrl+C wyżej
        return code_download

    build_args = argparse.Namespace(**vars(args))
    build_args.inputs = [args.xml_dir]
    build_args.limit = None  # --limit dotyczy liczby AUKCJI, nie plików

    if args.dry_run and not collect_xml_files([args.xml_dir]):
        reporter.step(
            "Tryb --dry-run: w %s nie ma jeszcze plików XML, więc nie ma z czego "
            "budować arkusza." % args.xml_dir
        )
        return code_download

    if code_download in (EXIT_NETWORK, EXIT_NO_DATA):
        if collect_xml_files([args.xml_dir]):
            reporter.warn(
                "Pobieranie się nie powiodło, ale w %s są wcześniej pobrane pliki — "
                "buduję z nich." % args.xml_dir
            )
        else:
            return code_download

    code_build = cmd_build(build_args, reporter)
    if code_build != EXIT_OK:
        return code_build
    return EXIT_PARTIAL if code_download != EXIT_OK else EXIT_OK


# --------------------------------------------------------------------------- #
# Parser argumentów
# --------------------------------------------------------------------------- #


def _polish_argparse_error(message: str) -> str:
    """Tłumaczy komunikat składniowy argparse na polski.

    README obiecuje "pomoc po polsku", a to właśnie te komunikaty widzi
    użytkownik w chwili pomyłki — czyli dokładnie wtedy, gdy pomoc jest
    najbardziej potrzebna.  Nieznanych wzorców nie kaleczymy: wracają bez zmian.
    """
    match = re.match(r"^argument (?P<what>.+?): invalid choice: (?P<value>.+?) "
                     r"\(choose from (?P<opts>.+)\)$", message)
    if match:
        opts = match.group("opts").replace("'", "")
        if match.group("what").strip().lower() == "podkomenda":
            return ("nieznana podkomenda: %s — dostępne: %s"
                    % (match.group("value"), opts))
        return ("nieznana wartość %s dla %s — dostępne: %s"
                % (match.group("value"), match.group("what"), opts))
    match = re.match(r"^unrecognized arguments: (?P<rest>.+)$", message)
    if match:
        return ("nierozpoznany argument: %s — sprawdź pisownię "
                "(pełna lista opcji: --help)" % match.group("rest"))
    match = re.match(r"^the following arguments are required: (?P<rest>.+)$", message)
    if match:
        return "brakuje wymaganego argumentu: %s" % match.group("rest")
    match = re.match(r"^argument (?P<what>.+?): expected one argument$", message)
    if match:
        return "opcja %s wymaga podania wartości" % match.group("what")
    match = re.match(r"^argument (?P<what>.+?): expected at least one argument$", message)
    if match:
        return "opcja %s wymaga co najmniej jednej wartości" % match.group("what")
    match = re.match(r"^argument (?P<what>.+?): invalid (?P<kind>\w+) value: "
                     r"(?P<value>.+)$", message)
    if match:
        kinds = {"int": "liczbą całkowitą", "float": "liczbą"}
        opis = kinds.get(match.group("kind"), "wartością typu " + match.group("kind"))
        return ("wartość %s dla %s nie jest %s"
                % (match.group("value"), match.group("what"), opis))
    match = re.match(r"^ambiguous option: (?P<rest>.+)$", message)
    if match:
        return "niejednoznaczny skrót opcji: %s" % match.group("rest")
    if message == "too few arguments":  # pragma: no cover - starsze Pythony
        return "za mało argumentów"
    return message


def _polish_headers(text: str) -> str:
    """Podmienia angielskie resztki szkieletu argparse na polskie."""
    return (text
            .replace("usage: ", "użycie: ")
            .replace("positional arguments:", "argumenty pozycyjne:")
            .replace("options:", "opcje:")
            .replace("optional arguments:", "opcje:"))


class PolishArgumentParser(argparse.ArgumentParser):
    """``ArgumentParser`` mówiący po polsku — także przy błędach składni.

    Sam argparse ma szkielet po angielsku ("positional arguments", "options",
    "invalid choice", "unrecognized arguments") i nie da się go przetłumaczyć
    parametrem.  Podmieniamy więc tytuły grup, własną opcję ``--help`` oraz
    :meth:`error`.  Podparsery dziedziczą tę klasę automatycznie
    (``add_subparsers`` bierze ``parser_class = type(self)``).
    """

    def __init__(self, *args, **kwargs):
        with_help = kwargs.pop("add_help", True)
        kwargs["add_help"] = False
        super().__init__(*args, **kwargs)
        self._positionals.title = "argumenty pozycyjne"
        self._optionals.title = "opcje"
        if with_help:
            self.add_argument("-h", "--help", action="help",
                              help="pokaż tę pomoc i zakończ")

    def format_usage(self) -> str:  # noqa: D102
        return _polish_headers(super().format_usage())

    def format_help(self) -> str:  # noqa: D102
        return _polish_headers(super().format_help())

    def error(self, message: str):  # noqa: D102
        self.print_usage(sys.stderr)
        self.exit(EXIT_USAGE,
                  "%s: błąd: %s\n" % (self.prog, _polish_argparse_error(message)))


def _add_common(parser: argparse.ArgumentParser) -> None:
    """Opcje wspólne dla wszystkich podkomend."""
    parser.add_argument("--version", action="version",
                        version="%s %s" % (PROG, _version()),
                        help="pokaż numer wersji i zakończ")
    parser.add_argument("--quiet", "-q", action="store_true",
                        help="mniej komunikatów (błędy nadal widoczne)")
    parser.add_argument("--dry-run", action="store_true",
                        help="pokaż, co by się stało, ale nic nie zapisuj")
    parser.add_argument("--overwrite", action="store_true",
                        help="nadpisuj istniejące pliki (XML-e i plik .xlsx)")


def _add_download_options(parser: argparse.ArgumentParser) -> None:
    """Opcje pobierania z portalu."""
    group = parser.add_argument_group("pobieranie")
    group.add_argument("--base-url", default=DEFAULT_BASE_URL,
                       metavar="ADRES",
                       help="adres listy aukcji (domyślnie %(default)s)")
    group.add_argument("--cookie", default=None, metavar="CIASTECZKO",
                       help="wartość nagłówka Cookie skopiowana z przeglądarki "
                            "(gdy portal wymaga zalogowania)")
    group.add_argument("--auction-re", default=None, metavar="REGEX",
                       help="własne wyrażenie regularne rozpoznające adresy aukcji, "
                            "np. '/auction/(?P<id>[^/?#]+)'")
    group.add_argument("--match", default=None, metavar="REGEX",
                       help="pobierz tylko aukcje, których identyfikator lub tytuł "
                            "pasuje do wzorca (np. '2026')")
    group.add_argument("--limit", type=int, default=None, metavar="N",
                       help="pobierz najwyżej N aukcji (przydatne na próbę)")
    group.add_argument("--max-pages", type=int, default=50, metavar="N",
                       help="ile stron listy aukcji przejrzeć (domyślnie %(default)s)")
    group.add_argument("--max-lots", type=int, default=200, metavar="N",
                       help="ile lotów w jednej aukcji odwiedzić (domyślnie %(default)s)")
    group.add_argument("--probe-limit", type=int, default=25, metavar="N",
                       help="ile niepewnych linków sprawdzić po Content-Type "
                            "(domyślnie %(default)s)")
    group.add_argument("--descend", choices=("auto", "always", "never"), default="auto",
                       help="czy schodzić ze strony aukcji na strony lotów "
                            "(domyślnie %(default)s)")
    group.add_argument("--delay", type=float, default=0.5, metavar="SEKUNDY",
                       help="uprzejma przerwa między żądaniami (domyślnie %(default)s)")
    group.add_argument("--timeout", type=float, default=30.0, metavar="SEKUNDY",
                       help="limit czasu jednej operacji na gnieździe; cała "
                            "odpowiedź ma na siebie dziesięciokrotność tego "
                            "czasu (domyślnie %(default)s)")
    group.add_argument("--retries", type=int, default=4, metavar="N",
                       help="ile razy ponowić nieudane żądanie (domyślnie %(default)s)")
    group.add_argument("--user-agent", default=None, metavar="TEKST",
                       help="własny nagłówek User-Agent")
    group.add_argument("--diagnose", action="store_true",
                       help="gdy nic nie znaleziono, wypisz, co jest na stronie")


def _add_build_options(parser: argparse.ArgumentParser) -> None:
    """Opcje budowania arkusza."""
    group = parser.add_argument_group("budowanie XLSX")
    group.add_argument("--per-auction", action="store_true",
                       help="dodaj osobny arkusz dla każdej aukcji")
    group.add_argument("--repeat", choices=("join", "index"), default="join",
                       help="powtarzające się pola w rekordzie: sklej w jedną komórkę "
                            "(join) albo rozbij na pole[1], pole[2] (index); "
                            "domyślnie %(default)s")
    group.add_argument("--join-sep", default=" | ", metavar="TEKST",
                       help="czym sklejać przy --repeat join (domyślnie ' | ')")
    group.add_argument("--record-path", default=None, metavar="SCIEZKA",
                       help="wymuś ścieżkę elementu powtarzalnego, np. 'batch/lot/items/item' "
                            "albo sam 'item'; domyślnie wykrywana automatycznie")
    group.add_argument("--keep-ns", action="store_true",
                       help="zostaw przestrzenie nazw XML w nazwach kolumn")
    group.add_argument("--no-unify", action="store_true",
                       help="nie ujednolicaj ścieżki rekordu między plikami "
                            "(domyślnie plik z jedną pozycją dostaje takie same "
                            "kolumny jak pliki z wieloma)")
    group.add_argument("--no-typing", action="store_true",
                       help="nie zamieniaj tekstu na liczby/daty — wszystko jako tekst")
    group.add_argument("--skip-duplicates", action="store_true",
                       help="pomiń pliki o treści identycznej z innym plikiem "
                            "(ten sam pakiet pobrany dwa razy); domyślnie tylko "
                            "ostrzeżenie")
    group.add_argument("--csv", default=None, metavar="PLIK",
                       help="zapisz DODATKOWO arkusz zbiorczy do pliku CSV "
                            "(UTF-8 z BOM; przydatne, gdy .xlsx nie chce się "
                            "otworzyć albo dane idą do innego programu)")
    group.add_argument("--csv-sep", default=DEFAULT_CSV_SEP, metavar="ZNAK",
                       help="separator pól w CSV (domyślnie '%(default)s' — tak "
                            "czyta polski Excel; użyj ',' dla wersji angielskiej)")


def build_parser() -> argparse.ArgumentParser:
    """Buduje parser argumentów wiersza poleceń."""
    parser = PolishArgumentParser(
        prog=PROG,
        description="Pobiera pliki XML z portalu flexitauctions.com i scala je "
                    "w JEDEN plik .xlsx.",
        epilog="Przykłady:\n"
               "  python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx\n"
               "  python3 -m flexit2xlsx download --out xml_flexit --limit 3\n"
               "  python3 -m flexit2xlsx all --out aukcje.xlsx --per-auction\n",
        formatter_class=argparse.RawDescriptionHelpFormatter,
    )
    parser.add_argument("--version", action="version",
                        version="%s %s" % (PROG, _version()),
                        help="pokaż numer wersji i zakończ")
    subparsers = parser.add_subparsers(dest="command", metavar="PODKOMENDA")

    # --- download ---------------------------------------------------------
    download = subparsers.add_parser(
        "download",
        help="pobierz pliki XML z portalu do katalogu",
        description="Pobiera pliki XML 'zawartość pakietu' z portalu do katalogu.",
    )
    download.add_argument("--out", "-o", default=DEFAULT_XML_DIR, metavar="KATALOG",
                          help="katalog na pobrane pliki XML (domyślnie %(default)s)")
    _add_common(download)
    _add_download_options(download)
    download.set_defaults(func=cmd_download)

    # --- build ------------------------------------------------------------
    build = subparsers.add_parser(
        "build",
        help="zbuduj jeden plik .xlsx z plików XML",
        description="Scala pliki XML (dowolna struktura) w jeden plik .xlsx.",
    )
    build.add_argument("--in", "-i", dest="in_paths", action="append", nargs="+",
                       metavar="SCIEZKA",
                       help="katalog z plikami XML albo pojedyncze pliki "
                            "(opcję można powtórzyć)")
    build.add_argument("paths", nargs="*", metavar="PLIK",
                       help="dodatkowe pliki/katalogi podane bez --in")
    build.add_argument("--out", "-o", default=DEFAULT_XLSX, metavar="PLIK",
                       help="plik wynikowy .xlsx (domyślnie %(default)s)")
    build.add_argument("--limit", type=int, default=None, metavar="N",
                       help="weź najwyżej N plików XML")
    _add_common(build)
    _add_build_options(build)
    build.set_defaults(func=cmd_build)

    # --- all --------------------------------------------------------------
    every = subparsers.add_parser(
        "all",
        help="pobierz z portalu i od razu zbuduj .xlsx",
        description="Pobiera pliki XML z portalu, a następnie scala je w jeden .xlsx.",
    )
    every.add_argument("--out", "-o", default=DEFAULT_XLSX, metavar="PLIK",
                       help="plik wynikowy .xlsx (domyślnie %(default)s)")
    every.add_argument("--in", "-i", "--xml-dir", dest="xml_dir",
                       default=DEFAULT_XML_DIR, metavar="KATALOG",
                       help="katalog na pobrane pliki XML (domyślnie %(default)s)")
    _add_common(every)
    _add_download_options(every)
    _add_build_options(every)
    every.set_defaults(func=cmd_all)

    return parser


def _version() -> str:
    """Numer wersji pakietu (bez importu cyklicznego)."""
    from . import __version__

    return __version__


def _normalize_inputs(args: argparse.Namespace) -> None:
    """Zwija ``--in`` (można powtarzać) i argumenty pozycyjne w jedną listę."""
    inputs: List[str] = []
    for chunk in getattr(args, "in_paths", None) or []:
        inputs.extend(chunk)
    inputs.extend(getattr(args, "paths", None) or [])
    args.inputs = inputs


def main(argv: Optional[Sequence[str]] = None) -> int:
    """Punkt wejścia CLI.  Zwraca kod wyjścia (patrz docstring modułu)."""
    parser = build_parser()
    args = parser.parse_args(list(argv) if argv is not None else None)

    if not getattr(args, "command", None):
        parser.print_help()
        return EXIT_USAGE

    _normalize_inputs(args)
    reporter = Reporter(quiet=getattr(args, "quiet", False))

    try:
        return int(args.func(args, reporter))
    except KeyboardInterrupt:
        reporter.error("Przerwano przez użytkownika (Ctrl+C).")
        return EXIT_INTERRUPTED
    except BrokenPipeError:  # pragma: no cover - np. `| head`
        return EXIT_OK
    except MemoryError:
        # UWAGA: przy braku pamięci nie składamy długich napisów
        reporter.error(
            "Za mało pamięci. Podziel katalog na części, użyj --limit, "
            "pomiń --per-auction albo wskaż --record-path."
        )
        return EXIT_ERROR
    except Exception as exc:  # ostatnia siatka bezpieczeństwa
        # Użytkownik tego narzędzia nie ma czytać tracebacku — dostaje zdanie
        # po polsku i zdefiniowany kod wyjścia.
        reporter.error("Nieoczekiwany błąd: %s: %s" % (type(exc).__name__, _short(exc)))
        reporter.error(
            "Jeśli błąd się powtarza, uruchom polecenie ponownie z opcją --diagnose "
            "i zachowaj wypisany raport."
        )
        return EXIT_ERROR


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
