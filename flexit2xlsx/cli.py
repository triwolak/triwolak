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
* ``build`` NIE nadpisuje istniejącego pliku ``.xlsx`` bez ``--overwrite``
  (łatwo pomylić katalogi; zniszczony wynik pracy boli bardziej niż komunikat).
* Błąd pojedynczego pliku XML nigdy nie przerywa całości — trafia do arkusza
  ``Podsumowanie``, na ``stderr`` i do kodu wyjścia 3.
* Kolejność plików jest deterministyczna (sortowanie po ścieżce), żeby dwa
  uruchomienia dawały identyczny wynik.
"""

from __future__ import annotations

import argparse
import os
import re
import sys
from typing import Any, Dict, Iterator, List, Optional, Sequence, Tuple

from . import values, xlsxwrite, xmlflatten
from .xlsxwrite import Sheet

__all__ = [
    "main",
    "build_parser",
    "collect_xml_files",
    "load_documents",
    "build_sheets",
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
# Wejście: zbieranie plików XML
# --------------------------------------------------------------------------- #


def collect_xml_files(paths: Sequence[str], reporter: Optional[Reporter] = None) -> List[str]:
    """Zamienia listę ścieżek (katalogi i/lub pliki) na posortowaną listę plików XML.

    * katalog — przeszukiwany rekurencyjnie, brane pliki ``*.xml``
      (z pominięciem katalogów ukrytych i plików ``.part`` po przerwanym pobieraniu),
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
            if not batch:
                reporter.warn("Katalog %s nie zawiera plików .xml" % path)
            for item in batch:
                remember(item)
        elif os.path.isfile(path):
            remember(path)
        else:
            reporter.warn("Pomijam %s — nie ma takiego pliku ani katalogu" % path)
    return found


# --------------------------------------------------------------------------- #
# Wczytywanie dokumentów
# --------------------------------------------------------------------------- #


def load_documents(
    files: Sequence[str],
    *,
    repeat: str = "join",
    join_sep: str = " | ",
    record_path: Optional[str] = None,
    strip_ns: bool = True,
    reporter: Optional[Reporter] = None,
) -> Tuple[List[xmlflatten.ParsedDoc], List[Tuple[str, str]]]:
    """Parsuje pliki XML; błąd JEDNEGO pliku nie przerywa całości.

    Zwraca ``(dokumenty, błędy)``, gdzie błąd to para ``(ścieżka, komunikat)``.
    """
    reporter = reporter or Reporter(quiet=True)
    docs: List[xmlflatten.ParsedDoc] = []
    errors: List[Tuple[str, str]] = []
    for path in files:
        try:
            doc = xmlflatten.parse_file(
                path,
                repeat=repeat,
                join_sep=join_sep,
                record_path=record_path,
                strip_ns=strip_ns,
            )
        except ValueError as exc:  # XmlParseError też jest ValueError
            message = _short(exc)
            errors.append((path, message))
            reporter.warn("Pomijam %s — %s" % (os.path.basename(path), message))
            continue
        except OSError as exc:
            message = _short("nie udało się odczytać pliku: %s" % exc)
            errors.append((path, message))
            reporter.warn("Pomijam %s — %s" % (os.path.basename(path), message))
            continue
        docs.append(doc)
        reporter.info(
            "  %-40s aukcja=%s, %s"
            % (
                _short(os.path.basename(path), 40),
                doc.auction,
                _plural(len(doc.records), "pozycja", "pozycje", "pozycji"),
            )
        )
    return docs, errors


def unify_record_paths(
    docs: List[xmlflatten.ParsedDoc],
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
        for candidate in candidates:
            try:
                fresh = xmlflatten.parse_file(
                    doc.source,
                    repeat=repeat,
                    join_sep=join_sep,
                    record_path=candidate,
                    strip_ns=strip_ns,
                )
            except (ValueError, OSError):
                continue
            if fresh.record_path and fresh.records:
                docs[index] = fresh
                unified.add(fresh.source)
                reporter.info(
                    "  ujednolicam kolumny: %s -> ścieżka rekordu %s"
                    % (os.path.basename(doc.source), candidate)
                )
                break
    return unified


# --------------------------------------------------------------------------- #
# Budowanie arkuszy
# --------------------------------------------------------------------------- #


def _convert(value: Any, use_typing: bool) -> Any:
    """Zamienia napis na liczbę/datę/bool, gdy włączone jest rozpoznawanie typów."""
    if use_typing and isinstance(value, str):
        return values.coerce_value(value)
    return value


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
    docs: Sequence[xmlflatten.ParsedDoc],
    columns: Sequence[str],
    *,
    with_auction: bool,
    use_typing: bool,
) -> Iterator[List[Any]]:
    """Generuje wiersze arkusza: kolumny techniczne + wartości z XML-a.

    Świadomie jest to GENERATOR — arkusz może mieć setki tysięcy wierszy,
    a ``xlsxwrite`` zapisuje strumieniowo.
    """
    for doc in docs:
        filename = os.path.basename(doc.source) or doc.source
        for index, raw in enumerate(xmlflatten.rows_for(doc, columns), 1):
            cells = [_convert(value, use_typing) for value in raw]
            if with_auction:
                yield [doc.auction, filename, index] + cells
            else:
                yield [filename, index] + cells


def _summary_rows(
    docs: Sequence[xmlflatten.ParsedDoc],
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
                len(doc.records),
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
    total = sum(len(doc.records) for doc in docs)
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


def build_sheets(
    docs: Sequence[xmlflatten.ParsedDoc],
    errors: Sequence[Tuple[str, str]] = (),
    *,
    per_auction: bool = False,
    use_typing: bool = True,
    unified: Sequence[str] = (),
) -> List[Sheet]:
    """Składa listę arkuszy do zapisania przez :func:`xlsxwrite.write_workbook`."""
    data_columns = xmlflatten.merge_columns(list(docs))
    headers = _unique_headers(data_columns, TECH_COLUMNS)

    used_names: set = set()
    sheets: List[Sheet] = [
        Sheet(
            name=xlsxwrite.safe_sheet_name(SHEET_ALL, used_names),
            columns=TECH_COLUMNS + headers,
            rows=_iter_rows(docs, data_columns, with_auction=True, use_typing=use_typing),
        ),
        Sheet(
            name=xlsxwrite.safe_sheet_name(SHEET_SUMMARY, used_names),
            columns=list(SUMMARY_COLUMNS),
            rows=_summary_rows(docs, errors, len(data_columns), unified),
        ),
    ]

    if per_auction:
        groups: Dict[str, List[xmlflatten.ParsedDoc]] = {}
        for doc in docs:
            groups.setdefault(doc.auction, []).append(doc)
        for auction, group in groups.items():
            own_columns = xmlflatten.merge_columns(group)
            own_headers = _unique_headers(own_columns, TECH_COLUMNS_PER_AUCTION)
            sheets.append(
                Sheet(
                    name=xlsxwrite.safe_sheet_name(auction, used_names),
                    columns=TECH_COLUMNS_PER_AUCTION + own_headers,
                    rows=_iter_rows(
                        group, own_columns, with_auction=False, use_typing=use_typing
                    ),
                )
            )
    return sheets


# --------------------------------------------------------------------------- #
# Podkomenda: build
# --------------------------------------------------------------------------- #


def _prepare_output(path: str, overwrite: bool, reporter: Reporter) -> Optional[str]:
    """Sprawdza ścieżkę wyniku i tworzy brakujące katalogi.

    Zwraca ścieżkę albo ``None``, gdy zapisu nie wolno wykonać.
    """
    target = os.path.expanduser(path)
    if os.path.isdir(target):
        reporter.error(
            "%s to katalog, a potrzebna jest nazwa pliku (np. %s)"
            % (target, os.path.join(target, DEFAULT_XLSX))
        )
        return None
    if os.path.exists(target) and not overwrite:
        reporter.error(
            "Plik %s już istnieje. Dodaj opcję --overwrite, żeby go nadpisać, "
            "albo podaj inną nazwę w --out." % target
        )
        return None
    parent = os.path.dirname(os.path.abspath(target))
    try:
        os.makedirs(parent, exist_ok=True)
    except OSError as exc:
        reporter.error("Nie mogę utworzyć katalogu %s (%s)" % (parent, exc))
        return None
    if not target.lower().endswith(".xlsx"):
        reporter.warn(
            "Nazwa %s nie kończy się na .xlsx — Excel może nie skojarzyć pliku."
            % os.path.basename(target)
        )
    return target


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

    total_rows = sum(len(doc.records) for doc in docs)
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

    sheets = build_sheets(
        docs,
        errors,
        per_auction=args.per_auction,
        use_typing=not args.no_typing,
        unified=unified,
    )
    try:
        xlsxwrite.write_workbook(target, sheets)
    except OSError as exc:
        reporter.error("Nie udało się zapisać %s (%s)" % (target, exc))
        return EXIT_ERROR

    reporter.step(
        "Zapisano %s (zapis: %s)" % (os.path.abspath(target), xlsxwrite.backend_name())
    )
    _report_errors(errors, reporter)
    return EXIT_PARTIAL if errors else EXIT_OK


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

    opener = scrape.make_opener(
        cookie=args.cookie,
        timeout=args.timeout,
        retries=args.retries,
        delay=args.delay,
        diagnose=args.diagnose,
    )
    if args.user_agent:
        opener.user_agent = args.user_agent

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


def _add_common(parser: argparse.ArgumentParser) -> None:
    """Opcje wspólne dla wszystkich podkomend."""
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
                       help="limit czasu jednego żądania (domyślnie %(default)s)")
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


def build_parser() -> argparse.ArgumentParser:
    """Buduje parser argumentów wiersza poleceń."""
    parser = argparse.ArgumentParser(
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
                        version="%(prog)s " + _version())
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


if __name__ == "__main__":  # pragma: no cover
    sys.exit(main())
