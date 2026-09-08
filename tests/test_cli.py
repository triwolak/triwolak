# -*- coding: utf-8 -*-
"""Testy interfejsu wiersza poleceń (``flexit2xlsx.cli``).

Główny scenariusz jest end-to-end: kilka plików XML o RÓŻNYCH schematach
w katalogu tymczasowym -> jeden plik ``.xlsx`` -> odczyt przez ``openpyxl``
-> sprawdzenie, że wiersze ze wszystkich aukcji są obecne, kolumny są
zunifikowane, a brakujące pola są puste.

Weryfikacja wyniku idzie NIEZALEŻNĄ drogą (``openpyxl``), a nie przez ten sam
kod, który plik zapisał.  ``openpyxl`` jest w tym środowisku zainstalowane;
gdyby go zabrakło, testy odczytujące XLSX są pomijane, a reszta działa dalej.

Test pobierania (``download``) korzysta z atrapy portalu przygotowanej przez
autora ``scrape.py`` — klasy ``MockSite`` z ``tests/test_scrape.py``
(serwer ``http.server`` na losowym porcie, mapa tras w ``tests/mocksite``).
"""

from __future__ import annotations

import contextlib
import io
import os
import shutil
import subprocess
import sys
import tempfile
import unittest

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(TESTS_DIR)
for _path in (ROOT_DIR, TESTS_DIR):
    if _path not in sys.path:
        sys.path.insert(0, _path)

from flexit2xlsx import cli, xmlflatten  # noqa: E402
from flexit2xlsx import xlsxwrite  # noqa: E402

try:
    import openpyxl
except ImportError:  # pragma: no cover - openpyxl jest opcjonalne
    openpyxl = None

try:  # atrapa portalu udostępniona przez testy scrape.py
    from test_scrape import MockSite
except Exception:  # pragma: no cover - brak atrapy = pomijamy testy sieciowe
    MockSite = None


# ---------------------------------------------------------------------------
# Fixtury XML — pięć RÓŻNYCH schematów
# ---------------------------------------------------------------------------

#: 1. Polski schemat "aukcja -> loty", metadane sprzedawcy, liczby po polsku.
XML_LOTY = """<?xml version="1.0" encoding="UTF-8"?>
<aukcja numer="AUK-2026-01">
  <sprzedawca>
    <nazwa>Flexit sp. z o.o.</nazwa>
    <miasto>Łódź</miasto>
  </sprzedawca>
  <data>2026-02-26</data>
  <lots>
    <lot>
      <nazwa>Zestaw laptopów</nazwa>
      <cena>1 234,56</cena>
      <sztuk>12</sztuk>
      <opis><![CDATA[Stan: dobry & sprawdzony]]></opis>
    </lot>
    <lot>
      <nazwa>Monitory 24"</nazwa>
      <cena>499,00</cena>
      <sztuk>8</sztuk>
      <opis>Zażółć gęślą jaźń</opis>
    </lot>
    <lot>
      <nazwa>Stacje dokujące</nazwa>
      <cena>99,90</cena>
      <sztuk>30</sztuk>
    </lot>
  </lots>
</aukcja>
"""

#: 2. Angielski schemat "package -> items", atrybuty na rekordzie, inne pola.
XML_PAKIET = """<?xml version="1.0" encoding="UTF-8"?>
<package>
  <meta>
    <auction>BATCH-77</auction>
    <currency>EUR</currency>
  </meta>
  <items>
    <item sku="SKU-1" grade="A">
      <model>ThinkPad T480</model>
      <serial>PF1A2B3C</serial>
      <ram>16 GB</ram>
    </item>
    <item sku="SKU-2" grade="B">
      <model>OptiPlex 3060</model>
      <serial>DL0002</serial>
      <ram>8 GB</ram>
    </item>
  </items>
</package>
"""

#: 3. Płaska tabela bez metadanych — identyfikator aukcji z nazwy pliku.
XML_PLASKI = """<?xml version="1.0" encoding="UTF-8"?>
<root>
  <row><kod>A1</kod><ilosc>5</ilosc><aktywny>true</aktywny></row>
  <row><kod>A2</kod><ilosc>7</ilosc><aktywny>false</aktywny></row>
  <row><kod>A3</kod><ilosc>0</ilosc><aktywny>true</aktywny></row>
  <row><kod>A4</kod><ilosc>12</ilosc><aktywny>true</aktywny></row>
</root>
"""

#: 4. Przestrzenie nazw, kodowanie ISO-8859-2, wartość groźna dla Excela.
XML_NS = """<?xml version="1.0" encoding="ISO-8859-2"?>
<f:faktura xmlns:f="http://example.org/faktura" numer="FV-9">
  <f:pozycje>
    <f:pozycja>
      <f:towar>Zażółć gęślą jaźń</f:towar>
      <f:formula>=SUMA(A1:A9)</f:formula>
      <f:data>15.03.2024</f:data>
    </f:pozycja>
    <f:pozycja>
      <f:towar>Wąż strażacki</f:towar>
      <f:formula>-2+3</f:formula>
      <f:data>2024-04-01</f:data>
    </f:pozycja>
  </f:pozycje>
</f:faktura>
"""

#: 5. Dokument bez powtórzeń — jeden wiersz na cały plik.
XML_JEDEN = """<?xml version="1.0" encoding="UTF-8"?>
<raport>
  <auction_id>RAP-5</auction_id>
  <tytul>Raport zbiorczy</tytul>
  <suma>10 000,00</suma>
</raport>
"""

#: 6. Uszkodzony XML — sprawdza, że jeden zły plik nie wywraca całości.
XML_ZEPSUTY = "<a><b>bez zamkniecia</a>"

#: Wszystkie poprawne fixtury: nazwa pliku -> treść.
FIXTURES = {
    "aukcja_loty.xml": XML_LOTY,
    "pakiet_items.xml": XML_PAKIET,
    "plaski.xml": XML_PLASKI,
    "faktura_ns.xml": XML_NS,
    "raport_jeden.xml": XML_JEDEN,
}


def write_fixtures(directory: str, names=None) -> None:
    """Zapisuje fixtury do katalogu (w kodowaniu wskazanym w deklaracji XML)."""
    for name, text in FIXTURES.items():
        if names is not None and name not in names:
            continue
        encoding = "iso-8859-2" if "ISO-8859-2" in text.split("\n", 1)[0] else "utf-8"
        with open(os.path.join(directory, name), "wb") as handle:
            handle.write(text.encode(encoding))


# ---------------------------------------------------------------------------
# Narzędzia
# ---------------------------------------------------------------------------


@contextlib.contextmanager
def run_cli(argv):
    """Uruchamia CLI, przechwytując oba strumienie.

    Zwraca obiekt z polami ``code``, ``out``, ``err``.
    """
    out, err = io.StringIO(), io.StringIO()

    class Result:
        code = None

    result = Result()
    with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
        result.code = cli.main(argv)
    result.out = out.getvalue()
    result.err = err.getvalue()
    yield result


def call_cli(argv):
    """Skrót: uruchom CLI i zwróć wynik (bez menedżera kontekstu)."""
    with run_cli(argv) as result:
        pass
    return result


def sheet_rows(path, sheet_name):
    """Zwraca listę wierszy arkusza jako krotki wartości (razem z nagłówkiem)."""
    workbook = openpyxl.load_workbook(path)
    try:
        worksheet = workbook[sheet_name]
        return [tuple(row) for row in worksheet.iter_rows(values_only=True)]
    finally:
        workbook.close()


def as_dicts(rows):
    """Zamienia (nagłówek + wiersze) na listę słowników kolumna -> wartość."""
    header = rows[0]
    return [dict(zip(header, row)) for row in rows[1:]]


class CliTestCase(unittest.TestCase):
    """Baza: katalog tymczasowy sprzątany po teście."""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="flexit-cli-")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.xml_dir = os.path.join(self.tmp, "xml")
        os.makedirs(self.xml_dir)
        self.out = os.path.join(self.tmp, "wynik.xlsx")


# ---------------------------------------------------------------------------
# 1. Zbieranie plików wejściowych
# ---------------------------------------------------------------------------


class TestCollectXmlFiles(CliTestCase):
    """``collect_xml_files``: katalogi, pliki, śmieci, powtórzenia."""

    def test_directory_is_searched_recursively_and_sorted(self):
        write_fixtures(self.xml_dir)
        nested = os.path.join(self.xml_dir, "podkatalog")
        os.makedirs(nested)
        with open(os.path.join(nested, "zagniezdzony.xml"), "w", encoding="utf-8") as fh:
            fh.write(XML_PLASKI)
        found = cli.collect_xml_files([self.xml_dir])
        self.assertEqual(len(found), len(FIXTURES) + 1)
        self.assertIn(os.path.join(nested, "zagniezdzony.xml"), found)
        # kolejność jest deterministyczna: dwa przebiegi dają to samo
        self.assertEqual(found, cli.collect_xml_files([self.xml_dir]))
        # pliki z jednego katalogu są posortowane alfabetycznie
        top = [os.path.basename(p) for p in found if os.path.dirname(p) == self.xml_dir]
        self.assertEqual(top, sorted(top))

    def test_non_xml_files_are_ignored_in_directory(self):
        write_fixtures(self.xml_dir)
        with open(os.path.join(self.xml_dir, "notatka.txt"), "w", encoding="utf-8") as fh:
            fh.write("nie xml")
        with open(os.path.join(self.xml_dir, "batch.xml.part"), "w", encoding="utf-8") as fh:
            fh.write("<a/>")
        found = [os.path.basename(p) for p in cli.collect_xml_files([self.xml_dir])]
        self.assertEqual(sorted(found), sorted(FIXTURES))

    def test_explicit_file_is_taken_even_without_xml_extension(self):
        path = os.path.join(self.tmp, "dane.txt")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(XML_PLASKI)
        self.assertEqual(cli.collect_xml_files([path]), [path])

    def test_duplicates_are_removed(self):
        write_fixtures(self.xml_dir, {"plaski.xml"})
        path = os.path.join(self.xml_dir, "plaski.xml")
        self.assertEqual(cli.collect_xml_files([self.xml_dir, path]), [path])

    def test_missing_path_warns_but_does_not_raise(self):
        buffer = io.StringIO()
        reporter = cli.Reporter(err=buffer)
        self.assertEqual(cli.collect_xml_files(["/nie/ma/takiego"], reporter), [])
        self.assertIn("nie ma takiego pliku", buffer.getvalue())


# ---------------------------------------------------------------------------
# 2. Główny scenariusz end-to-end
# ---------------------------------------------------------------------------


@unittest.skipIf(openpyxl is None, "openpyxl potrzebny do weryfikacji wyniku")
class TestBuildEndToEnd(CliTestCase):
    """Pięć różnych schematów -> jeden XLSX -> weryfikacja przez openpyxl."""

    def setUp(self):
        super().setUp()
        write_fixtures(self.xml_dir)
        self.result = call_cli(["build", "--in", self.xml_dir, "--out", self.out])

    def test_exit_code_and_file_created(self):
        self.assertEqual(self.result.code, cli.EXIT_OK, self.result.err)
        self.assertTrue(os.path.isfile(self.out))

    def test_expected_sheets_exist(self):
        workbook = openpyxl.load_workbook(self.out)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook.sheetnames, [cli.SHEET_ALL, cli.SHEET_SUMMARY])

    def test_technical_columns_are_first(self):
        rows = sheet_rows(self.out, cli.SHEET_ALL)
        self.assertEqual(list(rows[0][:3]), cli.TECH_COLUMNS)

    def test_all_records_from_all_files_are_present(self):
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        # 3 loty + 2 pozycje + 4 wiersze + 2 pozycje faktury + 1 raport
        self.assertEqual(len(rows), 12)
        per_file = {}
        for row in rows:
            per_file[row["Plik"]] = per_file.get(row["Plik"], 0) + 1
        self.assertEqual(
            per_file,
            {
                "aukcja_loty.xml": 3,
                "pakiet_items.xml": 2,
                "plaski.xml": 4,
                "faktura_ns.xml": 2,
                "raport_jeden.xml": 1,
            },
        )

    def test_row_numbers_restart_for_each_file(self):
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        numbers = {}
        for row in rows:
            numbers.setdefault(row["Plik"], []).append(row["Nr pozycji"])
        self.assertEqual(numbers["aukcja_loty.xml"], [1, 2, 3])
        self.assertEqual(numbers["plaski.xml"], [1, 2, 3, 4])

    def test_columns_are_unified_and_missing_fields_are_empty(self):
        rows = sheet_rows(self.out, cli.SHEET_ALL)
        header = rows[0]
        # jeden nagłówek dla całości, bez powtórzeń nazw
        self.assertEqual(len(header), len(set(header)))
        data = as_dicts(rows)
        # każdy wiersz ma dokładnie tyle komórek, ile jest kolumn
        for row in rows[1:]:
            self.assertEqual(len(row), len(header))
        lot = next(r for r in data if r["Plik"] == "aukcja_loty.xml")
        item = next(r for r in data if r["Plik"] == "pakiet_items.xml")
        self.assertEqual(lot["nazwa"], "Zestaw laptopów")
        self.assertEqual(item["model"], "ThinkPad T480")
        # pola z obcego schematu są PUSTE, a nie "None"/""
        self.assertIsNone(lot["model"])
        self.assertIsNone(item["nazwa"])

    def test_auction_column_uses_id_from_xml_or_filename(self):
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        auctions = {row["Plik"]: row["Aukcja"] for row in rows}
        self.assertEqual(auctions["aukcja_loty.xml"], "AUK-2026-01")
        self.assertEqual(auctions["pakiet_items.xml"], "BATCH-77")
        self.assertEqual(auctions["raport_jeden.xml"], "RAP-5")
        # brak identyfikatora w XML -> nazwa pliku bez rozszerzenia
        self.assertEqual(auctions["plaski.xml"], "plaski")

    def test_context_fields_repeat_in_every_row(self):
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        loty = [r for r in rows if r["Plik"] == "aukcja_loty.xml"]
        for row in loty:
            self.assertEqual(row["aukcja/sprzedawca/nazwa"], "Flexit sp. z o.o.")
            self.assertEqual(row["aukcja/sprzedawca/miasto"], "Łódź")

    def test_types_are_native(self):
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        loty = [r for r in rows if r["Plik"] == "aukcja_loty.xml"]
        self.assertEqual(loty[0]["cena"], 1234.56)
        self.assertEqual(loty[0]["sztuk"], 12)
        plaskie = [r for r in rows if r["Plik"] == "plaski.xml"]
        self.assertIs(plaskie[0]["aktywny"], True)
        self.assertIs(plaskie[1]["aktywny"], False)
        self.assertEqual(plaskie[2]["ilosc"], 0)

    def test_polish_characters_and_iso_8859_2_survive(self):
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        faktura = [r for r in rows if r["Plik"] == "faktura_ns.xml"]
        self.assertEqual(faktura[0]["towar"], "Zażółć gęślą jaźń")
        self.assertEqual(faktura[1]["towar"], "Wąż strażacki")

    def test_cdata_and_entities_are_kept(self):
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        loty = [r for r in rows if r["Plik"] == "aukcja_loty.xml"]
        self.assertEqual(loty[0]["opis"], "Stan: dobry & sprawdzony")
        self.assertIsNone(loty[2]["opis"])  # brak pola = pusta komórka

    def test_formula_like_values_are_written_as_text(self):
        workbook = openpyxl.load_workbook(self.out)
        self.addCleanup(workbook.close)
        worksheet = workbook[cli.SHEET_ALL]
        header = [cell.value for cell in worksheet[1]]
        column = header.index("formula") + 1
        found = []
        for row in range(2, worksheet.max_row + 1):
            cell = worksheet.cell(row=row, column=column)
            if cell.value is not None:
                found.append((cell.value, cell.data_type))
        self.assertIn(("=SUMA(A1:A9)", "s"), found)
        self.assertIn(("-2+3", "s"), found)

    def test_summary_sheet_lists_every_file(self):
        rows = sheet_rows(self.out, cli.SHEET_SUMMARY)
        self.assertEqual(list(rows[0]), cli.SUMMARY_COLUMNS)
        data = as_dicts(rows)
        files = {row["Plik"]: row for row in data}
        for name in FIXTURES:
            self.assertIn(name, files)
            self.assertTrue(str(files[name]["Status"]).startswith("OK"))
        self.assertEqual(files["aukcja_loty.xml"]["Liczba pozycji"], 3)
        self.assertEqual(files["aukcja_loty.xml"]["Ścieżka rekordu"], "aukcja/lots/lot")
        self.assertEqual(files["raport_jeden.xml"]["Ścieżka rekordu"],
                         "(brak — cały plik jako 1 wiersz)")
        total = files["RAZEM"]
        self.assertEqual(total["Liczba pozycji"], 12)

    def test_header_is_frozen_and_filtered(self):
        workbook = openpyxl.load_workbook(self.out)
        self.addCleanup(workbook.close)
        worksheet = workbook[cli.SHEET_ALL]
        self.assertEqual(worksheet.freeze_panes, "A2")
        self.assertIsNotNone(worksheet.auto_filter.ref)
        self.assertTrue(worksheet["A1"].font.bold)


# ---------------------------------------------------------------------------
# 3. Opcje podkomendy build
# ---------------------------------------------------------------------------


@unittest.skipIf(openpyxl is None, "openpyxl potrzebny do weryfikacji wyniku")
class TestBuildOptions(CliTestCase):
    """``--per-auction``, ``--no-typing``, ``--repeat``, ``--record-path``…"""

    def setUp(self):
        super().setUp()
        write_fixtures(self.xml_dir)

    def test_per_auction_adds_one_sheet_per_auction(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out,
                           "--per-auction", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        workbook = openpyxl.load_workbook(self.out)
        self.addCleanup(workbook.close)
        self.assertEqual(workbook.sheetnames[:2], [cli.SHEET_ALL, cli.SHEET_SUMMARY])
        self.assertEqual(len(workbook.sheetnames), 2 + 5)
        for name in ("AUK-2026-01", "BATCH-77", "plaski", "FV-9", "RAP-5"):
            self.assertIn(name, workbook.sheetnames)
        worksheet = workbook["AUK-2026-01"]
        rows = [tuple(r) for r in worksheet.iter_rows(values_only=True)]
        self.assertEqual(list(rows[0][:2]), cli.TECH_COLUMNS_PER_AUCTION)
        self.assertEqual(len(rows), 1 + 3)

    def test_no_typing_keeps_everything_as_text(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out,
                           "--no-typing", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        loty = [r for r in rows if r["Plik"] == "aukcja_loty.xml"]
        self.assertEqual(loty[0]["cena"], "1 234,56")
        self.assertEqual(loty[0]["sztuk"], "12")

    def test_repeat_index_splits_repeated_children(self):
        path = os.path.join(self.xml_dir, "powtorki.xml")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(
                "<root><rec><tag>a</tag><tag>b</tag></rec>"
                "<rec><tag>c</tag><tag>d</tag></rec></root>"
            )
        result = call_cli(["build", "--in", path, "--out", self.out,
                           "--repeat", "index", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        self.assertEqual(rows[0]["tag[1]"], "a")
        self.assertEqual(rows[0]["tag[2]"], "b")

    def test_repeat_join_uses_custom_separator(self):
        path = os.path.join(self.xml_dir, "powtorki.xml")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(
                "<root><rec><tag>a</tag><tag>b</tag></rec>"
                "<rec><tag>c</tag><tag>d</tag></rec></root>"
            )
        result = call_cli(["build", "--in", path, "--out", self.out,
                           "--join-sep", " ; ", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        self.assertEqual(rows[0]["tag"], "a ; b")

    def test_record_path_can_be_forced(self):
        path = os.path.join(self.xml_dir, "aukcja_loty.xml")
        result = call_cli(["build", "--in", path, "--out", self.out,
                           "--record-path", "sprzedawca", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        summary = as_dicts(sheet_rows(self.out, cli.SHEET_SUMMARY))
        self.assertEqual(summary[0]["Ścieżka rekordu"], "aukcja/sprzedawca")

    def test_wrong_record_path_is_a_file_error_not_a_crash(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out,
                           "--record-path", "nie-ma-takiej-sciezki", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_NO_DATA)
        self.assertIn("Nie znaleziono elementów", result.err)

    def test_limit_cuts_number_of_files(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out,
                           "--limit", "2", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        self.assertEqual(len({row["Plik"] for row in rows}), 2)

    def test_keep_ns_changes_column_names(self):
        path = os.path.join(self.xml_dir, "faktura_ns.xml")
        result = call_cli(["build", "--in", path, "--out", self.out,
                           "--keep-ns", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        header = sheet_rows(self.out, cli.SHEET_ALL)[0]
        self.assertTrue(any("f:towar" in str(name) for name in header), header)


# ---------------------------------------------------------------------------
# 4. Ujednolicanie kolumn między plikami tego samego schematu
# ---------------------------------------------------------------------------


@unittest.skipIf(openpyxl is None, "openpyxl potrzebny do weryfikacji wyniku")
class TestUnifyRecordPath(CliTestCase):
    """Plik z JEDNĄ pozycją musi dostać te same kolumny co plik z wieloma."""

    JEDEN = ("<batch auction='X-1'><lot><title>Jeden</title><items>"
             "<item nr='1'><model>A</model></item></items></lot></batch>")
    WIELE = ("<batch auction='X-2'><lot><title>Wiele</title><items>"
             "<item nr='1'><model>B</model></item>"
             "<item nr='2'><model>C</model></item></items></lot></batch>")

    def setUp(self):
        super().setUp()
        for name, text in (("jeden.xml", self.JEDEN), ("wiele.xml", self.WIELE)):
            with open(os.path.join(self.xml_dir, name), "w", encoding="utf-8") as fh:
                fh.write(text)

    def test_columns_are_unified_by_default(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out, "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        self.assertEqual([row["model"] for row in rows], ["A", "B", "C"])
        summary = {row["Plik"]: row for row in as_dicts(sheet_rows(self.out, cli.SHEET_SUMMARY))}
        self.assertEqual(summary["jeden.xml"]["Ścieżka rekordu"], "batch/lot/items/item")
        self.assertIn("ujednolicona", summary["jeden.xml"]["Status"])

    def test_no_unify_leaves_the_original_split(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out,
                           "--no-unify", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        header = sheet_rows(self.out, cli.SHEET_ALL)[0]
        self.assertIn("lot/items/item/model", header)
        self.assertIn("model", header)

    def test_unification_does_not_touch_foreign_schemas(self):
        write_fixtures(self.xml_dir, {"raport_jeden.xml"})
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out, "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        summary = {row["Plik"]: row for row in as_dicts(sheet_rows(self.out, cli.SHEET_SUMMARY))}
        self.assertEqual(summary["raport_jeden.xml"]["Ścieżka rekordu"],
                         "(brak — cały plik jako 1 wiersz)")


# ---------------------------------------------------------------------------
# 5. Błędy pojedynczych plików
# ---------------------------------------------------------------------------


@unittest.skipIf(openpyxl is None, "openpyxl potrzebny do weryfikacji wyniku")
class TestBrokenFiles(CliTestCase):
    """Zły plik nie może przerwać całości — ma trafić do Podsumowania."""

    def setUp(self):
        super().setUp()
        write_fixtures(self.xml_dir)
        with open(os.path.join(self.xml_dir, "zepsuty.xml"), "w", encoding="utf-8") as fh:
            fh.write(XML_ZEPSUTY)
        with open(os.path.join(self.xml_dir, "pusty.xml"), "w", encoding="utf-8") as fh:
            fh.write("")
        self.result = call_cli(["build", "--in", self.xml_dir, "--out", self.out])

    def test_exit_code_signals_partial_success(self):
        self.assertEqual(self.result.code, cli.EXIT_PARTIAL)

    def test_good_rows_are_still_there(self):
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        self.assertEqual(len(rows), 12)

    def test_errors_are_listed_in_summary_sheet(self):
        summary = {row["Plik"]: row for row in as_dicts(sheet_rows(self.out, cli.SHEET_SUMMARY))}
        self.assertIn("zepsuty.xml", summary)
        self.assertTrue(str(summary["zepsuty.xml"]["Status"]).startswith("BŁĄD:"))
        self.assertTrue(str(summary["pusty.xml"]["Status"]).startswith("BŁĄD:"))
        self.assertIn("z błędem: 2 pliki", str(summary["RAZEM"]["Status"]))

    def test_errors_are_repeated_on_stderr(self):
        self.assertIn("zepsuty.xml", self.result.err)
        self.assertIn("Pominięto 2 pliki", self.result.err)

    def test_all_files_broken_gives_no_data(self):
        directory = os.path.join(self.tmp, "same_zepsute")
        os.makedirs(directory)
        with open(os.path.join(directory, "a.xml"), "w", encoding="utf-8") as fh:
            fh.write(XML_ZEPSUTY)
        result = call_cli(["build", "--in", directory,
                           "--out", os.path.join(self.tmp, "inny.xlsx"), "--quiet"])
        self.assertEqual(result.code, cli.EXIT_NO_DATA)
        self.assertFalse(os.path.exists(os.path.join(self.tmp, "inny.xlsx")))


# ---------------------------------------------------------------------------
# 6. Kolizje nazw kolumn technicznych
# ---------------------------------------------------------------------------


@unittest.skipIf(openpyxl is None, "openpyxl potrzebny do weryfikacji wyniku")
class TestTechnicalColumnCollision(CliTestCase):
    """XML z własnym polem ``Aukcja`` nie może nadpisać kolumny technicznej."""

    def test_colliding_column_gets_suffix(self):
        path = os.path.join(self.xml_dir, "kolizja.xml")
        with open(path, "w", encoding="utf-8") as fh:
            fh.write(
                "<root>"
                "<row><Aukcja>z pliku</Aukcja><Plik>x.pdf</Plik></row>"
                "<row><Aukcja>z pliku 2</Aukcja><Plik>y.pdf</Plik></row>"
                "</root>"
            )
        result = call_cli(["build", "--in", path, "--out", self.out, "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        rows = sheet_rows(self.out, cli.SHEET_ALL)
        header = list(rows[0])
        self.assertEqual(header[:3], cli.TECH_COLUMNS)
        self.assertIn("Aukcja (XML)", header)
        self.assertIn("Plik (XML)", header)
        data = as_dicts(rows)
        self.assertEqual(data[0]["Aukcja (XML)"], "z pliku")
        self.assertEqual(data[0]["Aukcja"], "kolizja")

    def test_unique_headers_helper(self):
        self.assertEqual(
            cli._unique_headers(["Aukcja", "model"], cli.TECH_COLUMNS),
            ["Aukcja (XML)", "model"],
        )


# ---------------------------------------------------------------------------
# 7. Zapis pliku wynikowego: --overwrite, --dry-run, ścieżki
# ---------------------------------------------------------------------------


class TestOutputHandling(CliTestCase):
    """Nadpisywanie, tryb próbny i tworzenie katalogów."""

    def setUp(self):
        super().setUp()
        write_fixtures(self.xml_dir)

    def test_existing_file_is_not_overwritten_without_flag(self):
        with open(self.out, "w", encoding="utf-8") as fh:
            fh.write("stara zawartość")
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out, "--quiet"])
        self.assertEqual(result.code, cli.EXIT_ERROR)
        self.assertIn("--overwrite", result.err)
        with open(self.out, "r", encoding="utf-8") as fh:
            self.assertEqual(fh.read(), "stara zawartość")

    def test_overwrite_flag_replaces_the_file(self):
        with open(self.out, "w", encoding="utf-8") as fh:
            fh.write("stara zawartość")
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out,
                           "--overwrite", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        with open(self.out, "rb") as fh:
            self.assertEqual(fh.read(2), b"PK")  # to już archiwum XLSX

    def test_dry_run_writes_nothing(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out, "--dry-run"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertFalse(os.path.exists(self.out))
        self.assertIn("dry-run", result.out)
        self.assertIn(cli.SHEET_ALL, result.out)

    def test_missing_parent_directory_is_created(self):
        target = os.path.join(self.tmp, "nowy", "katalog", "plik.xlsx")
        result = call_cli(["build", "--in", self.xml_dir, "--out", target, "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertTrue(os.path.isfile(target))

    def test_directory_as_output_is_refused(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.tmp, "--quiet"])
        self.assertEqual(result.code, cli.EXIT_ERROR)
        self.assertIn("katalog", result.err)

    def test_empty_input_directory_reports_no_data(self):
        empty = os.path.join(self.tmp, "pusty_katalog")
        os.makedirs(empty)
        result = call_cli(["build", "--in", empty, "--out", self.out, "--quiet"])
        self.assertEqual(result.code, cli.EXIT_NO_DATA)
        self.assertIn("Nie znalazłem żadnego pliku", result.err)

    def test_quiet_silences_progress_but_not_errors(self):
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out, "--quiet"])
        self.assertEqual(result.out, "")
        result = call_cli(["build", "--in", "/nie/ma", "--out", self.out, "--quiet"])
        self.assertNotEqual(result.err, "")


# ---------------------------------------------------------------------------
# 8. Parser i kody wyjścia
# ---------------------------------------------------------------------------


class TestParser(unittest.TestCase):
    """Składnia poleceń."""

    def test_no_arguments_prints_help(self):
        result = call_cli([])
        self.assertEqual(result.code, cli.EXIT_USAGE)
        self.assertIn("download", result.out)
        self.assertIn("build", result.out)
        self.assertIn("all", result.out)

    def test_unknown_subcommand_exits_with_usage_code(self):
        with self.assertRaises(SystemExit) as ctx:
            with contextlib.redirect_stderr(io.StringIO()):
                cli.main(["nieznana"])
        self.assertEqual(ctx.exception.code, cli.EXIT_USAGE)

    def test_version_flag(self):
        with self.assertRaises(SystemExit) as ctx:
            with contextlib.redirect_stdout(io.StringIO()):
                cli.main(["--version"])
        self.assertEqual(ctx.exception.code, 0)

    def test_in_accepts_several_paths_and_repeats(self):
        parser = cli.build_parser()
        args = parser.parse_args(["build", "--in", "a", "b", "--in", "c", "d"])
        cli._normalize_inputs(args)
        self.assertEqual(args.inputs, ["a", "b", "c", "d"])

    def test_positional_paths_are_accepted(self):
        parser = cli.build_parser()
        args = parser.parse_args(["build", "plik.xml"])
        cli._normalize_inputs(args)
        self.assertEqual(args.inputs, ["plik.xml"])

    def test_build_without_input_reports_no_data(self):
        result = call_cli(["build", "--out", "nieistotne.xlsx", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_NO_DATA)
        self.assertIn("--in", result.err)


# ---------------------------------------------------------------------------
# 9. Uruchomienie jako moduł (python3 -m flexit2xlsx)
# ---------------------------------------------------------------------------


class TestModuleEntryPoint(CliTestCase):
    """``python3 -m flexit2xlsx`` musi działać jak wywołanie funkcji."""

    def test_module_runs_and_builds_file(self):
        write_fixtures(self.xml_dir)
        process = subprocess.run(
            [sys.executable, "-m", "flexit2xlsx", "build",
             "--in", self.xml_dir, "--out", self.out, "--quiet"],
            cwd=ROOT_DIR, capture_output=True, text=True,
        )
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertTrue(os.path.isfile(self.out))

    def test_help_is_in_polish(self):
        process = subprocess.run(
            [sys.executable, "-m", "flexit2xlsx", "--help"],
            cwd=ROOT_DIR, capture_output=True, text=True,
        )
        self.assertEqual(process.returncode, 0, process.stderr)
        self.assertIn("Pobiera pliki XML", process.stdout)


# ---------------------------------------------------------------------------
# 10. Pobieranie z atrapy portalu
# ---------------------------------------------------------------------------


@unittest.skipIf(MockSite is None,
                 "brak atrapy portalu (tests/test_scrape.py nie udostępnia MockSite)")
class TestDownloadAgainstMockSite(CliTestCase):
    """``download`` i ``all`` przeciwko atrapie portalu z ``tests/mocksite``.

    Atrapa to serwer HTTP na localhoście z układem opisanym w SITE_NOTES.md:
    lista aukcji -> strona aukcji -> strony lotów -> XML "Download Batch Details".
    ``--delay 0`` wyłącza uprzejme opóźnienia, więc testy są szybkie.
    """

    @classmethod
    def setUpClass(cls):
        cls.site = MockSite()

    @classmethod
    def tearDownClass(cls):
        cls.site.stop()

    def base_args(self, *extra):
        return ["--base-url", self.site.base, "--delay", "0", "--timeout", "5"] + list(extra)

    def test_download_fetches_every_xml(self):
        result = call_cli(["download", "--out", self.xml_dir] + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        files = sorted(os.listdir(self.xml_dir))
        self.assertEqual(len(files), 6, files)
        for name in files:
            self.assertTrue(name.endswith(".xml"), name)
            with open(os.path.join(self.xml_dir, name), "rb") as handle:
                self.assertTrue(handle.read().lstrip().startswith(b"<?xml"))
        self.assertIn("Pobrano 6 plików", result.out)

    def test_second_run_skips_existing_files(self):
        call_cli(["download", "--out", self.xml_dir] + self.base_args("--quiet"))
        before = sorted(os.listdir(self.xml_dir))
        result = call_cli(["download", "--out", self.xml_dir] + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertEqual(sorted(os.listdir(self.xml_dir)), before)
        self.assertIn("pominięto 6 plików", result.out)

    def test_dry_run_downloads_nothing(self):
        result = call_cli(["download", "--out", self.xml_dir, "--dry-run"] + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertEqual(os.listdir(self.xml_dir), [])
        self.assertIn("dry-run", result.out)
        self.assertIn(".xml", result.out)

    def test_limit_stops_after_n_auctions(self):
        result = call_cli(["download", "--out", self.xml_dir, "--limit", "1"]
                          + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertLess(len(os.listdir(self.xml_dir)), 6)

    def test_match_filters_auctions(self):
        result = call_cli(["download", "--out", self.xml_dir, "--match", "1103", "--dry-run"]
                          + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertIn("1103", result.out)
        self.assertNotIn("/lot/dell-optiplex-mix-cd048", result.out)

    def test_no_auctions_found_gives_no_data_code(self):
        result = call_cli(["download", "--out", self.xml_dir,
                           "--base-url", self.site.url("/pusto"),
                           "--delay", "0", "--timeout", "5"])
        self.assertEqual(result.code, cli.EXIT_NO_DATA)
        self.assertIn("Nie znalazłem żadnej aukcji", result.err)

    def test_unreachable_page_gives_network_code(self):
        result = call_cli(["download", "--out", self.xml_dir,
                           "--base-url", self.site.url("/nie-ma-takiej-strony"),
                           "--delay", "0", "--timeout", "5", "--retries", "0"])
        self.assertEqual(result.code, cli.EXIT_NETWORK)
        self.assertIn("--cookie", result.err)

    def test_custom_auction_regex_is_used(self):
        result = call_cli(["download", "--out", self.xml_dir, "--dry-run",
                           "--auction-re", r"/auction/(?P<id>[^/?#]*1103)/?$"]
                          + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertIn("1103", result.out)
        self.assertNotIn("1087", result.out)

    def test_broken_regex_is_reported(self):
        result = call_cli(["download", "--out", self.xml_dir,
                           "--auction-re", "([niedomkniete"] + self.base_args())
        self.assertEqual(result.code, cli.EXIT_ERROR)
        self.assertIn("--auction-re", result.err)

    def test_cookie_header_reaches_the_server(self):
        self.site.state.reset()
        call_cli(["download", "--out", self.xml_dir, "--cookie", "sid=tajne",
                  "--limit", "1", "--dry-run"] + self.base_args())
        cookies = [head.get("cookie") for _, head in self.site.state.headers]
        self.assertTrue(cookies, "atrapa nie zapisała żadnych nagłówków")
        self.assertTrue(all(value == "sid=tajne" for value in cookies), cookies)

    def test_all_dry_run_does_not_build_anything(self):
        result = call_cli(["all", "--in", self.xml_dir, "--out", self.out, "--dry-run"]
                          + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertFalse(os.path.exists(self.out))
        self.assertEqual(os.listdir(self.xml_dir), [])
        self.assertIn("nie ma jeszcze plików XML", result.out)

    @unittest.skipIf(openpyxl is None, "openpyxl potrzebny do weryfikacji wyniku")
    def test_all_downloads_and_builds_in_one_go(self):
        result = call_cli(["all", "--in", self.xml_dir, "--out", self.out]
                          + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertEqual(len(os.listdir(self.xml_dir)), 6)
        self.assertTrue(os.path.isfile(self.out))
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        # 6 pobranych pakietów: 3 + 2 + 2 + 1 + 1 + 1 sztuk sprzętu
        self.assertEqual(len(rows), 10)
        self.assertEqual(len({row["Aukcja"] for row in rows}), 2)
        # kolumny zunifikowane mimo różnej liczby pozycji w pakietach:
        # KAŻDY wiersz ma wypełnioną tę samą kolumnę "model"
        self.assertEqual(sum(1 for row in rows if row["model"]), 10)
        # nazwy plików są bezpieczne — path traversal z portalu nie przechodzi
        for name in os.listdir(self.xml_dir):
            self.assertNotIn("..", name)
            self.assertEqual(name, os.path.basename(name))


# ---------------------------------------------------------------------------
# 11. Zgodność z kontraktem INTERFACES.md
# ---------------------------------------------------------------------------


class TestContract(unittest.TestCase):
    """Podkomendy i arkusze wymagane przez kontrakt."""

    def test_three_subcommands_exist(self):
        parser = cli.build_parser()
        for command in ("download", "build", "all"):
            args = parser.parse_args([command])
            self.assertEqual(args.command, command)

    def test_required_options_are_present(self):
        parser = cli.build_parser()
        args = parser.parse_args([
            "all", "--out", "w.xlsx", "--in", "kat", "--per-auction",
            "--repeat", "index", "--record-path", "a/b", "--no-typing",
            "--dry-run", "--cookie", "c=1", "--base-url", "http://x/",
            "--auction-re", "/a/", "--limit", "3", "--overwrite", "--quiet",
        ])
        self.assertEqual(args.out, "w.xlsx")
        self.assertEqual(args.xml_dir, "kat")
        self.assertTrue(args.per_auction)
        self.assertEqual(args.repeat, "index")
        self.assertEqual(args.record_path, "a/b")
        self.assertTrue(args.no_typing)
        self.assertTrue(args.dry_run)
        self.assertEqual(args.cookie, "c=1")
        self.assertEqual(args.base_url, "http://x/")
        self.assertEqual(args.auction_re, "/a/")
        self.assertEqual(args.limit, 3)
        self.assertTrue(args.overwrite)
        self.assertTrue(args.quiet)

    def test_sheet_names_from_contract(self):
        self.assertEqual(cli.SHEET_ALL, "Wszystkie aukcje")
        self.assertEqual(cli.SHEET_SUMMARY, "Podsumowanie")
        self.assertEqual(cli.TECH_COLUMNS, ["Aukcja", "Plik", "Nr pozycji"])

    def test_build_sheets_produces_generators_not_lists(self):
        doc = xmlflatten.parse_bytes(XML_PLASKI.encode("utf-8"), "x.xml")
        sheets = cli.build_sheets([doc])
        self.assertIsInstance(sheets[0], xlsxwrite.Sheet)
        self.assertEqual(sheets[0].name, cli.SHEET_ALL)
        self.assertFalse(isinstance(sheets[0].rows, list))
        self.assertEqual(len(list(sheets[0].rows)), 4)


if __name__ == "__main__":  # pragma: no cover
    unittest.main(verbosity=2)
