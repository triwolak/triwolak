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
import unittest.mock

TESTS_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(TESTS_DIR)
for _path in (ROOT_DIR, TESTS_DIR):
    if _path not in sys.path:
        sys.path.insert(0, _path)

from flexit2xlsx import cli, values, xmlflatten  # noqa: E402
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
        # zakładki są numerowane w kolejności z arkusza Podsumowanie, a sama
        # nazwa aukcji zostaje rozpoznawalna (patrz _auction_sheet_name)
        for name in ("AUK-2026-01", "BATCH-77", "plaski", "FV-9", "RAP-5"):
            self.assertTrue(any(tab.endswith(" " + name) for tab in workbook.sheetnames),
                            "brak zakładki dla %s w %s" % (name, workbook.sheetnames))
        worksheet = workbook["01 AUK-2026-01"]
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
        # 7, bo aukcja 1103 ma zbiorczy XML ORAZ pakiet lotu bcd12 —
        # "--descend auto" schodzi teraz do lotów także wtedy, gdy strona
        # aukcji sama coś dała
        self.assertEqual(len(files), 7, files)
        for name in files:
            self.assertTrue(name.endswith(".xml"), name)
            with open(os.path.join(self.xml_dir, name), "rb") as handle:
                self.assertTrue(handle.read().lstrip().startswith(b"<?xml"))
        self.assertIn("Pobrano 7 plików", result.out)

    def test_second_run_skips_existing_files(self):
        call_cli(["download", "--out", self.xml_dir] + self.base_args("--quiet"))
        before = sorted(os.listdir(self.xml_dir))
        result = call_cli(["download", "--out", self.xml_dir] + self.base_args())
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        self.assertEqual(sorted(os.listdir(self.xml_dir)), before)
        self.assertIn("pominięto 7 plików", result.out)

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
        self.assertEqual(len(os.listdir(self.xml_dir)), 7)
        self.assertTrue(os.path.isfile(self.out))
        rows = as_dicts(sheet_rows(self.out, cli.SHEET_ALL))
        # 7 pobranych pakietów: 3 + 2 + 2 + 1 + 1 + 1 + 1 sztuk sprzętu
        self.assertEqual(len(rows), 11)
        self.assertEqual(len({row["Aukcja"] for row in rows}), 2)
        # kolumny zunifikowane mimo różnej liczby pozycji w pakietach:
        # KAŻDY wiersz ma wypełnioną tę samą kolumnę "model"
        self.assertEqual(sum(1 for row in rows if row["model"]), 11)
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


# ---------------------------------------------------------------------------
# 12. REGRESJE po audycie adwersaryjnym
# ---------------------------------------------------------------------------


def _batch_xml(auction: str, items: int, first: int = 0) -> str:
    """Realistyczny "Batch Details": aukcja -> lot -> pozycje."""
    body = "".join(
        "<item nr='%d'><model>ThinkPad T%d</model><serial>SN%06d</serial>"
        "<grade>A</grade><ram>16 GB</ram></item>" % (i, 480 + i % 3, first + i)
        for i in range(items)
    )
    return ("<?xml version='1.0' encoding='UTF-8'?>"
            "<batch auction='%s'><lot><title>Pakiet</title><items>%s</items>"
            "</lot></batch>" % (auction, body))


class TestPamieciPotoku(CliTestCase):
    """REGRESJA: potok ma być strumieniowy — pamięć O(największy plik)."""

    def _corpus(self, files, items):
        for index in range(files):
            path = os.path.join(self.xml_dir, "b%03d.xml" % index)
            with open(path, "w", encoding="utf-8") as handle:
                handle.write(_batch_xml("auk-%03d" % index, items, index * 1000))

    def test_dokumenty_sa_zwalniane_po_przekroczeniu_budzetu(self):
        """Powyżej ``MEMORY_ROW_BUDGET`` rekordy nie zostają w pamięci.

        Dawniej ``load_documents`` trzymało wszystkie ``ParsedDoc`` aż do końca
        zapisu, więc szczyt pamięci rósł LINIOWO z rozmiarem całego korpusu
        (~715 B na wiersz; milion wierszy = ponad 800 MB), mimo że
        ``xlsxwrite`` zapisuje strumieniowo.
        """
        self._corpus(files=10, items=10)          # 100 wierszy razem
        files = cli.collect_xml_files([self.xml_dir])
        saved = cli.MEMORY_ROW_BUDGET
        try:
            cli.MEMORY_ROW_BUDGET = 25            # budżet mniejszy niż korpus
            docs, errors = cli.load_documents(files)
        finally:
            cli.MEMORY_ROW_BUDGET = saved
        self.assertEqual(errors, [])
        self.assertEqual(len(docs), 10)
        self.assertFalse(any(doc.cached for doc in docs),
                         "po przekroczeniu budżetu żaden dokument nie zostaje w RAM")
        # ...a mimo to metadane i wiersze są kompletne
        self.assertEqual(sum(doc.count for doc in docs), 100)
        columns = xmlflatten.merge_columns(docs)
        rows = list(cli._iter_rows(docs, columns, with_auction=True, use_typing=True))
        self.assertEqual(len(rows), 100)
        self.assertEqual(len({row[0] for row in rows}), 10)

    def test_male_korpusy_zostaja_w_pamieci(self):
        """Poniżej budżetu nie ma ponownego parsowania (szybciej)."""
        self._corpus(files=3, items=5)
        docs, _ = cli.load_documents(cli.collect_xml_files([self.xml_dir]))
        self.assertTrue(all(doc.cached for doc in docs))

    def test_wynik_jest_taki_sam_w_obu_trybach(self):
        """Tryb strumieniowy nie może zmienić ANI JEDNEGO wiersza."""
        self._corpus(files=6, items=7)
        args = ["build", "--in", self.xml_dir, "--quiet", "--overwrite"]
        out_a = os.path.join(self.tmp, "a.xlsx")
        out_b = os.path.join(self.tmp, "b.xlsx")
        saved = cli.MEMORY_ROW_BUDGET
        try:
            cli.MEMORY_ROW_BUDGET = 10 ** 9
            self.assertEqual(call_cli(args + ["--out", out_a]).code, cli.EXIT_OK)
            cli.MEMORY_ROW_BUDGET = 1
            self.assertEqual(call_cli(args + ["--out", out_b]).code, cli.EXIT_OK)
        finally:
            cli.MEMORY_ROW_BUDGET = saved
        if openpyxl is None:
            self.skipTest("openpyxl potrzebny do porównania wyników")
        self.assertEqual(sheet_rows(out_a, cli.SHEET_ALL),
                         sheet_rows(out_b, cli.SHEET_ALL))
        self.assertEqual(sheet_rows(out_a, cli.SHEET_SUMMARY),
                         sheet_rows(out_b, cli.SHEET_SUMMARY))

    def test_znikniety_plik_nie_wywraca_zapisu(self):
        """Gdy plik zniknie po wczytaniu metadanych, tracimy JEGO wiersze, nie całość."""
        self._corpus(files=3, items=4)
        files = cli.collect_xml_files([self.xml_dir])
        saved = cli.MEMORY_ROW_BUDGET
        try:
            cli.MEMORY_ROW_BUDGET = 1
            docs, _ = cli.load_documents(files)
        finally:
            cli.MEMORY_ROW_BUDGET = saved
        os.remove(docs[1].source)
        reporter = cli.Reporter(quiet=True, err=io.StringIO())
        columns = xmlflatten.merge_columns(docs)
        rows = list(cli._iter_rows(docs, columns, with_auction=True,
                                   use_typing=True, reporter=reporter))
        self.assertEqual(len(rows), 8)
        self.assertIn("Nie udało się ponownie wczytać", reporter.err.getvalue())


class TestBrakuPamieci(CliTestCase):
    """REGRESJA: ``MemoryError`` ma dawać komunikat po polsku, nie traceback."""

    def test_plik_ktory_nie_miesci_sie_w_pamieci_jest_pomijany(self):
        write_fixtures(self.xml_dir)
        zly = os.path.join(self.xml_dir, "za-duzy.xml")
        with open(zly, "w", encoding="utf-8") as handle:
            handle.write(XML_PLASKI)

        prawdziwy = xmlflatten.parse_file

        def czasem_brak_pamieci(path, **kw):
            if os.path.basename(path) == "za-duzy.xml":
                raise MemoryError()
            return prawdziwy(path, **kw)

        with unittest.mock.patch.object(xmlflatten, "parse_file", czasem_brak_pamieci):
            result = call_cli(["build", "--in", self.xml_dir, "--out", self.out])
        self.assertEqual(result.code, cli.EXIT_PARTIAL)
        self.assertIn("za mało pamięci", result.err)
        self.assertNotIn("Traceback", result.err)
        self.assertTrue(os.path.isfile(self.out), "reszta plików mimo wszystko trafia do wyniku")

    def test_main_lapie_memoryerror_z_dowolnego_miejsca(self):
        def wybuchowe(*args, **kw):
            raise MemoryError()

        with unittest.mock.patch.object(cli, "cmd_build", wybuchowe):
            result = call_cli(["build", "--in", self.xml_dir, "--out", self.out])
        self.assertEqual(result.code, cli.EXIT_ERROR)
        self.assertIn("Za mało pamięci", result.err)
        self.assertNotIn("Traceback", result.err)

    def test_main_nie_wypuszcza_zadnego_wyjatku(self):
        """Ostatnia siatka bezpieczeństwa: żaden błąd nie wychodzi tracebackiem."""
        def wybuchowe(*args, **kw):
            raise RuntimeError("coś się urwało w środku")

        with unittest.mock.patch.object(cli, "cmd_build", wybuchowe):
            result = call_cli(["build", "--in", self.xml_dir, "--out", self.out])
        self.assertEqual(result.code, cli.EXIT_ERROR)
        self.assertIn("Nieoczekiwany błąd", result.err)
        self.assertIn("coś się urwało w środku", result.err)


class TestKosztuUjednolicania(CliTestCase):
    """REGRESJA: ``unify_record_paths`` nie może mieć kosztu kwadratowego."""

    def test_liczba_parsowan_nie_jest_iloczynem(self):
        """Dawniej: N plików bez powtórzeń x M ścieżek = N*M parsowań.

        Pliki bez powtarzalnego elementu nie zawierają ŻADNEGO ze znaczników
        kandydatów, więc po przedfiltrze nie ma czego parsować.
        """
        for index in range(20):
            tag = "poz%03d" % index
            body = "".join("<%s><m>M%d</m></%s>" % (tag, j, tag) for j in range(3))
            with open(os.path.join(self.xml_dir, "wiele%03d.xml" % index), "w",
                      encoding="utf-8") as handle:
                handle.write("<?xml version='1.0'?><r%03d><items>%s</items></r%03d>"
                             % (index, body, index))
        for index in range(20):
            with open(os.path.join(self.xml_dir, "jeden%03d.xml" % index), "w",
                      encoding="utf-8") as handle:
                handle.write("<?xml version='1.0'?><rX><items>"
                             "<inny><m>M</m></inny></items></rX>")

        docs, _ = cli.load_documents(cli.collect_xml_files([self.xml_dir]))
        licznik = {"n": 0}
        prawdziwy = xmlflatten.parse_bytes

        def counted(*args, **kw):
            licznik["n"] += 1
            return prawdziwy(*args, **kw)

        with unittest.mock.patch.object(xmlflatten, "parse_bytes", counted):
            cli.unify_record_paths(docs)
        self.assertEqual(licznik["n"], 0, "przedfiltr odsiał wszystkich kandydatów")

    def test_najwyzej_kilka_prob_na_plik(self):
        """Gdy znacznik JEST w pliku, próbujemy najwyżej kilku ścieżek."""
        for index in range(5):
            body = "".join("<item><m>M%d</m></item>" % j for j in range(3))
            with open(os.path.join(self.xml_dir, "wiele%03d.xml" % index), "w",
                      encoding="utf-8") as handle:
                handle.write("<?xml version='1.0'?><r%03d><items>%s</items></r%03d>"
                             % (index, body, index))
        # plik jednopozycyjny z tym samym znacznikiem, ale w innym miejscu drzewa
        with open(os.path.join(self.xml_dir, "jeden.xml"), "w", encoding="utf-8") as handle:
            handle.write("<?xml version='1.0'?><inny><gdzies><item><m>M</m></item>"
                         "</gdzies></inny>")
        docs, _ = cli.load_documents(cli.collect_xml_files([self.xml_dir]))
        licznik = {"n": 0}
        prawdziwy = xmlflatten.parse_bytes

        def counted(*args, **kw):
            licznik["n"] += 1
            return prawdziwy(*args, **kw)

        with unittest.mock.patch.object(xmlflatten, "parse_bytes", counted):
            cli.unify_record_paths(docs)
        self.assertLessEqual(licznik["n"], cli.UNIFY_MAX_CANDIDATES)

    def test_ujednolicanie_dalej_dziala(self):
        """Sedno funkcji: plik z JEDNĄ pozycją dostaje kolumny jak reszta."""
        with open(os.path.join(self.xml_dir, "wiele.xml"), "w", encoding="utf-8") as handle:
            handle.write(_batch_xml("auk-1", 3))
        with open(os.path.join(self.xml_dir, "jeden.xml"), "w", encoding="utf-8") as handle:
            handle.write(_batch_xml("auk-2", 1))
        docs, _ = cli.load_documents(cli.collect_xml_files([self.xml_dir]))
        unified = cli.unify_record_paths(docs)
        self.assertEqual({os.path.basename(p) for p in unified}, {"jeden.xml"})
        self.assertEqual(len({doc.record_path for doc in docs}), 1)
        for doc in docs:
            self.assertIn("model", doc.columns)


class TestPamieciRozpoznawaniaTypow(CliTestCase):
    """REGRESJA: bufor wyników ``coerce_value`` nie może zmieniać wyników."""

    def test_bufor_daje_te_same_wartosci_co_bez_bufora(self):
        probki = ["1 234,56", "1234.56", "007", "true", "FALSE", "2026-06-18",
                  "18.06.2026", "16 GB", "1.2.3", "", "   ", "PF1A2B3C",
                  "-12", "+3.5", "e0c8f", "192.168.0.1", "1,5", "  7  "]
        cache = {}
        for raw in probki * 3:
            with self.subTest(raw=raw):
                self.assertEqual(cli._convert(raw, True, cache),
                                 values.coerce_value(raw))
        self.assertLessEqual(len(cache), len(set(probki)))

    def test_bufor_nie_rosnie_ponad_limit(self):
        cache = {}
        saved = cli.TYPE_CACHE_MAX
        try:
            cli.TYPE_CACHE_MAX = 5
            for index in range(50):
                cli._convert("wartosc-%d" % index, True, cache)
        finally:
            cli.TYPE_CACHE_MAX = saved
        self.assertEqual(len(cache), 5)

    def test_dlugie_napisy_nie_trafiaja_do_bufora(self):
        cache = {}
        dlugi = "x" * (cli.TYPE_CACHE_MAX_LEN + 1)
        self.assertEqual(cli._convert(dlugi, True, cache), values.coerce_value(dlugi))
        self.assertEqual(cache, {})

    def test_wylaczone_typowanie_zwraca_oryginal(self):
        self.assertEqual(cli._convert("1 234,56", False, {}), "1 234,56")


class TestPodzialuNaBlokiKolumn(CliTestCase):
    """REGRESJA: arkusz-kontynuacja też musi mieć kolumny techniczne."""

    def test_sheet_prosi_o_powtorzenie_kolumn_technicznych(self):
        doc = xmlflatten.parse_bytes(XML_PLASKI.encode("utf-8"), "x.xml")
        sheets = cli.build_sheets([doc], per_auction=True)
        self.assertEqual(getattr(sheets[0], "key_columns", None), len(cli.TECH_COLUMNS))
        self.assertEqual(getattr(sheets[-1], "key_columns", None),
                         len(cli.TECH_COLUMNS_PER_AUCTION))

    @unittest.skipIf(openpyxl is None, "openpyxl potrzebny do odczytu wyniku")
    def test_kolumny_techniczne_powtorzone_w_arkuszu_kontynuacji(self):
        """Bez tego wiersza z drugiego arkusza nie da się przypisać do aukcji."""
        saved = xlsxwrite.MAX_COLS_PER_SHEET
        try:
            xlsxwrite.MAX_COLS_PER_SHEET = 40
            pola = "".join("<p%03d>v%03d</p%03d>" % (i, i, i) for i in range(60))
            with open(os.path.join(self.xml_dir, "szeroki.xml"), "w",
                      encoding="utf-8") as handle:
                handle.write("<?xml version='1.0'?><batch><items>"
                             "<item>%s</item><item>%s</item></items></batch>"
                             % (pola, pola))
            result = call_cli(["build", "--in", self.xml_dir, "--out", self.out,
                               "--overwrite", "--quiet"])
        finally:
            xlsxwrite.MAX_COLS_PER_SHEET = saved
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        workbook = openpyxl.load_workbook(self.out)
        self.addCleanup(workbook.close)
        dalsze = [name for name in workbook.sheetnames if name.startswith(cli.SHEET_ALL)]
        self.assertGreater(len(dalsze), 1, workbook.sheetnames)
        for name in dalsze:
            header = [cell.value for cell in next(workbook[name].iter_rows(max_row=1))]
            self.assertEqual(header[:len(cli.TECH_COLUMNS)], cli.TECH_COLUMNS,
                             "arkusz %s bez kolumn technicznych" % name)


class TestNazwZakladekAukcji(CliTestCase):
    """REGRESJA: zakładka ma pozwolić rozpoznać aukcję."""

    def test_dlugi_identyfikator_zachowuje_koncowke(self):
        used = set()
        name = cli._auction_sheet_name("flexit-auctions-18-06-2026-numer-0007", 7, used)
        self.assertLessEqual(len(name), xlsxwrite.SHEET_NAME_MAX)
        self.assertTrue(name.startswith("07 "), name)
        self.assertTrue(name.endswith("numer-0007"), name)

    @unittest.skipIf(openpyxl is None, "openpyxl potrzebny do odczytu wyniku")
    def test_wiele_aukcji_o_podobnych_nazwach_da_sie_rozroznic(self):
        for index in range(12):
            auction = "flexit-auctions-18-06-2026-numer-%04d" % index
            with open(os.path.join(self.xml_dir, "b%02d.xml" % index), "w",
                      encoding="utf-8") as handle:
                handle.write(_batch_xml(auction, 2, index * 10))
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out,
                           "--per-auction", "--overwrite", "--quiet"])
        self.assertEqual(result.code, cli.EXIT_OK, result.err)
        workbook = openpyxl.load_workbook(self.out)
        self.addCleanup(workbook.close)
        zakladki = workbook.sheetnames[2:]
        self.assertEqual(len(zakladki), 12)
        for index, name in enumerate(zakladki):
            self.assertTrue(name.endswith("numer-%04d" % index),
                            "zakładka %r nie identyfikuje aukcji" % name)


class TestKolejnosciWAll(CliTestCase):
    """REGRESJA: ``all`` sprawdza plik wynikowy PRZED pobieraniem."""

    @unittest.skipIf(MockSite is None, "brak atrapy portalu")
    def test_istniejacy_plik_wynikowy_zatrzymuje_pobieranie(self):
        site = MockSite()
        self.addCleanup(site.stop)
        site.state.reset()
        with open(self.out, "w", encoding="utf-8") as handle:
            handle.write("stary wynik")
        result = call_cli(["all", "--in", self.xml_dir, "--out", self.out,
                           "--base-url", site.base, "--delay", "0"])
        self.assertEqual(result.code, cli.EXIT_ERROR)
        self.assertIn("już istnieje", result.err)
        self.assertIn("Nie zaczynam pobierania", result.err)
        self.assertEqual(site.state.log, [], "ani jednego żądania do portalu")
        self.assertEqual(os.listdir(self.xml_dir), [])

    @unittest.skipIf(MockSite is None, "brak atrapy portalu")
    def test_z_overwrite_all_dziala_normalnie(self):
        site = MockSite()
        self.addCleanup(site.stop)
        site.state.reset()
        with open(self.out, "w", encoding="utf-8") as handle:
            handle.write("stary wynik")
        result = call_cli(["all", "--in", self.xml_dir, "--out", self.out,
                           "--base-url", site.base, "--delay", "0",
                           "--overwrite", "--quiet"])
        self.assertIn(result.code, (cli.EXIT_OK, cli.EXIT_PARTIAL), result.err)
        self.assertGreater(len(os.listdir(self.xml_dir)), 0)


class TestZapisanejStronyLogowania(CliTestCase):
    """REGRESJA: ``build`` ma rozpoznać zapisaną stronę HTML."""

    def test_html_zamiast_xml_podpowiada_cookie(self):
        with open(os.path.join(self.xml_dir, "pakiet.xml"), "w", encoding="utf-8") as handle:
            handle.write("<!DOCTYPE html>\n<html><head><title>Zaloguj się</title></head>"
                         "<body><form><input name='login'></form></body></html>")
        with open(os.path.join(self.xml_dir, "dobry.xml"), "w", encoding="utf-8") as handle:
            handle.write(_batch_xml("auk-1", 2))
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out])
        self.assertEqual(result.code, cli.EXIT_PARTIAL)
        self.assertIn("strona HTML", result.err)
        self.assertIn("logowania", result.err)
        self.assertIn("--cookie", result.err)
        self.assertNotIn("mismatched tag", result.err)

    def test_zwykly_uszkodzony_xml_dostaje_komunikat_techniczny(self):
        """Nie każdy błąd to strona logowania — nie zgadujemy na siłę."""
        with open(os.path.join(self.xml_dir, "zly.xml"), "w", encoding="utf-8") as handle:
            handle.write("<?xml version='1.0'?><a><b></a>")
        with open(os.path.join(self.xml_dir, "dobry.xml"), "w", encoding="utf-8") as handle:
            handle.write(_batch_xml("auk-1", 2))
        result = call_cli(["build", "--in", self.xml_dir, "--out", self.out])
        self.assertEqual(result.code, cli.EXIT_PARTIAL)
        self.assertNotIn("strona logowania", result.err)


class TestZlegoCiasteczkaWCli(CliTestCase):
    """REGRESJA: ``--cookie`` ze znakiem końca linii nie wywala CLI."""

    def test_cookie_z_enterem_w_srodku_daje_komunikat_po_polsku(self):
        result = call_cli(["download", "--out", self.xml_dir,
                           "--base-url", "http://127.0.0.1:1/",
                           "--cookie", "SESSIONID=abc\nX-Zle: 1",
                           "--delay", "0", "--retries", "0"])
        self.assertEqual(result.code, cli.EXIT_ERROR)
        self.assertIn("końca linii", result.err)
        self.assertNotIn("Traceback", result.err)

    def test_user_agent_z_wstrzyknieciem_jest_odrzucany(self):
        result = call_cli(["download", "--out", self.xml_dir,
                           "--base-url", "http://127.0.0.1:1/",
                           "--user-agent", "UA\r\nX-Wstrzykniete: 1",
                           "--delay", "0", "--retries", "0"])
        self.assertEqual(result.code, cli.EXIT_ERROR)
        self.assertIn("końca linii", result.err)


class TestPomocyPoPolsku(unittest.TestCase):
    """REGRESJA: komunikaty argparse też mają być po polsku."""

    def test_naglowki_sekcji_pomocy(self):
        text = cli.build_parser().format_help()
        self.assertIn("argumenty pozycyjne", text)
        self.assertIn("opcje", text)
        self.assertNotIn("positional arguments", text)
        self.assertNotIn("show this help message and exit", text)

    def test_nieznana_podkomenda(self):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            with self.assertRaises(SystemExit) as ctx:
                cli.build_parser().parse_args(["zbuduj"])
        self.assertEqual(ctx.exception.code, cli.EXIT_USAGE)
        message = err.getvalue()
        self.assertIn("nieznana", message.lower())
        self.assertIn("download", message)
        self.assertNotIn("invalid choice", message)

    def test_literowka_w_opcji(self):
        out, err = io.StringIO(), io.StringIO()
        with contextlib.redirect_stdout(out), contextlib.redirect_stderr(err):
            with self.assertRaises(SystemExit):
                cli.build_parser().parse_args(["build", "--outt", "a.xlsx"])
        message = err.getvalue()
        self.assertIn("nierozpoznany", message.lower())
        self.assertNotIn("unrecognized arguments", message)

    def test_version_dziala_w_kazdej_podkomendzie(self):
        for argv in (["--version"], ["build", "--version"],
                     ["download", "--version"], ["all", "--version"]):
            with self.subTest(argv=argv):
                out = io.StringIO()
                with contextlib.redirect_stdout(out):
                    with self.assertRaises(SystemExit) as ctx:
                        cli.build_parser().parse_args(argv)
                self.assertEqual(ctx.exception.code, 0)
                self.assertIn(cli.PROG, out.getvalue())


if __name__ == "__main__":  # pragma: no cover
    unittest.main(verbosity=2)
