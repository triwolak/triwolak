# -*- coding: utf-8 -*-
"""Testy ADWERSARYJNE: skala i wydajność potoku ``flexit2xlsx build``.

Cel: sprawdzić, co się dzieje, gdy użytkownik faktycznie pobierze WSZYSTKIE
aukcje z portalu — setki plików XML, setki tysięcy pozycji, pliki po kilkadziesiąt
megabajtów, tysiące unikalnych kolumn.

Zasady tego pliku:

* nic nie jest zmieniane w kodzie produkcyjnym — dane testowe powstają w katalogu
  tymczasowym, a pomiary robione są w PODPROCESACH (własny ``ru_maxrss`` dziecka),
  żeby pamięć testu nie zafałszowała wyniku,
* asercje są DETERMINISTYCZNE tam, gdzie to możliwe (liczba ponownych parsowań,
  liczba wierszy w wyniku); pomiary czasu są tylko drukowane, nie są asercjami,
* testy ciężkie (300 plików x 500 pozycji, plik 100 MB, ponad milion wierszy)
  są domyślnie pominięte — włącza je zmienna środowiskowa ``FLEXIT_SKALA_PELNA=1``.

Uruchomienie:
    python3 -m unittest tests.adversarial.test_skala -v
    FLEXIT_SKALA_PELNA=1 python3 -m unittest tests.adversarial.test_skala -v
"""

import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import textwrap
import time
import unittest
import zipfile

KORZEN = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
if KORZEN not in sys.path:
    sys.path.insert(0, KORZEN)

from flexit2xlsx import cli, xlsxwrite, xmlflatten  # noqa: E402

PELNA = os.environ.get("FLEXIT_SKALA_PELNA") == "1"
POMIN_PELNE = unittest.skipUnless(
    PELNA, "test ciężki — włącz przez FLEXIT_SKALA_PELNA=1"
)

MB = 1024.0 * 1024.0


# --------------------------------------------------------------------------- #
# Narzędzia pomiarowe
# --------------------------------------------------------------------------- #


def _log(tekst):
    """Wypisuje wynik pomiaru tak, żeby był widoczny przy ``-v``."""
    sys.stderr.write("\n    [POMIAR] %s\n" % tekst)
    sys.stderr.flush()


def _uruchom_pomiar(kod, dodatkowe_srodowisko=None):
    """Uruchamia fragment kodu w podprocesie i zwraca jego pomiar (JSON).

    Do ASERCJI używamy ``tracemalloc`` (szczyt alokacji Pythona) — jest
    deterministyczny.  ``ru_maxrss`` podprocesu jest tylko informacyjny: po
    ``fork`` dziecko dziedziczy "high water mark" pamięci rodzica, więc przy
    dużym procesie testowym potrafi pokazać zawyżoną wartość startową.
    """
    naglowek = textwrap.dedent(
        """
        import json, os, resource, sys, time, tracemalloc
        sys.path.insert(0, %r)
        def szczyt_mb():
            return resource.getrusage(resource.RUSAGE_SELF).ru_maxrss / 1024.0
        _wynik = {}
        _start = time.monotonic()
        tracemalloc.start()
        """
        % KORZEN
    )
    stopka = textwrap.dedent(
        """
        _biezace, _szczyt = tracemalloc.get_traced_memory()
        tracemalloc.stop()
        _wynik["sekundy"] = time.monotonic() - _start
        _wynik["py_szczyt_mb"] = _szczyt / (1024.0 * 1024.0)
        _wynik["py_biezace_mb"] = _biezace / (1024.0 * 1024.0)
        _wynik["rss_szczyt_mb"] = szczyt_mb()
        sys.stdout.write("\\n@@POMIAR@@" + json.dumps(_wynik) + "\\n")
        """
    )
    srodowisko = dict(os.environ)
    srodowisko.setdefault("PYTHONHASHSEED", "0")
    if dodatkowe_srodowisko:
        srodowisko.update(dodatkowe_srodowisko)
    proces = subprocess.run(
        [sys.executable, "-c", naglowek + textwrap.dedent(kod) + stopka],
        cwd=KORZEN,
        capture_output=True,
        text=True,
        env=srodowisko,
    )
    if proces.returncode != 0:
        raise AssertionError(
            "podproces pomiarowy zakończył się kodem %d\nstdout:\n%s\nstderr:\n%s"
            % (proces.returncode, proces.stdout[-4000:], proces.stderr[-4000:])
        )
    for linia in proces.stdout.splitlines():
        if linia.startswith("@@POMIAR@@"):
            return json.loads(linia[len("@@POMIAR@@") :])
    raise AssertionError("brak znacznika pomiaru w wyjściu:\n%s" % proces.stdout[-4000:])


def _uruchom_cli(argumenty, srodowisko=None, limit_sekund=3600):
    """Uruchamia ``python3 -m flexit2xlsx ...`` i zwraca (kod, stdout, stderr, sekundy)."""
    env = dict(os.environ)
    if srodowisko:
        env.update(srodowisko)
    start = time.monotonic()
    proces = subprocess.run(
        [sys.executable, "-m", "flexit2xlsx"] + list(argumenty),
        cwd=KORZEN,
        capture_output=True,
        text=True,
        env=env,
        timeout=limit_sekund,
    )
    return proces.returncode, proces.stdout, proces.stderr, time.monotonic() - start


def _wiersze_arkuszy(sciezka_xlsx):
    """Zwraca listę liczb wierszy (z nagłówkiem) w kolejnych arkuszach pliku XLSX.

    Czyta strumieniowo (bez openpyxl), żeby dało się policzyć także pliki
    z ponad milionem wierszy.
    """
    wynik = []
    with zipfile.ZipFile(sciezka_xlsx) as archiwum:
        nazwy = sorted(
            (n for n in archiwum.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml$", n)),
            key=lambda n: int(re.search(r"(\d+)", n.rsplit("/", 1)[-1]).group(1)),
        )
        for nazwa in nazwy:
            licznik = 0
            with archiwum.open(nazwa) as strumien:
                for kawalek in strumien:
                    licznik += kawalek.count(b"<row ")
            wynik.append(licznik)
    return wynik


def _naglowek_arkusza(sciezka_xlsx, numer=1):
    """Zwraca teksty komórek pierwszego wiersza wskazanego arkusza."""
    with zipfile.ZipFile(sciezka_xlsx) as archiwum:
        dane = archiwum.read("xl/worksheets/sheet%d.xml" % numer).decode("utf-8", "replace")
    poczatek = dane.index("<row ")
    wiersz = dane[poczatek : dane.index("</row>", poczatek)]
    return re.findall(r"<t[^>]*>([^<]*)</t>", wiersz)


# --------------------------------------------------------------------------- #
# Generatory danych syntetycznych
# --------------------------------------------------------------------------- #


def generuj_korpus(katalog, liczba_plikow, pozycji_w_pliku):
    """Tworzy korpus XML o CZĘŚCIOWO różnych schematach (4 warianty)."""
    os.makedirs(katalog, exist_ok=True)
    for numer in range(liczba_plikow):
        wariant = numer % 4
        sciezka = os.path.join(katalog, "aukcja_%04d.xml" % numer)
        with open(sciezka, "w", encoding="utf-8") as plik:
            plik.write('<?xml version="1.0" encoding="UTF-8"?>\n')
            if wariant == 0:
                plik.write(
                    '<batch id="AUK-%04d"><auctionId>AUK-%04d</auctionId>'
                    "<title>Aukcja %d</title><closing>2026-06-18T14:00:00</closing>"
                    "<lot><items>\n" % (numer, numer, numer)
                )
                for i in range(pozycji_w_pliku):
                    plik.write(
                        '<item sku="S%05d"><model>Lenovo T%d</model><serial>0%06d</serial>'
                        "<grade>A</grade><price>1 234,%02d</price><ram>16 GB</ram>"
                        "<active>true</active><note>Zażółć gęślą jaźń %d</note></item>\n"
                        % (i, 400 + i % 100, i, i % 100, i)
                    )
                plik.write("</items></lot></batch>\n")
            elif wariant == 1:
                plik.write(
                    "<aukcja><naglowek><idAukcji>PL-%04d</idAukcji>"
                    "<data>18.06.2026</data></naglowek><pozycje>\n" % numer
                )
                for i in range(pozycji_w_pliku):
                    plik.write(
                        '<pozycja nr="%d"><model>Dell E%d</model>'
                        "<numerSeryjny>SN%06d</numerSeryjny><stan>uzywany</stan>"
                        "<cena>1.234,%02d</cena><procesor>i5-8250U</procesor></pozycja>\n"
                        % (i, 5000 + i % 900, i, i % 100)
                    )
                plik.write("</pozycje></aukcja>\n")
            elif wariant == 2:
                plik.write('<export batch="EXP-%04d"><rows>\n' % numer)
                for i in range(pozycji_w_pliku):
                    plik.write(
                        "<row><a>%d</a><b>tekst %d</b><c>2026-01-%02d</c>"
                        "<d>%0.2f</d><e>16 GB</e><f>1.2.%d</f></row>\n"
                        % (i, i, (i % 28) + 1, i / 3.0, i)
                    )
                plik.write("</rows></export>\n")
            else:
                plik.write(
                    "<catalogue><meta><auction>CAT-%04d</auction></meta><products>\n" % numer
                )
                for i in range(pozycji_w_pliku):
                    plik.write(
                        '<product id="%d"><name>Produkt %d</name><qty>%d</qty>'
                        "<tags><tag>t%d</tag><tag>u%d</tag></tags>"
                        "<spec><cpu>Ryzen %d</cpu><disk>512 GB</disk></spec></product>\n"
                        % (i, i, i % 7 + 1, i % 5, i % 3, i % 9)
                    )
                plik.write("</products></catalogue>\n")
    return liczba_plikow * pozycji_w_pliku


def generuj_jeden_duzy(sciezka, docelowe_mb):
    """Tworzy JEDEN plik XML o zadanym rozmiarze (jedna aukcja, wiele pozycji)."""
    cel = int(docelowe_mb * MB)
    licznik = 0
    with open(sciezka, "w", encoding="utf-8") as plik:
        plik.write(
            '<?xml version="1.0" encoding="UTF-8"?>\n'
            "<batch><auctionId>BIG-0001</auctionId><title>Wielka aukcja</title><items>\n"
        )
        while plik.tell() < cel:
            kawalki = []
            for _ in range(2000):
                kawalki.append(
                    '<item sku="S%07d"><model>Lenovo ThinkPad T%d</model>'
                    "<serial>0%08d</serial><grade>A</grade><price>1 234,56</price>"
                    "<ram>16 GB</ram><cpu>i5-8250U</cpu>"
                    "<opis>Zażółć gęślą jaźń %d</opis></item>\n"
                    % (licznik, 480 + licznik % 20, licznik, licznik)
                )
                licznik += 1
            plik.write("".join(kawalki))
        plik.write("</items></batch>\n")
    return licznik, os.path.getsize(sciezka)


def generuj_rzadkie_kolumny(katalog, liczba_plikow, pozycji, kolumn_na_plik):
    """Korpus, w którego UNII jest ``liczba_plikow * kolumn_na_plik`` kolumn."""
    os.makedirs(katalog, exist_ok=True)
    for numer in range(liczba_plikow):
        sciezka = os.path.join(katalog, "plik_%04d.xml" % numer)
        with open(sciezka, "w", encoding="utf-8") as plik:
            plik.write('<?xml version="1.0"?>\n<batch><auctionId>A%04d</auctionId><items>\n' % numer)
            for i in range(pozycji):
                plik.write("<item><id>%d</id>" % i)
                for k in range(kolumn_na_plik):
                    nazwa = "pole_%06d" % (numer * kolumn_na_plik + k)
                    plik.write("<%s>w%d</%s>" % (nazwa, i, nazwa))
                plik.write("</item>\n")
            plik.write("</items></batch>\n")
    return liczba_plikow * pozycji, liczba_plikow * kolumn_na_plik + 1


def generuj_dla_unify(katalog, liczba_schematow, liczba_pojedynczych, pozycji=8):
    """Odwzorowuje realia portalu: 1 plik = 1 lot; część lotów ma JEDNĄ sztukę.

    Pliki wieloelementowe mają RÓŻNE schematy (różne nazwy tagów), więc każdy
    wnosi inną wykrytą ścieżkę rekordu.  Pliki jednoelementowe nie mają nic
    powtarzalnego, więc trafiają do ujednolicania.
    """
    os.makedirs(katalog, exist_ok=True)
    for s in range(liczba_schematow):
        sciezka = os.path.join(katalog, "multi_%04d.xml" % s)
        with open(sciezka, "w", encoding="utf-8") as plik:
            plik.write('<?xml version="1.0"?>\n<root%d><auctionId>A%04d</auctionId><box%d>\n'
                       % (s, s, s))
            for i in range(pozycji):
                plik.write(
                    "<poz%d><model>M%d</model><serial>SN%06d</serial>"
                    "<price>1234,56</price></poz%d>\n" % (s, i, i, s)
                )
            plik.write("</box%d></root%d>\n" % (s, s))
    for j in range(liczba_pojedynczych):
        sciezka = os.path.join(katalog, "single_%04d.xml" % j)
        with open(sciezka, "w", encoding="utf-8") as plik:
            plik.write(
                '<?xml version="1.0"?>\n<inny><auctionId>S%04d</auctionId><zawartosc>'
                "<sztuka><model>X%d</model><serial>SN%06d</serial></sztuka>"
                "</zawartosc></inny>\n" % (j, j, j)
            )
    return liczba_schematow + liczba_pojedynczych


# --------------------------------------------------------------------------- #
# 1. Pełny potok build — czas i pamięć
# --------------------------------------------------------------------------- #


class TestSkalaPotoku(unittest.TestCase):
    """Ile trwa i ile pamięci zjada ``build`` na realistycznym korpusie."""

    @classmethod
    def setUpClass(cls):
        cls.katalog = tempfile.mkdtemp(prefix="skala-potok-")
        cls.xml_dir = os.path.join(cls.katalog, "xml")
        cls.liczba_plikow = 300 if PELNA else 40
        cls.pozycji = 500 if PELNA else 250
        cls.wierszy = generuj_korpus(cls.xml_dir, cls.liczba_plikow, cls.pozycji)
        cls.bajty = sum(
            os.path.getsize(os.path.join(cls.xml_dir, n)) for n in os.listdir(cls.xml_dir)
        )

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.katalog, ignore_errors=True)

    def _zbuduj(self, nazwa, srodowisko=None):
        wynik = os.path.join(self.katalog, nazwa)
        kod, out, err, sekundy = _uruchom_cli(
            ["build", "--in", self.xml_dir, "--out", wynik, "--overwrite", "--quiet"],
            srodowisko=srodowisko,
        )
        self.assertEqual(kod, 0, "build zakończył się kodem %d:\n%s\n%s" % (kod, out, err))
        return wynik, sekundy

    def test_build_calego_korpusu_nie_gubi_wierszy(self):
        """Wszystkie pozycje ze wszystkich plików muszą trafić do arkusza zbiorczego."""
        wynik, sekundy = self._zbuduj("wynik.xlsx")
        arkusze = _wiersze_arkuszy(wynik)
        _log(
            "korpus %d plików / %.1f MB / %d wierszy -> build %.1f s, plik wynikowy %.1f MB, "
            "arkusze (z nagłówkiem): %s"
            % (
                self.liczba_plikow,
                self.bajty / MB,
                self.wierszy,
                sekundy,
                os.path.getsize(wynik) / MB,
                arkusze,
            )
        )
        # arkusz 1 = dane, arkusz 2 = Podsumowanie
        self.assertEqual(arkusze[0] - 1, self.wierszy)
        self.assertEqual(arkusze[1] - 1, self.liczba_plikow + 1)  # + wiersz RAZEM

    def test_oba_backendy_daja_tyle_samo_wierszy_i_porownanie_czasu(self):
        """``stdlib`` i ``openpyxl`` muszą dać ten sam wynik; czas bywa różny."""
        wynik_op, czas_op = self._zbuduj("openpyxl.xlsx")
        wynik_std, czas_std = self._zbuduj(
            "stdlib.xlsx", srodowisko={"FLEXIT_XLSX_BACKEND": "stdlib"}
        )
        arkusze_op = _wiersze_arkuszy(wynik_op)
        arkusze_std = _wiersze_arkuszy(wynik_std)
        _log(
            "%d wierszy: openpyxl %.1f s, stdlib %.1f s (openpyxl wolniejszy %.1fx)"
            % (self.wierszy, czas_op, czas_std, czas_op / max(czas_std, 0.001))
        )
        self.assertEqual(arkusze_op, arkusze_std)

    def test_pamiec_nie_rosnie_liniowo_z_liczba_wierszy(self):
        """Potok JEST strumieniowy: 2x więcej wierszy NIE daje 2x więcej RAM.

        ``load_documents`` zwraca uchwyty (:class:`cli.DocHandle`): metryczka
        zostaje w pamięci, a rekordy powyżej budżetu ``MEMORY_ROW_BUDGET`` są
        porzucane i wczytywane ponownie dopiero wtedy, gdy generator wierszy
        dojdzie do danego pliku.  Szczyt pamięci spada z ``O(cały korpus)``
        do ``O(największy plik)``.

        Budżet obniżamy w podprocesie, żeby test był szybki — mierzymy sam
        mechanizm, nie wartość stałej.
        """
        maly = os.path.join(self.katalog, "maly")
        duzy = os.path.join(self.katalog, "duzy")
        # TYLE SAMO plików, dwa razy więcej wierszy w każdym — przy zapisie
        # strumieniowym zużycie pamięci nie może się przez to podwoić.
        wierszy_maly = generuj_korpus(maly, 40, 250)
        wierszy_duzy = generuj_korpus(duzy, 40, 500)

        szablon = """
            import gc, glob
            from flexit2xlsx import cli, xmlflatten
            cli.MEMORY_ROW_BUDGET = 1000          # wymuszamy tryb strumieniowy
            pliki = sorted(glob.glob(%r + "/*.xml"))
            gc.collect()
            _baza = tracemalloc.get_traced_memory()[0]
            docs, errs = cli.load_documents(pliki)
            gc.collect()
            _wynik["dane_mb"] = (tracemalloc.get_traced_memory()[0] - _baza) / (1024.0 * 1024.0)
            _wynik["wierszy"] = sum(d.count for d in docs)
            _wynik["plikow"] = len(docs)
            _wynik["w_pamieci"] = sum(1 for d in docs if d.cached)
            # pełne przejście generatora wierszy nie może zbudować listy w RAM
            kolumny = xmlflatten.merge_columns(docs)
            gc.collect()
            _przed = tracemalloc.get_traced_memory()[0]
            _ile = sum(1 for _ in cli._iter_rows(docs, kolumny,
                                                 with_auction=True, use_typing=False))
            gc.collect()
            _wynik["wierszy_z_generatora"] = _ile
            _wynik["po_generatorze_mb"] = (
                tracemalloc.get_traced_memory()[0] - _przed) / (1024.0 * 1024.0)
        """
        a = _uruchom_pomiar(szablon % maly)
        b = _uruchom_pomiar(szablon % duzy)
        self.assertEqual(a["wierszy"], wierszy_maly)
        self.assertEqual(b["wierszy"], wierszy_duzy)
        self.assertEqual(a["wierszy_z_generatora"], wierszy_maly)
        self.assertEqual(b["wierszy_z_generatora"], wierszy_duzy)

        przyrost_a = max(a["dane_mb"], 0.001)
        przyrost_b = max(b["dane_mb"], 0.001)
        stosunek = przyrost_b / przyrost_a
        _log(
            "%d wierszy -> %.2f MB ŻYWYCH obiektów po wczytaniu (%d/%d plików w RAM) | "
            "%d wierszy -> %.2f MB (%d/%d w RAM) | stosunek %.2f | "
            "po przejściu generatora: +%.2f MB / +%.2f MB"
            % (
                a["wierszy"], a["dane_mb"], a["w_pamieci"], a["plikow"],
                b["wierszy"], b["dane_mb"], b["w_pamieci"], b["plikow"],
                stosunek, a["po_generatorze_mb"], b["po_generatorze_mb"],
            )
        )
        self.assertEqual(a["w_pamieci"], 0, "korpus ponad budżet nie może zostać w RAM")
        self.assertEqual(b["w_pamieci"], 0)
        # Dwa razy więcej wierszy w tej samej liczbie plików: zużycie ma zostać
        # praktycznie takie samo (rośnie z liczbą UCHWYTÓW, nie z liczbą wierszy).
        self.assertLess(
            stosunek, 1.5,
            "pamięć rośnie z liczbą wierszy — potok przestał być strumieniowy",
        )
        na_wiersz = przyrost_b * MB / b["wierszy"]
        self.assertLess(
            na_wiersz, 100.0,
            "%.0f B pamięci na wiersz — rekordy najwyraźniej zostają w RAM" % na_wiersz,
        )
        # Generator nie odkłada wierszy w pamięci.
        for pomiar, ile in ((a, wierszy_maly), (b, wierszy_duzy)):
            self.assertLess(
                pomiar["po_generatorze_mb"], 5.0,
                "generator wierszy zostawił %.1f MB przy %d wierszach"
                % (pomiar["po_generatorze_mb"], ile),
            )

    def test_generowanie_wierszy_z_rozpoznawaniem_typow_jest_waskim_gardlem(self):
        """``coerce_value`` na KAŻDEJ komórce kosztuje wielokrotność samego składania wierszy."""
        pomiar = _uruchom_pomiar(
            """
            import glob, time
            from flexit2xlsx import cli, xmlflatten
            pliki = sorted(glob.glob(%r + "/*.xml"))
            docs, errs = cli.load_documents(pliki)
            kolumny = xmlflatten.merge_columns(docs)
            t = time.monotonic()
            n = sum(1 for _ in cli._iter_rows(docs, kolumny, with_auction=True, use_typing=True))
            _wynik["z_typami"] = time.monotonic() - t
            t = time.monotonic()
            m = sum(1 for _ in cli._iter_rows(docs, kolumny, with_auction=True, use_typing=False))
            _wynik["bez_typow"] = time.monotonic() - t
            _wynik["wierszy"] = n
            _wynik["kolumn"] = len(kolumny)
            """
            % self.xml_dir
        )
        _log(
            "składanie %d wierszy x %d kolumn: z rozpoznawaniem typów %.2f s, "
            "z --no-typing %.2f s (%.1fx szybciej)"
            % (
                pomiar["wierszy"], pomiar["kolumn"], pomiar["z_typami"], pomiar["bez_typow"],
                pomiar["z_typami"] / max(pomiar["bez_typow"], 0.001),
            )
        )
        self.assertEqual(pomiar["wierszy"], self.wierszy)


# --------------------------------------------------------------------------- #
# 2. Jeden bardzo duży plik XML
# --------------------------------------------------------------------------- #


class TestJedenDuzyPlik(unittest.TestCase):
    """Pojedynczy plik XML: ile RAM-u kosztuje 1 MB wejścia."""

    @classmethod
    def setUpClass(cls):
        cls.katalog = tempfile.mkdtemp(prefix="skala-duzy-")
        cls.xml_dir = os.path.join(cls.katalog, "xml")
        os.makedirs(cls.xml_dir)
        cls.mb = 30 if PELNA else 8
        cls.pozycji, cls.bajty = generuj_jeden_duzy(
            os.path.join(cls.xml_dir, "duzy.xml"), cls.mb
        )

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.katalog, ignore_errors=True)

    def test_pamiec_to_kilkanascie_krotnosc_rozmiaru_pliku(self):
        """Plik wczytywany jest w całości: bajty + tekst + drzewo ET + rekordy naraz."""
        pomiar = _uruchom_pomiar(
            """
            from flexit2xlsx import xmlflatten
            doc = xmlflatten.parse_file(%r)
            _wynik["rekordow"] = len(doc.records)
            _wynik["sciezka"] = doc.record_path
            """
            % os.path.join(self.xml_dir, "duzy.xml")
        )
        krotnosc = pomiar["py_szczyt_mb"] * MB / self.bajty
        _log(
            "plik %.1f MB (%d pozycji) -> szczyt %.0f MB obiektów Pythona przy samym "
            "parse_file = %.1fx rozmiaru pliku (po parsowaniu żywych %.0f MB), "
            "ru_maxrss %.0f MB, %.1f s"
            % (
                self.bajty / MB, self.pozycji, pomiar["py_szczyt_mb"], krotnosc,
                pomiar["py_biezace_mb"], pomiar["rss_szczyt_mb"], pomiar["sekundy"],
            )
        )
        self.assertEqual(pomiar["rekordow"], self.pozycji)
        self.assertGreater(
            krotnosc, 5.0,
            "parse_file zużywa mniej niż 5x rozmiaru pliku — parser stał się strumieniowy",
        )

    def test_build_duzego_pliku_nie_gubi_pozycji(self):
        wynik = os.path.join(self.katalog, "duzy.xlsx")
        kod, out, err, sekundy = _uruchom_cli(
            ["build", "--in", self.xml_dir, "--out", wynik, "--overwrite", "--quiet"]
        )
        self.assertEqual(kod, 0, err)
        arkusze = _wiersze_arkuszy(wynik)
        _log(
            "build jednego pliku %.1f MB: %.1f s, arkusze %s"
            % (self.bajty / MB, sekundy, arkusze)
        )
        self.assertEqual(arkusze[0] - 1, self.pozycji)

    @POMIN_PELNE
    def test_plik_100_mb(self):
        """Skrajny przypadek z zadania: pojedynczy plik 100 MB."""
        katalog = tempfile.mkdtemp(prefix="skala-100mb-")
        try:
            xml_dir = os.path.join(katalog, "xml")
            os.makedirs(xml_dir)
            pozycji, bajty = generuj_jeden_duzy(os.path.join(xml_dir, "duzy.xml"), 100)
            wynik = os.path.join(katalog, "duzy100.xlsx")
            kod, out, err, sekundy = _uruchom_cli(
                ["build", "--in", xml_dir, "--out", wynik, "--overwrite", "--quiet"]
            )
            self.assertEqual(kod, 0, err)
            arkusze = _wiersze_arkuszy(wynik)
            _log(
                "plik %.1f MB / %d pozycji -> build %.1f s, wynik %.1f MB, arkusze %s"
                % (bajty / MB, pozycji, sekundy, os.path.getsize(wynik) / MB, arkusze)
            )
            self.assertEqual(arkusze[0] - 1, pozycji)
        finally:
            shutil.rmtree(katalog, ignore_errors=True)


# --------------------------------------------------------------------------- #
# 3. Brak pamięci — jak kończy się program
# --------------------------------------------------------------------------- #


class TestZachowanieBezPamieci(unittest.TestCase):
    """Co widzi użytkownik, gdy RAM się skończy (stary laptop, duży korpus)."""

    def _uruchom_z_limitem(self, katalog, xml_dir, megabajty):
        """Uruchamia ``build`` w podprocesie z twardym limitem ``RLIMIT_AS``."""
        skrypt = os.path.join(katalog, "limit_%d.py" % megabajty)
        with open(skrypt, "w", encoding="utf-8") as plik:
            plik.write(
                textwrap.dedent(
                    """
                    import resource, sys
                    resource.setrlimit(resource.RLIMIT_AS, (%d, %d))
                    sys.path.insert(0, %r)
                    from flexit2xlsx.cli import main
                    sys.exit(main(sys.argv[1:]))
                    """
                    % (megabajty * 1024 * 1024, megabajty * 1024 * 1024, KORZEN)
                )
            )
        return subprocess.run(
            [
                sys.executable, skrypt, "build", "--in", xml_dir,
                "--out", os.path.join(katalog, "oom_%d.xlsx" % megabajty),
                "--overwrite", "--quiet",
            ],
            cwd=KORZEN, capture_output=True, text=True, timeout=600,
        )

    def test_ok_duzy_plik_miesci_sie_w_skromnym_limicie_pamieci(self):
        """12 MB XML-a przechodzi przy 220 MB adresowalnej pamięci.

        Zapis jest strumieniowy, więc cały korpus nie musi się mieścić w RAM.
        """
        katalog = tempfile.mkdtemp(prefix="skala-oom-")
        try:
            xml_dir = os.path.join(katalog, "xml")
            os.makedirs(xml_dir)
            pozycji, bajty = generuj_jeden_duzy(os.path.join(xml_dir, "duzy.xml"), 12)
            proces = self._uruchom_z_limitem(katalog, xml_dir, 220)
            _log(
                "plik %.1f MB (%d pozycji) przy limicie 220 MB: kod=%d"
                % (bajty / MB, pozycji, proces.returncode)
            )
            self.assertEqual(proces.returncode, 0, proces.stderr[-2000:])
            self.assertNotIn("Traceback", proces.stderr)
            self.assertTrue(os.path.exists(os.path.join(katalog, "oom_220.xlsx")))
        finally:
            shutil.rmtree(katalog, ignore_errors=True)

    def test_ok_brak_pamieci_daje_komunikat_po_polsku_a_nie_traceback(self):
        """Gdy pamięci naprawdę zabraknie, użytkownik dostaje radę, nie ślad stosu.

        ``MemoryError`` jest łapany tam, gdzie powstaje (wczytanie pliku), plik
        jest pomijany z ostrzeżeniem, a CLI kończy się ustalonym kodem wyjścia.
        """
        katalog = tempfile.mkdtemp(prefix="skala-oom2-")
        try:
            xml_dir = os.path.join(katalog, "xml")
            os.makedirs(xml_dir)
            _pozycji, bajty = generuj_jeden_duzy(os.path.join(xml_dir, "duzy.xml"), 12)
            proces = self._uruchom_z_limitem(katalog, xml_dir, 60)
            _log(
                "plik %.1f MB przy limicie 60 MB: kod=%d, MemoryError w stderr=%s, "
                "komunikat po polsku=%s"
                % (
                    bajty / MB, proces.returncode,
                    "MemoryError" in proces.stderr,
                    any(z in proces.stderr for z in ("BŁĄD", "Pomijam", "za mało")),
                )
            )
            self.assertEqual(proces.returncode, cli.EXIT_NO_DATA, proces.stderr[-2000:])
            self.assertNotIn("Traceback", proces.stderr)
            self.assertNotIn("MemoryError", proces.stderr)
            self.assertIn("Pomijam duzy.xml", proces.stderr)
            self.assertIn("za mało pamięci", proces.stderr)
        finally:
            shutil.rmtree(katalog, ignore_errors=True)

# --------------------------------------------------------------------------- #
# 4. Ujednolicanie ścieżki rekordu — koszt kwadratowy
# --------------------------------------------------------------------------- #


class TestUnifyKwadratowy(unittest.TestCase):
    """``cli.unify_record_paths`` ma koszt LINIOWY, nie ``pliki x ścieżki``.

    Naiwna wersja parsowała każdy plik bez powtórzeń raz na KAŻDĄ znalezioną
    ścieżkę rekordu (``N x M`` parsowań i ``N x M`` odczytów z dysku).  Dziś
    bajty czytamy raz na plik, kandydatów odsiewamy tanim przedfiltrem po nazwie
    znacznika, a i tak próbujemy najwyżej ``cli.UNIFY_MAX_CANDIDATES`` ścieżek.
    """

    def _policz_parsowania(self, katalog):
        """Deterministyczny licznik ponownych parsowań w ``unify_record_paths``."""
        pliki = sorted(
            os.path.join(katalog, n) for n in os.listdir(katalog) if n.endswith(".xml")
        )
        docs, errors = cli.load_documents(pliki)
        bez_powtorzen = sum(1 for d in docs if d.record_path is None)
        sciezki = len({d.record_path for d in docs if d.record_path})
        oryginal_bytes = xmlflatten.parse_bytes
        oryginal_file = xmlflatten.parse_file
        licznik = {"wywolan": 0, "bajtow": 0}

        def szpieg_bytes(data, source, **kw):
            licznik["wywolan"] += 1
            licznik["bajtow"] += len(data)
            return oryginal_bytes(data, source, **kw)

        def szpieg_file(path, **kw):
            licznik["wywolan"] += 1
            licznik["bajtow"] += os.path.getsize(path)
            return oryginal_file(path, **kw)

        xmlflatten.parse_bytes = szpieg_bytes
        xmlflatten.parse_file = szpieg_file
        try:
            start = time.monotonic()
            cli.unify_record_paths(docs)
            sekundy = time.monotonic() - start
        finally:
            xmlflatten.parse_bytes = oryginal_bytes
            xmlflatten.parse_file = oryginal_file
        return bez_powtorzen, sciezki, licznik, sekundy

    def test_liczba_ponownych_parsowan_jest_liniowa_a_nie_kwadratowa(self):
        """N plików bez powtórzeń i M ścieżek: najwyżej ``N * UNIFY_MAX_CANDIDATES``."""
        katalog = tempfile.mkdtemp(prefix="skala-unify-")
        try:
            for schematow, pojedynczych in ((20, 20), (60, 60)):
                podkatalog = os.path.join(katalog, "u%d" % schematow)
                generuj_dla_unify(podkatalog, schematow, pojedynczych)
                bez, sciezek, licznik, sekundy = self._policz_parsowania(podkatalog)
                _log(
                    "%d plików (%d bez powtórzeń, %d różnych ścieżek): unify wykonał "
                    "%d ponownych parsowań (%.2f MB) w %.2f s — kwadrat dałby %d"
                    % (
                        schematow + pojedynczych, bez, sciezek, licznik["wywolan"],
                        licznik["bajtow"] / MB, sekundy, bez * sciezek,
                    )
                )
                self.assertEqual(bez, pojedynczych)
                self.assertEqual(sciezek, schematow)
                self.assertLessEqual(licznik["wywolan"], bez * cli.UNIFY_MAX_CANDIDATES)
                self.assertLess(licznik["wywolan"], bez * sciezek)
        finally:
            shutil.rmtree(katalog, ignore_errors=True)

    def test_jeden_duzy_plik_bez_powtorzen_nie_jest_czytany_wielokrotnie(self):
        """Jeden plik bez powtarzalnego elementu czytamy raz, nie raz na schemat."""
        katalog = tempfile.mkdtemp(prefix="skala-unify-big-")
        try:
            generuj_dla_unify(katalog, 60, 0)
            sciezka = os.path.join(katalog, "raport_jednorazowy.xml")
            with open(sciezka, "w", encoding="utf-8") as plik:
                plik.write('<?xml version="1.0"?>\n<raport><auctionId>R-1</auctionId>')
                for i in range(30000):
                    plik.write("<pole_%06d>wartosc %d</pole_%06d>" % (i, i, i))
                plik.write("</raport>\n")
            rozmiar = os.path.getsize(sciezka)
            bez, sciezek, licznik, sekundy = self._policz_parsowania(katalog)
            _log(
                "1 plik %.2f MB bez powtórzeń + %d różnych schematów: unify "
                "przeparsował %.1f MB (%d parsowań, %.1fx rozmiaru pliku) w %.1f s"
                % (
                    rozmiar / MB, sciezek, licznik["bajtow"] / MB, licznik["wywolan"],
                    licznik["bajtow"] / max(rozmiar, 1), sekundy,
                )
            )
            self.assertEqual(bez, 1)
            self.assertLessEqual(licznik["wywolan"], cli.UNIFY_MAX_CANDIDATES)
            self.assertLess(licznik["bajtow"], cli.UNIFY_MAX_CANDIDATES * rozmiar + 1)
        finally:
            shutil.rmtree(katalog, ignore_errors=True)


# --------------------------------------------------------------------------- #
# 5. Bardzo dużo kolumn
# --------------------------------------------------------------------------- #


class TestDuzoKolumn(unittest.TestCase):
    """2000 unikalnych kolumn (i ponad limit Excela) — czy dane przeżywają."""

    def test_2000_unikalnych_kolumn(self):
        katalog = tempfile.mkdtemp(prefix="skala-kolumny-")
        try:
            xml_dir = os.path.join(katalog, "xml")
            wierszy, kolumn = generuj_rzadkie_kolumny(xml_dir, 100, 20 if not PELNA else 100, 20)
            wynik = os.path.join(katalog, "kolumny.xlsx")
            kod, out, err, sekundy = _uruchom_cli(
                ["build", "--in", xml_dir, "--out", wynik, "--overwrite", "--quiet"]
            )
            self.assertEqual(kod, 0, err)
            naglowek = _naglowek_arkusza(wynik, 1)
            arkusze = _wiersze_arkuszy(wynik)
            _log(
                "%d wierszy x %d kolumn danych: build %.1f s, nagłówek ma %d kolumn, "
                "plik %.1f MB, arkusze %s"
                % (
                    wierszy, kolumn, sekundy, len(naglowek),
                    os.path.getsize(wynik) / MB, arkusze,
                )
            )
            self.assertGreaterEqual(len(naglowek), 2000)
            self.assertEqual(arkusze[0] - 1, wierszy)
            self.assertEqual(naglowek[:3], ["Aukcja", "Plik", "Nr pozycji"])
        finally:
            shutil.rmtree(katalog, ignore_errors=True)

    def test_powyzej_limitu_kolumn_excela_dane_ida_do_drugiego_arkusza(self):
        """>16384 kolumn: dane nie giną, a arkusz-kontynuacja POWTARZA kolumny techniczne."""
        katalog = tempfile.mkdtemp(prefix="skala-kolumny-limit-")
        try:
            xml_dir = os.path.join(katalog, "xml")
            wierszy, kolumn = generuj_rzadkie_kolumny(xml_dir, 400, 5, 50)
            wynik = os.path.join(katalog, "duzo.xlsx")
            kod, out, err, sekundy = _uruchom_cli(
                ["build", "--in", xml_dir, "--out", wynik, "--overwrite", "--quiet"]
            )
            self.assertEqual(kod, 0, err)
            arkusze = _wiersze_arkuszy(wynik)
            naglowek1 = _naglowek_arkusza(wynik, 1)
            naglowek2 = _naglowek_arkusza(wynik, 2)
            _log(
                "%d kolumn danych -> %d arkuszy danych po %d wierszy; pierwsze kolumny "
                "arkusza-kontynuacji: %s, build %.1f s"
                % (kolumn, len(arkusze) - 1, arkusze[0] - 1, naglowek2[:3], sekundy)
            )
            self.assertIn("kolumn", err)  # ostrzeżenie na stderr
            self.assertEqual(arkusze[0] - 1, wierszy)
            self.assertEqual(arkusze[1] - 1, wierszy)
            self.assertEqual(naglowek1[:3], ["Aukcja", "Plik", "Nr pozycji"])
            self.assertEqual(naglowek2[:3], ["Aukcja", "Plik", "Nr pozycji"],
                             "kontynuacja musi dać się połączyć z pierwszym arkuszem")
        finally:
            shutil.rmtree(katalog, ignore_errors=True)


# --------------------------------------------------------------------------- #
# 6. Limit wierszy Excela
# --------------------------------------------------------------------------- #


class TestLimitWierszy(unittest.TestCase):
    """Podział na kolejne arkusze po przekroczeniu limitu wierszy — bez utraty danych."""

    def test_podzial_nie_gubi_ani_nie_duplikuje_wierszy(self):
        """Limit wstrzykiwany przez stałą modułu (tak przewiduje dokumentacja modułu)."""
        katalog = tempfile.mkdtemp(prefix="skala-wiersze-")
        oryginal = xlsxwrite.MAX_ROWS_PER_SHEET
        try:
            wierszy = 5000
            xlsxwrite.MAX_ROWS_PER_SHEET = 1000
            wynik = os.path.join(katalog, "podzial.xlsx")
            arkusz = xlsxwrite.Sheet(
                name="Wszystkie aukcje",
                columns=["Aukcja", "Nr pozycji", "model"],
                rows=(["A-1", i, "model %d" % i] for i in range(wierszy)),
            )
            xlsxwrite.write_workbook(wynik, [arkusz])
            arkusze = _wiersze_arkuszy(wynik)
            dane = sum(n - 1 for n in arkusze)
            _log(
                "limit 1000 wierszy na arkusz, %d wierszy wejścia -> %d arkuszy %s, "
                "suma wierszy danych = %d"
                % (wierszy, len(arkusze), arkusze, dane)
            )
            self.assertEqual(dane, wierszy)
        finally:
            xlsxwrite.MAX_ROWS_PER_SHEET = oryginal
            shutil.rmtree(katalog, ignore_errors=True)

    @POMIN_PELNE
    def test_ponad_milion_wierszy_realnie(self):
        """Prawdziwe przekroczenie limitu 1 048 576 wierszy (bardzo wolne)."""
        katalog = tempfile.mkdtemp(prefix="skala-milion-")
        try:
            xml_dir = os.path.join(katalog, "xml")
            wierszy = generuj_korpus(xml_dir, 2100, 500)
            wynik = os.path.join(katalog, "milion.xlsx")
            kod, out, err, sekundy = _uruchom_cli(
                ["build", "--in", xml_dir, "--out", wynik, "--overwrite", "--quiet"]
            )
            self.assertEqual(kod, 0, err)
            arkusze = _wiersze_arkuszy(wynik)
            _log(
                "%d wierszy: build %.1f s, arkusze %s, plik %.1f MB"
                % (wierszy, sekundy, arkusze, os.path.getsize(wynik) / MB)
            )
            self.assertEqual(arkusze[0] - 1 + arkusze[1] - 1, wierszy)
        finally:
            shutil.rmtree(katalog, ignore_errors=True)


# --------------------------------------------------------------------------- #
# 7. Wiele aukcji naraz (--per-auction)
# --------------------------------------------------------------------------- #


class TestWieluAukcji(unittest.TestCase):
    """Setki aukcji w jednym pliku wynikowym."""

    def test_nazwy_arkuszy_zachowuja_numer_aukcji(self):
        """Identyfikator aukcji dłuższy niż 31 znaków = nieczytelne zakładki.

        Limit Excela to 31 znaków nazwy arkusza.  Slug portalu
        ``flexit-auctions-18-06-2026-1103`` ma dokładnie 31 znaków, więc mieści
        się o włos — ale każdy dłuższy identyfikator (dopisek, dłuższy numer,
        tytuł aukcji) traci KOŃCÓWKĘ, czyli właśnie to, co odróżnia aukcje.
        Zakładki zostają wtedy rozróżnione tylko automatycznym ``(2)``, ``(3)``…
        Dane nie giną (pełne nazwy są w arkuszu ``Podsumowanie`` i w kolumnie
        ``Aukcja`` arkusza zbiorczego), ale nawigacja po zakładkach przestaje
        działać.
        """
        katalog = tempfile.mkdtemp(prefix="skala-aukcje-")
        try:
            xml_dir = os.path.join(katalog, "xml")
            os.makedirs(xml_dir)
            aukcji, pozycji = 60, 3
            for numer in range(aukcji):
                with open(os.path.join(xml_dir, "a%03d.xml" % numer), "w", encoding="utf-8") as plik:
                    plik.write(
                        '<?xml version="1.0"?><batch><auctionId>'
                        "flexit-auctions-18-06-2026-numer-%04d</auctionId><items>" % numer
                    )
                    for i in range(pozycji):
                        plik.write("<item><model>M%d</model><serial>S%d</serial></item>" % (i, i))
                    plik.write("</items></batch>")
            wynik = os.path.join(katalog, "aukcje.xlsx")
            kod, out, err, sekundy = _uruchom_cli(
                ["build", "--in", xml_dir, "--out", wynik, "--overwrite", "--quiet", "--per-auction"]
            )
            self.assertEqual(kod, 0, err)
            with zipfile.ZipFile(wynik) as archiwum:
                nazwy = re.findall(
                    r'<sheet[^>]*name="([^"]*)"',
                    archiwum.read("xl/workbook.xml").decode("utf-8"),
                )
            arkusze = _wiersze_arkuszy(wynik)
            # czytelna zakładka = kończy się czterocyfrowym numerem aukcji
            nierozroznialne = sum(
                1
                for n in nazwy
                if n.startswith("flexit-auctions-18-06-2026")
                and not re.search(r"-\d{4}$", n)
            )
            _log(
                "%d aukcji, --per-auction: %d arkuszy, wszystkie nazwy unikalne=%s, "
                "arkuszy bez czytelnego numeru aukcji=%d, przykłady: %s, build %.1f s"
                % (
                    aukcji, len(nazwy), len(set(nazwy)) == len(nazwy),
                    nierozroznialne, nazwy[3:6], sekundy,
                )
            )
            self.assertEqual(len(set(nazwy)), len(nazwy))          # brak kolizji
            self.assertEqual(arkusze[0] - 1, aukcji * pozycji)      # dane kompletne
            self.assertIn("Aukcji jest %d" % aukcji, err)           # ostrzeżenie CLI
            self.assertEqual(nierozroznialne, 0,
                             "każda zakładka ma kończyć się numerem aukcji")
        finally:
            shutil.rmtree(katalog, ignore_errors=True)


if __name__ == "__main__":
    unittest.main(verbosity=2)
