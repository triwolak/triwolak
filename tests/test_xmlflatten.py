# -*- coding: utf-8 -*-
"""Testy generycznego spłaszczania XML (``flexit2xlsx.xmlflatten``).

Fixtury odwzorowują różne realistyczne kształty XML z portali aukcyjnych –
schemat flexitauctions.com nie jest znany, więc sprawdzamy zachowanie na
szerokim spektrum struktur, kodowań i przypadków złośliwych.
"""

from __future__ import annotations

import contextlib
import inspect
import io
import itertools
import os
import random
import sys
import tempfile
import tracemalloc
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flexit2xlsx import xmlflatten  # noqa: E402
from flexit2xlsx.xmlflatten import (  # noqa: E402
    ParsedDoc,
    XmlParseError,
    detect_record_path,
    merge_columns,
    parse_bytes,
    parse_file,
    rows_for,
)

try:  # lxml służy wyłącznie do NIEZALEŻNEJ weryfikacji wyników
    from lxml import etree as lxml_etree
except ImportError:  # pragma: no cover
    lxml_etree = None


# ---------------------------------------------------------------------------
# Fixtury – 1..12 różnych kształtów XML
# ---------------------------------------------------------------------------

# 1) klasyczne <auction><lots><lot>
XML_LOTS = b"""<?xml version="1.0" encoding="UTF-8"?>
<auction id="A-100">
  <title>Sprzet IT z likwidacji</title>
  <seller><name>Flexit</name><city>Krakow</city></seller>
  <lots count="3">
    <lot nr="1"><name>Laptop</name><qty>2</qty><price>100.50</price></lot>
    <lot nr="2"><name>Monitor</name><qty>5</qty><price>50</price></lot>
    <lot nr="3"><name>Mysz</name><qty>10</qty><price>5</price></lot>
  </lots>
</auction>"""

# 2) <package><items><item> – dane wyłącznie w atrybutach
XML_PACKAGE = b"""<?xml version="1.0"?>
<package auction="AUK/2024/7" generated="2024-05-01">
  <items>
    <item id="1" sku="X1" qty="3"/>
    <item id="2" sku="X2" qty="4"/>
    <item id="3" sku="X3" qty="1"/>
  </items>
</package>"""

# 3) plaski <root><row> z atrybutami zamiast elementow
XML_FLAT_ROWS = (b"<root><row a=\"1\" b=\"2\" c=\"3\"/>"
                 b"<row a=\"4\" b=\"5\" c=\"6\"/>"
                 b"<row a=\"7\" b=\"8\" c=\"9\"/></root>")

# 4) przestrzenie nazw: domyslna + dwa prefiksy
XML_NS = b"""<?xml version="1.0"?>
<a:auction xmlns:a="urn:auc" xmlns="urn:def" xmlns:x="urn:extra">
  <a:auction_id>NS-1</a:auction_id>
  <positions>
    <position x:code="k1"><label>Alfa</label></position>
    <position x:code="k2"><label>Beta</label></position>
  </positions>
</a:auction>"""

# 5) gleboko zagniezdzone specyfikacje w parach klucz-wartosc
XML_SPECS = b"""<?xml version="1.0"?>
<auction>
  <auction_id>S-9</auction_id>
  <lots>
    <lot>
      <name>PC</name>
      <specs>
        <spec name="RAM" value="16 GB"/>
        <spec name="CPU" value="i7"/>
        <spec name="Dysk" value="512"/>
      </specs>
    </lot>
    <lot>
      <name>Notebook</name>
      <specs>
        <spec name="RAM" value="8 GB"/>
        <spec name="CPU" value="i5"/>
        <spec name="Kolor" value="czarny"/>
      </specs>
    </lot>
  </lots>
</auction>"""

# 5b) pary klucz-wartosc jako elementy <name>/<value> oraz tresc elementu
XML_SPECS_ELEMENTS = b"""<?xml version="1.0"?>
<doc>
  <produkt>
    <parametry>
      <parametr><nazwa>Waga</nazwa><wartosc>2 kg</wartosc></parametr>
      <parametr><nazwa>Kolor</nazwa><wartosc>bialy</wartosc></parametr>
    </parametry>
    <cechy><cecha name="Gwarancja">24 mies.</cecha></cechy>
  </produkt>
  <produkt>
    <parametry>
      <parametr><nazwa>Waga</nazwa><wartosc>3 kg</wartosc></parametr>
      <parametr><nazwa>Moc</nazwa><wartosc>500 W</wartosc></parametr>
    </parametry>
    <cechy><cecha name="Gwarancja">12 mies.</cecha></cechy>
  </produkt>
</doc>"""

# 6) dokument bez powtorzen -> jeden wiersz
XML_SINGLE = b"""<?xml version="1.0"?>
<document>
  <id>D1</id>
  <title>Pojedynczy pakiet</title>
  <meta><created>2024-01-02</created><author>Kowalski</author></meta>
</document>"""

# 7) metadane aukcji na gorze + pozycje nizej (po polsku)
XML_META = b"""<?xml version="1.0" encoding="UTF-8"?>
<aukcja numer_aukcji="2024/11">
  <termin>2024-05-01</termin>
  <organizator><nazwa>Flexit</nazwa><email>x@example.com</email></organizator>
  <pozycje>
    <pozycja><nazwa>Serwer</nazwa><ilosc>1</ilosc></pozycja>
    <pozycja><nazwa>Switch</nazwa><ilosc>2</ilosc></pozycja>
    <pozycja><nazwa>Szafa</nazwa><ilosc>3</ilosc></pozycja>
  </pozycje>
</aukcja>"""

# 8) kodowania inne niz UTF-8
_PL_DOC = ('<?xml version="1.0" encoding="{enc}"?>'
           '<lista><wpis><opis>Zażółć gęślą jaźń</opis></wpis>'
           '<wpis><opis>Ćma łódź</opis></wpis></lista>')
XML_ISO88592 = _PL_DOC.format(enc="ISO-8859-2").encode("iso-8859-2")
XML_CP1250 = _PL_DOC.format(enc="windows-1250").encode("cp1250")
XML_UTF8_BOM = b"\xef\xbb\xbf" + _PL_DOC.format(enc="UTF-8").encode("utf-8")
XML_UTF16 = _PL_DOC.format(enc="UTF-16").encode("utf-16")   # z BOM
XML_NO_DECL = ("<lista><wpis><opis>Zażółć</opis></wpis>"
               "<wpis><opis>gęślą</opis></wpis></lista>").encode("utf-8")

# 9) CDATA, encje wbudowane, puste elementy, tresc mieszana
XML_MIXED = b"""<?xml version="1.0"?>
<root>
  <entry>
    <desc>Stan <b>bardzo dobry</b> (uzywany)</desc>
    <empty/>
    <cdata><![CDATA[a & b < c > d]]></cdata>
    <ent>Sygnatura &amp; nr &lt;7&gt;</ent>
    <blank>   </blank>
  </entry>
  <entry>
    <desc>Nowy</desc>
    <empty>x</empty>
    <cdata>zwykly tekst</cdata>
    <ent>bez encji</ent>
    <blank>pelne</blank>
  </entry>
</root>"""

# 10) rekordy z zagniezdzonymi, powtarzalnymi dziecmi
XML_NESTED = b"""<?xml version="1.0"?>
<a><lots>
  <lot><id>1</id><parts><part><pn>a</pn><q>1</q></part><part><pn>b</pn><q>2</q></part></parts></lot>
  <lot><id>2</id><parts><part><pn>c</pn><q>3</q></part><part><pn>d</pn><q>4</q></part></parts></lot>
</lots></a>"""

# 11) duzo drobnych <spec> kontra kilka bogatych <item>
XML_MANY_SPECS = (b"<a><items>" + b"".join(
    b"<item><n>%d</n><s>" % i + b"".join(
        b"<spec name='k%d' value='v%d'/>" % (j, j) for j in range(12)) + b"</s></item>"
    for i in range(3)) + b"</items></a>")

# 13) prawdopodobny ksztalt eksportu "Download Batch Details" z portalu
#     (lot = pakiet, w srodku lista sztuk: model, numer seryjny, stan)
XML_BATCH = b"""<?xml version="1.0" encoding="UTF-8"?>
<BatchDetails>
  <Lot id="1103-e0c8f">
    <Title>12x Lenovo 8th-10th Gen Laptop Mix</Title>
    <Auction>flexit-auctions-18-06-2026-1103</Auction>
    <Currency>EUR</Currency>
  </Lot>
  <Units>
    <Unit>
      <Model>ThinkPad T480</Model><Serial>PF0ABCDE</Serial><Grade>B</Grade>
      <Specification>
        <Item name="CPU" value="i5-8350U"/>
        <Item name="RAM" value="16 GB"/>
      </Specification>
    </Unit>
    <Unit>
      <Model>ThinkPad T490</Model><Serial>PF0FGHIJ</Serial><Grade>A</Grade>
      <Specification>
        <Item name="CPU" value="i7-8565U"/>
        <Item name="RAM" value="8 GB"/>
      </Specification>
    </Unit>
  </Units>
</BatchDetails>"""

# 12) powtarzalne dziecko wewnatrz rekordu (tryby join/index)
XML_REPEATED_CHILD = b"""<r>
  <i><t>a</t><t>b</t><t>c</t></i>
  <i><t>d</t></i>
</r>"""


class TestDetectRecordPath(unittest.TestCase):
    """Heurystyka wyboru elementu powtarzalnego."""

    def _path(self, data: bytes) -> str:
        return parse_bytes(data, "x.xml").record_path

    def test_lots(self):
        self.assertEqual(self._path(XML_LOTS), "auction/lots/lot")

    def test_items_z_atrybutami(self):
        self.assertEqual(self._path(XML_PACKAGE), "package/items/item")

    def test_plaskie_wiersze(self):
        self.assertEqual(self._path(XML_FLAT_ROWS), "root/row")

    def test_przestrzenie_nazw(self):
        self.assertEqual(self._path(XML_NS), "auction/positions/position")

    def test_specyfikacje_nie_wygrywaja_z_lotem(self):
        # 6 elementow <spec> kontra 2 elementy <lot> - rekordem musi byc <lot>
        self.assertEqual(self._path(XML_SPECS), "auction/lots/lot")

    def test_duzo_specow_nie_wygrywa_z_item(self):
        # 36 elementow <spec> kontra 3 elementy <item>
        self.assertEqual(self._path(XML_MANY_SPECS), "a/items/item")

    def test_brak_powtorzen_daje_none(self):
        self.assertIsNone(self._path(XML_SINGLE))

    def test_polskie_nazwy(self):
        self.assertEqual(self._path(XML_META), "aukcja/pozycje/pozycja")

    def test_zagniezdzone_czesci_nie_wygrywaja(self):
        self.assertEqual(self._path(XML_NESTED), "a/lots/lot")

    def test_batch_details_z_portalu(self):
        # metadane lotu na gorze, sztuki nizej - rekordem jest <Unit>
        doc = parse_bytes(XML_BATCH, "batch.xml")
        self.assertEqual(doc.record_path, "BatchDetails/Units/Unit")
        self.assertEqual(len(doc.records), 2)
        self.assertEqual(doc.auction, "flexit-auctions-18-06-2026-1103")
        self.assertEqual(doc.context["BatchDetails/Lot@id"], "1103-e0c8f")
        self.assertEqual(doc.context["BatchDetails/Lot/Currency"], "EUR")
        self.assertEqual(doc.records[0]["Serial"], "PF0ABCDE")
        self.assertEqual(doc.records[0]["Specification/CPU"], "i5-8350U")
        self.assertEqual(doc.records[1]["Specification/RAM"], "8 GB")
        # tytul pakietu powtarza sie w kazdym wierszu
        idx = doc.columns.index("BatchDetails/Lot/Title")
        for wiersz in rows_for(doc, doc.columns):
            self.assertEqual(wiersz[idx], "12x Lenovo 8th-10th Gen Laptop Mix")

    def test_lisc_przegrywa_z_rekordem(self):
        # 4x <t> (czysty lisc) kontra 2x <i> (element ze struktura)
        self.assertEqual(self._path(XML_REPEATED_CHILD), "r/i")

    def test_dziala_na_drzewie_z_elementtree(self):
        # detect_record_path przyjmuje dowolny Element, takze z xml.etree
        import xml.etree.ElementTree as ET
        root = ET.fromstring(XML_LOTS.decode("utf-8"))
        self.assertEqual(detect_record_path(root), "auction/lots/lot")

    def test_none_i_dokument_bez_dzieci(self):
        self.assertIsNone(detect_record_path(None))
        self.assertIsNone(parse_bytes(b"<r>tekst</r>", "r.xml").record_path)

    def test_jawna_sciezka_ma_pierwszenstwo(self):
        doc = parse_bytes(XML_SPECS, "x.xml", record_path="auction/lots/lot/specs/spec")
        self.assertEqual(doc.record_path, "auction/lots/lot/specs/spec")
        self.assertEqual(len(doc.records), 6)

    def test_jawna_sciezka_jako_sufiks(self):
        doc = parse_bytes(XML_LOTS, "x.xml", record_path="lot")
        self.assertEqual(doc.record_path, "auction/lots/lot")
        self.assertEqual(len(doc.records), 3)

    def test_jawna_sciezka_nieistniejaca(self):
        with self.assertRaises(ValueError) as ctx:
            parse_bytes(XML_LOTS, "x.xml", record_path="nie_ma_takiej")
        self.assertIn("nie_ma_takiej", str(ctx.exception))

    def test_determinizm(self):
        # ta sama heurystyka na tych samych danych zawsze daje ten sam wynik
        wyniki = {parse_bytes(XML_LOTS, "x.xml").record_path for _ in range(10)}
        self.assertEqual(wyniki, {"auction/lots/lot"})


class TestFlattening(unittest.TestCase):
    """Kolumny, wartości, kontekst, atrybuty."""

    def test_kolumny_i_wartosci_lotow(self):
        doc = parse_bytes(XML_LOTS, "lots.xml")
        self.assertEqual(len(doc.records), 3)
        self.assertEqual(doc.records[0],
                         {"lot@nr": "1", "name": "Laptop", "qty": "2", "price": "100.50"})
        self.assertIn("lot@nr", doc.columns)
        self.assertIn("auction/title", doc.columns)
        # kolumny kontekstu ida przed kolumnami rekordu
        self.assertLess(doc.columns.index("auction/title"), doc.columns.index("name"))

    def test_context_zbiera_metadane_przodkow(self):
        doc = parse_bytes(XML_LOTS, "lots.xml")
        self.assertEqual(doc.context, {
            "auction@id": "A-100",
            "auction/title": "Sprzet IT z likwidacji",
            "auction/seller/name": "Flexit",
            "auction/seller/city": "Krakow",
            "auction/lots@count": "3",
        })
        # rekordy NIE zawieraja kontekstu - jest doklejany dopiero w rows_for
        self.assertNotIn("auction/title", doc.records[0])

    def test_context_nie_wciaga_tresci_rekordow(self):
        doc = parse_bytes(XML_META, "m.xml")
        self.assertNotIn("aukcja/pozycje", doc.context)
        for value in doc.context.values():
            self.assertNotIn("Serwer", str(value))

    def test_atrybuty_rekordu_i_przodkow(self):
        doc = parse_bytes(XML_PACKAGE, "p.xml")
        self.assertEqual(doc.records[0],
                         {"item@id": "1", "item@sku": "X1", "item@qty": "3"})
        self.assertEqual(doc.context["package@auction"], "AUK/2024/7")
        self.assertEqual(doc.context["package@generated"], "2024-05-01")

    def test_element_pusty_z_atrybutami_nie_tworzy_pustej_kolumny(self):
        doc = parse_bytes(XML_PACKAGE, "p.xml")
        self.assertNotIn("item", doc.columns)

    def test_plaskie_wiersze_tylko_atrybuty(self):
        doc = parse_bytes(XML_FLAT_ROWS, "f.xml")
        self.assertEqual(doc.columns, ["row@a", "row@b", "row@c"])
        self.assertEqual([r["row@a"] for r in doc.records], ["1", "4", "7"])

    def test_pary_klucz_wartosc_z_atrybutow(self):
        doc = parse_bytes(XML_SPECS, "s.xml")
        self.assertEqual(doc.records[0]["specs/RAM"], "16 GB")
        self.assertEqual(doc.records[0]["specs/CPU"], "i7")
        self.assertEqual(doc.records[1]["specs/Kolor"], "czarny")
        # kolumna z pierwszego rekordu, ktorej brak w drugim -> None w wierszu
        self.assertIsNone(doc.records[1].get("specs/Dysk"))
        self.assertIn("specs/Dysk", doc.columns)

    def test_pary_klucz_wartosc_z_elementow_i_tresci(self):
        doc = parse_bytes(XML_SPECS_ELEMENTS, "s2.xml")
        self.assertEqual(doc.record_path, "doc/produkt")
        self.assertEqual(doc.records[0]["parametry/Waga"], "2 kg")
        self.assertEqual(doc.records[1]["parametry/Moc"], "500 W")
        # <cecha name="Gwarancja">24 mies.</cecha> - wartoscia jest tresc elementu
        self.assertEqual(doc.records[0]["cechy/Gwarancja"], "24 mies.")

    def test_para_klucz_wartosc_nie_gubi_danych(self):
        # <image name=".." url=".."> nie jest jednoznaczna para - url nie moze zniknac
        data = (b"<r><i><imgs><image name='a' url='u1'/><image name='b' url='u2'/></imgs>"
                b"<x>1</x></i><i><imgs><image name='c' url='u3'/>"
                b"<image name='d' url='u4'/></imgs><x>2</x></i></r>")
        doc = parse_bytes(data, "img.xml")
        self.assertEqual(doc.records[0]["imgs/image@url"], "u1 | u2")
        self.assertEqual(doc.records[0]["imgs/image@name"], "a | b")
        self.assertNotIn("imgs/a", doc.records[0])

    def test_tresc_mieszana_cdata_encje_i_puste(self):
        doc = parse_bytes(XML_MIXED, "mix.xml")
        rec = doc.records[0]
        self.assertEqual(rec["desc"], "Stan bardzo dobry (uzywany)")   # pelny tekst
        self.assertEqual(rec["desc/b"], "bardzo dobry")                # i osobno dziecko
        self.assertIsNone(rec["empty"])                                # pusty element
        self.assertIsNone(rec["blank"])                                # same biale znaki
        self.assertEqual(rec["cdata"], "a & b < c > d")                # CDATA
        self.assertEqual(rec["ent"], "Sygnatura & nr <7>")             # encje wbudowane
        self.assertEqual(doc.records[1]["empty"], "x")

    def test_normalizacja_bialych_znakow(self):
        doc = parse_bytes(b"<r><i><a>  wiele\n\t  spacji \xc2\xa0i NBSP </a></i>"
                          b"<i><a>x</a></i></r>", "w.xml")
        self.assertEqual(doc.records[0]["a"], "wiele spacji i NBSP")

    def test_repeat_join_domyslnie(self):
        doc = parse_bytes(XML_REPEATED_CHILD, "r.xml")
        self.assertEqual(doc.records, [{"t": "a | b | c"}, {"t": "d"}])

    def test_repeat_join_wlasny_separator(self):
        doc = parse_bytes(XML_REPEATED_CHILD, "r.xml", join_sep=";")
        self.assertEqual(doc.records[0]["t"], "a;b;c")

    def test_repeat_index(self):
        doc = parse_bytes(XML_REPEATED_CHILD, "r.xml", repeat="index")
        self.assertEqual(doc.columns, ["t[1]", "t[2]", "t[3]"])
        self.assertEqual(doc.records[0], {"t[1]": "a", "t[2]": "b", "t[3]": "c"})
        # jednokrotne wystapienie tez dostaje [1], zeby trafic w te sama kolumne
        self.assertEqual(doc.records[1], {"t[1]": "d"})

    def test_repeat_niepoprawny(self):
        with self.assertRaises(ValueError):
            parse_bytes(XML_REPEATED_CHILD, "r.xml", repeat="dowolny")

    def test_zagniezdzone_powtorzenia_sa_sklejane(self):
        doc = parse_bytes(XML_NESTED, "n.xml")
        self.assertEqual(doc.records[0]["parts/part/pn"], "a | b")
        self.assertEqual(doc.records[1]["parts/part/q"], "3 | 4")

    def test_jeden_wiersz_dla_dokumentu_bez_powtorzen(self):
        doc = parse_bytes(XML_SINGLE, "s.xml")
        self.assertIsNone(doc.record_path)
        self.assertEqual(doc.context, {})
        self.assertEqual(len(doc.records), 1)
        self.assertEqual(doc.records[0], {
            "id": "D1", "title": "Pojedynczy pakiet",
            "meta/created": "2024-01-02", "meta/author": "Kowalski",
        })

    def test_komentarze_i_instrukcje_przetwarzania_sa_pomijane(self):
        dane = (b"<?xml version='1.0'?><!-- naglowek --><r><?pi x?>"
                b"<i><a>x<!--w srodku-->y</a></i><i><a>z</a></i></r>")
        doc = parse_bytes(dane, "c.xml")
        self.assertEqual(doc.records, [{"a": "xy"}, {"a": "z"}])

    def test_rekord_bedacy_lisciem_z_atrybutami_i_trescia(self):
        doc = parse_bytes(b"<r><i a='1'>tekst</i><i a='2'>inny</i></r>", "t.xml")
        # tresc rekordu trafia do kolumny o nazwie jego tagu, atrybut tez
        self.assertEqual(doc.records[0], {"i@a": "1", "i": "tekst"})

    def test_brak_kolizji_nazw_kontekstu_i_rekordu(self):
        dane = b"<a><name>AUKCJA</name><i><name>X</name></i><i><name>Y</name></i></a>"
        doc = parse_bytes(dane, "k.xml")
        self.assertEqual(doc.context, {"a/name": "AUKCJA"})
        self.assertEqual(doc.records, [{"name": "X"}, {"name": "Y"}])
        wiersze = list(rows_for(doc, doc.columns))
        self.assertEqual(wiersze, [["AUKCJA", "X"], ["AUKCJA", "Y"]])

    def test_rekordy_zagniezdzone_w_rekordach_nie_dubluja_wierszy(self):
        dane = (b"<a><lot><sub><lot><x>1</x></lot></sub></lot>"
                b"<lot><sub><lot><x>2</x></lot></sub></lot></a>")
        doc = parse_bytes(dane, "z.xml")
        self.assertEqual(doc.record_path, "a/lot")
        self.assertEqual(doc.records, [{"sub/lot/x": "1"}, {"sub/lot/x": "2"}])

    def test_dokument_z_samym_korzeniem(self):
        doc = parse_bytes(b"<r/>", "r.xml")
        self.assertIsNone(doc.record_path)
        self.assertEqual(doc.records, [{"r": None}])

    def test_kolejnosc_kolumn_wg_pierwszego_wystapienia(self):
        doc = parse_bytes(XML_LOTS, "l.xml")
        self.assertEqual(doc.columns[-4:], ["lot@nr", "name", "qty", "price"])


class TestAuctionId(unittest.TestCase):
    """Identyfikator aukcji: z XML albo z nazwy pliku."""

    def test_z_atrybutu_korzenia(self):
        self.assertEqual(parse_bytes(XML_LOTS, "plik.xml").auction, "A-100")

    def test_z_atrybutu_o_polskiej_nazwie(self):
        self.assertEqual(parse_bytes(XML_META, "plik.xml").auction, "2024/11")

    def test_z_elementu(self):
        self.assertEqual(parse_bytes(XML_SPECS, "plik.xml").auction, "S-9")

    def test_z_przestrzeni_nazw(self):
        self.assertEqual(parse_bytes(XML_NS, "plik.xml").auction, "NS-1")

    def test_awaryjnie_z_nazwy_pliku(self):
        doc = parse_bytes(XML_FLAT_ROWS, "/tmp/dane/AUKCJA-77.xml")
        self.assertEqual(doc.auction, "AUKCJA-77")

    def test_nie_bierze_id_pozycji(self):
        # <id> wewnatrz rekordu nie moze byc identyfikatorem aukcji
        doc = parse_bytes(XML_NESTED, "/tmp/aukcja_5.xml")
        self.assertEqual(doc.auction, "aukcja_5")


class TestNamespaces(unittest.TestCase):
    """Przestrzenie nazw – domyślna, prefiksowana, niezadeklarowana."""

    def test_strip_ns_domyslnie(self):
        doc = parse_bytes(XML_NS, "ns.xml")
        self.assertEqual(doc.context, {"auction/auction_id": "NS-1"})
        self.assertEqual(doc.records[0], {"position@code": "k1", "label": "Alfa"})

    def test_bez_strip_ns_zachowuje_prefiksy(self):
        doc = parse_bytes(XML_NS, "ns.xml", strip_ns=False)
        self.assertEqual(doc.record_path, "a:auction/{urn:def}positions/{urn:def}position")
        self.assertEqual(doc.records[0]["position@x:code"], "k1")
        self.assertIn("a:auction/a:auction_id", doc.context)

    def test_niezadeklarowany_prefiks_nie_wywala_parsera(self):
        # brak xmlns:x - parser wraca do trybu bez przestrzeni nazw
        dane = b"<r><x:foo a='1'>v</x:foo><x:foo a='2'>w</x:foo></r>"
        doc = parse_bytes(dane, "u.xml")
        self.assertEqual(doc.record_path, "r/foo")
        self.assertEqual(doc.records[0], {"foo@a": "1", "foo": "v"})
        self.assertEqual(parse_bytes(dane, "u.xml", strip_ns=False).record_path, "r/x:foo")

    def test_wiele_przestrzeni_ten_sam_lokalny_tag(self):
        data = (b"<r xmlns:p='urn:p' xmlns:q='urn:q'>"
                b"<i><p:v>1</p:v><q:v>2</q:v></i><i><p:v>3</p:v><q:v>4</q:v></i></r>")
        # przy strip_ns=True pierwsza przestrzen dostaje gola nazwe, druga
        # zachowuje prefiks - dwa rozne pola NIE moga trafic do jednej komorki
        with contextlib.redirect_stderr(io.StringIO()):
            doc = parse_bytes(data, "n.xml")
        self.assertEqual(doc.records[0], {"v": "1", "q:v": "2"})
        # bez strip_ns pozostaja rozdzielne
        doc = parse_bytes(data, "n.xml", strip_ns=False)
        self.assertEqual(doc.records[0], {"p:v": "1", "q:v": "2"})


class TestEncodings(unittest.TestCase):
    """Deklaracje kodowania, BOM, niespójne kodowania."""

    OCZEKIWANE = ["Zażółć gęślą jaźń",
                  "Ćma łódź"]

    def _opisy(self, data: bytes):
        doc = parse_bytes(data, "enc.xml")
        return [r["opis"] for r in doc.records]

    def test_iso_8859_2(self):
        self.assertEqual(self._opisy(XML_ISO88592), self.OCZEKIWANE)

    def test_windows_1250(self):
        self.assertEqual(self._opisy(XML_CP1250), self.OCZEKIWANE)

    def test_utf8_z_bom(self):
        self.assertEqual(self._opisy(XML_UTF8_BOM), self.OCZEKIWANE)

    def test_utf16_z_bom(self):
        self.assertEqual(self._opisy(XML_UTF16), self.OCZEKIWANE)

    def test_iso_8859_2_z_bom_utf8_niespojnym(self):
        # BOM klamie, deklaracja mowi ISO-8859-2 - dekodowanie awaryjne ratuje tresc
        self.assertEqual(self._opisy(b"\xef\xbb\xbf" + XML_ISO88592), self.OCZEKIWANE)

    def test_bez_deklaracji_domyslnie_utf8(self):
        doc = parse_bytes(XML_NO_DECL, "enc.xml")
        self.assertEqual(doc.records[0]["opis"], "Zażółć")

    def test_deklaracja_nie_jest_ignorowana(self):
        # te same bajty zinterpretowane jako UTF-8 dalyby krzaki albo blad
        self.assertNotEqual(XML_ISO88592.decode("latin-1"), XML_ISO88592.decode("iso-8859-2"))
        self.assertEqual(self._opisy(XML_ISO88592), self.OCZEKIWANE)

    def test_polskie_znaki_w_nazwach_kolumn(self):
        data = '<r><i><ilosc_sztuk>1</ilosc_sztuk><cena_brutto>2</cena_brutto></i>' \
               '<i><ilosc_sztuk>3</ilosc_sztuk><cena_brutto>4</cena_brutto></i></r>'
        doc = parse_bytes(data.encode("utf-8"), "k.xml")
        self.assertEqual(doc.columns, ["ilosc_sztuk", "cena_brutto"])


class TestSecurity(unittest.TestCase):
    """XXE i bomby encyjne – MUSZĄ być zablokowane."""

    BOMBA = b"""<?xml version="1.0"?>
<!DOCTYPE lolz [
 <!ENTITY lol "lol">
 <!ENTITY lol1 "&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;&lol;">
 <!ENTITY lol2 "&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;&lol1;">
 <!ENTITY lol3 "&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;&lol2;">
 <!ENTITY lol4 "&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;&lol3;">
]>
<lolz><i>&lol4;</i><i>&lol4;</i></lolz>"""

    def test_bomba_encyjna_billion_laughs(self):
        with self.assertRaises(XmlParseError) as ctx:
            parse_bytes(self.BOMBA, "bomba.xml")
        self.assertIn("bomb", str(ctx.exception).lower())

    def test_bomba_encyjna_kwadratowa(self):
        # jedna duza encja uzyta tysiace razy tez musi zostac zatrzymana
        data = (b'<!DOCTYPE d [<!ENTITY a "' + b"A" * 1000 + b'">]><d>'
                + b"&a;" * 6000 + b"</d>")
        with self.assertRaises(XmlParseError):
            parse_bytes(data, "kwadrat.xml")

    def test_xxe_plik_lokalny(self):
        with tempfile.TemporaryDirectory() as tmp:
            sekret = os.path.join(tmp, "sekret.txt")
            with open(sekret, "w", encoding="utf-8") as fh:
                fh.write("TAJNE-HASLO-12345")
            data = ('<?xml version="1.0"?><!DOCTYPE d [<!ENTITY xxe SYSTEM "file://%s">]>'
                    '<d><a>&xxe;</a></d>' % sekret).encode("utf-8")
            with self.assertRaises(XmlParseError) as ctx:
                parse_bytes(data, "xxe.xml")
            self.assertIn("XXE", str(ctx.exception))
            self.assertNotIn("TAJNE-HASLO", str(ctx.exception))

    def test_xxe_encja_parametryczna(self):
        data = (b'<!DOCTYPE d [<!ENTITY % zew SYSTEM "http://127.0.0.1:9/zly.dtd"> %zew;]>'
                b'<d/>')
        with self.assertRaises(XmlParseError):
            parse_bytes(data, "xxe2.xml")

    def test_zewnetrzne_dtd_nie_jest_pobierane(self):
        # DOCTYPE z SYSTEM nie moze skutkowac pobraniem zasobu; dokument parsuje sie
        # normalnie, bo encje parametryczne sa wylaczone.
        with tempfile.TemporaryDirectory() as tmp:
            dtd = os.path.join(tmp, "zly.dtd")
            with open(dtd, "w", encoding="utf-8") as fh:
                fh.write('<!ENTITY wstrzykniete "TAJNE-HASLO-12345">')
            data = ('<?xml version="1.0"?><!DOCTYPE d SYSTEM "file://%s"><d><a>ok</a></d>'
                    % dtd).encode("utf-8")
            doc = parse_bytes(data, "dtd.xml")
            self.assertEqual(doc.records[0]["a"], "ok")

    def test_bezpieczna_encja_wewnetrzna_dziala(self):
        data = (b'<!DOCTYPE d [<!ENTITY firma "Flexit sp. z o.o.">]>'
                b'<d><i><n>&firma;</n></i><i><n>x</n></i></d>')
        doc = parse_bytes(data, "e.xml")
        self.assertEqual(doc.records[0]["n"], "Flexit sp. z o.o.")

    def test_zbyt_wiele_deklaracji_encji(self):
        decls = b"".join(b'<!ENTITY e%d "x">' % i for i in range(5000))
        with self.assertRaises(XmlParseError):
            parse_bytes(b"<!DOCTYPE d [" + decls + b"]><d/>", "e.xml")

    def test_laczna_tresc_encji_jest_ograniczona(self):
        wielka = b"x" * 100_000
        decls = b"".join(b'<!ENTITY e%d "%s">' % (i, wielka) for i in range(5))
        with self.assertRaises(XmlParseError) as ctx:
            parse_bytes(b"<!DOCTYPE d [" + decls + b"]><d/>", "e.xml")
        self.assertIn("bomb", str(ctx.exception).lower())


class TestBledy(unittest.TestCase):
    """Uszkodzone wejście -> XmlParseError (podklasa ValueError)."""

    def test_xmlparseerror_jest_valueerror(self):
        self.assertTrue(issubclass(XmlParseError, ValueError))

    def test_niezamkniety_tag(self):
        with self.assertRaises(XmlParseError) as ctx:
            parse_bytes(b"<a><b></a>", "zly.xml")
        self.assertIn("zly.xml", str(ctx.exception))

    def test_pusty_dokument(self):
        for dane in (b"", b"   \n  "):
            with self.assertRaises(XmlParseError):
                parse_bytes(dane, "pusty.xml")

    def test_nie_xml(self):
        with self.assertRaises(XmlParseError):
            parse_bytes(b"to nie jest zaden xml", "s.xml")

    def test_komunikat_po_polsku_z_pozycja(self):
        with self.assertRaises(XmlParseError) as ctx:
            parse_bytes(b"<a><b></a>", "zly.xml")
        self.assertIn("wiersz", str(ctx.exception))

    def test_encje_htmlowe_sa_ratowane(self):
        # eksporty portali bywaja "HTML-owe"; &nbsp; nie moze wywalac calego pliku
        with contextlib.redirect_stderr(io.StringIO()):
            doc = parse_bytes(b"<r><i><a>x&nbsp;y</a><b>&mdash;</b><c>&nieznana;</c></i>"
                              b"<i><a>z</a></i></r>", "html.xml")
        self.assertEqual(doc.records[0]["a"], "x y")
        self.assertEqual(doc.records[0]["b"], "—")
        # encja spoza HTML5 zostaje DOSLOWNIE - nic nie znika po cichu
        self.assertEqual(doc.records[0]["c"], "&nieznana;")


class TestMergeIRows(unittest.TestCase):
    """Scalanie kolumn i generowanie wierszy."""

    def setUp(self):
        self.d1 = parse_bytes(b"<r><i><a>1</a><b>2</b></i><i><a>3</a><b>4</b></i></r>", "1.xml")
        self.d2 = parse_bytes(b"<r><i><b>5</b><c>6</c></i><i><b>7</b><c>8</c></i></r>", "2.xml")

    def test_unia_kolumn_w_kolejnosci_wystapien(self):
        self.assertEqual(merge_columns([self.d1, self.d2]), ["a", "b", "c"])
        self.assertEqual(merge_columns([self.d2, self.d1]), ["b", "c", "a"])

    def test_unia_jest_deterministyczna(self):
        docs = [self.d1, self.d2, parse_bytes(XML_LOTS, "3.xml")]
        wyniki = {tuple(merge_columns(docs)) for _ in range(20)}
        self.assertEqual(len(wyniki), 1)

    def test_bez_duplikatow(self):
        kolumny = merge_columns([self.d1, self.d1, self.d2])
        self.assertEqual(len(kolumny), len(set(kolumny)))

    def test_pusta_lista(self):
        self.assertEqual(merge_columns([]), [])

    def test_rows_for_wyrownuje_do_kolumn(self):
        kolumny = merge_columns([self.d1, self.d2])
        self.assertEqual(list(rows_for(self.d1, kolumny)), [["1", "2", None], ["3", "4", None]])
        self.assertEqual(list(rows_for(self.d2, kolumny)), [[None, "5", "6"], [None, "7", "8"]])

    def test_rows_for_powtarza_kontekst(self):
        doc = parse_bytes(XML_LOTS, "l.xml")
        wiersze = list(rows_for(doc, doc.columns))
        self.assertEqual(len(wiersze), 3)
        idx = doc.columns.index("auction@id")
        self.assertEqual([w[idx] for w in wiersze], ["A-100"] * 3)

    def test_rows_for_jest_generatorem(self):
        wynik = rows_for(self.d1, ["a"])
        self.assertIsInstance(wynik, types.GeneratorType)
        self.assertEqual(list(wynik), [["1"], ["3"]])

    def test_rows_for_dla_nieznanej_kolumny(self):
        self.assertEqual(list(rows_for(self.d1, ["nie_ma"])), [[None], [None]])

    def test_rows_for_zachowuje_kolejnosc_kolumn(self):
        self.assertEqual(list(rows_for(self.d1, ["b", "a"])), [["2", "1"], ["4", "3"]])


class TestParseFile(unittest.TestCase):
    """Wczytywanie z dysku."""

    def test_parse_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            sciezka = os.path.join(tmp, "AUKCJA-42.xml")
            with open(sciezka, "wb") as fh:
                fh.write(XML_ISO88592)
            doc = parse_file(sciezka)
            self.assertEqual(doc.source, sciezka)
            self.assertEqual(doc.auction, "AUKCJA-42")
            self.assertEqual(doc.records[0]["opis"],
                             "Zażółć gęślą jaźń")

    def test_parse_file_przekazuje_opcje(self):
        with tempfile.TemporaryDirectory() as tmp:
            sciezka = os.path.join(tmp, "a.xml")
            with open(sciezka, "wb") as fh:
                fh.write(XML_REPEATED_CHILD)
            doc = parse_file(sciezka, repeat="index", source="wlasne-zrodlo")
            self.assertEqual(doc.source, "wlasne-zrodlo")
            self.assertEqual(doc.records[0]["t[2]"], "b")


class TestKontrakt(unittest.TestCase):
    """Zgodność z INTERFACES.md – sygnatury i pola muszą zostać stabilne."""

    def test_pola_parseddoc(self):
        doc = parse_bytes(XML_LOTS, "l.xml")
        for pole in ("source", "auction", "context", "records", "record_path", "columns"):
            self.assertTrue(hasattr(doc, pole), pole)
        self.assertIsInstance(doc, ParsedDoc)
        self.assertIsInstance(doc.source, str)
        self.assertIsInstance(doc.auction, str)
        self.assertIsInstance(doc.context, dict)
        self.assertIsInstance(doc.records, list)
        self.assertIsInstance(doc.columns, list)

    def test_sygnatura_parse_bytes(self):
        sig = inspect.signature(parse_bytes)
        self.assertEqual(list(sig.parameters), ["data", "source", "strip_ns", "repeat",
                                                "join_sep", "record_path"])
        for nazwa in ("strip_ns", "repeat", "join_sep", "record_path"):
            self.assertEqual(sig.parameters[nazwa].kind, inspect.Parameter.KEYWORD_ONLY)
        self.assertIs(sig.parameters["strip_ns"].default, True)
        self.assertEqual(sig.parameters["repeat"].default, "join")
        self.assertEqual(sig.parameters["join_sep"].default, " | ")
        self.assertIsNone(sig.parameters["record_path"].default)

    def test_sygnatury_pozostalych_funkcji(self):
        self.assertEqual(list(inspect.signature(detect_record_path).parameters), ["root"])
        self.assertEqual(list(inspect.signature(merge_columns).parameters), ["docs"])
        self.assertEqual(list(inspect.signature(rows_for).parameters), ["doc", "columns"])

    def test_modul_uzywa_tylko_biblioteki_standardowej(self):
        import flexit2xlsx.xmlflatten as modul
        zrodlo = inspect.getsource(modul)
        for zakazane in ("import lxml", "import openpyxl", "import requests",
                         "from lxml", "from openpyxl"):
            self.assertNotIn(zakazane, zrodlo)


@unittest.skipIf(lxml_etree is None, "lxml niedostepny")
class TestWeryfikacjaLxml(unittest.TestCase):
    """Niezależna kontrola wyników drugim parserem (tylko w testach)."""

    def _lxml_root(self, data: bytes):
        parser = lxml_etree.XMLParser(resolve_entities=False, no_network=True)
        return lxml_etree.fromstring(data, parser)

    def test_liczba_rekordow_zgadza_sie_z_xpath(self):
        przypadki = [
            (XML_LOTS, "//lot"),
            (XML_PACKAGE, "//item"),
            (XML_FLAT_ROWS, "//row"),
            (XML_META, "//pozycja"),
            (XML_MANY_SPECS, "//item"),
        ]
        for dane, xpath in przypadki:
            with self.subTest(xpath=xpath):
                oczekiwane = len(self._lxml_root(dane).xpath(xpath))
                self.assertEqual(len(parse_bytes(dane, "x.xml").records), oczekiwane)

    def test_wartosci_zgadzaja_sie_z_xpath(self):
        root = self._lxml_root(XML_LOTS)
        nazwy = [e.text for e in root.xpath("//lot/name")]
        doc = parse_bytes(XML_LOTS, "x.xml")
        self.assertEqual([r["name"] for r in doc.records], nazwy)

    def test_zaden_tekst_z_xml_nie_ginie(self):
        # kazdy niepusty tekst z dokumentu musi pojawic sie w ktorejs komorce
        doc = parse_bytes(XML_META, "x.xml")
        komorki = " || ".join(
            str(v) for w in rows_for(doc, doc.columns) for v in w if v is not None)
        for tekst in self._lxml_root(XML_META).xpath("//text()"):
            tekst = " ".join(str(tekst).split())
            if tekst:
                self.assertIn(tekst, komorki)


# ---------------------------------------------------------------------------
# Testy REGRESJI – każdy odpowiada wadzie zgłoszonej przez testerów
# ---------------------------------------------------------------------------

BATCH_1_ZE_ZDJECIAMI = b"""<?xml version="1.0" encoding="UTF-8"?>
<batch lotHash="bbbbb" auction="flexit-auctions-18-06-2026-1103">
  <lot><title>1x Dell OptiPlex</title>
    <photos><photo>https://media/1.jpg</photo><photo>https://media/2.jpg</photo>
            <photo>https://media/3.jpg</photo><photo>https://media/4.jpg</photo></photos>
    <items>
      <item nr="1"><model>OptiPlex 7050</model><serial>DL9</serial><grade>A</grade></item>
    </items></lot>
</batch>"""


def _stderr(funkcja, *args, **kw):
    """Uruchamia funkcję i zwraca (wynik, tekst wypisany na stderr)."""
    bufor = io.StringIO()
    with contextlib.redirect_stderr(bufor):
        wynik = funkcja(*args, **kw)
    return wynik, bufor.getvalue()


class TestRegresjaWykrywaniaRekordu(unittest.TestCase):
    """WADA: zdjęcia/załączniki wygrywały z pozycją pakietu (wiersze-widma)."""

    def test_zdjecia_nie_wygrywaja_z_pozycja_pakietu(self):
        doc = parse_bytes(BATCH_1_ZE_ZDJECIAMI, "batch_bbbbb.xml")
        self.assertEqual(doc.record_path, "batch/lot/items/item")
        self.assertEqual(len(doc.records), 1)
        self.assertEqual(doc.records[0]["model"], "OptiPlex 7050")
        # zdjęcia lądują w JEDNEJ kolumnie kontekstu, sklejone separatorem
        self.assertEqual(doc.context["batch/lot/photos/photo"].count("|"), 3)
        # dane sztuki NIE mogą się dublować w kontekście
        self.assertNotIn("batch/lot/items/item/model", doc.context)

    def test_pakiet_wielosztukowy_dalej_dziala(self):
        dane = (b"<batch><lot><title>3x</title><items>"
                b"<item nr='1'><model>A</model><serial>1</serial></item>"
                b"<item nr='2'><model>B</model><serial>2</serial></item>"
                b"<item nr='3'><model>C</model><serial>3</serial></item>"
                b"</items></lot></batch>")
        doc = parse_bytes(dane, "b.xml")
        self.assertEqual(doc.record_path, "batch/lot/items/item")
        self.assertEqual(len(doc.records), 3)

    def test_kolumny_zgodne_miedzy_plikiem_1_i_3_sztukowym(self):
        trzy = (b"<batch lotHash='aaaaa'><lot><title>3x</title><items>"
                b"<item nr='1'><model>A</model><serial>1</serial><grade>B</grade></item>"
                b"<item nr='2'><model>B</model><serial>2</serial><grade>A</grade></item>"
                b"<item nr='3'><model>C</model><serial>3</serial><grade>C</grade></item>"
                b"</items></lot></batch>")
        docs = [parse_bytes(trzy, "a.xml"), parse_bytes(BATCH_1_ZE_ZDJECIAMI, "b.xml")]
        kolumny = merge_columns(docs)
        self.assertIn("model", kolumny)
        self.assertNotIn("batch/lot/items/item/model", kolumny)
        self.assertEqual(sum(len(d.records) for d in docs), 4)

    def test_brak_powtorzen_dalej_daje_jeden_wiersz(self):
        # kontrakt: gdy NIC się nie powtarza, dokument jest jednym wierszem
        self.assertIsNone(parse_bytes(XML_SINGLE, "s.xml").record_path)
        self.assertIsNone(parse_bytes(b"<pakiet/>", "p.xml").record_path)

    def test_powtarzalny_lisc_wygrywa_gdy_nie_ma_lepszego(self):
        # bez elementu o nazwie typowej dla pozycji zostaje stara heurystyka
        doc = parse_bytes(b"<root><t>a</t><t>b</t><t>c</t></root>", "t.xml")
        self.assertEqual(doc.record_path, "root/t")
        self.assertEqual(len(doc.records), 3)


class TestRegresjaParKluczWartosc(unittest.TestCase):
    """WADY: niespójne spłaszczanie cech, kolizje nazw i gubiona treść."""

    def test_rekord_z_jedna_cecha_ma_te_same_kolumny_co_z_dwiema(self):
        dane = (b"<aukcja id='A-1'>"
                b"<pozycja><sn>1</sn><opcja nazwa='Kolor'>czarny</opcja>"
                b"<opcja nazwa='Rozmiar'>M</opcja></pozycja>"
                b"<pozycja><sn>2</sn><opcja nazwa='Kolor'>bialy</opcja></pozycja>"
                b"</aukcja>")
        doc = parse_bytes(dane, "o.xml")
        self.assertEqual(doc.records[0], {"sn": "1", "Kolor": "czarny", "Rozmiar": "M"})
        self.assertEqual(doc.records[1], {"sn": "2", "Kolor": "bialy"})
        self.assertNotIn("opcja", doc.columns)
        self.assertNotIn("opcja@nazwa", doc.columns)

    def test_nazwa_cechy_kolidujaca_z_tagiem_daje_osobna_kolumne(self):
        dane = (b"<aukcja><item><model>ThinkPad X230</model>"
                b"<spec name='model'>i5-3320M</spec></item>"
                b"<item><model>T440</model><spec name='model'>i5-4300U</spec></item>"
                b"</aukcja>")
        doc = parse_bytes(dane, "k.xml")
        self.assertEqual(doc.records[0], {"model": "ThinkPad X230", "spec/model": "i5-3320M"})
        self.assertEqual(doc.records[1], {"model": "T440", "spec/model": "i5-4300U"})

    def test_klucz_z_ukosnikiem_i_malpa_jest_eskejpowany(self):
        dane = (b"<a><i><spec name='a/b'>1</spec><spec name='c@d'>2</spec></i>"
                b"<i><spec name='a/b'>3</spec><spec name='c@d'>4</spec></i></a>")
        doc = parse_bytes(dane, "e.xml")
        self.assertEqual(doc.records[0], {"a_b": "1", "c_d": "2"})

    def test_tekst_elementu_pary_nie_ginie(self):
        dane = (b"<aukcja><item><spec name='RAM' value='16 GB'>po rozbudowie</spec></item>"
                b"<item><spec name='RAM' value='8 GB'>fabryczne</spec></item></aukcja>")
        doc = parse_bytes(dane, "s.xml")
        self.assertEqual(doc.records[0]["RAM"], "16 GB")
        self.assertEqual(doc.records[0]["RAM (tekst)"], "po rozbudowie")
        self.assertEqual(doc.records[1]["RAM (tekst)"], "fabryczne")

    def test_atrybuty_elementu_klucza_nie_gina(self):
        dane = (b"<aukcja><item>"
                b"<cecha><nazwa jezyk='pl' id='C1'>Kolor</nazwa><wartosc>czarny</wartosc></cecha>"
                b"<cecha><nazwa jezyk='pl' id='C2'>Rozmiar</nazwa><wartosc>M</wartosc></cecha>"
                b"</item><item><sn>2</sn></item></aukcja>")
        doc = parse_bytes(dane, "c.xml")
        self.assertEqual(doc.records[0]["Kolor"], "czarny")
        self.assertEqual(doc.records[0]["Rozmiar"], "M")
        self.assertIn("C1", doc.records[0]["cecha/nazwa@id"])
        self.assertIn("pl", doc.records[0]["cecha/nazwa@jezyk"])

    def test_para_bez_wartosci_dalej_jest_spłaszczana_normalnie(self):
        dane = (b"<r><i><imgs><image name='a' url='u1'/><image name='b' url='u2'/></imgs>"
                b"<x>1</x></i><i><imgs><image name='c' url='u3'/>"
                b"<image name='d' url='u4'/></imgs><x>2</x></i></r>")
        doc = parse_bytes(dane, "img.xml")
        self.assertEqual(doc.records[0]["imgs/image@url"], "u1 | u2")
        self.assertNotIn("imgs/a", doc.records[0])

    def test_para_klucz_wartosc_nie_zjada_rekordow(self):
        # <value nazwa=... wartosc=...> wygląda na parę, ale w środku są rekordy
        dane = (b"<a><value nazwa='K' wartosc='W'>"
                b"<item id='1'>x</item><item id='2'>y</item></value></a>")
        doc = parse_bytes(dane, "z.xml")
        self.assertEqual(doc.record_path, "a/value/item")
        self.assertEqual(len(doc.records), 2)
        self.assertNotIn("K", doc.context)


class TestRegresjaPrzestrzeniNazw(unittest.TestCase):
    """WADA: obcinanie prefiksów zlewało dwa różne pola w jedną komórkę."""

    XML = (b"<b xmlns:dc='http://purl.org/dc/elements/1.1/' xmlns:f='http://flexit/'>"
           b"<i><dc:title>Katalogowy</dc:title><f:title>Sprzedazy</f:title></i>"
           b"<i><dc:title>A</dc:title><f:title>B</f:title></i></b>")

    def test_dwie_przestrzenie_daja_dwie_kolumny(self):
        doc, err = _stderr(parse_bytes, self.XML, "ns.xml")
        self.assertEqual(len(doc.columns), 2)
        self.assertEqual(doc.records[0]["title"], "Katalogowy")
        self.assertEqual(doc.records[0]["f:title"], "Sprzedazy")
        self.assertIn("keep-ns", err)

    def test_bez_kolizji_nazwa_zostaje_goła(self):
        dane = (b"<b xmlns:dc='http://purl.org/dc/elements/1.1/'>"
                b"<i><dc:title>A</dc:title><sn>1</sn></i>"
                b"<i><dc:title>B</dc:title><sn>2</sn></i></b>")
        doc, err = _stderr(parse_bytes, dane, "ns2.xml")
        self.assertEqual(doc.records[0], {"title": "A", "sn": "1"})
        self.assertEqual(err, "")


class TestRegresjaEncji(unittest.TestCase):
    """WADY: poprawne encje odrzucały plik, a naprawa psuła treść CDATA."""

    def test_encja_z_zaeskejpowanym_ampersandem_dziala(self):
        doc = parse_bytes(b'<!DOCTYPE a [<!ENTITY firma "Kowalski &amp; Syn">]>'
                          b"<a><i><n>&firma;</n></i><i><n>b</n></i></a>", "e.xml")
        self.assertEqual(doc.records[0]["n"], "Kowalski & Syn")

    def test_dluga_encja_dziala(self):
        dane = (b'<!DOCTYPE a [<!ENTITY op "' + b"x" * 1100 + b'">]>'
                b"<a><i><n>&op;</n></i><i><n>b</n></i></a>")
        self.assertEqual(len(parse_bytes(dane, "d.xml").records[0]["n"]), 1100)

    def test_siedemdziesiat_deklaracji_encji_dziala(self):
        decls = b"".join(b'<!ENTITY e%d "v%d">' % (i, i) for i in range(70))
        doc = parse_bytes(b"<!DOCTYPE a [" + decls + b"]>"
                          b"<a><i><n>&e0;</n></i><i><n>&e69;</n></i></a>", "w.xml")
        self.assertEqual([r["n"] for r in doc.records], ["v0", "v69"])

    def test_bomba_encyjna_dalej_zablokowana(self):
        with self.assertRaises(XmlParseError):
            parse_bytes(b"<!DOCTYPE a [<!ENTITY a 'aaaaaaaaaa'>"
                        b"<!ENTITY b '&a;&a;&a;&a;&a;&a;&a;&a;&a;&a;'>"
                        b"<!ENTITY c '&b;&b;&b;&b;&b;&b;&b;&b;&b;&b;'>]>"
                        b"<a><i><n>&c;</n></i><i><n>x</n></i></a>", "b.xml")

    def test_naprawa_encji_nie_rusza_CDATA(self):
        dane = ("<?xml version='1.0' encoding='UTF-8'?><lot><opis>Cena&nbsp;netto</opis>"
                "<items><item><sku>A1</sku>"
                "<note><![CDATA[Firma &raquo; model &kod_producenta; koniec]]></note></item>"
                "<item><sku>A2</sku><note><![CDATA[R&D 100]]></note></item>"
                "</items></lot>").encode("utf-8")
        doc, err = _stderr(parse_bytes, dane, "cdata.xml")
        self.assertEqual(doc.records[0]["note"],
                         "Firma &raquo; model &kod_producenta; koniec")
        self.assertEqual(doc.records[0]["sku"], "A1")
        self.assertEqual(doc.records[1]["note"], "R&D 100")
        self.assertIn("encje", err)      # użytkownik wie, że coś naprawiono

    def test_nieznana_encja_zostaje_doslownie_i_ostrzega(self):
        doc, err = _stderr(parse_bytes,
                           b"<a><i><n>AT&amp;T &nieznana; koniec</n></i>"
                           b"<i><n>b</n></i></a>", "u.xml")
        self.assertEqual(doc.records[0]["n"], "AT&T &nieznana; koniec")
        self.assertIn("encje", err)


class TestRegresjaKodowan(unittest.TestCase):
    """WADY: ciche krzaki przy błędnej deklaracji i przy zgadywaniu kodowania."""

    SZABLON = ("<?xml version='1.0' encoding='%s'?>"
               "<a><i><n>zażółć gęślą</n></i><i><n>b</n></i></a>")

    def test_utf8_z_bledna_deklaracja_iso(self):
        dane = (self.SZABLON % "ISO-8859-2").encode("utf-8")
        doc, err = _stderr(parse_bytes, dane, "k.xml")
        self.assertEqual(doc.records[0]["n"], "zażółć gęślą")
        self.assertIn("UTF-8", err)

    def test_iso88592_bez_deklaracji(self):
        dane = "<a><i><n>zażółć gęślą</n></i><i><n>b</n></i></a>".encode("iso-8859-2")
        doc, err = _stderr(parse_bytes, dane, "x.xml")
        self.assertEqual(doc.records[0]["n"], "zażółć gęślą")
        self.assertIn("iso-8859-2", err)

    def test_cp1250_bez_deklaracji(self):
        dane = "<a><i><n>zażółć gęślą</n></i><i><n>b</n></i></a>".encode("cp1250")
        doc, err = _stderr(parse_bytes, dane, "x.xml")
        self.assertEqual(doc.records[0]["n"], "zażółć gęślą")
        self.assertIn("cp1250", err)

    def test_nieznana_nazwa_kodowania_ostrzega(self):
        dane = ("<?xml version='1.0' encoding='x-mac-central-europe'?>"
                "<a><i><n>zazolc</n></i><i><n>b</n></i></a>").encode("mac-latin2")
        doc, err = _stderr(parse_bytes, dane, "m.xml")
        self.assertEqual(doc.records[0]["n"], "zazolc")
        self.assertIn("nieznana nazwa kodowania", err)

    def test_deklaracje_zgodne_z_trescia_nie_ostrzegaja(self):
        for kodowanie in ("UTF-8", "ISO-8859-2", "windows-1250"):
            with self.subTest(kodowanie=kodowanie):
                dane = (self.SZABLON % kodowanie).encode(
                    "utf-8" if kodowanie == "UTF-8" else kodowanie)
                doc, err = _stderr(parse_bytes, dane, "z.xml")
                self.assertEqual(doc.records[0]["n"], "zażółć gęślą")
                self.assertEqual(err, "")


class TestRegresjaTresciRodzica(unittest.TestCase):
    """WADA: tekst elementu-rodzica rekordów znikał bez śladu."""

    def test_tekst_rodzica_rekordow_trafia_do_kontekstu(self):
        dane = ("<aukcja id='A-1'>Sprzet uzywany, odbior osobisty."
                "<pozycja><nazwa>Laptop</nazwa></pozycja>"
                "<pozycja><nazwa>Monitor</nazwa></pozycja></aukcja>").encode("utf-8")
        doc = parse_bytes(dane, "m.xml")
        self.assertEqual(doc.context["aukcja"], "Sprzet uzywany, odbior osobisty.")
        self.assertEqual(doc.context["aukcja@id"], "A-1")

    def test_tekst_przodka_rekordow_nie_wciaga_tresci_pozycji(self):
        dane = (b"<a><lot>Opis lotu<items><item><m>X</m></item>"
                b"<item><m>Y</m></item></items></lot></a>")
        doc = parse_bytes(dane, "p.xml")
        self.assertEqual(doc.context["a/lot"], "Opis lotu")
        self.assertNotIn("X", str(doc.context))

    def test_fuzzing_strukturalny_nie_gubi_wartosci(self):
        """Losowe dokumenty: KAŻDA wartość musi trafić do komórki albo nagłówka."""
        rnd = random.Random(7)
        tagi = ["pozycja", "spec", "opcja", "cecha", "name", "value", "nazwa", "lot", "x"]
        atrybuty = ["id", "name", "nazwa", "value", "wartosc", "typ"]

        def buduj(glebokosc, licznik):
            tag = rnd.choice(tagi)
            attrs = {rnd.choice(atrybuty): "V%05d" % next(licznik)
                     for _ in range(rnd.randint(0, 2))}
            dzieci, tekst = [], ""
            if glebokosc > 0 and rnd.random() < 0.7:
                dzieci = [buduj(glebokosc - 1, licznik) for _ in range(rnd.randint(1, 3))]
                if rnd.random() < 0.2:
                    tekst = "V%05d" % next(licznik)
            else:
                tekst = "V%05d" % next(licznik)
            return (tag, attrs, tekst, dzieci)

        def rysuj(wezel):
            tag, attrs, tekst, dzieci = wezel
            opis = "".join(' %s="%s"' % (k, v) for k, v in attrs.items())
            srodek = tekst + "".join(rysuj(k) for k in dzieci)
            if not srodek:
                return "<%s%s/>" % (tag, opis)
            return "<%s%s>%s</%s>" % (tag, opis, srodek, tag)

        def wartosci(wezel, zbior):
            tag, attrs, tekst, dzieci = wezel
            zbior.update(attrs.values())
            if tekst:
                zbior.add(tekst)
            for kid in dzieci:
                wartosci(kid, zbior)
            return zbior

        for numer in range(200):
            licznik = itertools.count(1)
            korzen = ("aukcja", {"id": "A%04d" % numer}, "",
                      [buduj(3, licznik) for _ in range(rnd.randint(2, 5))])
            xml = "<?xml version='1.0' encoding='UTF-8'?>" + rysuj(korzen)
            doc = parse_bytes(xml.encode("utf-8"), "fuzz.xml")
            blob = "\n".join(
                [str(v) for v in doc.context.values()]
                + [str(v) for r in doc.records for v in r.values()]
                + list(doc.context.keys())
                + [k for r in doc.records for k in r.keys()])
            brakujace = [v for v in wartosci(korzen, set()) if v not in blob]
            self.assertEqual(brakujace, [], "dokument %d gubi wartości:\n%s"
                             % (numer, xml[:400]))


class TestRegresjaZagniezdzenia(unittest.TestCase):
    """WADA: zagnieżdżenie ponad limit odrzucało CAŁY plik."""

    def test_gleboka_galaz_jest_obcinana_a_plik_wczytany(self):
        srodek = ("".join("<n%d>" % i for i in range(500)) + "X"
                  + "".join("</n%d>" % i for i in reversed(range(500))))
        dane = ("<root><item><tytul>Lot 1</tytul><sn>SN1</sn>%s</item>"
                "<item><tytul>Lot 2</tytul><sn>SN2</sn>%s</item></root>"
                % (srodek, srodek)).encode("utf-8")
        doc, err = _stderr(parse_bytes, dane, "deep.xml")
        self.assertEqual(len(doc.records), 2)
        self.assertEqual(doc.records[0]["tytul"], "Lot 1")
        self.assertEqual(doc.records[1]["sn"], "SN2")
        self.assertIn("obcięto", err)

    def test_plytki_dokument_nie_ostrzega(self):
        doc, err = _stderr(parse_bytes, XML_LOTS, "l.xml")
        self.assertEqual(err, "")
        self.assertEqual(len(doc.records), 3)


class TestRegresjaPamieci(unittest.TestCase):
    """WADA: poprawny plik 191 kB wyczerpywał 900 MB RAM (kwadratowy _scan)."""

    @staticmethod
    def _gleboki(glebokosc: int, liscie: int) -> bytes:
        czesci = ["<?xml version='1.0'?>"]
        czesci += ["<n%d>" % i for i in range(glebokosc)]
        czesci += ["<f%d>v</f%d>" % (j, j) for j in range(liscie)]
        czesci += ["</n%d>" % i for i in reversed(range(glebokosc))]
        return "".join(czesci).encode("utf-8")

    def test_pamiec_jest_proporcjonalna_do_rozmiaru(self):
        dane = self._gleboki(120, 1500)
        tracemalloc.start()
        try:
            parse_bytes(dane, "bomba.xml")
            _, szczyt = tracemalloc.get_traced_memory()
        finally:
            tracemalloc.stop()
        krotnosc = szczyt / len(dane)
        self.assertLess(krotnosc, 300,
                        "plik %d B zajął %.1f MB (%.0f-krotność rozmiaru)"
                        % (len(dane), szczyt / 1e6, krotnosc))

    def test_limit_liczby_sciezek(self):
        poprzedni = xmlflatten._MAX_PATHS
        xmlflatten._MAX_PATHS = 50
        self.addCleanup(setattr, xmlflatten, "_MAX_PATHS", poprzedni)
        dane = (b"<a>" + b"".join(b"<t%d>v</t%d>" % (i, i) for i in range(200))
                + b"</a>")
        with self.assertRaises(XmlParseError) as ctx:
            parse_bytes(dane, "szeroki.xml")
        self.assertIn("record-path", str(ctx.exception))

    def test_limit_liczby_elementow(self):
        poprzedni = xmlflatten._MAX_ELEMENTS
        xmlflatten._MAX_ELEMENTS = 50
        self.addCleanup(setattr, xmlflatten, "_MAX_ELEMENTS", poprzedni)
        dane = b"<a>" + b"<i><x>v</x></i>" * 100 + b"</a>"
        with self.assertRaises(XmlParseError):
            parse_bytes(dane, "duzy.xml")


class TestRegresjaTrybStrumieniowy(unittest.TestCase):
    """Duże pliki idą przez dwa przebiegi – wynik musi być IDENTYCZNY."""

    PROBKI = [XML_LOTS, XML_PACKAGE, XML_FLAT_ROWS, XML_NS, XML_SPECS,
              XML_SPECS_ELEMENTS, XML_SINGLE, XML_META, XML_MIXED, XML_NESTED,
              XML_MANY_SPECS, XML_BATCH, XML_REPEATED_CHILD,
              BATCH_1_ZE_ZDJECIAMI]

    def test_oba_tryby_daja_ten_sam_wynik(self):
        for numer, dane in enumerate(self.PROBKI):
            tekst = xmlflatten._decode(dane, "x.xml")
            for repeat in ("join", "index"):
                for strip in (True, False):
                    with self.subTest(numer=numer, repeat=repeat, strip_ns=strip):
                        w_pamieci = xmlflatten._parse_in_memory(
                            tekst, "x.xml", strip, repeat, " | ", None)
                        strumieniowo = xmlflatten._parse_streaming(
                            tekst, "x.xml", strip, repeat, " | ", None)
                        self.assertEqual(w_pamieci, strumieniowo)

    def test_jawna_sciezka_rekordu_w_obu_trybach(self):
        for wskazana in ("lot", "auction/lots/lot", "auction"):
            with self.subTest(record_path=wskazana):
                tekst = xmlflatten._decode(XML_LOTS, "x.xml")
                w_pamieci = xmlflatten._parse_in_memory(
                    tekst, "x.xml", True, "join", " | ", wskazana)
                strumieniowo = xmlflatten._parse_streaming(
                    tekst, "x.xml", True, "join", " | ", wskazana)
                self.assertEqual(w_pamieci, strumieniowo)
        # rekordem jest sam korzeń -> kontekst pusty, bez dublowania wartości
        doc = parse_bytes(XML_LOTS, "x.xml", record_path="auction")
        self.assertEqual(doc.context, {})
        self.assertEqual(len(doc.records), 1)

    def test_publiczne_api_w_trybie_strumieniowym(self):
        poprzedni = xmlflatten._STREAM_MIN_CHARS
        xmlflatten._STREAM_MIN_CHARS = 1
        self.addCleanup(setattr, xmlflatten, "_STREAM_MIN_CHARS", poprzedni)
        doc = parse_bytes(XML_LOTS, "l.xml")
        self.assertEqual(doc.record_path, "auction/lots/lot")
        self.assertEqual(len(doc.records), 3)
        self.assertEqual(doc.context["auction/title"], "Sprzet IT z likwidacji")
        self.assertEqual(doc.records[0]["lot@nr"], "1")
        # jawna ścieżka rekordu też działa bez drzewa
        doc = parse_bytes(XML_SPECS, "s.xml", record_path="spec")
        self.assertEqual(len(doc.records), 6)

    def test_duzy_plik_idzie_strumieniowo_i_ma_komplet_wierszy(self):
        rekord = ("<item nr='%d'><model>T480</model><serial>SN%08d</serial>"
                  "<grade>B</grade><opis>Sprzet powystawowy, stan dobry</opis></item>")
        dane = ("<?xml version='1.0' encoding='UTF-8'?><batch auction='A-1'>"
                "<lot><title>Mix</title><items>"
                + "".join(rekord % (i, i) for i in range(70000))
                + "</items></lot></batch>").encode("utf-8")
        self.assertGreater(len(dane), xmlflatten._STREAM_MIN_CHARS)
        doc = parse_bytes(dane, "duzy.xml")
        self.assertEqual(doc.record_path, "batch/lot/items/item")
        self.assertEqual(len(doc.records), 70000)
        self.assertEqual(doc.records[0]["serial"], "SN00000000")
        self.assertEqual(doc.records[-1]["item@nr"], "69999")
        self.assertEqual(doc.context, {"batch@auction": "A-1", "batch/lot/title": "Mix"})


if __name__ == "__main__":
    unittest.main()
