# -*- coding: utf-8 -*-
"""Testy generycznego spłaszczania XML (``flexit2xlsx.xmlflatten``).

Fixtury odwzorowują różne realistyczne kształty XML z portali aukcyjnych –
schemat flexitauctions.com nie jest znany, więc sprawdzamy zachowanie na
szerokim spektrum struktur, kodowań i przypadków złośliwych.
"""

from __future__ import annotations

import inspect
import os
import sys
import tempfile
import types
import unittest

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

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
                         {"@nr": "1", "name": "Laptop", "qty": "2", "price": "100.50"})
        self.assertIn("@nr", doc.columns)
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
        self.assertEqual(doc.records[0], {"@id": "1", "@sku": "X1", "@qty": "3"})
        self.assertEqual(doc.context["package@auction"], "AUK/2024/7")
        self.assertEqual(doc.context["package@generated"], "2024-05-01")

    def test_element_pusty_z_atrybutami_nie_tworzy_pustej_kolumny(self):
        doc = parse_bytes(XML_PACKAGE, "p.xml")
        self.assertNotIn("item", doc.columns)

    def test_plaskie_wiersze_tylko_atrybuty(self):
        doc = parse_bytes(XML_FLAT_ROWS, "f.xml")
        self.assertEqual(doc.columns, ["@a", "@b", "@c"])
        self.assertEqual([r["@a"] for r in doc.records], ["1", "4", "7"])

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
        # tresc rekordu trafia do kolumny o nazwie jego tagu
        self.assertEqual(doc.records[0], {"@a": "1", "i": "tekst"})

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
        self.assertEqual(doc.columns[-4:], ["@nr", "name", "qty", "price"])


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
        self.assertEqual(doc.records[0], {"@code": "k1", "label": "Alfa"})

    def test_bez_strip_ns_zachowuje_prefiksy(self):
        doc = parse_bytes(XML_NS, "ns.xml", strip_ns=False)
        self.assertEqual(doc.record_path, "a:auction/{urn:def}positions/{urn:def}position")
        self.assertEqual(doc.records[0]["@x:code"], "k1")
        self.assertIn("a:auction/a:auction_id", doc.context)

    def test_niezadeklarowany_prefiks_nie_wywala_parsera(self):
        # brak xmlns:x - parser wraca do trybu bez przestrzeni nazw
        dane = b"<r><x:foo a='1'>v</x:foo><x:foo a='2'>w</x:foo></r>"
        doc = parse_bytes(dane, "u.xml")
        self.assertEqual(doc.record_path, "r/foo")
        self.assertEqual(doc.records[0], {"@a": "1", "foo": "v"})
        self.assertEqual(parse_bytes(dane, "u.xml", strip_ns=False).record_path, "r/x:foo")

    def test_wiele_przestrzeni_ten_sam_lokalny_tag(self):
        data = (b"<r xmlns:p='urn:p' xmlns:q='urn:q'>"
                b"<i><p:v>1</p:v><q:v>2</q:v></i><i><p:v>3</p:v><q:v>4</q:v></i></r>")
        # przy strip_ns=True oba tagi zlewaja sie w jedna kolumne (sklejenie)
        self.assertEqual(parse_bytes(data, "n.xml").records[0], {"v": "1 | 2"})
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
        decls = b"".join(b'<!ENTITY e%d "x">' % i for i in range(200))
        with self.assertRaises(XmlParseError):
            parse_bytes(b"<!DOCTYPE d [" + decls + b"]><d/>", "e.xml")

    def test_zbyt_gleboki_dokument(self):
        data = b"<a>" + b"<b>" * 500 + b"x" + b"</b>" * 500 + b"</a>"
        with self.assertRaises(XmlParseError) as ctx:
            parse_bytes(data, "deep.xml")
        self.assertIn("zagnie", str(ctx.exception).lower())


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
        doc = parse_bytes(b"<r><i><a>x&nbsp;y</a><b>&mdash;</b><c>&nieznana;</c></i>"
                          b"<i><a>z</a></i></r>", "html.xml")
        self.assertEqual(doc.records[0]["a"], "x y")
        self.assertEqual(doc.records[0]["b"], "—")
        self.assertIsNone(doc.records[0]["c"])


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


if __name__ == "__main__":
    unittest.main()
