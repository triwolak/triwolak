# -*- coding: utf-8 -*-
"""Generyczne spłaszczanie dowolnego XML-a do wierszy arkusza.

Moduł nie zna schematu portalu aukcyjnego i nie może go zakładać. Cała logika
opiera się więc na heurystykach strukturalnych opisanych niżej. Używa wyłącznie
biblioteki standardowej.

Heurystyka wykrywania elementu powtarzalnego (``detect_record_path``)
--------------------------------------------------------------------
Dla każdej ścieżki elementu (np. ``auction/lots/lot``) zbieramy:

* ``n``  – liczbę wystąpień elementu o tej ścieżce w całym dokumencie,
* ``f``  – liczbę RÓŻNYCH ścieżek pól-liści (elementy liściowe + atrybuty)
  spotkanych wewnątrz elementów o tej ścieżce (suma zbiorów po wystąpieniach),
* premie nazewnicze i strukturalne.

Kandydatem jest każda ścieżka o ``n >= 2``. Wynik punktowy:

``score = (log2(n) + 1) * f * premia_nazwy * premia_struktury / (1 + 0,08 * głębokość)``

* ``log2(n)`` zamiast ``n`` celowo TŁUMI przewagę elementów bardzo licznych,
  ale ubogich w treść (klasyczny przypadek: 10 lotów × 12 ``spec`` = 120
  wystąpień ``spec``; bez tłumienia ``spec`` wygrałby z ``lot``).
* ``f`` realizuje wymaganie „ważona liczbą pól-liści” – prawdziwy rekord ma
  wiele różnych pól, a pomocniczy węzeł (np. ``kw``, ``spec``) ma ich 1–2.
  Liczymy ścieżki RÓŻNE, więc powtórzenia wewnątrz rekordu nie zawyżają wyniku
  rodzica w nieskończoność.
* premia nazwy: ×3 dla nazw typowych dla rekordu (item, lot, product,
  position, pozycja, asset, device, row, record, entry, artikel, part,
  komponent, sprzet…), ×0,4 dla nazw typowych dla par klucz–wartość i
  załączników (spec, attribute, property, param, field, image, keyword…).
* premia struktury: ×1,6 gdy rodzic wygląda na liczbę mnogą dziecka
  (``lots/lot``, ``items/item``, ``pozycje/pozycja``) i ×1,2 gdy rodzic zawiera
  wyłącznie dzieci o tym samym tagu (jednorodny kontener).
* ×0,35 dla czystych liści (element bez dzieci i bez atrybutów) – taki
  „rekord” dałby arkusz z jedną bezimienną kolumną.
* delikatna kara za głębokość rozstrzyga remisy na korzyść płytszych ścieżek.

Gdy żadna ścieżka nie powtarza się (``n < 2`` wszędzie) – ``detect_record_path``
zwraca ``None``, a cały dokument staje się JEDNYM wierszem.

Wyjątek ratunkowy: jeśli JEDYNE powtarzalne ścieżki są ubogie (czysty liść albo
mniej niż dwa różne pola – zdjęcia, tagi, linki), a w dokumencie istnieje
element o nazwie typowej dla pozycji i z prawdziwą treścią, to rekordem zostaje
ON – choćby wystąpił raz. Bez tego pakiet z jedną sztuką i czterema zdjęciami
dawał cztery wiersze-widma zamiast jednego wiersza z danymi sztuki.

Pozostałe reguły
----------------
* Klucz kolumny rekordu = ścieżka WZGLĘDNA wobec elementu rekordu
  (``specs/spec``), atrybuty jako ``sciezka@atrybut``; atrybut samego rekordu
  dostaje nazwę tego elementu (``item@nr``), żeby nagłówek dało się przeczytać.
* Wszystko poza poddrzewami rekordów trafia do ``context`` z kluczem
  BEZWZGLĘDNYM (z nazwą korzenia, np. ``auction/seller/name``). Dzięki temu
  metadane aukcji nigdy nie kolidują z kolumnami rekordu. Własny tekst
  elementu-rodzica rekordów też tam trafia (pod kluczem jego ścieżki) – żadna
  wartość z XML-a nie może zniknąć bez śladu.
* Pary klucz–wartość (``<spec name="RAM" value="16 GB"/>``) zamieniamy na
  kolumnę ``specs/RAM``. Decyzję podejmujemy RAZ dla całego dokumentu (dla
  ścieżki elementu), a nie osobno w każdym rekordzie – inaczej pozycja z jedną
  cechą miałaby inne kolumny niż pozycja z dwiema. Klucz pochodzi z DANYCH,
  więc: ``/`` i ``@`` w nim eskejpujemy, a przy kolizji z prawdziwym tagiem
  (``<model>`` obok ``<spec name="model">``) chowamy go w przestrzeni nazw
  pojemnika (``spec/model``). Niezużyta treść i atrybuty pary trafiają do
  dodatkowych kolumn (``RAM (tekst)``, ``cecha/nazwa@jezyk``).
* Bezpieczeństwo: własny parser na ``xml.parsers.expat`` blokuje encje
  zewnętrzne (XXE) i encje rekurencyjne (bomba encyjna) – ``xml.etree`` sam z
  siebie rozwija bomby encyjne, dlatego nie używamy go do parsowania.
  Blokujemy odwołania ``&nazwa;``/``%nazwa;`` w treści encji (a nie każdy
  ampersand), pilnujemy ŁĄCZNEJ długości encji i wielkości dokumentu
  (liczba elementów, liczba różnych ścieżek). Zbyt głęboka gałąź jest
  OBCINANA z ostrzeżeniem, a nie odrzucana razem z całym plikiem.
* Wydajność: dokumenty powyżej :data:`_STREAM_MIN_CHARS` czytamy dwoma
  przebiegami (statystyki bez drzewa, potem wiersze spłaszczane w locie), więc
  w pamięci nie leży naraz pełne drzewo i komplet wierszy.
* Ostrzeżenia (stderr): naprawione encje, zgadnięte/sprzeczne kodowanie,
  kolizje nazw po obcięciu przestrzeni nazw, obcięte gałęzie. Nic nie zmienia
  się w danych po cichu.
"""

from __future__ import annotations

import codecs
import itertools
import math
import os
import re
import sys
import xml.etree.ElementTree as ET
import xml.parsers.expat as expat
from dataclasses import dataclass, field
from html.entities import html5 as _HTML5_ENTITIES
from typing import Any, Dict, Iterator, List, Optional, Sequence, Set, Tuple

__all__ = [
    "XmlParseError",
    "ParsedDoc",
    "parse_bytes",
    "parse_file",
    "detect_record_path",
    "merge_columns",
    "rows_for",
]

# --- limity bezpieczeństwa ---------------------------------------------------
_MAX_DEPTH = 200            # maksymalne zagnieżdżenie XML (ochrona stosu);
                            # głębsze gałęzie są OBCINANE, nie odrzucamy pliku
_MAX_ENTITY_DECLS = 4096    # maksymalna liczba deklaracji encji
_MAX_ENTITY_TOTAL = 256 * 1024   # łączna długość treści wszystkich encji
_TEXT_AMPLIFICATION = 20    # ile razy tekst może przerosnąć źródło
_MIN_TEXT_LIMIT = 4 * 1024 * 1024   # dolna granica limitu tekstu
_MAX_ELEMENTS = 2_000_000   # łączna liczba elementów w dokumencie
_MAX_PATHS = 100_000        # łączna liczba RÓŻNYCH ścieżek elementów

#: Minimalna liczba różnych pól, żeby ścieżka mogła uchodzić za rekord.
#: Element bez pól (czysty liść, np. ``<photo>url</photo>``) albo z jednym
#: polem daje arkusz z jedną, bezimienną kolumną – to prawie zawsze pomyłka.
_MIN_RECORD_FIELDS = 2

# --- słowniki heurystyk ------------------------------------------------------
_PREFERRED_RECORD_NAMES = frozenset({
    "item", "lot", "product", "position", "pozycja", "asset", "device", "row",
    "record", "entry", "artikel", "part", "komponent", "sprzet", "sprzęt",
    "produkt", "przedmiot", "towar", "aukcja", "auction", "pakiet", "package",
    # nazwy typowe dla eksportu „Batch Details” (jedna sztuka w pakiecie)
    "unit", "line", "lineitem", "line_item", "sztuka", "egzemplarz",
    "urzadzenie", "urządzenie",
})
_DISCOURAGED_RECORD_NAMES = frozenset({
    "spec", "specs", "specification", "attribute", "attr", "property", "prop",
    "param", "parameter", "parametr", "cecha", "atrybut", "wlasciwosc",
    "właściwość", "feature", "field", "pole", "tag", "keyword", "kw", "image",
    "img", "photo", "zdjecie", "zdjęcie", "file", "plik", "link", "url",
    "category", "kategoria", "note", "uwaga", "value", "wartosc", "wartość",
})
_KV_ELEMENT_NAMES = frozenset({
    "spec", "specs", "specification", "attribute", "attr", "property", "prop",
    "param", "parameter", "parametr", "cecha", "atrybut", "wlasciwosc",
    "właściwość", "feature", "field", "pole", "detail", "szczegol", "szczegół",
    "info", "characteristic",
})
_KEY_NAMES = frozenset({
    "name", "nazwa", "key", "klucz", "label", "etykieta", "cecha", "parametr",
    "param", "attribute", "atrybut", "property", "wlasciwosc", "właściwość",
})
_VALUE_NAMES = frozenset({
    "value", "val", "wartosc", "wartość", "content", "text", "tekst", "tresc",
    "treść",
})

# Białe znaki + NBSP, spacje typograficzne i BOM w środku tekstu.
_WS_RE = re.compile("[\\s\\u00a0\\u1680\\u2000-\\u200a\\u2028\\u2029"
                    "\\u202f\\u205f\\u3000\\ufeff]+")
_DECL_ENC_RE = re.compile(
    r"""<\?xml[^>]*?encoding\s*=\s*['"]([A-Za-z0-9_.:+-]+)['"]""", re.IGNORECASE
)
_ENTITY_REF_RE = re.compile(r"&([A-Za-z][A-Za-z0-9._-]*);")
_BUILTIN_ENTITIES = frozenset({"amp", "lt", "gt", "quot", "apos"})
#: Fragmenty dokumentu, w których „&nazwa;” NIE jest encją (treść dosłowna).
_CDATA_SPLIT_RE = re.compile(r"(<!\[CDATA\[.*?\]\]>|<!--.*?-->|<\?.*?\?>)", re.S)
#: Odwołanie do encji ogólnej wewnątrz treści encji (rekurencja = bomba).
_ENTITY_SELF_REF_RE = re.compile(r"&(?!(?:amp|lt|gt|quot|apos);|#)")
#: Odwołanie do encji parametrycznej wewnątrz treści encji.
_PARAM_REF_RE = re.compile(r"%[A-Za-z_][A-Za-z0-9._:-]*;")

_BOMS: Tuple[Tuple[bytes, str], ...] = (
    (codecs.BOM_UTF32_LE, "utf-32-le"),
    (codecs.BOM_UTF32_BE, "utf-32-be"),
    (codecs.BOM_UTF8, "utf-8"),
    (codecs.BOM_UTF16_LE, "utf-16-le"),
    (codecs.BOM_UTF16_BE, "utf-16-be"),
)
_FALLBACK_ENCODINGS = ("utf-8", "cp1250", "iso-8859-2", "latin-1")

#: Bajty, które w ISO-8859-2 są polskimi literami (Ą Ś Ź ą ś), a w cp1250
#: rzadko spotykaną interpunkcją (ˇ ¦ ¬ ± ¶) – rozstrzygają zgadywanie.
_ISO88592_HINT_BYTES = frozenset({0xA1, 0xA6, 0xAC, 0xB1, 0xB6})


def _warn(message: str) -> None:
    """Wypisuje ostrzeżenie na stderr; nigdy nie przerywa przetwarzania."""
    try:
        sys.stderr.write("UWAGA: %s\n" % message)
    except Exception:  # pragma: no cover - stderr może być zamknięty
        pass


class XmlParseError(ValueError):
    """Błąd parsowania XML (uszkodzony dokument, XXE, bomba encyjna)."""


@dataclass
class ParsedDoc:
    """Wynik spłaszczenia jednego pliku XML."""

    source: str                                  # ścieżka/nazwa pliku źródłowego
    auction: str                                 # identyfikator aukcji
    context: Dict[str, Any] = field(default_factory=dict)   # pola z poziomu aukcji
    records: List[Dict[str, Any]] = field(default_factory=list)  # wiersze
    record_path: Optional[str] = None            # wykryta ścieżka rekordu
    columns: List[str] = field(default_factory=list)  # kolejność kolumn


# ---------------------------------------------------------------------------
# Pomocnicze operacje na tekście i nazwach
# ---------------------------------------------------------------------------

def _norm_ws(text: Optional[str]) -> str:
    """Zwija białe znaki (także NBSP) i przycina; zachowuje polskie znaki."""
    if not text:
        return ""
    return _WS_RE.sub(" ", text).strip()


def _local_name(tag: str) -> str:
    """Zwraca nazwę lokalną: ``{uri}tag`` -> ``tag``, ``pre:tag`` -> ``tag``."""
    if not isinstance(tag, str):
        return ""
    if "}" in tag:
        tag = tag.rsplit("}", 1)[1]
    if ":" in tag:
        tag = tag.rsplit(":", 1)[1]
    return tag


def _is_element(node: Any) -> bool:
    """True dla zwykłych elementów (odsiewa komentarze i instrukcje PI)."""
    return isinstance(getattr(node, "tag", None), str)


def _join_path(prefix: str, name: str) -> str:
    return name if not prefix else prefix + "/" + name


# ---------------------------------------------------------------------------
# Dekodowanie: BOM + deklaracja kodowania w XML
# ---------------------------------------------------------------------------

def _sniff_encoding(data: bytes) -> Tuple[bytes, Optional[str], Optional[str]]:
    """Zwraca (dane bez BOM, kodowanie z BOM, kodowanie z deklaracji XML)."""
    body = data
    bom_encoding: Optional[str] = None
    for bom, enc in _BOMS:
        if data.startswith(bom):
            body = data[len(bom):]
            bom_encoding = enc
            break
    else:
        # UTF-16 bez BOM rozpoznajemy po zerowych bajtach wokół "<"
        if data.startswith(b"<\x00"):
            bom_encoding = "utf-16-le"
        elif data.startswith(b"\x00<"):
            bom_encoding = "utf-16-be"
    match = _DECL_ENC_RE.search(body[:1024].decode("latin-1", "replace"))
    return body, bom_encoding, (match.group(1) if match else None)


def _is_single_byte_encoding(name: str) -> bool:
    """True dla kodowań jednobajtowych (ISO-8859-*, cp1250, latin-1…)."""
    try:
        info = codecs.lookup(name)
    except (LookupError, ValueError):
        return False
    return not info.name.startswith(("utf", "u8", "u16", "u32"))


def _decodes_as_utf8(body: bytes) -> bool:
    """True, gdy bajty są poprawnym UTF-8 i zawierają znaki spoza ASCII."""
    if not any(byte >= 0x80 for byte in body[:65536]):
        return False        # czysty ASCII – wybór kodowania niczego nie zmienia
    try:
        body.decode("utf-8")
    except UnicodeDecodeError:
        return False
    return True


def _guess_legacy_encoding(body: bytes) -> str:
    """Rozstrzyga cp1250 vs ISO-8859-2 dla bajtów bez BOM i bez deklaracji.

    W ISO-8859-2 zakres 0x80-0x9F to znaki sterujące C1 – prawdziwy tekst ich
    nie zawiera, więc ich obecność wskazuje na cp1250 (siedzą tam m.in. ``ś``
    i ``ź``).  W drugą stronę bajty 0xA1/0xA6/0xAC/0xB1/0xB6 to w ISO-8859-2
    polskie litery, a w cp1250 rzadka interpunkcja.
    """
    if any(0x80 <= byte <= 0x9F for byte in body):
        return "cp1250"
    if any(byte in _ISO88592_HINT_BYTES for byte in body):
        return "iso-8859-2"
    return "cp1250"


def _decode(data: bytes, source: str) -> str:
    """Dekoduje bajty wg BOM/deklaracji, z awaryjnym łańcuchem kodowań.

    Dwie sytuacje wymagają ostrożności (obie spotykane w eksportach portali):

    * deklaracja kłamie – plik jest zapisany w UTF-8, a w nagłówku stoi np.
      ``ISO-8859-2``.  Gdy bajty są POPRAWNYM UTF-8 ze znakami spoza ASCII,
      wierzymy zawartości, nie deklaracji (i mówimy o tym na stderr);
    * brak jakiejkolwiek wskazówki – wtedy po nieudanej próbie UTF-8
      rozstrzygamy cp1250 vs ISO-8859-2 heurystyką bajtową i ostrzegamy,
      że kodowanie zostało ZGADNIĘTE.
    """
    if not data or not data.strip():
        raise XmlParseError("Pusty dokument XML: %s" % source)
    body, bom_encoding, declared = _sniff_encoding(data)

    # (kodowanie, ostrzeżenie do wypisania, gdy to właśnie ono zadziała)
    candidates: List[Tuple[str, Optional[str]]] = []
    if bom_encoding:
        candidates.append((bom_encoding, None))
    if declared:
        try:
            codecs.lookup(declared)
        except (LookupError, ValueError):
            _warn("%s: nieznana nazwa kodowania %r w deklaracji XML — "
                  "rozpoznaję kodowanie automatycznie" % (source, declared))
        else:
            if (not bom_encoding and _is_single_byte_encoding(declared)
                    and _decodes_as_utf8(body)):
                candidates.append((
                    "utf-8",
                    "%s: deklaracja mówi o kodowaniu %r, ale zawartość jest "
                    "poprawnym UTF-8 — czytam jako UTF-8" % (source, declared),
                ))
            candidates.append((declared, None))

    if not candidates:
        guess = _guess_legacy_encoding(body)
        powod = ("nierozpoznana deklaracja kodowania" if declared
                 else "brak deklaracji kodowania")
        candidates.append(("utf-8", None))
        if guess != "utf-8":
            candidates.append((guess, "%s: %s — zgaduję %s"
                                      % (source, powod, guess)))
    candidates.extend((encoding, None) for encoding in _FALLBACK_ENCODINGS)

    for candidate, warning in candidates:
        try:
            text = body.decode(candidate)
        except (LookupError, UnicodeDecodeError, ValueError):
            continue
        if warning:
            _warn(warning)
        break
    else:  # pragma: no cover - latin-1 nigdy nie zawodzi
        text = body.decode("latin-1", "replace")
    text = text.lstrip("\ufeff")
    if not text.startswith("<"):
        text = text.lstrip()
    if not text:
        raise XmlParseError("Pusty dokument XML: %s" % source)
    return text


def _repair_entities(text: str, declared: Set[str], source: str) -> str:
    """Zamienia nieznane encje (np. ``&nbsp;``) na tekst – ratunek dla eksportów.

    Uruchamiane WYŁĄCZNIE po błędzie „undefined entity”, żeby nie modyfikować
    poprawnych dokumentów. Encje wbudowane i zadeklarowane w dokumencie zostają
    nietknięte.

    Dwie reguły chronią treść przed cichą zmianą:

    * sekcje ``<![CDATA[...]]>``, komentarze i instrukcje przetwarzania są
      POMIJANE – tam ``&nazwa;`` to zwykły tekst, a nie encja;
    * encja spoza HTML5 nie jest usuwana, tylko zapisywana dosłownie
      (``&amp;nazwa;``), więc żaden fragment treści nie znika.

    O każdej naprawie informujemy na stderr.
    """
    naprawione = [0]
    doslowne = [0]

    def _sub(match: "re.Match[str]") -> str:
        name = match.group(1)
        if name in _BUILTIN_ENTITIES or name in declared:
            return match.group(0)
        repl = _HTML5_ENTITIES.get(name + ";") or _HTML5_ENTITIES.get(name)
        if repl is None:
            doslowne[0] += 1
            return "&amp;%s;" % name          # zachowujemy dosłowną postać
        naprawione[0] += 1
        return (
            repl.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        )

    parts = _CDATA_SPLIT_RE.split(text)
    for index in range(0, len(parts), 2):     # nieparzyste = CDATA/komentarz/PI
        parts[index] = _ENTITY_REF_RE.sub(_sub, parts[index])
    if naprawione[0] or doslowne[0]:
        _warn("%s: napotkano encje spoza XML-a — zamieniono na tekst: %d, "
              "zachowano dosłownie: %d (treść CDATA pozostała nietknięta)"
              % (source, naprawione[0], doslowne[0]))
    return "".join(parts)


# ---------------------------------------------------------------------------
# Bezpieczny parser XML (expat + ET.TreeBuilder)
# ---------------------------------------------------------------------------

_URI_TOKEN_RE = re.compile(r"[^0-9A-Za-z]+")
_URI_SKIP_TOKENS = frozenset({
    "http", "https", "urn", "www", "com", "org", "net", "pl", "io", "xml",
    "ns", "schema", "schemas", "xsd", "spec", "elements", "index",
})


def _uri_hint(uri: str) -> str:
    """Krótka, czytelna etykieta przestrzeni nazw bez zadeklarowanego prefiksu."""
    tokens = [token for token in _URI_TOKEN_RE.split(uri) if token]
    for token in reversed(tokens):
        if token.lower() not in _URI_SKIP_TOKENS and not token.isdigit():
            return token
    return tokens[-1] if tokens else "ns"


def _run_parser(text: str, source: str, strip_ns: bool, ns_aware: bool,
                declared: Set[str], builder: Any = None, warn: bool = True) -> Any:
    """Przetwarza dokument własnym parserem expat z blokadą encji.

    ``builder`` to obiekt zgodny z :class:`xml.etree.ElementTree.TreeBuilder`
    (``start``/``data``/``end``/``close``). Domyślnie budujemy pełne drzewo,
    ale przy dużych plikach podstawiamy budowniczych, którzy zbierają same
    statystyki albo spłaszczają rekordy w locie – dzięki temu 100-megabajtowy
    XML nie musi w całości mieszkać w pamięci.
    """
    if builder is None:
        builder = ET.TreeBuilder()
    parser = expat.ParserCreate("utf-8", "}" if ns_aware else None)
    parser.buffer_text = True
    try:
        parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    except (AttributeError, expat.error):  # pragma: no cover
        pass

    prefixes: Dict[str, Optional[str]] = {}
    cache: Dict[str, str] = {}
    owners: Dict[str, str] = {}   # nazwa lokalna -> URI, który ją zajął
    state = {"depth": 0, "entities": 0, "chars": 0, "entity_chars": 0,
             "truncated": 0, "ns_clash": 0}
    # Tekst po rozwinięciu encji nie może wielokrotnie przerastać źródła –
    # to tania ochrona przed „kwadratową” bombą encyjną.
    text_limit = max(_MIN_TEXT_LIMIT, _TEXT_AMPLIFICATION * len(text))

    def fixname(name: str) -> str:
        cached = cache.get(name)
        if cached is not None:
            return cached
        fixed = name
        if "}" in name:
            uri, local = name.split("}", 1)
            if strip_ns:
                # Obcinanie prefiksów nie może ZLAĆ dwóch różnych pól:
                # pierwsza przestrzeń nazw dostaje gołą nazwę, każda kolejna
                # zachowuje prefiks (``dc:title`` obok ``title``).
                owner = owners.setdefault(local, uri)
                if owner == uri:
                    fixed = local
                else:
                    state["ns_clash"] += 1
                    fixed = "%s:%s" % (prefixes.get(uri) or _uri_hint(uri), local)
            else:
                prefix = prefixes.get(uri)
                fixed = "%s:%s" % (prefix, local) if prefix else "{%s}%s" % (uri, local)
        elif strip_ns and ":" in name:
            fixed = name.split(":", 1)[1]
            owners.setdefault(fixed, "")
        elif strip_ns:
            owners.setdefault(name, "")
        cache[name] = fixed
        return fixed

    def start(tag: str, attrs: Dict[str, str]) -> None:
        state["depth"] += 1
        if state["depth"] > _MAX_DEPTH:
            # Zamiast odrzucać CAŁY plik obcinamy zbyt głęboką gałąź –
            # płytsze pola (tytuł, numery seryjne) da się jeszcze odczytać.
            state["truncated"] += 1
            return
        builder.start(fixname(tag), {fixname(k): v for k, v in attrs.items()})

    def end(tag: str) -> None:
        depth = state["depth"]
        state["depth"] = depth - 1
        if depth > _MAX_DEPTH:
            return
        builder.end(fixname(tag))

    def ns_decl(prefix: Optional[str], uri: str) -> None:
        if uri not in prefixes:
            prefixes[uri] = prefix

    def entity_decl(name, is_parameter, value, base, system_id, public_id,
                    notation_name) -> None:
        if system_id or public_id or notation_name:
            raise XmlParseError(
                "Zablokowano encję zewnętrzną '%s' (ochrona przed XXE): %s"
                % (name, source)
            )
        if is_parameter:
            raise XmlParseError(
                "Zablokowano encję parametryczną '%s' w DTD: %s" % (name, source)
            )
        state["entities"] += 1
        if state["entities"] > _MAX_ENTITY_DECLS:
            raise XmlParseError(
                "Zbyt wiele deklaracji encji (limit %d) – możliwa bomba encyjna: %s"
                % (_MAX_ENTITY_DECLS, source)
            )
        if value is None:
            raise XmlParseError(
                "Zablokowano encję '%s' bez treści (możliwa encja zewnętrzna): %s"
                % (name, source)
            )
        state["entity_chars"] += len(value)
        if state["entity_chars"] > _MAX_ENTITY_TOTAL:
            raise XmlParseError(
                "Łączna treść encji przekroczyła %d znaków – możliwa bomba "
                "encyjna: %s" % (_MAX_ENTITY_TOTAL, source)
            )
        # Blokujemy tylko PRAWDZIWE odwołania (&nazwa; / %nazwa;), bo to one
        # dają rekurencję. Zwykły, zaeskejpowany ampersand („Kowalski &amp; Syn”)
        # jest poprawną treścią i musi się rozwinąć.
        if _ENTITY_SELF_REF_RE.search(value) or _PARAM_REF_RE.search(value):
            raise XmlParseError(
                "Zablokowano rekurencyjną encję '%s' "
                "(ochrona przed bombą encyjną): %s" % (name, source)
            )
        declared.add(name)

    def external_ref(context, base, system_id, public_id) -> int:
        raise XmlParseError(
            "Zablokowano odwołanie do zasobu zewnętrznego (ochrona przed XXE): %s"
            % source
        )

    def data(chunk: str) -> None:
        state["chars"] += len(chunk)
        if state["chars"] > text_limit:
            raise XmlParseError(
                "Tekst po rozwinięciu encji przekroczył limit %d znaków "
                "(ochrona przed bombą encyjną): %s" % (text_limit, source)
            )
        if state["depth"] > _MAX_DEPTH:
            return          # treść obciętej gałęzi
        builder.data(chunk)

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = data
    parser.EntityDeclHandler = entity_decl
    parser.ExternalEntityRefHandler = external_ref
    if ns_aware:
        parser.StartNamespaceDeclHandler = ns_decl

    parser.Parse(text.encode("utf-8"), True)
    result = builder.close()
    if result is None:  # pragma: no cover - expat zgłosiłby błąd wcześniej
        raise XmlParseError("Dokument XML nie zawiera elementu głównego: %s" % source)
    if warn and state["truncated"]:
        _warn("%s: dokument zagnieżdżony głębiej niż %d poziomów — obcięto %d "
              "zbyt głębokich elementów (płytsze pola wczytano)"
              % (source, _MAX_DEPTH, state["truncated"]))
    if warn and state["ns_clash"]:
        _warn("%s: po obcięciu przestrzeni nazw kolidowały nazwy pól — kolejne "
              "przestrzenie zachowały prefiks (użyj --keep-ns, żeby zobaczyć "
              "pełne nazwy)" % source)
    return result


def _parse_xml(text: str, source: str, strip_ns: bool,
               builder_factory=None) -> Tuple[Any, str, bool]:
    """Parsuje z dwiema próbami naprawy: nieznany prefiks i nieznana encja.

    Zwraca ``(wynik budowniczego, użyty tekst, czy_z_przestrzeniami_nazw)`` –
    dwa ostatnie pola pozwalają uruchomić DRUGI przebieg (spłaszczanie
    strumieniowe) bez powtarzania prób naprawy.
    """
    ns_aware = True
    repaired = False
    declared: Set[str] = set()
    for _ in range(4):
        try:
            builder = builder_factory() if builder_factory is not None else None
            return (_run_parser(text, source, strip_ns, ns_aware, declared, builder),
                    text, ns_aware)
        except expat.ExpatError as exc:
            code = getattr(exc, "code", None)
            if ns_aware and code == expat.errors.codes[expat.errors.XML_ERROR_UNBOUND_PREFIX]:
                # Niezadeklarowany prefiks – parsujemy bez obsługi przestrzeni nazw.
                ns_aware = False
                continue
            if (not repaired
                    and code == expat.errors.codes[expat.errors.XML_ERROR_UNDEFINED_ENTITY]):
                # Eksporty portali bywają „HTML-owe” (&nbsp;) – jedna próba naprawy.
                repaired = True
                text = _repair_entities(text, declared, source)
                continue
            raise XmlParseError(
                "Uszkodzony XML w %s: %s (wiersz %s, kolumna %s)"
                % (source, expat.ErrorString(code) if code else exc,
                   getattr(exc, "lineno", "?"), getattr(exc, "offset", "?"))
            ) from exc
        except ValueError as exc:
            if isinstance(exc, XmlParseError):
                raise
            raise XmlParseError("Uszkodzony XML w %s: %s" % (source, exc)) from exc
    raise XmlParseError("Nie udało się sparsować XML: %s" % source)  # pragma: no cover


# ---------------------------------------------------------------------------
# Spłaszczanie pojedynczego poddrzewa
# ---------------------------------------------------------------------------

class _Plan:
    """Decyzje podejmowane RAZ dla CAŁEGO dokumentu.

    Bez tego ten sam element byłby spłaszczany różnie w różnych rekordach:
    ``<opcja nazwa="Kolor">`` w pozycji z dwiema cechami trafiał do kolumny
    ``Kolor``, a w pozycji z jedną cechą do kolumn ``opcja``/``opcja@nazwa``.
    Filtr po kolumnie „Kolor” po cichu gubił taki wiersz.

    * ``kv_paths``   – bezwzględne ścieżki elementów, które w CAŁYM dokumencie
      wyglądają na pary klucz–wartość (powtarzają się gdziekolwiek albo mają
      nazwę typową dla pojemnika na cechy);
    * ``child_tags`` – tagi dzieci każdej ścieżki; służą do wykrycia kolizji
      klucza pochodzącego z DANYCH z prawdziwym tagiem;
    * ``record_path`` – ścieżka rekordu, której poddrzewa nie wchodzą do
      kontekstu (zamiast zbioru ``id()`` wszystkich rekordów, który przy
      500 tys. pozycji sam ważył dziesiątki MB).
    """

    __slots__ = ("kv_paths", "child_tags", "record_path")

    def __init__(self, kv_paths: Set[str], child_tags: Dict[str, Set[str]],
                 record_path: Optional[str]) -> None:
        self.kv_paths = kv_paths
        self.child_tags = child_tags
        self.record_path = record_path


def _build_plan(root: ET.Element, record_path: Optional[str]) -> _Plan:
    """Jeden przebieg po drzewie: kandydaci na pary k–v i tagi dzieci."""
    kv_paths: Set[str] = set()
    child_tags: Dict[str, Set[str]] = {}

    def visit(elem: ET.Element, path: str, depth: int) -> None:
        kids = [c for c in elem if _is_element(c)]
        if not kids:
            return
        tags = child_tags.get(path)
        if tags is None:
            tags = child_tags[path] = set()
        counts: Dict[str, int] = {}
        for kid in kids:
            tags.add(kid.tag)
            counts[kid.tag] = counts.get(kid.tag, 0) + 1
        for tag, number in counts.items():
            if number >= 2 or _local_name(tag).lower() in _KV_ELEMENT_NAMES:
                kv_paths.add(_join_path(path, tag))
        if depth < _MAX_DEPTH:
            for kid in kids:
                visit(kid, _join_path(path, kid.tag), depth + 1)

    visit(root, root.tag, 0)
    return _Plan(kv_paths, child_tags, record_path)


def _kv_key(text: str) -> str:
    """Klucz z DANYCH nie może udawać ścieżki ani atrybutu."""
    return text.replace("/", "_").replace("@", "_")


def _kv_info(elem: ET.Element):
    """Rozpoznaje parę klucz–wartość, np. ``<spec name="RAM">16 GB</spec>``.

    Zwraca ``(klucz, wartość, dodatki, reszta_tekstu, nieużyte_dzieci)`` albo
    ``None``, gdy element nie jest parą (wtedy spłaszczamy go normalnie).

    „Dodatki” to niezużyte atrybuty – zarówno samego elementu, jak i elementów
    pełniących rolę klucza/wartości (``<nazwa jezyk="pl">``). Razem z „resztą
    tekstu” gwarantują, że zamiana na parę k–v NIE gubi żadnej informacji.
    """
    kids = [c for c in elem if _is_element(c)]

    key_text = ""
    key_attr: Optional[str] = None
    key_kid: Optional[ET.Element] = None
    for attr, val in elem.attrib.items():
        if _local_name(attr).lower() in _KEY_NAMES:
            key_text = _norm_ws(val)
            if key_text:
                key_attr = attr
                break
    if not key_text:
        for kid in kids:
            if (_local_name(kid.tag).lower() in _KEY_NAMES
                    and not any(_is_element(g) for g in kid)):
                key_text = _norm_ws("".join(kid.itertext()))
                if key_text:
                    key_kid = kid
                    break
    if not key_text:
        return None

    value: Any = None
    value_attr: Optional[str] = None
    value_kid: Optional[ET.Element] = None
    for attr, val in elem.attrib.items():
        if attr == key_attr:
            continue
        if _local_name(attr).lower() in _VALUE_NAMES:
            value_attr = attr
            value = _norm_ws(val) or None
            break
    if value_attr is None:
        for kid in kids:
            if kid is key_kid:
                continue
            if (_local_name(kid.tag).lower() in _VALUE_NAMES
                    and not any(_is_element(g) for g in kid)):
                value_kid = kid
                value = _norm_ws("".join(kid.itertext())) or None
                break

    own = _norm_ws("".join(t for t in itertools.chain(
        [elem.text], (k.tail for k in kids)) if t))
    leftover: Optional[str] = None
    if value_attr is None and value_kid is None:
        if not own:
            return None       # klucz bez wartości – to nie jest para
        value = own
    else:
        leftover = own or None

    extras: List[Tuple[str, Any]] = []
    for attr, val in elem.attrib.items():
        if attr == key_attr or attr == value_attr:
            continue
        extras.append(("@" + attr, _norm_ws(val) or None))
    for kid in (key_kid, value_kid):
        if kid is None:
            continue
        for attr, val in kid.attrib.items():
            extras.append(("/%s@%s" % (kid.tag, attr), _norm_ws(val) or None))

    unused = [kid for kid in kids if kid is not key_kid and kid is not value_kid]
    return key_text, value, extras, leftover, unused


def _add(out: Dict[str, List[Any]], key: str, value: Any) -> None:
    out.setdefault(key, []).append(value)


def _collect(elem: ET.Element, prefix: str, out: Dict[str, List[Any]],
             plan: _Plan, abs_path: str, depth: int = 0) -> None:
    """Rekurencyjnie zbiera wartości poddrzewa do mapy ``klucz -> [wartości]``."""
    if depth > _MAX_DEPTH:
        return
    # Atrybuty elementu rekordu dostają jego nazwę (``item@nr``), a nie samo
    # „@nr” – tak opisuje kolumny README i tak są czytelne dla człowieka.
    self_key = prefix or _local_name(elem.tag)
    for attr, val in elem.attrib.items():
        _add(out, "%s@%s" % (self_key, attr), _norm_ws(val) or None)

    children = [c for c in elem if _is_element(c)]
    if not children:
        text = _norm_ws("".join(elem.itertext()))
        # Element bez treści, ale z atrybutami (np. <item id="1"/>) nie tworzy
        # osobnej pustej kolumny – cała informacja jest już w atrybutach.
        if text or not elem.attrib:
            _add(out, self_key, text or None)
        return

    record_path = plan.record_path
    kept = ([c for c in children
             if _join_path(abs_path, c.tag) != record_path]
            if record_path else children)
    # Czy GŁĘBIEJ w tym poddrzewie siedzą rekordy? Jeśli tak, treść mieszana
    # NIE może brać całego ``itertext`` – wciągnęłaby wartości pozycji do
    # kontekstu i powielała je w każdym wierszu.
    ponad_rekordami = bool(record_path) and record_path.startswith(abs_path + "/")

    # Treść mieszana: element ma dzieci ORAZ własny tekst.
    own = _norm_ws("".join(t for t in itertools.chain(
        [elem.text], (c.tail for c in children)) if t))
    if own:
        if ponad_rekordami:
            # Bierzemy WYŁĄCZNIE własny tekst – żeby go nie zgubić, ale też
            # żeby nie skopiować do kontekstu treści rekordów.
            _add(out, self_key, own)
        else:
            _add(out, self_key, _norm_ws("".join(elem.itertext())))

    siblings = plan.child_tags.get(abs_path) or frozenset()
    for child in kept:
        child_path = _join_path(prefix, child.tag)
        child_abs = _join_path(abs_path, child.tag)
        # Element, w którego wnętrzu siedzą rekordy, nigdy nie jest parą
        # klucz–wartość: para „zjadłaby” pozycje i wciągnęła je do kontekstu.
        dziecko_ponad_rekordami = (
            bool(record_path) and record_path.startswith(child_abs + "/"))
        kv = (_kv_info(child)
              if child_abs in plan.kv_paths and not dziecko_ponad_rekordami
              else None)
        if kv is None:
            _collect(child, child_path, out, plan, child_abs, depth + 1)
            continue
        key_text, value, extras, leftover, unused = kv
        key = _kv_key(key_text)
        # Klucz pochodzi z DANYCH, więc może wskazać istniejącą kolumnę
        # (``<model>`` obok ``<spec name="model">``). Przy kolizji chowamy go
        # w przestrzeni nazw elementu-pojemnika: „spec/model”.
        if key in siblings:
            key_path = _join_path(child_path, key)
        else:
            key_path = _join_path(prefix, key)
        _add(out, key_path, value)
        if leftover:
            _add(out, key_path + " (tekst)", leftover)
        for suffix, extra_value in extras:
            _add(out, child_path + suffix, extra_value)
        for grand in unused:
            grand_abs = _join_path(child_abs, grand.tag)
            if grand_abs == record_path:
                continue                       # to jest rekord, nie kontekst
            _collect(grand, _join_path(child_path, grand.tag), out, plan,
                     grand_abs, depth + 1)


def _finalize(out: Dict[str, List[Any]], repeat: str, join_sep: str) -> Dict[str, Any]:
    """Zamienia listy wartości na pojedyncze pola wg trybu ``repeat``."""
    result: Dict[str, Any] = {}
    for key, values in out.items():
        if len(values) == 1:
            result[key] = values[0]
        elif repeat == "index":
            for num, value in enumerate(values, 1):
                result["%s[%d]" % (key, num)] = value
        else:
            parts = [str(v) for v in values if v is not None and str(v) != ""]
            result[key] = join_sep.join(parts) if parts else None
    return result


_INDEXED_KEY_RE = re.compile(r"^(?P<base>.+)\[(?P<num>\d+)\]$")


def _harmonize_index(records: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Ujednolica kolumny w trybie ``index`` w obrębie jednego dokumentu.

    Gdy w jednym rekordzie pole wystąpiło wielokrotnie (``foto[1]``, ``foto[2]``),
    a w innym tylko raz (``foto``), to jednokrotne też dostaje ``[1]`` – inaczej
    ta sama informacja trafiłaby w arkuszu do dwóch różnych kolumn.
    """
    bases: Set[str] = set()
    for record in records:
        for key in record:
            match = _INDEXED_KEY_RE.match(key)
            if match:
                bases.add(match.group("base"))
    if not bases:
        return records
    result: List[Dict[str, Any]] = []
    for record in records:
        result.append({(key + "[1]" if key in bases else key): value
                       for key, value in record.items()})
    return result


def _flatten(elem: ET.Element, prefix: str, repeat: str, join_sep: str,
             plan: _Plan, abs_path: str) -> Dict[str, Any]:
    out: Dict[str, List[Any]] = {}
    _collect(elem, prefix, out, plan, abs_path)
    return _finalize(out, repeat, join_sep)


# ---------------------------------------------------------------------------
# Wykrywanie ścieżki rekordu
# ---------------------------------------------------------------------------

def _is_plural_of(parent: str, child: str) -> bool:
    """``lots``/``lot``, ``items``/``item``, ``pozycje``/``pozycja``…"""
    parent, child = parent.lower(), child.lower()
    if not parent or not child or parent == child:
        return False
    common = 0
    for a, b in zip(parent, child):
        if a != b:
            break
        common += 1
    return common >= max(3, len(child) - 2)


def _scan(root: ET.Element) -> Dict[str, Dict[str, Any]]:
    """Zbiera statystyki wszystkich ścieżek w dokumencie.

    Zużycie pamięci jest LINIOWE względem liczby RÓŻNYCH ścieżek. Wcześniej
    każda ścieżka trzymała zbiór wszystkich względnych ścieżek pól potomnych,
    sklejanych łańcuchowo w górę drzewa – koszt „głębokość × szerokość”
    potrafił zjeść setki MB na poprawnym pliku ważącym 200 kB. Teraz ścieżka
    pamięta wyłącznie SWOJE atrybuty i to, czy bywa gołym liściem, a liczbę
    różnych pól w poddrzewie sumujemy jednym przebiegiem od najgłębszych
    ścieżek (patrz :func:`_sum_fields`).
    """
    stats: Dict[str, Dict[str, Any]] = {}
    counter = itertools.count()
    totals = {"elements": 0}

    def visit(elem: ET.Element, path: str, depth: int, parent_tag: str,
              parent_path: Optional[str]) -> None:
        st = stats.get(path)
        if st is None:
            if len(stats) >= _MAX_PATHS:
                raise XmlParseError(
                    "Zbyt wiele różnych ścieżek w dokumencie (limit %d) – "
                    "wskaż element pozycji opcją --record-path" % _MAX_PATHS
                )
            st = stats[path] = {
                "count": 0, "fields": 0, "depth": depth,
                "order": next(counter), "tag": elem.tag,
                "parent": parent_tag, "parent_path": parent_path,
                "homogeneous": True, "attrs": set(), "bare_leaf": False,
            }
        st["count"] += 1
        totals["elements"] += 1
        if totals["elements"] > _MAX_ELEMENTS:
            raise XmlParseError(
                "Zbyt wiele elementów w dokumencie (limit %d) – wskaż element "
                "pozycji opcją --record-path" % _MAX_ELEMENTS
            )
        if elem.attrib:
            st["attrs"].update(elem.attrib)
        kids = [c for c in elem if _is_element(c)]
        if not kids:
            if not elem.attrib:
                st["bare_leaf"] = True
        elif depth < _MAX_DEPTH:
            for kid in kids:
                visit(kid, _join_path(path, kid.tag), depth + 1, elem.tag, path)
        if len({c.tag for c in kids}) > 1:
            for kid in kids:
                kid_stats = stats.get(_join_path(path, kid.tag))
                if kid_stats is not None:
                    kid_stats["homogeneous"] = False

    visit(root, root.tag, 0, "", None)
    _sum_fields(stats)
    return stats


def _sum_fields(stats: Dict[str, Dict[str, Any]]) -> None:
    """Uzupełnia ``fields`` – liczbę RÓŻNYCH pól-liści w poddrzewie ścieżki.

    Polem jest atrybut (``sciezka@atrybut``) albo goły liść bez atrybutów.
    Każde pole liczy się raz dla swojej ścieżki i raz dla każdego przodka,
    więc sumujemy poddrzewa od najgłębszych ścieżek w górę.
    """
    contrib = {path: len(st["attrs"]) + (1 if st["bare_leaf"] else 0)
               for path, st in stats.items()}
    subtotal = dict(contrib)
    for path in sorted(stats, key=lambda p: -stats[p]["depth"]):
        parent = stats[path]["parent_path"]
        if parent is not None and parent in subtotal:
            subtotal[parent] += subtotal[path]
    for path, st in stats.items():
        # pola przodka to jego atrybuty + wszystko, co wnoszą potomkowie
        st["fields"] = len(st["attrs"]) + subtotal[path] - contrib[path]


class _ScanBuilder:
    """Budowniczy, który zamiast drzewa zbiera statystyki i plan spłaszczania.

    Używany w PIERWSZYM przebiegu nad dużymi plikami: daje dokładnie to samo,
    co :func:`_scan` + :func:`_build_plan` na gotowym drzewie, ale nie trzyma
    ani jednego węzła w pamięci.
    """

    __slots__ = ("stats", "kv_paths", "child_tags", "_stack", "_order",
                 "_elements")

    def __init__(self) -> None:
        self.stats: Dict[str, Dict[str, Any]] = {}
        self.kv_paths: Set[str] = set()
        self.child_tags: Dict[str, Set[str]] = {}
        self._stack: List[List[Any]] = []      # [path, tag, {tag: n}, ma_atrybuty]
        self._order = itertools.count()
        self._elements = 0

    def start(self, tag: str, attrs: Dict[str, str]) -> None:
        parent = self._stack[-1] if self._stack else None
        path = _join_path(parent[0], tag) if parent is not None else tag
        st = self.stats.get(path)
        if st is None:
            if len(self.stats) >= _MAX_PATHS:
                raise XmlParseError(
                    "Zbyt wiele różnych ścieżek w dokumencie (limit %d) – "
                    "wskaż element pozycji opcją --record-path" % _MAX_PATHS
                )
            st = self.stats[path] = {
                "count": 0, "fields": 0, "depth": len(self._stack),
                "order": next(self._order), "tag": tag,
                "parent": parent[1] if parent is not None else "",
                "parent_path": parent[0] if parent is not None else None,
                "homogeneous": True, "attrs": set(), "bare_leaf": False,
            }
        st["count"] += 1
        self._elements += 1
        if self._elements > _MAX_ELEMENTS:
            raise XmlParseError(
                "Zbyt wiele elementów w dokumencie (limit %d) – wskaż element "
                "pozycji opcją --record-path" % _MAX_ELEMENTS
            )
        if attrs:
            st["attrs"].update(attrs)
        if parent is not None:
            parent[2][tag] = parent[2].get(tag, 0) + 1
        self._stack.append([path, tag, {}, bool(attrs)])

    def data(self, chunk: str) -> None:
        pass

    def end(self, tag: str) -> None:
        path, _, kids, has_attrs = self._stack.pop()
        st = self.stats[path]
        if not kids:
            if not has_attrs:
                st["bare_leaf"] = True
            return
        tags = self.child_tags.get(path)
        if tags is None:
            tags = self.child_tags[path] = set()
        tags.update(kids)
        mixed = len(kids) > 1
        for kid_tag, number in kids.items():
            kid_path = _join_path(path, kid_tag)
            if number >= 2 or _local_name(kid_tag).lower() in _KV_ELEMENT_NAMES:
                self.kv_paths.add(kid_path)
            if mixed:
                kid_stats = self.stats.get(kid_path)
                if kid_stats is not None:
                    kid_stats["homogeneous"] = False

    def close(self) -> "_ScanBuilder":
        _sum_fields(self.stats)
        return self


class _RecordBuilder(ET.TreeBuilder):
    """Budowniczy drzewa, który SPŁASZCZA i zwalnia rekordy w trakcie parsowania.

    Dzięki temu w pamięci nigdy nie leży jednocześnie całe drzewo i wszystkie
    wiersze: po zamknięciu elementu rekordu zostaje po nim pusty znacznik
    (potrzebny tylko po to, żeby kontekst wiedział, że tu były pozycje).
    """

    def __init__(self, record_path: str, plan: _Plan, repeat: str,
                 join_sep: str, records: List[Dict[str, Any]]) -> None:
        super().__init__()
        self._record_path = record_path
        self._plan = plan
        self._repeat = repeat
        self._join_sep = join_sep
        self._records = records
        self._path: List[str] = []

    def start(self, tag, attrs):        # type: ignore[override]
        parent = self._path[-1] if self._path else None
        self._path.append(_join_path(parent, tag) if parent is not None else tag)
        return super().start(tag, attrs)

    def end(self, tag):                 # type: ignore[override]
        elem = super().end(tag)
        path = self._path.pop()
        if path == self._record_path:
            self._records.append(_flatten(elem, "", self._repeat, self._join_sep,
                                          self._plan, path))
            elem.clear()                # zwalniamy poddrzewo rekordu
        return elem


def _score(path: str, st: Dict[str, Any]) -> float:
    name = _local_name(st["tag"]).lower()
    fields = max(1, st["fields"])
    bonus = 1.0
    if name in _PREFERRED_RECORD_NAMES:
        bonus *= 3.0
    elif name in _DISCOURAGED_RECORD_NAMES:
        bonus *= 0.4
    if _is_plural_of(_local_name(st["parent"]), name):
        bonus *= 1.6
    if st["homogeneous"]:
        bonus *= 1.2
    if not st["fields"]:
        # Czysty liść (bez dzieci i atrybutów) to kiepski kandydat na rekord –
        # dałby arkusz z jedną, bezimienną kolumną.
        bonus *= 0.35
    return ((math.log2(st["count"]) + 1.0) * fields * bonus
            / (1.0 + 0.08 * st["depth"]))


def _preferred_single_path(stats: Dict[str, Dict[str, Any]]) -> Optional[str]:
    """Ratunkowy kandydat na rekord wśród ścieżek NIEpowtarzających się.

    Używany tylko wtedy, gdy jedyne powtarzalne ścieżki to „śmieci” (zdjęcia,
    załączniki, słowa kluczowe – czyste liście albo pojedyncze pole). Realny
    przypadek z portalu: pakiet z JEDNĄ sztuką i czterema zdjęciami; wygrywały
    zdjęcia i arkusz dostawał cztery wiersze-widma zamiast jednego.
    """
    pool = [
        path for path, st in stats.items()
        if st["depth"] > 0                       # korzeń to nie rekord
        and st["fields"] >= _MIN_RECORD_FIELDS
        and _local_name(st["tag"]).lower() in _PREFERRED_RECORD_NAMES
    ]
    if not pool:
        return None
    # Rekordem jest element najbardziej WEWNĘTRZNY: gdy w środku <lot> siedzą
    # <item>-y, rekordem jest <item>, a metadane lotu trafiają do kontekstu.
    inner = [path for path in pool
             if not any(other.startswith(path + "/") for other in pool)]
    return min(inner, key=lambda path: (-_score(path, stats[path]),
                                        stats[path]["depth"],
                                        stats[path]["order"], path))


def _choose_record_path(stats: Dict[str, Dict[str, Any]]) -> Optional[str]:
    """Wybiera ścieżkę rekordu na podstawie statystyk z :func:`_scan`."""
    best: Optional[Tuple[float, int, int, str]] = None
    for path, st in stats.items():
        if st["count"] < 2:
            continue
        key = (-_score(path, st), st["depth"], st["order"], path)
        if best is None or key < best:
            best = key
    if best is None:
        return None
    if stats[best[3]]["fields"] < _MIN_RECORD_FIELDS:
        # Powtarza się tylko coś ubogiego (zdjęcia, tagi, linki). Jeśli w
        # dokumencie jest element o nazwie typowej dla pozycji i z prawdziwą
        # treścią – to ON jest rekordem, choćby wystąpił raz.
        fallback = _preferred_single_path(stats)
        if fallback is not None and fallback != best[3]:
            return fallback
    return best[3]


def detect_record_path(root: Any) -> Optional[str]:
    """Zwraca ścieżkę elementu powtarzalnego (np. ``auction/lots/lot``).

    ``None`` oznacza, że nic się nie powtarza i dokument jest jednym wierszem.
    """
    if root is None or not _is_element(root):
        return None
    return _choose_record_path(_scan(root))


def _resolve_record_path(stats: Dict[str, Dict[str, Any]], wanted: str) -> str:
    """Dopasowuje ścieżkę podaną przez użytkownika (dopuszcza sam sufiks)."""
    wanted = wanted.strip().strip("/")
    if wanted in stats:
        return wanted
    matches = [p for p in stats if p.endswith("/" + wanted)]
    if not matches:
        matches = [p for p in stats if _local_name(p.rsplit("/", 1)[-1]) == wanted]
    if not matches:
        repeated = sorted(p for p, st in stats.items() if st["count"] >= 2)
        raise ValueError(
            "Nie znaleziono elementów dla ścieżki rekordu %r. "
            "Powtarzalne ścieżki w dokumencie: %s"
            % (wanted, ", ".join(repeated) or "brak")
        )
    matches.sort(key=lambda p: (-stats[p]["count"], stats[p]["depth"], stats[p]["order"]))
    return matches[0]


def _find_records(root: ET.Element, path: str) -> List[ET.Element]:
    """Elementy o zadanej ścieżce (bez schodzenia w rekordy zagnieżdżone)."""
    found: List[ET.Element] = []

    def walk(elem: ET.Element, current: str) -> None:
        if current == path:
            found.append(elem)
            return
        if not path.startswith(current + "/"):
            return
        for kid in elem:
            if _is_element(kid):
                walk(kid, _join_path(current, kid.tag))

    walk(root, root.tag)
    return found


# ---------------------------------------------------------------------------
# Identyfikator aukcji
# ---------------------------------------------------------------------------

_AUCTION_STRONG_RE = re.compile(
    r"^(auction|aukcja|aukcji|aukcje)[_\- ]?(id|no|nr|num|numer|number|code|kod|sygnatura)?$"
    r"|^(id|nr|numer|number|kod|code|sygnatura)[_\- ]?(aukcji|auction)$",
    re.IGNORECASE,
)
_AUCTION_OWNER_RE = re.compile(r"^(auction|aukcj|package|pakiet|oferta|offer)", re.IGNORECASE)
_AUCTION_WEAK = frozenset({
    "id", "identyfikator", "sygnatura", "nr", "numer", "number", "code", "kod",
    "reference", "ref",
})


def _source_stem(source: str) -> str:
    base = os.path.basename(str(source).replace("\\", "/"))
    stem = os.path.splitext(base)[0]
    return stem or str(source)


def _detect_auction(meta: Dict[str, Any], source: str) -> str:
    """Szuka identyfikatora aukcji w metadanych; awaryjnie nazwa pliku."""
    best: Optional[Tuple[int, int, int, str]] = None
    for order, (key, value) in enumerate(meta.items()):
        if value is None:
            continue
        text = _norm_ws(str(value))
        if not text:
            continue
        tail = key.rsplit("/", 1)[-1]
        owner, _, attr = tail.partition("@")
        name = _local_name(attr or owner).lower()
        owner = _local_name(owner).lower()
        if _AUCTION_STRONG_RE.match(name):
            rank = 0
        elif attr and _AUCTION_OWNER_RE.match(owner) and name in _AUCTION_WEAK:
            # np. <auction id="..."> albo <aukcja nr="...">
            rank = 0
        elif name in _AUCTION_WEAK:
            rank = 1
        else:
            continue
        candidate = (rank, key.count("/"), order, text)
        if best is None or candidate < best:
            best = candidate
    return best[3] if best else _source_stem(source)


# ---------------------------------------------------------------------------
# API publiczne
# ---------------------------------------------------------------------------

# ---------------------------------------------------------------------------
# Dwa tryby wczytywania: w pamięci (małe pliki) i strumieniowy (duże)
# ---------------------------------------------------------------------------

#: Powyżej tylu znaków dokumentu przechodzimy na tryb strumieniowy. Poniżej
#: nie warto płacić drugim przebiegiem parsera (koszt ok. +40% czasu); powyżej
#: – pełne drzewo razem z gotowymi wierszami potrafi zająć kilkanaście razy
#: więcej pamięci niż sam plik, a to na laptopie z 4-8 GB RAM kończy się
#: swapem albo zabiciem procesu.
_STREAM_MIN_CHARS = 8 * 1024 * 1024


def _parse_in_memory(text: str, source: str, strip_ns: bool, repeat: str,
                     join_sep: str, record_path: Optional[str]
                     ) -> Tuple[Optional[str], Dict[str, Any], List[Dict[str, Any]]]:
    """Buduje całe drzewo i spłaszcza je – prosty tryb dla typowych plików."""
    root = _parse_xml(text, source, strip_ns)[0]
    stats = _scan(root)
    if record_path:
        path: Optional[str] = _resolve_record_path(stats, record_path)
    else:
        path = _choose_record_path(stats)
    elements = _find_records(root, path) if path else []
    if not elements:
        # Brak powtórzeń – cały dokument to jeden wiersz.
        plan = _build_plan(root, None)
        return None, {}, [_flatten(root, "", repeat, join_sep, plan, root.tag)]

    plan = _build_plan(root, path)
    # Gdy rekordem jest sam korzeń (wymuszone --record-path), poza rekordem nie
    # ma już nic – kontekst zostaje pusty zamiast dublować wszystkie wartości.
    context = ({} if path == root.tag
               else _flatten(root, root.tag, repeat, join_sep, plan, root.tag))
    records = []
    for elem in elements:
        records.append(_flatten(elem, "", repeat, join_sep, plan, path))
        # Poddrzewo rekordu nie jest już potrzebne – zwalniamy je od razu,
        # żeby szczyt pamięci nie był sumą „drzewo + wszystkie wiersze”.
        elem.clear()
    return path, context, records


def _parse_streaming(text: str, source: str, strip_ns: bool, repeat: str,
                     join_sep: str, record_path: Optional[str]
                     ) -> Tuple[Optional[str], Dict[str, Any], List[Dict[str, Any]]]:
    """Dwa przebiegi: statystyki bez drzewa, potem wiersze spłaszczane w locie.

    Wynik jest identyczny jak w :func:`_parse_in_memory`, ale w pamięci nie
    leży jednocześnie pełne drzewo dokumentu i komplet wierszy.
    """
    scan, text, ns_aware = _parse_xml(text, source, strip_ns, _ScanBuilder)
    stats = scan.stats
    if record_path:
        path: Optional[str] = _resolve_record_path(stats, record_path)
    else:
        path = _choose_record_path(stats)

    if not path or path not in stats:
        plan = _Plan(scan.kv_paths, scan.child_tags, None)
        root = _run_parser(text, source, strip_ns, ns_aware, set(), warn=False)
        return None, {}, [_flatten(root, "", repeat, join_sep, plan, root.tag)]

    plan = _Plan(scan.kv_paths, scan.child_tags, path)
    records: List[Dict[str, Any]] = []
    builder = _RecordBuilder(path, plan, repeat, join_sep, records)
    root = _run_parser(text, source, strip_ns, ns_aware, set(), builder, warn=False)
    if not records:  # pragma: no cover - ścieżka ze statystyk zawsze istnieje
        plan = _Plan(scan.kv_paths, scan.child_tags, None)
        return None, {}, [_flatten(root, "", repeat, join_sep, plan, root.tag)]
    context = ({} if path == root.tag
               else _flatten(root, root.tag, repeat, join_sep, plan, root.tag))
    return path, context, records


def parse_bytes(data: bytes, source: str, *, strip_ns: bool = True,
                repeat: str = "join", join_sep: str = " | ",
                record_path: Optional[str] = None) -> ParsedDoc:
    """Parsuje bajty XML i spłaszcza je do ``ParsedDoc``.

    Kodowanie odczytywane jest z BOM albo deklaracji ``<?xml … encoding="…"?>``.
    """
    if repeat not in ("join", "index"):
        raise ValueError("Nieznany tryb repeat=%r (dozwolone: 'join', 'index')" % repeat)
    if isinstance(data, str):
        data = data.encode("utf-8")
    text = _decode(bytes(data), source)
    data = b""            # bajty nie są już potrzebne – oddajemy pamięć

    if len(text) >= _STREAM_MIN_CHARS:
        path, context, records = _parse_streaming(
            text, source, strip_ns, repeat, join_sep, record_path)
    else:
        path, context, records = _parse_in_memory(
            text, source, strip_ns, repeat, join_sep, record_path)
    text = ""             # przy 100-megabajtowym pliku to realny zysk pamięci
    if repeat == "index":
        records = _harmonize_index(records)

    columns: List[str] = []
    seen: Set[str] = set()
    for key in itertools.chain(context.keys(), *(r.keys() for r in records)):
        if key not in seen:
            seen.add(key)
            columns.append(key)

    # Identyfikator szukamy w metadanych aukcji; pola rekordu (np. id lotu)
    # bierzemy pod uwagę tylko wtedy, gdy dokument jest jednym wierszem.
    meta = context if context else (records[0] if path is None else {})
    auction = _detect_auction(meta, source)
    return ParsedDoc(source=str(source), auction=auction, context=context,
                     records=records, record_path=path, columns=columns)


def parse_file(path, **kw) -> ParsedDoc:
    """Wczytuje plik z dysku i przekazuje go do :func:`parse_bytes`."""
    source = kw.pop("source", None) or str(path)
    with open(path, "rb") as handle:
        # bez pośredniej zmiennej: bajty żyją tylko wewnątrz parse_bytes
        return parse_bytes(handle.read(), source, **kw)


def merge_columns(docs: Sequence[ParsedDoc]) -> List[str]:
    """Unia kolumn wielu dokumentów w kolejności pierwszego wystąpienia."""
    merged: Dict[str, None] = {}
    for doc in docs:
        for column in doc.columns:
            merged.setdefault(column, None)
    return list(merged)


def rows_for(doc: ParsedDoc, columns: Sequence[str]) -> Iterator[List[Any]]:
    """Generuje wiersze dokumentu w kolejności zadanych kolumn.

    Pola z ``context`` powtarzają się w każdym wierszu; przy kolizji nazw
    pierwszeństwo ma wartość z rekordu.
    """
    for record in doc.records:
        merged = dict(doc.context)
        merged.update(record)
        yield [merged.get(column) for column in columns]
