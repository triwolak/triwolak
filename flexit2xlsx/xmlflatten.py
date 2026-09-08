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

Pozostałe reguły
----------------
* Klucz kolumny rekordu = ścieżka WZGLĘDNA wobec elementu rekordu
  (``specs/spec``), atrybuty jako ``sciezka@atrybut``; atrybut samego rekordu to
  ``@atrybut``.
* Wszystko poza poddrzewami rekordów trafia do ``context`` z kluczem
  BEZWZGLĘDNYM (z nazwą korzenia, np. ``auction/seller/name``). Dzięki temu
  metadane aukcji nigdy nie kolidują z kolumnami rekordu.
* Pary klucz–wartość (``<spec name="RAM" value="16 GB"/>``) zamieniamy na
  kolumnę ``specs/RAM``. Robimy to tylko wtedy, gdy nazwa elementu wygląda na
  pojemnik na cechy ALBO element powtarza się w rodzicu ≥2 razy – i gdy nic
  poza parą nie zostałoby zgubione.
* Bezpieczeństwo: własny parser na ``xml.parsers.expat`` blokuje encje
  zewnętrzne (XXE) i encje rekurencyjne (bomba encyjna) – ``xml.etree`` sam z
  siebie rozwija bomby encyjne, dlatego nie używamy go do parsowania.
"""

from __future__ import annotations

import codecs
import itertools
import math
import os
import re
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
_MAX_DEPTH = 200            # maksymalne zagnieżdżenie XML (ochrona stosu)
_MAX_ENTITY_DECLS = 64      # maksymalna liczba deklaracji encji
_MAX_ENTITY_VALUE = 1024    # maksymalna długość treści encji
_TEXT_AMPLIFICATION = 20    # ile razy tekst może przerosnąć źródło
_MIN_TEXT_LIMIT = 4 * 1024 * 1024   # dolna granica limitu tekstu

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

_BOMS: Tuple[Tuple[bytes, str], ...] = (
    (codecs.BOM_UTF32_LE, "utf-32-le"),
    (codecs.BOM_UTF32_BE, "utf-32-be"),
    (codecs.BOM_UTF8, "utf-8"),
    (codecs.BOM_UTF16_LE, "utf-16-le"),
    (codecs.BOM_UTF16_BE, "utf-16-be"),
)
_FALLBACK_ENCODINGS = ("utf-8", "cp1250", "iso-8859-2", "latin-1")


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

def _sniff_encoding(data: bytes) -> Tuple[bytes, List[str]]:
    """Zwraca (dane bez BOM, lista kandydatów na kodowanie) wg BOM i deklaracji.

    Kolejność ma znaczenie: BOM, potem deklaracja ``<?xml … encoding="…"?>``,
    na końcu awaryjny łańcuch. Gdy BOM kłamie (zdarza się w eksportach), i tak
    wracamy do kodowania zadeklarowanego w dokumencie.
    """
    candidates: List[str] = []
    body = data
    for bom, enc in _BOMS:
        if data.startswith(bom):
            body = data[len(bom):]
            candidates.append(enc)
            break
    else:
        # UTF-16 bez BOM rozpoznajemy po zerowych bajtach wokół "<"
        if data.startswith(b"<\x00"):
            candidates.append("utf-16-le")
        elif data.startswith(b"\x00<"):
            candidates.append("utf-16-be")
    match = _DECL_ENC_RE.search(body[:1024].decode("latin-1", "replace"))
    if match:
        candidates.append(match.group(1))
    candidates.extend(_FALLBACK_ENCODINGS)
    return body, candidates


def _decode(data: bytes, source: str) -> str:
    """Dekoduje bajty wg BOM/deklaracji, z awaryjnym łańcuchem kodowań."""
    if not data or not data.strip():
        raise XmlParseError("Pusty dokument XML: %s" % source)
    body, candidates = _sniff_encoding(data)
    for candidate in candidates:
        try:
            text = body.decode(candidate)
        except (LookupError, UnicodeDecodeError, ValueError):
            continue
        break
    else:  # pragma: no cover - latin-1 nigdy nie zawodzi
        text = body.decode("latin-1", "replace")
    text = text.lstrip("\ufeff")
    if not text.startswith("<"):
        text = text.lstrip()
    if not text:
        raise XmlParseError("Pusty dokument XML: %s" % source)
    return text


def _repair_entities(text: str, declared: Set[str]) -> str:
    """Zamienia nieznane encje (np. ``&nbsp;``) na tekst – ratunek dla eksportów.

    Uruchamiane WYŁĄCZNIE po błędzie „undefined entity”, żeby nie modyfikować
    poprawnych dokumentów. Encje wbudowane i zadeklarowane w dokumencie zostają
    nietknięte.
    """

    def _sub(match: "re.Match[str]") -> str:
        name = match.group(1)
        if name in _BUILTIN_ENTITIES or name in declared:
            return match.group(0)
        repl = _HTML5_ENTITIES.get(name + ";") or _HTML5_ENTITIES.get(name)
        if repl is None:
            return ""
        return (
            repl.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        )

    return _ENTITY_REF_RE.sub(_sub, text)


# ---------------------------------------------------------------------------
# Bezpieczny parser XML (expat + ET.TreeBuilder)
# ---------------------------------------------------------------------------

def _run_parser(text: str, source: str, strip_ns: bool, ns_aware: bool,
                declared: Set[str]) -> ET.Element:
    """Buduje drzewo ``ET.Element`` własnym parserem expat z blokadą encji."""
    builder = ET.TreeBuilder()
    parser = expat.ParserCreate("utf-8", "}" if ns_aware else None)
    parser.buffer_text = True
    try:
        parser.SetParamEntityParsing(expat.XML_PARAM_ENTITY_PARSING_NEVER)
    except (AttributeError, expat.error):  # pragma: no cover
        pass

    prefixes: Dict[str, Optional[str]] = {}
    cache: Dict[str, str] = {}
    state = {"depth": 0, "entities": 0, "chars": 0}
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
                fixed = local
            else:
                prefix = prefixes.get(uri)
                fixed = "%s:%s" % (prefix, local) if prefix else "{%s}%s" % (uri, local)
        elif strip_ns and ":" in name:
            fixed = name.split(":", 1)[1]
        cache[name] = fixed
        return fixed

    def start(tag: str, attrs: Dict[str, str]) -> None:
        state["depth"] += 1
        if state["depth"] > _MAX_DEPTH:
            raise XmlParseError(
                "Zbyt głęboko zagnieżdżony XML (limit %d): %s" % (_MAX_DEPTH, source)
            )
        builder.start(fixname(tag), {fixname(k): v for k, v in attrs.items()})

    def end(tag: str) -> None:
        state["depth"] -= 1
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
        if value is None or len(value) > _MAX_ENTITY_VALUE or "&" in value or "%" in value:
            raise XmlParseError(
                "Zablokowano rekurencyjną/nadmiarową encję '%s' "
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
        builder.data(chunk)

    parser.StartElementHandler = start
    parser.EndElementHandler = end
    parser.CharacterDataHandler = data
    parser.EntityDeclHandler = entity_decl
    parser.ExternalEntityRefHandler = external_ref
    if ns_aware:
        parser.StartNamespaceDeclHandler = ns_decl

    parser.Parse(text.encode("utf-8"), True)
    root = builder.close()
    if root is None:  # pragma: no cover - expat zgłosiłby błąd wcześniej
        raise XmlParseError("Dokument XML nie zawiera elementu głównego: %s" % source)
    return root


def _parse_xml(text: str, source: str, strip_ns: bool) -> ET.Element:
    """Parsuje z dwiema próbami naprawy: nieznany prefiks i nieznana encja."""
    ns_aware = True
    repaired = False
    declared: Set[str] = set()
    for _ in range(4):
        try:
            return _run_parser(text, source, strip_ns, ns_aware, declared)
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
                text = _repair_entities(text, declared)
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

def _kv_info(elem: ET.Element) -> Optional[Tuple[str, Any, Set[str], Set[int]]]:
    """Rozpoznaje parę klucz–wartość, np. ``<spec name="RAM">16 GB</spec>``.

    Zwraca ``(klucz, wartość, zużyte_atrybuty, zużyte_dzieci)`` albo ``None``,
    gdy element nie jest jednoznaczną parą (wtedy spłaszczamy go normalnie).
    """
    used_attrs: Set[str] = set()
    used_kids: Set[int] = set()
    key_text = ""
    for attr, val in elem.attrib.items():
        if _local_name(attr).lower() in _KEY_NAMES:
            key_text = _norm_ws(val)
            used_attrs.add(attr)
            break
    kids = [c for c in elem if _is_element(c)]
    if not key_text:
        for kid in kids:
            if _local_name(kid.tag).lower() in _KEY_NAMES and not any(
                    _is_element(g) for g in kid):
                key_text = _norm_ws("".join(kid.itertext()))
                used_kids.add(id(kid))
                break
    if not key_text:
        return None

    for attr, val in elem.attrib.items():
        if attr in used_attrs:
            continue
        if _local_name(attr).lower() in _VALUE_NAMES:
            used_attrs.add(attr)
            return key_text, (_norm_ws(val) or None), used_attrs, used_kids
    for kid in kids:
        if id(kid) in used_kids:
            continue
        if _local_name(kid.tag).lower() in _VALUE_NAMES and not any(
                _is_element(g) for g in kid):
            used_kids.add(id(kid))
            return (key_text, (_norm_ws("".join(kid.itertext())) or None),
                    used_attrs, used_kids)

    # Wariant „wartość to treść elementu” – tylko gdy nic innego nie zginie.
    if len(used_attrs) == len(elem.attrib) and all(id(k) in used_kids for k in kids):
        own = _norm_ws("".join(t for t in itertools.chain(
            [elem.text], (k.tail for k in kids)) if t))
        if own:
            return key_text, own, used_attrs, used_kids
    return None


def _add(out: Dict[str, List[Any]], key: str, value: Any) -> None:
    out.setdefault(key, []).append(value)


def _collect(elem: ET.Element, prefix: str, out: Dict[str, List[Any]],
             skip: Set[int], depth: int = 0) -> None:
    """Rekurencyjnie zbiera wartości poddrzewa do mapy ``klucz -> [wartości]``."""
    if depth > _MAX_DEPTH:
        return
    for attr, val in elem.attrib.items():
        _add(out, "%s@%s" % (prefix, attr), _norm_ws(val) or None)

    children = [c for c in elem if _is_element(c)]
    if not children:
        text = _norm_ws("".join(elem.itertext()))
        # Element bez treści, ale z atrybutami (np. <item id="1"/>) nie tworzy
        # osobnej pustej kolumny – cała informacja jest już w atrybutach.
        if text or not elem.attrib:
            _add(out, prefix or _local_name(elem.tag), text or None)
        return

    kept = [c for c in children if id(c) not in skip]
    skipped_any = len(kept) != len(children)

    # Treść mieszana: element ma dzieci ORAZ własny tekst – zapisujemy całość.
    if kept and not skipped_any:
        own = _norm_ws("".join(t for t in itertools.chain(
            [elem.text], (c.tail for c in children)) if t))
        if own:
            _add(out, prefix or _local_name(elem.tag),
                 _norm_ws("".join(elem.itertext())))

    counts: Dict[str, int] = {}
    for child in kept:
        counts[child.tag] = counts.get(child.tag, 0) + 1

    for child in kept:
        local = _local_name(child.tag).lower()
        child_path = _join_path(prefix, child.tag)
        kv = None
        if local in _KV_ELEMENT_NAMES or counts[child.tag] >= 2:
            kv = _kv_info(child)
        if kv is not None:
            key_text, value, used_attrs, used_kids = kv
            _add(out, _join_path(prefix, key_text), value)
            for attr, val in child.attrib.items():
                if attr not in used_attrs:
                    _add(out, "%s@%s" % (child_path, attr), _norm_ws(val) or None)
            for grand in child:
                if _is_element(grand) and id(grand) not in used_kids:
                    _collect(grand, _join_path(child_path, grand.tag), out, skip,
                             depth + 1)
        else:
            _collect(child, child_path, out, skip, depth + 1)


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
             skip: Optional[Set[int]] = None) -> Dict[str, Any]:
    out: Dict[str, List[Any]] = {}
    _collect(elem, prefix, out, skip or set())
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
    """Zbiera statystyki wszystkich ścieżek w dokumencie."""
    stats: Dict[str, Dict[str, Any]] = {}
    counter = itertools.count()

    def visit(elem: ET.Element, path: str, depth: int, parent_tag: str) -> Set[str]:
        st = stats.get(path)
        if st is None:
            st = stats[path] = {
                "count": 0, "fields": set(), "depth": depth,
                "order": next(counter), "tag": elem.tag,
                "parent": parent_tag, "homogeneous": True,
            }
        st["count"] += 1
        rel: Set[str] = {"@" + a for a in elem.attrib}
        kids = [c for c in elem if _is_element(c)]
        if depth < _MAX_DEPTH:
            for kid in kids:
                sub = visit(kid, _join_path(path, kid.tag), depth + 1, elem.tag)
                if sub:
                    rel.update(kid.tag + "/" + s for s in sub)
                else:
                    rel.add(kid.tag)
        homogeneous = len({c.tag for c in kids}) <= 1
        for kid in kids:
            kid_stats = stats.get(_join_path(path, kid.tag))
            if kid_stats is not None and not homogeneous:
                kid_stats["homogeneous"] = False
        st["fields"].update(rel)
        return rel

    visit(root, root.tag, 0, "")
    return stats


def _score(path: str, st: Dict[str, Any]) -> float:
    name = _local_name(st["tag"]).lower()
    fields = max(1, len(st["fields"]))
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


def detect_record_path(root: Any) -> Optional[str]:
    """Zwraca ścieżkę elementu powtarzalnego (np. ``auction/lots/lot``).

    ``None`` oznacza, że nic się nie powtarza i dokument jest jednym wierszem.
    """
    if root is None or not _is_element(root):
        return None
    stats = _scan(root)
    best: Optional[Tuple[float, int, int, str]] = None
    for path, st in stats.items():
        if st["count"] < 2:
            continue
        key = (-_score(path, st), st["depth"], st["order"], path)
        if best is None or key < best:
            best = key
    return best[3] if best else None


def _iter_paths(root: ET.Element) -> Iterator[Tuple[str, ET.Element]]:
    stack = [(root.tag, root)]
    while stack:
        path, elem = stack.pop()
        yield path, elem
        for kid in reversed([c for c in elem if _is_element(c)]):
            stack.append((_join_path(path, kid.tag), kid))


def _resolve_record_path(root: ET.Element, wanted: str) -> str:
    """Dopasowuje ścieżkę podaną przez użytkownika (dopuszcza sam sufiks)."""
    wanted = wanted.strip().strip("/")
    stats = _scan(root)
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
    root = _parse_xml(text, source, strip_ns)

    if record_path:
        path = _resolve_record_path(root, record_path)
    else:
        path = detect_record_path(root)

    records: List[Dict[str, Any]] = []
    context: Dict[str, Any] = {}
    if path:
        elements = _find_records(root, path)
    else:
        elements = []

    if elements:
        skip = {id(e) for e in elements}
        context = _flatten(root, root.tag, repeat, join_sep, skip)
        for elem in elements:
            records.append(_flatten(elem, "", repeat, join_sep))
        if repeat == "index":
            records = _harmonize_index(records)
    else:
        # Brak powtórzeń – cały dokument to jeden wiersz.
        path = None
        records.append(_flatten(root, "", repeat, join_sep))

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
        data = handle.read()
    return parse_bytes(data, source, **kw)


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
