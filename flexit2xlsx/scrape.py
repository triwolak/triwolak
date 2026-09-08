# -*- coding: utf-8 -*-
"""Pobieranie plików XML ("zawartość pakietu") z portalu aukcyjnego.

Moduł używa **wyłącznie biblioteki standardowej** (``urllib.request``,
``http.cookiejar``, ``html.parser``, ``re``).  Schemat portalu NIE jest znany,
więc cała logika jest heurystyczna i defensywna: gdy czegoś nie znajdzie,
zwraca pustą listę i pozwala wypisać diagnostykę (:func:`describe_page`),
zamiast wybuchać wyjątkiem.

Hierarchia portalu (wg ``SITE_NOTES.md``) jest TRZYPOZIOMOWA::

    lista aukcji  ->  strona aukcji  ->  strony lotów  ->  "Download Batch Details" (XML)

Podjęte decyzje projektowe (przypadki niejednoznaczne)
-----------------------------------------------------
* **Ile prób.** ``retries=4`` oznacza 4 PONOWIENIA po pierwszej próbie
  (razem 5 żądań), bo tylko wtedy sekwencja opóźnień to dokładnie 2, 4, 8, 16 s
  wymagane przez kontrakt.
* **Wstrzykiwalny sen.** Backoff i uprzejme opóźnienie idą przez
  ``opener.sleep`` (parametr ``sleep=`` w :func:`make_opener`), a domyślnie przez
  modułową funkcję :data:`SLEEP_FUNCTION`.  Testy nigdy nie śpią naprawdę.
* **Granica "tej samej witryny".**  Linki spoza witryny są odrzucane BEZ żądania
  sieciowego.  Za tę samą witrynę uznajemy równy host albo równą domenę
  rejestrowalną przybliżoną dwiema ostatnimi etykietami — dzięki temu działa
  ``media.flexitauctions.com`` przy bazie ``flexitauctions.com``, a
  ``evil.example.com`` jest odrzucane.  Dla adresów IP wymagana jest równość
  hosta ORAZ portu (inny port = inne źródło).  Przekierowanie poza witrynę
  przerywa żądanie (:class:`SiteBlockedError`) — inaczej wyciekłoby ciasteczko.
* **Schematy inne niż http/https** (``file:``, ``javascript:``, ``data:``,
  ``mailto:``) są odrzucane już na etapie wyciągania linków.
* **Zejście do lotów** (``descend="auto"``) następuje tylko wtedy, gdy strona
  aukcji sama nie dała żadnego XML-a.  Portal może linkować XML na obu
  poziomach, a "auto" nie generuje setek zbędnych żądań.
* **Sondowanie Content-Type** dotyczy tylko linków NIEPEWNYCH (przycisk
  "Download Batch Details" bez rozszerzenia).  Budżet sond jest ograniczony
  (``probe_limit``), a wynik cache'owany w openerze, żeby nie odpytywać dwa razy
  tego samego adresu.
* **Nazwy plików** są sprowadzane do ASCII ``[A-Za-z0-9_.-]`` (a nie do
  unicode'owego ``\\w``), bo plik ma być przenośny między systemami plików.
  Zawsze brana jest sama nazwa bazowa — ``../../etc/passwd`` daje ``passwd``.
* **Pobrany plik, który jest stroną HTML** (ściana logowania) to błąd
  :class:`NotXmlError`, a nie "pobrany plik".  Inaczej śmieć wylądowałby w
  katalogu wyjściowym i przy kolejnym uruchomieniu zostałby POMINIĘTY jako
  "już pobrany".
* **Błąd sieci przy PIERWSZEJ stronie listy** jest wyjątkiem (użytkownik musi
  wiedzieć, że portal nie odpowiada), przy kolejnych stronach paginacji — tylko
  ostrzeżeniem.  Brak DOPASOWAŃ to zawsze pusta lista, nigdy wyjątek.
"""

from __future__ import annotations

import hashlib
import http.client
import http.cookiejar
import os
import posixpath
import re
import socket
import sys
import time
import urllib.error
import urllib.parse
import urllib.request
from dataclasses import dataclass, field
from html.parser import HTMLParser
from typing import Any, Callable, Optional

__all__ = [
    "AuctionRef",
    "XmlRef",
    "ScrapeError",
    "FetchError",
    "SiteBlockedError",
    "NotXmlError",
    "Opener",
    "make_opener",
    "fetch",
    "fetch_full",
    "discover_auctions",
    "find_xml_links",
    "download_xml",
    "describe_page",
    "diagnose_last",
    "safe_filename",
    "absolutize",
    "is_same_site",
    "DEFAULT_USER_AGENT",
    "DEFAULT_AUCTION_RE",
    "DEFAULT_DELAY",
    "BACKOFF_BASE",
    "MAX_BACKOFF",
    "RETRY_STATUSES",
    "SLEEP_FUNCTION",
]

# ---------------------------------------------------------------------------
# Stałe konfiguracyjne
# ---------------------------------------------------------------------------

#: Nagłówek ``User-Agent``.  Portale aukcyjne potrafią blokować "gołe" klienty
#: urllib, dlatego podszywamy się pod przeglądarkę, ale uczciwie dopisujemy
#: nazwę narzędzia.
DEFAULT_USER_AGENT = (
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/124.0.0.0 Safari/537.36 flexit2xlsx/1.0"
)

#: Domyślny wzorzec adresu aukcji: ``/auction/<slug>`` z opcjonalną zakładką
#: ``/info``.  Grupa ``id`` to identyfikator aukcji (slug).
DEFAULT_AUCTION_RE = r"/auction/(?P<id>[^/?#]+)/?(?:info/?)?$"

#: Podstawa wykładniczego backoffu — opóźnienia 2, 4, 8, 16 s.
BACKOFF_BASE = 2.0
BACKOFF_FACTOR = 2.0
MAX_BACKOFF = 60.0

#: Uprzejme opóźnienie między kolejnymi żądaniami (sekundy).
DEFAULT_DELAY = 0.5

#: Kody HTTP, po których ma sens ponowienie.
RETRY_STATUSES = frozenset({408, 425, 429, 500, 502, 503, 504})

#: Twardy limit rozmiaru pojedynczej odpowiedzi (ochrona przed zip-bombą/RAM).
DEFAULT_MAX_BYTES = 64 * 1024 * 1024

#: Typy MIME jednoznacznie oznaczające XML.
XML_CONTENT_TYPES = frozenset(
    {"application/xml", "text/xml", "application/rss+xml", "application/atom+xml"}
)

#: Typy MIME, przy których trzeba jeszcze zajrzeć w treść.
AMBIGUOUS_CONTENT_TYPES = frozenset(
    {
        "",
        "application/octet-stream",
        "binary/octet-stream",
        "application/download",
        "application/force-download",
        "text/plain",
        "application/x-download",
    }
)

#: Rozszerzenia, których na pewno nie sondujemy (oszczędność żądań).
_BORING_EXT = (
    ".jpg", ".jpeg", ".png", ".gif", ".webp", ".svg", ".ico", ".bmp", ".avif",
    ".css", ".js", ".mjs", ".map", ".woff", ".woff2", ".ttf", ".eot", ".otf",
    ".pdf", ".zip", ".rar", ".7z", ".gz", ".mp4", ".webm", ".mp3", ".doc",
    ".docx", ".xls", ".xlsx", ".csv", ".txt", ".json",
)

#: Słowa-klucze w treści/atrybutach linku sugerujące plik z zawartością pakietu.
_XML_HINT_RE = re.compile(
    r"(batch|manifest|packing|content|details|specification|export|download|"
    r"pobierz|zawarto|pakiet|specyfikacj|wykaz|zestawienie|xml)",
    re.IGNORECASE,
)

#: Wzorzec strony lotu: ``/lot/<slug>-<hash>`` albo ``/auction/<aukcja>/<slug>-<hash>``.
_LOT_RE = re.compile(
    r"^/(?:lot/|auction/[^/?#]+/)(?P<slug>[^/?#]*?)-(?P<hash>[0-9a-f]{4,8})/?$",
    re.IGNORECASE,
)

#: Segmenty pod ``/auction/<slug>/``, które NIE są lotami.
_NOT_LOT_SEGMENTS = frozenset(
    {
        "info", "terms", "conditions", "faq", "help", "contact", "lots",
        "gallery", "map", "documents", "regulamin", "kontakt", "search",
    }
)

#: Parametry zapytania, w których portal potrafi trzymać nazwę pliku.
_FILENAME_PARAMS = ("file", "filename", "name", "doc", "document", "path", "f")

#: Atrybuty HTML, w których może siedzieć adres (także w SPA).
_URL_ATTRS = (
    "href", "data-href", "data-url", "data-download", "data-file", "data-xml",
    "data-src", "data-link",
)

#: Maksymalna długość nazwy pliku (z rozszerzeniem).
MAX_FILENAME_LEN = 120

#: Domyślna funkcja "śpij" — testy podmieniają ten atrybut modułu albo
#: przekazują ``sleep=`` do :func:`make_opener`.
SLEEP_FUNCTION: Callable[[float], None] = time.sleep


# ---------------------------------------------------------------------------
# Wyjątki
# ---------------------------------------------------------------------------


class ScrapeError(RuntimeError):
    """Błąd pobierania z portalu (nadklasa wszystkich błędów modułu)."""


class FetchError(ScrapeError):
    """Nie udało się pobrać adresu (HTTP lub sieć) po wszystkich próbach."""

    def __init__(self, message: str, *, url: str = "", status: Optional[int] = None):
        super().__init__(message)
        self.url = url
        self.status = status


class SiteBlockedError(ScrapeError):
    """Adres wskazuje poza witrynę (albo używa niedozwolonego schematu)."""


class NotXmlError(ScrapeError):
    """Pobrana treść nie jest XML-em (najczęściej strona logowania)."""


# ---------------------------------------------------------------------------
# Struktury danych z kontraktu
# ---------------------------------------------------------------------------


@dataclass
class AuctionRef:
    """Odnośnik do pojedynczej aukcji."""

    url: str
    id: str
    title: str


@dataclass
class XmlRef:
    """Odnośnik do pliku XML z zawartością pakietu."""

    url: str
    filename: str
    auction_id: str
    auction_title: str


# ---------------------------------------------------------------------------
# Pomocnicze: komunikaty
# ---------------------------------------------------------------------------


def _warn(message: str, *, stream=None) -> None:
    """Wypisuje ostrzeżenie na ``stderr`` (rozwiązywany w momencie wywołania)."""
    print("[scrape] " + message, file=stream if stream is not None else sys.stderr)


# ---------------------------------------------------------------------------
# Pomocnicze: adresy URL
# ---------------------------------------------------------------------------

_BAD_SCHEMES = (
    "javascript:", "mailto:", "tel:", "data:", "about:", "file:", "ftp:",
    "blob:", "sms:", "callto:", "ws:", "wss:",
)


def _normalize_url(url: str) -> str:
    """Kanonizuje adres: małe litery w hoście, bez fragmentu, bez ``..`` w ścieżce."""
    try:
        parts = urllib.parse.urlsplit(url)
        host = (parts.hostname or "").lower()
        port = parts.port
    except ValueError:
        return url
    scheme = parts.scheme.lower()
    default_port = {"http": 80, "https": 443}.get(scheme)
    netloc = host
    if port is not None and port != default_port:
        netloc = "%s:%d" % (host, port)
    path = parts.path or "/"
    if "./" in path or path.endswith(("/.", "/..")):
        trailing = path.endswith("/")
        path = posixpath.normpath(path)
        if trailing and not path.endswith("/"):
            path += "/"
    if not path.startswith("/"):
        path = "/" + path
    return urllib.parse.urlunsplit((scheme, netloc, path, parts.query, ""))


def absolutize(base_url: str, href: Optional[str]) -> Optional[str]:
    """Zamienia ``href`` (względny lub bezwzględny) na znormalizowany adres http(s).

    Zwraca ``None`` dla pustych kotwic i schematów, których nie wolno pobierać
    (``javascript:``, ``mailto:``, ``data:``, ``file:``...).
    """
    if not href:
        return None
    href = href.strip().replace("\n", "").replace("\r", "").replace("\t", "")
    if not href or href.startswith("#"):
        return None
    low = href.lower()
    for bad in _BAD_SCHEMES:
        if low.startswith(bad):
            return None
    try:
        url = urllib.parse.urljoin(base_url, href)
    except ValueError:
        return None
    try:
        scheme = urllib.parse.urlsplit(url).scheme.lower()
    except ValueError:
        return None
    if scheme not in ("http", "https"):
        return None
    return _normalize_url(url)


def _is_ip_literal(host: str) -> bool:
    """Czy host jest literałem IP (v4 lub v6)?"""
    if not host:
        return False
    if host.startswith("[") or ":" in host:
        return True
    return bool(re.fullmatch(r"\d{1,3}(?:\.\d{1,3}){3}", host))


def _registrable(host: str) -> str:
    """Przybliżenie domeny rejestrowalnej: dwie ostatnie etykiety."""
    labels = [p for p in host.split(".") if p]
    if len(labels) <= 2:
        return ".".join(labels)
    return ".".join(labels[-2:])


def is_same_site(base_url: str, url: str) -> bool:
    """Czy ``url`` należy do tej samej witryny co ``base_url``?

    Poddomeny tej samej domeny są akceptowane (``media.flexitauctions.com``),
    obce domeny nie.  Dla adresów IP wymagana jest zgodność hosta i portu.
    """
    try:
        a = urllib.parse.urlsplit(base_url)
        b = urllib.parse.urlsplit(url)
        ha, hb = (a.hostname or "").lower(), (b.hostname or "").lower()
        pa, pb = a.port, b.port
    except ValueError:
        return False
    if b.scheme.lower() not in ("http", "https"):
        return False
    if not ha or not hb:
        return False
    da = {"http": 80, "https": 443}.get(a.scheme.lower())
    db = {"http": 80, "https": 443}.get(b.scheme.lower())
    if ha == hb:
        if _is_ip_literal(ha):
            return (pa or da) == (pb or db)
        return True
    if _is_ip_literal(ha) or _is_ip_literal(hb):
        return False
    return bool(_registrable(ha)) and _registrable(ha) == _registrable(hb)


# ---------------------------------------------------------------------------
# Pomocnicze: nazwy plików
# ---------------------------------------------------------------------------

#: Nazwy zarezerwowane w Windows (nie mogą być nazwą pliku).
_RESERVED_NAMES = frozenset(
    ["con", "prn", "aux", "nul"]
    + ["com%d" % i for i in range(1, 10)]
    + ["lpt%d" % i for i in range(1, 10)]
)


def _cap_name(stem: str, ext: str) -> str:
    """Przycina nazwę do :data:`MAX_FILENAME_LEN`, dopisując skrót dla unikalności."""
    limit = max(8, MAX_FILENAME_LEN - len(ext))
    if len(stem) > limit:
        digest = hashlib.sha1(stem.encode("utf-8", "replace")).hexdigest()[:8]
        stem = stem[: limit - 9] + "-" + digest
    return stem + ext


def safe_filename(name: Optional[str], *, default: str = "plik.xml",
                  ensure_ext: str = ".xml") -> str:
    """Sprowadza dowolny tekst do bezpiecznej nazwy pliku.

    * bierze wyłącznie nazwę bazową (``../../etc/passwd`` -> ``passwd``),
      rozumiejąc oba rodzaje separatorów i procentowe kodowanie,
    * dopuszcza tylko ASCII ``[A-Za-z0-9_.-]``, resztę zamienia na ``_``,
    * usuwa wiodące kropki i ciągi ``..`` (brak path traversal, brak plików ukrytych),
    * pilnuje długości i nazw zarezerwowanych w Windows,
    * dokleja ``ensure_ext``, jeśli nazwa nie ma tego rozszerzenia
      (``ensure_ext=""`` wyłącza doklejanie).
    """
    text = name or ""
    if "%" in text:
        try:
            text = urllib.parse.unquote(text)
        except Exception:  # pragma: no cover - unquote praktycznie nie rzuca
            pass
    # NUL i znaki sterujące wycinamy zanim policzymy nazwę bazową
    text = "".join(ch for ch in text if ch >= " " and ch != "\x7f")
    text = text.replace("\\", "/")
    text = text.split("?", 1)[0].split("#", 1)[0]
    base = posixpath.basename(text).strip()
    base = re.sub(r"[^A-Za-z0-9_.-]", "_", base, flags=re.ASCII)
    base = re.sub(r"_{2,}", "_", base)           # bez ciągów podkreśleń
    base = re.sub(r"\.{2,}", ".", base)          # ".." nie przetrwa
    base = base.strip("._-")
    if not base:
        base = default or "plik.xml"
        base = re.sub(r"[^A-Za-z0-9_.-]", "_", base, flags=re.ASCII).strip("._-")
        if not base:
            base = "plik.xml"
    stem, ext = posixpath.splitext(base)
    if ensure_ext and ext.lower() != ensure_ext.lower():
        stem, ext = base, ensure_ext
    if not stem:
        stem = "plik"
    if stem.lower() in _RESERVED_NAMES:
        stem = "_" + stem
    return _cap_name(stem, ext)


def _unique_name(name: str, used) -> str:
    """Dokleja ``-2``, ``-3``... gdy nazwa już wystąpiła w bieżącym zestawie."""
    if name.lower() not in used:
        used.add(name.lower())
        return name
    stem, ext = posixpath.splitext(name)
    i = 2
    while True:
        cand = "%s-%d%s" % (stem, i, ext)
        if cand.lower() not in used:
            used.add(cand.lower())
            return cand
        i += 1


def _filename_for(url: str, *, auction_id: str = "", hint: Optional[str] = None) -> str:
    """Wylicza nazwę pliku dla adresu XML (z prefiksem aukcji, bezpieczna)."""
    parts = urllib.parse.urlsplit(url)
    path = urllib.parse.unquote(parts.path or "")
    candidate = ""
    if hint and posixpath.basename(hint.replace("\\", "/")).strip():
        candidate = hint
    if not candidate:
        base = posixpath.basename(path)
        if base.lower().endswith(".xml"):
            candidate = base
    if not candidate:
        for key, value in urllib.parse.parse_qsl(parts.query, keep_blank_values=False):
            if key.lower() in _FILENAME_PARAMS and value:
                candidate = value
                break
    if not candidate:
        base = posixpath.basename(path)
        extra = ""
        if parts.query:
            # różne parametry -> różne pliki; skrót zapytania trzyma je osobno
            extra = "-" + hashlib.sha1(parts.query.encode("utf-8")).hexdigest()[:6]
        candidate = (base or "batch") + extra
    name = safe_filename(candidate, default="batch.xml")
    prefix = safe_filename(auction_id or "", default="", ensure_ext="") if auction_id else ""
    if prefix and prefix.lower() not in name.lower():
        stem, ext = posixpath.splitext(name)
        # separator "-" (a nie "__"), bo safe_filename zwija ciągi podkreśleń,
        # a nazwa musi przetrwać PONOWNĄ sanityzację w download_xml bez zmian
        name = _cap_name(prefix + "-" + stem, ext)
    return name


# ---------------------------------------------------------------------------
# Opener
# ---------------------------------------------------------------------------


def _default_sleep(seconds: float) -> None:
    """Domyślny sen — czyta :data:`SLEEP_FUNCTION` przy KAŻDYM wywołaniu.

    Dzięki temu podmiana atrybutu modułu działa także dla openerów utworzonych
    wcześniej (wygodne w testach).
    """
    SLEEP_FUNCTION(seconds)


class _GuardedRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Blokuje przekierowania poza witrynę (ochrona ciasteczka sesyjnego)."""

    owner: Optional["Opener"] = None

    def redirect_request(self, req, fp, code, msg, headers, newurl):  # noqa: D102
        owner = self.owner
        if owner is not None and owner.site:
            if not is_same_site(owner.site, newurl):
                raise SiteBlockedError(
                    "Przekierowanie poza witrynę zablokowane: %s -> %s"
                    % (req.full_url, newurl)
                )
        return super().redirect_request(req, fp, code, msg, headers, newurl)


@dataclass
class Opener:
    """Konfiguracja + stan sesji HTTP.

    Obiekt zwracany przez :func:`make_opener`; wszystkie funkcje modułu
    przyjmują go jako pierwszy argument.
    """

    director: urllib.request.OpenerDirector
    user_agent: str = DEFAULT_USER_AGENT
    cookie: Optional[str] = None
    timeout: float = 30.0
    retries: int = 4
    delay: float = DEFAULT_DELAY
    max_bytes: int = DEFAULT_MAX_BYTES
    #: Adres bazowy witryny (ustawiany przy pierwszym użyciu, patrz ``bind_site``).
    site: Optional[str] = None
    diagnose: bool = False
    extra_headers: dict = field(default_factory=dict)
    cookie_jar: Optional[http.cookiejar.CookieJar] = None
    #: Funkcja usypiająca (backoff + uprzejme opóźnienie) — wstrzykiwalna.
    sleep: Callable[[float], None] = _default_sleep
    clock: Callable[[], float] = time.monotonic
    stats: dict = field(default_factory=lambda: {
        "requests": 0, "retries": 0, "bytes": 0, "blocked": 0, "probes": 0,
    })
    #: Ostatnio pobrana strona HTML: ``(url, dane, content_type)`` — dla diagnostyki.
    last_page: Optional[tuple] = None
    _ctype_cache: dict = field(default_factory=dict)
    _last_request_at: Optional[float] = None

    # -- pomocnicze ---------------------------------------------------------

    def headers(self) -> dict:
        """Nagłówki wysyłane z każdym żądaniem."""
        head = {
            "User-Agent": self.user_agent,
            "Accept": (
                "text/html,application/xhtml+xml,application/xml;q=0.9,"
                "text/xml;q=0.9,*/*;q=0.8"
            ),
            "Accept-Language": "pl,en;q=0.8",
            # urllib nie umie sam rozpakować gzip — prosimy o brak kompresji
            "Accept-Encoding": "identity",
            "Connection": "close",
        }
        if self.cookie:
            head["Cookie"] = self.cookie
        head.update(self.extra_headers or {})
        return head

    def bind_site(self, url: str) -> None:
        """Przypina opener do witryny (pierwszy adres wygrywa)."""
        if not self.site and url:
            self.site = _normalize_url(url)

    def allows(self, url: str) -> bool:
        """Czy wolno pobrać ten adres?"""
        try:
            scheme = urllib.parse.urlsplit(url).scheme.lower()
        except ValueError:
            return False
        if scheme not in ("http", "https"):
            return False
        if not self.site:
            return True
        return is_same_site(self.site, url)

    def pace(self) -> None:
        """Uprzejme opóźnienie między żądaniami (używa wstrzykniętego snu)."""
        if self.delay <= 0:
            return
        now = self.clock()
        if self._last_request_at is not None:
            wait = self.delay - (now - self._last_request_at)
            if wait > 0:
                self.sleep(wait)
        self._last_request_at = self.clock()

    def backoff(self, attempt: int) -> float:
        """Opóźnienie przed próbą numer ``attempt`` (0-indeksowaną): 2, 4, 8, 16 s."""
        return min(MAX_BACKOFF, BACKOFF_BASE * (BACKOFF_FACTOR ** attempt))


def make_opener(*, user_agent: str = DEFAULT_USER_AGENT, cookie: Optional[str] = None,
                timeout: float = 30.0, retries: int = 4,
                delay: float = DEFAULT_DELAY,
                sleep: Optional[Callable[[float], None]] = None,
                clock: Optional[Callable[[], float]] = None,
                site: Optional[str] = None,
                max_bytes: int = DEFAULT_MAX_BYTES,
                diagnose: bool = False,
                extra_headers: Optional[dict] = None) -> Any:
    """Buduje sesję HTTP (ciasteczka, nagłówki, limity, backoff).

    :param cookie: surowa wartość nagłówka ``Cookie`` (np. skopiowana z
        przeglądarki), gdy portal wymaga zalogowania.
    :param retries: liczba PONOWIEŃ po nieudanej pierwszej próbie.
    :param sleep: funkcja usypiająca — wstrzykiwana w testach, żeby backoff
        nie trwał naprawdę.  Domyślnie modułowa :data:`SLEEP_FUNCTION`.
    :param site: adres witryny, poza którą nie wolno wyjść; gdy pominięty,
        opener przypina się do pierwszego pobieranego adresu.
    """
    jar = http.cookiejar.CookieJar()
    redirect = _GuardedRedirectHandler()
    director = urllib.request.build_opener(
        urllib.request.HTTPCookieProcessor(jar),
        redirect,
    )
    # build_opener dokleja własny User-Agent — usuwamy, nagłówki dajemy sami
    director.addheaders = []
    opener = Opener(
        director=director,
        user_agent=user_agent,
        cookie=cookie,
        timeout=float(timeout),
        retries=max(0, int(retries)),
        delay=max(0.0, float(delay)),
        max_bytes=int(max_bytes),
        site=_normalize_url(site) if site else None,
        diagnose=bool(diagnose),
        extra_headers=dict(extra_headers or {}),
        cookie_jar=jar,
    )
    if sleep is not None:
        opener.sleep = sleep
    if clock is not None:
        opener.clock = clock
    redirect.owner = opener
    return opener


# ---------------------------------------------------------------------------
# Pobieranie
# ---------------------------------------------------------------------------


def _content_type(headers) -> str:
    """Wyciąga sam typ MIME (bez parametrów), małymi literami."""
    raw = ""
    try:
        raw = headers.get("Content-Type", "") or ""
    except AttributeError:  # pragma: no cover - nietypowy obiekt nagłówków
        raw = ""
    return raw.split(";", 1)[0].strip().lower()


def _read_limited(response, limit: int) -> bytes:
    """Czyta odpowiedź kawałkami, pilnując limitu rozmiaru."""
    chunks = []
    total = 0
    while True:
        chunk = response.read(65536)
        if not chunk:
            break
        total += len(chunk)
        if limit and total > limit:
            raise ScrapeError(
                "Odpowiedź przekracza limit %d bajtów (podnieś max_bytes)" % limit
            )
        chunks.append(chunk)
    return b"".join(chunks)


def _retry_after(headers, fallback: float) -> float:
    """Honoruje nagłówek ``Retry-After`` (tylko postać sekundowa)."""
    try:
        raw = (headers.get("Retry-After", "") or "").strip()
    except AttributeError:
        return fallback
    if not raw:
        return fallback
    try:
        seconds = float(raw)
    except ValueError:
        return fallback          # postać datowa — nie kombinujemy, backoff jak zwykle
    if seconds < 0:
        return fallback
    return min(MAX_BACKOFF, seconds)


def fetch_full(opener, url: str) -> tuple:
    """Jak :func:`fetch`, ale zwraca też nagłówki: ``(dane, content_type, headers)``."""
    url = _normalize_url(url)
    if not opener.allows(url):
        opener.stats["blocked"] += 1
        raise SiteBlockedError("Adres spoza witryny (albo zły schemat): %s" % url)
    last_message = ""
    attempts = opener.retries + 1
    for attempt in range(attempts):
        opener.pace()
        try:
            request = urllib.request.Request(url, headers=opener.headers())
            response = opener.director.open(request, timeout=opener.timeout)
            try:
                data = _read_limited(response, opener.max_bytes)
                ctype = _content_type(response.headers)
                headers = response.headers
            finally:
                response.close()
            opener.stats["requests"] += 1
            opener.stats["bytes"] += len(data)
            return data, ctype, headers
        except urllib.error.HTTPError as exc:
            opener.stats["requests"] += 1
            status = exc.code
            try:
                exc.read()
            except Exception:  # pragma: no cover
                pass
            finally:
                exc.close()
            last_message = "HTTP %s %s" % (status, exc.reason)
            if status in RETRY_STATUSES and attempt < attempts - 1:
                opener.stats["retries"] += 1
                opener.sleep(_retry_after(exc.headers, opener.backoff(attempt)))
                continue
            raise FetchError(
                "Nie udało się pobrać %s: %s" % (url, last_message),
                url=url, status=status,
            ) from exc
        except (urllib.error.URLError, socket.timeout, TimeoutError,
                http.client.HTTPException, ConnectionError, OSError) as exc:
            opener.stats["requests"] += 1
            reason = getattr(exc, "reason", exc)
            last_message = "%s: %s" % (type(exc).__name__, reason)
            if attempt < attempts - 1:
                opener.stats["retries"] += 1
                opener.sleep(opener.backoff(attempt))
                continue
            raise FetchError(
                "Nie udało się pobrać %s po %d próbach: %s"
                % (url, attempts, last_message),
                url=url,
            ) from exc
    raise FetchError(  # pragma: no cover - pętla zawsze kończy się wcześniej
        "Nie udało się pobrać %s: %s" % (url, last_message), url=url
    )


def fetch(opener, url: str) -> tuple:
    """Pobiera adres i zwraca ``(treść, content_type)``.

    Ponawia próby przy błędach sieci i kodach 408/425/429/5xx z wykładniczym
    backoffem 2, 4, 8, 16 s.  Przy trwałej porażce rzuca :class:`FetchError`.
    """
    data, ctype, _headers = fetch_full(opener, url)
    return data, ctype


# ---------------------------------------------------------------------------
# Parsowanie HTML
# ---------------------------------------------------------------------------


@dataclass
class Link:
    """Pojedynczy odnośnik znaleziony na stronie."""

    url: str                      # bezwzględny, znormalizowany
    text: str = ""
    tag: str = "a"
    attrs: dict = field(default_factory=dict)
    source: str = "html"          # "html" albo "script" (osadzony JSON/JS)

    @property
    def rel(self) -> str:
        return (self.attrs.get("rel") or "").lower()

    def haystack(self) -> str:
        """Tekst + istotne atrybuty — do dopasowywania słów kluczowych."""
        bits = [self.text, self.url]
        for key in ("title", "aria-label", "class", "id", "download", "rel", "alt"):
            value = self.attrs.get(key)
            if value:
                bits.append(str(value))
        return " ".join(bits)


class _PageParser(HTMLParser):
    """Zbiera linki, tytuł i treść skryptów.  Nigdy nie rzuca wyjątkiem."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.raw_links: list = []       # (tag, attrs, text)
        self.base_href: Optional[str] = None
        self.scripts: list = []
        self.title: str = ""
        self.forms: list = []
        self._anchor: Optional[list] = None   # [tag, attrs, [teksty]]
        self._in_script = 0
        self._script_buf: list = []
        self._in_title = False

    # -- zdarzenia parsera --------------------------------------------------

    def handle_starttag(self, tag, attrs):
        adict = {}
        for key, value in attrs:
            if key is None:
                continue
            adict[key.lower()] = value if value is not None else ""
        if tag == "base" and adict.get("href"):
            self.base_href = adict["href"]
        elif tag == "title":
            self._in_title = True
        elif tag in ("script", "template"):
            self._in_script += 1
            self._script_buf = []
        elif tag == "form":
            self.forms.append(adict)
        if tag in ("a", "area"):
            self._close_anchor()
            self._anchor = [tag, adict, []]
            if adict.get("alt"):
                self._anchor[2].append(adict["alt"])
            return
        if tag == "img" and self._anchor is not None and adict.get("alt"):
            self._anchor[2].append(adict["alt"])
        # linki poza <a>: <link rel=next>, data-* na dowolnym elemencie
        for attr in _URL_ATTRS:
            if adict.get(attr):
                self.raw_links.append((tag, adict, adict.get("title", "")))
                break

    def handle_startendtag(self, tag, attrs):
        self.handle_starttag(tag, attrs)
        if tag in ("a", "area"):
            self._close_anchor()
        elif tag in ("script", "template"):
            self._in_script = max(0, self._in_script - 1)

    def handle_endtag(self, tag):
        if tag in ("a", "area"):
            self._close_anchor()
        elif tag in ("script", "template"):
            if self._in_script:
                self._in_script -= 1
                if self._script_buf:
                    self.scripts.append("".join(self._script_buf))
                self._script_buf = []
        elif tag == "title":
            self._in_title = False

    def handle_data(self, data):
        if self._in_script:
            self._script_buf.append(data)
            return
        if self._in_title:
            self.title += data
        if self._anchor is not None:
            self._anchor[2].append(data)

    def error(self, message):  # pragma: no cover - Python 3 i tak nie woła
        pass

    # -- finalizacja --------------------------------------------------------

    def _close_anchor(self):
        if self._anchor is None:
            return
        tag, adict, texts = self._anchor
        self._anchor = None
        self.raw_links.append((tag, adict, " ".join(texts)))

    def close(self):
        super().close()
        self._close_anchor()
        if self._script_buf:
            self.scripts.append("".join(self._script_buf))
            self._script_buf = []


@dataclass
class Page:
    """Sparsowana strona: linki (z HTML i ze skryptów) plus metadane."""

    url: str
    links: list = field(default_factory=list)
    title: str = ""
    text_len: int = 0
    scripts: int = 0
    encoding: str = "utf-8"
    forms: int = 0


_CHARSET_RE = re.compile(rb"""charset\s*=\s*["']?\s*([A-Za-z0-9_.:+-]+)""", re.IGNORECASE)

_RAW_URL_RE = re.compile(
    r"""["'(]\s*((?:https?://|/)[^\s"'()<>]{1,400})""",
    re.IGNORECASE,
)


def _decode_html(data: bytes, content_type: str = "") -> tuple:
    """Dekoduje HTML: BOM -> nagłówek -> ``<meta charset>`` -> UTF-8 (z podmianą)."""
    if data.startswith(b"\xef\xbb\xbf"):
        return data[3:].decode("utf-8", "replace"), "utf-8-sig"
    if data.startswith((b"\xff\xfe", b"\xfe\xff")):
        return data.decode("utf-16", "replace"), "utf-16"
    encoding = ""
    match = _CHARSET_RE.search((content_type or "").encode("ascii", "replace"))
    if match:
        encoding = match.group(1).decode("ascii", "replace")
    if not encoding:
        head = data[:4096]
        match = _CHARSET_RE.search(head)
        if match:
            encoding = match.group(1).decode("ascii", "replace")
    for candidate in (encoding, "utf-8"):
        if not candidate:
            continue
        try:
            return data.decode(candidate), candidate.lower()
        except (LookupError, UnicodeDecodeError):
            continue
    return data.decode("utf-8", "replace"), "utf-8"


def _unescape_js(text: str) -> str:
    """Odwraca ucieczki spotykane w osadzonym JSON: ``\\/`` i ``\\u002F``."""
    return (
        text.replace("\\/", "/")
        .replace("\\u002F", "/")
        .replace("\\u002f", "/")
        .replace("\\u0026", "&")
        .replace("&amp;", "&")
    )


def parse_page(data, url: str, content_type: str = "") -> Page:
    """Parsuje stronę na obiekt :class:`Page` (linki bezwzględne, dedupilikowane)."""
    if isinstance(data, bytes):
        text, encoding = _decode_html(data, content_type)
    else:
        text, encoding = str(data), "str"
    parser = _PageParser()
    try:
        parser.feed(text)
        parser.close()
    except Exception as exc:  # pragma: no cover - html.parser jest wyrozumiały
        _warn("Nietypowy HTML na %s (%s) — używam tylko skanu tekstowego" % (url, exc))
    base = url
    if parser.base_href:
        candidate = absolutize(url, parser.base_href)
        if candidate:
            base = candidate
    page = Page(url=url, title=re.sub(r"\s+", " ", parser.title).strip(),
                text_len=len(text), scripts=len(parser.scripts), encoding=encoding,
                forms=len(parser.forms))
    seen = set()
    for tag, attrs, anchor_text in parser.raw_links:
        href = ""
        for attr in _URL_ATTRS:
            if attrs.get(attr):
                href = attrs[attr]
                break
        absolute = absolutize(base, _unescape_js(href))
        if not absolute:
            continue
        clean_text = re.sub(r"\s+", " ", anchor_text or "").strip()
        key = (absolute, clean_text, tag)
        if key in seen:
            continue
        seen.add(key)
        page.links.append(Link(url=absolute, text=clean_text, tag=tag, attrs=attrs))
    # skan surowej treści — portal może być SPA i trzymać adresy w JSON-ie
    known = {link.url for link in page.links}
    haystacks = list(parser.scripts)
    haystacks.append(text)
    for blob in haystacks:
        for match in _RAW_URL_RE.finditer(_unescape_js(blob)):
            absolute = absolutize(base, match.group(1))
            if not absolute or absolute in known:
                continue
            known.add(absolute)
            page.links.append(Link(url=absolute, text="", tag="script",
                                   attrs={}, source="script"))
    return page


# ---------------------------------------------------------------------------
# Diagnostyka
# ---------------------------------------------------------------------------


def describe_page(data, url: str = "", *, content_type: str = "", stream=None,
                  limit: int = 20, pattern=None) -> str:
    """Opisuje, CO właściwie jest na stronie — gdy nic nie pasuje do wzorców.

    Zwraca raport (tekst) i — gdy podano ``stream`` — wypisuje go tam.
    Nigdy nie rzuca wyjątkiem, nawet dla śmieci zamiast HTML-a.
    """
    lines = ["--- diagnostyka strony: %s ---" % (url or "(bez adresu)")]
    try:
        page = parse_page(data, url or "http://example.invalid/", content_type)
    except Exception as exc:  # pragma: no cover - parse_page łapie swoje błędy
        lines.append("Nie udało się sparsować strony: %s" % exc)
        report = "\n".join(lines)
        if stream is not None:
            print(report, file=stream)
        return report
    same, other = [], []
    for link in page.links:
        (same if (not url or is_same_site(url, link.url)) else other).append(link)
    lines.append("Tytuł strony: %s" % (page.title or "(brak)"))
    lines.append("Typ treści: %s, kodowanie: %s, długość: %d znaków"
                 % (content_type or "(nieznany)", page.encoding, page.text_len))
    lines.append("Linków razem: %d (w tej witrynie: %d, poza witryną: %d), "
                 "bloków <script>: %d, formularzy: %d"
                 % (len(page.links), len(same), len(other), page.scripts, page.forms))
    if pattern is not None:
        rx = re.compile(pattern) if isinstance(pattern, str) else pattern
        hits = [link for link in same if rx.search(urllib.parse.urlsplit(link.url).path)]
        lines.append("Pasujących do wzorca %r: %d" % (getattr(rx, "pattern", pattern), len(hits)))
    xmlish = [link for link in page.links if _url_mentions_xml(link.url)]
    if xmlish:
        lines.append("Adresy wyglądające na XML: %d" % len(xmlish))
        for link in xmlish[:limit]:
            lines.append("  XML? %s" % link.url)
    if not page.links:
        lines.append("UWAGA: nie znaleziono ŻADNEGO linku — strona może być "
                     "aplikacją JS (SPA) albo wymagać zalogowania (--cookie).")
    lines.append("Przykładowe linki (max %d):" % limit)
    for link in page.links[:limit]:
        label = (link.text[:60] + "…") if len(link.text) > 60 else link.text
        lines.append("  [%s] %s  <- %s" % (link.source, link.url, label or "(bez tekstu)"))
    if len(page.links) > limit:
        lines.append("  ... oraz %d dalszych" % (len(page.links) - limit))
    if other:
        lines.append("Odrzucone (obca domena), max %d:" % limit)
        for link in other[:limit]:
            lines.append("  %s" % link.url)
    lines.append("--- koniec diagnostyki ---")
    report = "\n".join(lines)
    if stream is not None:
        print(report, file=stream)
    return report


def diagnose_last(opener, *, stream=None, limit: int = 20) -> str:
    """Diagnostyka OSTATNIEJ pobranej strony (wygodne dla CLI po pustym wyniku)."""
    if not getattr(opener, "last_page", None):
        report = "[scrape] Brak zapamiętanej strony do diagnostyki."
        if stream is not None:
            print(report, file=stream)
        return report
    url, data, ctype = opener.last_page
    return describe_page(data, url, content_type=ctype, stream=stream, limit=limit)


def _maybe_diagnose(opener, what: str) -> None:
    """Wypisuje ostrzeżenie (i pełną diagnostykę, gdy ``diagnose=True``)."""
    _warn("Nic nie znaleziono: %s" % what)
    if getattr(opener, "diagnose", False):
        diagnose_last(opener, stream=sys.stderr)
    else:
        _warn("Uruchom z diagnostyką (make_opener(diagnose=True) / --diagnose), "
              "aby zobaczyć, co jest na stronie.")


# ---------------------------------------------------------------------------
# Odkrywanie aukcji
# ---------------------------------------------------------------------------

_PAGE_PARAM_RE = re.compile(r"(?:^|[?&])(page|p|strona|pageNumber|pageIndex)=(\d+)",
                            re.IGNORECASE)
#: Teksty linku oznaczające "następna strona" (bez samych cyfr — te akceptujemy
#: wyłącznie razem z parametrem ``page=``, żeby nie łazić po całym serwisie).
_NEXT_WORD_RE = re.compile(r"^\s*(next|nast(ę|e)pn\w*|dalej|wi(ę|e)cej|»|›|>>?)\s*$",
                           re.IGNORECASE)


def _remember_page(opener, url: str, data: bytes, ctype: str) -> None:
    """Zapamiętuje stronę do późniejszej diagnostyki (z limitem rozmiaru)."""
    try:
        opener.last_page = (url, data[:512 * 1024], ctype)
    except Exception:  # pragma: no cover - opener może być atrapą bez atrybutu
        pass


def _auction_id(match, url: str) -> str:
    """Wyciąga identyfikator aukcji z dopasowania regexu albo ze ścieżki."""
    ident = ""
    if match is not None:
        ident = (match.groupdict().get("id") or "") if match.groupdict() else ""
        if not ident and match.groups():
            ident = match.group(1) or ""
    if not ident:
        path = urllib.parse.urlsplit(url).path.rstrip("/")
        parts = [p for p in path.split("/") if p]
        while parts and parts[-1].lower() in _NOT_LOT_SEGMENTS:
            parts.pop()
        ident = parts[-1] if parts else url
    return ident.strip("/")


def _canonical_score(url: str) -> tuple:
    """Im mniejsza krotka, tym bardziej "kanoniczny" adres aukcji."""
    path = urllib.parse.urlsplit(url).path.rstrip("/")
    return (len([p for p in path.split("/") if p]), len(url))


def discover_auctions(opener, base_url: str, *, auction_re: Optional[str] = None,
                      max_pages: int = 50) -> list:
    """Zbiera listę aukcji, przechodząc paginację listy.

    Zwraca listę :class:`AuctionRef` w kolejności odkrycia, bez duplikatów
    (po identyfikatorze aukcji).  Gdy nic nie pasuje — pustą listę i ostrzeżenie
    na ``stderr``; wyjątek leci tylko wtedy, gdy nie da się pobrać PIERWSZEJ
    strony listy.
    """
    opener.bind_site(base_url)
    pattern = re.compile(auction_re) if auction_re else re.compile(DEFAULT_AUCTION_RE)
    start = _normalize_url(base_url)
    queue = [start]
    visited = set()
    found = {}
    pages_done = 0
    while queue and pages_done < max(1, int(max_pages)):
        page_url = queue.pop(0)
        if page_url in visited:
            continue
        visited.add(page_url)
        try:
            data, ctype = fetch(opener, page_url)
        except ScrapeError as exc:
            if pages_done == 0:
                raise
            _warn("Pomijam stronę listy %s (%s)" % (page_url, exc))
            continue
        pages_done += 1
        _remember_page(opener, page_url, data, ctype)
        page = parse_page(data, page_url, ctype)
        for link in page.links:
            if not is_same_site(page_url, link.url):
                continue
            path = urllib.parse.urlsplit(link.url).path
            match = pattern.search(path)
            if not match:
                continue
            ident = _auction_id(match, link.url)
            if not ident:
                continue
            title = link.text or link.attrs.get("title") or link.attrs.get("aria-label") or ""
            title = re.sub(r"\s+", " ", title).strip() or ident
            existing = found.get(ident)
            if existing is None:
                found[ident] = AuctionRef(url=link.url, id=ident, title=title)
            else:
                # ten sam identyfikator pod kilkoma adresami (np. /info) —
                # zostawiamy najbardziej kanoniczny adres i pierwszy sensowny tytuł
                if _canonical_score(link.url) < _canonical_score(existing.url):
                    existing.url = link.url
                if existing.title == existing.id and title != ident:
                    existing.title = title
        for candidate in _next_page_urls(page, page_url):
            if candidate not in visited and candidate not in queue:
                queue.append(candidate)
    if not found:
        _maybe_diagnose(opener, "brak aukcji na %s (wzorzec %r)"
                        % (base_url, pattern.pattern))
    return list(found.values())


def _next_page_urls(page: Page, page_url: str) -> list:
    """Znajduje adresy kolejnych stron listy (``rel=next``, ``?page=N``, "Dalej")."""
    result = []
    current_path = urllib.parse.urlsplit(page_url).path
    current_no = 0
    match = _PAGE_PARAM_RE.search(page_url)
    if match:
        current_no = int(match.group(2))
    for link in page.links:
        if not is_same_site(page_url, link.url):
            continue
        if link.url == page_url:
            continue
        parts = urllib.parse.urlsplit(link.url)
        page_match = _PAGE_PARAM_RE.search(link.url)
        same_path = parts.path == current_path
        number = int(page_match.group(2)) if page_match else None
        # link "wstecz" (np. "1" albo "Poprzednia") — pomijamy, żeby nie chodzić w kółko
        if page_match and same_path and number is not None and number <= current_no:
            continue
        is_next_rel = "next" in link.rel
        looks_next_word = bool(_NEXT_WORD_RE.match(link.text or ""))
        if is_next_rel or page_match is not None or looks_next_word:
            if link.url not in result:
                result.append(link.url)
    return result


# ---------------------------------------------------------------------------
# Szukanie linków XML
# ---------------------------------------------------------------------------


def _url_mentions_xml(url: str) -> bool:
    """Czy adres sam w sobie zdradza XML (rozszerzenie, ścieżka albo parametr)?"""
    parts = urllib.parse.urlsplit(url)
    path = urllib.parse.unquote(parts.path or "").lower()
    if path.endswith(".xml"):
        return True
    if re.search(r"(^|[/._-])xml([/._-]|$)", path):
        return True
    query = urllib.parse.unquote(parts.query or "").lower()
    for key, value in urllib.parse.parse_qsl(query, keep_blank_values=True):
        if "xml" in key or "xml" in value:
            return True
    return False


def _has_boring_ext(url: str) -> bool:
    """Czy adres ma rozszerzenie, którego na pewno nie warto sondować?"""
    path = urllib.parse.unquote(urllib.parse.urlsplit(url).path or "").lower()
    return path.endswith(_BORING_EXT)


def _looks_like_xml(data: bytes) -> bool:
    """Czy początek treści wygląda na XML (a nie na HTML)?"""
    if not data:
        return False
    head = data.lstrip(b"\xef\xbb\xbf").lstrip()[:512].lower()
    if not head.startswith(b"<"):
        return False
    if head.startswith(b"<!doctype html") or head.startswith(b"<html"):
        return False
    return True


@dataclass
class _Candidate:
    url: str
    hint: Optional[str]
    certain: bool
    why: str


def _xml_candidates(page: Page, opener) -> list:
    """Wybiera z linków strony kandydatów na plik XML (pewnych i do sondowania)."""
    candidates = []
    seen = set()
    for link in page.links:
        url = link.url
        if url in seen:
            continue
        if not opener.allows(url):
            continue
        hint = link.attrs.get("download") or None
        if _url_mentions_xml(url):
            seen.add(url)
            candidates.append(_Candidate(url, hint, True, "adres wskazuje XML"))
            continue
        if hint and hint.lower().endswith(".xml"):
            seen.add(url)
            candidates.append(_Candidate(url, hint, True, "atrybut download=*.xml"))
            continue
        if _has_boring_ext(url):
            continue
        reason = ""
        if "download" in link.attrs:
            reason = "atrybut download"
        elif link.source == "html" and _XML_HINT_RE.search(link.haystack()):
            reason = "słowo kluczowe w linku"
        if reason:
            seen.add(url)
            candidates.append(_Candidate(url, hint, False, reason))
    return candidates


def _confirm_by_content_type(opener, url: str, fetch_fn) -> bool:
    """Sonduje adres i mówi, czy to XML (wynik zapamiętany w openerze)."""
    cache = getattr(opener, "_ctype_cache", None)
    if cache is not None and url in cache:
        return cache[url]
    verdict = False
    try:
        data, ctype = fetch_fn(opener, url)
        opener.stats["probes"] = opener.stats.get("probes", 0) + 1
        ctype = (ctype or "").split(";", 1)[0].strip().lower()
        if ctype in XML_CONTENT_TYPES or ctype.endswith("+xml"):
            verdict = True
        elif ctype in AMBIGUOUS_CONTENT_TYPES and _looks_like_xml(data):
            verdict = True
    except ScrapeError as exc:
        _warn("Nie udało się sprawdzić %s (%s)" % (url, exc))
        verdict = False
    if cache is not None:
        cache[url] = verdict
    return verdict


def _lot_urls(page: Page, auction: AuctionRef, opener) -> list:
    """Znajduje adresy stron lotów na stronie aukcji (z deduplikacją po hashu)."""
    by_hash = {}
    generic = []
    auction_path = urllib.parse.urlsplit(auction.url).path.rstrip("/")
    if auction_path.endswith("/info"):
        auction_path = auction_path[: -len("/info")]
    for link in page.links:
        if not opener.allows(link.url) or not is_same_site(page.url, link.url):
            continue
        path = urllib.parse.urlsplit(link.url).path
        match = _LOT_RE.match(path.rstrip("/") or "/")
        if match:
            key = match.group("hash").lower()
            current = by_hash.get(key)
            if current is None or _canonical_score(link.url) < _canonical_score(current):
                by_hash[key] = link.url
            continue
        # zapasowo: dokładnie jeden segment poniżej ścieżki aukcji
        if auction_path and path.startswith(auction_path + "/"):
            rest = path[len(auction_path) + 1:].strip("/")
            if rest and "/" not in rest and rest.lower() not in _NOT_LOT_SEGMENTS:
                if link.url not in generic:
                    generic.append(link.url)
    urls = list(by_hash.values())
    for url in generic:
        if url not in urls:
            urls.append(url)
    return urls


def find_xml_links(opener, auction: AuctionRef, *, opener_fetch=None,
                   descend: str = "auto", max_lots: int = 200,
                   probe_limit: int = 25) -> list:
    """Znajduje pliki XML "zawartość pakietu" dla jednej aukcji.

    Sprawdza stronę aukcji, a gdy nie ma tam XML-i (``descend="auto"``) — także
    strony poszczególnych lotów.  ``descend="always"`` schodzi zawsze,
    ``descend="never"`` nigdy.

    :param opener_fetch: podmiana funkcji :func:`fetch` (testy, cache).
    :param max_lots: górny limit odwiedzanych stron lotów.
    :param probe_limit: ile najwyżej linków wolno sprawdzić przez Content-Type.
    """
    fetch_fn = opener_fetch or fetch
    opener.bind_site(auction.url)
    used_names = set()
    results = []
    seen_urls = set()
    probes_left = max(0, int(probe_limit))

    def collect(page: Page) -> int:
        """Dokłada do wyniku XML-e znalezione na jednej stronie."""
        nonlocal probes_left
        added = 0
        for cand in _xml_candidates(page, opener):
            if cand.url in seen_urls:
                continue
            if not cand.certain:
                if probes_left <= 0:
                    _warn("Wyczerpany budżet sondowania — pomijam %s" % cand.url)
                    continue
                probes_left -= 1
                if not _confirm_by_content_type(opener, cand.url, fetch_fn):
                    continue
            seen_urls.add(cand.url)
            name = _unique_name(
                _filename_for(cand.url, auction_id=auction.id, hint=cand.hint),
                used_names,
            )
            results.append(XmlRef(url=cand.url, filename=name,
                                  auction_id=auction.id,
                                  auction_title=auction.title))
            added += 1
        return added

    try:
        data, ctype = fetch_fn(opener, auction.url)
    except ScrapeError as exc:
        _warn("Nie udało się pobrać strony aukcji %s (%s)" % (auction.url, exc))
        return []
    _remember_page(opener, auction.url, data, ctype)
    if _looks_like_xml(data) and _url_mentions_xml(auction.url):
        # sama "aukcja" bywa plikiem XML — nie ma czego parsować
        name = _unique_name(_filename_for(auction.url, auction_id=auction.id), used_names)
        return [XmlRef(url=auction.url, filename=name, auction_id=auction.id,
                       auction_title=auction.title)]
    page = parse_page(data, auction.url, ctype)
    direct = collect(page)

    mode = (descend or "auto").lower()
    should_descend = mode == "always" or (mode == "auto" and direct == 0)
    if should_descend and mode != "never":
        lots = _lot_urls(page, auction, opener)
        if len(lots) > max_lots:
            _warn("Aukcja %s ma %d lotów — ograniczam do %d (max_lots)"
                  % (auction.id, len(lots), max_lots))
            lots = lots[:max_lots]
        for lot_url in lots:
            if lot_url in seen_urls:
                continue
            try:
                lot_data, lot_ctype = fetch_fn(opener, lot_url)
            except ScrapeError as exc:
                _warn("Pomijam lot %s (%s)" % (lot_url, exc))
                continue
            _remember_page(opener, lot_url, lot_data, lot_ctype)
            if _looks_like_xml(lot_data) and _url_mentions_xml(lot_url):
                seen_urls.add(lot_url)
                name = _unique_name(
                    _filename_for(lot_url, auction_id=auction.id), used_names)
                results.append(XmlRef(url=lot_url, filename=name,
                                      auction_id=auction.id,
                                      auction_title=auction.title))
                continue
            collect(parse_page(lot_data, lot_url, lot_ctype))

    if not results:
        _maybe_diagnose(opener, "brak linków XML w aukcji %s (%s)"
                        % (auction.id, auction.url))
    return results


# ---------------------------------------------------------------------------
# Pobieranie plików
# ---------------------------------------------------------------------------


def download_xml(opener, ref: XmlRef, out_dir, *, overwrite: bool = False,
                 fetch_fn=None, verify_xml: bool = True) -> str:
    """Pobiera jeden plik XML do ``out_dir`` i zwraca ścieżkę zapisanego pliku.

    * nazwa pliku jest sanityzowana (brak path traversal, brak ``..``),
    * gdy plik już istnieje i ``overwrite=False`` — NIE pobiera niczego,
    * zapis jest atomowy (``.part`` + ``os.replace``), więc przerwane pobranie
      nie zostawia uszkodzonego pliku, który później zostałby pominięty,
    * gdy odpowiedź jest stroną HTML (ściana logowania) — :class:`NotXmlError`.
    """
    out_dir = os.fspath(out_dir)
    os.makedirs(out_dir, exist_ok=True)
    fallback = safe_filename(
        "%s-%s" % (ref.auction_id or "aukcja",
                   hashlib.sha1((ref.url or "").encode("utf-8")).hexdigest()[:8]),
        default="batch.xml",
    )
    name = safe_filename(ref.filename, default=fallback)
    path = os.path.join(out_dir, name)
    root = os.path.realpath(out_dir)
    if os.path.dirname(os.path.realpath(path)) != root:
        raise ScrapeError("Nazwa pliku wyprowadza poza katalog docelowy: %r" % ref.filename)
    if os.path.exists(path) and not overwrite:
        return path
    data, ctype = (fetch_fn or fetch)(opener, ref.url)
    if verify_xml:
        if not data.strip():
            raise NotXmlError("Pusta odpowiedź dla %s" % ref.url)
        if not _looks_like_xml(data):
            head = data.lstrip()[:200].decode("utf-8", "replace")
            raise NotXmlError(
                "Treść spod %s nie jest XML-em (Content-Type: %s). "
                "Portal mógł zwrócić stronę logowania — spróbuj z opcją --cookie.\n"
                "Początek odpowiedzi: %s" % (ref.url, ctype or "?", head)
            )
    tmp = path + ".part"
    try:
        with open(tmp, "wb") as handle:
            handle.write(data)
        os.replace(tmp, path)
    finally:
        if os.path.exists(tmp):
            try:
                os.remove(tmp)
            except OSError:  # pragma: no cover
                pass
    return path
