# -*- coding: utf-8 -*-
"""Testy pobieracza (``flexit2xlsx.scrape``) na ATRAPIE portalu.

Portal flexitauctions.com jest nieosiągalny z tej sesji, więc testy serwują
statyczne strony z ``tests/mocksite/`` przez ``http.server`` na losowym wolnym
porcie (wątek w tle).  Atrapa naśladuje realny układ portalu:

    lista aukcji (paginacja)  ->  strona aukcji  ->  strony lotów  ->  XML

Testy są SZYBKIE: backoff i uprzejme opóźnienia idą przez wstrzykniętą
funkcję :class:`FakeClock.sleep`, która nic nie usypia, tylko zapisuje żądane
opóźnienia i przesuwa sztuczny zegar.

Uruchamiane przez ``python3 -m unittest`` oraz ``python3 -m pytest``.
"""

from __future__ import annotations

import ast
import contextlib
import gzip
import io
import json
import os
import shutil
import socket
import sys
import tempfile
import threading
import time
import unittest
import zlib
import xml.etree.ElementTree as ET
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from flexit2xlsx import scrape  # noqa: E402
from flexit2xlsx.scrape import (  # noqa: E402
    AuctionRef,
    FetchError,
    NotXmlError,
    ScrapeError,
    SiteBlockedError,
    XmlRef,
    absolutize,
    describe_page,
    diagnose_last,
    discover_auctions,
    download_xml,
    fetch,
    find_xml_links,
    is_same_site,
    make_opener,
    safe_filename,
)

try:  # lxml służy wyłącznie do NIEZALEŻNEJ weryfikacji pobranych plików
    from lxml import etree as lxml_etree
except ImportError:  # pragma: no cover
    lxml_etree = None

MOCKSITE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "mocksite")

AUCTION_1087 = "flexit-auctions-26-02-2026-1087"
AUCTION_1103 = "flexit-auctions-18-06-2026-1103"


# ---------------------------------------------------------------------------
# Atrapa portalu
# ---------------------------------------------------------------------------


class _State:
    """Stan serwera atrapy: trasy, licznik żądań, dziennik."""

    def __init__(self, routes, root):
        self.routes = routes
        self.root = root
        self.log = []          # lista celów żądań (ścieżka + query)
        self.headers = []      # lista (cel, dict nagłówków)
        self.counters = {}

    def reset(self):
        self.log.clear()
        self.headers.clear()
        self.counters.clear()

    def count(self, target: str) -> int:
        """Ile razy poproszono o dany cel."""
        return sum(1 for item in self.log if item == target)


class _MockHandler(BaseHTTPRequestHandler):
    """Serwuje pliki z ``tests/mocksite`` wg ``routes.json``."""

    server_version = "FlexitMock/1.0"
    protocol_version = "HTTP/1.0"

    # -- infrastruktura -----------------------------------------------------

    def log_message(self, fmt, *args):  # cisza w wyniku testów
        pass

    def _send(self, status, body=b"", content_type="text/html; charset=utf-8",
              extra=None):
        try:
            self.send_response(status)
            if content_type:
                self.send_header("Content-Type", content_type)
            self.send_header("Content-Length", str(len(body)))
            for key, value in (extra or {}).items():
                self.send_header(key, value)
            self.end_headers()
            if body:
                self.wfile.write(body)
        except (BrokenPipeError, ConnectionResetError):
            # klient zdążył się rozłączyć (test timeoutu) — to nie jest błąd
            self.close_connection = True

    # -- właściwa obsługa ---------------------------------------------------

    def do_GET(self):  # noqa: N802 - nazwa narzucona przez http.server
        state = self.server.state
        target = self.path
        state.log.append(target)
        state.headers.append((target, {k.lower(): v for k, v in self.headers.items()}))
        spec = state.routes.get(target)
        if spec is None:
            spec = state.routes.get(target.split("?", 1)[0])
        if spec is None or not isinstance(spec, dict):
            return self._send(404, b"<html><body><h1>404 Not Found</h1></body></html>")

        seen = state.counters.get(target, 0)
        state.counters[target] = seen + 1
        fail_times = int(spec.get("fail_times", 0))
        if fail_times == -1 or seen < fail_times:
            extra = {}
            if spec.get("retry_after"):
                extra["Retry-After"] = str(spec["retry_after"])
            return self._send(int(spec.get("fail_status", 503)),
                              b"<html><body>Service Unavailable</body></html>",
                              extra=extra)

        delay = float(spec.get("delay", 0) or 0)
        if delay:
            time.sleep(delay)

        status = int(spec.get("status", 200))
        if status in (301, 302, 303, 307, 308):
            return self._send(status, b"", extra={"Location": spec["location"]})

        with open(os.path.join(state.root, spec["file"]), "rb") as handle:
            body = handle.read()
        content_type = spec.get("content_type")
        if not content_type:
            content_type = ("application/xml" if spec["file"].endswith(".xml")
                            else "text/html; charset=utf-8")
        return self._send(status, body, content_type, spec.get("headers"))

    def do_HEAD(self):  # noqa: N802
        self.do_GET()


class MockSite:
    """Serwer atrapy na losowym wolnym porcie, w wątku demona."""

    def __init__(self, root=MOCKSITE):
        with open(os.path.join(root, "routes.json"), "r", encoding="utf-8") as handle:
            routes = json.load(handle)
        self.server = ThreadingHTTPServer(("127.0.0.1", 0), _MockHandler)
        self.server.daemon_threads = True
        self.server.state = _State(routes, root)
        self.thread = threading.Thread(target=self.server.serve_forever, daemon=True)
        self.thread.start()

    @property
    def state(self):
        return self.server.state

    @property
    def base(self) -> str:
        host, port = self.server.server_address[:2]
        return "http://%s:%d/" % (host, port)

    def url(self, path: str) -> str:
        return self.base.rstrip("/") + path

    def stop(self):
        self.server.shutdown()
        self.server.server_close()
        self.thread.join(timeout=5)


SITE = None          # ustawiane w setUpModule


def setUpModule():   # noqa: N802 - nazwa narzucona przez unittest
    global SITE
    SITE = MockSite()


def tearDownModule():  # noqa: N802
    if SITE is not None:
        SITE.stop()


# ---------------------------------------------------------------------------
# Narzędzia testowe
# ---------------------------------------------------------------------------


class FakeClock:
    """Zegar i sen bez prawdziwego czekania — testy mają być SZYBKIE."""

    def __init__(self):
        self.now = 0.0
        self.slept = []

    def sleep(self, seconds):
        self.slept.append(round(float(seconds), 6))
        self.now += float(seconds)

    def monotonic(self):
        return self.now


@contextlib.contextmanager
def capture_stderr():
    """Przechwytuje ostrzeżenia modułu (wypisywane na ``stderr``)."""
    buffer = io.StringIO()
    with contextlib.redirect_stderr(buffer):
        yield buffer


class SiteTestCase(unittest.TestCase):
    """Wspólna baza: czysty dziennik żądań i szybki opener."""

    def setUp(self):
        SITE.state.reset()
        self.clock = FakeClock()
        self.tmp = tempfile.mkdtemp(prefix="flexit-scrape-")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def opener(self, **kw):
        params = dict(
            timeout=5.0,
            retries=4,
            delay=0.0,
            sleep=self.clock.sleep,
            clock=self.clock.monotonic,
            site=SITE.base,
        )
        params.update(kw)
        return make_opener(**params)

    def auction(self, ident=AUCTION_1087, path=None, title="Aukcja testowa"):
        return AuctionRef(url=SITE.url(path or ("/auction/" + ident)),
                          id=ident, title=title)


# ---------------------------------------------------------------------------
# 1. Adresy URL i nazwy plików (bez sieci)
# ---------------------------------------------------------------------------


class TestSafeFilename(unittest.TestCase):
    """Bezpieczne nazwy plików — w tym próby path traversal."""

    CASES = [
        ("../../etc/passwd", "passwd.xml"),
        ("../../../../etc/passwd", "passwd.xml"),
        ("/etc/passwd", "passwd.xml"),
        ("..\\..\\windows\\win.ini", "win.ini.xml"),
        ("....//....//etc/shadow", "shadow.xml"),
        ("%2e%2e%2f%2e%2e%2fetc%2fpasswd", "passwd.xml"),
        ("batch.xml", "batch.xml"),
        ("batch-1087-e0c8f.xml", "batch-1087-e0c8f.xml"),
        ("..", "plik.xml"),
        (".", "plik.xml"),
        ("", "plik.xml"),
        (None, "plik.xml"),
        ("   ", "plik.xml"),
        (".ukryty", "ukryty.xml"),
        # litery spoza ASCII są TRANSLITEROWANE, a nie kasowane — nazwa
        # dalej identyfikuje plik (dawniej zostawało samo "xml.xml")
        ("plik z polskimi znakami ąęć.xml", "plik_z_polskimi_znakami_aec.xml"),
        ("a/b/c/d.xml", "d.xml"),
        ("plik\x00.xml", "plik.xml"),
        ("plik\n\r\t.xml", "plik.xml"),
        ("con", "_con.xml"),
        ("COM1.xml", "_COM1.xml"),
        ("$(rm -rf ~).xml", "rm_-rf_.xml"),
        ("a" * 300 + ".xml", None),          # sprawdzane osobno (długość)
    ]

    def test_cases(self):
        for raw, expected in self.CASES:
            with self.subTest(raw=raw):
                got = safe_filename(raw)
                self.assertNotIn("/", got)
                self.assertNotIn("\\", got)
                self.assertNotIn("..", got)
                self.assertFalse(got.startswith("."))
                self.assertTrue(got.lower().endswith(".xml"))
                if expected is not None:
                    self.assertEqual(got, expected)

    def test_long_name_is_capped_and_stable(self):
        long_name = "b" * 400 + ".xml"
        first = safe_filename(long_name)
        self.assertLessEqual(len(first), scrape.MAX_FILENAME_LEN)
        self.assertEqual(first, safe_filename(long_name))          # deterministyczne
        other = safe_filename("c" * 400 + ".xml")
        self.assertNotEqual(first, other)                          # bez kolizji

    def test_ascii_only(self):
        got = safe_filename("Zażółć_gęślą.xml")
        self.assertTrue(all(ord(ch) < 128 for ch in got), got)

    def test_default_used_when_nothing_left(self):
        self.assertEqual(safe_filename("///", default="zapas.xml"), "zapas.xml")

    def test_ensure_ext_can_be_disabled(self):
        self.assertEqual(safe_filename("katalog", ensure_ext=""), "katalog")


class TestUrlHelpers(unittest.TestCase):
    """Sklejanie, normalizacja i granica witryny."""

    def test_absolutize_relative(self):
        base = "http://host/lot/apple-imac-mix-2a98f"
        self.assertEqual(absolutize(base, "../download/batch-2a98f"),
                         "http://host/download/batch-2a98f")
        self.assertEqual(absolutize(base, "/media/a.xml"), "http://host/media/a.xml")
        self.assertEqual(absolutize(base, "?page=2"),
                         "http://host/lot/apple-imac-mix-2a98f?page=2")
        self.assertEqual(absolutize("http://host/a/b/", "c.xml"), "http://host/a/b/c.xml")

    def test_absolutize_absolute_and_normalization(self):
        self.assertEqual(absolutize("http://host/", "HTTP://HOST:80/x/../y.xml"),
                         "http://host/y.xml")
        self.assertEqual(absolutize("http://host/", "https://Other.Example.COM/a"),
                         "https://other.example.com/a")

    def test_absolutize_rejects_dangerous_schemes(self):
        for href in ("javascript:alert(1)", "mailto:a@b.pl", "data:text/xml,<a/>",
                     "file:///etc/passwd", "tel:+48123", "#kotwica", "", "   ", None):
            with self.subTest(href=href):
                self.assertIsNone(absolutize("http://host/", href))

    def test_same_site(self):
        base = "https://flexitauctions.com/"
        self.assertTrue(is_same_site(base, "https://flexitauctions.com/auction/x"))
        self.assertTrue(is_same_site(base, "http://www.flexitauctions.com/x"))
        self.assertTrue(is_same_site(base, "https://media.flexitauctions.com/a.xml"))
        self.assertFalse(is_same_site(base, "https://evil.example.com/a.xml"))
        self.assertFalse(is_same_site(base, "https://flexitauctions.com.evil.pl/a"))
        self.assertFalse(is_same_site(base, "ftp://flexitauctions.com/a"))

    def test_same_site_ip_requires_same_port(self):
        self.assertTrue(is_same_site("http://127.0.0.1:8080/", "http://127.0.0.1:8080/x"))
        self.assertFalse(is_same_site("http://127.0.0.1:8080/", "http://127.0.0.1:9090/x"))
        self.assertFalse(is_same_site("http://127.0.0.1:8080/", "http://127.0.0.2:8080/x"))


# ---------------------------------------------------------------------------
# 2. fetch: ponawianie, backoff, timeouty, granica witryny
# ---------------------------------------------------------------------------


class TestFetch(SiteTestCase):
    """Warstwa transportowa."""

    def test_fetch_returns_bytes_and_content_type(self):
        opener = self.opener()
        data, ctype = fetch(opener, SITE.url("/media/batch-1087-e0c8f.xml"))
        self.assertTrue(data.startswith(b"<?xml"))
        self.assertEqual(ctype, "application/xml")
        self.assertEqual(self.clock.slept, [])          # nic nie czekało

    def test_content_type_without_parameters(self):
        opener = self.opener()
        _data, ctype = fetch(opener, SITE.url("/api/lot/cd048/batch?format=xml"))
        self.assertEqual(ctype, "text/xml")             # bez "; charset=utf-8"

    def test_headers_contain_user_agent_and_cookie(self):
        opener = self.opener(cookie="sid=abc123")
        fetch(opener, SITE.url("/media/batch-1087-e0c8f.xml"))
        _target, headers = SITE.state.headers[-1]
        self.assertIn("flexit2xlsx", headers.get("user-agent", ""))
        self.assertEqual(headers.get("cookie"), "sid=abc123")

    def test_404_is_not_retried(self):
        opener = self.opener()
        with self.assertRaises(FetchError) as ctx:
            fetch(opener, SITE.url("/nie-ma-takiej-strony"))
        self.assertEqual(ctx.exception.status, 404)
        self.assertEqual(len(SITE.state.log), 1)
        self.assertEqual(self.clock.slept, [])

    def test_503_is_retried_with_backoff(self):
        opener = self.opener()
        data, _ctype = fetch(opener, SITE.url("/flaky/batch.xml"))
        self.assertTrue(data.startswith(b"<?xml"))
        self.assertEqual(self.clock.slept, [2.0, 4.0])
        self.assertEqual(SITE.state.count("/flaky/batch.xml"), 3)

    def test_exponential_backoff_2_4_8_16(self):
        opener = self.opener(retries=4)
        with self.assertRaises(FetchError) as ctx:
            fetch(opener, SITE.url("/flaky-forever.xml"))
        self.assertEqual(self.clock.slept, [2.0, 4.0, 8.0, 16.0])
        self.assertEqual(SITE.state.count("/flaky-forever.xml"), 5)
        self.assertEqual(ctx.exception.status, 503)
        self.assertIn("503", str(ctx.exception))
        self.assertEqual(ctx.exception.url, SITE.url("/flaky-forever.xml"))

    def test_retries_zero_means_single_attempt(self):
        opener = self.opener(retries=0)
        with self.assertRaises(FetchError):
            fetch(opener, SITE.url("/flaky-forever.xml"))
        self.assertEqual(self.clock.slept, [])
        self.assertEqual(SITE.state.count("/flaky-forever.xml"), 1)

    def test_retry_after_header_is_honoured(self):
        opener = self.opener()
        fetch(opener, SITE.url("/retry-after.xml"))
        self.assertEqual(self.clock.slept, [1.0])

    def test_status_418_is_not_retried(self):
        opener = self.opener()
        with self.assertRaises(FetchError) as ctx:
            fetch(opener, SITE.url("/teapot.xml"))
        self.assertEqual(ctx.exception.status, 418)
        self.assertEqual(SITE.state.count("/teapot.xml"), 1)

    def test_timeout_is_reported_and_fast(self):
        opener = self.opener(timeout=0.05, retries=0)
        started = time.monotonic()
        with self.assertRaises(FetchError) as ctx:
            fetch(opener, SITE.url("/slow.xml"))
        self.assertLess(time.monotonic() - started, 3.0)
        self.assertIn("slow.xml", str(ctx.exception))

    def test_timeout_is_retried(self):
        opener = self.opener(timeout=0.05, retries=2)
        with self.assertRaises(FetchError):
            fetch(opener, SITE.url("/slow.xml"))
        self.assertEqual(self.clock.slept, [2.0, 4.0])   # backoff także dla timeoutu

    def test_offsite_url_is_blocked_without_request(self):
        opener = self.opener()
        with self.assertRaises(SiteBlockedError):
            fetch(opener, "https://evil.example.com/auction/hacked-999/info")
        self.assertEqual(SITE.state.log, [])

    def test_other_port_is_a_different_site(self):
        opener = self.opener()
        with self.assertRaises(SiteBlockedError):
            fetch(opener, "http://127.0.0.1:1/x.xml")

    def test_file_scheme_is_blocked(self):
        opener = self.opener()
        with self.assertRaises(SiteBlockedError):
            fetch(opener, "file:///etc/passwd")

    def test_max_bytes_guard(self):
        opener = self.opener(max_bytes=10)
        with self.assertRaises(ScrapeError):
            fetch(opener, SITE.url("/media/batch-1087-e0c8f.xml"))

    def test_polite_delay_between_requests(self):
        opener = self.opener(delay=0.5)
        url = SITE.url("/media/batch-1087-e0c8f.xml")
        fetch(opener, url)
        fetch(opener, url)
        fetch(opener, url)
        self.assertEqual(self.clock.slept, [0.5, 0.5])   # przed 2. i 3. żądaniem

    def test_redirect_inside_site_is_followed(self):
        opener = self.opener()
        data, ctype = fetch(opener, SITE.url("/redirect-in"))
        self.assertTrue(data.startswith(b"<?xml"))
        self.assertEqual(ctype, "application/xml")

    def test_redirect_outside_site_is_blocked(self):
        opener = self.opener()
        with self.assertRaises(SiteBlockedError):
            fetch(opener, SITE.url("/redirect-out"))

    def test_module_level_sleep_hook(self):
        """Backoff da się wstrzyknąć także przez atrybut modułu."""
        recorded = []
        original = scrape.SLEEP_FUNCTION
        scrape.SLEEP_FUNCTION = recorded.append
        try:
            opener = make_opener(retries=1, delay=0.0, timeout=5.0, site=SITE.base)
            with self.assertRaises(FetchError):
                fetch(opener, SITE.url("/flaky-forever.xml"))
        finally:
            scrape.SLEEP_FUNCTION = original
        self.assertEqual(recorded, [2.0])


# ---------------------------------------------------------------------------
# 3. Odkrywanie aukcji (paginacja)
# ---------------------------------------------------------------------------


class TestDiscoverAuctions(SiteTestCase):
    """Lista aukcji + paginacja."""

    def test_finds_all_auctions_across_pages(self):
        opener = self.opener()
        auctions = discover_auctions(opener, SITE.base)
        self.assertEqual(
            [a.id for a in auctions],
            [AUCTION_1087, AUCTION_1103, "flexit-auctions-02-01-2026-1023",
             "flexit-auctions-15-03-2025-1037", "flexit-auctions-30-11-2025-1011"],
        )
        self.assertTrue(all(a.title for a in auctions))
        self.assertEqual(auctions[0].title, "Flexit Auctions 26-02-2026 (Krakow)")

    def test_pagination_visits_each_page_once(self):
        opener = self.opener()
        discover_auctions(opener, SITE.base)
        listing = [t for t in SITE.state.log if t in ("/", "/?page=1", "/?page=2", "/?page=3")]
        self.assertEqual(sorted(listing), ["/", "/?page=2", "/?page=3"])

    def test_duplicate_links_are_merged_to_canonical_url(self):
        opener = self.opener()
        auctions = discover_auctions(opener, SITE.base)
        ids = [a.id for a in auctions]
        self.assertEqual(len(ids), len(set(ids)))
        ref = next(a for a in auctions if a.id == AUCTION_1087)
        # ta sama aukcja była pod /info i pod adresem kanonicznym
        self.assertTrue(ref.url.endswith("/auction/" + AUCTION_1087))

    def test_offsite_auctions_are_rejected(self):
        opener = self.opener()
        auctions = discover_auctions(opener, SITE.base)
        self.assertTrue(all("evil.example.com" not in a.url for a in auctions))
        self.assertTrue(all("hacked-999" not in a.id for a in auctions))

    def test_lot_links_are_not_auctions(self):
        opener = self.opener()
        auctions = discover_auctions(opener, SITE.base)
        self.assertTrue(all("/lot/" not in a.url for a in auctions))

    def test_max_pages_limits_crawl(self):
        opener = self.opener()
        auctions = discover_auctions(opener, SITE.base, max_pages=1)
        self.assertEqual([a.id for a in auctions], [AUCTION_1087, AUCTION_1103])

    def test_custom_regex(self):
        opener = self.opener()
        auctions = discover_auctions(
            opener, SITE.base, auction_re=r"/auction/(?P<id>[^/?#]+)/info$")
        self.assertEqual([a.id for a in auctions],
                         [AUCTION_1087, "flexit-auctions-02-01-2026-1023"])

    def test_regex_without_named_group(self):
        opener = self.opener()
        auctions = discover_auctions(opener, SITE.base, auction_re=r"/auction/([^/?#]+)$")
        self.assertIn(AUCTION_1103, [a.id for a in auctions])

    def test_no_results_returns_empty_list_not_exception(self):
        opener = self.opener()
        with capture_stderr() as err:
            auctions = discover_auctions(opener, SITE.url("/pusto"))
        self.assertEqual(auctions, [])
        self.assertIn("Nic nie znaleziono", err.getvalue())

    def test_first_page_error_raises(self):
        opener = self.opener()
        with self.assertRaises(FetchError):
            discover_auctions(opener, SITE.url("/nie-ma-listy"))

    def test_diagnostics_printed_when_enabled(self):
        opener = self.opener(diagnose=True)
        with capture_stderr() as err:
            discover_auctions(opener, SITE.url("/pusto"))
        report = err.getvalue()
        self.assertIn("diagnostyka strony", report)
        self.assertIn("Tytuł strony: Nothing here", report)
        self.assertIn("/about", report)


# ---------------------------------------------------------------------------
# 4. Szukanie linków XML
# ---------------------------------------------------------------------------


class TestFindXmlLinks(SiteTestCase):
    """Warianty wykrywania pliku XML pod aukcją."""

    def refs_1087(self, **kw):
        opener = kw.pop("opener", None) or self.opener()
        return find_xml_links(opener, self.auction(AUCTION_1087), **kw)

    def test_all_four_variants_found(self):
        refs = self.refs_1087()
        urls = sorted(r.url.replace(SITE.base.rstrip("/"), "") for r in refs)
        self.assertEqual(urls, [
            "/api/lot/cd048/batch?format=xml",          # wariant: ?format=xml
            "/dl/53439",                                # wariant: atrybut download
            "/download/batch-2a98f",                    # wariant: tylko Content-Type
            "/media/batch-1087-e0c8f.xml",              # wariant: href .xml
            "/media/download?file=../../../../etc/passwd&format=xml",
        ])
        self.assertTrue(all(isinstance(r, XmlRef) for r in refs))
        self.assertTrue(all(r.auction_id == AUCTION_1087 for r in refs))

    def test_filenames_are_unique_and_safe(self):
        refs = self.refs_1087()
        names = [r.filename for r in refs]
        self.assertEqual(len(names), len(set(names)))
        for name in names:
            with self.subTest(name=name):
                self.assertNotIn("/", name)
                self.assertNotIn("..", name)
                self.assertTrue(name.endswith(".xml"))
                self.assertTrue(name.startswith(AUCTION_1087 + "-"))

    def test_path_traversal_in_query_is_neutralised(self):
        refs = self.refs_1087()
        ref = next(r for r in refs if "passwd" in r.url)
        self.assertEqual(ref.filename, AUCTION_1087 + "-passwd.xml")

    def test_duplicate_links_collapse(self):
        refs = self.refs_1087()
        urls = [r.url for r in refs]
        self.assertEqual(len(urls), len(set(urls)))
        # ten sam lot linkowany dwoma adresami -> pobrany raz
        self.assertEqual(SITE.state.count("/lot/lenovo-8th-10th-gen-laptop-mix-e0c8f"), 1)
        self.assertEqual(
            SITE.state.count("/auction/%s/lenovo-8th-10th-gen-laptop-mix-e0c8f"
                             % AUCTION_1087), 0)

    def test_relative_link_is_resolved(self):
        refs = self.refs_1087()
        self.assertIn(SITE.url("/download/batch-2a98f"), [r.url for r in refs])

    def test_offsite_links_are_rejected(self):
        refs = self.refs_1087()
        joined = " ".join(r.url for r in refs)
        self.assertNotIn("evil.example.com", joined)
        self.assertNotIn("media.other-domain.net", joined)

    def test_html_answer_is_rejected_after_probe(self):
        refs = self.refs_1087()
        self.assertNotIn(SITE.url("/dl/other"), [r.url for r in refs])
        self.assertEqual(SITE.state.count("/dl/other"), 1)      # sondowany, odrzucony

    def test_boring_extensions_are_not_probed(self):
        self.refs_1087()
        self.assertEqual(SITE.state.count("/assets/lot.css"), 0)

    def test_probe_budget(self):
        with capture_stderr() as err:
            refs = self.refs_1087(probe_limit=0)
        self.assertIn("budżet sondowania", err.getvalue())
        urls = [r.url for r in refs]
        self.assertNotIn(SITE.url("/download/batch-2a98f"), urls)   # wymagał sondy
        self.assertIn(SITE.url("/media/batch-1087-e0c8f.xml"), urls)
        self.assertEqual(len(refs), 4)
        self.assertEqual(SITE.state.count("/dl/other"), 0)

    def test_content_type_probe_is_cached(self):
        opener = self.opener()
        self.refs_1087(opener=opener)
        self.refs_1087(opener=opener)
        self.assertEqual(SITE.state.count("/dl/other"), 1)

    def test_descend_auto_schodzi_do_lotow_mimo_xml_na_stronie_aukcji(self):
        """REGRESJA: zbiorczy XML aukcji nie może ukryć pakietów lotów.

        Strona aukcji 1103 ma własny plik "całej aukcji" ORAZ listę lotów.
        Dawniej ``--descend auto`` po znalezieniu tego jednego pliku nie schodził
        już do żadnego lotu — cały pakiet "15x ThinkPad T480 Mix" znikał
        z arkusza bez ostrzeżenia i z kodem wyjścia 0.
        """
        opener = self.opener()
        with capture_stderr() as err:
            refs = find_xml_links(opener, self.auction(AUCTION_1103))
        self.assertEqual(
            sorted(r.url for r in refs),
            sorted([SITE.url("/media/batch-1103-all.xml"),
                    SITE.url("/media/batch-1103-bcd12.xml")]),
        )
        self.assertEqual(SITE.state.count("/lot/thinkpad-t480-mix-bcd12"), 1)
        self.assertIn("schodzę też do lotów", err.getvalue())

    def test_descend_always_visits_lots(self):
        opener = self.opener()
        refs = find_xml_links(opener, self.auction(AUCTION_1103), descend="always")
        self.assertEqual(
            sorted(r.url for r in refs),
            sorted([SITE.url("/media/batch-1103-all.xml"),
                    SITE.url("/media/batch-1103-bcd12.xml")]),
        )

    def test_descend_never(self):
        opener = self.opener()
        with capture_stderr():
            refs = find_xml_links(opener, self.auction(AUCTION_1087), descend="never")
        self.assertEqual(refs, [])
        self.assertEqual(SITE.state.count("/lot/pusty-lot-6bfb2"), 0)

    def test_max_lots(self):
        opener = self.opener()
        with capture_stderr() as err:
            refs = find_xml_links(opener, self.auction(AUCTION_1087), max_lots=2)
        self.assertLessEqual(len(refs), 2)
        self.assertIn("max_lots", err.getvalue())

    def test_auction_without_xml_returns_empty_list(self):
        opener = self.opener()
        with capture_stderr() as err:
            refs = find_xml_links(opener, self.auction("flexit-auctions-02-01-2026-1023"))
        self.assertEqual(refs, [])
        self.assertIn("Nic nie znaleziono", err.getvalue())

    def test_unreachable_auction_page_returns_empty_list(self):
        opener = self.opener(retries=0)
        auction = AuctionRef(url=SITE.url("/auction/nie-ma-1"), id="nie-ma-1", title="X")
        with capture_stderr() as err:
            refs = find_xml_links(opener, auction)
        self.assertEqual(refs, [])
        self.assertIn("Nie udało się pobrać strony aukcji", err.getvalue())

    def test_auction_url_pointing_directly_at_xml(self):
        opener = self.opener()
        auction = AuctionRef(url=SITE.url("/media/batch-1087-e0c8f.xml"),
                             id=AUCTION_1087, title="Aukcja")
        refs = find_xml_links(opener, auction)
        self.assertEqual(len(refs), 1)
        self.assertEqual(refs[0].url, SITE.url("/media/batch-1087-e0c8f.xml"))

    def test_injected_opener_fetch_works_without_network(self):
        """``opener_fetch`` pozwala działać w całości offline."""
        pages = {
            "http://portal.test/auction/a-1": (
                b"<html><body><a href='/lot/x-aaaaa'>Lot</a></body></html>",
                "text/html",
            ),
            "http://portal.test/lot/x-aaaaa": (
                b"<html><body><a href='/dane/x.xml'>Download Batch Details</a>"
                b"</body></html>",
                "text/html",
            ),
        }
        calls = []

        def fake_fetch(opener, url):
            calls.append(url)
            if url in pages:
                return pages[url]
            raise FetchError("brak", url=url)

        opener = make_opener(delay=0.0, sleep=self.clock.sleep, site="http://portal.test/")
        refs = find_xml_links(
            opener,
            AuctionRef(url="http://portal.test/auction/a-1", id="a-1", title="A"),
            opener_fetch=fake_fetch,
        )
        self.assertEqual([r.url for r in refs], ["http://portal.test/dane/x.xml"])
        self.assertEqual(calls, ["http://portal.test/auction/a-1",
                                 "http://portal.test/lot/x-aaaaa"])
        self.assertEqual(SITE.state.log, [])          # nic nie poszło do sieci

    def test_spa_json_link_is_found(self):
        """Adres XML ukryty w osadzonym JSON (``__NEXT_DATA__``) też się liczy."""
        page = (
            b'<html><body><h1>Lot</h1>'
            b'<script id="__NEXT_DATA__" type="application/json">'
            b'{"lot":{"batch":"\\/media\\/ukryty.xml"}}</script></body></html>'
        )

        def fake_fetch(opener, url):
            return page, "text/html"

        opener = make_opener(delay=0.0, sleep=self.clock.sleep, site="http://portal.test/")
        refs = find_xml_links(
            opener,
            AuctionRef(url="http://portal.test/auction/a-1", id="a-1", title="A"),
            opener_fetch=fake_fetch,
        )
        self.assertEqual([r.url for r in refs], ["http://portal.test/media/ukryty.xml"])


# ---------------------------------------------------------------------------
# 5. Pobieranie plików
# ---------------------------------------------------------------------------


class TestDownloadXml(SiteTestCase):
    """Zapis plików na dysk."""

    def ref(self, path="/media/batch-1087-e0c8f.xml", filename="batch-e0c8f.xml"):
        return XmlRef(url=SITE.url(path), filename=filename,
                      auction_id=AUCTION_1087, auction_title="Aukcja")

    def test_downloads_and_returns_path(self):
        opener = self.opener()
        path = download_xml(opener, self.ref(), self.tmp)
        self.assertEqual(path, os.path.join(self.tmp, "batch-e0c8f.xml"))
        with open(path, "rb") as handle:
            data = handle.read()
        with open(os.path.join(MOCKSITE, "batch_e0c8f.xml"), "rb") as handle:
            self.assertEqual(data, handle.read())
        self.assertEqual(ET.fromstring(data).tag, "batch")

    def test_creates_missing_directory(self):
        opener = self.opener()
        target = os.path.join(self.tmp, "a", "b", "c")
        path = download_xml(opener, self.ref(), target)
        self.assertTrue(os.path.isfile(path))

    def test_existing_file_is_skipped_without_network(self):
        opener = self.opener()
        path = download_xml(opener, self.ref(), self.tmp)
        SITE.state.reset()
        again = download_xml(opener, self.ref(), self.tmp)
        self.assertEqual(path, again)
        self.assertEqual(SITE.state.log, [])

    def test_overwrite_forces_download(self):
        opener = self.opener()
        path = download_xml(opener, self.ref(), self.tmp)
        with open(path, "wb") as handle:
            handle.write(b"<stare/>")
        SITE.state.reset()
        download_xml(opener, self.ref(), self.tmp, overwrite=True)
        self.assertEqual(len(SITE.state.log), 1)
        with open(path, "rb") as handle:
            self.assertTrue(handle.read().startswith(b"<?xml"))

    def test_path_traversal_in_filename(self):
        opener = self.opener()
        outside = os.path.join(self.tmp, "poza")
        os.makedirs(outside)
        out_dir = os.path.join(self.tmp, "pobrane")
        for evil in ("../../../../etc/passwd", "../poza/wyciek.xml",
                     "..\\..\\windows\\win.ini", "/etc/passwd"):
            with self.subTest(evil=evil):
                path = download_xml(opener, self.ref(filename=evil), out_dir,
                                    overwrite=True)
                self.assertEqual(os.path.dirname(os.path.realpath(path)),
                                 os.path.realpath(out_dir))
        self.assertEqual(os.listdir(outside), [])
        for name in os.listdir(out_dir):
            self.assertNotIn("..", name)
        self.assertIn("passwd.xml", os.listdir(out_dir))

    def test_login_page_instead_of_xml_raises(self):
        opener = self.opener()
        ref = self.ref("/login.xml", "login.xml")
        with self.assertRaises(NotXmlError) as ctx:
            download_xml(opener, ref, self.tmp)
        self.assertIn("--cookie", str(ctx.exception))
        self.assertEqual(os.listdir(self.tmp), [])          # bez śmieci i bez .part

    def test_empty_answer_raises(self):
        opener = self.opener()
        with self.assertRaises(NotXmlError):
            download_xml(opener, self.ref("/empty.xml", "empty.xml"), self.tmp)
        self.assertEqual(os.listdir(self.tmp), [])

    def test_verification_can_be_disabled(self):
        opener = self.opener()
        path = download_xml(opener, self.ref("/login.xml", "login.xml"), self.tmp,
                            verify_xml=False)
        self.assertTrue(os.path.isfile(path))

    def test_no_part_files_left_behind(self):
        opener = self.opener()
        download_xml(opener, self.ref(), self.tmp)
        self.assertEqual([n for n in os.listdir(self.tmp) if n.endswith(".part")], [])

    def test_server_supplied_filename_is_ignored(self):
        """Content-Disposition z ``../../etc/passwd`` nie ma prawa nic zmienić."""
        opener = self.opener()
        refs = find_xml_links(opener, self.auction(AUCTION_1087))
        ref = next(r for r in refs if r.url.endswith("/dl/53439"))
        path = download_xml(opener, ref, self.tmp)
        self.assertEqual(os.path.dirname(path), self.tmp)
        self.assertNotIn("passwd", os.path.basename(path))
        self.assertEqual(os.path.basename(path), AUCTION_1087 + "-batch-53439.xml")

    def test_download_error_propagates(self):
        opener = self.opener(retries=0)
        with self.assertRaises(FetchError):
            download_xml(opener, self.ref("/nie-ma.xml", "brak.xml"), self.tmp)


# ---------------------------------------------------------------------------
# 6. Diagnostyka
# ---------------------------------------------------------------------------


class TestDescribePage(SiteTestCase):
    """Funkcja diagnostyczna — co właściwie było na stronie."""

    def test_report_lists_links_and_offsite(self):
        opener = self.opener()
        data, ctype = fetch(opener, SITE.base)
        report = describe_page(data, SITE.base, content_type=ctype)
        self.assertIn("Tytuł strony: Online Auctions", report)
        self.assertIn("/auction/" + AUCTION_1087, report)
        self.assertIn("Odrzucone (obca domena)", report)
        self.assertIn("evil.example.com", report)
        self.assertIn("Linków razem:", report)

    def test_report_can_be_printed_to_stream(self):
        buffer = io.StringIO()
        result = describe_page(b"<html><a href='/a.xml'>x</a></html>",
                               "http://host/", stream=buffer)
        self.assertEqual(buffer.getvalue().strip(), result.strip())
        self.assertIn("Adresy wyglądające na XML", result)

    def test_report_for_page_without_links(self):
        report = describe_page(b"<html><body>nic</body></html>", "http://host/")
        self.assertIn("SPA", report)

    def test_report_survives_garbage(self):
        report = describe_page(b"\x00\xff\xfe\x01 to nie jest HTML", "http://host/")
        self.assertIn("diagnostyka strony", report)

    def test_report_counts_pattern_matches(self):
        opener = self.opener()
        data, ctype = fetch(opener, SITE.base)
        report = describe_page(data, SITE.base, content_type=ctype,
                               pattern=scrape.DEFAULT_AUCTION_RE)
        self.assertIn("Pasujących do wzorca", report)

    def test_diagnose_last_uses_last_page(self):
        opener = self.opener()
        with capture_stderr():
            discover_auctions(opener, SITE.url("/pusto"))
        report = diagnose_last(opener)
        self.assertIn("/pusto", report)
        self.assertIn("Nothing here", report)

    def test_diagnose_last_without_page(self):
        opener = self.opener()
        self.assertIn("Brak zapamiętanej strony", diagnose_last(opener))


# ---------------------------------------------------------------------------
# 7. Przebieg całościowy
# ---------------------------------------------------------------------------


class TestEndToEnd(SiteTestCase):
    """Cały scenariusz: lista -> aukcje -> loty -> pliki XML na dysku."""

    def test_full_run(self):
        opener = self.opener(delay=0.5)
        auctions = discover_auctions(opener, SITE.base)
        self.assertEqual(len(auctions), 5)
        paths = []
        with capture_stderr():
            for auction in auctions:
                for ref in find_xml_links(opener, auction):
                    paths.append(download_xml(opener, ref, self.tmp))
        # 7, a nie 6: aukcja 1103 ma zbiorczy XML *oraz* pakiet lotu bcd12,
        # do którego "auto" teraz schodzi
        self.assertEqual(len(paths), 7)
        self.assertEqual(len(set(paths)), 7)
        self.assertEqual(sorted(os.listdir(self.tmp)),
                         sorted(os.path.basename(p) for p in paths))
        for path in paths:
            with self.subTest(path=path):
                root = ET.parse(path).getroot()
                self.assertEqual(root.tag, "batch")
                if lxml_etree is not None:      # niezależna weryfikacja
                    self.assertEqual(lxml_etree.parse(path).getroot().tag, "batch")
        # uprzejme opóźnienie zadziałało przed każdym żądaniem poza pierwszym
        self.assertEqual(self.clock.slept, [0.5] * (opener.stats["requests"] - 1))

    def test_second_run_skips_downloaded_files(self):
        opener = self.opener()
        auction = self.auction(AUCTION_1087)
        with capture_stderr():
            refs = find_xml_links(opener, auction)
            for ref in refs:
                download_xml(opener, ref, self.tmp)
        before = sorted(os.listdir(self.tmp))
        SITE.state.reset()
        for ref in refs:
            download_xml(opener, ref, self.tmp)
        self.assertEqual(SITE.state.log, [])
        self.assertEqual(sorted(os.listdir(self.tmp)), before)


# ---------------------------------------------------------------------------
# 7b. Parsowanie stron (kodowania, <base>)
# ---------------------------------------------------------------------------


class TestParsePage(unittest.TestCase):
    """Dekodowanie HTML-a i sklejanie adresów."""

    def test_meta_charset_windows_1250(self):
        html = (
            "<html><head><meta charset='windows-1250'><title>Zażółć gęślą</title>"
            "</head><body><a href='/a.xml'>Pobierz</a></body></html>"
        ).encode("windows-1250")
        page = scrape.parse_page(html, "http://host/")
        self.assertEqual(page.title, "Zażółć gęślą")
        self.assertEqual(page.encoding, "windows-1250")

    def test_charset_from_content_type_header(self):
        html = "<html><title>Łódź</title></html>".encode("iso-8859-2")
        page = scrape.parse_page(html, "http://host/", "text/html; charset=ISO-8859-2")
        self.assertEqual(page.title, "Łódź")

    def test_bom_is_stripped(self):
        html = "\ufeff<html><title>Aukcje</title></html>".encode("utf-8")
        page = scrape.parse_page(html, "http://host/")
        self.assertEqual(page.title, "Aukcje")

    def test_broken_encoding_does_not_crash(self):
        page = scrape.parse_page(b"<html><title>\xff\xfe</title><a href='/x.xml'>x</a>",
                                 "http://host/")
        self.assertIn("http://host/x.xml", [link.url for link in page.links])

    def test_base_href_is_respected(self):
        html = b"<html><head><base href='/media/'></head><body><a href='b.xml'>x</a></body></html>"
        page = scrape.parse_page(html, "http://host/lot/abc")
        self.assertIn("http://host/media/b.xml", [link.url for link in page.links])

    def test_malformed_html_still_yields_links(self):
        html = (b"<html><body><a href='/a.xml'>pierwszy<a href='/b.xml'>drugi"
                b"<div><p>bez zamkniecia</body>")
        page = scrape.parse_page(html, "http://host/")
        self.assertEqual([link.url for link in page.links],
                         ["http://host/a.xml", "http://host/b.xml"])


# ---------------------------------------------------------------------------
# 8. Kontrakt modułu
# ---------------------------------------------------------------------------


class TestModuleContract(unittest.TestCase):
    """Sygnatury z INTERFACES.md i zakaz zależności zewnętrznych."""

    STDLIB = {
        "gzip", "hashlib", "http", "os", "posixpath", "re", "socket", "sys",
        "time", "unicodedata", "urllib", "zlib", "dataclasses", "html",
        "typing", "__future__",
    }

    def test_only_stdlib_imports(self):
        path = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
                            "flexit2xlsx", "scrape.py")
        with open(path, "r", encoding="utf-8") as handle:
            tree = ast.parse(handle.read(), filename=path)
        modules = set()
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                modules.update(alias.name.split(".")[0] for alias in node.names)
            elif isinstance(node, ast.ImportFrom) and node.level == 0 and node.module:
                modules.add(node.module.split(".")[0])
        self.assertTrue(modules <= self.STDLIB, "obce zależności: %s" % (modules - self.STDLIB))

    def test_public_signatures(self):
        import dataclasses
        import inspect

        self.assertEqual([f.name for f in dataclasses.fields(AuctionRef)],
                         ["url", "id", "title"])
        self.assertEqual([f.name for f in dataclasses.fields(XmlRef)],
                         ["url", "filename", "auction_id", "auction_title"])

        sig = inspect.signature(scrape.make_opener)
        for name, default in (("user_agent", None), ("cookie", None),
                              ("timeout", 30.0), ("retries", 4)):
            self.assertIn(name, sig.parameters)
            self.assertEqual(sig.parameters[name].kind,
                             inspect.Parameter.KEYWORD_ONLY)
            if default is not None:
                self.assertEqual(sig.parameters[name].default, default)

        self.assertEqual(list(inspect.signature(scrape.fetch).parameters),
                         ["opener", "url"])
        discover = inspect.signature(scrape.discover_auctions).parameters
        self.assertEqual(list(discover)[:2], ["opener", "base_url"])
        self.assertIsNone(discover["auction_re"].default)
        self.assertEqual(discover["max_pages"].default, 50)
        find = inspect.signature(scrape.find_xml_links).parameters
        self.assertEqual(list(find)[:2], ["opener", "auction"])
        self.assertIsNone(find["opener_fetch"].default)
        download = inspect.signature(scrape.download_xml).parameters
        self.assertEqual(list(download)[:3], ["opener", "ref", "out_dir"])
        self.assertIs(download["overwrite"].default, False)


# ---------------------------------------------------------------------------
# 8. REGRESJE po audycie adwersaryjnym
#
# Każdy test odpowiada jednemu ustaleniu.  Serwer HTTP z ``http.server`` nie
# pozwala udawać obciętej odpowiedzi ani gzipa wbrew negocjacji, więc te testy
# używają WŁASNEGO serwera na surowym gnieździe TCP.
# ---------------------------------------------------------------------------


class RawSocketServer:
    """Serwer TCP odpowiadający dokładnie tymi bajtami, które poda handler.

    Potrzebny, bo ``http.server`` sam pilnuje poprawności odpowiedzi — nie da
    się nim udawać serwera, który kłamie w ``Content-Length``, pakuje wbrew
    ``Accept-Encoding`` albo sączy dane po bajcie.
    """

    def __init__(self, handler):
        self.handler = handler
        self.sock = socket.socket()
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(16)
        self.port = self.sock.getsockname()[1]
        self.hits = []                      # ścieżki kolejnych żądań
        self.cookies = []                   # (ścieżka, nagłówek Cookie)
        self._lock = threading.Lock()
        self._stop = False
        threading.Thread(target=self._accept_loop, daemon=True).start()

    def _accept_loop(self):
        while not self._stop:
            try:
                conn, _ = self.sock.accept()
            except OSError:
                return
            threading.Thread(target=self._serve, args=(conn,), daemon=True).start()

    def _serve(self, conn):
        try:
            conn.settimeout(15)
            raw = b""
            while b"\r\n\r\n" not in raw:
                chunk = conn.recv(4096)
                if not chunk:
                    return
                raw += chunk
            head = raw.split(b"\r\n", 1)[0].decode("latin-1").split(" ")
            path = head[1] if len(head) > 1 else "/"
            cookie = ""
            for line in raw.decode("latin-1", "replace").split("\r\n"):
                if line.lower().startswith("cookie:"):
                    cookie = line.split(":", 1)[1].strip()
            with self._lock:
                self.hits.append(path)
                self.cookies.append((path, cookie))
            self.handler(self, conn, path, raw)
        except Exception:                   # serwer testowy nigdy nie psuje testu
            pass
        finally:
            try:
                conn.close()
            except OSError:
                pass

    @property
    def base(self):
        return "http://127.0.0.1:%d/" % self.port

    def url(self, path):
        return "http://127.0.0.1:%d%s" % (self.port, path)

    def count(self, path):
        with self._lock:
            return self.hits.count(path)

    def close(self):
        self._stop = True
        try:
            self.sock.close()
        except OSError:
            pass


#: Wzorcowy plik "zawartość pakietu" używany przez testy surowego gniazda.
RAW_XML = (b"<?xml version='1.0' encoding='UTF-8'?>\n"
           b"<batch><item><model>ThinkPad T480</model></item>"
           b"<item><model>Latitude 5400</model></item></batch>\n")


def raw_response(conn, status="200 OK", headers=(), body=b""):
    text = "HTTP/1.1 %s\r\n" % status
    for key, value in headers:
        text += "%s: %s\r\n" % (key, value)
    conn.sendall(text.encode("latin-1") + b"\r\n" + body)


def raw_body(conn, body, ctype="application/xml", extra=()):
    if isinstance(body, str):
        body = body.encode("utf-8")
    raw_response(conn, "200 OK",
                 list(extra) + [("Content-Type", ctype),
                                ("Content-Length", str(len(body)))], body)


class RawServerTestCase(unittest.TestCase):
    """Baza dla testów na surowym gnieździe: serwer, opener, katalog tymczasowy."""

    def serve(self, handler):
        server = RawSocketServer(handler)
        self.addCleanup(server.close)
        return server

    def opener(self, server, **kw):
        self.clock = FakeClock()
        params = dict(delay=0.0, site=server.base, timeout=5.0, retries=0,
                      sleep=self.clock.sleep, clock=self.clock.monotonic)
        params.update(kw)
        return make_opener(**params)

    def tmpdir(self):
        path = tempfile.mkdtemp(prefix="flexit-reg-")
        self.addCleanup(shutil.rmtree, path, True)
        return path

    @staticmethod
    def ref(server, path="batch.xml", name="batch.xml"):
        return XmlRef(url=server.base + path, filename=name,
                      auction_id="auk1", auction_title="Aukcja")

    @staticmethod
    def auction_ref(server, path="auction/auk1", ident="auk1"):
        return AuctionRef(url=server.base + path, id=ident, title="Aukcja")


class TestIntegralnoscOdpowiedzi(RawServerTestCase):
    """REGRESJA: obcięta odpowiedź nie może udawać kompletnego pliku."""

    def test_niedobor_wzgledem_content_length_jest_bledem(self):
        """Serwer deklaruje 5085 B, wysyła 45 B i się rozłącza.

        ``http.client`` świadomie NIE zgłasza tu ``IncompleteRead``, więc bez
        własnego sprawdzenia ogryzek trafiał na dysk jako komplet i przy
        kolejnym uruchomieniu był POMIJANY jako "już pobrany" — uszkodzenie
        było trwałe, a CLI zgłaszało je jako błąd parsowania XML.
        """
        def handler(srv, conn, path, req):
            conn.sendall(("HTTP/1.1 200 OK\r\nContent-Type: application/xml\r\n"
                          "Content-Length: %d\r\n\r\n" % (len(RAW_XML) + 5000)
                          ).encode("latin-1") + RAW_XML[:45])
            conn.close()

        server = self.serve(handler)
        opener = self.opener(server)
        out = self.tmpdir()
        with self.assertRaises(FetchError) as ctx:
            download_xml(opener, self.ref(server), out)
        self.assertIn("IncompleteRead", str(ctx.exception))
        self.assertEqual(os.listdir(out), [], "nic uszkodzonego nie zostaje na dysku")

    def test_niedobor_jest_ponawiany_z_backoffem(self):
        """Niedobór to błąd sieci — ma prawo do ponowień jak każdy inny."""
        def handler(srv, conn, path, req):
            conn.sendall(b"HTTP/1.1 200 OK\r\nContent-Type: application/xml\r\n"
                         b"Content-Length: 40\r\n\r\n<?xml versi")
            conn.close()

        server = self.serve(handler)
        opener = self.opener(server, retries=2)
        with self.assertRaises(FetchError):
            fetch(opener, server.base + "b.xml")
        self.assertEqual(len(server.hits), 3)
        self.assertEqual(self.clock.slept, [2.0, 4.0])

    def test_nadmiar_wzgledem_content_length_nie_przeszkadza(self):
        """Serwer, który wysyła WIĘCEJ niż obiecał, nie jest błędem transmisji."""
        def handler(srv, conn, path, req):
            raw_response(conn, "200 OK",
                         [("Content-Type", "application/xml"),
                          ("Content-Length", "10")], RAW_XML)
            conn.close()

        server = self.serve(handler)
        data, _ = fetch(self.opener(server), server.base + "b.xml")
        self.assertTrue(data.startswith(b"<?xml"))

    def test_odpowiedz_bez_content_length_dziala_jak_dawniej(self):
        """HTTP/1.0 bez ``Content-Length`` (koniec = rozłączenie)."""
        def handler(srv, conn, path, req):
            conn.sendall(b"HTTP/1.0 200 OK\r\nContent-Type: text/xml\r\n\r\n" + RAW_XML)
            conn.close()

        server = self.serve(handler)
        data, _ = fetch(self.opener(server), server.base + "b.xml")
        self.assertEqual(data, RAW_XML)


class TestKodowanieTransportowe(RawServerTestCase):
    """REGRESJA: ``Content-Encoding`` musi być honorowany."""

    def test_gzip_jest_rozpakowywany(self):
        """Cloudflare i nginx z ``gzip_static`` pakują mimo ``identity``.

        Wcześniej ``download_xml`` widział bajty ``\\x1f\\x8b``, ogłaszał
        "to nie XML" i radził ``--cookie`` — diagnoza całkowicie myląca.
        """
        body = gzip.compress(RAW_XML)

        def handler(srv, conn, path, req):
            raw_response(conn, "200 OK",
                         [("Content-Type", "application/xml"),
                          ("Content-Encoding", "gzip"),
                          ("Content-Length", str(len(body)))], body)

        server = self.serve(handler)
        opener = self.opener(server)
        data, ctype = fetch(opener, server.base + "batch.xml")
        self.assertEqual(data, RAW_XML)
        self.assertEqual(ctype, "application/xml")
        path = download_xml(opener, self.ref(server), self.tmpdir())
        with open(path, "rb") as handle:
            self.assertEqual(handle.read(), RAW_XML)

    def test_gzip_na_liscie_aukcji_nie_gubi_linkow(self):
        page = ('<html><body><a href="/auction/flexit-auctions-18-06-2026-1103">'
                'Aukcja</a></body></html>').encode("utf-8")
        body = gzip.compress(page)

        def handler(srv, conn, path, req):
            raw_response(conn, "200 OK",
                         [("Content-Type", "text/html; charset=utf-8"),
                          ("Content-Encoding", "gzip"),
                          ("Content-Length", str(len(body)))], body)

        server = self.serve(handler)
        with capture_stderr():
            found = discover_auctions(self.opener(server), server.base)
        self.assertEqual([a.id for a in found], ["flexit-auctions-18-06-2026-1103"])

    def test_deflate_jest_rozpakowywany(self):
        body = zlib.compress(RAW_XML)

        def handler(srv, conn, path, req):
            raw_response(conn, "200 OK",
                         [("Content-Type", "application/xml"),
                          ("Content-Encoding", "deflate"),
                          ("Content-Length", str(len(body)))], body)

        server = self.serve(handler)
        data, _ = fetch(self.opener(server), server.base + "b.xml")
        self.assertEqual(data, RAW_XML)

    def test_uszkodzony_gzip_konczy_sie_fetcherror(self):
        """Nie da się rozpakować -> błąd transmisji, a nie śmieci na dysku."""
        def handler(srv, conn, path, req):
            body = b"\x1f\x8b" + b"popsute"
            raw_response(conn, "200 OK",
                         [("Content-Type", "application/xml"),
                          ("Content-Encoding", "gzip"),
                          ("Content-Length", str(len(body)))], body)

        server = self.serve(handler)
        with self.assertRaises(FetchError):
            fetch(self.opener(server), server.base + "b.xml")

    def test_bomba_zip_jest_zatrzymana_po_rozpakowaniu(self):
        """Limit ``max_bytes`` obowiązuje także PO dekompresji."""
        body = gzip.compress(b"<a>" + b"x" * 5_000_000 + b"</a>")

        def handler(srv, conn, path, req):
            raw_response(conn, "200 OK",
                         [("Content-Type", "application/xml"),
                          ("Content-Encoding", "gzip"),
                          ("Content-Length", str(len(body)))], body)

        server = self.serve(handler)
        opener = self.opener(server, max_bytes=100_000)
        with self.assertRaises(ScrapeError) as ctx:
            fetch(opener, server.base + "b.xml")
        self.assertIn("limit", str(ctx.exception))


class TestRozpoznawaniaXml(RawServerTestCase):
    """REGRESJA: XML w UTF-16 i XML podany jako ``text/html``."""

    UTF16 = "<?xml version='1.0' encoding='UTF-16'?><batch><i>ąćż</i></batch>"

    def test_looks_like_xml_rozumie_bomy(self):
        for encoding in ("utf-16", "utf-16-le", "utf-16-be", "utf-32", "utf-8-sig"):
            with self.subTest(encoding=encoding):
                self.assertTrue(scrape._looks_like_xml(self.UTF16.encode(encoding)))
        self.assertFalse(scrape._looks_like_xml("<html><body>x".encode("utf-16")))
        self.assertFalse(scrape._looks_like_xml(b"<!DOCTYPE html><html>"))

    def test_xml_w_utf16_trafia_na_dysk(self):
        """Eksporty z narzędzi windowsowych bywają w UTF-16LE.

        Dawniej ``_looks_like_xml`` widziało pierwszy bajt ``\\xff`` i mówiło
        "to nie XML" — cała aukcja przepadała po cichu.
        """
        body = self.UTF16.encode("utf-16")

        def handler(srv, conn, path, req):
            raw_body(conn, body)

        server = self.serve(handler)
        path = download_xml(self.opener(server), self.ref(server), self.tmpdir())
        self.assertEqual(ET.parse(path).getroot().tag, "batch")

    def test_xml_podany_jako_text_html_bez_rozszerzenia(self):
        """Endpoint eksportu ze źle ustawionym typem MIME (bardzo częste)."""
        page = ('<html><body><a href="/api/lot/11588/batch">'
                'Download Batch Details</a></body></html>')

        def handler(srv, conn, path, req):
            if path.startswith("/api/"):
                raw_body(conn, RAW_XML, ctype="text/html; charset=utf-8")
            else:
                raw_body(conn, page, ctype="text/html; charset=utf-8")

        server = self.serve(handler)
        opener = self.opener(server)
        refs = find_xml_links(opener, self.auction_ref(server))
        self.assertEqual(len(refs), 1)
        path = download_xml(opener, refs[0], self.tmpdir())
        with open(path, "rb") as handle:
            self.assertEqual(handle.read(), RAW_XML)

    def test_prawdziwa_strona_html_dalej_jest_odrzucana(self):
        """Sniffing treści nie może wpuścić zwykłej strony HTML."""
        page = ('<html><body><a href="/api/batch">Download Batch Details</a>'
                '</body></html>')

        def handler(srv, conn, path, req):
            raw_body(conn, page, ctype="text/html")

        server = self.serve(handler)
        with capture_stderr():
            refs = find_xml_links(self.opener(server), self.auction_ref(server))
        self.assertEqual(refs, [])


class TestSondowaniaTypu(RawServerTestCase):
    """REGRESJA: sonda ``Content-Type`` nie może kosztować drugiego pobrania."""

    def test_niepewny_link_pobierany_tylko_raz(self):
        page = ('<html><body><a href="/api/batch" download>'
                'Download Batch Details</a></body></html>')

        def handler(srv, conn, path, req):
            if path.startswith("/api"):
                raw_body(conn, RAW_XML, ctype="application/octet-stream")
            else:
                raw_body(conn, page, ctype="text/html")

        server = self.serve(handler)
        opener = self.opener(server)
        refs = find_xml_links(opener, self.auction_ref(server, "auction/a1", "a1"))
        self.assertEqual(len(refs), 1)
        path = download_xml(opener, refs[0], self.tmpdir())
        with open(path, "rb") as handle:
            self.assertEqual(handle.read(), RAW_XML)
        self.assertEqual(server.count("/api/batch"), 1,
                         "sonda i pobranie to dawniej były dwa transfery")


class TestLacznegoTerminu(RawServerTestCase):
    """REGRESJA: ``--timeout`` musi ograniczać CAŁE pobranie, nie jeden odczyt."""

    def test_saczenie_po_bajcie_jest_przerywane(self):
        """Serwer sączy po bajcie: żaden pojedynczy odczyt nie łamie timeoutu.

        Bez terminu na całą odpowiedź wrogi lub przeciążony portal trzymał
        narzędzie dowolnie długo (pomiar w audycie: 13-krotność ``--timeout``).
        """
        def handler(srv, conn, path, req):
            conn.sendall(b"HTTP/1.1 200 OK\r\nContent-Type: application/xml\r\n"
                         b"Content-Length: 400\r\n\r\n")
            for _ in range(400):
                try:
                    conn.sendall(b"x")
                    time.sleep(0.05)
                except OSError:
                    return

        server = self.serve(handler)
        # realny zegar: przedmiotem pomiaru jest faktyczny upływ czasu
        opener = make_opener(delay=0.0, site=server.base, timeout=0.2, retries=0,
                             sleep=lambda seconds: None)
        started = time.monotonic()
        with self.assertRaises(FetchError):
            fetch(opener, server.base + "drip.xml")
        elapsed = time.monotonic() - started
        limit = 0.2 * scrape.RESPONSE_TIMEOUT_FACTOR
        self.assertLess(elapsed, limit + 2.0,
                        "pobranie trwało %.1f s przy limicie %.1f s" % (elapsed, limit))


class TestSklejaniaCiasteczek(RawServerTestCase):
    """REGRESJA: ``--cookie`` nie może unieważniać ``http.cookiejar``."""

    @staticmethod
    def _handler(srv, conn, path, req):
        if path.endswith(".xml"):
            raw_body(conn, RAW_XML)
        else:
            raw_body(conn, '<html><body><a href="/b.xml">Batch</a></body></html>',
                     ctype="text/html",
                     extra=[("Set-Cookie", "csrf=ABC123; Path=/")])

    def test_statyczne_cookie_i_jar_ida_razem(self):
        """Portal dosyła token CSRF — musi wrócić RAZEM z sesją z ``--cookie``."""
        server = self.serve(self._handler)
        opener = self.opener(server, cookie="SESSION=tajne")
        refs = find_xml_links(opener, self.auction_ref(server))
        download_xml(opener, refs[0], self.tmpdir())
        sent = dict(server.cookies)["/b.xml"]
        self.assertIn("csrf=ABC123", sent)
        self.assertIn("SESSION=tajne", sent)

    def test_bez_cookie_jar_dziala_jak_dawniej(self):
        server = self.serve(self._handler)
        opener = self.opener(server)
        refs = find_xml_links(opener, self.auction_ref(server))
        download_xml(opener, refs[0], self.tmpdir())
        self.assertEqual(dict(server.cookies)["/b.xml"], "csrf=ABC123")

    def test_ciasteczko_serwera_wygrywa_przy_tej_samej_nazwie(self):
        """Świeższa wartość (od portalu) nie może być zdublowana starą."""
        merged = scrape._merge_cookie_header("csrf=NOWY", "csrf=STARY; extra=1")
        self.assertEqual(merged, "csrf=NOWY; extra=1")


class TestNaglowkowNiepoprawnych(unittest.TestCase):
    """REGRESJA: wklejone ciasteczko z Enterem nie może wywalić CLI."""

    def test_koncowy_enter_jest_obcinany(self):
        opener = make_opener(cookie="SESSIONID=abc\n", delay=0.0)
        self.assertEqual(opener.cookie, "SESSIONID=abc")
        opener = make_opener(user_agent="  MojUA/1.0  ", delay=0.0)
        self.assertEqual(opener.user_agent, "MojUA/1.0")

    def test_znak_konca_linii_w_srodku_to_czytelny_blad(self):
        for kwargs in ({"cookie": "A=1\nX-Zle: 1"},
                       {"user_agent": "UA\r\nX-Wstrzykniete: 1"}):
            with self.subTest(kwargs=kwargs):
                with self.assertRaises(ScrapeError) as ctx:
                    make_opener(delay=0.0, **kwargs)
                self.assertIn("końca linii", str(ctx.exception))

    def test_puste_ciasteczko_znaczy_brak_ciasteczka(self):
        self.assertIsNone(make_opener(cookie="   ", delay=0.0).cookie)


class TestGranicyZrodla(unittest.TestCase):
    """REGRESJA: ciasteczko sesyjne tylko do źródła portalu i jego poddomen."""

    def test_obca_domena_w_tym_samym_sufiksie_publicznym(self):
        cases = [
            ("https://sklep.example.co.uk/", "https://evil.co.uk/kradnij"),
            ("https://aukcje.com.pl/", "https://zlodziej.com.pl/kradnij"),
            ("https://przetargi.gov.pl/", "https://falszywy.gov.pl/kradnij"),
            ("https://klient.github.io/", "https://atakujacy.github.io/kradnij"),
        ]
        for base, url in cases:
            with self.subTest(base=base):
                self.assertFalse(is_same_site(base, url))

    def test_inny_port_to_inne_zrodlo_takze_dla_nazw(self):
        self.assertFalse(is_same_site("http://intranet:8080/", "http://intranet:9200/"))
        self.assertTrue(is_same_site("http://intranet:8080/", "http://intranet:8080/a"))

    def test_poddomeny_portalu_dalej_dzialaja(self):
        base = "https://flexitauctions.com/"
        self.assertTrue(is_same_site(base, "https://media.flexitauctions.com/a.xml"))
        self.assertTrue(is_same_site(base, "http://www.flexitauctions.com/x"))
        self.assertTrue(is_same_site("https://www.flexitauctions.com/",
                                     "https://flexitauctions.com/x"))


class TestWyciekuCiasteczka(RawServerTestCase):
    """REGRESJA: przekierowanie na inny port nie dostaje ciasteczka."""

    def test_przekierowanie_na_inny_port_jest_blokowane(self):
        widziane = []

        def ofiara(srv, conn, path, req):
            widziane.append(req.decode("latin-1", "replace"))
            raw_body(conn, "<html>ok</html>", ctype="text/html")

        victim = self.serve(ofiara)

        def portal(srv, conn, path, req):
            raw_response(conn, "302 Found",
                         [("Location", victim.url("/kradnij")), ("Content-Length", "0")])

        server = self.serve(portal)
        opener = self.opener(server, cookie="SESSIONID=TAJNY-TOKEN")
        with self.assertRaises(ScrapeError):
            fetch(opener, server.base)
        self.assertEqual(widziane, [], "serwer-ofiara nie zobaczył żadnego żądania")


class TestSekwencjiSterujacych(RawServerTestCase):
    """REGRESJA: treść z portalu nie steruje terminalem użytkownika."""

    HTML = ("<html><head><title>Aukcje\x1b]0;PRZEJETY-TYTUL\x07</title></head><body>"
            "<a href='/auction/flexit-\x1b[2J\x1b[31mKRWAWY-1103/'>"
            "Lot \x1b[5mMIGA\x1b[0m</a></body></html>").encode("utf-8")

    def test_id_tytul_i_raport_bez_znakow_sterujacych(self):
        def handler(srv, conn, path, req):
            raw_body(conn, self.HTML, ctype="text/html; charset=utf-8")

        server = self.serve(handler)
        auctions = discover_auctions(self.opener(server), server.base)
        self.assertEqual(len(auctions), 1)
        for text in (auctions[0].id, auctions[0].title):
            self.assertNotIn("\x1b", text)
            self.assertNotIn("\x07", text)
        report = describe_page(self.HTML, "http://x/", content_type="text/html")
        self.assertNotIn("\x1b", report)
        self.assertNotIn("\x07", report)

    def test_ostrzezenia_modulu_sa_czyszczone(self):
        with capture_stderr() as err:
            scrape._warn("adres \x1b[2J/kradnij")
        self.assertNotIn("\x1b", err.getvalue())


class TestSkanowaniaSpa(RawServerTestCase):
    """REGRESJA: względny adres ``.xml`` w osadzonym JSON-ie."""

    def test_wzgledny_adres_xml_w_json_jest_znajdowany(self):
        spa = ('<html><body><script type="application/json">'
               '{"batchUrl":"batch/11588.xml"}</script></body></html>')

        def handler(srv, conn, path, req):
            if path.endswith(".xml"):
                raw_body(conn, RAW_XML)
            else:
                raw_body(conn, spa, ctype="text/html")

        server = self.serve(handler)
        refs = find_xml_links(
            self.opener(server),
            self.auction_ref(server, "auction/auk1/lot-abc12"),
        )
        self.assertEqual([r.url for r in refs],
                         [server.url("/auction/auk1/batch/11588.xml")])

    def test_atrybut_download_nie_jest_mylony_z_adresem(self):
        """``download="batch.xml"`` to nazwa pliku, nie ścieżka do pobrania."""
        page = ('<html><body><a href="/dl/7" download="batch-7.xml">'
                'Download Batch Details</a></body></html>')

        def handler(srv, conn, path, req):
            if path.startswith("/dl/"):
                raw_body(conn, RAW_XML)
            else:
                raw_body(conn, page, ctype="text/html")

        server = self.serve(handler)
        refs = find_xml_links(self.opener(server), self.auction_ref(server))
        self.assertEqual([r.url for r in refs], [server.url("/dl/7")])


class TestNazwZPolskimiZnakami(unittest.TestCase):
    """REGRESJA: nazwa spoza ASCII ma dalej identyfikować plik."""

    def test_transliteracja_zamiast_kasowania(self):
        self.assertEqual(safe_filename("zawartość pakietu.xml"), "zawartosc_pakietu.xml")
        self.assertEqual(safe_filename("ł ą ż.xml"), "l_a_z.xml")
        self.assertEqual(safe_filename("ąęó.xml"), "aeo.xml")

    def test_wynik_jest_dalej_czystym_ascii(self):
        for raw in ("zażółć gęślą.xml", "Straße.xml", "Ærø.xml", "日本語.xml"):
            with self.subTest(raw=raw):
                got = safe_filename(raw)
                self.assertTrue(all(ord(ch) < 128 for ch in got), got)
                self.assertTrue(got.endswith(".xml"))

    def test_nazwy_z_samych_znakow_niedrukowalnych_maja_zapas(self):
        self.assertEqual(safe_filename("日本語.xml"), "xml.xml")


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
