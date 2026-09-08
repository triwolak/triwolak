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
import io
import json
import os
import shutil
import sys
import tempfile
import threading
import time
import unittest
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
        ("plik z polskimi znakami ąęć.xml", "plik_z_polskimi_znakami_.xml"),
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

    def test_descend_auto_stops_when_auction_page_has_xml(self):
        opener = self.opener()
        refs = find_xml_links(opener, self.auction(AUCTION_1103))
        self.assertEqual([r.url for r in refs], [SITE.url("/media/batch-1103-all.xml")])
        self.assertEqual(SITE.state.count("/lot/thinkpad-t480-mix-bcd12"), 0)

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
        self.assertEqual(len(paths), 6)
        self.assertEqual(len(set(paths)), 6)
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
        "hashlib", "http", "os", "posixpath", "re", "socket", "sys", "time",
        "urllib", "dataclasses", "html", "typing", "__future__",
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


if __name__ == "__main__":  # pragma: no cover
    unittest.main()
