# -*- coding: utf-8 -*-
"""Testy ADWERSARIALNE modułu :mod:`flexit2xlsx.scrape`.

Cel: złamać scraper na wrogim serwerze HTTP zbudowanym na SUROWYM gnieździe TCP
(``http.server`` nie pozwala udawać obciętych odpowiedzi, gzipa wbrew
``Accept-Encoding`` ani sączenia bajtów po kropli).

Konwencja nazw: ``test_ok_*`` — zachowanie WYMAGANE, test pilnuje, żeby nie
było regresji.  (Historycznie były też testy ``test_blad_*`` utrwalające
znalezione defekty; po ich naprawieniu zostały przepisane na asercje stanu
poprawnego.)

Żaden test nie śpi naprawdę na backoffie — sen jest wstrzykiwany
(``make_opener(sleep=...)``).  Wyjątkiem są dwa testy „wolnej odpowiedzi”,
gdzie realny upływ czasu JEST przedmiotem pomiaru (łącznie < 5 s).
"""

from __future__ import annotations

import gzip
import os
import shutil
import socket
import sys
import tempfile
import threading
import time
import unittest
import xml.etree.ElementTree as ET

_HERE = os.path.dirname(os.path.abspath(__file__))
_ROOT = os.path.dirname(os.path.dirname(_HERE))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from flexit2xlsx import scrape  # noqa: E402


# ---------------------------------------------------------------------------
# Wrogi serwer na surowym gnieździe
# ---------------------------------------------------------------------------


class RawServer:
    """Serwer TCP, który odpowiada DOKŁADNIE tymi bajtami, które każe handler.

    Handler dostaje ``(serwer, gniazdo, ścieżka, surowe_żądanie)`` i sam pisze
    odpowiedź — można więc łamać HTTP do woli (obcięty ``Content-Length``,
    ``Content-Encoding`` wbrew negocjacji, sączenie po bajcie, brak nagłówków).
    """

    def __init__(self, handler):
        self.handler = handler
        self.sock = socket.socket()
        self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
        self.sock.bind(("127.0.0.1", 0))
        self.sock.listen(32)
        self.port = self.sock.getsockname()[1]
        self.hits = []                       # ścieżki kolejnych żądań
        self.cookies = []                    # (ścieżka, nagłówek Cookie)
        self._lock = threading.Lock()
        self._stop = False
        self._thread = threading.Thread(target=self._accept_loop, daemon=True)
        self._thread.start()

    # -- pętla akceptująca --------------------------------------------------

    def _accept_loop(self):
        while not self._stop:
            try:
                conn, _addr = self.sock.accept()
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
            first = raw.split(b"\r\n", 1)[0].decode("latin-1").split(" ")
            path = first[1] if len(first) > 1 else "/"
            cookie = ""
            for line in raw.decode("latin-1", "replace").split("\r\n"):
                if line.lower().startswith("cookie:"):
                    cookie = line.split(":", 1)[1].strip()
            with self._lock:
                self.hits.append(path)
                self.cookies.append((path, cookie))
            self.handler(self, conn, path, raw)
        except Exception:                    # serwer testowy nigdy nie wywraca testu
            pass
        finally:
            try:
                conn.close()
            except OSError:
                pass

    # -- API dla testów -----------------------------------------------------

    @property
    def base(self) -> str:
        return "http://127.0.0.1:%d/" % self.port

    def url(self, path: str) -> str:
        return "http://127.0.0.1:%d%s" % (self.port, path)

    def count(self, path: str) -> int:
        with self._lock:
            return self.hits.count(path)

    def close(self):
        self._stop = True
        try:
            self.sock.close()
        except OSError:
            pass


def send_raw(conn, status="200 OK", headers=(), body=b""):
    """Wysyła odpowiedź HTTP z podanymi nagłówkami (bez żadnej magii)."""
    head = "HTTP/1.1 %s\r\n" % status
    for key, value in headers:
        head += "%s: %s\r\n" % (key, value)
    conn.sendall(head.encode("latin-1") + b"\r\n" + body)


def send_html(conn, body, status="200 OK", ctype="text/html; charset=utf-8", extra=()):
    """Wysyła stronę HTML z poprawnym ``Content-Length``."""
    if isinstance(body, str):
        body = body.encode("utf-8")
    send_raw(conn, status,
             list(extra) + [("Content-Type", ctype), ("Content-Length", str(len(body)))],
             body)


def send_xml(conn, body=None, ctype="application/xml", extra=()):
    """Wysyła plik XML."""
    body = XML if body is None else body
    send_raw(conn, "200 OK",
             list(extra) + [("Content-Type", ctype), ("Content-Length", str(len(body)))],
             body)


#: Wzorcowy „plik zawartości pakietu”.
XML = (b"<?xml version='1.0' encoding='UTF-8'?>\n"
       b"<batch><item><model>ThinkPad T480</model><serial>007</serial></item>"
       b"<item><model>Latitude 5400</model><serial>008</serial></item></batch>\n")


class FakeClock:
    """Zegar + sen bez realnego czekania (dowód, że backoff jest wstrzykiwalny)."""

    def __init__(self):
        self.now = 0.0
        self.slept = []

    def sleep(self, seconds):
        self.slept.append(round(float(seconds), 3))
        self.now += float(seconds)

    def monotonic(self):
        return self.now


class Base(unittest.TestCase):
    """Wspólna infrastruktura: serwer, opener z fałszywym snem, katalog tymczasowy."""

    def serve(self, handler) -> RawServer:
        server = RawServer(handler)
        self.addCleanup(server.close)
        return server

    def opener(self, server, **kw):
        kw.setdefault("delay", 0.0)
        kw.setdefault("site", server.base)
        self.clock = FakeClock()
        kw.setdefault("sleep", self.clock.sleep)
        kw.setdefault("clock", self.clock.monotonic)
        return scrape.make_opener(**kw)

    def tmpdir(self) -> str:
        path = tempfile.mkdtemp(prefix="flexit-adv-")
        self.addCleanup(shutil.rmtree, path, True)
        return path

    @staticmethod
    def auction(url, ident="auk1", title="Aukcja testowa"):
        return scrape.AuctionRef(url=url, id=ident, title=title)

    def quiet(self):
        """Wycisza ``stderr`` modułu (scrape sypie ostrzeżeniami do diagnostyki)."""
        import io
        saved = sys.stderr
        sys.stderr = io.StringIO()
        self.addCleanup(lambda: setattr(sys, "stderr", saved))
        return sys.stderr


# ===========================================================================
# 1. Integralność pobranych danych
# ===========================================================================


class TestIntegralnoscDanych(Base):
    """Czy to, co wyląduje na dysku, jest tym, co serwer obiecał wysłać?"""

    def test_ok_obcieta_odpowiedz_nie_laduje_na_dysku(self):
        """Niedobór wobec ``Content-Length`` to błąd sieci, nie „gotowy plik''.

        Serwer deklaruje 5085 bajtów, wysyła 45 i zamyka połączenie.
        ``http.client`` w tej sytuacji NIE rzuca ``IncompleteRead`` (świadoma
        decyzja CPythona), więc bez własnej kontroli ogryzek trafiłby na dysk
        jako komplet i przy kolejnym uruchomieniu zostałby POMINIĘTY —
        uszkodzenie byłoby trwałe.
        """
        def handler(srv, conn, path, req):
            head = ("HTTP/1.1 200 OK\r\nContent-Type: application/xml\r\n"
                    "Content-Length: %d\r\n\r\n" % (len(XML) + 5000))
            conn.sendall(head.encode("latin-1") + XML[:45])
            conn.close()

        server = self.serve(handler)
        opener = self.opener(server, retries=0)
        out = self.tmpdir()
        ref = scrape.XmlRef(url=server.base + "batch.xml", filename="batch.xml",
                            auction_id="auk1", auction_title="Aukcja")

        with self.assertRaises(scrape.FetchError) as ctx:
            scrape.download_xml(opener, ref, out)
        self.assertIn("IncompleteRead", str(ctx.exception))
        self.assertEqual(os.listdir(out), [], "nic nie mogło trafić na dysk")

    def test_ok_urwane_polaczenie_w_srodku_ciala_daje_wyjatek(self):
        """Zerwanie w połowie transmisji kończy się ``FetchError``, nie ciszą.

        Serwer deklaruje 40 bajtów, wysyła 12 i rozłącza się.
        """
        def handler(srv, conn, path, req):
            conn.sendall(b"HTTP/1.1 200 OK\r\nContent-Type: application/xml\r\n"
                         b"Content-Length: 40\r\n\r\n")
            conn.sendall(b"<?xml versi")
            conn.close()

        server = self.serve(handler)
        opener = self.opener(server, retries=0)
        with self.assertRaises(scrape.FetchError):
            scrape.fetch(opener, server.base + "b.xml")

    def test_ok_urwane_chunked_jest_wykrywane(self):
        """Kodowanie ``chunked`` urwane w środku JEST wykrywane (IncompleteRead)."""
        def handler(srv, conn, path, req):
            conn.sendall(b"HTTP/1.1 200 OK\r\nContent-Type: text/xml\r\n"
                         b"Transfer-Encoding: chunked\r\n\r\n")
            conn.sendall(b"%x\r\n" % 20 + XML[:20] + b"\r\n")
            conn.close()

        server = self.serve(handler)
        opener = self.opener(server, retries=0)
        with self.assertRaises(scrape.FetchError):
            scrape.fetch(opener, server.base + "b.xml")

    def test_ok_chunked_kompletny(self):
        """Poprawne ``chunked`` (bez ``Content-Length``) jest sklejane w całość."""
        def handler(srv, conn, path, req):
            conn.sendall(b"HTTP/1.1 200 OK\r\nContent-Type: text/xml\r\n"
                         b"Transfer-Encoding: chunked\r\n\r\n")
            for part in (XML[:30], XML[30:70], XML[70:]):
                conn.sendall(b"%x\r\n" % len(part) + part + b"\r\n")
            conn.sendall(b"0\r\n\r\n")

        server = self.serve(handler)
        data, _ = scrape.fetch(self.opener(server), server.base + "b.xml")
        self.assertEqual(data, XML)

    def test_ok_brak_content_length_http10(self):
        """HTTP/1.0 bez ``Content-Length`` (koniec = rozłączenie) działa."""
        def handler(srv, conn, path, req):
            conn.sendall(b"HTTP/1.0 200 OK\r\nContent-Type: text/xml\r\n\r\n" + XML)
            conn.close()

        server = self.serve(handler)
        data, _ = scrape.fetch(self.opener(server), server.base + "b.xml")
        self.assertEqual(data, XML)

    def test_ok_ogromna_odpowiedz_jest_ucinana_wyjatkiem(self):
        """Odpowiedź ponad ``max_bytes`` przerywa pobieranie, a nie pamięć."""
        big = b"<?xml version='1.0'?><r>" + b"<i>x</i>" * 200000 + b"</r>"

        def handler(srv, conn, path, req):
            send_xml(conn, big)

        server = self.serve(handler)
        opener = self.opener(server, max_bytes=100000, retries=0)
        with self.assertRaises(scrape.ScrapeError) as ctx:
            scrape.fetch(opener, server.base + "big.xml")
        self.assertIn("limit", str(ctx.exception))


# ===========================================================================
# 2. Kodowanie transportowe i rozpoznawanie XML
# ===========================================================================


class TestKodowanieIRozpoznawanie(Base):
    """gzip, UTF-16, ``Content-Type`` niezgodny z treścią."""

    def test_ok_gzip_jest_rozpakowywany(self):
        """``Content-Encoding: gzip`` jest rozpakowywany przez :mod:`gzip`.

        Opener prosi o ``Accept-Encoding: identity``, ale CDN-y (Cloudflare,
        nginx z ``gzip_static``) i tak pakują.  Bez rozpakowania użytkownik
        dostawał komunikat o STRONIE LOGOWANIA i radę ``--cookie`` — diagnozę
        całkowicie mylącą.
        """
        body = gzip.compress(XML)

        def handler(srv, conn, path, req):
            send_raw(conn, "200 OK",
                     [("Content-Type", "application/xml"),
                      ("Content-Encoding", "gzip"),
                      ("Content-Length", str(len(body)))], body)

        server = self.serve(handler)
        opener = self.opener(server)

        data, ctype = scrape.fetch(opener, server.base + "batch.xml")
        self.assertEqual(ctype, "application/xml")
        self.assertEqual(data, XML)

        ref = scrape.XmlRef(url=server.base + "batch.xml", filename="b.xml",
                            auction_id="auk1", auction_title="A")
        out = self.tmpdir()
        sciezka = scrape.download_xml(opener, ref, out)
        with open(sciezka, "rb") as uchwyt:
            self.assertEqual(uchwyt.read(), XML)

    def test_ok_gzip_na_stronie_html_nie_gubi_linkow(self):
        """Spakowana lista aukcji jest rozpakowywana i linki się znajdują."""
        page = ('<html><body><a href="/auction/flexit-auctions-18-06-2026-1103">'
                'Aukcja</a></body></html>').encode("utf-8")
        body = gzip.compress(page)

        def handler(srv, conn, path, req):
            send_raw(conn, "200 OK",
                     [("Content-Type", "text/html; charset=utf-8"),
                      ("Content-Encoding", "gzip"),
                      ("Content-Length", str(len(body)))], body)

        server = self.serve(handler)
        self.quiet()
        found = scrape.discover_auctions(self.opener(server), server.base)
        self.assertEqual(len(found), 1)
        self.assertEqual(found[0].id, "flexit-auctions-18-06-2026-1103")

    def test_ok_xml_w_utf16_jest_rozpoznawany(self):
        """XML w UTF-16 z BOM jest rozpoznawany i zapisywany.

        ``_looks_like_xml`` musi znać BOM-y UTF-16/UTF-32, bo eksporty
        z narzędzi Windows bardzo często są w UTF-16LE, a kontrakt
        ``xmlflatten`` wymienia takie pliki wprost.
        """
        body = "<?xml version='1.0' encoding='UTF-16'?><batch><i>ąćż</i></batch>".encode("utf-16")

        def handler(srv, conn, path, req):
            send_xml(conn, body)

        server = self.serve(handler)
        opener = self.opener(server)

        self.assertEqual(ET.fromstring(body).tag, "batch")
        self.assertTrue(scrape._looks_like_xml(body))
        ref = scrape.XmlRef(url=server.base + "batch.xml", filename="b.xml",
                            auction_id="auk1", auction_title="A")
        out = self.tmpdir()
        sciezka = scrape.download_xml(opener, ref, out)
        with open(sciezka, "rb") as uchwyt:
            self.assertEqual(uchwyt.read(), body)

    def test_ok_xml_podany_jako_text_html_bez_rozszerzenia_jest_znajdowany(self):
        """Endpoint API zwracający XML z ``Content-Type: text/html`` też się liczy.

        SITE_NOTES ostrzega, że „Download Batch Details'' może być endpointem
        API bez rozszerzenia ``.xml``, a taki endpoint bardzo często ma źle
        ustawiony typ MIME.  Skoro treść i tak została pobrana, trzeba w nią
        zajrzeć (``_looks_like_xml``), zamiast wierzyć nagłówkowi.
        """
        page = '<html><body><a href="/api/lot/11588/batch">Download Batch Details</a></body></html>'

        def handler(srv, conn, path, req):
            if path.startswith("/api/"):
                send_xml(conn, XML, ctype="text/html; charset=utf-8")
            else:
                send_html(conn, page)

        server = self.serve(handler)
        self.quiet()
        opener = self.opener(server)
        refs = scrape.find_xml_links(opener, self.auction(server.base + "auction/auk1"))
        self.assertEqual(len(refs), 1)
        self.assertTrue(refs[0].url.endswith("/api/lot/11588/batch"))
        sciezka = scrape.download_xml(opener, refs[0], self.tmpdir())
        with open(sciezka, "rb") as uchwyt:
            self.assertEqual(uchwyt.read(), XML)

    def test_ok_xml_jako_text_html_ale_z_rozszerzeniem_dziala(self):
        """Ten sam serwer, ale adres kończy się ``.xml`` — wtedy jest dobrze."""
        page = '<html><body><a href="/api/lot/11588/batch.xml">Download Batch Details</a></body></html>'

        def handler(srv, conn, path, req):
            if path.startswith("/api/"):
                send_xml(conn, XML, ctype="text/html")
            else:
                send_html(conn, page)

        server = self.serve(handler)
        opener = self.opener(server)
        refs = scrape.find_xml_links(opener, self.auction(server.base + "auction/auk1"))
        self.assertEqual(len(refs), 1)
        path = scrape.download_xml(opener, refs[0], self.tmpdir())
        with open(path, "rb") as handle:
            self.assertEqual(handle.read(), XML)


# ===========================================================================
# 3. Przekierowania
# ===========================================================================


class TestPrzekierowania(Base):

    def test_ok_redirect_na_obca_domene_blokowany_bez_wycieku_ciasteczka(self):
        """302 poza witrynę = ``SiteBlockedError``; ciasteczko NIE wychodzi."""
        def handler(srv, conn, path, req):
            if path == "/go":
                send_raw(conn, "302 Found",
                         [("Location", "http://localhost:%d/evil" % srv.port),
                          ("Content-Length", "0")])
            else:
                send_html(conn, "<html>evil</html>")

        server = self.serve(handler)
        opener = self.opener(server, cookie="SESSION=tajne")
        with self.assertRaises(scrape.SiteBlockedError):
            scrape.fetch(opener, server.base + "go")
        self.assertEqual(server.hits, ["/go"], "obca domena nie została odpytana")

    def test_ok_petla_przekierowan_konczy_sie_bledem(self):
        """Nieskończona pętla 302 kończy się ``FetchError``, nie zawieszeniem."""
        def handler(srv, conn, path, req):
            number = path.rsplit("/", 1)[-1] or "0"
            send_raw(conn, "302 Found",
                     [("Location", "/r/%d" % (int(number) + 1)), ("Content-Length", "0")])

        server = self.serve(handler)
        opener = self.opener(server, retries=1)
        with self.assertRaises(scrape.FetchError):
            scrape.fetch(opener, server.base + "r/0")
        self.assertLess(len(server.hits), 40, "urllib ograniczył liczbę skoków")

    def test_ok_301_na_samego_siebie(self):
        """301 wskazujące na ten sam adres nie kręci się w kółko."""
        def handler(srv, conn, path, req):
            send_raw(conn, "301 Moved Permanently",
                     [("Location", srv.url("/self")), ("Content-Length", "0")])

        server = self.serve(handler)
        with self.assertRaises(scrape.FetchError):
            scrape.fetch(self.opener(server, retries=1), server.base + "self")

    def test_ok_redirect_na_strone_logowania_konczy_sie_czytelnym_bledem(self):
        """Ściana logowania (302 -> /login) daje ``NotXmlError`` z radą ``--cookie``."""
        def handler(srv, conn, path, req):
            if path.endswith(".xml"):
                send_raw(conn, "302 Found", [("Location", "/login"), ("Content-Length", "0")])
            elif path == "/login":
                send_html(conn, "<html><head><title>Zaloguj</title></head>"
                                "<body><form action='/login'></form></body></html>")
            else:
                send_html(conn, '<html><body><a href="/batch.xml">'
                                'Download Batch Details</a></body></html>')

        server = self.serve(handler)
        opener = self.opener(server)
        refs = scrape.find_xml_links(opener, self.auction(server.base + "auction/auk1"))
        self.assertEqual(len(refs), 1)
        out = self.tmpdir()
        with self.assertRaises(scrape.NotXmlError) as ctx:
            scrape.download_xml(opener, refs[0], out)
        self.assertIn("--cookie", str(ctx.exception))
        self.assertEqual(os.listdir(out), [], "śmieć nie ląduje na dysku")


# ===========================================================================
# 4. Ponawianie, backoff, kody błędów
# ===========================================================================


class TestPonawianie(Base):

    def test_ok_backoff_jest_wstrzykiwalny_i_testy_nie_spia(self):
        """503 -> 4 ponowienia z opóźnieniami 2/4/8/16 s, ZERO realnego snu."""
        def handler(srv, conn, path, req):
            send_raw(conn, "503 Service Unavailable", [("Content-Length", "0")])

        server = self.serve(handler)
        opener = self.opener(server)
        started = time.monotonic()
        with self.assertRaises(scrape.FetchError) as ctx:
            scrape.fetch(opener, server.base + "x")
        elapsed = time.monotonic() - started
        self.assertEqual(ctx.exception.status, 503)
        self.assertEqual(len(server.hits), 5, "1 próba + 4 ponowienia")
        self.assertEqual(self.clock.slept, [2.0, 4.0, 8.0, 16.0])
        self.assertLess(elapsed, 2.0, "30 s backoffu przespane W CAŁOŚCI wirtualnie")

    def test_ok_modulowy_SLEEP_FUNCTION_tez_jest_podmienialny(self):
        """Bez ``sleep=`` backoff idzie przez modułowe ``scrape.SLEEP_FUNCTION``."""
        def handler(srv, conn, path, req):
            send_raw(conn, "500 Internal Server Error", [("Content-Length", "0")])

        server = self.serve(handler)
        slept = []
        original = scrape.SLEEP_FUNCTION
        scrape.SLEEP_FUNCTION = slept.append
        self.addCleanup(lambda: setattr(scrape, "SLEEP_FUNCTION", original))
        opener = scrape.make_opener(delay=0.0, site=server.base, retries=2)
        started = time.monotonic()
        with self.assertRaises(scrape.FetchError):
            scrape.fetch(opener, server.base + "x")
        self.assertEqual(slept, [2.0, 4.0])
        self.assertLess(time.monotonic() - started, 2.0)

    def test_ok_404_nie_jest_ponawiane(self):
        """404 to porażka trwała — jedno żądanie, zero snu."""
        def handler(srv, conn, path, req):
            send_raw(conn, "404 Not Found", [("Content-Length", "0")])

        server = self.serve(handler)
        opener = self.opener(server)
        with self.assertRaises(scrape.FetchError) as ctx:
            scrape.fetch(opener, server.base + "nie-ma")
        self.assertEqual(ctx.exception.status, 404)
        self.assertEqual(len(server.hits), 1)
        self.assertEqual(self.clock.slept, [])

    def test_ok_429_honoruje_retry_after(self):
        """``Retry-After: 7`` wygrywa z wykładniczym backoffem."""
        state = {"n": 0}

        def handler(srv, conn, path, req):
            state["n"] += 1
            if state["n"] <= 2:
                send_raw(conn, "429 Too Many Requests",
                         [("Retry-After", "7"), ("Content-Length", "0")])
            else:
                send_xml(conn)

        server = self.serve(handler)
        opener = self.opener(server)
        data, _ = scrape.fetch(opener, server.base + "x.xml")
        self.assertEqual(data, XML)
        self.assertEqual(self.clock.slept, [7.0, 7.0])

    def test_ok_timeout_ogranicza_calkowity_czas_pobrania(self):
        """``--timeout`` obejmuje też CAŁĄ odpowiedź, nie tylko jeden odczyt.

        Serwer wysyła nagłówki natychmiast, a potem sączy po jednym bajcie.
        Żaden pojedynczy odczyt nie przekracza limitu gniazda, więc bez
        osobnego terminu dla całej odpowiedzi wrogi (albo przeciążony) portal
        mógłby trzymać narzędzie dowolnie długo — a przy ``--per-auction``
        dla każdego lotu z osobna.  Termin to
        ``timeout * scrape.RESPONSE_TIMEOUT_FACTOR``.
        """
        def handler(srv, conn, path, req):
            conn.sendall(b"HTTP/1.1 200 OK\r\nContent-Type: application/xml\r\n"
                         b"Content-Length: 400\r\n\r\n")
            for _ in range(400):
                try:
                    conn.sendall(b"x")
                    time.sleep(0.02)
                except OSError:
                    return

        server = self.serve(handler)
        # PRAWDZIWY zegar: termin odpowiedzi liczy się w czasie rzeczywistym,
        # a nie w krokach fałszywego ``sleep`` (ten podczas czytania nie działa).
        opener = self.opener(server, timeout=0.2, retries=0, clock=time.monotonic)
        limit = 0.2 * scrape.RESPONSE_TIMEOUT_FACTOR
        naturalny_czas = 400 * 0.02      # ile trwałoby sączenie do końca
        started = time.monotonic()
        with self.assertRaises(scrape.FetchError) as pulapka:
            scrape.fetch(opener, server.base + "drip.xml")
        elapsed = time.monotonic() - started
        self.assertIn("czas odpowiedzi", str(pulapka.exception))
        self.assertLess(elapsed, naturalny_czas / 2,
                        "pobranie trwało %.1f s przy terminie %.1f s"
                        % (elapsed, limit))

    def test_ok_brak_odpowiedzi_konczy_sie_timeoutem_i_ponowieniem(self):
        """Serwer milczący w ogóle: timeout gniazda -> ponowienie -> FetchError."""
        def handler(srv, conn, path, req):
            time.sleep(2.0)

        server = self.serve(handler)
        opener = self.opener(server, timeout=0.3, retries=1)
        with self.assertRaises(scrape.FetchError):
            scrape.fetch(opener, server.base + "cisza")
        self.assertEqual(len(server.hits), 2)
        self.assertEqual(self.clock.slept, [2.0])


# ===========================================================================
# 5. Parsowanie HTML i wyszukiwanie linków
# ===========================================================================


class TestParsowaniaHtml(Base):

    def test_ok_zle_zagniezdzony_html(self):
        """Niedomknięte ``<li>``, ``<a>`` przeplecione z ``<div>``, ``<table>`` bez ``</tr>``."""
        page = """<html><body>
<ul><li><a href="/auction/aa-1001"><b>Aukcja AA<li><a href='/auction/bb-1002'>Aukcja BB
<div><a href=/auction/cc-1003 >Aukcja CC</div></a></b>
<table><tr><td><a href="/auction/dd-1004">Aukcja DD</table>
<a href="/auction/ee-1005" >Aukcja EE
</body>"""

        def handler(srv, conn, path, req):
            send_html(conn, page)

        server = self.serve(handler)
        found = scrape.discover_auctions(self.opener(server), server.base)
        self.assertEqual(sorted(a.id for a in found),
                         ["aa-1001", "bb-1002", "cc-1003", "dd-1004", "ee-1005"])

    def test_ok_linki_wzgledne_z_dwiema_kropkami(self):
        """``../``, ``./`` i ``../../`` są rozwijane względem adresu lotu."""
        page = ('<html><body><a href="../../files/pakiet-a.xml">A</a>'
                '<a href="./pakiet-b.xml">B</a><a href="../pakiet-c.xml">C</a></body></html>')

        def handler(srv, conn, path, req):
            send_xml(conn) if path.endswith(".xml") else send_html(conn, page)

        server = self.serve(handler)
        refs = scrape.find_xml_links(self.opener(server),
                                     self.auction(server.url("/auction/auk1/lot-abc12")))
        paths = sorted(r.url[len(server.base) - 1:] for r in refs)
        self.assertEqual(paths, ["/auction/auk1/pakiet-b.xml",
                                 "/auction/pakiet-c.xml",
                                 "/files/pakiet-a.xml"])
        for ref in refs:
            self.assertNotIn("..", ref.filename)

    def test_ok_linki_w_javascript_bezwzgledne_i_escapowane(self):
        """SPA: adresy w ``__NEXT_DATA__`` z ``\\u002F`` są odnajdywane."""
        spa = ('<html><body><div id="root"></div>'
               '<script id="__NEXT_DATA__" type="application/json">'
               '{"a":"\\u002Fapi\\u002Flot\\u002F11588\\u002Fbatch.xml",'
               '"b":"/media/pakiet-2.xml"}</script></body></html>')

        def handler(srv, conn, path, req):
            send_xml(conn) if path.endswith(".xml") else send_html(conn, spa)

        server = self.serve(handler)
        refs = scrape.find_xml_links(self.opener(server),
                                     self.auction(server.base + "auction/auk1"))
        self.assertEqual(sorted(r.url[len(server.base) - 1:] for r in refs),
                         ["/api/lot/11588/batch.xml", "/media/pakiet-2.xml"])

    def test_ok_wzgledny_adres_xml_w_javascript_jest_znajdowany(self):
        """Adres WZGLĘDNY w osadzonym JSON-ie (``"batch/11588.xml"``) też się liczy.

        To postać typowa dla danych Next.js, a SITE_NOTES każe szukać „w całej
        treści adresów pasujących do ``\\.xml``''.
        """
        spa = ('<html><body><script type="application/json">'
               '{"batchUrl":"batch/11588.xml"}</script></body></html>')

        def handler(srv, conn, path, req):
            send_xml(conn) if path.endswith(".xml") else send_html(conn, spa)

        server = self.serve(handler)
        self.quiet()
        refs = scrape.find_xml_links(self.opener(server),
                                     self.auction(server.base + "auction/auk1/lot-abc12"))
        self.assertEqual(len(refs), 1)
        self.assertTrue(refs[0].url.endswith("/auction/auk1/batch/11588.xml"),
                        refs[0].url)

    def test_ok_obca_domena_w_javascript_jest_odrzucana(self):
        """Adres do obcej domeny w JSON-ie nie generuje żądania sieciowego."""
        spa = ('<html><body><script type="application/json">'
               '{"x":"https://evil.example.invalid/batch.xml"}</script></body></html>')

        def handler(srv, conn, path, req):
            send_html(conn, spa)

        server = self.serve(handler)
        self.quiet()
        refs = scrape.find_xml_links(self.opener(server),
                                     self.auction(server.base + "auction/auk1"))
        self.assertEqual(refs, [])

    def test_ok_base_href_na_obca_domene_nie_wyprowadza_scrapera(self):
        """``<base href>`` wskazujące poza witrynę nie przekierowuje pobierania."""
        page = ('<html><head><base href="https://evil.example.invalid/"></head>'
                '<body><a href="batch.xml">Batch</a></body></html>')

        def handler(srv, conn, path, req):
            send_html(conn, page)

        server = self.serve(handler)
        self.quiet()
        refs = scrape.find_xml_links(self.opener(server),
                                     self.auction(server.base + "auction/auk1"))
        self.assertEqual(refs, [])
        self.assertEqual(server.hits, ["/auction/auk1"])

    def test_ok_smieci_zamiast_html_nie_wywracaja_diagnostyki(self):
        """``describe_page`` na binarnych śmieciach zwraca raport, nie wyjątek."""
        report = scrape.describe_page(b"\x00\x01\x02\xff\xfe\x80 garbage \xc3(",
                                      "http://x.invalid/",
                                      content_type="application/octet-stream")
        self.assertIn("diagnostyka strony", report)


# ===========================================================================
# 6. Paginacja i duplikaty
# ===========================================================================


class TestPaginacjaIDuplikaty(Base):

    def test_ok_paginacja_bez_konca_jest_ograniczona(self):
        """Paginator generujący w nieskończoność ``rel=next`` zatrzymuje się na ``max_pages``."""
        def handler(srv, conn, path, req):
            number = 1
            if "page=" in path:
                number = int(path.split("page=")[1].split("&")[0])
            send_html(conn, '<html><body><a href="/auction/a-%d">A%d</a>'
                            '<a rel="next" href="/?page=%d">Dalej</a></body></html>'
                            % (number, number, number + 1))

        server = self.serve(handler)
        started = time.monotonic()
        found = scrape.discover_auctions(self.opener(server), server.base, max_pages=7)
        self.assertEqual(len(server.hits), 7)
        self.assertEqual(len(found), 7)
        self.assertLess(time.monotonic() - started, 5.0)

    def test_ok_paginacja_w_kolko_nie_zapetla(self):
        """Strony 1<->2 linkujące nawzajem: każdy adres odwiedzany raz."""
        def handler(srv, conn, path, req):
            other = "/?page=2" if "page=2" not in path else "/?page=1"
            send_html(conn, '<html><body><a href="/auction/x-1">X</a>'
                            '<a rel="next" href="%s">Dalej</a></body></html>' % other)

        server = self.serve(handler)
        scrape.discover_auctions(self.opener(server), server.base, max_pages=50)
        self.assertLessEqual(len(server.hits), 3)

    def test_ok_ta_sama_aukcja_pod_wieloma_url_deduplikowana(self):
        """``/auction/x``, ``/auction/x/``, ``/auction/x/info``, ``?order=`` -> jedna aukcja."""
        slug = "flexit-auctions-18-06-2026-1103"
        page = ("<html><body>"
                '<a href="/auction/{s}/info">Info</a>'
                '<a href="/auction/{s}">Aukcja czerwcowa</a>'
                '<a href="/auction/{s}/">Aukcja (slash)</a>'
                '<a href="/auction/{s}?order=closedSoonest">Sortowanie</a>'
                "</body></html>").format(s=slug)

        def handler(srv, conn, path, req):
            send_html(conn, page)

        server = self.serve(handler)
        found = scrape.discover_auctions(self.opener(server), server.base)
        self.assertEqual([a.id for a in found], [slug])
        self.assertEqual(found[0].url, server.url("/auction/%s" % slug),
                         "wybrany najbardziej kanoniczny adres")

    def test_ok_ten_sam_lot_pod_dwoma_url_pobierany_raz(self):
        """Ryzyko 5 ze SITE_NOTES: ``/lot/<slug>-<hash>`` i ``/auction/../<slug>-<hash>``."""
        auction_page = ('<html><body><a href="/lot/lenovo-mix-11588">Lot</a>'
                        '<a href="/auction/auk1/lenovo-mix-11588">Ten sam lot</a></body></html>')
        lot_page = ('<html><body><a href="/api/lot/11588/batch.xml" download>'
                    'Download Batch Details</a></body></html>')

        def handler(srv, conn, path, req):
            if path.endswith(".xml"):
                send_xml(conn)
            elif "11588" in path:
                send_html(conn, lot_page)
            else:
                send_html(conn, auction_page)

        server = self.serve(handler)
        refs = scrape.find_xml_links(self.opener(server),
                                     self.auction(server.url("/auction/auk1")))
        self.assertEqual(len(refs), 1)
        self.assertEqual(len(server.hits), 2, "strona aukcji + JEDNA strona lotu")

    def test_ok_niepewny_link_nie_jest_pobierany_dwa_razy(self):
        """Sondowanie ``Content-Type`` zapamiętuje treść — bez drugiego pobrania.

        ``_confirm_by_content_type`` robi zwykłe GET (nie HEAD, nie ``Range``),
        więc treść jest już w pamięci; ``download_xml`` musi z niej skorzystać.
        Każdy link bez ``.xml`` w adresie (czyli dokładnie przycisk „Download
        Batch Details'' z SITE_NOTES) kosztowałby inaczej podwójny transfer
        i podwójne uprzejme opóźnienie ``--delay``.
        """
        page = ('<html><body><a href="/api/batch" download>'
                'Download Batch Details</a></body></html>')

        def handler(srv, conn, path, req):
            if path.startswith("/api"):
                send_xml(conn, XML, ctype="application/octet-stream")
            else:
                send_html(conn, page)

        server = self.serve(handler)
        opener = self.opener(server)
        refs = scrape.find_xml_links(opener, self.auction(server.base + "auction/a1", "a1"))
        self.assertEqual(len(refs), 1)
        sciezka = scrape.download_xml(opener, refs[0], self.tmpdir())
        self.assertEqual(server.count("/api/batch"), 1,
                         "treść z sondy ma wystarczyć — żadnego drugiego GET-a")
        with open(sciezka, "rb") as uchwyt:
            self.assertEqual(uchwyt.read(), XML)


# ===========================================================================
# 7. Nazwy plików
# ===========================================================================


class TestNazwPlikow(Base):

    def test_ok_path_traversal_i_nul_w_nazwie(self):
        """``../``, ``..%2F``, NUL, ścieżki Windows i nazwy zarezerwowane."""
        cases = {
            "../../../etc/passwd": "passwd.xml",
            "..%2F..%2Fetc%2Fshadow": "shadow.xml",
            "%2e%2e%2f%2e%2e%2fx.xml": "x.xml",
            "a\x00b.xml": "ab.xml",
            "\x00\x00\x00": "plik.xml",
            "C:\\Windows\\system32\\evil.xml": "evil.xml",
            "....//....//x.xml": "x.xml",
            "con.xml": "_con.xml",
            "": "plik.xml",
            "/": "plik.xml",
            "..": "plik.xml",
        }
        for raw, expected in cases.items():
            with self.subTest(nazwa=raw):
                got = scrape.safe_filename(raw)
                self.assertEqual(got, expected)
                self.assertNotIn("/", got)
                self.assertNotIn("\\", got)
                self.assertNotIn("\x00", got)
                self.assertNotIn("..", got)

    def test_ok_traversal_przez_atrybut_download_nie_wychodzi_z_katalogu(self):
        """``download="../../../../tmp/EVIL.xml"`` zapisuje się WEWNĄTRZ ``out_dir``."""
        page = ('<html><body><a href="/dl?id=1" download="../../../../tmp/EVIL.xml">'
                'Batch</a></body></html>')

        def handler(srv, conn, path, req):
            send_xml(conn) if path.startswith("/dl") else send_html(conn, page)

        server = self.serve(handler)
        opener = self.opener(server)
        refs = scrape.find_xml_links(opener, self.auction(server.base + "auction/auk1"))
        self.assertEqual(len(refs), 1)
        self.assertNotIn("..", refs[0].filename)
        out = self.tmpdir()
        saved = scrape.download_xml(opener, refs[0], out)
        self.assertEqual(os.path.dirname(os.path.realpath(saved)), os.path.realpath(out))
        self.assertFalse(os.path.exists("/tmp/EVIL.xml"))

    def test_ok_nul_w_href_nie_wywraca_wyszukiwania(self):
        """Bajt NUL w ``href`` nie przerywa przetwarzania pozostałych linków."""
        page = '<html><body><a href="/ba\x00tch.xml">Zły</a><a href="/ok.xml">OK</a></body></html>'

        def handler(srv, conn, path, req):
            send_xml(conn) if path.endswith(".xml") and "\x00" not in path else send_html(conn, page)

        server = self.serve(handler)
        opener = self.opener(server, retries=0)
        refs = scrape.find_xml_links(opener, self.auction(server.base + "auction/auk1"))
        names = sorted(r.filename for r in refs)
        self.assertEqual(names, ["auk1-batch.xml", "auk1-ok.xml"])
        for name in names:
            self.assertNotIn("\x00", name)

    def test_ok_dwa_pliki_o_tej_samej_nazwie_w_jednej_aukcji(self):
        """Dwa ``batch.xml`` z różnych lotów dostają różne nazwy (``-2``)."""
        auction_page = ('<html><body><a href="/lot/a-aaa11">A</a>'
                        '<a href="/lot/b-bbb22">B</a></body></html>')

        def handler(srv, conn, path, req):
            if path.endswith(".xml"):
                send_xml(conn)
            elif path.startswith("/lot/"):
                slug = path.rsplit("/", 1)[-1]
                send_html(conn, '<html><body><a href="/files/%s/batch.xml">Batch</a>'
                                '</body></html>' % slug)
            else:
                send_html(conn, auction_page)

        server = self.serve(handler)
        opener = self.opener(server)
        refs = scrape.find_xml_links(opener, self.auction(server.url("/auction/auk1")))
        self.assertEqual(len(refs), 2)
        self.assertEqual(len({r.filename for r in refs}), 2, "nazwy nie kolidują")
        out = self.tmpdir()
        for ref in refs:
            scrape.download_xml(opener, ref, out)
        self.assertEqual(len(os.listdir(out)), 2, "oba pliki na dysku")

    def test_ok_nazwa_z_polskimi_znakami_zachowuje_tozsamosc(self):
        """Nazwa spoza ASCII nie może zredukować się do ``xml.xml``.

        ``zawartość pakietu.xml`` po transliteracji zostaje czytelną nazwą,
        więc kolumna „Plik'' w arkuszu dalej identyfikuje lot.  Wynik mieści
        się w ``[\\w.-]``, tak jak wymaga kontrakt (INTERFACES.md).
        """
        self.assertEqual(scrape.safe_filename("zawartość pakietu.xml"),
                         "zawartosc_pakietu.xml")
        self.assertEqual(scrape.safe_filename("ł ą ż.xml"), "l_a_z.xml")
        self.assertEqual(scrape.safe_filename("ąęó.xml"), "aeo.xml")


# ===========================================================================
# 8. Uwierzytelnianie
# ===========================================================================


class TestUwierzytelniania(Base):

    def test_ok_cookie_z_opcji_laczy_sie_z_ciasteczkami_serwera(self):
        """``--cookie`` UZUPEŁNIA ``http.cookiejar``, a nie go unieważnia.

        Gdyby ``Opener.headers()`` wstawiało nagłówek ``Cookie`` na sztywno,
        ``HTTPCookieProcessor`` (który dokłada ciasteczka sesji tylko wtedy,
        gdy żądanie nagłówka ``Cookie`` jeszcze nie ma) zostałby wyłączony:
        token CSRF, ``cf_clearance`` czy świeży identyfikator sesji NIGDY nie
        wracałyby na serwer — dokładnie w scenariuszu, dla którego ``--cookie``
        powstało, portal odpowiadałby 403.
        """
        def handler(srv, conn, path, req):
            if path.endswith(".xml"):
                send_xml(conn)
            else:
                send_html(conn, '<html><body><a href="/b.xml">Batch</a></body></html>',
                          extra=[("Set-Cookie", "csrf=ABC123; Path=/")])

        # (a) z --cookie: obie wartości docierają na serwer
        server = self.serve(handler)
        opener = self.opener(server, cookie="SESSION=tajne")
        refs = scrape.find_xml_links(opener, self.auction(server.base + "auction/auk1"))
        scrape.download_xml(opener, refs[0], self.tmpdir())
        wyslane = dict(server.cookies)["/b.xml"]
        self.assertIn("SESSION=tajne", wyslane)
        self.assertIn("csrf=ABC123", wyslane, "ciasteczko serwera musi wrócić")

        # (b) bez --cookie: ten sam serwer, ciasteczko wraca poprawnie
        server2 = self.serve(handler)
        opener2 = self.opener(server2)
        refs2 = scrape.find_xml_links(opener2, self.auction(server2.base + "auction/auk1"))
        scrape.download_xml(opener2, refs2[0], self.tmpdir())
        self.assertEqual(dict(server2.cookies)["/b.xml"], "csrf=ABC123")

    def test_ok_403_konczy_sie_bledem_bez_ponawiania(self):
        """403 (brak uprawnień) nie jest ponawiane — nie ma sensu dobijać portalu."""
        def handler(srv, conn, path, req):
            send_raw(conn, "403 Forbidden", [("Content-Length", "0")])

        server = self.serve(handler)
        opener = self.opener(server)
        with self.assertRaises(scrape.FetchError) as ctx:
            scrape.fetch(opener, server.base + "b.xml")
        self.assertEqual(ctx.exception.status, 403)
        self.assertEqual(len(server.hits), 1)

    def test_ok_strona_logowania_zamiast_listy_daje_pusta_liste(self):
        """Cała witryna za logowaniem: zero aukcji, ZERO wyjątków."""
        login = ("<html><head><title>Zaloguj się</title></head><body>"
                 "<form action='/login' method='post'><input name='u'></form></body></html>")

        def handler(srv, conn, path, req):
            send_html(conn, login)

        server = self.serve(handler)
        err = self.quiet()
        opener = self.opener(server)
        found = scrape.discover_auctions(opener, server.base)
        self.assertEqual(found, [])
        self.assertIn("Nic nie znaleziono", err.getvalue())
        # diagnostyka potrafi powiedzieć, CO tam było
        report = scrape.diagnose_last(opener)
        self.assertIn("Zaloguj", report)
        self.assertIn("formularzy: 1", report)


if __name__ == "__main__":
    unittest.main(verbosity=2)
