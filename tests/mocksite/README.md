# Atrapa portalu flexitauctions.com

Statyczne strony używane przez `tests/test_scrape.py`. Portal nie jest osiągalny
z tej sesji (brak sieci), więc scraper testujemy na wiernej makiecie serwowanej
przez `http.server` na losowym wolnym porcie w wątku demona.

## Układ (taki sam jak w SITE_NOTES.md)

    lista aukcji (3 strony paginacji)
      -> strona aukcji  (/auction/<slug>, /auction/<slug>/info)
        -> strona lotu  (/lot/<slug>-<hash5>, /auction/<slug>/<slug>-<hash5>)
          -> plik XML "Download Batch Details"

## Co która strona sprawdza

| Plik | Rola w testach |
|---|---|
| `list_page1..3.html` | paginacja (`rel="next"`, `?page=N`, adresy względne), duplikat aukcji, link do obcej domeny, link do lotu (NIE aukcji) |
| `empty_list.html` | brak wyników — pusta lista i diagnostyka zamiast wyjątku |
| `auction_1087.html` | aukcja bez XML-a wprost: trzeba zejść do 6 lotów; ten sam lot pod dwoma adresami; lot z obcej domeny |
| `auction_1103.html` | XML linkowany wprost ze strony aukcji (`descend="auto"` nie schodzi do lotów) |
| `auction_pusta.html` | aukcja bez lotów i bez XML-a |
| `lot_e0c8f.html` | wariant 1: `href` kończący się `.xml` + ten sam link dwa razy (duplikat) |
| `lot_cd048.html` | wariant 2: `?format=xml` + ten sam adres w osadzonym JSON (`__NEXT_DATA__`) |
| `lot_53439.html` | wariant 3: atrybut `download="batch-53439.xml"`; obok `download` prowadzący do HTML-a (odrzucany po sondowaniu) |
| `lot_2a98f.html` | wariant 4: adres WZGLĘDNY bez rozszerzenia — XML rozpoznany dopiero po `Content-Type`; obok link do obcej domeny |
| `lot_11588.html` | path traversal w parametrze: `?file=../../../../etc/passwd&format=xml` |
| `lot_6bfb2.html` | lot bez żadnego XML-a |
| `login.html` | ściana logowania podana pod adresem `.xml` (błąd `NotXmlError`) |
| `not_xml.html`, `style.css` | treści do odrzucenia przy sondowaniu Content-Type |
| `batch_*.xml` | realistyczna "zawartość pakietu" (sprzęt, numery seryjne, grade) |
| `empty.xml` | pusta odpowiedź |

## `routes.json`

Mapa `cel żądania -> opis odpowiedzi`. Klucz to dosłowna ścieżka wraz z `?query`.
Pola: `file`, `content_type`, `status`, `location`, `headers`, `delay` (wolna
odpowiedź — test timeoutu), `fail_times` (ile pierwszych żądań ma zwrócić błąd,
`-1` = zawsze), `fail_status`, `retry_after`.

Trasy techniczne: `/flaky/batch.xml` (2x 503, potem 200), `/flaky-forever.xml`
(zawsze 503 — backoff 2/4/8/16 s), `/retry-after.xml` (nagłówek `Retry-After`),
`/teapot.xml` (418 — bez ponawiania), `/slow.xml` (timeout),
`/redirect-in` i `/redirect-out` (przekierowanie w witrynie i poza nią).
