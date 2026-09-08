# Ustalenia o portalu flexitauctions.com

Zebrane WYŁĄCZNIE z wyników wyszukiwarki — sesja nie miała dostępu sieciowego do portalu
(egress proxy blokuje domenę, potwierdzone: CONNECT -> 403). Nic nie zostało zweryfikowane
na żywym serwisie. Traktuj jako wskazówki do domyślnych wzorców, nie jako pewnik.

## Zaobserwowane adresy URL

- Strona główna / lista aukcji: `https://flexitauctions.com/`  (tytuł: "Online Auctions")
- Strona aukcji:        `https://flexitauctions.com/auction/flexit-auctions-18-06-2026-1103`
- Zakładka info aukcji: `https://flexitauctions.com/auction/flexit-auctions-26-02-2026-1087/info`
- Lot wewnątrz aukcji:  `https://flexitauctions.com/auction/<slug-aukcji>-<id>/<slug-lotu>-<hash5>`
- Lot samodzielnie:     `https://flexitauctions.com/lot/<slug-lotu>-<hash5>`
- Subdomena mediów:     `https://media.flexitauctions.com/`  (prawdopodobne źródło plików do pobrania)

Wzorce:
- slug aukcji: `flexit-auctions-DD-MM-YYYY-<numer>` (numery rosnące: 1023, 1037, 1087, 1088, 1097, 1098, 1103)
- hash lotu: 5 znaków hex, np. `e0c8f`, `cd048`, `bcd12`, `53439`, `2a98f`, `6bfb2`, `11588`

## Parametry zapytań widziane w indeksie

- `?custom-category=(Mobile)Workstations`
- `?order=closedSoonest`
- `?lotState=ViewingOnly&lotState=OpenNoBids&lotState=OpenReserveNotMet&lotState=OpenSelling&lotState=Sold`
- `?lotSlug=<slug-innego-lotu>`   (parametr nawigacyjny między lotami)

## NAJWAŻNIEJSZE: gdzie jest XML

Snippet wyszukiwarki dla stron lotów wskazuje na przycisk **"Download Batch Details"**.
To jest najpewniej właśnie "plik XML z zawartością pakietu", o którym mówi użytkownik:
lot = pakiet (np. "12x Lenovo 8th-10th Gen Laptop Mix"), a XML zawiera listę sztuk
w pakiecie (model, specyfikacja, numery seryjne, stan/grade).

Konsekwencja dla scrapera — hierarchia jest TRZYPOZIOMOWA:

    lista aukcji  ->  strona aukcji  ->  strony lotów  ->  "Download Batch Details" (XML)

Scraper NIE może zakładać, że XML jest linkowany bezpośrednio ze strony aukcji.
Musi zejść poziom niżej, do każdego lotu.

## Ryzyka do obsłużenia

1. **Portal może być aplikacją JS (SPA).** Wtedy linków nie ma w statycznym HTML.
   Scraper musi dodatkowo przeszukiwać osadzony JSON/JS: `__NEXT_DATA__`,
   `window.__INITIAL_STATE__`, tagi `<script type="application/json">`, oraz szukać
   w całej treści adresów pasujących do `\.xml` i do `media.flexitauctions.com`.
2. **"Download Batch Details" może być endpointem API**, np. `/api/lot/<id>/batch.xml`
   albo `?format=xml` / `?export=xml`. Wykrywanie musi iść po tekście linku/przycisku
   ("batch", "details", "download", "xml"), nie tylko po rozszerzeniu pliku.
3. **Może być wymagane logowanie** do pobrania szczegółów pakietu — stąd opcja `--cookie`.
4. **Waluta EUR, treść po angielsku** — nagłówki w XML będą angielskie.
5. Ten sam lot bywa dostępny pod dwoma URL-ami (`/lot/...` i `/auction/.../...`) — deduplikacja
   po hashu lotu jest konieczna, inaczej pobierzemy wszystko dwa razy.
6. Aukcje archiwalne sięgają co najmniej 2025 roku — potrzebny limit zakresu
   (`--limit`, filtr po roku/slugu), żeby nie ściągać całej historii przypadkiem.

## Wniosek dla użytkownika

Ponieważ portal jest nieosiągalny z tej sesji, część pobierająca jest napisana "w ciemno"
na podstawie powyższych wzorców i MUSI zostać zweryfikowana lokalnie.
Część scalająca XML -> XLSX jest niezależna od portalu i w pełni przetestowana.
