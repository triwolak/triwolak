# Kontrakt interfejsów (WIĄŻĄCY — nie zmieniaj sygnatur)

Cel projektu: pobrać pliki XML z portalu aukcyjnego flexitauctions.com i scalić
zawartość WSZYSTKICH aukcji w JEDEN plik `.xlsx`.

## Zasady globalne

- **Zero wymaganych zależności zewnętrznych.** Tylko biblioteka standardowa Pythona 3.9+.
  `openpyxl` jest OPCJONALNE (lepsze formatowanie); gdy go brak, działa wbudowany zapis XLSX.
- Kodowanie wyjścia i komunikatów: UTF-8, polskie znaki muszą działać.
- Nie używaj `Date.now`-podobnej niedeterministyczności w testach.
- Każdy moduł ma testy w `tests/test_<modul>.py`, uruchamiane przez `python3 -m pytest`
  ORAZ przez `python3 -m unittest` (pisz testy w stylu `unittest`, żeby działały bez pytest).

## `flexit2xlsx/values.py` — normalizacja i bezpieczeństwo komórek

```python
def normalize_ws(text: str) -> str
    # zwija białe znaki, przycina; zachowuje polskie znaki

def coerce_value(text: str | None) -> Any
    # "" / None / same białe znaki -> None
    # liczby całkowite -> int   (ale NIE gdy wiodące zero: "007" zostaje str)
    # liczby dziesiętne -> float (obsłuż "1 234,56", "1.234,56", "1,234.56", "1234.56", spacja NBSP)
    # "true"/"false" (dowolna wielkość liter) -> bool ; "tak"/"nie" NIE są konwertowane
    # daty: "YYYY-MM-DD", "DD.MM.YYYY", "DD-MM-YYYY", ISO-8601 z czasem -> datetime.date / datetime.datetime
    # NIE konwertuj: numerów seryjnych, wersji ("1.2.3"), numerów telefonu, ciągów z jednostkami ("16 GB")
    # zwraca oryginalny str, gdy nic nie pasuje

def sanitize_cell(value: Any) -> Any
    # usuwa znaki niedozwolone w XLSX (control chars poza \t \n \r), przycina do 32767 znaków,
    # zamienia NaN/Inf na str, zwraca wartość gotową do zapisu
    # NIE dodaje apostrofu — ochrona przed formula injection jest po stronie zapisu (typ tekstowy)

def looks_like_formula(value: Any) -> bool
    # True dla str zaczynających się od = + - @ \t \r (po odcięciu białych znaków) — używane przez xlsxwrite
```

## `flexit2xlsx/xmlflatten.py` — generyczne spłaszczanie XML do wierszy

Schemat XML z portalu NIE jest znany. Kod musi działać z DOWOLNĄ strukturą.

```python
@dataclass
class ParsedDoc:
    source: str                    # ścieżka/nazwa pliku źródłowego
    auction: str                   # identyfikator aukcji (z XML albo z nazwy pliku)
    context: dict[str, Any]        # pola z poziomu aukcji (powtarzane w każdym wierszu)
    records: list[dict[str, Any]]  # jeden dict = jeden wiersz (pozycja/lot/pakiet)
    record_path: str | None        # wykryta ścieżka elementu powtarzalnego, np. "auction/lots/lot"
    columns: list[str]             # kolumny w kolejności pierwszego wystąpienia

def parse_bytes(data: bytes, source: str, *, strip_ns: bool = True,
                repeat: str = "join", join_sep: str = " | ",
                record_path: str | None = None) -> ParsedDoc
def parse_file(path, **kw) -> ParsedDoc
def detect_record_path(root) -> str | None
def merge_columns(docs: list[ParsedDoc]) -> list[str]
def rows_for(doc: ParsedDoc, columns: list[str]) -> Iterator[list[Any]]
```

Wymagania:
- Wykrywanie elementu powtarzalnego: ścieżka o największej liczbie wystąpień (>=2),
  ważona liczbą pól-liści; preferuj nazwy: item, lot, product, position, pozycja,
  asset, device, row, record, entry, artikel, part, komponent, sprzet.
  Gdy nic się nie powtarza — cały dokument to JEDEN wiersz.
- Klucz kolumny = ścieżka względna z `/`; atrybuty jako `sciezka@atrybut`.
- Powtarzające się dzieci wewnątrz rekordu: `repeat="join"` skleja `join_sep`,
  `repeat="index"` tworzy `pole[1]`, `pole[2]`.
- Pola przodków rekordu (metadane aukcji) trafiają do `context` i są powtarzane w każdym wierszu.
- Bezpieczeństwo: brak rozwijania encji zewnętrznych (XXE), odporność na bomby encyjne.
- Obsłuż: przestrzenie nazw, CDATA, deklaracje kodowania inne niż UTF-8 (np. ISO-8859-2,
  windows-1250), BOM, mieszane treści, puste elementy, atrybuty na elemencie rekordu.
- Uszkodzony XML: rzuć `XmlParseError` (podklasa `ValueError`) z czytelnym komunikatem.

Kolumny techniczne dodaje CLI, nie ten moduł.

## `flexit2xlsx/xlsxwrite.py` — zapis XLSX

```python
@dataclass
class Sheet:
    name: str
    columns: list[str]
    rows: Iterable[Sequence[Any]]

def backend_name() -> str            # "openpyxl" albo "stdlib"
def safe_sheet_name(name: str, used: set[str]) -> str   # <=31 znaków, bez []:*?/\, unikalna
def write_workbook(path, sheets: list[Sheet], *, freeze_header: bool = True,
                   autofilter: bool = True, auto_width: bool = True) -> None
```

Wymagania:
- Plik musi otwierać się w Excelu/LibreOffice — poprawny OOXML (`xl/workbook.xml`,
  `xl/worksheets/sheetN.xml`, `sharedStrings.xml` lub inline, `[Content_Types].xml`, `_rels`).
- Nagłówek pogrubiony, zamrożony wiersz 1, autofiltr, sensowne szerokości kolumn.
- Wartości tekstowe zaczynające się od `=` `+` `-` `@` zapisywane JAKO TEKST (nie formuła).
- Typy natywne: int/float jako liczby, date/datetime jako daty z formatem, bool jako PRAWDA/FAŁSZ.
- Limit Excela: 1 048 576 wierszy i 16 384 kolumn — przy przekroczeniu dziel na kolejne arkusze
  i ostrzegaj na stderr.
- `rows` może być generatorem (duże dane, bez ładowania wszystkiego do pamięci).

## `flexit2xlsx/scrape.py` — pobieranie z portalu (tylko stdlib: urllib + html.parser)

```python
@dataclass
class AuctionRef:
    url: str; id: str; title: str
@dataclass
class XmlRef:
    url: str; filename: str; auction_id: str; auction_title: str

def make_opener(*, user_agent: str = ..., cookie: str | None = None,
                timeout: float = 30.0, retries: int = 4) -> Any
def fetch(opener, url: str) -> tuple[bytes, str]        # (treść, content_type)
def discover_auctions(opener, base_url: str, *, auction_re: str | None = None,
                      max_pages: int = 50) -> list[AuctionRef]
def find_xml_links(opener, auction: AuctionRef, *, opener_fetch=None) -> list[XmlRef]
def download_xml(opener, ref: XmlRef, out_dir, *, overwrite: bool = False) -> str  # ścieżka pliku
```

Wymagania:
- Wykrywanie linków do aukcji: domyślny wzorzec `/auction/<slug>/` (zaobserwowany na portalu),
  plus paginacja (`?page=N`, `rel="next"`), plus konfigurowalny regex.
- Wykrywanie XML: `href` kończący się `.xml`, zawierający `xml` w ścieżce/parametrach,
  albo link, którego `Content-Type` to `application/xml` / `text/xml`.
- Ponawianie z wykładniczym backoffem (2s, 4s, 8s, 16s), uprzejme opóźnienie między żądaniami,
  `--dry-run` (tylko lista), pomijanie już pobranych plików, bezpieczne nazwy plików
  (bez path traversal, bez `..`, bez znaków spoza [\w.-]).
- Kod ma być ODPORNY na nieznany HTML: gdy nic nie znajdzie, zwróć pustą listę i pozwól CLI
  wypisać czytelną diagnostykę (nie rzucaj wyjątkiem).

## `flexit2xlsx/cli.py` — spinacz

Podkomendy:
- `download` — pobierz XML-e z portalu do katalogu
- `build`    — zbuduj JEDEN plik XLSX z katalogu/listy plików XML
- `all`      — download + build

Arkusze w wyniku:
1. `Wszystkie aukcje` — suma wszystkich wierszy ze wszystkich plików (unia kolumn)
2. `Podsumowanie` — plik, aukcja, liczba pozycji, wykryta ścieżka rekordu
3. opcjonalnie `--per-auction` — osobny arkusz na aukcję

Kolumny techniczne na początku arkusza zbiorczego: `Aukcja`, `Plik`, `Nr pozycji`.
