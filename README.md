# flexit2xlsx — wszystkie aukcje w jednym pliku Excela

Narzędzie pobiera pliki **XML "zawartość pakietu"** z portalu aukcyjnego
[flexitauctions.com](https://flexitauctions.com/) i scala je w **JEDEN plik `.xlsx`**,
który otworzysz w Excelu, LibreOffice albo Arkuszach Google.

Kluczowa cecha: **schemat XML nie musi być znany**. Program czyta dowolną strukturę
XML, sam znajduje w niej powtarzający się element (pozycję / lot / sztukę sprzętu),
zamienia go na wiersze, a resztę pól (dane aukcji, sprzedawcy, pakietu) powtarza
w każdym wierszu. Pliki o różnych schematach lądują w jednej tabeli — brakujące
pola są po prostu puste.

> ⚠️ **WAŻNE OSTRZEŻENIE.** Część **pobierająca** (`download`) powstała bez dostępu
> do sieci — sesja, w której pisano ten kod, miała zablokowane połączenie z portalem.
> Wzorce adresów pochodzą wyłącznie z wyników wyszukiwarki i **nie zostały sprawdzone
> na żywym serwisie**. Może się zdarzyć, że portal ma inny układ stron i pobieranie
> nie znajdzie nic — patrz rozdział
> [„Gdy skrypt nie znajduje aukcji lub XML-i”](#gdy-skrypt-nie-znajduje-aukcji-lub-xml-i).
> Część **scalająca** (`build`) jest w pełni przetestowana i działa niezależnie od portalu:
> zawsze możesz pobrać pliki ręcznie i użyć samego `build`.

---

## Spis treści

1. [Wymagania](#wymagania)
2. [Instrukcja krok po kroku — Windows](#instrukcja-krok-po-kroku--windows)
3. [Instrukcja krok po kroku — macOS i Linux](#instrukcja-krok-po-kroku--macos-i-linux)
4. [Przykłady poleceń](#przykłady-poleceń)
5. [Co znajdziesz w pliku wynikowym](#co-znajdziesz-w-pliku-wynikowym)
6. [Wszystkie opcje](#wszystkie-opcje)
7. [Gdy skrypt nie znajduje aukcji lub XML-i](#gdy-skrypt-nie-znajduje-aukcji-lub-xml-i)
8. [Logowanie i ciasteczka (`--cookie`)](#logowanie-i-ciasteczka---cookie)
9. [Kody wyjścia](#kody-wyjścia)
10. [Najczęstsze problemy](#najczęstsze-problemy)
11. [Dla programistów](#dla-programistów)
12. [Znane ograniczenia](#znane-ograniczenia)

---

## Wymagania

| Co | Wersja | Uwagi |
|---|---|---|
| Python | **3.9 lub nowszy** | na Windows pobierz z [python.org](https://www.python.org/downloads/) |
| Biblioteki zewnętrzne | **żadne** | wszystko działa na bibliotece standardowej |
| `openpyxl` | opcjonalnie | gdy jest zainstalowane, zostanie użyte automatycznie; gdy go nie ma, program zapisuje XLSX własnym kodem |
| Dostęp do internetu | tylko dla `download` | podkomenda `build` działa całkowicie offline |

Nic nie trzeba instalować — wystarczy skopiować katalog `flexit2xlsx` i uruchomić
go poleceniem `python3 -m flexit2xlsx`.

Sprawdzenie, czy Python jest zainstalowany:

```bash
python3 --version        # macOS / Linux
py --version             # Windows
```

---

## Instrukcja krok po kroku — Windows

1. **Zainstaluj Pythona.** Wejdź na <https://www.python.org/downloads/>, kliknij
   żółty przycisk *Download Python*. W instalatorze **koniecznie zaznacz
   „Add python.exe to PATH”**, potem *Install Now*.
2. **Rozpakuj narzędzie.** Załóżmy, że katalog z projektem (ten, w którym leży
   plik `README.md` i katalog `flexit2xlsx`) trafił na pulpit:
   `C:\Users\TwojaNazwa\Desktop\triwolak`.
3. **Otwórz wiersz poleceń.** Naciśnij `Win`, wpisz `cmd`, `Enter`.
4. **Przejdź do katalogu projektu** (uwaga na własną nazwę użytkownika):

   ```bat
   cd C:\Users\TwojaNazwa\Desktop\triwolak
   ```

5. **Sprawdź, czy działa:**

   ```bat
   py -m flexit2xlsx --help
   ```

   Powinna wypisać się pomoc po polsku. Jeśli pojawia się „py nie jest rozpoznawane
   jako polecenie” — Python nie został dodany do PATH; zainstaluj go ponownie
   z zaznaczoną opcją z punktu 1.
6. **Pobierz pliki i zbuduj arkusz — jednym poleceniem:**

   ```bat
   py -m flexit2xlsx all --in xml_flexit --out aukcje.xlsx
   ```

   * `--in xml_flexit` — katalog, w którym wylądują pobrane pliki XML (utworzy się sam),
   * `--out aukcje.xlsx` — plik wynikowy.
7. **Otwórz `aukcje.xlsx`** — leży w katalogu projektu, kliknij dwa razy.

**Jeśli krok 6 nic nie pobiera** (a to możliwe, patrz ostrzeżenie na górze), pobierz
pliki XML ręcznie z przeglądarki, wrzuć je do katalogu `xml_flexit` i uruchom samo
scalanie:

```bat
py -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx
```

---

## Instrukcja krok po kroku — macOS i Linux

1. **Sprawdź Pythona.** W Terminalu:

   ```bash
   python3 --version
   ```

   Jeśli wypisze wersję 3.9 lub wyższą — jest dobrze. Na macOS bez Pythona:
   zainstaluj z <https://www.python.org/downloads/> albo przez `brew install python`.
   Na Ubuntu/Debianie: `sudo apt install python3`.
2. **Przejdź do katalogu projektu:**

   ```bash
   cd ~/Desktop/triwolak
   ```

3. **Sprawdź, czy działa:**

   ```bash
   python3 -m flexit2xlsx --help
   ```

4. **Pobierz i zbuduj arkusz:**

   ```bash
   python3 -m flexit2xlsx all --in xml_flexit --out aukcje.xlsx
   ```

5. **Otwórz wynik:**

   ```bash
   open aukcje.xlsx        # macOS
   xdg-open aukcje.xlsx    # Linux
   ```

Gdy pobieranie nic nie znajdzie — pobierz XML-e ręcznie i użyj samego `build`
(jak w instrukcji dla Windows).

---

## Przykłady poleceń

W przykładach używamy `python3`; na Windows zamień je na `py`.

```bash
# 1. NAJPROSTSZE: mam już katalog z plikami XML, chcę jeden Excel
python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx

# 2. Wskazanie pojedynczych plików zamiast katalogu
python3 -m flexit2xlsx build --in pakiet1.xml pakiet2.xml --out aukcje.xlsx

# 3. Dodatkowo osobna zakładka dla każdej aukcji
python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx --per-auction

# 4. Nadpisanie istniejącego wyniku (bez tego program odmówi)
python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx --overwrite

# 5. Podgląd bez zapisu: co by się znalazło w arkuszu?
python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx --dry-run

# 6. Wszystko jako tekst (gdy Excel psuje numery seryjne albo kody)
python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx --no-typing

# 7. Powtarzające się pola w osobnych kolumnach: cecha[1], cecha[2]…
python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx --repeat index

# 8. Wymuszenie, co jest "wierszem" w XML-u (gdy automat wybrał źle)
python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx --record-path item

# 9. POBIERANIE: tylko lista, nic nie ściągaj (test na sucho)
python3 -m flexit2xlsx download --out xml_flexit --dry-run

# 10. Pobranie 3 pierwszych aukcji — na próbę, żeby nie ściągać całej historii
python3 -m flexit2xlsx download --out xml_flexit --limit 3

# 11. Tylko aukcje z 2026 roku
python3 -m flexit2xlsx download --out xml_flexit --match 2026

# 12. Portal wymaga zalogowania — ciasteczko z przeglądarki
python3 -m flexit2xlsx download --out xml_flexit --cookie "sessionid=abc123; csrftoken=xyz"

# 13. Pobieranie + scalanie za jednym zamachem, ciszej i z podziałem na aukcje
python3 -m flexit2xlsx all --in xml_flexit --out aukcje.xlsx --per-auction --quiet

# 14. Portal nie odpowiada tak, jak się spodziewamy — pokaż, co widać na stronie
python3 -m flexit2xlsx download --out xml_flexit --diagnose
```

---

## Co znajdziesz w pliku wynikowym

### Arkusz `Wszystkie aukcje`

Wszystkie wiersze ze wszystkich plików, jedna wspólna tabela. Na początku trzy
kolumny techniczne dodawane przez program:

| Kolumna | Znaczenie |
|---|---|
| `Aukcja` | identyfikator aukcji odczytany z XML-a, a gdy go nie ma — nazwa pliku |
| `Plik` | nazwa pliku XML, z którego pochodzi wiersz |
| `Nr pozycji` | numer pozycji w obrębie jednego pliku (1, 2, 3…) |

Dalej idą kolumny z XML-a. Nazwa kolumny to ścieżka w dokumencie
(`lot/items/item/model` → `model`, atrybut → `element@atrybut`). Pola, których
dany plik nie ma, zostają **puste** — dzięki temu pliki o różnych schematach
mieszczą się w jednej tabeli.

Dodatkowo: nagłówek jest pogrubiony i zamrożony, włączony jest autofiltr,
a szerokości kolumn są dobrane automatycznie. Liczby, daty i wartości
logiczne zapisywane są jako **prawdziwe typy Excela**, a nie tekst.

### Arkusz `Podsumowanie`

Po jednym wierszu na każdy plik: nazwa pliku, aukcja, liczba pozycji, wykryta
ścieżka rekordu, liczba kolumn i status. Pliki, których nie udało się wczytać,
też tu są — ze statusem `BŁĄD: …`. Ostatni wiersz `RAZEM` podaje sumy.

### Arkusze per aukcja (opcja `--per-auction`)

Osobna zakładka dla każdej aukcji, z kolumnami technicznymi `Plik` i `Nr pozycji`.
Nazwy zakładek są przycinane do 31 znaków (limit Excela) i odróżniane
przyrostkiem `(2)`, `(3)`… gdy się powtarzają.

---

## Wszystkie opcje

### Wspólne

| Opcja | Działanie |
|---|---|
| `--quiet`, `-q` | mniej komunikatów (błędy nadal widać) |
| `--dry-run` | pokaż, co by się stało, ale niczego nie zapisuj |
| `--overwrite` | nadpisz istniejące pliki (XML-e przy pobieraniu, `.xlsx` przy budowaniu) |
| `--version` | numer wersji |
| `--help` | pomoc (działa też dla każdej podkomendy: `build --help`) |

### `download` — pobieranie z portalu

| Opcja | Działanie |
|---|---|
| `--out KATALOG` | katalog na pobrane pliki XML (domyślnie `xml_flexit`) |
| `--base-url ADRES` | adres listy aukcji (domyślnie `https://flexitauctions.com/`) |
| `--cookie CIASTECZKO` | nagłówek `Cookie` skopiowany z przeglądarki (logowanie) |
| `--auction-re REGEX` | własny wzorzec adresów aukcji |
| `--match REGEX` | pobierz tylko aukcje, których identyfikator lub tytuł pasuje do wzorca |
| `--limit N` | pobierz najwyżej N aukcji |
| `--max-pages N` | ile stron listy przejrzeć (domyślnie 50) |
| `--max-lots N` | ile lotów w jednej aukcji odwiedzić (domyślnie 200) |
| `--descend auto\|always\|never` | czy schodzić ze strony aukcji na strony lotów |
| `--delay SEKUNDY` | uprzejma przerwa między żądaniami (domyślnie 0.5 s) |
| `--timeout`, `--retries` | limit czasu i liczba ponowień (domyślnie 30 s, 4 ponowienia) |
| `--user-agent TEKST` | własny nagłówek `User-Agent` |
| `--diagnose` | gdy nic nie znaleziono, wypisz, co właściwie jest na stronie |

### `build` — scalanie do XLSX

| Opcja | Działanie |
|---|---|
| `--in SCIEZKA…` | katalog z XML-ami albo pojedyncze pliki (można powtarzać) |
| `--out PLIK` | plik wynikowy `.xlsx` (domyślnie `aukcje.xlsx`) |
| `--per-auction` | dodatkowy arkusz na każdą aukcję |
| `--repeat join\|index` | powtarzające się pola: sklej w jedną komórkę (`join`) albo rozbij na `pole[1]`, `pole[2]` (`index`) |
| `--join-sep TEKST` | czym sklejać przy `--repeat join` (domyślnie ` \| `) |
| `--record-path SCIEZKA` | wymuś, który element XML-a jest „wierszem” (np. `item` albo `batch/lot/items/item`) |
| `--keep-ns` | zostaw przestrzenie nazw w nazwach kolumn |
| `--no-typing` | nie zamieniaj tekstu na liczby/daty — wszystko jako tekst |
| `--no-unify` | nie ujednolicaj ścieżki rekordu między plikami (patrz niżej) |
| `--limit N` | weź najwyżej N plików XML |

**O `--no-unify`.** Plik z JEDNĄ pozycją nie ma nic powtarzalnego, więc sam z siebie
zostałby potraktowany jako jeden wiersz całego dokumentu — i jego kolumny nazywałyby
się inaczej niż w plikach z wieloma pozycjami. Domyślnie program to naprawia:
bierze ścieżkę rekordu wykrytą w pozostałych plikach i wczytuje taki plik ponownie,
żeby kolumny się zgadzały. `--no-unify` wyłącza to zachowanie.

### `all` — pobierz i zbuduj

Przyjmuje wszystkie opcje obu powyższych. Uwaga na znaczenie ścieżek:
`--in` to **katalog na pobrane XML-e**, a `--out` to **plik `.xlsx`**.

---

## Gdy skrypt nie znajduje aukcji lub XML-i

To najbardziej prawdopodobny problem — portal nie był i nie mógł być sprawdzony
na żywo. Kolejność działań:

### Krok 1. Zobacz, co program w ogóle widzi

```bash
python3 -m flexit2xlsx download --out xml_flexit --dry-run --diagnose
```

`--diagnose` wypisze tytuł strony, liczbę linków, przykładowe adresy i ostrzeżenie,
gdy strona nie ma ŻADNYCH linków (to znak, że portal jest aplikacją JavaScript
albo wymaga zalogowania).

### Krok 2. Podaj własny wzorzec adresu aukcji

Domyślnie program szuka adresów pasujących do `/auction/<coś>`. Jeżeli na portalu
adresy wyglądają inaczej, otwórz stronę listy aukcji w przeglądarce, kliknij prawym
przyciskiem na link do aukcji → *Kopiuj adres linku* i zobacz, jak jest zbudowany.
Potem podaj własne wyrażenie regularne — z grupą `(?P<id>…)`, która wskazuje
identyfikator aukcji:

```bash
# adresy typu https://flexitauctions.com/auctions/1234-nazwa
python3 -m flexit2xlsx download --out xml_flexit \
    --auction-re "/auctions/(?P<id>[0-9]+-[^/?#]+)"

# adresy typu https://flexitauctions.com/pl/aukcja/abc123
python3 -m flexit2xlsx download --out xml_flexit \
    --auction-re "/pl/aukcja/(?P<id>[^/?#]+)"
```

Ściągawka z wyrażeń regularnych:

| Zapis | Znaczenie |
|---|---|
| `[^/?#]+` | jeden lub więcej znaków, ale nie `/`, `?` ani `#` |
| `[0-9]+` | jedna lub więcej cyfr |
| `(?P<id>…)` | ta część adresu zostanie użyta jako identyfikator aukcji |
| `$` | koniec adresu |

### Krok 3. Zmuś program, żeby zszedł do lotów

XML „Download Batch Details” bywa dopiero na stronie pojedynczego lotu, nie na
stronie aukcji:

```bash
python3 -m flexit2xlsx download --out xml_flexit --descend always --limit 1
```

`--limit 1` ogranicza próbę do jednej aukcji, żeby nie zamęczyć portalu.

### Krok 4. Pobierz pliki ręcznie i użyj tylko `build`

**To działa zawsze** i jest niezależne od portalu:

1. Otwórz w przeglądarce stronę aukcji, wejdź w kolejne loty.
2. Kliknij przycisk **„Download Batch Details”** (albo inny odnośnik do XML-a).
3. Zapisz wszystkie pliki `.xml` do jednego katalogu, np. `xml_flexit`.
   Nazwy plików mogą być dowolne — program i tak wypisze je w kolumnie `Plik`.
4. Uruchom scalanie:

   ```bash
   python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx
   ```

### Krok 5. Program pobrał pliki, ale arkusz wygląda dziwnie

* **Wszystko wylądowało w jednym wierszu na plik** — automat nie znalazł
  powtarzającego się elementu. Otwórz XML w Notatniku, znajdź znacznik, który się
  powtarza (np. `<item>`, `<unit>`, `<line>`) i podaj go wprost:

  ```bash
  python3 -m flexit2xlsx build --in xml_flexit --out aukcje.xlsx --record-path item
  ```

* **Kolumn jest podejrzanie dużo** — automat wybrał zbyt „płytki” element.
  Zobacz w arkuszu `Podsumowanie`, kolumna `Ścieżka rekordu`, jaką ścieżkę wykryto,
  i wskaż inną przez `--record-path`.
* **Excel psuje numery seryjne albo kody** (np. `007` → `7`, długie numery →
  notacja wykładnicza) — użyj `--no-typing`.

---

## Logowanie i ciasteczka (`--cookie`)

Portal może wymagać zalogowania, żeby wydać plik z zawartością pakietu. Program
**nie umie się logować sam** (nie wysyła formularzy), ale potrafi korzystać
z sesji, którą już otworzyłeś w przeglądarce.

### Jak skopiować ciasteczko — Chrome / Edge

1. Zaloguj się na portalu w przeglądarce.
2. Naciśnij `F12` (narzędzia deweloperskie) → zakładka **Network** / **Sieć**.
3. Odśwież stronę (`F5`) i kliknij pierwszy wpis na liście.
4. W panelu po prawej znajdź sekcję **Request Headers** → wiersz **`Cookie:`**.
5. Skopiuj **całą** wartość po dwukropku (bywa długa, to normalne).

### Jak skopiować ciasteczko — Firefox

To samo, zakładka **Sieć**, sekcja **Nagłówki żądania**, wiersz `Cookie`.

### Użycie

```bash
python3 -m flexit2xlsx download --out xml_flexit \
    --cookie "sessionid=abc123; csrftoken=xyz789; other=1"
```

Uwagi bezpieczeństwa:

* Ciasteczko to **klucz do Twojego konta** — nie wklejaj go na forach ani do zgłoszeń
  błędów, nie zapisuj w plikach współdzielonych.
* Program wysyła ciasteczko **wyłącznie do tej samej witryny**, z której pobiera dane;
  przekierowanie na obcą domenę jest przerywane, żeby ciasteczko nie wyciekło.
* Sesja wygasa — po kilku godzinach trzeba skopiować ciasteczko na nowo.
* Objaw wygasłej sesji: komunikat *„Treść spod … nie jest XML-em … Portal mógł
  zwrócić stronę logowania”*.

---

## Kody wyjścia

Przydatne, gdy uruchamiasz narzędzie z innego skryptu:

| Kod | Znaczenie |
|---|---|
| `0` | wszystko się udało |
| `1` | błąd ogólny (np. plik wynikowy już istnieje i nie podano `--overwrite`) |
| `2` | błąd składni polecenia |
| `3` | zrobione częściowo — plik powstał, ale część XML-i się nie wczytała |
| `4` | nie znaleziono danych (żadnej aukcji, żadnego XML-a, żadnego pliku) |
| `5` | błąd komunikacji z portalem (sieć, HTTP, blokada) |
| `130` | przerwane przez użytkownika (Ctrl+C) |

---

## Najczęstsze problemy

| Objaw | Co zrobić |
|---|---|
| `python3: command not found` / `py nie jest rozpoznawane` | Python nie jest zainstalowany albo nie ma go w PATH — patrz instrukcje wyżej |
| `No module named flexit2xlsx` | jesteś w złym katalogu; wejdź (`cd`) do katalogu, w którym leży podkatalog `flexit2xlsx` |
| `Plik aukcje.xlsx już istnieje` | dodaj `--overwrite` albo podaj inną nazwę w `--out` |
| `Nie znalazłem żadnego pliku .xml` | sprawdź ścieżkę po `--in`; pliki muszą mieć rozszerzenie `.xml` |
| `Nie znalazłem żadnej aukcji` | patrz [rozdział o braku aukcji](#gdy-skrypt-nie-znajduje-aukcji-lub-xml-i) |
| `Treść spod … nie jest XML-em` | portal zwrócił stronę logowania — użyj `--cookie` |
| Excel: „plik jest uszkodzony” | zgłoś to jako błąd; spróbuj też `--no-typing` i innej nazwy pliku wyjściowego |
| Polskie znaki wyglądają źle | otwórz plik `.xlsx` (nie CSV) w Excelu — kodowanie jest w środku i zawsze jest UTF-8 |
| Bardzo dużo wierszy | powyżej 1 048 576 wierszy program sam dzieli dane na kolejne arkusze i ostrzega o tym |

---

## Dla programistów

### Struktura

```
flexit2xlsx/
    values.py      normalizacja i bezpieczeństwo pojedynczej komórki
    xmlflatten.py  generyczne spłaszczanie DOWOLNEGO XML-a do wierszy
    xlsxwrite.py   zapis XLSX: wbudowany generator OOXML albo openpyxl
    scrape.py      pobieranie z portalu (urllib + html.parser, zero zależności)
    cli.py         podkomendy download / build / all
tests/             testy jednostkowe i end-to-end (unittest)
INTERFACES.md      wiążący kontrakt interfejsów między modułami
SITE_NOTES.md      co wiadomo o portalu (i skąd)
```

### Testy

```bash
python3 -m unittest discover -s tests           # cały zestaw
python3 -m unittest discover -s tests -v        # z nazwami testów
python3 -m unittest tests.test_cli              # jeden moduł
```

Testy używają `openpyxl` i `lxml` do **niezależnej weryfikacji** wyników — sam
program ich nie potrzebuje. Część sieciowa jest testowana przeciwko atrapie
portalu (`tests/mocksite`, serwer `http.server` na losowym porcie).

### Wymuszenie sposobu zapisu XLSX

```bash
FLEXIT_XLSX_BACKEND=stdlib   python3 -m flexit2xlsx build --in xml_flexit --out a.xlsx
FLEXIT_XLSX_BACKEND=openpyxl python3 -m flexit2xlsx build --in xml_flexit --out a.xlsx
```

Domyślnie (`auto`) używane jest `openpyxl`, gdy jest zainstalowane.

### Użycie jako biblioteki

```python
from flexit2xlsx import cli, xmlflatten, xlsxwrite

docs, errors = cli.load_documents(cli.collect_xml_files(["xml_flexit"]))
xlsxwrite.write_workbook("aukcje.xlsx", cli.build_sheets(docs, errors))
```

---

## Znane ograniczenia

* **Pobieranie nie było sprawdzone na żywym portalu** — patrz ostrzeżenie na początku.
* Program **nie loguje się sam** (brak wysyłania formularzy, obsługi CSRF i 2FA);
  jedyna droga to `--cookie`.
* **Nie wykonuje JavaScriptu.** Jeżeli portal buduje listę aukcji dopiero w przeglądarce,
  scraper przeszuka jeszcze osadzony JSON (`__NEXT_DATA__` i podobne), ale nie zawsze
  to wystarczy.
* **Brak równoległego pobierania** — świadomie, żeby nie obciążać portalu.
* Jeden dokument XML daje **jedną tabelę**; gdy plik zawiera dwie niezależne listy,
  wybrana zostanie ta, która wygląda na główną.
* Rozpoznawanie typów jest **ostrożne**: `"16 GB"`, `"1.2.3"`, `"007"` i numery
  telefonów zostają tekstem. Wartości takie jak `"1,234"` traktowane są po polsku
  (jeden przecinek = separator dziesiętny).
* Otwarcie wyniku w Excelu i LibreOffice **należy potwierdzić lokalnie** — środowisko,
  w którym powstał kod, nie miało zainstalowanego arkusza kalkulacyjnego. Poprawność
  formatu OOXML jest sprawdzana w testach na poziomie struktury pliku i przez `openpyxl`.
