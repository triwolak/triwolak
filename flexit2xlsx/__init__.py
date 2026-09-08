"""flexit2xlsx — pobieranie XML-i z portalu aukcyjnego i scalanie ich w JEDEN plik XLSX.

Pakiet składa się z pięciu modułów:

* :mod:`flexit2xlsx.values`     — normalizacja i bezpieczeństwo pojedynczej komórki,
* :mod:`flexit2xlsx.xmlflatten` — generyczne spłaszczanie DOWOLNEGO XML-a do wierszy,
* :mod:`flexit2xlsx.xlsxwrite`  — zapis skoroszytu ``.xlsx`` (stdlib albo openpyxl),
* :mod:`flexit2xlsx.scrape`     — pobieranie stron i plików z portalu (tylko stdlib),
* :mod:`flexit2xlsx.cli`        — spinacz: podkomendy ``download``, ``build``, ``all``.

Uruchomienie z linii poleceń::

    python3 -m flexit2xlsx build --in katalog_z_xml --out aukcje.xlsx

Do działania wystarczy Python 3.9+ z biblioteką standardową.  ``openpyxl`` jest
opcjonalne — gdy jest zainstalowane, zostanie użyte automatycznie.

Moduły ładują się LENIWIE (przez ``__getattr__``), żeby ``import flexit2xlsx``
nie ciągnął za sobą całej sieciowej części, gdy potrzebne jest tylko scalanie.
"""

from __future__ import annotations

import importlib
from typing import Any

__version__ = "1.0.0"

__all__ = [
    "__version__",
    "values",
    "xmlflatten",
    "xlsxwrite",
    "scrape",
    "cli",
    "main",
]

#: Nazwy podmodułów ładowanych na żądanie.
_SUBMODULES = frozenset({"values", "xmlflatten", "xlsxwrite", "scrape", "cli"})


def __getattr__(name: str) -> Any:
    """Leniwy import podmodułów i funkcji :func:`flexit2xlsx.cli.main`."""
    if name in _SUBMODULES:
        module = importlib.import_module("." + name, __name__)
        globals()[name] = module
        return module
    if name == "main":
        function = importlib.import_module(".cli", __name__).main
        globals()["main"] = function
        return function
    raise AttributeError("moduł %r nie ma atrybutu %r" % (__name__, name))


def __dir__() -> list:
    """Uzupełnia podpowiedzi o leniwie ładowane nazwy."""
    return sorted(set(globals()) | set(__all__))
