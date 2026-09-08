"""Punkt wejścia dla ``python3 -m flexit2xlsx``.

Cała logika siedzi w :mod:`flexit2xlsx.cli`; tutaj tylko przekazujemy kod
wyjścia do powłoki.
"""

from __future__ import annotations

import sys

from .cli import main

if __name__ == "__main__":  # pragma: no cover - uruchamiane przez interpreter
    sys.exit(main())
