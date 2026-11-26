"""``python -m app`` — та же точка входа, что и exe (п. 3.1/3.2 ТЗ)."""

from __future__ import annotations

import multiprocessing
import sys

from app.main import main

if __name__ == "__main__":
    multiprocessing.freeze_support()  # PyInstaller onefile требует вызова в начале
    sys.exit(main())
