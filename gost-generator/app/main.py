"""Точка входа приложения: служебный режим ``--internal-run`` или запуск GUI (п. 3.2 ТЗ).

Зачем нужен ``--internal-run``: внутри PyInstaller-сборки ``sys.executable`` — это сам
exe, а не интерпретатор Python, поэтому напрямую запустить ``exe script.py`` нельзя.
exe сам выступает «интерпретатором»: читает JSON с готовым argv и выполняет скрипт через
``runpy.run_path``. В этой ветке НИКАКОГО импорта GUI (customtkinter не тянется, окно не
создаётся) — иначе собранный --windowed exe зависнет на Tk при фоновом запуске скрипта.
"""

from __future__ import annotations

import json
import os
import runpy
import sys
import traceback

INTERNAL_RUN_FLAG = "--internal-run"


def run_internal(script_path: str, params_json_path: str) -> int:
    """Выполнить скрипт-плагин от имени внутреннего интерпретатора.

    Формат params_json: ``{"argv": [...], "cwd": "..."|null}`` — его пишет runner.
    Выход (stdout) полностью совместим с прямым запуском python: все print-строки,
    traceback исключений тоже печатается в stdout (раннер объединяет stderr=STDOUT,
    но в frozen-режиме консоли нет, поэтому traceback направляем явно в stdout).

    :return: код возврата скрипта
    """
    try:
        with open(params_json_path, encoding="utf-8") as fh:
            payload = json.load(fh)
    except (OSError, ValueError) as exc:
        print(f"internal-run: не удалось прочитать файл параметров: {exc}")
        return 2

    if not os.path.isfile(script_path):
        print(f"internal-run: скрипт не найден: {script_path}")
        return 2

    argv = payload.get("argv") or []
    cwd = payload.get("cwd")
    sys.argv = [script_path, *[str(a) for a in argv]]
    if cwd:
        try:
            os.chdir(cwd)
        except OSError as exc:
            print(f"internal-run: не удалось сменить каталог {cwd}: {exc}")
            return 2

    try:
        runpy.run_path(script_path, run_name="__main__")
    except SystemExit as exc:
        # sys.exit(str) — сообщение в stderr поведений CPython; тут пишем в stdout.
        if isinstance(exc.code, str):
            print(exc.code)
            return 1
        return int(exc.code) if exc.code is not None else 0
    except BaseException:  # KeyboardInterrupt тоже считаем аварийным завершением скрипта
        traceback.print_exc(file=sys.stdout)
        return 1
    return 0


def _parse_internal_args(args: list[str]) -> tuple[int, list[str]]:
    """Выделить из аргументов команду ``--internal-run <script> <params_json>``.

    :return: (позиция флага или -1, остальные аргументы)
    """
    if args and args[0] == INTERNAL_RUN_FLAG:
        return 0, args[1:]
    return -1, args


def main(argv: list[str] | None = None) -> int:
    """Разобрать аргументы и выбрать режим; возвращает код для sys.exit."""
    from app.cli_core import configure_console  # только stdlib, GUI не тянется
    configure_console()
    args = list(sys.argv[1:] if argv is None else argv)
    pos, rest = _parse_internal_args(args)
    if pos >= 0:
        if len(rest) < 2:
            print(f"Использование: {INTERNAL_RUN_FLAG} <script_path> <params_json_path>")
            return 2
        return run_internal(rest[0], rest[1])

    # GUI ветка: ленивый импорт, чтобы --internal-run не тянул customtkinter.
    try:
        from app.ui import app_window  # импортируется ТОЛЬКО здесь
    except ImportError:
        traceback.print_exc(file=sys.stderr)
        print("GUI недоступен: установите зависимости (customtkinter) и повторите запуск.",
              file=sys.stderr)
        return 1

    from app.core import config

    try:
        dirs = config.DataDirs()          # порядок каталогов данных: п. 3.3/19.6
    except Exception as exc:              # битый gost.json/состояние — окно не открывается молча
        print(f"Не удалось определить каталоги данных: {exc}", file=sys.stderr)
        return 6
    app_window.run_gui(dirs)
    return 0


if __name__ == "__main__":
    import multiprocessing

    multiprocessing.freeze_support()  # обязательно до всего остального (п. 3.2.6)
    sys.exit(main())
