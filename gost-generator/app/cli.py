"""gost — headless CLI для ИИ-агентов (ТЗ §19).

Ядро бизнес-логики — app.core; команды только оркестрируют его вызовы
(один источник истины, п. 19.1). Модуль НЕ импортирует графическую оболочку
и работает без дисплея. Контракт вывода (п. 19.4): stdout — только результат
(JSON-конверт при --json, иначе текст), stderr — логи/прогресс.
Коды возврата — п. 19.5; приоритет настроек — п. 19.6
(флаги > env > gost.json > состояние > дефолты).
Обработчики команд — в app/cli_cmds.py и app/cli_docs.py, адаптеры к ядру —
в app/cli_core.py, тексты артефактов — в app/cli_init.py.
"""

from __future__ import annotations

import argparse
import sys
from typing import Callable

from app import constants
from app import cli_cmds, cli_core, cli_docs, cli_init
from app.cli_core import CliError, make_error

# --- Парсер и точка входа ----------------------------------------------------------


def _add_json(p: argparse.ArgumentParser) -> None:
    """Флаг --json: единый конверт п. 19.4 в stdout."""
    p.add_argument("--json", action="store_true", help="печатать JSON-конверт в stdout")


def build_parser() -> argparse.ArgumentParser:
    """Дерево подкоманд ровно по п. 19.3 ТЗ."""
    parser = argparse.ArgumentParser(
        prog="gost",
        description=f"{constants.APP_TITLE} — headless CLI для агентов (ТЗ §19)")
    parser.add_argument("--version", action="version", version=f"%(prog)s {constants.APP_VERSION}")
    subs = parser.add_subparsers(dest="command", required=True)

    p = subs.add_parser("init", help="инициализация каталогов и артефактов агентов")
    p.add_argument("--agents", action="store_true",
                   help="записать AGENTS.md, schemas/*.json, examples/*.json")
    p.add_argument("--dir", metavar="PATH", help="целевой каталог (по умолчанию текущий)")
    _add_json(p)
    p.set_defaults(func=cli_cmds.cmd_init)

    p = subs.add_parser("doctor", help="самодиагностика окружения")
    _add_json(p)
    p.set_defaults(func=cli_cmds.cmd_doctor)

    p = subs.add_parser("rules", help="пресеты правил оформления")
    rsubs = p.add_subparsers(dest="rsub", required=True)
    for name in ("list",):
        q = rsubs.add_parser(name)
        _add_json(q)
        q.set_defaults(func=cli_cmds.cmd_rules)
    q = rsubs.add_parser("show"); q.add_argument("key"); _add_json(q); q.set_defaults(func=cli_cmds.cmd_rules)
    q = rsubs.add_parser("validate"); q.add_argument("key", nargs="?"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_rules)
    q = rsubs.add_parser("export"); q.add_argument("key"); q.add_argument("dest"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_rules)
    q = rsubs.add_parser("import"); q.add_argument("src"); q.add_argument("--key"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_rules)
    q = rsubs.add_parser("diff"); q.add_argument("left"); q.add_argument("right"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_rules)

    p = subs.add_parser("profile", help="профили автора и работы")
    prsubs = p.add_subparsers(dest="psub", required=True)
    q = prsubs.add_parser("list"); _add_json(q); q.set_defaults(func=cli_cmds.cmd_profile)
    q = prsubs.add_parser("show"); q.add_argument("key"); _add_json(q); q.set_defaults(func=cli_cmds.cmd_profile)
    q = prsubs.add_parser("set"); q.add_argument("key")
    q.add_argument("values", nargs="+", metavar="key=value"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_profile)
    q = prsubs.add_parser("use"); q.add_argument("key"); _add_json(q); q.set_defaults(func=cli_cmds.cmd_profile)
    q = prsubs.add_parser("export"); q.add_argument("key"); q.add_argument("dest"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_profile)
    q = prsubs.add_parser("import"); q.add_argument("src"); q.add_argument("--key"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_profile)

    p = subs.add_parser("scripts", help="скрипты-плагины")
    ssubs = p.add_subparsers(dest="ssub", required=True)
    q = ssubs.add_parser("list"); _add_json(q); q.set_defaults(func=cli_cmds.cmd_scripts)
    q = ssubs.add_parser("describe"); q.add_argument("name"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_scripts)
    q = ssubs.add_parser("add"); q.add_argument("path"); _add_json(q); q.set_defaults(func=cli_cmds.cmd_scripts)
    q = ssubs.add_parser("validate"); q.add_argument("path"); _add_json(q)
    q.set_defaults(func=cli_cmds.cmd_scripts)

    p = subs.add_parser("run", help="запуск скрипта-генератора")
    p.add_argument("script", metavar="SCRIPT")
    p.add_argument("--input", metavar="P")
    p.add_argument("--output", metavar="P")
    p.add_argument("--params", metavar="JSON|@file", help="параметры: JSON или @файл")
    p.add_argument("--rules", metavar="NAME|PATH")
    p.add_argument("--profile", metavar="NAME")
    p.add_argument("--normalize", action=argparse.BooleanOptionalAction, default=None,
                   help="применить правила после генерации (по умолчанию для rules_hook=postprocess)")
    p.add_argument("--check", action="store_true", help="проверить результат генерации")
    p.add_argument("--fail-on", choices=("error", "warning", "never"), default="error")
    p.add_argument("--dry-run", action="store_true", help="план запуска без исполнения и записей")
    _add_json(p)
    p.set_defaults(func=cli_docs.cmd_run)

    p = subs.add_parser("check", help="проверка документов по правилам")
    p.add_argument("paths", nargs="+", metavar="PATH")
    p.add_argument("--rules", metavar="NAME|PATH")
    p.add_argument("--strict", action=argparse.BooleanOptionalAction, default=None,
                   help="строгий режим (по умолчанию из пресета)")
    p.add_argument("--format", choices=("json", "text", "md"), default="text",
                   help="формат текстового результата (без --json)")
    p.add_argument("--report", metavar="P", help="сохранить отчёт в файл")
    p.add_argument("--fail-on", choices=("error", "warning", "never"), default="error")
    _add_json(p)
    p.set_defaults(func=cli_docs.cmd_check)

    p = subs.add_parser("fix", help="исправление оформления (нормализация, идемпотентно)")
    p.add_argument("paths", nargs="+", metavar="PATH")
    p.add_argument("--rules", metavar="NAME|PATH")
    p.add_argument("--dry-run", action="store_true", help="план действий без записи файлов")
    p.add_argument("--backup", action=argparse.BooleanOptionalAction, default=True,
                   help="создать бэкап <name>.orig.docx")
    _add_json(p)
    p.set_defaults(func=cli_docs.cmd_fix)

    p = subs.add_parser("batch", help="пакетная обработка каталога документов")
    p.add_argument("directory", metavar="DIR")
    p.add_argument("--glob", default="*.docx")
    p.add_argument("--check", action="store_true")
    p.add_argument("--fix", action="store_true")
    p.add_argument("--report", metavar="P")
    p.add_argument("--full", action="store_true", help="включать issues в элементы сводки")
    _add_json(p)
    p.set_defaults(func=cli_docs.cmd_batch)

    p = subs.add_parser("job", help="декларативная задача (п. 19.7)")
    jsubs = p.add_subparsers(dest="jsub", required=True)
    q = jsubs.add_parser("run")
    q.add_argument("task", metavar="task.json")
    _add_json(q)
    q.set_defaults(func=cli_docs.cmd_job)

    p = subs.add_parser("schema", help="JSON-схема артефакта")
    p.add_argument("name", choices=cli_init.SCHEMA_NAMES)
    _add_json(p)
    p.set_defaults(func=cli_cmds.cmd_schema)

    p = subs.add_parser("explain", help="расшифровка кода нарушения или ошибки CLI")
    p.add_argument("code")
    _add_json(p)
    p.set_defaults(func=cli_cmds.cmd_explain)

    p = subs.add_parser("context", help="контекст проекта для агента (первый вызов)")
    _add_json(p)
    p.set_defaults(func=cli_cmds.cmd_context)

    return parser


def main(argv: list[str] | None = None) -> int:
    """Точка входа CLI (п. 19.2). Возвращает код возврата из п. 19.5."""
    parser = build_parser()
    args = parser.parse_args(argv)
    command = args.command
    try:
        func: Callable[..., Result] = args.func
        result = func(args)
    except CliError as exc:
        return _fail(command, args, exc.code, exc.message, exc.path)
    except KeyboardInterrupt:  # pragma: no cover
        cli_core.log("прервано пользователем")
        return constants.EXIT_USAGE
    except SystemExit:
        raise
    except Exception as exc:  # noqa: BLE001 — traceback в stderr, в конверт только код (п. 12)
        import traceback
        traceback.print_exc(file=sys.stderr)
        return _fail(command, args, "INTERNAL_ERROR", f"Внутренняя ошибка: {exc}", None)

    for w in result.warnings:
        cli_core.log(f"предупреждение: {w.get('message') if isinstance(w, dict) else w}")
    rc = cli_core.exit_code_for(result.errors) if result.errors else constants.EXIT_OK
    if _wants_json(args):
        cli_core.emit_json(command, result.ok, result.data, result.errors, result.warnings)
    elif result.text:
        print(result.text)
    if rc != constants.EXIT_OK and not _wants_json(args):
        for e in result.errors:
            cli_core.log(f"ошибка[{e['code']}]: {e['message']}"
                         + (f" ({e['path']})" if e.get("path") else ""))
    return rc


def _wants_json(args: argparse.Namespace) -> bool:
    """Единый критерий JSON-вывода: явный `--json` или `--format json`.

    Иначе агент, вызвавший `gost check --format json`, получил бы «голый» data без
    конверта п. 19.4 и сломал бы парсинг.
    """
    return bool(getattr(args, "json", False)) or getattr(args, "format", None) == "json"


def _fail(command: str, args: argparse.Namespace, code: str,
          message: str, path: str | None) -> int:
    """Оформить ошибку как конверт (--json) или сообщение в stderr; вернуть код возврата."""
    errs = [make_error(code, message, path)]
    if _wants_json(args):
        cli_core.emit_json(command, False, {}, errs, [])
    else:
        cli_core.log(f"ошибка[{code}]: {message}" + (f" ({path})" if path else ""))
    return cli_core.exit_code_for(errs) or constants.EXIT_USAGE


if __name__ == "__main__":
    sys.exit(main())
