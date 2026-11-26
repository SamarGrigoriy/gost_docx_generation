"""Публичный API проверки `.docx` по активным правилам (контракт — INTERFACES.md).

Модуль тонкий: собирает модель документа (`docx_parser`), прогоняет слои проверок
(`docx_checks` — страница/текст/заголовки, `docx_content_checks` — рисунки/таблицы/
библиография) и отдаёт список `Issue`. Реальной логикой проверок владеют те модули,
здесь — только оркестрация и сериализация результата для CLI/GUI.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping, Sequence

from app import constants
from app.core.docx_checks import (
    FONT_ALIASES,
    HINTS,
    SEVERITY_BY_CODE,
    CheckContext,
    Issue,
)
from app.core.docx_checks import CHECKS_LAYOUT
from app.core.docx_content_checks import CHECKS_CONTENT, MANUAL_ITEMS
from app.core.docx_parser import DocxParser
from app.core.docx_xml import DocxError

__all__ = ["DocxError", "Issue", "check_docx", "check_docx_result", "summarize", "explain",
           "SEVERITY_BY_CODE", "HINTS"]

# Чек-лист «проверить вручную» (п. 9.1 ТЗ, код MANUAL).
_MANUAL_ITEMS: Sequence[tuple[str, str]] = MANUAL_ITEMS


def check_docx(path: str | Path, rules: Mapping[str, Any]) -> list[Issue]:
    """Проверить документ по правилам; вернуть замечания в порядке значимости.

    `rules` — словарь пресета (п. 7.1 ТЗ). Битый zip/XML → `DocxError` (вызывающий
    код перехватывает и показывает «Ошибка чтения», пакетный режим не падает).
    """
    model = DocxParser(mono_fonts=constants.MONOSPACE_FONTS).parse_file(str(path))
    ctx = CheckContext(model, rules)
    for check in (*CHECKS_LAYOUT, *CHECKS_CONTENT):
        check(ctx)
    return sort_issues(ctx.issues)


def sort_issues(issues: Sequence[Issue]) -> list[Issue]:
    """Сортировка: сначала ошибки, затем предупреждения/рекомендации, внутри — по абзацу."""
    return sorted(issues, key=lambda i: (
        constants.SEVERITY_ORDER.get(i.severity, 9),
        i.code,
        (i.location or {}).get("paragraph") if (i.location or {}).get("paragraph")
        is not None else 10 ** 9,
    ))


def summarize(issues: Sequence[Issue]) -> dict[str, int]:
    """Счётчики по severity для сводки отчёта и JSON-конверта."""
    out = {"error": 0, "warning": 0, "info": 0}
    for issue in issues:
        out[issue.severity] = out.get(issue.severity, 0) + 1
    return out


def check_docx_result(path: str | Path, rules: Mapping[str, Any]) -> dict[str, Any]:
    """Результат проверки в виде словаря (конверт п. 19.4 ТЗ).

    Ошибку чтения не поднимаем: `read_error` + пустой список замечаний — так
    пакетный режим («Проверка папки») показывает статус, а не падает.
    """
    file_path = Path(path)
    preset = str(rules.get("name") or rules.get("preset") or constants.DEFAULT_PRESET_KEY)
    strict = bool(rules.get("strict", True))
    base: dict[str, Any] = {
        "file": str(file_path),
        "preset": preset,
        "strict": strict,
        "summary": {"error": 0, "warning": 0, "info": 0},
        "issues": [],
        "read_error": None,
    }
    try:
        issues = check_docx(file_path, rules)
    except DocxError as exc:
        base["read_error"] = str(exc)
        return base
    except Exception as exc:  # noqa: BLE001 — проверка не должна ронять приложение
        base["read_error"] = f"Сбой проверки: {exc.__class__.__name__}: {exc}"
        return base
    base["issues"] = [i.to_dict() for i in issues]
    base["summary"] = summarize(issues)
    return base


def explain(code: str) -> dict[str, Any]:
    """Расшифровка кода нарушения (для `gost explain CODE`)."""
    return {
        "code": code,
        "severity": SEVERITY_BY_CODE.get(code, "unknown"),
        "hint": HINTS.get(code, "Код неизвестен: сверьтесь с таблицей проверок п. 9.1 ТЗ."),
        "manual": code in {"MANUAL", "VOLUME_FILL"},
    }


def font_equivalents(font: str) -> set[str]:
    """Метрические аналоги шрифта — нужны, чтобы LibreOffice-замена не считалась нарушением."""
    key = str(font).lower()
    return set(FONT_ALIASES.get(key, {key}))
