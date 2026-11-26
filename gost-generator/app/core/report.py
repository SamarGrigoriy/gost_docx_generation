"""Отчёты проверки DOCX: markdown, текст, пакетная сводка (п. 9.2, 10.3 ТЗ).

Модуль НЕ знает про UI и не читает файлы DOCX — он рендерит результат работы
`docx_inspector.check_docx_result` (или такой же dict, пришедший из JSON-конверта CLI),
поэтому пригоден и для GUI, и для CLI, и для пост-хука раннера.

Структура отчёта фиксирована:
шапка (файл, дата, пресет, режим) → сводка счётчиков → таблица замечаний
(Severity | Код | Замечание | Где | Как исправить) → блок «Проверить вручную».
"""

from __future__ import annotations

import datetime as _dt
import os
import textwrap
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from app import constants
from app.core.docx_inspector import HINTS

__all__ = ["render_markdown", "render_text", "save_report", "render_batch_markdown",
           "render_batch_text", "save_batch_report", "export_pdf", "format_summary",
           "MANUAL_CODE", "SEVERITY_LABELS", "SEVERITY_SHORT"]

SEVERITY_LABELS: dict[str, str] = {
    "error": "Ошибка", "warning": "Предупреждение", "info": "Рекомендация",
}
# Сокращения для узкой таблицы в терминале/GUI.
SEVERITY_SHORT: dict[str, str] = {"error": "ERR", "warning": "WARN", "info": "INFO"}
MANUAL_CODE = "MANUAL"
MD_FORMATS = ("md", "markdown")
TXT_FORMATS = ("txt", "text")


# --- Адаптеры ---------------------------------------------------------------
def _hint_for(code: str) -> str:
    """Подсказка «как исправить» по коду (для замечаний, где проверка её не передала)."""
    return HINTS.get(str(code), "Код неизвестен: сверьтесь с таблицей проверок п. 9.1 ТЗ.")


def _as_dict(issue: Any) -> dict[str, Any]:
    """Issue (dataclass) или словарь из JSON-конверта → единый dict."""
    if isinstance(issue, Mapping):
        return dict(issue)
    return {"code": getattr(issue, "code", ""), "severity": getattr(issue, "severity", ""),
            "message": getattr(issue, "message", ""),
            "location": dict(getattr(issue, "location", {}) or {}),
            "hint": getattr(issue, "hint", "")}


def _issues(result: Mapping[str, Any]) -> list[dict[str, Any]]:
    return [_as_dict(i) for i in (result.get("issues") or [])]


def _where(issue: Mapping[str, Any]) -> str:
    """Колонка «Где»: страница (эвристика), номер абзаца и фрагмент текста."""
    loc = issue.get("location") or {}
    parts: list[str] = []
    page = loc.get("page")
    if page:
        parts.append(f"стр. ~{page}")
    para = loc.get("paragraph")
    if para is not None:
        parts.append(f"абз. {int(para) + 1}")
    snippet = str(loc.get("snippet") or "").strip()
    if snippet:
        parts.append(f"«{snippet}»")
    return ", ".join(parts) if parts else "весь документ"


def _esc_md(text: str) -> str:
    """Пипы и переносы ломают таблицу Markdown — экранируем/сворачиваем."""
    return str(text).replace("|", "\\|").replace("\n", " ")


def _plural(count: int, one: str, few: str, many: str) -> str:
    """Русское склонение числительного: 1 ошибка, 2 ошибки, 5 ошибок."""
    mod10, mod100 = count % 10, count % 100
    if mod10 == 1 and mod100 != 11:
        word = one
    elif 2 <= mod10 <= 4 and not 12 <= mod100 <= 14:
        word = few
    else:
        word = many
    return f"{count} {word}"


def format_summary(summary: Mapping[str, int] | None) -> str:
    """«1 ошибка, 3 предупреждения, 9 рекомендаций» — одна строка для лога и статуса."""
    data = summary or {}
    errors = int(data.get("error", 0))
    warnings = int(data.get("warning", 0))
    infos = int(data.get("info", 0))
    if not (errors or warnings or infos):
        return "Замечаний нет"
    parts = []
    if errors:
        parts.append(_plural(errors, "ошибка", "ошибки", "ошибок"))
    if warnings:
        parts.append(_plural(warnings, "предупреждение", "предупреждения", "предупреждений"))
    if infos:
        parts.append(_plural(infos, "рекомендация", "рекомендации", "рекомендаций"))
    return ", ".join(parts)


def _header_lines(result: Mapping[str, Any], rules: Mapping[str, Any],
                  *, stamp: str | None = None) -> list[str]:
    """Шапка: файл, дата, пресет, режим строгости, версия приложения."""
    created = stamp or _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    preset = str(result.get("preset") or rules.get("name") or "—")
    strict = bool(result.get("strict", rules.get("strict", True)))
    path = str(result.get("file") or "—")
    return [
        f"# Отчёт проверки: {os.path.basename(path)}",
        "",
        f"- **Файл:** `{path}`",
        f"- **Дата:** {created}",
        f"- **Пресет:** {preset}",
        f"- **Режим:** {'строгий (strict)' if strict else 'мягкий (strict: false — ошибки снижены до предупреждений)'}",
        f"- **Результат:** {format_summary(result.get('summary'))}",
        f"- **Приложение:** {constants.APP_NAME} {constants.APP_VERSION}",
    ]


def _manual_items(issues: Sequence[Mapping[str, Any]]) -> list[tuple[str, str]]:
    """Чек-лист из замечаний кода MANUAL: (что проверить, как)."""
    out = []
    for issue in issues:
        if issue.get("code") == MANUAL_CODE:
            title = str(issue.get("message", "")).replace("Проверить вручную:", "").strip()
            out.append((title, str(issue.get("hint", "")).strip()))
    return out


def _read_error_block(result: Mapping[str, Any]) -> list[str]:
    return [
        "## Файл не удалось прочитать",
        "",
        f"> {result.get('read_error')}",
        "",
        "Проверка не выполнялась. Убедитесь, что файл является DOCX (zip-пакет OOXML) "
        "и не повреждён.",
    ]


# --- Markdown ---------------------------------------------------------------
def render_markdown(result: Mapping[str, Any], rules: Mapping[str, Any],
                    *, stamp: str | None = None) -> str:
    """Отчёт в Markdown: сводка, таблица замечаний, чек-лист «Проверить вручную»."""
    issues = _issues(result)
    lines = _header_lines(result, rules, stamp=stamp)
    lines.append("")
    if result.get("read_error"):
        lines.extend(_read_error_block(result))
        return "\n".join(lines) + "\n"

    summary = result.get("summary") or {}
    lines += ["## Сводка", "",
              "| Ошибки | Предупреждения | Рекомендации |",
              f"|---:|---:|---:|",
              f"| {int(summary.get('error', 0))} | {int(summary.get('warning', 0))} | "
              f"{int(summary.get('info', 0))} |", ""]

    table = [i for i in issues if i.get("code") != MANUAL_CODE]
    lines.append("## Замечания")
    if not table:
        lines += ["", "Замечаний нет — документ соответствует пресету.", ""]
    else:
        lines += ["", "| Severity | Код | Замечание | Где | Как исправить |",
                  "|---|---|---|---|---|"]
        for issue in table:
            severity = str(issue.get("severity", "info"))
            code = str(issue.get("code", ""))
            message = _esc_md(issue.get("message", ""))
            hint = _esc_md(issue.get("hint") or _hint_for(code))
            lines.append(f"| {SEVERITY_LABELS.get(severity, severity)} | `{code}` | {message} "
                         f"| {_esc_md(_where(issue))} | {hint} |")
        lines.append("")

    manual = _manual_items(issues)
    if manual:
        lines += ["## Проверить вручную", ""]
        lines += [f"- [ ] **{title}** — {how}" for title, how in manual]
        lines.append("")
    return "\n".join(lines) + "\n"


# --- Текст ------------------------------------------------------------------
def render_text(result: Mapping[str, Any], rules: Mapping[str, Any],
                *, stamp: str | None = None, width: int = 88) -> str:
    """Тот же отчёт для терминала/лога: без таблиц, с переносами по ширине."""
    issues = _issues(result)
    out = [f"ОТЧЁТ ПРОВЕРКИ: {os.path.basename(str(result.get('file') or '—'))}",
           "=" * min(width, 60)]
    created = stamp or _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    strict = bool(result.get("strict", rules.get("strict", True)))
    out += [f"Файл: {result.get('file') or '—'}", f"Дата: {created}",
            f"Пресет: {result.get('preset') or rules.get('name') or '—'}",
            f"Режим: {'strict' if strict else 'мягкий'}",
            f"Итог: {format_summary(result.get('summary'))}", ""]
    if result.get("read_error"):
        out += ["Файл не читается:", str(result["read_error"])]
        return "\n".join(out) + "\n"
    if not issues:
        out.append("Замечаний нет — документ соответствует пресету.")
    for num, issue in enumerate(issues, start=1):
        severity = SEVERITY_SHORT.get(str(issue.get("severity")), "INFO")
        out.append(f"{num:>3}. [{severity}] {issue.get('code')} — {_where(issue)}")
        for line in textwrap.wrap(str(issue.get("message", "")).strip(), width) or [""]:
            out.append(f"      {line}".rstrip())
        hint = str(issue.get("hint") or explain(str(issue.get("code", "")))).strip()
        if hint and issue.get("code") != MANUAL_CODE:
            for line in textwrap.wrap(f"Как исправить: {hint}", width):
                out.append(f"      {line}".rstrip())
        out.append("")
    return "\n".join(out).rstrip() + "\n"


def save_report(result: Mapping[str, Any], rules: Mapping[str, Any], dest_path: Any,
                fmt: str = "md") -> Path:
    """Сохранить отчёт; возвращает путь. `fmt`: md/markdown либо txt/text."""
    fmt_norm = str(fmt or "md").lower().lstrip(".")
    if fmt_norm in MD_FORMATS:
        text, suffix = render_markdown(result, rules), constants.REPORT_MD_EXT
    elif fmt_norm in TXT_FORMATS:
        text, suffix = render_text(result, rules), constants.REPORT_TXT_EXT
    else:
        raise ValueError(f"Неизвестный формат отчёта: {fmt} (ожидаются md или txt)")
    dest = Path(str(dest_path))
    if dest.suffix.lower() not in (".md", ".txt"):
        dest = dest.with_name(dest.name + suffix)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(text, encoding="utf-8")
    return dest


# --- Пакетный режим ---------------------------------------------------------
STATUS_LABELS = {
    "ok": "Без замечаний",
    "violations": "Есть ошибки",
    "warnings": "Есть замечания",
    "read_error": "Ошибка чтения",
    "normalize_error": "Ошибка нормализации",
    "script_failed": "Сбой скрипта",
    "error": "Ошибка",
}


def _status_label(status: str) -> str:
    """Машинный ключ статуса из CLI → человекочитаемая строка отчёта (п. 19.4 vs §12)."""
    return STATUS_LABELS.get(status.strip().lower(), status)


def render_batch_markdown(items: Iterable[Mapping[str, Any]], *,
                          stamp: str | None = None) -> str:
    """Сводная таблица по нескольким файлам (вкладка «Проверка», режим «Папка»).

    `items` = `[{"file", "summary", "status", "result"?}]`; статус может быть
    машинным ключом CLI («ok», «violations») или готовой строкой UI — переводится
    только ключ, строка пользователя печатается как есть.
    """
    created = stamp or _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    rows = list(items)
    total = {"error": 0, "warning": 0, "info": 0}
    lines = ["# Пакетная проверка документов", "",
             f"- **Дата:** {created}",
             f"- **Файлов:** {len(rows)}", "",
             "| Файл | Ошибок | Предупреждений | Рекомендаций | Статус |",
             "|---|---:|---:|---:|---|"]
    for item in rows:
        summary = item.get("summary") or {}
        for key in total:
            total[key] += int(summary.get(key, 0) or 0)
        status = _status_label(str(item.get("status") or "")) if item.get("status") else ""
        note = ""
        result = item.get("result") or {}
        if result.get("read_error"):
            note = f" — {result['read_error']}"
        name = os.path.basename(str(item.get("file") or "—"))
        lines.append(f"| `{_esc_md(name)}` | {int(summary.get('error', 0) or 0)} | "
                     f"{int(summary.get('warning', 0) or 0)} | "
                     f"{int(summary.get('info', 0) or 0)} | "
                     f"{_esc_md(status or _fallback_status(summary, result.get('read_error')))}"
                     f"{_esc_md(note)} |")
    lines += ["", f"**Итого:** {format_summary(total)}", ""]
    broken = [item for item in rows if (item.get("result") or {}).get("read_error")]
    if broken:
        lines += ["## Файлы, которые не удалось прочитать", ""]
        lines += [f"- `{os.path.basename(str(item.get('file', '')))}`: "
                  f"{item['result']['read_error']}" for item in broken]
        lines.append("")
    return "\n".join(lines) + "\n"


def _fallback_status(summary: Mapping[str, int], read_error: Any) -> str:
    """Статус, если вызывающий код его не передал (например, в JSON из CLI)."""
    if read_error:
        return "Ошибка чтения"
    if int(summary.get("error", 0) or 0):
        return "Есть ошибки"
    if int(summary.get("warning", 0) or 0) or int(summary.get("info", 0) or 0):
        return "Есть замечания"
    return "Без замечаний"


def render_batch_text(items: Iterable[Mapping[str, Any]], *, stamp: str | None = None) -> str:
    """Пакетная сводка для терминала: одна строка на файл плюс итоги."""
    rows = list(items)
    created = stamp or _dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    total: dict[str, int] = {"error": 0, "warning": 0, "info": 0}
    out = [_plural(len(rows), "файл", "файла", "файлов") + " проверено",
           "=" * 60, f"Дата: {created}", ""]
    for num, item in enumerate(rows, start=1):
        summary = item.get("summary") or {}
        for key in total:
            total[key] += int(summary.get(key, 0) or 0)
        result = item.get("result") or {}
        read_error = result.get("read_error")
        status = _status_label(str(item.get("status") or _fallback_status(summary, read_error)))
        name = os.path.basename(str(item.get("file") or "—"))
        out.append(f"{num:>3}. {name} — {format_summary(summary)} — {status}")
        if read_error:
            out.append(f"      {read_error}")
    out += ["", f"Итого: {format_summary(total)}"]
    return "\n".join(out).rstrip() + "\n"


def save_batch_report(items: Sequence[Mapping[str, Any]], dest_path: Any,
                      fmt: str = "md") -> Path:
    """Сохранить сводку пакетной проверки (md или txt)."""
    fmt_norm = str(fmt or "md").lower().lstrip(".")
    dest = Path(str(dest_path))
    if fmt_norm in TXT_FORMATS:
        text = render_batch_text(items)
        target_ext = constants.REPORT_TXT_EXT
    elif fmt_norm in MD_FORMATS:
        text = render_batch_markdown(items)
        target_ext = constants.REPORT_MD_EXT
    else:
        raise ValueError(f"Неизвестный формат отчёта: {fmt} (ожидаются md или txt)")
    if dest.suffix.lower() not in (".md", ".txt"):
        dest = dest.with_name(dest.name + target_ext)
    dest.parent.mkdir(parents=True, exist_ok=True)
    dest.write_text(text, encoding="utf-8")
    return dest


# --- Заделы v2 (архитектурные хуки, не заглушки фич v1) --------------------
def export_pdf(path: Any) -> Path:
    """Экспорт отчёта/документа в PDF через LibreOffice headless (п. 13 ТЗ, v2).

    Задел: в v1 отчёт отдаётся в .md/.txt, PDF не входит в требования приёмки.
    Планируемая реализация: `soffice --headless --convert-to pdf --outdir <dir> <file>`
    с поиском бинаря по `shutil.which`, таймаутом и кодом возврата `EXIT_SCRIPT_FAILED`;
    вызов подключать в UI после появления настройки «Каталог LibreOffice».
    """
    raise NotImplementedError(
        "Экспорт в PDF запланирован в v2 (LibreOffice headless). См. docstring export_pdf.")


def _build_parser() -> Any:
    """Задел модуля-парсера библиографии (п. 13 ТЗ, v2 — «редактор списка литературы»).

    В v1 список литературы проверяется эвристикой `docx_content_checks` (номера, годы,
    затекстовые ссылки). Полноценный парсер ГОСТ-записи (автор/заглавие/издательство/объём)
    выносится в отдельный модуль, чтобы не смешивать со слоистым сравнением строк.
    Парсер в v1 СОЗНАТЕЛЬНО не реализован.
    """
    raise NotImplementedError(
        "Парсер библиографических записей — задел v2; в v1 используется эвристическая "
        "проверка BIB_ORDER/BIB_AGE/BIB_COVERAGE.")
