"""Проверки содержательных объектов: рисунки, таблицы, формулы, сноски, библиография.

Второй слой проверок (первый — `docx_checks.py`: страница/текст/заголовки).
Модули разбиты из-за требования «файлы > 600 строк разбивать».

Всё, что проверяется, — из ключей правил: шаблоны подписей компилируются в regex
(`dp.template_to_regex`), а при выключенном ключе (`require_reference_in_text: false`,
`max_source_age_years: 0`, `toc.required: false`) проверка просто не запускается.
"""

from __future__ import annotations

import datetime as _dt
import re
from typing import Any, Callable, Iterable, Mapping, Sequence

import xml.etree.ElementTree as ET

from app import constants
from app.core import docx_checks as base
from app.core import docx_parser as dp
from app.core import docx_xml as dx
from app.core.docx_checks import CheckContext
from app.core.docx_parser import ParagraphInfo
from app.core.units import half_points_to_pt, pt_to_half_points

# Список того, что автоматикой не проверяется (п. 9.1 ТЗ, код MANUAL).
MANUAL_ITEMS: Sequence[tuple[str, str]] = (
    ("Орфография и пунктуация", "Прогоните проверку орфографии в Word (F7): приложение её не делает."),
    ("Титульный лист по шаблону вуза",
     "Сверьте реквизиты титула с актуальной методичкой вашего вуза (вуз, кафедра, город, год)."),
    ("Графика не сканами",
     "Рисунки и таблицы должны быть перенесены из Word/Excel, а не вставлены картинкой-сканом."),
    ("Оригинальность текста",
     "Проверьте в Антиплагиат.ВУЗ/ETXT от введения до конца списка литературы; накрутки запрещены."),
    ("Заполненность страниц",
     "Оценка объёма в приложении эвристическая: реально проверьте, что страница заполнена ≥ 80 %."),
    ("Выводы по каждой главе (ВКР)",
     "После каждой главы должны быть краткие выводы — содержание проверить нельзя."),
    ("Соответствие типа работы",
     "Не перепутаны ли типы документов (например слово «диплом» в курсовой)."),
    ("Совпадение заголовков с содержанием",
     "Обновите поля оглавления (Ctrl+A → F9) и сверьте названия разделов с текстом."),
)

_CITATION_RE = re.compile(r"\[\s*(\d{1,3})(?:\s*[,;]\s*(?:с\.?|стр\.?)?\s*[\d\-–—,\s]+)?\s*\]")
_YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2})\b")
_RECORD_RE = re.compile(r"^\s*(\d{1,3})[.)]\s+\S")


# --- Рисунки ---------------------------------------------------------------
def _figure_captions(ctx: CheckContext) -> list[tuple[ParagraphInfo, re.Match[str]]]:
    rx = ctx.regexes.get("figure")
    if rx is None:
        return []
    out: list[tuple[ParagraphInfo, re.Match[str]]] = []
    for para in ctx.model.paragraphs:
        if para.in_table:
            continue
        match = rx.match(" ".join(para.text.split()).strip())
        if match:
            out.append((para, match))
    return out


def _prev_content_paragraph(ctx: CheckContext, index: int) -> ParagraphInfo | None:
    """Предшествующий значимый абзац (для «подпись под рисунком» / «название над таблицей»)."""
    for para in reversed(ctx.model.paragraphs):
        if para.index >= index:
            continue
        if para.in_table:
            continue
        return para if (para.text.strip() or para.has_drawing) else None
    return None


def _next_block_paragraph(ctx: CheckContext, index: int) -> ParagraphInfo | None:
    """Следующий значимый абзац; абзацы ячеек возвращаются — вызывающий сверяет их с таблицей."""
    for para in ctx.model.paragraphs:
        if para.index <= index:
            continue
        return para if (para.text.strip() or para.has_drawing) else None
    return None


def check_figures(ctx: CheckContext) -> None:
    """FIG_CAPTION / FIG_COUNT / FIG_ORDER / FIG_REF (п. 5.3 ТЗ)."""
    figures = ctx.rules.get("figures") or {}
    if not figures:
        return
    captions = _figure_captions(ctx)
    rx = ctx.regexes.get("figure")
    images = ctx.model.images

    # FIG_COUNT: подписей должно быть не меньше, чем изображений
    if images and len(captions) < len(images):
        caption_indexes = {p.index for p, _ in captions}
        orphan = next((img.paragraph_index for img in images
                       if _caption_for_image(ctx, img.paragraph_index, caption_indexes) is None),
                      None)
        ctx.add("FIG_COUNT",
                f"Изображений {len(images)}, а подписей «Рисунок N – …» — {len(captions)}. "
                f"Часть рисунков без подписи.",
                paragraph=orphan)

    want_align = dp.JC_ALIASES.get(str(figures.get("caption_align", "center")), "center")
    dash = str(figures.get("dash", "–"))
    for para, match in captions:
        problems: list[str] = []
        prev = _prev_content_paragraph(ctx, para.index)
        if prev is None or not prev.has_drawing:
            problems.append("подпись стоит не под изображением")
        if (para.eff.jc or "left") != want_align:
            problems.append(f"выравнивание «{para.eff.jc or 'left'}» вместо «{want_align}»")
        if not figures.get("caption_indent", False) and (para.eff.first_line or 0) > 10:
            problems.append("есть абзацный отступ (должен быть 0)")
        if figures.get("caption_no_trailing_dot", True) and para.text.strip().endswith("."):
            problems.append("точка в конце подписи")
        if dash and dash not in para.text:
            problems.append(f"нет тире «{dash}» между номером и названием")
        if not match.group(2).strip():
            problems.append("пустое название после тире")
        if problems:
            ctx.add("FIG_CAPTION",
                    f"Подпись «{dp.snippet_of(para.text, 60)}»: " + "; ".join(problems) + ".",
                    paragraph=para.index)

    # FIG_ORDER: сквозная нумерация без пропусков и повторов
    if figures.get("numbering", "continuous") != "none":
        numbers = [_as_number(match.group(1)) for _, match in captions]
        problems = _sequence_problems(numbers, per_section=(figures.get("numbering") == "by_section"))
        if problems:
            ctx.add("FIG_ORDER", "Нумерация рисунков: " + "; ".join(problems) + ".",
                    paragraph=captions[0][0].index if captions else None)

    # FIG_REF: ссылка в тексте на каждый рисунок
    if figures.get("require_reference_in_text", True):
        template = str(figures.get("reference_template", "в соответствии с рисунком {n}"))
        body_text = " ".join(p.text for p in ctx.model.paragraphs
                             if not p.in_table and not _is_figure_caption(p, rx))
        for para, match in captions:
            num = match.group(1)
            if not _has_ref(body_text, template, num):
                ctx.add("FIG_REF",
                        f"На рисунок {num} нет ссылки в тексте "
                        f"(ожидается «{str(figures.get('reference_template', '')).format(n=num)}»).",
                        paragraph=para.index)


def _is_figure_caption(para: ParagraphInfo, rx: re.Pattern[str] | None) -> bool:
    return bool(rx is not None and rx.match(" ".join(para.text.split()).strip()))


def _has_ref(text: str, template: str, num: str) -> bool:
    """Есть ли в тексте ссылка вида шаблона («в соответствии с рисунком 1») или любая
    корректная отсылка «рисунок N» — второй вариант принимаем, чтобы не плодить ложные ошибки."""
    literal = dp._flexi_literal(str(template).replace("{n}", num))
    if re.search(literal, text, re.IGNORECASE):
        return True
    return bool(re.search(rf"рисунк\w*\s*№?\s*{re.escape(num)}\b", text, re.IGNORECASE))


def _caption_for_image(ctx: CheckContext, image_index: int,
                       caption_indexes: set[int]) -> int | None:
    """Индекс подписи, относящейся к рисунку (ближайшая непустая подпись НИЖЕ рисунка)."""
    for para in ctx.model.paragraphs:
        if para.index <= image_index or para.in_table or para.index not in caption_indexes:
            continue
        return para.index
    return None


def _as_number(raw: str) -> Any:
    """Номер подписи: «12» → 12, «1.2» → (1, 2) — чтобы сравнивать разделовую нумерацию."""
    if "." in raw:
        return tuple(int(x) for x in raw.split("."))
    return int(raw)


def _sequence_problems(numbers: Iterable[Any], *, per_section: bool = False) -> list[str]:
    """Проверка сквозной нумерации: монотонность, пропуски, повторы."""
    seq = [n for n in numbers if n is not None]
    problems: list[str] = []
    if per_section:
        return problems  # по разделам своя арифметика — не проверяем порядок
    expected = 1
    for value in seq:
        if not isinstance(value, int):
            problems.append(f"номер {value} не соответствует сквозной нумерации")
            continue
        if value == expected:
            expected += 1
        elif value == expected - 1:
            problems.append(f"повтор номера {value}")
        else:
            problems.append(f"пропуск: ожидается {expected}, стоит {value}")
            expected = value + 1
    return problems


# --- Таблицы ---------------------------------------------------------------
def _table_captions(ctx: CheckContext) -> list[tuple[ParagraphInfo, re.Match[str]]]:
    rx = ctx.regexes.get("table")
    if rx is None:
        return []
    out: list[tuple[ParagraphInfo, re.Match[str]]] = []
    for para in ctx.model.paragraphs:
        if para.in_table:
            continue
        match = rx.match(" ".join(para.text.split()).strip())
        if match:
            out.append((para, match))
    return out


def check_tables(ctx: CheckContext) -> None:
    """TBL_CAPTION / TBL_HEADER / TBL_BORDERS / TBL_ORDER / TBL_REF (п. 5.4 ТЗ)."""
    tables = ctx.rules.get("tables") or {}
    if not tables:
        return
    rx = ctx.regexes.get("table")
    captions = _table_captions(ctx)
    caption_by_para = {p.index: m for p, m in captions}

    # Название над каждой таблицей + корректное форматирование
    want_align = dp.JC_ALIASES.get(str(tables.get("caption_align", "left")), "left")
    dash = str(tables.get("dash", "–"))
    for table in ctx.model.tables:
        prev = _prev_content_paragraph(ctx, _first_para_index(ctx, table.paragraph_indexes))
        if prev is None:
            ctx.add("TBL_CAPTION", "Таблица без названия: перед ней нет абзаца вообще.",
                    paragraph=table.paragraph_indexes[0] if table.paragraph_indexes else None)
            continue
        match = caption_by_para.get(prev.index) if rx is not None else None
        if match is None:
            if prev.text.strip() and rx is not None and not rx.match(
                    " ".join(prev.text.split()).strip()):
                ctx.add("TBL_CAPTION",
                        f"Перед таблицей нет названия вида «{tables.get('caption_template', '')}»: "
                        f"стоит «{dp.snippet_of(prev.text, 50)}».",
                        paragraph=prev.index)
            continue
        problems: list[str] = []
        next_para = _next_block_paragraph(ctx, prev.index)
        if next_para is None or next_para.index not in table.paragraph_indexes:
            problems.append("название не непосредственно над таблицей")
        if (prev.eff.jc or "left") != want_align:
            problems.append(f"выравнивание «{prev.eff.jc or 'left'}» вместо «{want_align}» "
                            f"(название таблицы ставится слева)")
        if not tables.get("caption_indent", False) and (prev.eff.first_line or 0) > 10:
            problems.append("абзацный отступ должен быть 0")
        if tables.get("caption_no_trailing_dot", True) and prev.text.strip().endswith("."):
            problems.append("точка в конце названия")
        if dash and dash not in prev.text:
            problems.append(f"нет тире «{dash}»")
        if problems:
            ctx.add("TBL_CAPTION",
                    f"Название «{dp.snippet_of(prev.text, 60)}»: " + "; ".join(problems) + ".",
                    paragraph=prev.index)

    for table in ctx.model.tables:
        _check_table_header(ctx, tables, table)
        _check_table_borders(ctx, tables, table)

    # TBL_ORDER
    if tables.get("caption_template"):
        numbers = [_as_number(m.group(1)) for _, m in captions]
        problems = _sequence_problems(numbers)
        if problems:
            ctx.add("TBL_ORDER", "Нумерация таблиц: " + "; ".join(problems) + ".",
                    paragraph=captions[0][0].index if captions else None)

    # TBL_REF
    if tables.get("require_reference_in_text", True):
        body_text = " ".join(p.text for p in ctx.model.paragraphs
                             if not p.in_table and p.index not in caption_by_para)
        for para, match in captions:
            num = match.group(1)
            if not re.search(rf"таблиц\w*\s*(?:№\s*)?{re.escape(num)}\b", body_text, re.IGNORECASE):
                ctx.add("TBL_REF", f"На таблицу {num} нет ссылки в тексте («в таблице {num}»).",
                        paragraph=para.index)

    # «Продолжение таблицы N»: если стоит — должен быть корректный номер и правый край
    cont_rx = ctx.regexes.get("continuation")
    if cont_rx is not None:
        want = dp.JC_ALIASES.get(str(tables.get("continuation_align", "right")), "right")
        for para in ctx.model.paragraphs:
            if para.in_table:
                continue
            text = " ".join(para.text.split()).strip()
            if cont_rx.match(text) and (para.eff.jc or "left") != want:
                ctx.add("TBL_CAPTION",
                        f"«{dp.snippet_of(text, 50)}» выровнено «{para.eff.jc or 'left'}», "
                        f"над продолжением таблицы требуется «{want}».", paragraph=para.index)


def _first_para_index(ctx: CheckContext, indexes: list[int]) -> int:
    """Индекс первого абзаца таблицы (для поиска предыдущего блока)."""
    return indexes[0] if indexes else -1


def _check_table_header(ctx: CheckContext, tables: Mapping[str, Any],
                        table: dp.TableInfo) -> None:
    """TBL_HEADER: повтор шапки (w:trPr/w:tblHeader) и её полужирание."""
    if not table.rows:
        return
    head = table.rows[0]
    problems: list[str] = []
    if tables.get("repeat_header", True) and not head.is_header:
        problems.append("нет признака заголовка строки (w:trPr/w:tblHeader) — шапка не повторится")
    if tables.get("header_bold", True):
        indexes = {i for i in _row_paragraph_indexes(ctx, head)}
        bold_ratio = _bold_ratio_of(ctx, indexes)
        if indexes and bold_ratio < 0.6:
            problems.append(f"шапка не полужирная (жирного {bold_ratio * 100:.0f} % символов)")
    if problems:
        first = table.paragraph_indexes[0] if table.paragraph_indexes else None
        ctx.add("TBL_HEADER", "Таблица: " + "; ".join(problems) + ".", paragraph=first)


def _row_paragraph_indexes(ctx: CheckContext, row: dp.RowInfo) -> list[int]:
    by_id = {id(p.element): p.index for p in ctx.model.paragraphs}
    return [by_id[id(el)] for el in row.element.iter(dx.qn("w:p")) if id(el) in by_id]


def _bold_ratio_of(ctx: CheckContext, indexes: Iterable[int]) -> float:
    total = 0
    bold = 0
    wanted = set(indexes)
    for para in ctx.model.paragraphs:
        if para.index not in wanted:
            continue
        for run in para.runs:
            weight = len(run.text.strip())
            if not weight:
                continue
            total += weight
            if run.props.bold:
                bold += weight
    return (bold / total) if total else 0.0


def _check_table_borders(ctx: CheckContext, tables: Mapping[str, Any],
                         table: dp.TableInfo) -> None:
    """TBL_BORDERS: тонкие сплошные границы (tblBorders, sz в 1/8 пункта)."""
    kind = str(tables.get("borders", "single_thin"))
    if kind in ("none", "", "style"):
        return
    max_sz = int(tables.get("border_size_eighths_pt", 4))
    tbl_pr = dx.find_child(table.element, "w:tblPr")
    borders = dx.find_child(tbl_pr, "w:tblBorders") if tbl_pr is not None else None
    needed = ("top", "left", "bottom", "right", "insideH", "insideV")
    if borders is None:
        ctx.add("TBL_BORDERS",
                "У таблицы нет явных границ (w:tblBorders): границы могут быть не заданы или "
                "взяться из стиля.",
                paragraph=table.paragraph_indexes[0] if table.paragraph_indexes else None)
        return
    missing = [side for side in needed
               if dx.find_child(borders, f"w:{side}") is None]
    thick = []
    not_single = []
    for side in needed:
        el = dx.find_child(borders, f"w:{side}")
        if el is None:
            continue
        val = (el.get(dx.qn("w:val")) or "single").lower()
        if val not in ("single",):
            not_single.append(f"{side}={val}")
        size = dx.get_int(el, "w:sz")
        if size is not None and size > max(max_sz, 4):
            thick.append(f"{side}={size / 8:g} pt")
    problems = []
    if missing:
        problems.append(f"нет границ: {', '.join(missing)}")
    if not_single:
        problems.append(f"границы не сплошные: {', '.join(not_single)}")
    if thick:
        problems.append(f"границы толще {max_sz / 8:g} pt: {', '.join(thick)}")
    if problems:
        ctx.add("TBL_BORDERS", "Таблица: " + "; ".join(problems) + ".",
                paragraph=table.paragraph_indexes[0] if table.paragraph_indexes else None)


# --- Формулы ---------------------------------------------------------------
def check_formulas(ctx: CheckContext) -> None:
    """FORMULA: номер (N) у правого края и «где» без двоеточия (п. 5.5 ТЗ)."""
    formulas = ctx.rules.get("formulas") or {}
    if not formulas:
        return
    num_fmt = str(formulas.get("number_format", "({n})"))
    num_rx = dp.template_to_regex(num_fmt, anchored=False)
    numbers: list[int] = []
    for para in ctx.model.paragraphs:
        if para.in_table or not _looks_like_formula(para):
            continue
        tail = para.text.strip()
        match = None
        for match in num_rx.finditer(tail):
            pass  # берём последнее вхождение: номер стоит в конце строки
        if match is None or not tail.endswith(match.group(0).strip()):
            ctx.add("FORMULA",
                    f"Формула без номера у правого края «{num_fmt}»: "
                    f"«{dp.snippet_of(para.text, 60)}».", paragraph=para.index)
        else:
            numbers.append(int(match.group(1)))
    problems = _sequence_problems(numbers)
    if problems:
        ctx.add("FORMULA", "Нумерация формул: " + "; ".join(problems) + ".", severity="warning")

    word = str(formulas.get("explanation_first_word", "где"))
    if formulas.get("explanation_no_colon", True) and word:
        for para in ctx.model.paragraphs:
            text = para.text.strip()
            if re.match(rf"^{re.escape(word)}\s*:", text, re.IGNORECASE):
                ctx.add("FORMULA",
                        f"Пояснение формулы: «{dp.snippet_of(text, 40)}» — после «{word}» "
                        f"двоеточие не ставится.", paragraph=para.index)


def _looks_like_formula(para: ParagraphInfo) -> bool:
    """Формула: Math-элемент или строка со знаком «=» и не текстовая подписью."""
    if any(True for _ in para.element.iter(dx.qn("m:oMath"))):
        return True
    text = para.text.strip()
    return bool(text) and "=" in text and len(text) <= 120 and not text.endswith(".")


# --- Сноски ----------------------------------------------------------------
def check_footnotes(ctx: CheckContext) -> None:
    """FOOTNOTE_FONT: word/footnotes.xml — кегль в [min_allowed, size] и шрифт."""
    notes = ctx.rules.get("footnotes") or {}
    root = ctx.model.footnotes
    if root is None or not notes:
        return  # подстрочных ссылок в документе нет — проверять нечего
    want = pt_to_half_points(float(notes.get("size_pt", 12)))
    low = pt_to_half_points(float(notes.get("min_allowed_size_pt", notes.get("size_pt", 12))))
    want_font = str(notes.get("font", constants.DEFAULT_FONT)).lower()
    bad_size: list[float] = []
    bad_font: set[str] = set()
    for note in dx.find_all(root, "w:footnote"):
        ntype = note.get(dx.qn("w:type"))
        if ntype in ("separator", "continuationSeparator"):
            continue
        for para in dx.find_all(note, "w:p"):
            direct = ctx.model.registry.raw_para_props(dx.find_child(para, "w:pPr"))
            style_id = direct.get("style")
            for run in para.iter(dx.qn("w:r")):
                props = ctx.model.registry.run_props(style_id,
                                                     ctx.model.registry.raw_run_props(
                                                         dx.find_child(run, "w:rPr")))
                raw = dp.to_run_props(props, ctx.model.mono_fonts)
                if not run_text(run).strip():
                    continue
                if raw.size is not None and not (low <= raw.size <= want):
                    bad_size.append(half_points_to_pt(int(raw.size)))
                if raw.font and raw.font.lower() not in (want_font,) and not raw.mono:
                    bad_font.add(raw.font)
    if bad_size:
        uniq = ", ".join(f"{v:g} pt" for v in sorted(set(bad_size)))
        ctx.add("FOOTNOTE_FONT",
                f"Кегль текста сносок {uniq}, правило допускает от "
                f"{half_points_to_pt(low):g} до {half_points_to_pt(want):g} pt.")
    if bad_font:
        ctx.add("FOOTNOTE_FONT",
                f"Шрифт сносок {', '.join(sorted(bad_font))}, правило требует "
                f"«{notes.get('font')}».")


def run_text(r_el: ET.Element) -> str:
    """Текст run'а в произвольной части пакета (сноски, колонтитулы)."""
    out: list[str] = []
    for child in list(r_el):
        tag = dx.local_name(child.tag)
        if tag == "t":
            out.append(child.text or "")
        elif tag in ("tab", "br", "cr", "sym"):
            out.append(" ")
    return "".join(out)


# --- Содержание ------------------------------------------------------------
def check_toc(ctx: CheckContext) -> None:
    """TOC_PRESENT: автосодержание (поле TOC в document.xml) + updateFields в settings.xml."""
    toc = ctx.rules.get("toc") or {}
    if not toc:
        return
    found = any(re.search(r"\bTOC\b", i.text or "") for i in
                ctx.model.document.iter(dx.qn("w:instrText")))
    if not found:
        found = any(re.search(r"\bTOC\b", el.get(dx.qn("w:instr")) or "")
                    for el in ctx.model.document.iter(dx.qn("w:fldSimple")))
    if not found and toc.get("required", True):
        ctx.add("TOC_PRESENT", "Автосодержание отсутствует: в документе нет поля TOC.")
    elif not found:
        ctx.add("TOC_PRESENT", "Автосодержание отсутствует (пресет не требует его).",
                severity="info")
        return
    if toc.get("auto", True) and ctx.model.settings is not None:
        # ВАЖНО: проверка `is None`, а не `not el` — пустой Element ложноподобен (DeprecationWarning)
        if dx.find_child(ctx.model.settings, "w:updateFields") is None:
            ctx.add("TOC_PRESENT",
                    "Поле TOC есть, но в settings.xml нет w:updateFields: при открытии "
                    "содержание не обновится автоматически.", severity="warning")


# --- Библиография ----------------------------------------------------------
def _biblio_heading_index(ctx: CheckContext) -> int | None:
    """Индекс абзаца «СПИСОК … ИСТОЧНИКОВ»: от него ищем записи."""
    for para in ctx.model.paragraphs:
        if para.in_table:
            continue
        text = " ".join(para.text.split()).strip().upper().rstrip(".")
        if text.startswith("СПИСОК") and ("ИСТОЧН" in text or "ЛИТЕРАТУР" in text):
            return para.index
    return None


def bibliography_records(ctx: CheckContext) -> list[tuple[ParagraphInfo, int]]:
    """Записи списка литературы: абзацы вида «1. …» после заголовка списка."""
    start = _biblio_heading_index(ctx)
    if start is None:
        return []
    out: list[tuple[ParagraphInfo, int]] = []
    for para in ctx.model.paragraphs:
        if para.index <= start or para.in_table:
            continue
        text = para.text.strip()
        if not text:
            continue
        if dp.is_structural_heading(text, ctx.rules) and not _RECORD_RE.match(text):
            break  # начался ПРИЛОЖЕНИЕ — список закончился
        match = _RECORD_RE.match(text)
        if match:
            out.append((para, int(match.group(1))))
    return out


def cited_numbers(ctx: CheckContext) -> list[int]:
    """Номера затекстовых ссылок в порядке появления (стиль square_brackets)."""
    if str(dp.rget(ctx.rules, "references.style", "square_brackets")) != "square_brackets":
        return []
    start = _biblio_heading_index(ctx)
    out: list[int] = []
    for para in ctx.model.paragraphs:
        if para.in_table:
            continue
        if start is not None and para.index >= start:
            break
        text = para.text
        if dp.is_caption(text, ctx.regexes):
            continue
        for match in _CITATION_RE.finditer(text):
            out.append(int(match.group(1)))
    return out


def check_bibliography(ctx: CheckContext) -> None:
    """BIB_ORDER / BIB_AGE / BIB_COVERAGE (п. 5.6 ТЗ)."""
    refs = ctx.rules.get("references") or {}
    records = bibliography_records(ctx)
    if not records:
        if dp.rget(refs, "min_cited_percent"):
            ctx.add("BIB_COVERAGE",
                    "Список литературы не найден (нет заголовка «СПИСОК … ИСТОЧНИКОВ» с "
                    "нумерованными записями).", severity="info")
        return

    # BIB_ORDER: нумерация записей и соответствие порядку появления ссылок
    if str(refs.get("bibliography_order", "by_appearance")) == "by_appearance":
        numbers = [n for _, n in records]
        problems = _sequence_problems(numbers)
        cites = cited_numbers(ctx)
        appearance: list[int] = []
        for value in cites:
            if value not in appearance:
                appearance.append(value)
        if not problems and appearance and appearance[: len(numbers)] != numbers:
            wrong = next((i for i, (rec, num) in enumerate(zip(records, numbers))
                          if num != appearance[i]), None)
            ctx.add("BIB_ORDER",
                    "Список источников не совпадает с порядком появления ссылок в тексте: "
                    f"по тексту {appearance[:8]}…, в списке {numbers[:8]}…",
                    paragraph=records[wrong][0].index if wrong is not None else records[0][0].index)
        elif problems:
            ctx.add("BIB_ORDER", "Нумерация источников: " + "; ".join(problems) + ".",
                    paragraph=records[0][0].index)

    # BIB_AGE
    max_age = int(refs.get("max_source_age_years", 0) or 0)
    if max_age > 0:
        year_now = _dt.date.today().year
        limit = year_now - max_age
        stale = []
        for para, num in records:
            years = [int(y) for y in _YEAR_RE.findall(para.text)]
            newest = max(years) if years else None
            if newest is None:
                stale.append((num, para, "год не указан"))
            elif newest < limit:
                stale.append((num, para, f"{newest} г."))
        for num, para, why in stale[:5]:
            ctx.add("BIB_AGE",
                    f"Источник {num}: {why} — требуется не старше {limit} г. "
                    f"(правило: {max_age} лет).", paragraph=para.index)
        if len(stale) > 5:
            ctx.add("BIB_AGE", f"И ещё {len(stale) - 5} устаревших источников.",
                    paragraph=stale[5][1].index)

    # BIB_COVERAGE
    min_percent = int(refs.get("min_cited_percent", 0) or 0)
    if min_percent > 0:
        cited = set(cited_numbers(ctx))
        record_numbers = {n for _, n in records}
        used = cited & record_numbers
        percent = len(used) / len(record_numbers) * 100 if record_numbers else 100.0
        if percent < min_percent:
            missed = sorted(record_numbers - cited)[:8]
            ctx.add("BIB_COVERAGE",
                    f"Процитировано {percent:.0f} % источников (нужно ≥ {min_percent} %). "
                    f"Без ссылки остались: {', '.join(str(x) for x in missed)}.",
                    paragraph=records[0][0].index)


# --- Приложения ------------------------------------------------------------
def allowed_appendix_letters(rules: Mapping[str, Any]) -> tuple[str, ...]:
    """Буквы приложений: кириллица без Ё, З, Й, О, Ч, Ъ, Ы, Ь (п. 5.6 ТЗ)."""
    exclude = set(str(ch).upper() for ch in dp.rget(rules, "appendices.letters_exclude",
                                                   list(constants.APPENDIX_LETTERS)) or [])
    return tuple(ch for ch in "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ" if ch not in exclude)


def check_appendices(ctx: CheckContext) -> None:
    """APPENDIX: буквы, порядок, новая страница, выравнивание метки."""
    app_rules = ctx.rules.get("appendices") or {}
    rx = ctx.regexes.get("appendix")
    if rx is None:
        return
    allowed = allowed_appendix_letters(ctx.rules)
    found: list[tuple[ParagraphInfo, str]] = []
    for para in ctx.model.paragraphs:
        if para.in_table:
            continue
        text = " ".join(para.text.split()).strip()
        match = rx.match(text)
        if match and text.upper().startswith("ПРИЛОЖЕН"):
            found.append((para, match.group(1).upper()))
    if not found:
        return
    want_align = dp.JC_ALIASES.get(str(app_rules.get("label_align", "center")), "center")
    for para, letter in found:
        problems: list[str] = []
        if letter not in allowed:
            problems.append(f"буква «{letter}» не используется для приложений "
                            f"(исключены {', '.join(sorted(set('АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ') - set(allowed)))})")
        if (para.eff.jc or "left") != want_align:
            problems.append(f"метка выровнена «{para.eff.jc or 'left'}», нужно «{want_align}»")
        if app_rules.get("new_page", True) and not base._starts_new_page(ctx, para):
            problems.append("приложение начинается без новой страницы")
        if para.text.strip().endswith("."):
            problems.append("точка после обозначения приложения")
        if problems:
            ctx.add("APPENDIX", f"Приложение {letter}: " + "; ".join(problems) + ".",
                    paragraph=para.index)
    letters = [letter for _, letter in found]
    if letters != [l for l in allowed if l in letters]:
        ctx.add("APPENDIX",
                "Приложения идут не по порядку букв: " + " → ".join(letters) + ".",
                paragraph=found[0][0].index)
    body_text = " ".join(p.text for p in ctx.model.paragraphs
                         if not p.in_table and p not in [f[0] for f in found])
    for para, letter in found:
        if not re.search(rf"приложен\w*\s*№?\s*{re.escape(letter)}\b", body_text, re.IGNORECASE):
            ctx.add("APPENDIX", f"На приложение {letter} нет ссылки в тексте.",
                    paragraph=para.index, severity="info")


# --- Объём и «проверить вручную» ------------------------------------------
def check_volume(ctx: CheckContext) -> None:
    """VOLUME_FILL: эвристика заполненности страниц (info)."""
    volume = ctx.rules.get("volume") or {}
    min_fill = float(volume.get("min_page_fill_percent", 0) or 0)
    if not min_fill or not ctx.model.paragraphs:
        return
    total = ctx.pages.total_pages
    fill = ctx.pages.min_fill_excluding_last()
    ctx.add("VOLUME_FILL",
            f"Оценка объёма: ~{total} стр., минимальная заполненность страницы ~{fill:.0f} % "
            f"(правило: ≥ {min_fill:.0f} %). Пересчёт условный — зависит от шрифта и таблиц.",
            severity="info")


def check_manual(ctx: CheckContext) -> None:
    """MANUAL: чек-лист того, что нельзя проверить автоматически (п. 9.1 ТЗ)."""
    for title, howto in MANUAL_ITEMS:
        ctx.add("MANUAL", f"Проверить вручную: {title}.", hint=howto, severity="info")


CHECKS_CONTENT: Sequence[Callable[[CheckContext], None]] = (
    check_figures,
    check_tables,
    check_formulas,
    check_footnotes,
    check_toc,
    check_bibliography,
    check_appendices,
    check_volume,
    check_manual,
)
