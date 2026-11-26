"""Правки `word/document.xml`: текст, заголовки, подписи, таблицы, сноски.

Нормализатор правит ПРЯМОЕ форматирование только там, где оно конфликтует с правилами
(требование ТЗ): сравнение идёт со значениями из модели инспектора, поэтому на уже
корректном документе ни один атрибут не переустанавливается — на этом держится
идемпотентность (два прогона = байт-в-байт одинаковый XML).

Что сознательно НЕ делается: содержимое не сочиняется. Отсутствующие ссылки «(рис. 1)»,
устаревшие источники, незаполненные страницы — содержательные дефекты, их исправляет
генерирующий скрипт, а не пост-обработка.
"""

from __future__ import annotations

from typing import Any, Mapping

import xml.etree.ElementTree as ET

from app import constants
from app.core import docx_parser as dp
from app.core import docx_xml as dx
from app.core.docx_norm_util import (
    JC_TO_OOXML,
    NormCtx,
    TextSpec,
    ensure_bold,
    ensure_flag,
    ensure_fonts,
    ensure_indent,
    ensure_jc,
    ensure_size,
    ensure_spacing,
    paragraph_ppr,
    run_rpr,
    set_attr,
    text_spec,
)
from app.core.units import half_points_to_pt, pt_to_half_points

__all__ = [
    "normalize_run_fonts",
    "normalize_body_paragraphs",
    "normalize_structural_headings",
    "normalize_captions",
    "normalize_heading_styles",
    "normalize_tables",
    "normalize_footnotes",
]


def is_mono_run(run: dp.RunInfo, model: dp.DocxModel) -> bool:
    """Защищённый run: моноширинный шрифт из `constants.MONOSPACE_FONTS` не трогаем —
    это код-блок, его оформление задаёт правило code_blocks, а не правило текста."""
    font = str(run.props.font or "").lower()
    return font in {f.lower() for f in model.mono_fonts} or font in {
        f.lower() for f in constants.MONOSPACE_FONTS}


def _fix_run_props(rpr: ET.Element, t: TextSpec, notes: list[str], *,
                   size: int | None = None, bold: bool | None = None) -> None:
    """Шрифт (и, если задано, кегль/полужирание) одного run'а — с отчётом в `notes`."""
    if ensure_fonts(rpr, t.font):
        notes.append(f"шрифт run → {t.font}")
    if size is not None and ensure_size(rpr, size):
        notes.append(f"кегль run → {half_points_to_pt(size):g} pt")
    if bold is not None and ensure_bold(rpr, bold):
        notes.append("полужирный" if bold else "обычный")


def _para_runs(ctx: NormCtx, para: dp.ParagraphInfo, *, create: bool = False) -> list[ET.Element]:
    """`w:rPr` значимых run'ов абзаца.

    Пропускаем: моноширинные (код-блок — правило code_blocks), пустые по тексту (run'ы
    поля PAGE/TOC, рисунки `w:drawing`, разрывы) и, при `create=False`, run'ы без `w:rPr`.
    Создавать rPr у run'а без текста нельзя: он и так наследует свойства, а лишние
    элементы меняли бы байты корректного файла при каждом прогоне.
    """
    out: list[ET.Element] = []
    for run in para.runs:
        if is_mono_run(run, ctx.model) or not run.text.strip():
            continue
        rpr = run_rpr(run.element, create=create)
        if rpr is not None:
            out.append(rpr)
    return out


def _strip_trailing_dot(para: dp.ParagraphInfo, notes: list[str]) -> None:
    """Точка в конце заголовка/подписи — нарушение; снимаем с последнего непустого w:t."""
    for run in reversed(para.runs):
        done = False
        for t_el in reversed(dx.find_all(run.element, "w:t")):
            text = (t_el.text or "").rstrip()
            if not text:
                continue
            if text.endswith("."):
                t_el.text = text[:-1].rstrip()
                notes.append("убрана точка в конце")
            done = True
            break
        if done:
            break


def _upper_runs(para: dp.ParagraphInfo, notes: list[str]) -> None:
    """Прописной регистр структурного элемента — правим текст в `w:t`, границы run'ов не трогаем."""
    for run in para.runs:
        for t_el in dx.find_all(run.element, "w:t"):
            text = t_el.text or ""
            if text and text != text.upper():
                t_el.text = text.upper()
                notes.append("регистр → прописные")


def _first_content_paragraph(model: dp.DocxModel) -> int:
    """Индекс первого непустого абзаца: ему pageBreakBefore не ставим (иначе — пустая страница)."""
    for para in model.paragraphs:
        if not para.is_empty:
            return para.index
    return -1


# --- Основной текст ---------------------------------------------------------
def normalize_run_fonts(ctx: NormCtx) -> None:
    """Явный `w:rFonts` во всех не-моно run'ах → шрифт из правила.

    Без этого правка стиля бесполезна: Arial, заданный прямым форматированием, перебивает
    стиль (именно так старый чекер «не видел» подмену шрифта — п. 11.2 №1 ТЗ).
    """
    t = text_spec(ctx.rules)
    if not t.font:
        return
    cells = {p.index for p in ctx.model.paragraphs if p.in_table}
    touched = 0
    for para in ctx.model.paragraphs:
        if para.is_empty or para.instr or para.index in cells:
            continue  # ячейки ведёт normalize_tables; текст поля (TOC/PAGE) перезапишет Word
        changed = False
        for rpr in _para_runs(ctx, para, create=True):
            changed = ensure_fonts(rpr, t.font) or changed
        touched += 1 if changed else 0
    if touched:
        ctx.note(f"document.xml: шрифт прямо форматированных run'ов ({touched} абзацев) "
                 f"приведён к «{t.font}»")


def _looks_like_formula(text: str) -> bool:
    """Формула: короткая строка со «=» без точки в конце (как в docx_content_checks).

    Нормализатор не трогает такие абзацы: формула стоит по центру с номером у правого
    края, и «приведение к тексту» (justify + firstLine) испортило бы её оформление.
    """
    clean = text.strip()
    return bool(clean) and "=" in clean and len(clean) <= 120 and not clean.endswith(".")


def normalize_body_paragraphs(ctx: NormCtx) -> None:
    """Абзацы основного текста: интервал, отступ, выравнивание, кегль — строго по правилу."""
    rules = ctx.rules
    if not rules.get("text"):
        return
    t = text_spec(rules)
    counts: dict[str, int] = {}
    for para in dp.main_paragraphs(ctx.model, rules):
        if _looks_like_formula(para.text):
            continue
        ppr = paragraph_ppr(para.element, ctx)
        if ppr is None:
            continue
        if rget(rules, "text.line_spacing") is not None and \
                ensure_spacing(ppr, line=t.line, before=t.before, after=t.after):
            counts["межстрочный интервал"] = counts.get("межстрочный интервал", 0) + 1
        if "first_line_indent_cm" in (rules.get("text") or {}) and ensure_indent(ppr, t.indent):
            counts["абзацный отступ"] = counts.get("абзацный отступ", 0) + 1
        if rget(rules, "text.alignment") and ensure_jc(ppr, t.jc):
            counts["выравнивание"] = counts.get("выравнивание", 0) + 1
        if rget(rules, "text.size_pt") is not None:
            for rpr in _para_runs(ctx, para):
                if ensure_size(rpr, t.size_half):
                    counts["кегль"] = counts.get("кегль", 0) + 1
    for label, count in counts.items():
        ctx.note(f"Основной текст: {label} → по правилу ({count} абзац(ов)); "
                 f"интервал {t.line / 240:g}, отступ {t.indent} twips, кегль "
                 f"{half_points_to_pt(t.size_half):g} pt")


def rget(rules: Mapping[str, Any], path: str, default: Any = None) -> Any:
    """Краткая обёртка над `dp.rget` (единый способ чтения вложенного ключа правил)."""
    return dp.rget(rules, path, default)


# --- Заголовки --------------------------------------------------------------
def normalize_structural_headings(ctx: NormCtx) -> None:
    """РЕФЕРАТ/СОДЕРЖАНИЕ/ВВЕДЕНИЕ/…: прописные, полужирные, по центру, с новой страницы."""
    headings = ctx.rules.get("headings") or {}
    if not headings:
        return
    t = text_spec(ctx.rules)
    want_upper = bool(headings.get("structural_uppercase", True))
    want_bold = bool(headings.get("structural_bold", True))
    want_jc = JC_TO_OOXML.get(str(headings.get("structural_align", "center")).lower(), "center")
    want_new_page = bool(headings.get("structural_new_page", True))
    want_no_dot = bool(headings.get("structural_no_trailing_dot", True))
    first_index = _first_content_paragraph(ctx.model)

    for para in ctx.model.paragraphs:
        if para.in_table or not para.text.strip():
            continue
        if not dp.is_structural_heading(para.text, ctx.rules):
            continue
        notes: list[str] = []
        if want_upper:
            _upper_runs(para, notes)
        if want_no_dot:
            _strip_trailing_dot(para, notes)

        ppr = paragraph_ppr(para.element, ctx)
        if ppr is not None:
            if ensure_jc(ppr, want_jc):
                notes.append(f"выравнивание → «{want_jc}»")
            # Центровой заголовок с firstLine уезжает вправо — отступ обнуляем.
            if want_jc in ("center", "right") and ensure_indent(ppr, 0):
                notes.append("абзацный отступ → 0")
            if ensure_flag(ppr, "w:keepNext", bool(headings.get("keep_with_next", True))):
                notes.append("keepNext")
            if want_new_page and para.index != first_index and \
                    ensure_flag(ppr, "w:pageBreakBefore", True):
                notes.append("новая страница (pageBreakBefore)")

        for rpr in _para_runs(ctx, para, create=True):
            _fix_run_props(rpr, t, notes, size=t.size_half, bold=True if want_bold else None)

        if notes:
            ctx.note(f"Структурный заголовок «{dp.snippet_of(para.text, 40)}»: "
                     + "; ".join(dict.fromkeys(notes)))


def normalize_heading_styles(ctx: NormCtx) -> None:
    """Заголовки разделов → `w:pStyle HeadingN` + keepNext + кегль не больше основного.

    Стиль навешивается только при совпадении профиля: короткий, полужирный, без полей,
    не в таблице. Иначе под стиль попали бы обычные фразы вида «1 вариант ответа».
    """
    headings = ctx.rules.get("headings") or {}
    if not headings:
        return
    t = text_spec(ctx.rules)
    for para in ctx.model.paragraphs:
        if para.in_table or para.is_empty or para.instr:
            continue
        if dp.is_structural_heading(para.text, ctx.rules):
            continue  # структурные уже обработаны отдельной функцией
        if not dp.looks_like_heading(para, ctx.rules):
            continue
        level = dp.heading_level(para.text, para)
        bold_ratio = _bold_ratio(para)
        notes: list[str] = []
        ppr = paragraph_ppr(para.element, ctx)
        if ppr is None:
            continue
        if level and bold_ratio >= 0.6 and len(para.text.strip()) <= 160:
            style = dx.find_child(ppr, "w:pStyle")
            want_style = f"Heading{level}"
            if style is None or (style.get(dx.qn("w:val")) or "") != want_style:
                dx.ensure_child(ppr, "w:pStyle", replace=True, attrs={"w:val": want_style})
                notes.append(f"стиль → {want_style}")
        if bool(headings.get("keep_with_next", True)) and ensure_flag(ppr, "w:keepNext", True):
            notes.append("keepNext")
        if bool(headings.get("no_trailing_dot", True)):
            _strip_trailing_dot(para, notes)
        if bool(headings.get("section_bold", True)):
            for rpr in _para_runs(ctx, para, create=True):
                if ensure_bold(rpr, True):
                    notes.append("полужирный")
        if bool(headings.get("same_size_as_body", True)):
            for rpr in _para_runs(ctx, para):
                if ensure_size(rpr, t.size_half):
                    notes.append(f"кегль → {half_points_to_pt(t.size_half):g} pt")
        if notes:
            ctx.note(f"Заголовок «{dp.snippet_of(para.text, 40)}»: "
                     + "; ".join(dict.fromkeys(notes)))


def _bold_ratio(para: dp.ParagraphInfo) -> float:
    total = sum(len(r.text.strip()) for r in para.runs)
    if not total:
        return 1.0 if para.eff_run.bold else 0.0
    return sum(len(r.text.strip()) for r in para.runs if r.props.bold) / total


# --- Подписи ---------------------------------------------------------------
def normalize_captions(ctx: NormCtx) -> None:
    """«Таблица N – …» — слева, без отступа, с keepNext; «Рисунок N – …» — по центру, без отступа."""
    tables = ctx.rules.get("tables") or {}
    figures = ctx.rules.get("figures") or {}
    regexes = dp.caption_regexes(ctx.rules)
    if not regexes:
        return

    for para in ctx.model.paragraphs:
        if para.in_table or not para.text.strip():
            continue
        kind = dp.is_caption(para.text, regexes)
        if kind == "table":
            want_jc = JC_TO_OOXML.get(str(tables.get("caption_align", "left")).lower(), "left")
            keep, no_dot = True, bool(tables.get("caption_no_trailing_dot", True))
        elif kind == "continuation":
            want_jc = JC_TO_OOXML.get(str(tables.get("continuation_align", "right")).lower(),
                                      "right")
            keep, no_dot = False, False
        elif kind == "figure":
            want_jc = JC_TO_OOXML.get(str(figures.get("caption_align", "center")).lower(),
                                      "center")
            keep, no_dot = False, bool(figures.get("caption_no_trailing_dot", True))
        else:
            continue
        notes: list[str] = []
        ppr = paragraph_ppr(para.element, ctx)
        if ppr is None:
            continue
        if ensure_jc(ppr, want_jc):
            notes.append(f"выравнивание → «{want_jc}»")
        if ensure_indent(ppr, 0):
            notes.append("абзацный отступ → 0")
        if keep and ensure_flag(ppr, "w:keepNext", True):
            notes.append("keepNext (название не отрывается от таблицы)")
        if no_dot:
            _strip_trailing_dot(para, notes)
        if notes:
            ctx.note(f"Подпись «{dp.snippet_of(para.text, 40)}»: " + "; ".join(dict.fromkeys(notes)))

    # Рисунок «держит» подпись: keepNext ставится абзацу с изображением, если подпись под ним.
    if str(figures.get("caption_position", "below")).lower() != "below":
        return
    rx = regexes.get("figure")
    if rx is None:
        return
    by_index = {p.index: p for p in ctx.model.paragraphs}
    for img in ctx.model.images:
        holder = by_index.get(img.paragraph_index)
        if holder is None or holder.in_table:
            continue
        nxt = by_index.get(holder.index + 1)
        if nxt is None or nxt.in_table:
            continue
        if not rx.match(" ".join(nxt.text.split()).strip()):
            continue
        ppr = paragraph_ppr(holder.element, ctx)
        if ppr is not None and ensure_flag(ppr, "w:keepNext", True):
            ctx.note(f"Абзац {holder.index + 1} с рисунком: keepNext (подпись не отрывается)")


# --- Таблицы ---------------------------------------------------------------
def normalize_tables(ctx: NormCtx) -> None:
    """Границы, повтор шапки, жирная шапка, фиксированный макет, vAlign, кегль ячеек.

    `w:cantSplit` НЕ ставится: правило tables.allow_row_split означает, что высокую строку
    можно разрывать между страницами. Найденный cantSplit при allow_row_split убираем.
    """
    tables = ctx.rules.get("tables") or {}
    if not tables:
        return
    t = text_spec(ctx.rules)
    cell_half = pt_to_half_points(float(tables["font_size_pt"])) if tables.get(
        "font_size_pt") else None
    cell_line = tables.get("line_spacing")
    want_valign = {"center": "center", "top": "top", "bottom": "bottom"}.get(
        str(tables.get("cell_vertical_align", "center")).lower())
    by_table: dict[int, list[dp.ParagraphInfo]] = {}
    for para in ctx.model.paragraphs:
        if para.table_index is not None:
            by_table.setdefault(para.table_index, []).append(para)

    for table in ctx.model.tables:
        notes: list[str] = []
        tbl_pr = dx.find_child(table.element, "w:tblPr")
        if tbl_pr is None:
            tbl_pr = dx.ensure_child(table.element, "w:tblPr", index=0, replace=False)
        if str(tables.get("borders", "single_thin")) == "single_thin" and \
                _ensure_borders(tbl_pr, int(tables.get("border_size_eighths_pt", 4))):
            notes.append("границы single по всем сторонам")
        if str(tables.get("layout", "")).lower() == "fixed":
            layout = dx.find_child(tbl_pr, "w:tblLayout")
            if layout is None:
                layout = dx.ensure_child(tbl_pr, "w:tblLayout", replace=True)
            if set_attr(layout, "w:type", "fixed"):
                notes.append("макет таблицы — фиксированный")

        if tables.get("repeat_header", True) and table.rows:
            trpr = dx.find_child(table.rows[0].element, "w:trPr")
            if trpr is None:
                trpr = dx.ensure_child(table.rows[0].element, "w:trPr", index=0,
                                       replace=False)
            if ensure_flag(trpr, "w:tblHeader", True):
                notes.append("первая строка — заголовок (tblHeader)")
        if tables.get("allow_row_split", True):
            removed = 0
            for row in table.rows:
                row_trpr = dx.find_child(row.element, "w:trPr")
                if row_trpr is not None:
                    removed += dx.remove_children(row_trpr, "w:cantSplit")
            if removed:
                notes.append(f"убрано w:cantSplit ({removed} строк)")

        head_indexes = _head_paragraph_indexes(ctx, table)
        for para in by_table.get(table.index, []):
            ppr = paragraph_ppr(para.element, ctx)
            if ppr is not None and cell_line is not None and \
                    ensure_spacing(ppr, line=int(round(float(cell_line) * 240))):
                notes.append(f"интервал в ячейках → {float(cell_line):g}")
            is_head = para.index in head_indexes
            for rpr in _para_runs(ctx, para, create=True):
                if is_head and tables.get("header_bold", True) and ensure_bold(rpr, True):
                    notes.append("шапка полужирная")
                if cell_half is not None and ensure_size(rpr, cell_half):
                    notes.append(f"кегль в ячейках → {half_points_to_pt(cell_half):g} pt")
                if rget(ctx.rules, "text.font") and ensure_fonts(rpr, t.font):
                    notes.append(f"шрифт в ячейках → {t.font}")
        if want_valign:
            moved = 0
            for tc in dx.find_all(table.element, "w:tc"):
                tcpr = dx.find_child(tc, "w:tcPr")
                if tcpr is None:
                    tcpr = dx.ensure_child(tc, "w:tcPr", index=0, replace=False)
                el = dx.find_child(tcpr, "w:vAlign")
                if el is None:
                    el = dx.ensure_child(tcpr, "w:vAlign", replace=True)
                if set_attr(el, "w:val", want_valign):
                    moved += 1
            if moved:
                notes.append(f"вертикальное выравнивание ячеек → «{want_valign}» ({moved})")
        if notes:
            ctx.note(f"Таблица {table.index + 1}: " + "; ".join(dict.fromkeys(notes)))


def _head_paragraph_indexes(ctx: NormCtx, table: dp.TableInfo) -> set[int]:
    """Индексы абзацев первой (шапки) строки таблицы."""
    if not table.rows:
        return set()
    by_id = {id(p.element): p.index for p in ctx.model.paragraphs}
    return {by_id[id(el)] for el in table.rows[0].element.iter(dx.qn("w:p")) if id(el) in by_id}


def _ensure_borders(tbl_pr: ET.Element, size_eighths: int) -> bool:
    """w:tblBorders: шесть сторон, val=single, sz в 1/8 пункта, цвет — чёрный."""
    needed = ("top", "left", "bottom", "right", "insideH", "insideV")
    borders = dx.find_child(tbl_pr, "w:tblBorders")
    if borders is not None:
        good = True
        for side in needed:
            el = dx.find_child(borders, f"w:{side}")
            if el is None or (el.get(dx.qn("w:val")) or "").lower() != "single" or \
                    (dx.get_int(el, "w:sz") or 0) > max(size_eighths, 4):
                good = False
                break
        if good:
            return False
        tbl_pr.remove(borders)
    borders = dx.ensure_child(tbl_pr, "w:tblBorders", replace=False)
    for side in needed:
        el = ET.SubElement(borders, dx.qn(f"w:{side}"))
        el.set(dx.qn("w:val"), "single")
        el.set(dx.qn("w:sz"), str(max(size_eighths, 4)))
        el.set(dx.qn("w:space"), "0")
        el.set(dx.qn("w:color"), "000000")
    return True


# --- Сноски ----------------------------------------------------------------
def normalize_footnotes(ctx: NormCtx) -> None:
    """`word/footnotes.xml`: кегль сносок в допустимый диапазон правила и нужный шрифт.

    Надстрочная номерная run остаётся надстрочной — правим только rFonts/sz, структуру
    ссылки (`w:footnoteReference`) не трогаем.
    """
    notes = ctx.rules.get("footnotes") or {}
    root = ctx.model.footnotes
    if not notes or root is None:
        return
    want = pt_to_half_points(float(notes.get("size_pt", 12)))
    low = pt_to_half_points(float(notes.get("min_allowed_size_pt",
                                            notes.get("size_pt", 12))))
    font = str(notes.get("font", "") or "")
    mono = {f.lower() for f in constants.MONOSPACE_FONTS}
    changed_size = False
    changed_font = False
    for note in dx.find_all(root, "w:footnote"):
        if note.get(dx.qn("w:type")) in ("separator", "continuationSeparator"):
            continue
        # run'ы лежат внутри w:p, а не непосредственно в w:footnote (CT_FtnRecContent)
        for para in dx.find_all(note, "w:p"):
            for r_el in dx.find_all(para, "w:r"):
                rpr = dx.find_child(r_el, "w:rPr")
                if rpr is None:
                    continue
                fonts = dx.find_child(rpr, "w:rFonts")
                if fonts is not None and (fonts.get(dx.qn("w:ascii")) or "").lower() in mono:
                    continue
                size = dx.get_int(dx.find_child(rpr, "w:sz"), "w:val")
                if size is not None and not (low <= size <= want):
                    changed_size = ensure_size(rpr, want) or changed_size
                if font and fonts is not None:
                    changed_font = set_attr(fonts, "w:ascii", font) or changed_font
                    changed_font = set_attr(fonts, "w:hAnsi", font) or changed_font
    if changed_size or changed_font:
        ctx.note("Сноски: " + "; ".join(filter(None, [
            f"кегль → {half_points_to_pt(want):g} pt" if changed_size else "",
            f"шрифт → {font}" if changed_font else ""])))
