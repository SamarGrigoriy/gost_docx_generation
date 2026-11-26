"""Проверки OOXML по АКТИВНЫМ правилам: страница, текст, нумерация, заголовки.

Все коды — из таблицы §9.1 ТЗ. Значения берутся ТОЛЬКО из `rules`, зашитых констант
нет (п. 11.2 ТЗ: старый чекер сверял с хардкодом 12 pt / интервал 1,0).

Общие механизмы:
* `CheckContext` считает `location` (номер абзаца + оценка страницы + фрагмент ≤120 симв.);
* severity берётся из `SEVERITIES_BY_CODE`, при `strict: false` error понижается до warning;
* проверка снимается с контроля, если соответствующего ключа правил нет или он выключен
  (`toc.required: false`, `references.max_source_age_years: 0`,
  `figures.require_reference_in_text: false` и т.п.).
"""

from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from typing import Any, Callable, Mapping, Sequence

import xml.etree.ElementTree as ET

from app import constants
from app.core import docx_parser as dp
from app.core import docx_xml as dx
from app.core.docx_parser import DocxModel, ParagraphInfo
from app.core.units import (
    cm_to_twips,
    half_points_to_pt,
    mm_to_twips,
    page_size,
    pt_to_half_points,
    spacing_to_line,
    twips_to_mm,
)

# --- Реестр кодов и базовых severity (п. 9.1 ТЗ) ---------------------------
SEVERITY_BY_CODE: dict[str, str] = {
    "PAGE_MARGINS": "error",
    "PAGE_SIZE": "error",
    "FONT_MAIN": "error",
    "SIZE_MAIN": "error",
    "LINE_SPACING": "error",
    "INDENT_VALUE": "error",
    "ALIGNMENT": "warning",
    "SPACING_PARA": "warning",
    "NUM_POSITION": "error",
    "NUM_START": "error",
    "HEAD_STRUCT": "error",
    "HEAD_KEEP": "warning",
    "HEAD_SIZE": "warning",
    "FIG_CAPTION": "error",
    "FIG_COUNT": "error",
    "FIG_ORDER": "warning",
    "FIG_REF": "warning",
    "TBL_CAPTION": "error",
    "TBL_HEADER": "warning",
    "TBL_BORDERS": "warning",
    "TBL_ORDER": "warning",
    "TBL_REF": "warning",
    "FORMULA": "warning",
    "FOOTNOTE_FONT": "warning",
    "TOC_PRESENT": "error",
    "BIB_ORDER": "warning",
    "BIB_AGE": "warning",
    "BIB_COVERAGE": "warning",
    "APPENDIX": "warning",
    "VOLUME_FILL": "info",
    "MANUAL": "info",
}

# Подсказки «как исправить» — статичная часть; проверка может передать свою.
HINTS: dict[str, str] = {
    "PAGE_MARGINS": "Задайте поля в макете страницы: левое 30 мм, правое 15 мм, верхнее и "
                    "нижнее 20 мм (или измените пресет).",
    "PAGE_SIZE": "Формат A4, ориентация по пресету: макет страницы → размер → A4.",
    "FONT_MAIN": "Выделите основной текст и задайте шрифт из правил (вручную, а не только стилем).",
    "SIZE_MAIN": "Кегль основного текста должен соответствовать правилу text.size_pt.",
    "LINE_SPACING": "Абзац → межстрочный интервал = правилу (1,5 для ГОСТ), в абзацах текста.",
    "INDENT_VALUE": "Абзацный отступ = правилу (1,25 см = 709 twips), а не «на глаз».",
    "ALIGNMENT": "Основной текст — по ширине (w:jc=both).",
    "SPACING_PARA": "Интервалы до и после абзаца = 0 (абзац → отступы и интервалы).",
    "NUM_POSITION": "Номер страницы — внизу по центру: колонтитул → нижний, выравнивание по центру.",
    "NUM_START": "Одна секция +Different First Page (titlePg) без номера; не ставьте pgNumType start.",
    "HEAD_STRUCT": "Структурные элементы — ПРОПИСНЫМИ, полужирные, по центру, с новой страницы, "
                   "без точки.",
    "HEAD_KEEP": "У заголовка включите «не отрывать от следующего» (keepNext).",
    "HEAD_SIZE": "Кегль заголовка не должен превышать кегль основного текста.",
    "FIG_CAPTION": "Подпись под рисунком по центру: «Рисунок N – Название», без точки.",
    "FIG_COUNT": "У каждого изображения должна быть подпись «Рисунок N – …».",
    "FIG_ORDER": "Нумерация рисунков сквозная, без пропусков и повторов.",
    "FIG_REF": "Добавьте ссылку в тексте: «в соответствии с рисунком N».",
    "TBL_CAPTION": "Название над таблицей, СЛЕВА и без абзацного отступа: «Таблица N – …».",
    "TBL_HEADER": "Шапку таблицы повторите на каждой странице (свойство строки «заголовок») и "
                  "сделайте полужирной.",
    "TBL_BORDERS": "Границы таблицы — тонкие сплошные (0,5 pt) со всех сторон.",
    "TBL_ORDER": "Номера таблиц должны идти по порядку появления.",
    "TBL_REF": "Добавьте ссылку в тексте: «в таблице N».",
    "FORMULA": "Формула — отдельной строкой по центру, номер (N) у правого края; «где» без "
               "двоеточия.",
    "FOOTNOTE_FONT": "Текст сноски — TNR 12 pt (допускается 10 pt).",
    "TOC_PRESENT": "Вставьте автосодержание: ссылки → оглавление (поле TOC), затем обновите поля.",
    "BIB_ORDER": "Список источников — в порядке появления ссылок в тексте.",
    "BIB_AGE": "Замените источники старше предельного возраста (правило references)."
               "",
    "BIB_COVERAGE": "Проставьте затекстовые ссылки на все (≥ минимального процента) источники.",
    "APPENDIX": "Приложения — буквы кириллицы кроме Ё, З, Й, О, Ч, Ъ, Ы, Ь; каждое с новой "
                "страницы, по центру.",
    "VOLUME_FILL": "Страница должна быть заполнена не менее чем на 80 % (последняя в разделе "
                   "— от 1/3).",
    "MANUAL": "Проверить вручную.",
}

# Метрические аналоги шрифтов: если документ использует их вместо TNR — это не Arial.
FONT_ALIASES: dict[str, set[str]] = {
    "times new roman": {"times new roman", "liberation serif", "nimbus roman",
                        "thorndale amt", "thorndale amr", "tinos", "serif"},
    "arial": {"arial", "liberation sans", "helvetica", "dejavu sans", "arimo", "sans-serif"},
    "calibri": {"calibri", "carlito"},
    "pt sans": {"pt sans", "pt sans caption"},
}

# Шаблоны полей колонтитулов: позиция → требуемое выравнивание.
_ALIGN_BY_POSITION = {
    "bottom_center": "center", "top_center": "center", "center": "center",
    "bottom_right": "right", "top_right": "right", "right": "right",
    "bottom_left": "left", "top_left": "left", "left": "left",
}


@dataclass
class Issue:
    """Замечание инспектора: код, severity, место (абзац/страница/фрагмент) и подсказка."""

    code: str
    severity: str
    message: str
    location: dict[str, Any] = field(default_factory=dict)
    hint: str = ""

    def to_dict(self) -> dict[str, Any]:
        """Для JSON-конверта CLI (п. 19.4: issues[{code,severity,message,location,hint}])."""
        return {"code": self.code, "severity": self.severity, "message": self.message,
                "location": dict(self.location), "hint": self.hint}


class CheckContext:
    """Контекст проверки: модель, правила, счётчик страниц, накопленные замечания."""

    def __init__(self, model: DocxModel, rules: Mapping[str, Any]) -> None:
        self.model = model
        self.rules = rules
        self.strict = bool(rules.get("strict", True))
        self.regexes = dp.caption_regexes(rules)
        self.pages = dp.estimate_pages(model, rules)
        self.issues: list[Issue] = []
        self._main: list[ParagraphInfo] | None = None
        self._title: tuple[int | None] | None = None

    # --- утилиты ---------------------------------------------------------
    @property
    def first_structural(self) -> int | None:
        """Индекс начала основного текста: первый структурный элемент либо конец титула.

        Если структурного элемента нет (документ без РЕФЕРАТА), границей служит абзац
        после первой принудительной страницы — строки титульного листа проверяются
        по шаблону вуза, а не как заголовки разделов.
        """
        if self._title is None:
            boundary = dp.first_structural_index(self.model, self.rules)
            if boundary is None:
                boundary = next(
                    (p.index for p in self.model.paragraphs
                     if not p.in_table
                     and (p.has_page_break_run or p.props.page_break_before
                          or p.eff.page_break_before)),
                    None)
            self._title = (boundary,)
        return self._title[0]

    # --- утилиты ---------------------------------------------------------
    def add(self, code: str, message: str, *, paragraph: int | None = None,
            hint: str | None = None, severity: str | None = None,
            snippet: str | None = None) -> None:
        """Записать замечание с поправкой на строгий режим."""
        base = severity or SEVERITY_BY_CODE.get(code, "warning")
        severity_final = "warning" if (base == "error" and not self.strict) else base
        location: dict[str, Any] = {"paragraph": paragraph,
                                   "page": self.pages.page_of.get(paragraph) if paragraph
                                   is not None else None,
                                   "snippet": snippet if snippet is not None
                                   else (self.text_at(paragraph) if paragraph is not None else "")}
        self.issues.append(Issue(code=code, severity=severity_final, message=message,
                                 location=location,
                                 hint=hint or HINTS.get(code, "Проверьте по методичке.")))

    def text_at(self, paragraph: int | None) -> str:
        """Фрагмент текста абзаца (≤120 символов) для колонки «Где»."""
        if paragraph is None:
            return ""
        for para in self.model.paragraphs:
            if para.index == paragraph:
                return dp.snippet_of(para.text)
        return ""

    def para(self, index: int | None) -> ParagraphInfo | None:
        for para in self.model.paragraphs:
            if para.index == index:
                return para
        return None

    @property
    def main(self) -> list[ParagraphInfo]:
        """Абзацы основного текста (вне таблиц/кода/подписей/заголовков) — считаются один раз."""
        if self._main is None:
            self._main = dp.main_paragraphs(self.model, self.rules, self.regexes)
        return self._main

    @property
    def section(self) -> dp.SectionInfo | None:
        return self.model.main_section


# --- 1. Страница -----------------------------------------------------------
def check_page_size(ctx: CheckContext) -> None:
    """PAGE_SIZE: формат и ориентация из `page.size`/`page.orientation`."""
    page = ctx.rules.get("page") or {}
    if "size" not in page and "orientation" not in page:
        return
    want_w, want_h = page_size(str(page.get("size", "A4")), str(page.get("orientation", "portrait")))
    want_landscape = str(page.get("orientation", "portrait")) == "landscape"
    for sect in ctx.model.sections:
        if sect.pg_w is None and sect.pg_h is None:
            continue  # размер не задан — Word подставит значение по умолчанию, не наш случай
        ok = sect.pg_w == want_w and sect.pg_h == want_h
        if ok and want_landscape and sect.orient not in (None, "landscape"):
            ok = False
        if not ok:
            ctx.add("PAGE_SIZE",
                    f"Размер страницы {sect.pg_w}×{sect.pg_h} twips "
                    f"({twips_to_mm(sect.pg_w or 0):.0f}×{twips_to_mm(sect.pg_h or 0):.0f} мм), "
                    f"ожидается {want_w}×{want_h} "
                    f"({'альбомная' if want_landscape else 'книжная'} A4).")


def check_page_margins(ctx: CheckContext) -> None:
    """PAGE_MARGINS: поля секции против правил, допуск ±0,5 мм (п. 9.1 ТЗ)."""
    page = ctx.rules.get("page") or {}
    wanted = {
        "left": page.get("margin_left_mm"),
        "right": page.get("margin_right_mm"),
        "top": page.get("margin_top_mm"),
        "bottom": page.get("margin_bottom_mm"),
    }
    tol = mm_to_twips(constants.MARGIN_TOLERANCE_MM)
    for sect in ctx.model.sections:
        bad: list[str] = []
        for key, mm in wanted.items():
            if mm is None:
                continue
            actual = sect.mar.get(key)
            if actual is None:
                bad.append(f"{key}: не задано (ожидалось {mm} мм)")
                continue
            if abs(actual - mm_to_twips(float(mm))) > tol:
                bad.append(f"{key}: {twips_to_mm(actual):.1f} мм вместо {mm} мм")
        if bad:
            ctx.add("PAGE_MARGINS", "Поля страницы не соответствуют правилам: " + "; ".join(bad))


# --- 2. Основной текст (реальные rFonts/sz/spacing, вне таблиц и кода) -----
def check_font_main(ctx: CheckContext) -> None:
    """FONT_MAIN: фактический шрифт run'ов основного текста (п. 11.2 №1 — Arial ловится)."""
    expected = (ctx.rules.get("text") or {}).get("font")
    if not expected:
        return
    weights = [(r.props.font, len(r.text)) for para in ctx.main for r in para.runs if r.text.strip()]
    actual = dp.dominant_weighted(weights)
    if actual is None:
        return  # основного текста нет — не о чем сообщать
    # Совпадением считается и метрический аналог (Liberation Serif вместо TNR в LibreOffice),
    # но НЕ другой реальный шрифт: Arial обязан находить FONT_MAIN (п. 11.2 №1).
    expected_keys = FONT_ALIASES.get(str(expected).lower(), {str(expected).lower()})
    actual_keys = FONT_ALIASES.get(str(actual).lower(), {str(actual).lower()})
    if expected_keys & actual_keys:
        return
    first = next((p.index for p in ctx.main
                  if any(r.text.strip() and str(r.props.font or "").lower() == str(actual).lower()
                         for r in p.runs)), None)
    ctx.add("FONT_MAIN",
            f"Основной текст набран шрифтом «{actual}», правило требует «{expected}».",
            paragraph=first,
            hint=HINTS["FONT_MAIN"] + f" Сейчас доминирует «{actual}».")


def check_size_main(ctx: CheckContext) -> None:
    """SIZE_MAIN: доминирующий кегль основного текста (таблицы/код — отдельно)."""
    expected_pt = (ctx.rules.get("text") or {}).get("size_pt")
    if expected_pt is None:
        return
    expected = pt_to_half_points(float(expected_pt))
    weights = [(r.props.size, len(r.text)) for para in ctx.main for r in para.runs if r.text.strip()]
    actual = dp.dominant_weighted(weights)
    if actual is None or actual == expected:
        return
    first = next((p.index for p in ctx.main
                  if any(r.props.size == actual and r.text.strip() for r in p.runs)), None)
    ctx.add("SIZE_MAIN",
            f"Кегль основного текста {half_points_to_pt(int(actual)):g} pt, "
            f"правило требует {expected_pt:g} pt.", paragraph=first)


def check_line_spacing(ctx: CheckContext) -> None:
    """LINE_SPACING: w:line при lineRule=auto против правила text.line_spacing."""
    expected_spacing = (ctx.rules.get("text") or {}).get("line_spacing")
    if expected_spacing is None:
        return
    expected = spacing_to_line(float(expected_spacing))
    values = [p.eff.line for p in ctx.main if p.eff.line]
    actual = dp.dominant(values)
    if actual is None:
        return
    rule_auto = all((p.eff.line_rule or "auto") == "auto" for p in ctx.main if p.eff.line)
    if actual == expected and rule_auto:
        return
    first = next((p.index for p in ctx.main if p.eff.line == actual), None)
    detail = "" if rule_auto else " (интервал задан как точный, lineRule≠auto)"
    ctx.add("LINE_SPACING",
            f"Межстрочный интервал основного текста {dp.rget(ctx.rules, 'text.line_spacing')}: "
            f"w:line={actual} вместо {expected}{detail}.", paragraph=first)


def check_indent_value(ctx: CheckContext) -> None:
    """INDENT_VALUE: firstLine = правило ±5 twips (п. 11.2 №3 — проверяется ЗНАЧЕНИЕ)."""
    if "first_line_indent_cm" not in (ctx.rules.get("text") or {}):
        return
    target = cm_to_twips(float(ctx.rules["text"]["first_line_indent_cm"]))
    tol = constants.INDENT_TOLERANCE_TWIPS

    def indented(para: ParagraphInfo) -> bool:
        if abs((para.eff.first_line or 0) - target) <= tol:
            return True
        # Списки и библиография оформляются сдвигом влево (w:ind w:left) без firstLine —
        # это эквивалентный способ задать абзацный отступ по ГОСТ 7.32.
        return (para.eff.first_line or 0) == 0 and abs((para.eff.left or 0) - target) <= tol

    paras = [p for p in ctx.main if p.text.strip()]
    if not paras:
        return
    bad = [p for p in paras if not indented(p)]
    # Порог 1/3: точечные отступления (например, «приплывшие» подписи) не должны
    # превращаться в ошибку всего документа — нарушение систематическое или нет.
    if len(bad) < max(1, len(paras) // 3):
        return
    actual = dp.dominant([(p.eff.first_line or 0) for p in bad])
    ctx.add(
        "INDENT_VALUE",
        f"Абзацный отступ firstLine={actual} twips "
        f"({twips_to_mm(int(actual or 0)):.1f} мм) у {len(bad)} из {len(paras)} абзацев, "
        f"правило требует {target} twips "
        f"({float(ctx.rules['text']['first_line_indent_cm']):g} см).",
        paragraph=bad[0].index)


def check_alignment(ctx: CheckContext) -> None:
    """ALIGNMENT: выравнивание основного текста (warning по ТЗ)."""
    expected = (ctx.rules.get("text") or {}).get("alignment")
    if not expected:
        return
    want = dp.JC_ALIASES.get(str(expected), str(expected))
    values = [p.eff.jc for p in ctx.main if p.eff.jc]
    actual = dp.dominant(values)
    if actual is None or actual == want:
        return
    first = next((p.index for p in ctx.main if p.eff.jc == actual), None)
    ctx.add("ALIGNMENT", f"Основной текст выровнен «{actual}», правило требует «{want}».",
            paragraph=first)


def check_spacing_para(ctx: CheckContext) -> None:
    """SPACING_PARA: интервалы до/после абзаца = 0 (ГОСТ: только отступ первой строки)."""
    want_before = float((ctx.rules.get("text") or {}).get("space_before_pt", 0))
    want_after = float((ctx.rules.get("text") or {}).get("space_after_pt", 0))
    bad = [p for p in ctx.main
           if (p.eff.before or 0) > pt_to_half_points(want_before) * 10
           or (p.eff.after or 0) > pt_to_half_points(want_after) * 10]
    if not bad:
        return
    p0 = bad[0]
    ctx.add("SPACING_PARA",
            f"У {len(bad)} абзацев заданы интервалы до/после "
            f"(before={p0.eff.before}, after={p0.eff.after} twips), "
            f"правило требует {want_before:g}/{want_after:g} pt.",
            paragraph=p0.index)


# --- 3. Нумерация страниц --------------------------------------------------
def _part_has_page_field(root: ET.Element) -> bool:
    """Есть ли в части колонтитула поле PAGE (instrText или fldSimple/@w:instr)."""
    for instr in root.iter(dx.qn("w:instrText")):
        if re.search(r"\bPAGE\b", instr.text or "", re.IGNORECASE):
            return True
    for fld in root.iter(dx.qn("w:fldSimple")):
        if re.search(r"\bPAGE\b", fld.get(dx.qn("w:instr")) or "", re.IGNORECASE):
            return True
    return False


def _ftr_alignment(ctx: CheckContext, root: ET.Element) -> str | None:
    """Выравнивание абзаца, в котором лежит поле PAGE (effective jc с учётом стиля колонтитула)."""
    for p in dx.find_all(root, "w:p"):
        if not _part_has_page_field(p):
            continue
        direct = ctx.model.registry.raw_para_props(dx.find_child(p, "w:pPr"))
        jc = direct.get("jc") or (ctx.model.registry.para_props(direct.get("style")).get("jc"))
        return dp.JC_ALIASES.get(str(jc), jc) if jc else None
    return None


def check_numbering_position(ctx: CheckContext) -> None:
    """NUM_POSITION: номер страницы внизу по центру (footer + поле PAGE)."""
    numbering = ctx.rules.get("numbering") or {}
    position = str(numbering.get("position", "bottom_center"))
    if position in ("none", "", "off"):
        return
    want_align = _ALIGN_BY_POSITION.get(position, "center")
    sections = ctx.model.sections
    footer_parts = {name: root for name, root in ctx.model.footer_parts.items()}
    header_parts = {name: root for name, root in ctx.model.header_parts.items()}
    used_footers = {p for sect in sections for p in sect.footers.values()}
    used_headers = {p for sect in sections for p in sect.headers.values()}

    in_footer = [n for n in footer_parts if n in used_footers and _part_has_page_field(footer_parts[n])]
    in_any_footer = [n for n in footer_parts if _part_has_page_field(footer_parts[n])]
    in_header = [n for n in header_parts if n in used_headers and _part_has_page_field(header_parts[n])]

    if not in_footer:
        if in_header:
            ctx.add("NUM_POSITION",
                    "Номер страницы вставлен в верхний колонтитул "
                    f"({', '.join(sorted(in_header))}), а должен быть внизу по центру.",
                    hint=HINTS["NUM_POSITION"] + f" Правило: {position}.")
        elif in_any_footer:
            ctx.add("NUM_POSITION",
                    "Нижний колонтитул с полем PAGE существует, но не привязан к секции "
                    "(нет w:footerReference): номера печататься не будут.")
        elif not footer_parts:
            ctx.add("NUM_POSITION", "В документе нет нижнего колонтитула с номером страницы.")
        else:
            ctx.add("NUM_POSITION",
                    "В нижнем колонтитуле нет поля PAGE — номера страниц не выводятся.")
        return
    aligns = {(_ftr_alignment(ctx, footer_parts[n]) or "left") for n in in_footer}
    if want_align not in aligns:
        n0 = in_footer[0]
        ctx.add("NUM_POSITION",
                f"Номер страницы выровнен «{', '.join(sorted(aligns))}», требуется "
                f"«{want_align}» (правило numbering.position={position}).",
                hint=HINTS["NUM_POSITION"])


def check_numbering_start(ctx: CheckContext) -> None:
    """NUM_START: титул без номера, вторая страница показывает «2» (п. 5.1 ТЗ, §11.1 №1)."""
    numbering = ctx.rules.get("numbering") or {}
    if "title_page_numbered" not in numbering and "second_page_shows" not in numbering:
        return
    titled = bool(numbering.get("title_page_numbered", False))
    second = int(numbering.get("second_page_shows", 1 if titled else 2))
    sections = ctx.model.sections
    if not sections:
        return
    has_title_pg = any(sect.title_pg for sect in sections)

    if not titled and not has_title_pg:
        ctx.add("NUM_START",
                "Титульный лист не отделён от основной части (нет w:titlePg): номер стоит на "
                "первой странице, а вторая может показать «1» вместо «2».")
    # pgNumType start во 2-й и последующих секциях = обнуление нумерации на второй
    # физической странице (ровно дефект sect_body из п. 11.1 №1).
    for pos, sect in enumerate(sections):
        start = sect.pg_num_start
        if start is None:
            continue
        if pos > 0:
            ctx.add("NUM_START",
                    f"Нумерация перезапущена в секции {pos + 1} (w:pgNumType w:start=\"{start}\"): "
                    f"на второй физической странице будет «{start}» вместо {second}. "
                    f"Правильно — одна секция + titlePg.")
        elif not titled and start != max(1, second - 1):
            ctx.add("NUM_START",
                    f"w:pgNumType w:start=\"{start}\": вторая страница покажет {start + 1} "
                    f"вместо {second}.")
    if not titled and has_title_pg:
        # titlePg без ссылки на «первый» нижний колонтитул = номер на титуле всё равно виден
        for sect in sections:
            if sect.title_pg and "first" not in sect.footers and sect.footers:
                ctx.add("NUM_START",
                        "Режим titlePg включён, но пустой колонтитул первой страницы "
                        "(w:footerReference type=\"first\") не задан — номер на титуле останется.",
                        severity="warning")
                break


# --- 4. Заголовки ---------------------------------------------------------
def _bold_ratio(para: ParagraphInfo) -> float:
    """Доля жирных символов в абзаце: заголовок может набран полужирным не целиком."""
    total = sum(len(r.text.strip()) for r in para.runs)
    if not total:
        return 1.0 if para.eff_run.bold else 0.0
    bold = sum(len(r.text.strip()) for r in para.runs if r.props.bold)
    return bold / total


def _starts_new_page(ctx: CheckContext, para: ParagraphInfo) -> bool:
    """Начинается ли абзац с новой страницы: pageBreakBefore, явный w:br или конец секции выше."""
    if para.eff.page_break_before or para.props.page_break_before:
        return True
    for prev in reversed([p for p in ctx.model.paragraphs if p.index < para.index]):
        if prev.in_table:
            continue
        if prev.has_sect_break:  # разрыв секции с типом nextPage тоже даёт новую страницу
            return True
        return prev.has_page_break_run
    return True  # самый первый абзац документа


def check_structural_headings(ctx: CheckContext) -> None:
    """HEAD_STRUCT: элементы STRUCTURAL_HEADINGS — ПРОПИСНЫЕ, полужирные, центр, новая страница."""
    headings = ctx.rules.get("headings") or {}
    if not headings:
        return
    want_upper = bool(headings.get("structural_uppercase", True))
    want_bold = bool(headings.get("structural_bold", True))
    want_align = dp.JC_ALIASES.get(str(headings.get("structural_align", "center")), "center")
    want_new_page = bool(headings.get("structural_new_page", True))
    want_no_dot = bool(headings.get("structural_no_trailing_dot", True))
    for para in ctx.model.paragraphs:
        if para.in_table or para.is_toc or not para.text.strip():
            continue  # «СОДЕРЖАНИЕ 2» из автосодержания — не структурный элемент
        text = para.text.strip()
        if not dp.is_structural_heading(text, ctx.rules):
            continue
        problems: list[str] = []
        if want_upper and text != text.upper():
            problems.append("не в верхнем регистре")
        if want_bold and _bold_ratio(para) < 0.6:
            problems.append("не полужирный")
        if want_align and (para.eff.jc or "left") != want_align:
            problems.append(f"выравнивание «{para.eff.jc or 'left'}» вместо «{want_align}»")
        if want_new_page and not _starts_new_page(ctx, para):
            problems.append("без новой страницы (нет pageBreakBefore)")
        if want_no_dot and text.endswith("."):
            problems.append("точка в конце")
        if problems:
            ctx.add("HEAD_STRUCT",
                    f"Структурный элемент «{dp.snippet_of(text, 60)}»: " + "; ".join(problems) + ".",
                    paragraph=para.index)


def check_headings_style_and_size(ctx: CheckContext) -> None:
    """HEAD_KEEP + HEAD_SIZE: keepNext у заголовков, кегль не больше основного."""
    headings = ctx.rules.get("headings") or {}
    if not headings:
        return
    body_size = pt_to_half_points(float((ctx.rules.get("text") or {}).get("size_pt", 14)))
    want_keep = bool(headings.get("keep_with_next", True))
    want_same_size = bool(headings.get("same_size_as_body", True))
    title_end = ctx.first_structural
    for para in ctx.model.paragraphs:
        if para.is_empty or not dp.looks_like_heading(para, ctx.rules):
            continue
        # Строки титульного листа («ГБУ ВО …», название работы) жирные и короткие —
        # на них эвристика заголовка не распространяется: это не заголовки разделов.
        if title_end is not None and para.index < title_end:
            continue
        if want_keep and not (para.eff.keep_next or para.props.keep_next):
            ctx.add("HEAD_KEEP",
                    f"Заголовок «{dp.snippet_of(para.text, 60)}» без keepNext: может оторваться "
                    f"от текста в конце страницы.", paragraph=para.index)
        if want_same_size and para.eff_run.size and para.eff_run.size > body_size:
            ctx.add("HEAD_SIZE",
                    f"Кегль заголовка «{dp.snippet_of(para.text, 60)}» "
                    f"{half_points_to_pt(int(para.eff_run.size)):g} pt больше основного текста "
                    f"({half_points_to_pt(body_size):g} pt).", paragraph=para.index)


CHECKS_LAYOUT: Sequence[Callable[[CheckContext], None]] = (
    check_page_size,
    check_page_margins,
    check_font_main,
    check_size_main,
    check_line_spacing,
    check_indent_value,
    check_alignment,
    check_spacing_para,
    check_numbering_position,
    check_numbering_start,
    check_structural_headings,
    check_headings_style_and_size,
)
