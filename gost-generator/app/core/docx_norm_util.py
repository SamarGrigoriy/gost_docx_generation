"""Сеттеры OOXML-свойств для нормализатора + общий контекст правок.

Модуль отвечает на вопрос «как ПРАВИЛЬНО поправить элемент», а не «что править»:

* значения всегда сравниваются с текущими — иначе нормализатор переставляет атрибуты
  на каждом прогоне и идемпотентность (два прогона = байт-в-байт) ломается;
* элементы создаются только через `docx_xml.ensure_child`, чтобы не нарушать порядок
  детей `w:pPr`/`w:rPr`/`w:sectPr` (Word иначе считает пакет битым);
* логические флаги (`w:b`, `w:keepNext`, `w:tblHeader`) в OOXML triple-state:
  элемент без `w:val` = true, `w:val="0"|"false"` = false, отсутствие = не задано.

`NormCtx` — то, что передают все функции правок: модель, правила, список изменений
(человекочитаемый) и реестр «грязных» частей пакета.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable, Mapping, MutableMapping

import xml.etree.ElementTree as ET

from app import constants
from app.core import docx_xml as dx
from app.core.docx_parser import DocxModel, rget
from app.core.units import cm_to_twips, pt_to_half_points, spacing_to_line

__all__ = [
    "TextSpec",
    "NormCtx",
    "JC_TO_OOXML",
    "text_spec",
    "set_attr",
    "ensure_flag",
    "ensure_fonts",
    "ensure_size",
    "ensure_bold",
    "ensure_spacing",
    "ensure_indent",
    "ensure_jc",
    "paragraph_ppr",
    "run_rpr",
]

# Правила говорят «justify», в OOXML пишется «both».
JC_TO_OOXML: dict[str, str] = {
    "justify": "both", "both": "both", "left": "left", "center": "center",
    "right": "right", "distribute": "distribute",
}

FALSE_VALUES = ("0", "false", "none", "no")


@dataclass
class TextSpec:
    """Итоговые значения основного текста в единицах OOXML (twips/полупоинты)."""

    font: str
    size_half: int
    line: int
    indent: int
    jc: str
    before: int
    after: int


def text_spec(rules: Mapping[str, Any]) -> TextSpec:
    """Пересчёт секции `text` правил в атрибуты OOXML.

    Отсутствие ключа трактуется как «правило выключено»: берём константу, но вызывающий
    обязан сам проверить наличие ключа, если не хочет править то, что пресет не задаёт.
    """
    size_pt = float(rget(rules, "text.size_pt", constants.DEFAULT_SIZE_PT) or 0)
    spacing = float(rget(rules, "text.line_spacing", constants.DEFAULT_LINE_SPACING) or 1.0)
    indent_cm = float(rget(rules, "text.first_line_indent_cm", 0) or 0)
    return TextSpec(
        font=str(rget(rules, "text.font", constants.DEFAULT_FONT) or constants.DEFAULT_FONT),
        size_half=pt_to_half_points(size_pt) if size_pt else pt_to_half_points(
            constants.DEFAULT_SIZE_PT),
        line=spacing_to_line(spacing),
        indent=cm_to_twips(indent_cm),
        jc=JC_TO_OOXML.get(str(rget(rules, "text.alignment",
                                    constants.DEFAULT_ALIGN)).lower(), "both"),
        before=int(round(float(rget(rules, "text.space_before_pt", 0) or 0) * 20)),
        after=int(round(float(rget(rules, "text.space_after_pt", 0) or 0) * 20)),
    )


@dataclass
class NormCtx:
    """Контекст одной сессии нормализации.

    `dirty` — части, дерево которых изменено (оркестратор сериализует их обратно);
    `add_part(kind, xml)` — выделение имени/rel для НОВОЙ части пакета (нужно только
    для колонтитулов, которых в документе не было; media и существующие rels не трогаются).
    """

    model: DocxModel
    rules: Mapping[str, Any]
    changes: list[str] = field(default_factory=list)
    dirty: MutableMapping[str, ET.Element] = field(default_factory=dict)
    add_part: Callable[[str, str], str] = lambda kind, xml: ""

    def note(self, message: str) -> None:
        """Зафиксировать изменение человекочитаемой строкой (попадает в результат)."""
        self.changes.append(message)


# --- Атрибуты ---------------------------------------------------------------
def set_attr(el: ET.Element, name: str, value: Any) -> bool:
    """Установить (value=None — снять) атрибут, только если значение отличается.

    Возвращает факт правки: именно на ней держится идемпотентность — повторный прогон
    не должен менять ни байта в уже нормализованной части.
    """
    qname = dx.qn(name)
    if value is None:
        if qname in el.attrib:
            del el.attrib[qname]
            return True
        return False
    want = str(value)
    if el.get(qname) == want:
        return False
    el.set(qname, want)
    return True


def ensure_flag(parent: ET.Element, tag: str, want: bool) -> bool:
    """Логический элемент-флаг (`w:keepNext`, `w:b`, `w:tblHeader`, `w:titlePg`)."""
    existing = dx.find_child(parent, tag)
    if want:
        if existing is not None:
            raw = existing.get(dx.qn("w:val"))
            if raw is None or raw.strip().lower() not in FALSE_VALUES:
                return False
            # элемент есть, но выключен — снимаем w:val (отсутствие = true)
            existing.attrib.pop(dx.qn("w:val"), None)
            return True
        dx.ensure_child(parent, tag, replace=False)
        return True
    if existing is None:
        return False
    parent.remove(existing)
    return True


# --- rPr / pPr --------------------------------------------------------------
def paragraph_ppr(p_el: ET.Element, ctx: NormCtx, *, create: bool = True) -> ET.Element | None:
    """`w:pPr` абзаца (создаётся первым ребёнком — так требует CT_P)."""
    ppr = dx.find_child(p_el, "w:pPr")
    if ppr is not None:
        return ppr
    if not create:
        return None
    ppr = ET.Element(dx.qn("w:pPr"))
    p_el.insert(0, ppr)
    return ppr


def run_rpr(r_el: ET.Element, *, create: bool = True) -> ET.Element | None:
    """`w:rPr` run'а: по схеме это ВСЕГДА первый ребёнок `w:r`."""
    rpr = dx.find_child(r_el, "w:rPr")
    if rpr is not None:
        return rpr
    if not create:
        return None
    rpr = ET.Element(dx.qn("w:rPr"))
    r_el.insert(0, rpr)
    return rpr


def ensure_fonts(rpr: ET.Element, font: str) -> bool:
    """rFonts: шрифт обязан быть задан явно, иначе Word подставит значение из темы.

    Правим только `ascii`/`hAnsi` (кириллица в Word идёт по hAnsi) и те из `cs`/`eastAsia`,
    которые в документе УЖЕ есть: добавлять их «для полноты» значит менять байты на
    корректном файле и ломать идемпотентность.
    """
    fonts = dx.find_child(rpr, "w:rFonts")
    if fonts is None:
        fonts = dx.ensure_child(rpr, "w:rFonts", replace=True)
    changed = False
    for attr in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
        if attr in ("w:cs", "w:eastAsia") and dx.qn(attr) not in fonts.attrib:
            continue
        changed = set_attr(fonts, attr, font) or changed
    return changed


def ensure_size(rpr: ET.Element, half_points: int) -> bool:
    """w:sz; w:szCs синхронизируем, только если он уже есть (см. rationale `ensure_fonts`)."""
    changed = False
    sz = dx.find_child(rpr, "w:sz")
    if sz is None:
        sz = dx.ensure_child(rpr, "w:sz", replace=True)
    changed = set_attr(sz, "w:val", half_points)
    szcs = dx.find_child(rpr, "w:szCs")
    if szcs is not None:
        changed = set_attr(szcs, "w:val", half_points) or changed
    return changed


def ensure_bold(rpr: ET.Element, want: bool) -> bool:
    """w:b (+ w:bCs, только если он уже есть — см. rationale в `ensure_fonts`)."""
    changed = ensure_flag(rpr, "w:b", want)
    bcs = dx.find_child(rpr, "w:bCs")
    if want and bcs is not None:
        changed = ensure_flag(rpr, "w:bCs", True) or changed
    elif not want and bcs is not None:
        changed = ensure_flag(rpr, "w:bCs", False) or changed
    return changed


def ensure_spacing(ppr: ET.Element, *, line: int | None = None,
                   before: int | None = None, after: int | None = None) -> bool:
    """w:spacing; lineRule всегда «auto»: только он делает `w:line` множителем интервала."""
    sp = dx.find_child(ppr, "w:spacing")
    if sp is None:
        if line is None and before is None and after is None:
            return False
        sp = dx.ensure_child(ppr, "w:spacing", replace=True)
    changed = False
    if line is not None:
        changed = set_attr(sp, "w:line", line) or changed
        changed = set_attr(sp, "w:lineRule", "auto") or changed
    if before is not None:
        changed = set_attr(sp, "w:before", before) or changed
        changed = set_attr(sp, "w:beforeLines", None) or changed
    if after is not None:
        changed = set_attr(sp, "w:after", after) or changed
        changed = set_attr(sp, "w:afterLines", None) or changed
    return changed


def ensure_indent(ppr: ET.Element, first_line: int) -> bool:
    """w:ind firstLine; firstLineChars/hanging имеют приоритет над firstLine в Word — снимаем."""
    ind = dx.find_child(ppr, "w:ind")
    if ind is None:
        ind = dx.ensure_child(ppr, "w:ind", replace=True)
    changed = set_attr(ind, "w:firstLine", first_line)
    for extra in ("w:firstLineChars", "w:hanging", "w:hangingChars"):
        if extra in ind.attrib:
            del ind.attrib[dx.qn(extra)]
            changed = True
    return changed


def ensure_jc(ppr: ET.Element, jc: str) -> bool:
    """w:jc со значением схемы («both», а не «justify»)."""
    el = dx.find_child(ppr, "w:jc")
    if el is None:
        el = dx.ensure_child(ppr, "w:jc", replace=True)
    return set_attr(el, "w:val", jc)
