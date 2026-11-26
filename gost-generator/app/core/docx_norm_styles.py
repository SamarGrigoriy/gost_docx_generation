"""Приведение `word/styles.xml` к правилам пресета.

Почему нормализация начинается со стилей: инспектор считает РЕАЛЬНЫЕ шрифт/кегль/интервал
по цепочке наследования (run → метка абзаца → basedOn → docDefaults). Если править только
прямое форматирование, документ «починится» визуально, но стиль Normal продолжит тянуть
Arial/12 pt/1,0 во все новые абзацы. Поэтому:

* `docDefaults/rPrDefault/rPr` — шрифт и кегль по правилу `text.font`/`text.size_pt`;
* `docDefaults/pPrDefault/pPr` — интервал, отступ, выравнивание, интервалы до/после;
* стиль `Normal` — то же самое (он default="1" и наследует docDefaults, но генераторы
  часто дублируют свойства вразрез с ними);
* `Heading1..Heading4` — полужирание, кегль не больше основного, keepNext, одиночный
  интервал, корректное встроенное имя `heading N` (иначе Word не считает стиль
  заголовком уровня и автосодержание его не подхватит).
"""

from __future__ import annotations

import re
from typing import Any, Mapping

import xml.etree.ElementTree as ET

from app.core import docx_xml as dx
from app.core.docx_norm_util import (
    TextSpec,
    ensure_bold,
    ensure_flag,
    ensure_fonts,
    ensure_indent,
    ensure_jc,
    ensure_size,
    ensure_spacing,
    set_attr,
    text_spec,
)
from app.core.docx_parser import rget
from app.core.units import half_points_to_pt

__all__ = ["normalize_styles"]

_HEADING_ID = re.compile(r"^heading\s*([1-4])$", re.IGNORECASE)
_HEADING_NAME = re.compile(r"^(?:heading|заголов(?:ок)?)[\s_-]*([1-4])$", re.IGNORECASE)


def _styles_root(styles: ET.Element) -> list[ET.Element]:
    """Все `w:style` (пути задаём явно: docDefaults тоже лежит в корне)."""
    return dx.find_all(styles, "w:style")


def _style_id(style: ET.Element) -> str:
    return style.get(dx.qn("w:styleId")) or ""


def _style_name(style: ET.Element) -> str:
    name = dx.find_child(style, "w:name")
    return (name.get(dx.qn("w:val")) if name is not None else "") or ""


def _ensure_defaults(styles: ET.Element, t: TextSpec, ctx: Any) -> None:
    """docDefaults: базис, который виден ВСЕМ стилям и прямому форматированию без rFonts."""
    rpr = dx.find_child(styles, "w:docDefaults/w:rPrDefault/w:rPr")
    if rpr is not None:
        if ensure_fonts(rpr, t.font):
            ctx.note(f"styles.xml: базовый шрифт (docDefaults) → {t.font}")
        if ensure_size(rpr, t.size_half):
            ctx.note(f"styles.xml: базовый кегль (docDefaults) → "
                     f"{half_points_to_pt(t.size_half):g} pt")
    ppr = dx.find_child(styles, "w:docDefaults/w:pPrDefault/w:pPr")
    if ppr is not None:
        changed = False
        if rget(ctx.rules, "text.line_spacing") is not None:
            changed = ensure_spacing(ppr, line=t.line, before=t.before, after=t.after) or changed
        if rget(ctx.rules, "text.first_line_indent_cm") is not None:
            changed = ensure_indent(ppr, t.indent) or changed
        if rget(ctx.rules, "text.alignment") is not None:
            changed = ensure_jc(ppr, t.jc) or changed
        if changed:
            ctx.note("styles.xml: абзац по умолчанию — "
                     f"интервал {t.line / 240:g}, отступ {t.indent} twips, "
                     f"выравнивание «{t.jc}», интервалы до/после {t.before}/{t.after} twips")


def _ensure_style_props(style: ET.Element, t: TextSpec, *, indent: bool) -> tuple[bool, bool]:
    """Свойства текстового стиля; возвращает (изменено_в_rPr, изменено_в_pPr)."""
    rpr = dx.find_child(style, "w:rPr")
    if rpr is None:
        rpr = dx.ensure_child(style, "w:rPr", replace=False)
    font_changed = ensure_fonts(rpr, t.font)
    size_changed = ensure_size(rpr, t.size_half)

    ppr = dx.find_child(style, "w:pPr")
    if ppr is None:
        ppr = dx.ensure_child(style, "w:pPr", index=0, replace=False)
    para_changed = ensure_spacing(ppr, line=t.line, before=t.before, after=t.after)
    if indent:
        para_changed = ensure_indent(ppr, t.indent) or para_changed
    para_changed = ensure_jc(ppr, t.jc) or para_changed
    return (font_changed or size_changed), para_changed


def _find_or_create_heading(styles: ET.Element, level: int) -> ET.Element:
    """Стиль заголовка уровня N: по styleId, затем по встроенному имени; иначе — создаём."""
    for style in _styles_root(styles):
        if style.get(dx.qn("w:type")) != "paragraph":
            continue
        sid = _style_id(style)
        if _HEADING_ID.match(sid) and sid.lower() == f"heading{level}":
            return style
    for style in _styles_root(styles):
        match = _HEADING_NAME.match(_style_name(style).strip())
        if match and int(match.group(1)) == level:
            return style
    style = ET.SubElement(styles, dx.qn("w:style"))
    style.set(dx.qn("w:type"), "paragraph")
    style.set(dx.qn("w:styleId"), f"Heading{level}")
    name = ET.SubElement(style, dx.qn("w:name"))
    name.set(dx.qn("w:val"), f"heading {level}")
    return style


def _ensure_heading_style(styles: ET.Element, level: int, t: TextSpec,
                          rules: Mapping[str, Any], ctx: Any) -> None:
    """Heading N: полужирный, кегль ≤ основного, keepNext, одиночный интервал, outlineLvl."""
    style = _find_or_create_heading(styles, level)
    notes: list[str] = []

    name_el = dx.find_child(style, "w:name")
    if name_el is None:
        name_el = dx.ensure_child(style, "w:name", index=0, replace=True)
    if set_attr(name_el, "w:val", f"heading {level}"):
        notes.append(f"имя → «heading {level}»")

    # basedOn на Normal даёт наследование шрифта/интервала текста; Normal на себя — цикл
    based = dx.find_child(style, "w:basedOn")
    if based is None and _style_id(style) != "Normal":
        dx.ensure_child(style, "w:basedOn", index=1, attrs={"w:val": "Normal"},
                        replace=True)
        notes.append("наследование от Normal")

    rpr = dx.find_child(style, "w:rPr")
    if rpr is None:
        rpr = dx.ensure_child(style, "w:rPr", replace=False)
    if ensure_fonts(rpr, t.font):
        notes.append(f"шрифт → {t.font}")
    # same_size_as_body: кегль заголовка не должен превышать основной (п. 9.1 HEAD_SIZE)
    current = dx.get_int(dx.find_child(rpr, "w:sz"), "w:val")
    if current is None or (rget(rules, "headings.same_size_as_body", True)
                           and current > t.size_half):
        if ensure_size(rpr, t.size_half):
            notes.append(f"кегль → {half_points_to_pt(t.size_half):g} pt")
    want_bold = bool(rget(rules, "headings.section_bold", True))
    if ensure_bold(rpr, want_bold):
        notes.append("полужирный" if want_bold else "обычный")

    ppr = dx.find_child(style, "w:pPr")
    if ppr is None:
        ppr = dx.ensure_child(style, "w:pPr", index=0, replace=False)
    head_line = rget(rules, "headings.multiline_line_spacing", 1.0)
    if ensure_spacing(ppr, line=int(round(float(head_line or 1.0) * 240)), before=0, after=0):
        notes.append(f"интервал {float(head_line or 1.0):g}")
    # Заголовок не должен получать абзацный отступ «по центру» — при center/right firstLine
    # сдвигает строку, поэтому для центрированных структурных заголовков он обнуляется.
    if rget(rules, "headings.section_from_indent", True):
        if ensure_indent(ppr, t.indent):
            notes.append(f"отступ {t.indent} twips")
    if ensure_jc(ppr, "left"):
        notes.append("выравнивание заголовка → по левому краю")
    if ensure_flag(ppr, "w:keepNext", bool(rget(rules, "headings.keep_with_next", True))):
        notes.append("keepNext (заголовок не отрывается от текста)")
    outline = dx.find_child(ppr, "w:outlineLvl")
    if outline is None:
        outline = dx.ensure_child(ppr, "w:outlineLvl", replace=True)
    if set_attr(outline, "w:val", level - 1):
        notes.append(f"уровень структуры {level - 1}")

    if notes:
        ctx.note(f"styles.xml: стиль Heading{level} — " + "; ".join(notes))


def normalize_styles(styles: ET.Element | None, rules: Mapping[str, Any], ctx: Any) -> None:
    """Полная нормализация стилей; `ctx.dirty` помечается, если что-то изменилось."""
    if styles is None:
        return
    before = len(ctx.changes)
    t = text_spec(rules)
    _ensure_defaults(styles, t, ctx)

    for style in _styles_root(styles):
        if style.get(dx.qn("w:type")) != "paragraph":
            continue
        sid = _style_id(style)
        heading = _HEADING_ID.match(sid) or _HEADING_NAME.match(_style_name(style).strip())
        if heading:
            continue  # заголовки — ниже, своей логикой
        is_default = sid.lower() == "normal" or style.get(dx.qn("w:default")) == "1"
        if not is_default:
            continue
        # Normal, basedOn на самого себя — типичная ошибка генераторов: цепочка наследования
        # превращается в цикл и свойства текста перестают применяться.
        based = dx.find_child(style, "w:basedOn")
        if based is not None and (based.get(dx.qn("w:val")) or "") == sid:
            style.remove(based)
            ctx.note("styles.xml: у стиля Normal убран самонаследование basedOn")
        r_changed, p_changed = _ensure_style_props(style, t, indent=True)
        if r_changed or p_changed:
            ctx.note(f"styles.xml: стиль Normal приведён к правилу текста "
                     f"(«{t.font}», {half_points_to_pt(t.size_half):g} pt, "
                     f"интервал {t.line / 240:g}, отступ {t.indent})")

    for level in (1, 2, 3, 4):
        _ensure_heading_style(styles, level, t, rules, ctx)

    if len(ctx.changes) > before:
        ctx.dirty["word/styles.xml"] = styles
