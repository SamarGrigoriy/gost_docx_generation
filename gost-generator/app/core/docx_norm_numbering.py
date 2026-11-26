"""Нормализация секций, нумерации страниц и полей (TOC) в пакете DOCX.

Здесь собраны правки, которые задевают НЕ только `document.xml`:

* `w:sectPr` — размер/поля страницы, `w:titlePg`, `w:pgNumType`;
* части `word/footerN.xml` / `word/headerN.xml` — поле PAGE внизу по центру и ПУСТОЙ
  колонтитул первой страницы (титульный лист без номера);
* `word/settings.xml` — `w:updateFields` (иначе автосодержание не обновится при открытии);
* вставка поля `TOC`, если требование `toc.required` есть, а поля в документе нет.

Ключевой OOXML-нюанс: правильная нумерация ГОСТа — ОДНА секция + `w:titlePg` + пустой
первый нижний колонтитул. Перезапуск нумерации через `w:pgNumType w:start` во второй
секции даёт «1» на второй физической странице — ровно дефект п. 11.1 №1 ТЗ, поэтому
нормализатор эти `start` снимает.
"""

from __future__ import annotations

import re
from typing import Any

import xml.etree.ElementTree as ET

from app.core import docx_parser as dp
from app.core import docx_xml as dx
from app.core.docx_norm_util import (
    JC_TO_OOXML,
    NormCtx,
    ensure_flag,
    ensure_fonts,
    ensure_indent,
    ensure_jc,
    ensure_size,
    ensure_spacing,
    run_rpr,
    set_attr,
    text_spec,
)
from app.core.units import mm_to_twips, page_size, pt_to_half_points

__all__ = ["normalize_sections", "normalize_page_numbering", "ensure_update_fields",
           "insert_toc_field"]

# `xml:space` — зарезервированный префикс, его нет в таблице пространств qn().
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"

# Правило numbering.position → выравнивание поля в колонтитуле.
_ALIGN_BY_POSITION = {
    "bottom_center": "center", "top_center": "center", "center": "center",
    "bottom_right": "right", "top_right": "right", "right": "right",
    "bottom_left": "left", "top_left": "left", "left": "left",
}

_PAGE_RE = re.compile(r"\bPAGE\b", re.IGNORECASE)
_TOC_RE = re.compile(r"\bTOC\b", re.IGNORECASE)


def _has_page_field(el: ET.Element) -> bool:
    """Есть ли в элементе (части или абзаце) поле PAGE: instrText либо `w:fldSimple`."""
    for instr in el.iter(dx.qn("w:instrText")):
        if _PAGE_RE.search(instr.text or ""):
            return True
    for fld in el.iter(dx.qn("w:fldSimple")):
        if _PAGE_RE.search(fld.get(dx.qn("w:instr")) or ""):
            return True
    return False


# --- Секции -----------------------------------------------------------------
def normalize_sections(ctx: NormCtx) -> None:
    """`w:sectPr` всех секций: pgSz/pgMar по правилам страницы."""
    page = ctx.rules.get("page") or {}
    if not page:
        return
    before = len(ctx.changes)
    width = height = None
    landscape = False
    if page.get("size") or page.get("orientation"):
        landscape = str(page.get("orientation") or "portrait") == "landscape"
        width, height = page_size(str(page.get("size") or "A4"),
                                 str(page.get("orientation") or "portrait"))
    mm_map = {
        "w:left": page.get("margin_left_mm"), "w:right": page.get("margin_right_mm"),
        "w:top": page.get("margin_top_mm"), "w:bottom": page.get("margin_bottom_mm"),
        # расстояние до колонтитулов в методичке задают вместе с нижним полем
        "w:header": page.get("footer_distance_mm"), "w:footer": page.get("footer_distance_mm"),
    }
    mm_map = {k: v for k, v in mm_map.items() if v is not None}

    for sect in ctx.model.sections:
        notes: list[str] = []
        if width and height:
            pg_sz = dx.find_child(sect.element, "w:pgSz")
            if pg_sz is None:
                pg_sz = dx.ensure_child(sect.element, "w:pgSz", replace=False)
            changed = set_attr(pg_sz, "w:w", width)
            changed = set_attr(pg_sz, "w:h", height) or changed
            # orient пишется только для альбомной: «portrait» в файле Word не пишет
            changed = set_attr(pg_sz, "w:orient", "landscape" if landscape else None) or changed
            if changed:
                notes.append(f"размер {width}×{height} twips"
                             + (" (альбомная)" if landscape else ""))
        if mm_map:
            pg_mar = dx.find_child(sect.element, "w:pgMar")
            if pg_mar is None:
                pg_mar = dx.ensure_child(sect.element, "w:pgMar", replace=False)
            changed = False
            for attr, mm in mm_map.items():
                changed = set_attr(pg_mar, attr, mm_to_twips(float(mm))) or changed
            if changed:
                notes.append("поля/отступы колонтитулов → по правилу (мм → twips)")
        if notes:
            ctx.note("Секция: " + "; ".join(dict.fromkeys(notes)))
    if len(ctx.changes) > before:
        ctx.dirty["word/document.xml"] = ctx.model.document


# --- Нумерация страниц ------------------------------------------------------
def normalize_page_numbering(ctx: NormCtx) -> None:
    """Номер внизу по центру (поле PAGE в footer), титул без номера, без перезапуска нумерации."""
    numbering = ctx.rules.get("numbering") or {}
    if not numbering:
        return
    position = str(numbering.get("position", "bottom_center"))
    doc_dirty = _drop_pg_num_start(ctx)
    if position not in ("none", "off", ""):
        doc_dirty = _ensure_footer_page_field(ctx, _ALIGN_BY_POSITION.get(position, "center"),
                                              numbering) or doc_dirty
        doc_dirty = _strip_page_fields_from_headers(ctx) or doc_dirty
    doc_dirty = _ensure_title_page_numbering(ctx, bool(numbering.get("title_page_numbered",
                                                                    False))) or doc_dirty
    if doc_dirty:
        ctx.dirty["word/document.xml"] = ctx.model.document


def _drop_pg_num_start(ctx: NormCtx) -> bool:
    """Снять `w:pgNumType w:start`: нумерация должна идти подряд, титул отсекается titlePg."""
    changed = False
    for sect in ctx.model.sections:
        pg_num = dx.find_child(sect.element, "w:pgNumType")
        if pg_num is not None and pg_num.get(dx.qn("w:start")) is not None:
            pg_num.attrib.pop(dx.qn("w:start"))
            if not len(pg_num):
                sect.element.remove(pg_num)
            ctx.note("Секция: снят w:pgNumType w:start — вторая страница показывает «2», "
                     "нумерация не перезапускается")
            changed = True
    return changed


def _footer_rid_by_part(ctx: NormCtx) -> dict[str, str]:
    """имя части → r:id (берём из уже существующих relationships, не создавая дублей)."""
    out: dict[str, str] = {}
    for rid, (_rtype, target) in ctx.model.rels.items():
        base = target.rsplit("/", 1)[-1]
        if base.startswith("footer"):
            out.setdefault(target, rid)
    return out


def _page_footers(ctx: NormCtx) -> list[str]:
    return [name for name, root in ctx.model.footer_parts.items() if _has_page_field(root)]


def _add_footer_reference(sect: ET.Element, ftype: str, rid: str) -> None:
    """Ссылка в законном месте sectPr; replace=False: footerReference не уникальны по схеме."""
    dx.ensure_child(sect, "w:footerReference", replace=False,
                    attrs={"w:type": ftype, "r:id": rid})


def _ensure_footer_page_field(ctx: NormCtx, want_align: str, numbering: dict[str, Any]) -> bool:
    """В каждой секции — ссылка на нижний колонтитул с полем PAGE, выровненным по правилу."""
    sections = ctx.model.sections
    if not sections:
        return False
    doc_dirty = False
    page_footers = _page_footers(ctx)
    # несём поле в тот колонтитул, на который секция уже ссылается как default (rels не трогаем)
    target = next((sect.footers.get("default") for sect in sections
                   if sect.footers.get("default") in page_footers), None)
    if target is None:
        target = next((sect.footers.get("default") for sect in sections
                       if sect.footers.get("default")), None)
    if target is None:
        # ссылок нет вовсе: берём существующую часть footer*.xml, иначе — создаём новую
        rid_map = _footer_rid_by_part(ctx)
        spare = next((n for n in ctx.model.footer_parts), None)
        target = spare
        if spare is None:
            rid = ctx.add_part("footer", _footer_xml(ctx, want_align,
                                                    numbering.get("font_size_pt")))
            if rid:
                for sect in sections:
                    _add_footer_reference(sect.element, "default", rid)
                ctx.note("Создан нижний колонтитул с полем PAGE, добавлены "
                         "w:footerReference type=\"default\" во все секции")
                doc_dirty = True
        else:
            rid = rid_map.get(spare) or ""
            if rid:
                for sect in sections:
                    _add_footer_reference(sect.element, "default", rid)
                ctx.note(f"Секции: добавлена ссылка w:footerReference type=\"default\" "
                         f"→ {spare} (иначе номера не печатаются)")
                doc_dirty = True

    if target is not None:
        root = ctx.model.footer_parts.get(target)
        if root is None:
            return doc_dirty
        if not _has_page_field(root):
            _insert_page_field(ctx, root, want_align, numbering)
            ctx.note(f"{target}: вставлено поле PAGE (номер страницы "
                     f"{_align_word(want_align)})")
            ctx.dirty[target] = root
        elif _align_footer(ctx, root, want_align, numbering):
            ctx.note(f"{target}: номер страницы выровнен «{_align_word(want_align)}», "
                     f"кегль по правилу numbering")
            ctx.dirty[target] = root
        # остальные колонтитулы с полем тоже приводим к нужному краю (иначе NUM_POSITION
        # видит смешанное выравнивание среди используемых частей)
        for name in page_footers:
            if name == target or name not in _referenced_footers(ctx):
                continue
            other = ctx.model.footer_parts.get(name)
            if other is not None and _align_footer(ctx, other, want_align, numbering):
                ctx.note(f"{name}: выравнивание номера → «{_align_word(want_align)}»")
                ctx.dirty[name] = other
    return doc_dirty


def _referenced_footers(ctx: NormCtx) -> set[str]:
    return {part for sect in ctx.model.sections for part in sect.footers.values()}


def _align_word(jc: str) -> str:
    return {"center": "по центру", "left": "по левому краю",
            "right": "по правому краю"}.get(jc, jc)


def _align_footer(ctx: NormCtx, root: ET.Element, want_align: str,
                  numbering: dict[str, Any]) -> bool:
    """Выравнивание абзаца с полем PAGE + кегль/шрифт номера из `numbering`."""
    changed = False
    t = text_spec(ctx.rules)
    half = pt_to_half_points(float(numbering["font_size_pt"])) if numbering.get(
        "font_size_pt") else None
    for para in dx.find_all(root, "w:p"):
        if not _has_page_field(para):
            continue
        ppr = dx.find_child(para, "w:pPr")
        if ppr is None:
            ppr = dx.ensure_child(para, "w:pPr", index=0, replace=False)
        # Правим ровно выравнивание (единственное, что требует правило позиции):
        # интервалы/отступы колонтитула не трогаем — иначе нормализатор «улучшал» бы
        # уже корректный файл на каждом прогоне.
        changed = ensure_jc(ppr, JC_TO_OOXML.get(want_align, want_align)) or changed
        for r in dx.find_all(para, "w:r"):
            rpr = dx.find_child(r, "w:rPr")
            if rpr is None:
                continue
            if half is not None:
                changed = ensure_size(rpr, half) or changed
            changed = ensure_fonts(rpr, t.font) or changed
    return changed


def _insert_page_field(ctx: NormCtx, root: ET.Element, want_align: str,
                       numbering: dict[str, Any]) -> bool:
    """Добавить в первый абзац колонтитула последовательность детей поля PAGE."""
    para = dx.find_child(root, "w:p")
    if para is None:
        para = ET.SubElement(root, dx.qn("w:p"))
    t = text_spec(ctx.rules)
    half = pt_to_half_points(float(numbering["font_size_pt"])) if numbering.get(
        "font_size_pt") else t.size_half
    ppr = dx.find_child(para, "w:pPr")
    if ppr is None:
        ppr = dx.ensure_child(para, "w:pPr", index=0, replace=False)
    ensure_spacing(ppr, line=240, before=0, after=0)
    ensure_indent(ppr, 0)
    ensure_jc(ppr, JC_TO_OOXML.get(want_align, want_align))
    cached = str(int(numbering.get("second_page_shows", 2) or 2))
    for kind in ("begin", "instr", "separate", "text", "end"):
        r = ET.SubElement(para, dx.qn("w:r"))
        rpr = run_rpr(r, create=True)
        if rpr is not None:
            ensure_fonts(rpr, t.font)
            ensure_size(rpr, half)
        if kind in ("begin", "separate", "end"):
            el = ET.SubElement(r, dx.qn("w:fldChar"))
            el.set(dx.qn("w:fldCharType"), kind)
        elif kind == "instr":
            el = ET.SubElement(r, dx.qn("w:instrText"))
            el.set(XML_SPACE, "preserve")
            el.text = " PAGE "
        else:
            # закешированный результат поля: без него Word показывает «Не определено»
            el = ET.SubElement(r, dx.qn("w:t"))
            el.text = cached
    return True


def _strip_page_fields_from_headers(ctx: NormCtx) -> bool:
    """Убрать поле PAGE из используемых верхних колонтитулов: номер — только внизу."""
    used = {part for sect in ctx.model.sections for part in sect.headers.values()}
    changed = False
    for name, root in ctx.model.header_parts.items():
        if name not in used or not _has_page_field(root):
            continue
        local = False
        for para in dx.find_all(root, "w:p"):
            local = _drop_field(para) or local
        if local:
            ctx.dirty[name] = root
            ctx.note(f"{name}: поле PAGE из верхнего колонтитула удалено "
                     "(по правилу номер стоит внизу)")
            changed = True
    return changed


def _drop_field(p_el: ET.Element) -> bool:
    """Вырезать из абзаца run'ы поля begin…end (вместе с закешированным результатом)."""
    removed = 0
    inside = False
    for child in list(p_el):
        if dx.local_name(child.tag) == "r":
            fld = dx.find_child(child, "w:fldChar")
            instr = dx.find_child(child, "w:instrText")
            kind = fld.get(dx.qn("w:fldCharType")) if fld is not None else None
            if kind == "begin":
                inside = True
            if inside:
                p_el.remove(child)
                removed += 1
                if kind == "end":
                    inside = False
                continue
            if instr is not None and _PAGE_RE.search(instr.text or ""):
                p_el.remove(child)
                removed += 1
        elif dx.local_name(child.tag) == "fldSimple" and \
                _PAGE_RE.search(child.get(dx.qn("w:instr")) or ""):
            p_el.remove(child)
            removed += 1
    return removed > 0


def _ensure_title_page_numbering(ctx: NormCtx, titled: bool) -> bool:
    """Титульный лист без номера: `w:titlePg` + пустой «первый» нижний колонтитул."""
    sections = ctx.model.sections
    if not sections or titled:
        return False
    first = sections[0]
    changed = False
    if not first.title_pg:
        ensure_flag(first.element, "w:titlePg", True)
        ctx.note("Секция: включён w:titlePg — на титульном листе номер не печатается")
        changed = True
    if "first" not in first.footers:
        rid_map = _footer_rid_by_part(ctx)
        empty = _find_empty_footer(ctx)
        if empty is not None:
            rid = rid_map.get(empty) or ""
            if rid:
                _add_footer_reference(first.element, "first", rid)
                ctx.note(f"Секция: w:footerReference type=\"first\" → {empty} "
                         "(титульный лист без номера)")
                changed = True
        else:
            rid = ctx.add_part("footer", _footer_xml(ctx, "center", None, with_field=False))
            if rid:
                _add_footer_reference(first.element, "first", rid)
                ctx.note("Создан пустой нижний колонтитул первой страницы "
                         "(титульный лист без номера)")
                changed = True
    return changed


def _find_empty_footer(ctx: NormCtx) -> str | None:
    """Существующий нижний колонтитул БЕЗ поля, не занятый под default."""
    used_default = {sect.footers.get("default") for sect in ctx.model.sections}
    for name, root in ctx.model.footer_parts.items():
        if not _has_page_field(root) and name not in used_default:
            return name
    return None


def _footer_xml(ctx: NormCtx, want_align: str, size_pt: Any, *,
                with_field: bool = True) -> str:
    """XML нового нижнего колонтитула (часть создаётся «с нуля», поэтому строкой)."""
    t = text_spec(ctx.rules)
    size = pt_to_half_points(float(size_pt)) if size_pt else t.size_half
    rpr = (f'<w:rPr><w:rFonts w:ascii="{t.font}" w:hAnsi="{t.font}" w:cs="{t.font}"/>'
           f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr>')
    inner = ""
    if with_field:
        inner = (
            f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
            f'<w:r>{rpr}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
            f'<w:r>{rpr}<w:t>2</w:t></w:r>'
            f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
        )
    jc = JC_TO_OOXML.get(want_align, want_align)
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:ftr xmlns:w="{dx.W}" xmlns:r="{dx.R}">'
        f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" w:lineRule="auto"/>'
        f'<w:ind w:firstLine="0"/><w:jc w:val="{jc}"/></w:pPr>{inner}</w:p></w:ftr>'
    )


# --- Поля и settings -------------------------------------------------------
def insert_toc_field(ctx: NormCtx) -> bool:
    """Вставка поля автосодержания после заголовка «СОДЕРЖАНИЕ» (если `toc.required`)."""
    if not ctx.rules.get("toc") or not dp.rget(ctx.rules, "toc.required", True):
        return False
    document = ctx.model.document
    if any(_TOC_RE.search(i.text or "") for i in document.iter(dx.qn("w:instrText"))):
        return False
    if any(_TOC_RE.search(f.get(dx.qn("w:instr")) or "")
           for f in document.iter(dx.qn("w:fldSimple"))):
        return False
    body = dx.find_child(document, "w:body")
    if body is None:
        return False
    heading = None
    for para in ctx.model.paragraphs:
        text = " ".join(para.text.split()).strip().upper()
        if not para.in_table and (text.startswith("СОДЕРЖАНИЕ") or text.startswith("ОГЛАВЛЕНИЕ")):
            heading = para
            break
    if heading is None:
        return False  # вставлять некуда: заголовка нет — выдумывать структуру не будем
    t = text_spec(ctx.rules)
    rpr = (f'<w:rPr><w:rFonts w:ascii="{t.font}" w:hAnsi="{t.font}"/>'
           f'<w:sz w:val="{t.size_half}"/></w:rPr>')
    field = (
        f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rpr}<w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u '
        f'</w:instrText></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr}<w:t>Оглавление обновится при открытии документа</w:t></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
    )
    p_el = ET.Element(dx.qn("w:p"))
    ppr = dx.ensure_child(p_el, "w:pPr", index=0, replace=False)
    ensure_spacing(ppr, line=240, before=0, after=0)
    ensure_indent(ppr, 0)
    ensure_jc(ppr, "left")
    wrap = ET.fromstring(f'<wrap xmlns:w="{dx.W}" xmlns:r="{dx.R}">{field}</wrap>')
    for child in list(wrap):
        p_el.append(child)
    children = list(body)
    body.insert(children.index(heading.element) + 1, p_el)
    ctx.dirty["word/document.xml"] = document
    ctx.note("document.xml: вставлено поле автосодержания TOC после заголовка «СОДЕРЖАНИЕ»")
    return True


def ensure_update_fields(ctx: NormCtx) -> None:
    """`w:updateFields` в settings.xml — иначе TOC и номера не пересчитаются при открытии."""
    settings = ctx.model.settings
    if settings is None or not dp.rget(ctx.rules, "toc.auto", False):
        return
    if not any(_TOC_RE.search(i.text or "") for i in
               ctx.model.document.iter(dx.qn("w:instrText"))):
        return
    el = dx.find_child(settings, "w:updateFields")
    if el is not None and (el.get(dx.qn("w:val")) or "true").lower() not in ("0", "false", "none"):
        return
    if el is None:
        # порядок детей settings.xml важен: updateFields стоит перед compat/docVars —
        # ensure_child сам ставит элемент в законную позицию
        dx.ensure_child(settings, "w:updateFields", replace=False, attrs={"w:val": "true"})
    else:
        set_attr(el, "w:val", "true")
    ctx.dirty["word/settings.xml"] = settings
    ctx.note("settings.xml: включён w:updateFields — поля (TOC, PAGE) обновятся при открытии")
