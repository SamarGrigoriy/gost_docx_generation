"""Общий парсер пакета DOCX: абзацы, таблицы, изображения, колонтитулы, стили.

Задача модуля — один раз прочитать OOXML и отдать верхним слоям (инспектору и
нормализатору) одинаковую модель, чтобы их выводы не расходились.

Ключевые решения:
* свойства run/абзаца вычисляются «эффективными»: run → метка абзаца → цепочка
  стилей (basedOn) → docDefaults. Без этого Arial, заданный прямым форматированием,
  выглядел бы как Times New Roman из стиля (дефект №1 старого чекера, п. 11.2 ТЗ);
* текст ищется по всем `w:t`, табуляции и разрывы превращаются в пробелы, чтобы
  регулярки подписей не «спотыкались» о `w:tab`;
* `w:instrText` собирается отдельно — поля (PAGE, TOC) живут именно там.
"""

from __future__ import annotations

import math
import re
from dataclasses import dataclass, field
from typing import Any, Iterable, Mapping, MutableMapping

import xml.etree.ElementTree as ET

from app.core import docx_xml as dx
from app.core.docx_xml import DocxError, qn
from app.core.units import (
    cm_to_twips,
    mm_to_twips,
    page_size,
    pt_to_half_points,
    spacing_to_line,
)

# Псевдонимы выравнивания: в правилах «justify», в OOXML — «both».
JC_ALIASES: dict[str, str] = {
    "both": "justify",
    "justify": "justify",
    "left": "left",
    "center": "center",
    "right": "right",
    "distribute": "distribute",
}

OFF_VALUES = ("0", "false", "none", "no")


def _bool_of(el: ET.Element | None, attr: str = "w:val") -> bool | None:
    """Трибул `w:b`/`w:keepNext`: отсутствие — None, `w:val="0"` — False, иначе True."""
    if el is None:
        return None
    raw = el.get(qn(attr))
    if raw is None:
        return True
    return raw.strip().lower() not in OFF_VALUES


@dataclass
class RunProps:
    """Свойства фрагмента текста (полупоинты и twips — как в OOXML)."""

    font: str | None = None
    size: int | None = None
    bold: bool = False
    italic: bool = False
    superscript: bool = False
    underline: bool = False
    mono: bool = False


@dataclass
class RunInfo:
    text: str
    props: RunProps
    element: ET.Element


@dataclass
class ParaProps:
    """Набор атрибутов `w:pPr` (None = не задан, значение наследуется)."""

    style: str | None = None
    jc: str | None = None
    line: int | None = None
    line_rule: str | None = None
    before: int | None = None
    after: int | None = None
    first_line: int | None = None
    left: int | None = None
    hanging: int | None = None
    keep_next: bool | None = None
    keep_lines: bool | None = None
    page_break_before: bool | None = None


@dataclass
class ImageInfo:
    """Изображение: `wp:extent` в EMU + relationship на медиа-часть."""

    paragraph_index: int
    cx: int = 0
    cy: int = 0
    rid: str | None = None
    target: str | None = None
    kind: str = "drawing"  # drawing | pict


@dataclass
class ParagraphInfo:
    """Абзац со всеми свойствами, нужными проверкам и нормализатору."""

    index: int
    element: ET.Element
    text: str = ""
    runs: list[RunInfo] = field(default_factory=list)
    props: ParaProps = field(default_factory=ParaProps)          # только прямое форматирование
    eff: ParaProps = field(default_factory=ParaProps)            # с учётом стилей/docDefaults
    eff_run: RunProps = field(default_factory=RunProps)
    instr: str = ""
    in_table: bool = False
    table_index: int | None = None
    row_is_header: bool = False
    is_cell: bool = False
    images: list[ImageInfo] = field(default_factory=list)
    has_page_break_run: bool = False
    has_sect_break: bool = False
    is_empty: bool = True

    @property
    def has_drawing(self) -> bool:
        return bool(self.images)

    @property
    def is_code(self) -> bool:
        """Код-блок: моноширинный шрифт доминирует (Consolas и т.п.)."""
        mono = sum(len(r.text) for r in self.runs if r.props.mono)
        total = sum(len(r.text) for r in self.runs)
        return mono > 0 and mono >= max(1, total // 2)


@dataclass
class RowInfo:
    element: ET.Element
    is_header: bool
    cells: list[list[str]]


@dataclass
class TableInfo:
    index: int
    element: ET.Element
    rows: list[RowInfo]
    paragraph_indexes: list[int]
    prev_block_paragraph: int | None = None
    next_block_paragraph: int | None = None

    @property
    def is_empty(self) -> bool:
        return not self.rows


@dataclass
class SectionInfo:
    """`w:sectPr`: поля, размер, колонтитулы, тип нумерации."""

    element: ET.Element
    is_body_final: bool = False
    pg_w: int | None = None
    pg_h: int | None = None
    orient: str | None = None
    mar: dict[str, int] = field(default_factory=dict)
    title_pg: bool = False
    pg_num_start: int | None = None
    footers: dict[str, str] = field(default_factory=dict)  # type(default/first/even) → part
    headers: dict[str, str] = field(default_factory=dict)


class StyleRegistry:
    """Стили + docDefaults с обходом цепочки `basedOn`.

    Без наследования невозможно узнать РЕАЛЬНЫЙ шрифт: генераторы часто пишут
    `rFonts` только в `docDefaults`, а прямой `w:rPr` отсутствует у половины run'ов.
    """

    def __init__(self, styles_root: ET.Element | None) -> None:
        self.defaults_run = self._run_props(dx.find_child(styles_root, "w:docDefaults/w:rPrDefault/w:rPr"))
        self.defaults_para = self._para_props(dx.find_child(styles_root, "w:docDefaults/w:pPrDefault/w:pPr"))
        self.styles: dict[str, dict[str, Any]] = {}
        self.by_name: dict[str, str] = {}
        self.default_style: str | None = None
        if styles_root is None:
            return
        for st in dx.find_all(styles_root, "w:style"):
            sid = st.get(qn("w:styleId")) or ""
            name_el = dx.find_child(st, "w:name")
            self.styles[sid] = {
                "element": st,
                "id": sid,
                "type": st.get(qn("w:type")),
                "name": (name_el.get(qn("w:val")) if name_el is not None else sid),
                "based_on": (lambda e: e.get(qn("w:val")) if e is not None else None)(
                    dx.find_child(st, "w:basedOn")
                ),
                "is_default": (lambda e: _bool_of(e) is True)(dx.find_child(st, "w:default")),
                "run": self._run_props(dx.find_child(st, "w:rPr")),
                "para": self._para_props(dx.find_child(st, "w:pPr")),
            }
            if self.styles[sid]["name"]:
                self.by_name[str(self.styles[sid]["name"]).strip().lower()] = sid
            # Абзацный стиль с w:default="1" применяется к <w:p> без pStyle (напр. Normal)
            if self.styles[sid]["is_default"] and self.styles[sid]["type"] == "paragraph":
                self.default_style = sid

    @staticmethod
    def _run_props(rpr: ET.Element | None) -> dict[str, Any]:
        if rpr is None:
            return {}
        out: dict[str, Any] = {}
        fonts = dx.find_child(rpr, "w:rFonts")
        if fonts is not None:
            for attr, key in (("w:ascii", "font"), ("w:hAnsi", "hfont"), ("w:cs", "cfont"),
                              ("w:eastAsia", "efont")):
                val = fonts.get(qn(attr))
                if val:
                    out[key] = val
        sz = dx.find_child(rpr, "w:sz")
        if sz is not None:
            out["size"] = dx.get_int(sz, "w:val")
        bold = _bool_of(dx.find_child(rpr, "w:b"))
        if bold is not None:
            out["bold"] = bold
        ital = _bool_of(dx.find_child(rpr, "w:i"))
        if ital is not None:
            out["italic"] = ital
        va = dx.find_child(rpr, "w:vertAlign")
        if va is not None:
            out["superscript"] = va.get(qn("w:val")) in ("superscript", "superscript ")
        u = dx.find_child(rpr, "w:u")
        if u is not None and (u.get(qn("w:val")) or "single").lower() not in ("none", "0"):
            out["underline"] = True
        return out

    @staticmethod
    def _para_props(ppr: ET.Element | None) -> dict[str, Any]:
        if ppr is None:
            return {}
        out: dict[str, Any] = {}
        jc = dx.find_child(ppr, "w:jc")
        if jc is not None:
            out["jc"] = jc.get(qn("w:val"))
        sp = dx.find_child(ppr, "w:spacing")
        if sp is not None:
            for attr, key in (("w:line", "line"), ("w:lineRule", "line_rule"),
                              ("w:before", "before"), ("w:after", "after")):
                val = dx.get_int(sp, attr)
                if val is not None:
                    out[key] = val
        ind = dx.find_child(ppr, "w:ind")
        if ind is not None:
            first = dx.get_int(ind, "w:firstLine")
            chars = dx.get_int(ind, "w:firstLineChars")
            if first is not None:
                out["first_line"] = first
            elif chars is not None:
                out["first_line_chars"] = chars
            for attr, key in (("w:left", "left"), ("w:hanging", "hanging")):
                val = dx.get_int(ind, attr)
                if val is not None:
                    out[key] = val
        for tag, key in (("w:keepNext", "keep_next"), ("w:keepLines", "keep_lines"),
                         ("w:pageBreakBefore", "page_break_before")):
            val = _bool_of(dx.find_child(ppr, tag))
            if val is not None:
                out[key] = val
        style = dx.find_child(ppr, "w:pStyle")
        if style is not None:
            out["style"] = style.get(qn("w:val"))
        return out

    def raw_run_props(self, rpr: ET.Element | None) -> dict[str, Any]:
        """Свойства только из этого `w:rPr`, без наследования (для цепочки слияния)."""
        return self._run_props(rpr)

    def raw_para_props(self, ppr: ET.Element | None) -> dict[str, Any]:
        """Атрибуты только из этого `w:pPr`, без наследования."""
        return self._para_props(ppr)

    def _chain(self, style_id: str | None) -> list[dict[str, Any]]:
        """От корня к стилю (outermost first), с защитой от циклов basedOn.

        Абзац без `pStyle` наследует стиль по умолчанию (`w:default="1"`, иначе `Normal`):
        это свойство OOXML, а не оформительский выбор — без него реальный отступ/кегль
        текста читаются как отсутствующие, и проверка даёт ложные нарушения.
        """
        chain: list[dict[str, Any]] = []
        seen: set[str] = set()
        current = style_id or self.default_style or (
            "Normal" if "Normal" in self.styles else None)
        while current and current not in seen and current in self.styles:
            seen.add(current)
            chain.append(self.styles[current])
            current = self.styles[current]["based_on"]
        return list(reversed(chain))

    def run_props(self, style_id: str | None, *overlays: dict[str, Any]) -> dict[str, Any]:
        """Слитые свойства run: docDefaults → стили → overlays (от слабых к сильным)."""
        merged: dict[str, Any] = dict(self.defaults_run)
        for st in self._chain(style_id):
            merged.update(st["run"])
        for ov in overlays:
            merged.update(ov)
        return merged

    def para_props(self, style_id: str | None, *overlays: dict[str, Any]) -> dict[str, Any]:
        """Слитые свойства абзаца. `style` (pStyle) всегда берётся из overlays."""
        merged: dict[str, Any] = dict(self.defaults_para)
        for st in self._chain(style_id):
            merged.update({k: v for k, v in st["para"].items() if k != "style"})
        for ov in overlays:
            merged.update({k: v for k, v in ov.items() if k != "style"})
        merged["style"] = style_id
        return merged


def to_run_props(raw: dict[str, Any], mono_fonts: Iterable[str]) -> RunProps:
    """Словарь свойств → типизированный `RunProps` (шрифт = ascii/hAnsi по приоритету)."""
    font = raw.get("font") or raw.get("hfont") or raw.get("cfont")
    mono_set = {f.lower() for f in mono_fonts}
    return RunProps(
        font=font,
        size=raw.get("size"),
        bold=bool(raw.get("bold", False)),
        italic=bool(raw.get("italic", False)),
        superscript=bool(raw.get("superscript", False)),
        underline=bool(raw.get("underline", False)),
        mono=bool(font and font.lower() in mono_set),
    )


def to_para_props(raw: dict[str, Any]) -> ParaProps:
    """Словарь → `ParaProps`; firstLineChars пересчитывается в twips по кеглю."""
    first = raw.get("first_line")
    if first is None and raw.get("firstLineChars") is not None:  # pragma: no cover
        first = None
    jc = raw.get("jc")
    return ParaProps(
        style=raw.get("style"),
        jc=JC_ALIASES.get(str(jc), jc) if jc else None,
        line=raw.get("line"),
        line_rule=raw.get("line_rule"),
        before=raw.get("before"),
        after=raw.get("after"),
        first_line=first,
        left=raw.get("left"),
        hanging=raw.get("hanging"),
        keep_next=raw.get("keep_next"),
        keep_lines=raw.get("keep_lines"),
        page_break_before=raw.get("page_break_before"),
    )


@dataclass
class DocxModel:
    """Модель документа, пригодная и для проверок, и для правок."""

    path: str
    parts: MutableMapping[str, bytes]
    order: list[str]
    document: ET.Element
    styles: ET.Element | None
    settings: ET.Element | None
    registry: StyleRegistry
    paragraphs: list[ParagraphInfo]
    tables: list[TableInfo]
    images: list[ImageInfo]
    sections: list[SectionInfo]
    rels: dict[str, tuple[str, str]]
    footer_parts: dict[str, ET.Element]
    header_parts: dict[str, ET.Element]
    footnotes: ET.Element | None
    blocks: list[tuple[str, ET.Element]]
    mono_fonts: tuple[str, ...]

    @property
    def main_section(self) -> SectionInfo | None:
        """Последний `sectPr` body — «главная» секция для проверок полей/размера."""
        for sect in reversed(self.sections):
            if sect.is_body_final:
                return sect
        return self.sections[-1] if self.sections else None


class DocxParser:
    """Читает пакет и строит `DocxModel`. Один экземпляр = один документ."""

    def __init__(self, mono_fonts: Iterable[str] = ("Consolas", "Courier New", "Courier",
                                                    "Liberation Mono")) -> None:
        self.mono_fonts = tuple(mono_fonts)

    # --- публичный вход ---------------------------------------------------
    def parse_file(self, path: str) -> DocxModel:
        parts, order = dx.read_package(path)
        return self.parse_parts(path, parts, order)

    def parse_parts(self, path: str,
                    parts: MutableMapping[str, bytes],
                    order: Iterable[str]) -> DocxModel:
        order_list = list(order)
        document = dx.parse_part(parts, "word/document.xml")
        if document is None:  # pragma: no cover — read_package уже проверил наличие
            raise DocxError("Нет word/document.xml")
        styles = dx.parse_part(parts, "word/styles.xml")
        settings = dx.parse_part(parts, "word/settings.xml")
        footnotes = dx.parse_part(parts, "word/footnotes.xml")
        registry = StyleRegistry(styles)
        rels = dx.parse_rels(parts, "word/_rels/document.xml.rels")
        resolved = {rid: (rtype, self._resolve_target("word/", target))
                    for rid, (rtype, target) in rels.items()}

        paragraphs, tables, images, blocks = self._parse_body(document, registry, resolved)
        sections = self._parse_sections(document, resolved)
        footers, headers = self._parse_ftr_parts(parts, resolved)

        return DocxModel(
            path=path, parts=parts, order=order_list, document=document, styles=styles,
            settings=settings, registry=registry, paragraphs=paragraphs, tables=tables,
            images=images, sections=sections, rels=resolved, footer_parts=footers,
            header_parts=headers, footnotes=footnotes, blocks=blocks,
            mono_fonts=self.mono_fonts,
        )

    # --- внутреннее -------------------------------------------------------
    @staticmethod
    def _resolve_target(base: str, target: str) -> str:
        """Target relationship ('footer1.xml' или '../customXml/item1.xml') → путь в zip."""
        if target.startswith("/"):
            return target[1:]
        raw = base + target
        out: list[str] = []
        for seg in raw.split("/"):
            if seg == "..":
                if out:
                    out.pop()
            elif seg in ("", "."):
                continue
            else:
                out.append(seg)
        return "/".join(out)

    def _parse_body(self, document: ET.Element, registry: StyleRegistry,
                    rels: dict[str, tuple[str, str]],
                    ) -> tuple[list[ParagraphInfo], list[TableInfo], list[ImageInfo],
                               list[tuple[str, ET.Element]]]:
        body = dx.find_child(document, "w:body")
        paragraphs: list[ParagraphInfo] = []
        tables: list[TableInfo] = []
        images: list[ImageInfo] = []
        blocks: list[tuple[str, ET.Element]] = []
        if body is None:
            return paragraphs, tables, images, blocks

        # Параграфы внутри таблиц помечаются заранее: проверки основного текста
        # обязаны их исключать (дефект №2 старого чекера — кегль таблиц перебивает статистику).
        in_tbl: set[int] = set()
        table_els: list[ET.Element] = []
        for child in list(body):
            tag = dx.local_name(child.tag)
            blocks.append((tag, child))
            if tag == "tbl":
                table_els.append(child)
                in_tbl.update(id(p) for p in child.iter(qn("w:p")))

        # Индексация в документном порядке (включая ячейки) — это «номер абзаца» в location.
        by_element: dict[int, int] = {}
        for p_el in document.iter(qn("w:p")):
            info = self._parse_paragraph(p_el, registry, rels, len(paragraphs),
                                         id(p_el) in in_tbl)
            paragraphs.append(info)
            images.extend(info.images)
            by_element[id(p_el)] = info.index

        for tbl in table_els:
            rows: list[RowInfo] = []
            for tr in dx.find_all(tbl, "w:tr"):
                trpr = dx.find_child(tr, "w:trPr")
                is_header = bool(trpr is not None
                                 and _bool_of(dx.find_child(trpr, "w:tblHeader")))
                cells: list[list[str]] = []
                for tc in dx.find_all(tr, "w:tc"):
                    texts = [paragraph_text(p).strip() for p in dx.find_all(tc, "w:p")]
                    cells.append([t for t in texts if t] or [""])
                rows.append(RowInfo(element=tr, is_header=is_header, cells=cells))
            p_indexes = [by_element[id(p)] for p in tbl.iter(qn("w:p")) if id(p) in by_element]
            tables.append(TableInfo(index=len(tables), element=tbl, rows=rows,
                                    paragraph_indexes=p_indexes))
        for tbl in tables:
            head = self._first_row_paragraphs(tbl, by_element)
            for p_i in tbl.paragraph_indexes:
                paragraphs[p_i].table_index = tbl.index
                paragraphs[p_i].is_cell = True
                paragraphs[p_i].row_is_header = p_i in head
        return paragraphs, tables, images, blocks

    def _parse_paragraph(self, p_el: ET.Element, registry: StyleRegistry,
                         rels: dict[str, tuple[str, str]], index: int,
                         p_el_id_in_table: bool) -> ParagraphInfo:
        ppr = dx.find_child(p_el, "w:pPr")
        direct = registry.raw_para_props(ppr) if ppr is not None else {}
        style_id = direct.get("style")
        mark_rpr = dx.find_child(p_el, "w:pPr/w:rPr")
        mark_props = registry.raw_run_props(mark_rpr) if mark_rpr is not None else {}

        runs: list[RunInfo] = []
        # Обход ВСЕХ w:r в документном порядке: run'ы лежат и в w:hyperlink, и в w:ins,
        # и в w:smartTag — по прямым детям их легко потерять.
        for r_el in p_el.iter(qn("w:r")):
            r_props = registry.run_props(style_id, mark_props,
                                         registry.raw_run_props(dx.find_child(r_el, "w:rPr")))
            runs.append(RunInfo(text=self._run_text(r_el),
                                props=to_run_props(r_props, self.mono_fonts), element=r_el))
        instr_parts = [(i.text or "").strip() for i in p_el.iter(qn("w:instrText")) if i.text]
        # Рисунки: w:drawing (современный OOXML) и w:pict (устаревший VML) — оба считаются.
        images: list[ImageInfo] = []
        for drawing in p_el.iter(qn("w:drawing")):
            images.extend(self._parse_drawings(drawing, index, rels))
        for pict in p_el.iter(qn("w:pict")):
            images.extend(self._parse_picts(pict, index, rels))
        has_break = any(br.get(qn("w:type")) == "page" for br in p_el.iter(qn("w:br")))

        text = "".join(r.text for r in runs)
        eff_raw = registry.para_props(style_id, direct)
        eff_run = self._dominant_run(runs, to_run_props(registry.run_props(style_id, mark_props),
                                                        self.mono_fonts))
        first_line_chars = direct.get("first_line_chars") or eff_raw.get("first_line_chars")
        if eff_raw.get("first_line") is None and first_line_chars:
            # firstLineChars задаётся в сотых долях символа: переводим в twips по кеглю
            size = (eff_run.size or pt_to_half_points(12))
            eff_raw["first_line"] = int(round(first_line_chars / 100 * size * 10))

        info = ParagraphInfo(
            index=index,
            element=p_el,
            text=text,
            runs=runs,
            props=to_para_props({**direct, "style": style_id}),
            eff=to_para_props({**eff_raw, "style": style_id}),
            eff_run=eff_run,
            instr=" ".join(instr_parts).strip(),
            images=images,
            in_table=p_el_id_in_table,
            has_page_break_run=has_break,
            has_sect_break=ppr is not None and dx.has_child(ppr, "w:sectPr"),
            is_empty=not text.strip() and not images,
        )
        return info

    @staticmethod
    def _first_row_paragraphs(table: TableInfo, by_element: dict[int, int]) -> set[int]:
        """Индексы абзацев первой строки таблицы — по ней судим о «шапке»."""
        if not table.rows:
            return set()
        return {by_element[id(p)] for p in table.rows[0].element.iter(qn("w:p"))
                if id(p) in by_element}

    @staticmethod
    def _run_text(r_el: ET.Element) -> str:
        """Текст run: `w:t` + табуляции/разрывы как пробелы, поля (`w:instrText`) не в счёт."""
        out: list[str] = []
        for child in list(r_el):
            tag = dx.local_name(child.tag)
            if tag == "t":
                out.append(child.text or "")
            elif tag in ("tab", "br", "cr"):
                out.append(" ")
            elif tag == "sym":
                out.append(" ")
        return "".join(out)

    def _dominant_run(self, runs: list[RunInfo], base: RunProps) -> RunProps:
        """Итоговый шрифт/кегль абзаца по весу символов (заголовок из 1 run'а = его свойства)."""
        weights: dict[tuple[str | None, int | None], int] = {}
        for r in runs:
            if not r.text.strip():
                continue
            key = (r.props.font, r.props.size)
            weights[key] = weights.get(key, 0) + len(r.text)
        if not weights:
            return base
        (font, size), _ = max(weights.items(), key=lambda kv: kv[1])
        mono_set = {f.lower() for f in self.mono_fonts}
        return RunProps(font=font or base.font, size=size if size is not None else base.size,
                        bold=any(r.props.bold for r in runs if r.text.strip()) or base.bold,
                        italic=base.italic, superscript=base.superscript,
                        underline=base.underline,
                        mono=bool(font and font.lower() in mono_set))

    @staticmethod
    def _parse_drawings(el: ET.Element, p_index: int,
                        rels: dict[str, tuple[str, str]]) -> list[ImageInfo]:
        """`wp:inline|wp:anchor` → extent (EMU) + `a:blip/@r:embed` → медиа-часть."""
        out: list[ImageInfo] = []
        # Внутри w:drawing ровно один из вариантов: wp:inline (текст не обтекает) или wp:anchor.
        targets = [c for c in list(el) if dx.local_name(c.tag) in ("inline", "anchor")]
        for tgt in targets:
            extent = dx.find_child(tgt, "wp:extent")
            blip = None
            for candidate in tgt.iter(qn("a:blip")):
                blip = candidate
                break
            info = ImageInfo(paragraph_index=p_index, kind="drawing")
            if extent is not None:
                info.cx = dx.get_int(extent, "cx", 0) or 0
                info.cy = dx.get_int(extent, "cy", 0) or 0
            if blip is not None:
                rid = blip.get(qn("r:embed")) or blip.get(qn("r:link"))
                info.rid = rid
                if rid and rid in rels:
                    info.target = rels[rid][1]
            out.append(info)
        return out

    @staticmethod
    def _parse_picts(el: ET.Element, p_index: int,
                     rels: dict[str, tuple[str, str]]) -> list[ImageInfo]:
        """Устаревший VML (`w:pict`/`v:imagedata`) — тоже изображение, его считаем."""
        out: list[ImageInfo] = []
        for img in el.iter(qn("v:imagedata")):
            rid = img.get(qn("r:id"))
            out.append(ImageInfo(paragraph_index=p_index, kind="pict", rid=rid,
                                 target=rels[rid][1] if rid in rels else None))
        return out

    def _parse_sections(self, document: ET.Element,
                        rels: dict[str, tuple[str, str]]) -> list[SectionInfo]:
        """И `body/sectPr` (последняя секция), и `pPr/sectPr` (разрывы в середине)."""
        body = dx.find_child(document, "w:body")
        out: list[SectionInfo] = []
        if body is None:
            return out
        for p_el in body.iter(qn("w:p")):
            sect = dx.find_child(p_el, "w:pPr/w:sectPr")
            if sect is not None:
                out.append(self._section(sect, is_body_final=False, rels=rels))
        final = dx.find_child(body, "w:sectPr")
        if final is not None:
            out.append(self._section(final, is_body_final=True, rels=rels))
        return out

    def _section(self, sect: ET.Element, *, is_body_final: bool,
                 rels: dict[str, tuple[str, str]]) -> SectionInfo:
        info = SectionInfo(element=sect, is_body_final=is_body_final)
        pg_sz = dx.find_child(sect, "w:pgSz")
        if pg_sz is not None:
            info.pg_w = dx.get_int(pg_sz, "w:w")
            info.pg_h = dx.get_int(pg_sz, "w:h")
            info.orient = pg_sz.get(qn("w:orient"))
        pg_mar = dx.find_child(sect, "w:pgMar")
        if pg_mar is not None:
            for attr in ("w:top", "w:right", "w:bottom", "w:left", "w:header", "w:footer",
                         "w:gutter"):
                val = dx.get_int(pg_mar, attr)
                if val is not None:
                    info.mar[attr.split(":")[-1]] = val
        info.title_pg = dx.has_child(sect, "w:titlePg")
        pg_num = dx.find_child(sect, "w:pgNumType")
        if pg_num is not None:
            info.pg_num_start = dx.get_int(pg_num, "w:start")
        for ref in dx.iter_children(sect, "footerReference"):
            rtype = ref.get(qn("w:type")) or "default"
            rid = ref.get(qn("r:id"))
            if rid and rid in rels:
                info.footers[rtype] = rels[rid][1]
        for ref in dx.iter_children(sect, "headerReference"):
            rtype = ref.get(qn("w:type")) or "default"
            rid = ref.get(qn("r:id"))
            if rid and rid in rels:
                info.headers[rtype] = rels[rid][1]
        return info

    @staticmethod
    def _parse_ftr_parts(parts: MutableMapping[str, bytes],
                         rels: dict[str, tuple[str, str]],
                         ) -> tuple[dict[str, ET.Element], dict[str, ET.Element]]:
        """Парсим footer/header части пакета (в т.ч. не прикреплённые к sectPr — важно для NUM_POSITION)."""
        footers: dict[str, ET.Element] = {}
        headers: dict[str, ET.Element] = {}
        for name in dx.part_names(parts, "word/"):
            base = name.rsplit("/", 1)[-1]
            if base.startswith("footer") and base.endswith(".xml"):
                root = dx.parse_part(parts, name)
                if root is not None:
                    footers[name] = root
            elif base.startswith("header") and base.endswith(".xml"):
                root = dx.parse_part(parts, name)
                if root is not None:
                    headers[name] = root
        return footers, headers


# --- Вспомогательные функции для проверок/нормализатора ---------------------
def paragraph_text(p_el: ET.Element) -> str:
    """Текст абзаца без служебных элементов (для regex-распознавания подписей)."""
    return "".join(DocxParser._run_text(r) for r in p_el.iter(qn("w:r")))


def line_height_twips(para: ParagraphInfo, base_size_pt: float) -> float:
    """Высота строки в twips для оценки объёма/заполненности (эвристика).

    При `lineRule="auto"` `w:line` — доли одинарного интервала (240 = 1,0),
    при exact/atLeast — величина в 1/20 пункта, то есть численно равна twips.
    """
    size = (para.eff_run.size or pt_to_half_points(base_size_pt)) / 2.0
    single = size * 20.0 * 1.2
    line = para.eff.line or spacing_to_line(1.0)
    rule = (para.eff.line_rule or "auto").lower()
    if rule == "auto":
        return max(1.0, single * (line / 240.0))
    return max(1.0, float(line))


def text_area_twips(sect: SectionInfo | None, rules: dict[str, Any]) -> tuple[int, int]:
    """(ширина, высота) текстовой полосы в twips: из правил, поля секции — перекрывающе."""
    page = rules.get("page") or {}
    orientation = str(page.get("orientation") or "portrait")
    width, height = page_size(str(page.get("size") or "A4"), orientation)
    if sect is not None and sect.pg_w and sect.pg_h:
        width, height = sect.pg_w, sect.pg_h
    left = mm_to_twips(float(page.get("margin_left_mm", 30)))
    right = mm_to_twips(float(page.get("margin_right_mm", 15)))
    top = mm_to_twips(float(page.get("margin_top_mm", 20)))
    bottom = mm_to_twips(float(page.get("margin_bottom_mm", 20)))
    if sect and sect.mar:
        left = sect.mar.get("left", left)
        right = sect.mar.get("right", right)
        top = sect.mar.get("top", top)
        bottom = sect.mar.get("bottom", bottom)
    return max(100, width - left - right), max(200, height - top - bottom)


def avg_char_twips(size_half_points: int) -> float:
    """Средняя ширина символа для TNR-подобного шрифта ≈ половине кегля."""
    return max(1.0, size_half_points / 2.0 * 20.0 * 0.5)


@dataclass
class PageEstimate:
    """Результат эвристической разбивки на страницы (для location.page и VOLUME_FILL)."""

    page_of: dict[int, int]
    area_height: int
    page_used: list[float]

    @property
    def total_pages(self) -> int:
        return len(self.page_used)

    def fill_percent(self, page: int) -> float:
        """Заполненность страницы (1-based): занятые twips / высота полосы."""
        if page < 1 or page > len(self.page_used) or self.area_height <= 0:
            return 0.0
        return min(100.0, self.page_used[page - 1] / self.area_height * 100.0)

    def min_fill_excluding_last(self) -> float:
        """Минимальная заполненность среди всех страниц, кроме последней (она может быть короткой)."""
        body = self.page_used[:-1]
        if not body or self.area_height <= 0:
            return 100.0
        return min(100.0, min(body) / self.area_height * 100.0)


def estimate_pages(model: DocxModel, rules: dict[str, Any]) -> PageEstimate:
    """Номер страницы для каждого абзаца — ЭВРИСТИКА (точный layout знает только Word).

    Считаем высоту: строки абзаца (длина / ёмкость полосы) × высота строки; явные
    разрывы (`pageBreakBefore`, `w:br type="page"`, конец секции) обнуляют счётчик.
    Нужен для поля `location.page` и для info-проверки VOLUME_FILL.
    """
    area_w, area_h = text_area_twips(model.main_section, rules)
    body_size = pt_to_half_points(float((rules.get("text") or {}).get("size_pt", 14)))
    chars_per_line = max(15.0, area_w / avg_char_twips(body_size))
    table_size = float((rules.get("tables") or {}).get("font_size_pt",
                                                       (rules.get("text") or {}).get("size_pt", 14)))
    page_of: dict[int, int] = {}
    page_used: list[float] = []
    used = 0.0
    page = 1
    para_by_id = {id(p.element): p for p in model.paragraphs}

    def consume(height: float) -> None:
        """Добавляет блок на текущую страницу, при переполнении — новая страница."""
        nonlocal used, page
        if height <= 0:
            return
        if used + height > area_h and used > 0:
            page_used.append(used)
            page += 1
            used = 0.0
        used += height

    for kind, element in model.blocks:
        if kind == "tbl":
            table = next((t for t in model.tables if t.element is element), None)
            if table is None:
                continue
            line_h = line_height_twips(ParagraphInfo(index=-1, element=element), table_size)
            for row in table.rows:
                widths = [len(" ".join(cell)) for cell in row.cells] or [0]
                rows_in_row = max(1, math.ceil(max(widths) / max(8.0, chars_per_line / 4)))
                consume(rows_in_row * line_h + line_h * 0.15)  # + поля ячейки
            for p_i in table.paragraph_indexes:
                page_of[p_i] = page
            continue
        para = para_by_id.get(id(element))
        if para is None:
            continue
        if para.eff.page_break_before or para.has_page_break_run or para.has_sect_break:
            if used > 0:
                page_used.append(used)
                page += 1
                used = 0.0
        para_chars = max(8.0, area_w / avg_char_twips(para.eff_run.size or body_size))
        n_lines = max(1, math.ceil(len(para.text) / para_chars)) if para.text.strip() else 1
        consume(n_lines * line_height_twips(para, body_size / 2.0))
        page_of[para.index] = page

    page_used.append(used)
    return PageEstimate(page_of=page_of, area_height=area_h, page_used=page_used)


def paragraph_map(model: DocxModel) -> dict[int, ParagraphInfo]:
    """Быстрый доступ к абзацу по его номеру (для `location`)."""
    return {p.index: p for p in model.paragraphs}


def indent_cm_from_twips(value: int) -> float:
    """Обратный пересчёт для сообщений человеку."""
    return round(value / cm_to_twips(1.0), 3)


# --- Классификация абзацев (общая для инспектора и нормализатора) ----------
def rget(rules: Mapping[str, Any], path: str, default: Any = None) -> Any:
    """`rget(rules, "figures.caption_template")` — доступ к вложенному ключу правил."""
    node: Any = rules
    for seg in path.split("."):
        if not isinstance(node, Mapping) or seg not in node:
            return default
        node = node[seg]
    return node


def _flexi_literal(literal: str) -> str:
    """Литерал шаблона → regex с гибкими пробелами и любым тире (– — -)."""
    escaped = re.escape(literal)
    escaped = re.sub(r"(?:\\\s|\s)+", r"\\s+", escaped)
    return re.sub(r"[–—−-]", r"[-–—−]", escaped)


def template_to_regex(template: str, *, anchored: bool = True) -> re.Pattern[str]:
    """Регулярное выражение из шаблона правил (`Рисунок {n} – {title}`).

    `{n}` → число (сквозная или разделовая нумерация «1» / «1.2»),
    `{letter}` → буква приложения, `{title}`/`{page}` → текст.
    Пробелы и тире делаем гибкими: генераторы ставят дефис вместо андиша.
    """
    parts = re.split(r"(\{[a-zA-Z_]+\})", template)
    out: list[str] = []
    for part in parts:
        if part.startswith("{") and part.endswith("}"):
            key = part[1:-1]
            if key == "n":
                out.append(r"(\d{1,3}(?:\.\d{1,3})?)")
            elif key == "letter":
                out.append(r"([А-Яа-яA-Za-z]{1,2})")
            elif key == "title":
                out.append(r"(.*)")
            else:
                out.append(r"(.+?)")
        else:
            out.append(_flexi_literal(part))
    pattern = "".join(out)
    return re.compile((r"^" + pattern + r"$") if anchored else pattern,
                      re.IGNORECASE | re.DOTALL)


def caption_regexes(rules: Mapping[str, Any]) -> dict[str, re.Pattern[str]]:
    """Набор regex'ов подписей из правил; при отсутствии секции — пустой dict (чек отключён)."""
    out: dict[str, re.Pattern[str]] = {}
    fig = rget(rules, "figures.caption_template")
    if fig:
        out["figure"] = template_to_regex(fig)
    tbl = rget(rules, "tables.caption_template")
    if tbl:
        out["table"] = template_to_regex(tbl)
    cont = rget(rules, "tables.continuation_text")
    if cont:
        out["continuation"] = template_to_regex(cont, anchored=False)
    note = rget(rules, "notes.prefix")
    if note:
        out["note"] = re.compile(rf"^{re.escape(note)}\s*[-–—]?\s*\d*\.?\s*", re.IGNORECASE)
    app = rget(rules, "appendices.label")
    if app:
        out["appendix"] = template_to_regex(app)
    return out


def is_structural_heading(text: str, rules: Mapping[str, Any] | None = None) -> bool:
    """Абзац — структурный элемент (РЕФЕРАТ, ВВЕДЕНИЕ, ПРИЛОЖЕНИЕ …)?

    Список — `constants.STRUCTURAL_HEADINGS`. Совпадение строгое: либо ровно элемент,
    либо элемент + короткий хвост через пробел/дефис («ПРИЛОЖЕНИЕ А», «РАЗДЕЛ 2»).
    Иначе обычный абзац текста, начинающийся со слова «Приложение содержит…»,
    ложно считался бы заголовком.
    """
    from app import constants  # лениво: константы — единственный источник списка

    clean = " ".join(text.split()).strip().rstrip(".").upper()
    if not clean or len(clean) > 90:
        return False
    for head in constants.STRUCTURAL_HEADINGS:
        if clean == head:
            return True
        if clean.startswith(head) and len(clean) - len(head) <= 6:
            tail = clean[len(head):]
            if tail and tail[0] in " -–—\t":
                return True
    return False


def heading_level(text: str, para: ParagraphInfo) -> int | None:
    """Уровень заголовка 1..4 по номеру («1» → 1, «1.2.3» → 3) либо по стилю."""
    style = (para.eff.style or para.props.style or "").lower()
    match = re.search(r"(?:heading|заголов)\s*(\d)", style)
    numbered = re.match(r"^\s*(\d{1,2}(?:\.\d{1,2})*)\s+\S", text)
    if numbered:
        return min(4, numbered.group(1).count(".") + 1)
    if match:
        return min(4, max(1, int(match.group(1))))
    if style in ("title", "Heading1".lower(), "heading"):
        return 1
    return None


def looks_like_heading(para: ParagraphInfo, rules: Mapping[str, Any] | None = None) -> bool:
    """Заголовок раздела/подраздела: стиль, номер в начале или короткий жирный без точки."""
    if para.in_table or para.is_empty or para.instr:
        return False  # абзац с полем (TOC/PAGE) — не заголовок, даже если начинается с цифры
    text = para.text.strip()
    if len(text) > 160 or para.is_code:
        return False
    if is_structural_heading(text, rules):
        return False
    if heading_level(text, para) is not None:
        return True
    style = (para.eff.style or "").lower()
    if "heading" in style or "загол" in style or "title" in style:
        return True
    return bool(para.eff_run.bold) and len(text) <= 90 and not text.endswith(".")


def is_caption(text: str, regexes: Mapping[str, re.Pattern[str]]) -> str | None:
    """К какому виду подписей относится текст (figure/table/continuation/note/appendix)."""
    clean = " ".join(text.split()).strip()
    for kind, rx in regexes.items():
        if rx.match(clean):
            return kind
    return None


def first_structural_index(model: DocxModel,
                           rules: Mapping[str, Any] | None = None) -> int | None:
    """Индекс первого структурного элемента (РЕФЕРАТ и т.п.) — граница титульного листа.

    Всё, что раньше него, оформляется по шаблону вуза, а не как текст работы: такие
    абзацы исключаются из статистики основного текста и из проверки заголовков.
    """
    return min((p.index for p in model.paragraphs
                if not p.in_table and is_structural_heading(p.text, rules)),
               default=None)


def main_paragraphs(model: DocxModel, rules: Mapping[str, Any],
                    regexes: Mapping[str, re.Pattern[str]] | None = None,
                    ) -> list[ParagraphInfo]:
    """Абзацы ОСНОВНОГО текста: вне таблиц, вне код-блоков, не подписи и не заголовки.

    Исключаются также:
    * титульный лист (всё, что до первого структурного элемента) — он форматируется
      по шаблону вуза, а не как текст работы;
    * абзацы с полями (TOC/PAGE): «1 Раздел……5» из содержания перебило бы статистику.

    Именно по этому множеству считаются шрифт/кегль/интервал/отступ — иначе 12-pt
    таблицы и Consolas перебивают статистику (дефекты №1–2 старого чекера, п. 11.2 ТЗ).
    """
    rx = regexes if regexes is not None else caption_regexes(rules)
    structural_start = first_structural_index(model, rules)
    out: list[ParagraphInfo] = []
    for para in model.paragraphs:
        text = para.text.strip()
        if not text or para.in_table or para.is_code:
            continue
        if structural_start is not None and para.index < structural_start:
            continue
        if para.instr:
            continue
        if is_caption(text, rx):
            continue
        if is_structural_heading(text, rules) or looks_like_heading(para, rules):
            continue
        out.append(para)
    return out


def dominant(values: Iterable[Any]) -> Any:
    """Самое частое значение (реальный «mode» документа, а не «первое попавшееся»)."""
    counts: dict[Any, int] = {}
    for value in values:
        counts[value] = counts.get(value, 0) + 1
    if not counts:
        return None
    return max(counts.items(), key=lambda kv: (kv[1], str(kv[0])))[0]


def dominant_weighted(pairs: Iterable[tuple[Any, int]]) -> Any:
    """Самое частое значение с весом (весом служат символы текста — длинные абзацы важнее)."""
    totals: dict[Any, int] = {}
    for value, weight in pairs:
        if value is None:
            continue
        totals[value] = totals.get(value, 0) + weight
    if not totals:
        return None
    return max(totals.items(), key=lambda kv: (kv[1], str(kv[0])))[0]


def snippet_of(text: str, limit: int = 120) -> str:
    """Фрагмент для отчёта: не больше `limit` символов (иначе таблица замусоривается)."""
    clean = " ".join((text or "").split()).strip()
    return clean if len(clean) <= limit else clean[: limit - 1] + "…"

