"""Программная сборка мини-DOCX для тестов инспектора/нормализатора.

Фикстуры генерируются, а не хранятся бинарниками (п. 15 ТЗ), и НЕ зависят от
`scripts/mkdocx.py` — тот пишет другой агент, тесты должны быть изолированы.

`make_docx(...)` собирает документ, соответствующий пресету gost_732, и умеет
вносить конкретные дефекты параметрами (поля, шрифт, отступ, интервал, позиция
подписи таблицы, колонтитул с номером) — ровно набор из п. 11.2 ТЗ.
"""

from __future__ import annotations

import datetime as _dt
import zipfile
from dataclasses import dataclass, field
from typing import Any, Iterable, Sequence

from xml.sax.saxutils import escape


def esc(text: str) -> str:
    """Экранирование текста/атрибутов OOXML (п. 7 CHATBOT_TEMPLATE: & < > « »)."""
    return escape(str(text), {'"': "&quot;"})


PNG_BYTES = bytes.fromhex(
    "89504e470d0a1a0a0000000d49484452000000010000000108060000001f15c4"
    "890000000d49444154789c6360000000020001e221bc330000000049454e44ae426082"
)

DOC_NS = (
    'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
    'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
    'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" '
    'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" '
    'xmlns:v="urn:schemas-microsoft-com:vml" '
    'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" '
    'xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"'
)


# --- Низкоуровневые блоки --------------------------------------------------
def run(text: str, *, font: str | None = None, size: int | None = None,
        bold: bool = False, italic: bool = False, sup: bool = False) -> str:
    """`w:r`: порядок детей rPr строго rFonts → b → i → vertAlign → sz → szCs (п. 5 шаблона)."""
    props: list[str] = []
    if font:
        props.append(f'<w:rFonts w:ascii="{esc(font)}" w:hAnsi="{esc(font)}" w:cs="{esc(font)}"/>')
    if bold:
        props.append("<w:b/>")
    if italic:
        props.append("<w:i/>")
    if sup:
        props.append('<w:vertAlign w:val="superscript"/>')
    if size:
        props.append(f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>')
    rpr = f"<w:rPr>{''.join(props)}</w:rPr>" if props else ""
    # xml:space="preserve" обязателен: иначе Word съедал бы ведущие/двойные пробелы
    return f'<w:r>{rpr}<w:t xml:space="preserve">{esc(text)}</w:t></w:r>'


def para(content: str | Sequence[str] = "", *, jc: str | None = None, indent: int | None = None,
         line: int | None = None, keep: bool = False, brk: bool = False,
         style: str | None = None, before: int = 0, after: int = 0, raw: str = "") -> str:
    """Абзац: `w:pPr` в законном порядке pStyle→keepNext→pageBreakBefore→spacing→ind→jc."""
    props: list[str] = []
    if style:
        props.append(f'<w:pStyle w:val="{esc(style)}"/>')
    if keep:
        props.append("<w:keepNext/>")
    if brk:
        props.append("<w:pageBreakBefore/>")
    if line is not None or before or after:
        line_attr = f' w:line="{line}" w:lineRule="auto"' if line is not None else ""
        props.append(f'<w:spacing w:before="{before}" w:after="{after}"{line_attr}/>')
    if indent is not None:
        props.append(f'<w:ind w:firstLine="{indent}" w:left="0" w:right="0"/>')
    if jc:
        props.append(f'<w:jc w:val="{esc(jc)}"/>')
    ppr = f"<w:pPr>{''.join(props)}</w:pPr>" if props else ""
    if isinstance(content, str):
        body = raw or (run(content) if content else "")
    else:
        body = "".join(content)
    return f"<w:p>{ppr}{body}</w:p>"


def field_page(runs_text: str = "2", *, sz: int = 24, font: str = "Times New Roman") -> str:
    """Поле PAGE: begin → instrText → separate → кэш-результат → end (п. 5 шаблона)."""
    rpr = (f'<w:rPr><w:rFonts w:ascii="{font}" w:hAnsi="{font}"/>'
           f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>')
    return (
        f'<w:r>{rpr}<w:fldChar w:fldCharType="begin"/></w:r>'
        f'<w:r>{rpr}<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="separate"/></w:r>'
        f'<w:r>{rpr}<w:t xml:space="preserve">{esc(runs_text)}</w:t></w:r>'
        f'<w:r>{rpr}<w:fldChar w:fldCharType="end"/></w:r>'
    )


def toc_field() -> str:
    """Поле автосодержания (instrText TOC) — признак TOC_PRESENT."""
    return (
        '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
        '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
        '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
        f'{run("1 Раздел……5")}'
        '<w:r><w:fldChar w:fldCharType="end"/></w:r>'
    )


def drawing(rid: str, *, cx: int = 5400000, cy: int = 3240000, doc_pr: int = 100001) -> str:
    """`w:drawing`/`wp:inline` с extent в EMU и `a:blip/@r:embed` на медиа-часть."""
    return (
        f'<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
        f'<wp:extent cx="{cx}" cy="{cy}"/>'
        f'<wp:docPr id="{doc_pr}" name="Рисунок {doc_pr}"/>'
        f'<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
        f'<pic:pic><pic:nvPicPr><pic:cNvPr id="{doc_pr}" name="img.png"/><pic:cNvPicPr/>'
        f'</pic:nvPicPr><pic:blipFill><a:blip r:embed="{rid}"/><a:stretch><a:fillRect/>'
        f'</a:stretch></pic:blipFill>'
        f'<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
        f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic>'
        f'</a:graphicData></a:graphic></wp:inline></w:drawing>'
    )


def image_para(rid: str, *, jc: str = "center", doc_pr: int = 100001,
               keep: bool = True) -> str:
    """Абзац с рисунком: keepNext держит подпись под ним на той же странице."""
    keep_xml = "<w:keepNext/><w:keepLines/>" if keep else ""
    return (f'<w:p><w:pPr>{keep_xml}<w:jc w:val="{esc(jc)}"/></w:pPr>'
            f'<w:r>{drawing(rid, doc_pr=doc_pr)}</w:r></w:p>')


def table(rows: Sequence[Sequence[str]], *, borders: bool = True, header: bool = True,
          bold_header: bool = True, font: str = "Times New Roman", size: int = 24,
          line: int = 240, widths: Sequence[int] | None = None,
          layout: str = "fixed") -> str:
    """Таблица: порядок tblPr→tblGrid→tr; в tr: trPr→tc; каждая tc содержит ≥1 p."""
    cols = len(rows[0]) if rows else 1
    widths = list(widths or [9355 // cols] * cols)
    border_xml = ""
    if borders:
        sides = "".join(f'<w:{s} w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
                        for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
        border_xml = f"<w:tblBorders>{sides}</w:tblBorders>"
    tbl_pr = (f'<w:tblPr><w:tblW w:w="{sum(widths)}" w:type="dxa"/>{border_xml}'
              f'<w:tblLayout w:type="{esc(layout)}"/></w:tblPr>')
    grid = "<w:tblGrid>" + "".join(f'<w:gridCol w:w="{w}"/>' for w in widths) + "</w:tblGrid>"
    body: list[str] = []
    for row_i, row in enumerate(rows):
        trpr = "<w:trPr><w:tblHeader/></w:trPr>" if (row_i == 0 and header) else ""
        cells: list[str] = []
        for col_i, text in enumerate(row):
            is_head = row_i == 0 and header
            p = para(text, jc="center" if is_head else "left", indent=0, line=line,
                     raw=run(text, font=font, size=size, bold=(bold_header and is_head)))
            cells.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths[col_i]}" w:type="dxa"/>'
                         f'<w:vAlign w:val="center"/></w:tcPr>{p}</w:tc>')
        body.append(f"<w:tr>{trpr}{''.join(cells)}</w:tr>")
    return f"<w:tbl>{tbl_pr}{grid}{''.join(body)}</w:tbl>"


# --- Части пакета ----------------------------------------------------------
def styles_xml(*, font: str = "Times New Roman", size: int = 28, line: int = 360,
               indent: int = 709, jc: str = "both", broken_normal_basedon: bool = False,
               heading_keep: bool = True) -> str:
    """styles.xml: docDefaults + Normal + Heading1..4 (имена — встроенные, для TOC/навигации)."""
    based_on = '<w:basedOn w:val="Normal"/>' if broken_normal_basedon else ""
    keep = "<w:keepNext/>" if heading_keep else ""   # без него заголовок может оторваться от текста
    heads = ""
    for lvl in (1, 2, 3, 4):
        heads += (
            f'<w:style w:type="paragraph" w:styleId="Heading{lvl}">'
            f'<w:name w:val="heading {lvl}"/><w:basedOn w:val="Normal"/>'
            f'<w:next w:val="Normal"/><w:qFormat/>'
            f'<w:pPr>{keep}<w:spacing w:before="0" w:after="0" w:line="240" '
            f'w:lineRule="auto"/><w:ind w:firstLine="{indent}"/><w:jc w:val="left"/>'
            f'<w:outlineLvl w:val="{lvl - 1}"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{esc(font)}" w:hAnsi="{esc(font)}"/>'
            f'<w:b/><w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr></w:style>'
        )
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:styles {DOC_NS}>'
        '<w:docDefaults><w:rPrDefault><w:rPr>'
        f'<w:rFonts w:ascii="{esc(font)}" w:hAnsi="{esc(font)}" w:cs="{esc(font)}"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/>'
        f'<w:lang w:val="ru-RU"/></w:rPr></w:rPrDefault>'
        f'<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="{line}" '
        f'w:lineRule="auto"/><w:ind w:firstLine="{indent}"/><w:jc w:val="{esc(jc)}"/></w:pPr>'
        '</w:pPrDefault></w:docDefaults>'
        f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal">{based_on}'
        f'<w:name w:val="Normal"/><w:qFormat/>'
        f'<w:pPr><w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>'
        f'<w:ind w:firstLine="{indent}"/><w:jc w:val="{esc(jc)}"/></w:pPr>'
        f'<w:rPr><w:rFonts w:ascii="{esc(font)}" w:hAnsi="{esc(font)}" w:cs="{esc(font)}"/>'
        f'<w:sz w:val="{size}"/><w:szCs w:val="{size}"/></w:rPr></w:style>'
        f'<w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">'
        f'<w:name w:val="Default Paragraph Font"/></w:style>'
        '<w:style w:type="table" w:default="1" w:styleId="TableNormal">'
        '<w:name w:val="Normal Table"/></w:style>'
        '<w:style w:type="numbering" w:default="1" w:styleId="NoList">'
        '<w:name w:val="No List"/></w:style>'
        f'{heads}</w:styles>'
    )


def settings_xml(*, update_fields: bool = False) -> str:
    """settings.xml: w:updateFields нужен, чтобы TOC обновился при открытии."""
    upd = '<w:updateFields w:val="true"/>' if update_fields else ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:settings {DOC_NS}>{upd}'
            '<w:defaultTabStop w:val="708"/><w:characterSpacingControl w:val="compressPunctuation"/>'
            '<w:compat><w:compatSetting w:name="compatibilityMode" '
            'w:namespace="http://schemas.microsoft.com/office/word" w:val="15"/></w:compat>'
            '</w:settings>')


def footer_xml(*, with_page: bool = True, jc: str = "center", font: str = "Times New Roman",
               size: int = 24) -> str:
    """Нижний колонтитул: номер по центру (позиция bottom_center)."""
    body = field_page("2", sz=size, font=font) if with_page else ""
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:ftr {DOC_NS}><w:p><w:pPr><w:spacing w:before="0" w:after="0"/>'
            f'<w:ind w:firstLine="0"/><w:jc w:val="{esc(jc)}"/></w:pPr>{body}</w:p></w:ftr>')


def header_xml(*, with_page: bool = True, jc: str = "center") -> str:
    """Верхний колонтитул — используется дефектом «номер сверху»."""
    body = field_page("2") if with_page else run("Название работы")
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            f'<w:hdr {DOC_NS}><w:p><w:pPr><w:spacing w:before="0" w:after="0"/>'
            f'<w:ind w:firstLine="0"/><w:jc w:val="{esc(jc)}"/></w:pPr>{body}</w:p></w:hdr>')


def footnotes_xml(*, size: int = 24, font: str = "Times New Roman") -> str:
    """word/footnotes.xml: разделители + одна сноска заданного кегля."""
    return (
        f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        f'<w:footnotes {DOC_NS}>'
        '<w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" '
        'w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote>'
        '<w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing '
        'w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/>'
        '</w:r></w:p></w:footnote>'
        f'<w:footnote w:id="1"><w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="240" '
        f'w:lineRule="auto"/><w:ind w:firstLine="283"/></w:pPr>'
        f'{run("1 ", size=size, font=font, sup=True)}'
        f'{run("Текст сноски.", size=size, font=font)}'
        '</w:p></w:footnote></w:footnotes>'
    )


def sect_pr(*, left: int = 1701, right: int = 850, top: int = 1134, bottom: int = 1134,
            header: int = 709, footer: int = 709, page_w: int = 11906, page_h: int = 16838,
            orient: str | None = None, title_pg: bool = True, pg_start: int | None = None,
            rids: dict[str, str] | None = None) -> str:
    """`w:sectPr` в законном порядке: refs → pgSz → pgMar → pgNumType → titlePg (п. 5 шаблона)."""
    rids = rids or {}
    refs = ""
    if rids.get("header"):
        refs += f'<w:headerReference w:type="default" r:id="{rids["header"]}"/>'
    if rids.get("header_first"):
        refs += f'<w:headerReference w:type="first" r:id="{rids["header_first"]}"/>'
    if rids.get("footer"):
        refs += f'<w:footerReference w:type="default" r:id="{rids["footer"]}"/>'
    if rids.get("footer_first"):
        refs += f'<w:footerReference w:type="first" r:id="{rids["footer_first"]}"/>'
    orient_attr = f' w:orient="{esc(orient)}"' if orient else ""
    pg = f'<w:pgSz w:w="{page_w}" w:h="{page_h}"{orient_attr}/>'
    mar = (f'<w:pgMar w:top="{top}" w:right="{right}" w:bottom="{bottom}" w:left="{left}" '
           f'w:header="{header}" w:footer="{footer}" w:gutter="0"/>')
    num = f'<w:pgNumType w:start="{pg_start}"/>' if pg_start is not None else ""
    title = "<w:titlePg/>" if title_pg else ""
    return f"<w:sectPr>{refs}{pg}{mar}{num}{title}</w:sectPr>"


def content_types_xml(parts: Iterable[str]) -> str:
    """[Content_Types].xml: Default на rels/xml/png + Override на каждую XML-часть."""
    types = {
        "/word/document.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml"
                              ".document.main+xml",
        "/word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml"
                            ".styles+xml",
        "/word/settings.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml"
                              ".settings+xml",
        "/word/footnotes.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml"
                               ".footnotes+xml",
        "/word/fontTable.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml"
                               ".fontTable+xml",
    }
    overrides = []
    for name in parts:
        if name.endswith(".xml") and not name.startswith("_rels"):
            part = "/" + name
            ctype = types.get(part)
            if ctype is None and name.startswith("word/footer"):
                ctype = ("application/vnd.openxmlformats-officedocument.wordprocessingml"
                         ".footer+xml")
            elif ctype is None and name.startswith("word/header"):
                ctype = ("application/vnd.openxmlformats-officedocument.wordprocessingml"
                         ".header+xml")
            if ctype:
                overrides.append(f'<Override PartName="{part}" ContentType="{ctype}"/>')
    return (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package'
        '.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Default Extension="png" ContentType="image/png"/>'
        f'{"".join(overrides)}</Types>'
    )


def document_rels_xml(entries: Sequence[tuple[str, str, str]]) -> str:
    """word/_rels/document.xml.rels: (Id, Type, Target)."""
    items = "".join(f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
                    for rid, rtype, target in entries)
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f'{items}</Relationships>')


ROOT_RELS = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
             '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument'
             '/2006/relationships/officeDocument" Target="word/document.xml"/>'
             '</Relationships>')

REL_TYPES = {
    "footer": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    "header": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
    "image": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    "styles": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
    "settings": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings",
    "footnotes": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes",
}


# --- Высокоуровневая фикстура ---------------------------------------------
@dataclass
class DocSpec:
    """Переключатели оформления/дефектов. Значения по умолчанию = ГОСТ 7.32 без нарушений."""

    # страница
    margin_left: int = 1701
    margin_right: int = 850
    margin_top: int = 1134
    margin_bottom: int = 1134
    page_w: int = 11906
    page_h: int = 16838
    orient: str | None = None
    footer_distance: int = 709
    # основной текст
    body_font: str = "Times New Roman"
    body_size: int = 28
    body_line: int = 360
    body_indent: int = 709
    body_jc: str = "both"
    # заголовки
    heading_style: bool = True
    heading_keep: bool = True
    heading_bold: bool = True
    heading_indent: int = 709
    structural_case: str = "upper"          # upper | title
    structural_bold: bool = True
    structural_jc: str = "center"
    structural_brk: bool = True
    # подписи
    table_caption_jc: str = "left"
    table_caption_indent: int = 0
    table_caption_keep: bool = True
    figure_caption_jc: str = "center"
    figure_caption_indent: int = 0
    # таблицы
    table_borders: bool = True
    table_header: bool = True
    table_header_bold: bool = True
    table_font_size: int = 24
    # нумерация страниц
    page_field_in: str = "footer"           # footer | header | none
    title_pg: bool = True
    pg_start: int | None = None
    # содержательные блоки
    with_toc: bool = True
    with_footnotes: bool = True
    footnote_size: int = 24
    with_code_block: bool = True
    fig_ref: bool = True
    tbl_ref: bool = True
    app_ref: bool = True
    fig_caption: bool = True
    tbl_caption: bool = True
    source_years: Sequence[int] | None = None
    update_fields: bool = True


def _head(text: str, s: DocSpec) -> str:
    """Абзац основного текста."""
    return para(text, jc=s.body_jc, indent=s.body_indent, line=s.body_line,
                raw=run(text, font=s.body_font, size=s.body_size))


def _structural(text: str, s: DocSpec, *, style: str | None = "Heading1") -> str:
    shown = text.upper() if s.structural_case == "upper" else text.title()
    return para(shown, jc=s.structural_jc, indent=0, line=s.body_line, keep=True,
                brk=s.structural_brk, style=style,
                raw=run(shown, font=s.body_font, size=s.body_size, bold=s.structural_bold))


def make_docx(path: str | Any, s: DocSpec | None = None, **overrides: Any) -> str:
    """Собрать DOCX по спецификации; возвращает путь. Дефекты включаются overrides'ами."""
    spec = s or DocSpec()
    for key, value in overrides.items():
        if not hasattr(spec, key):
            raise KeyError(f"Неизвестный переключатель фикстуры: {key}")
        setattr(spec, key, value)
    year_now = _dt.date.today().year

    def head(text: str) -> str:
        return _head(text, spec)

    def struct(text: str, style: str | None = "Heading1") -> str:
        return _structural(text, spec, style=style)
    years = list(spec.source_years or [year_now - 1, year_now - 2])

    body: list[str] = []
    # --- титульный лист (форматируется по шаблону вуза, в статистику текста не входит)
    for line in ("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
                 "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ",
                 "Отчёт по практическому занятию",
                 f"Геническ, {year_now}"):
        body.append(para(line, jc="center", indent=0, line=spec.body_line,
                         raw=run(line, font=spec.body_font, size=spec.body_size)))

    body.append(struct("РЕФЕРАТ"))
    body.append(head("В работе рассмотрены приёмы оформления учебных документов "
                      f"по методичке {year_now} года, приведены примеры расчётов и таблиц."))

    body.append(struct("СОДЕРЖАНИЕ"))
    if spec.with_toc:
        body.append(para("", jc="left", indent=0, line=spec.body_line, raw=toc_field()))

    body.append(struct("ВВЕДЕНИЕ"))
    body.append(head("Актуальность темы обусловлена требованиями методички к учебным "
                      "работам [1]. Целями занятия являются освоение приёмов оформления "
                      "и практическое применение правил [2]."))
    body.append(head("В соответствии с рисунком 1 показана структура документа, а данные "
                      "сведены в таблице 1. Полный текст приведён в приложении А."
                      if spec.fig_ref and spec.tbl_ref and spec.app_ref else "Общий обзор литературы."))
    if spec.with_code_block:
        body.append(para("def greet(name):", jc="left", indent=0, line=240,
                         raw=run("def greet(name):\n    return 'привет, ' + name",
                                 font="Consolas", size=20)))

    body.append(para("1 Порядок оформления", jc=spec.body_jc, indent=spec.heading_indent,
                     line=spec.body_line, keep=spec.heading_keep,
                     style="Heading1" if spec.heading_style else None,
                     raw=run("1 Порядок оформления", font=spec.body_font, size=spec.body_size,
                             bold=spec.heading_bold)))
    body.append(head("Каждый элемент документа подчиняется правилам, зафиксированным в "
                      "пресете оформления; отклонения недопустимы [1]."))
    body.append(image_para("rIdImg", jc="center"))
    if spec.fig_caption:
        body.append(para("Рисунок 1 – Структура документа", jc=spec.figure_caption_jc,
                         indent=spec.figure_caption_indent, line=spec.body_line,
                         raw=run("Рисунок 1 – Структура документа", font=spec.body_font,
                                 size=spec.body_size)))
    body.append(para("F = m · a (1)", jc="center", indent=0, line=spec.body_line,
                     raw=run("F = m · a (1)", font=spec.body_font, size=spec.body_size)))
    body.append(head("где m — масса тела, кг; a — ускорение, м/с²."))
    if spec.tbl_caption:
        body.append(para("Таблица 1 – Параметры", jc=spec.table_caption_jc,
                         indent=spec.table_caption_indent, line=spec.body_line,
                         keep=spec.table_caption_keep,
                         raw=run("Таблица 1 – Параметры", font=spec.body_font,
                                 size=spec.body_size)))
    body.append(table([["Показатель", "Значение"], ["Ширина поля", "30 мм"],
                       ["Кегль", "14 pt"]],
                      borders=spec.table_borders, header=spec.table_header,
                      bold_header=spec.table_header_bold, font=spec.body_font,
                      size=spec.table_font_size))
    body.append(para("1.1 Требования к тексту", jc=spec.body_jc, indent=spec.heading_indent,
                     line=spec.body_line, keep=spec.heading_keep,
                     style="Heading2" if spec.heading_style else None,
                     raw=run("1.1 Требования к тексту", font=spec.body_font,
                             size=spec.body_size, bold=spec.heading_bold)))
    body.append(head("Текст набирается с абзацного отступа, выравниванием по ширине, без"
                      " лишних интервалов до и после абзаца [2]."))
    if spec.with_footnotes:
        body.append(head("Отдельные термины выносятся в сноску.") +
                    '<w:p><w:pPr><w:ind w:firstLine="%d"/><w:spacing w:before="0" w:after="0" '
                    'w:line="%d" w:lineRule="auto"/><w:jc w:val="%s"/></w:pPr>'
                    '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="%d"/></w:rPr>'
                    '<w:t xml:space="preserve">Термин</w:t></w:r>'
                    '<w:r><w:rPr><w:rFonts w:ascii="%s" w:hAnsi="%s"/><w:sz w:val="%d"/>'
                    '<w:vertAlign w:val="superscript"/></w:rPr>'
                    '<w:footnoteReference w:id="1"/></w:r></w:p>'
                    % (spec.body_indent, spec.body_line, spec.body_jc, spec.body_font,
                       spec.body_font, spec.body_size, spec.body_font, spec.body_font,
                       spec.body_size))

    body.append(struct("ЗАКЛЮЧЕНИЕ"))
    body.append(head("Соблюдение правил оформления делает документ пригодным к сдаче "
                      "и проверке комиссиями."))

    body.append(struct("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ"))
    body.append(head(f"1. Иванов И. И. Основы оформления : учебник. — Москва : Наука, "
                      f"{years[0]}. — 120 с."))
    body.append(head(f"2. Петрова А. П. Методичка по практике. — Геническ : ФГБОУ ВО, "
                      f"{years[1]}. — 48 с."))

    body.append(struct("ПРИЛОЖЕНИЕ А"))
    body.append(head("Приложение содержит пример технологической карты занятия."))

    # --- секция и колонтитулы -------------------------------------------
    rids: dict[str, str] = {}
    rels: list[tuple[str, str, str]] = [
        ("rIdStyles", REL_TYPES["styles"], "styles.xml"),
        ("rIdSettings", REL_TYPES["settings"], "settings.xml"),
    ]
    parts: dict[str, str] = {}
    if spec.with_footnotes:
        rids["footnotes"] = "rIdNotes"
        rels.append(("rIdNotes", REL_TYPES["footnotes"], "footnotes.xml"))
        parts["word/footnotes.xml"] = footnotes_xml(size=spec.footnote_size,
                                                    font=spec.body_font)
    rels.append(("rIdImg", REL_TYPES["image"], "media/image1.png"))
    if spec.page_field_in in ("footer", "none"):
        rels.append(("rIdF1", REL_TYPES["footer"], "footer1.xml"))
        parts["word/footer1.xml"] = footer_xml(with_page=spec.page_field_in == "footer",
                                               font=spec.body_font)
        rids["footer"] = "rIdF1"
        if spec.title_pg and spec.page_field_in == "footer":
            rels.append(("rIdF2", REL_TYPES["footer"], "footer2.xml"))
            parts["word/footer2.xml"] = footer_xml(with_page=False)
            rids["footer_first"] = "rIdF2"
    if spec.page_field_in == "header":
        rels.append(("rIdH1", REL_TYPES["header"], "header1.xml"))
        parts["word/header1.xml"] = header_xml(with_page=True)
        rids["header"] = "rIdH1"
        rels.append(("rIdF1", REL_TYPES["footer"], "footer1.xml"))
        parts["word/footer1.xml"] = footer_xml(with_page=False)
        rids["footer"] = "rIdF1"

    sect = sect_pr(left=spec.margin_left, right=spec.margin_right, top=spec.margin_top,
                   bottom=spec.margin_bottom, footer=spec.footer_distance,
                   page_w=spec.page_w, page_h=spec.page_h, orient=spec.orient,
                   title_pg=spec.title_pg, pg_start=spec.pg_start, rids=rids)
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                f'<w:document {DOC_NS}><w:body>{"".join(body)}{sect}</w:body></w:document>')

    names = ["[Content_Types].xml", "_rels/.rels", "word/document.xml", "word/styles.xml",
             "word/settings.xml", "word/_rels/document.xml.rels"] + sorted(parts) + \
            ["word/media/image1.png"]
    all_parts: dict[str, bytes] = {
        "[Content_Types].xml": content_types_xml(names).encode("utf-8"),
        "_rels/.rels": ROOT_RELS.encode("utf-8"),
        "word/document.xml": document.encode("utf-8"),
        "word/styles.xml": styles_xml(font=spec.body_font, size=spec.body_size,
                                      line=spec.body_line, indent=spec.body_indent,
                                      jc=spec.body_jc,
                                      heading_keep=spec.heading_keep).encode("utf-8"),
        "word/settings.xml": settings_xml(update_fields=spec.update_fields).encode("utf-8"),
        "word/_rels/document.xml.rels": document_rels_xml(rels).encode("utf-8"),
        "word/media/image1.png": PNG_BYTES,
    }
    for name, xml in parts.items():
        all_parts[name] = xml.encode("utf-8")

    target = str(path)
    with zipfile.ZipFile(target, "w", zipfile.ZIP_DEFLATED) as zf:
        for name in names:
            zf.writestr(name, all_parts[name])
    return target

