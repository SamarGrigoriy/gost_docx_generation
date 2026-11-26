#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_games.py — подборка дидактических/музыкальных игр (педагогический пресет).

Запуск:
    python3 -I scripts/make_games.py --out games.docx \
        [--rules presets/pedagogical.json] [--meta meta.json] [--params params.json]

Контент — только из параметров/--meta; при пустых параметрах генерируется структура
с пустыми заполняемыми полями (содержание не выдумывается).

Структура: шапка титула → ВВЕДЕНИЕ (intro) → таблица игр (название, цель, ход/правила)
→ заключение → приложение.
"""

from __future__ import annotations

import argparse
from typing import Any
import json
import os
import sys
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

SCRIPT_META = {
    "name": "Подборка дидактических и музыкальных игр",
    "description": "Подборка игр для занятий: альбомная A4, таблица игр "
                   "(цель/ход/правила), введение, приложение. Контент из параметров.",
    "category": "Педагогика",
    "version": "1.0",
    "input": "none",
    "output_ext": ".docx",
    "arg_template": ["--out", "{output}", "--rules", "{rules_json}", "--meta", "{meta_json}"],
    "rules_hook": "postprocess",
    "params": [
        {"id": "theme", "label": "Тема подборки", "type": "text", "default": ""},
        {"id": "age_group", "label": "Возрастная группа", "type": "select",
         "options": ["2-3", "3-4", "4-5", "5-6", "6-7"], "default": "3-4"},
        {"id": "intro", "label": "Введение (актуальность, задачи)", "type": "textarea",
         "default": ""},
        {"id": "games", "label": "Игры: JSON-список или строки «название|цель|ход|правила»",
         "type": "textarea", "default": ""},
        {"id": "concluding", "label": "Заключение", "type": "textarea", "default": ""},
        {"id": "appendix", "label": "Приложение (материалы, атрибуты)", "type": "textarea",
         "default": ""},
    ],
}

# Константы педагогического пресета; apply_rules пересчитывает из --rules
TWIPS_PER_MM = 56.6929
LINE_UNIT = 240
PT14, PT12 = 28, 24
LINE10 = 240
TNR = "Times New Roman"
PAGE_W, PAGE_H = 16838, 11906   # альбомная A4
MAR = 1134                      # 20 мм
IND = 0                         # абзацный отступ текста, twips (из правил text.first_line_indent_cm)
PT_HEAD = PT14
PT_TBL = PT12
LINE = LINE10
BDR = 4
FONT = TNR

DEFAULT_RULES: dict = {
    "page": {"size": "A4", "orientation": "landscape", "margin_left_mm": 20,
             "margin_right_mm": 20, "margin_top_mm": 20, "margin_bottom_mm": 20},
    "text": {"font": TNR, "size_pt": 14, "line_spacing": 1.0, "first_line_indent_cm": 0},
    "tables": {"font_size_pt": 12, "line_spacing": 1.0, "repeat_header": True,
               "header_bold": True, "border_size_eighths_pt": 4},
}


def apply_rules(rules: dict) -> None:
    """Пересчитать константы оформления из пресета (rules_hook postprocess совместим)."""
    global PT_HEAD, PT_TBL, LINE, PAGE_W, PAGE_H, MAR, BDR, FONT, IND
    page = rules.get("page") or {}
    m = page.get("margin_left_mm")
    if m is not None:
        MAR = int(round(m * TWIPS_PER_MM))
    if str(page.get("orientation", "landscape")).lower() != "landscape":
        PAGE_W, PAGE_H = PAGE_H, PAGE_W
    text = rules.get("text") or {}
    if text.get("size_pt"):
        PT_HEAD = int(round(float(text["size_pt"]) * 2))
    FONT = str(text.get("font", TNR))
    if text.get("line_spacing"):
        LINE = int(round(float(text["line_spacing"]) * LINE_UNIT))
    if "first_line_indent_cm" in text:
        IND = int(round(float(text["first_line_indent_cm"]) * 10 * TWIPS_PER_MM))
    tbl = rules.get("tables") or {}
    if tbl.get("font_size_pt"):
        PT_TBL = int(round(float(tbl["font_size_pt"]) * 2))
    if tbl.get("border_size_eighths_pt") is not None:
        BDR = int(tbl["border_size_eighths_pt"])


NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def esc(t: object) -> str:
    return escape(str(t), {'"': '&quot;'})


def run(t: str, sz: int | None = None, bold: bool = False) -> str:
    sz = sz if sz is not None else PT_HEAD
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
            f'{"<w:b/>" if bold else ""}'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(t)}</w:t></w:r>')


def para(text: str = "", sz: int | None = None, bold: bool = False, align: str = "both",
         line: int | None = None, keep: bool = False, ind: int = 0,
         brk: bool = False, style: str | None = None) -> str:
    """pPr-порядок: pStyle → keepNext → pageBreakBefore → spacing → ind → jc; w:t с xml:space (в run)."""
    line = line if line is not None else LINE
    p = ["<w:pPr>"]
    if style:
        p.append(f'<w:pStyle w:val="{style}"/>')
    if keep:
        p.append("<w:keepNext/>")
    if brk:
        p.append("<w:pageBreakBefore/>")
    p.append(f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>')
    p.append(f'<w:ind w:firstLine="{ind}"/>')
    p.append(f'<w:jc w:val="{align}"/></w:pPr>')
    return f"<w:p>{''.join(p)}{run(text, sz, bold) if text else ''}</w:p>"


def structural(text: str) -> str:
    """Структурный элемент: с новой страницы, ПРОПИСНЫЕ, полужирный, по центру, без точки.

    pStyle обязателен: без него Word не построит автосодержание и панель навигации.
    """
    return para(text.upper(), bold=True, align="center", keep=True, brk=True,
                style="Heading1")


def cell(width: int, text: str, sz: int | None = None, bold: bool = False,
         align: str = "left", span: int = 1) -> str:
    sz = sz if sz is not None else PT_TBL
    tc = f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>'
    if span > 1:
        tc += f'<w:gridSpan w:val="{span}"/>'
    tc += '<w:vAlign w:val="center"/></w:tcPr>'
    lines = [l for l in str(text).split("\n") if l.strip()] or [""]
    tc += "".join(para(l, sz=sz, bold=bold, align=align) for l in lines)
    return tc + "</w:tc>"


def table(col_w: list[int], rows_xml: list[str]) -> str:
    total = sum(col_w)
    borders = "".join(f'<w:{s} w:val="single" w:sz="{BDR}" w:space="0" w:color="000000"/>'
                      for s in ("top", "left", "bottom", "right", "insideH", "insideV"))
    return ('<w:tbl><w:tblPr>'
            f'<w:tblW w:w="{total}" w:type="dxa"/><w:tblBorders>{borders}</w:tblBorders>'
            '<w:tblLayout w:type="fixed"/></w:tblPr><w:tblGrid>'
            + "".join(f'<w:gridCol w:w="{w}"/>' for w in col_w) + "</w:tblGrid>"
            + "".join(rows_xml) + "</w:tbl>" + para(""))


def _rows_from_items(data: list[Any]) -> list[list[str]]:
    """Список игр → строки таблицы: объекты, списки или строки «название|цель|ход|правила»."""
    out: list[list[str]] = []
    for item in data:
        if isinstance(item, dict):
            out.append([str(item.get("name", item.get("название", ""))),
                        str(item.get("goal", item.get("цель", ""))),
                        str(item.get("course", item.get("ход", ""))),
                        str(item.get("rules", item.get("правила", "")))])
        elif isinstance(item, (list, tuple)):
            row = [str(c) for c in item][:4]
            out.append(row + [""] * (4 - len(row)))
        elif isinstance(item, str):
            parts = [p.strip() for p in item.split("|")]
            out.append((parts + ["", "", "", ""])[:4])
    return out


def parse_games(raw: Any) -> list[list[str]]:
    """Игры: список (уже распарсенный из --meta JSON), JSON-строка или строки с «|».

    Значение из меты приходит уже структурированным — str() по нему ломает содержимое,
    поэтому тип разбирается явно. Пустой ввод → одна пустая строка-заготовка.
    """
    if isinstance(raw, (list, tuple)):
        rows = _rows_from_items(list(raw))
        return rows or [["", "", "", ""]]
    if isinstance(raw, dict):
        rows = _rows_from_items(list(raw.values()))
        return rows or [["", "", "", ""]]
    text = str(raw or "").strip()
    if text:
        if text.startswith("["):
            try:
                rows = _rows_from_items(json.loads(text))
                if rows:
                    return rows
            except json.JSONDecodeError:
                print("ВНИМАНИЕ: games — невалидный JSON, паршу как строки с «|»",
                      file=sys.stderr)
        rows = []
        for ln in text.splitlines():
            if not ln.strip():
                continue
            parts = [p.strip() for p in ln.split("|")]
            parts = (parts + ["", "", "", ""])[:4]
            rows.append(parts)
        if rows:
            return rows
    return [["", "", "", ""]]   # пустая заполняемая строка


def build(out_path: str, meta: dict, do_verify: bool = True) -> int:
    g = lambda k, d="": str(meta.get(k, d) or d)
    raw = lambda k, d="": meta.get(k, d)
    body: list[str] = []

    # --- шапка титула (из меты; пустые реквизиты не печатаем) ---
    for key in ("ministry", "fsbe", "univ", "kaf"):
        v = g(key)
        if v:
            body.append(para(v, sz=PT12, align="center"))
    title = "ПОДБОРКА ДИДАКТИЧЕСКИХ И МУЗЫКАЛЬНЫХ ИГР"
    if g("theme"):
        title = f"{title}\nпо теме: «{g('theme')}»"
    for tl in title.split("\n"):
        body.append(para(tl, bold=True, align="center", keep=True))
    if g("age_group"):
        body.append(para(f"Возрастная группа: {g('age_group')} лет", align="center"))
    body.append(para(""))

    # --- введение ---
    body.append(structural("ВВЕДЕНИЕ"))
    intro = g("intro").strip()
    if intro:
        for pl in intro.splitlines():
            if pl.strip():
                body.append(para(pl.strip(), align="both", ind=IND))
    else:
        body.append(para("(введение заполняется педагогом)", align="both"))
        body.append(para(""))

    # --- таблица игр ---
    total = PAGE_W - 2 * MAR
    weights = [4, 22, 24, 26, 24]  # № | Название | Цель | Ход игры | Правила
    cols = [int(total * w / sum(weights)) for w in weights]
    cols[-1] += total - sum(cols)
    rows = ["<w:tr><w:trPr><w:tblHeader/></w:trPr>"
            + "".join(cell(w, t, bold=True, align="center") for w, t in
                      zip(cols, ["№", "Название игры", "Цель", "Ход игры", "Правила"]))
            + "</w:tr>"]
    for i, gr in enumerate(parse_games(raw("games")), 1):
        rows.append("<w:tr>" + cell(cols[0], str(i), align="center")
                    + "".join(cell(w, t) for w, t in zip(cols[1:], gr)) + "</w:tr>")
    body.append(table(cols, rows))

    # --- заключение / приложение ---
    body.append(structural("ЗАКЛЮЧЕНИЕ"))
    conc = g("concluding").strip()
    if conc:
        for pl in conc.splitlines():
            if pl.strip():
                body.append(para(pl.strip(), align="both", ind=IND))
    body.append(structural("ПРИЛОЖЕНИЕ"))
    app = g("appendix").strip()
    if app:
        for pl in app.splitlines():
            if pl.strip():
                body.append(para(pl.strip(), align="both", ind=IND))

    sect = ('<w:sectPr><w:footerReference w:type="default" r:id="rIdFtr"/>'
            f'<w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}" w:orient="landscape"/>'
            f'<w:pgMar w:top="{MAR}" w:right="{MAR}" w:bottom="{MAR}" w:left="{MAR}" '
            'w:header="709" w:footer="680" w:gutter="0"/>'
            '<w:docGrid w:linePitch="360"/></w:sectPr>')
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f'<w:document {NS}><w:body>{"".join(body)}{sect}</w:body></w:document>')
    write_docx(out_path, document)
    if do_verify and not verify(out_path):
        print("Самопроверка выявила замечания", file=sys.stderr)
    return 0


def write_docx(out_path: str, document: str) -> None:
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                     '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                     '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                     "</Types>")
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                 "</Relationships>")
    base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                f'<Relationship Id="rIdStyles" Type="{base}styles" Target="styles.xml"/>'
                f'<Relationship Id="rIdFtr" Type="{base}footer" Target="footer1.xml"/>'
                "</Relationships>")
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              f'<w:styles {NS}><w:docDefaults><w:rPrDefault>'
              f'<w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
              f'<w:sz w:val="{PT_HEAD}"/><w:szCs w:val="{PT_HEAD}"/></w:rPr></w:rPrDefault>'
              f'<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="{LINE}" '
              'w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
              f'<w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
              f'<w:sz w:val="{PT_HEAD}"/><w:szCs w:val="{PT_HEAD}"/></w:rPr></w:style>'
              # Heading1..4: кегль не больше основного (требование методички), keepNext,
              # интервал 1.0 для многострочных заголовков.
              + "".join(
                  f'<w:style w:type="paragraph" w:styleId="Heading{i}"><w:name w:val="heading {i}"/>'
                  f'<w:basedOn w:val="Normal"/><w:next w:val="Normal"/>'
                  f'<w:pPr><w:keepNext/><w:spacing w:before="0" w:after="0" w:line="{LINE10}" '
                  f'w:lineRule="auto"/><w:ind w:firstLine="0"/></w:pPr>'
                  f'<w:rPr><w:b/><w:sz w:val="{PT_HEAD}"/><w:szCs w:val="{PT_HEAD}"/></w:rPr></w:style>'
                  for i in (1, 2, 3, 4))
              + '</w:styles>')
    ftr_page = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                '<w:r><w:fldChar w:fldCharType="separate"/></w:r>' + run("1", sz=PT12)
                + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
    ftr = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<w:ftr {NS}>'
           f'<w:p><w:pPr><w:spacing w:before="0" w:after="0" w:line="{LINE}" '
           f'w:lineRule="auto"/><w:ind w:firstLine="0"/><w:jc w:val="center"/></w:pPr>'
           f'{ftr_page}</w:p></w:ftr>')
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles)
        z.writestr("word/footer1.xml", ftr)


def verify(path: str) -> bool:
    """Самопроверка: читаемость zip, валидность XML, секция и шапка таблицы игр."""
    ok = True
    try:
        with zipfile.ZipFile(path) as z:
            parts = {n: z.read(n) for n in z.namelist()}
    except Exception as exc:
        print(f"VERIFY FAIL: {exc}", file=sys.stderr)
        return False
    for name in ("word/document.xml", "word/styles.xml", "word/footer1.xml",
                 "[Content_Types].xml", "_rels/.rels", "word/_rels/document.xml.rels"):
        try:
            ET.fromstring(parts[name])
            print(f"VERIFY OK  : XML валиден: {name}", file=sys.stderr)
        except Exception as exc:
            print(f"VERIFY FAIL: {name}: {exc}", file=sys.stderr)
            ok = False
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    body = ET.fromstring(parts["word/document.xml"]).find(f"{W}body")
    sect = body.find(f"{W}sectPr")
    pg = sect.find(f"{W}pgSz") if sect is not None else None
    mar = sect.find(f"{W}pgMar") if sect is not None else None
    tbls = body.findall(f".//{W}tbl")
    checks = [
        (sect is not None and list(body)[-1] is sect, "sectPr — последний дочерний body"),
        (pg is not None and int(pg.get(f"{W}w", "0")) > int(pg.get(f"{W}h", "0")),
         "ориентация альбомная"),
        (mar is not None and all(int(mar.get(f"{W}{s}", "-1")) == MAR
                                 for s in ("top", "left", "right", "bottom")),
         f"поля {MAR} twips (20 мм)"),
        (bool(tbls) and tbls[0].find(f"{W}tr/{W}trPr/{W}tblHeader") is not None,
         "tblHeader в таблице игр"),
        (all(tc.find(f"{W}p") is not None for tc in body.iter(f"{W}tc")),
         "каждая ячейка содержит абзац"),
        # Структурные заголовки обязаны начинаться с новой страницы и иметь pStyle —
        # иначе автосодержание/навигация Word не работают (ТЗ §11.1.3, §11.1.6).
        (all(p.find(f"{W}pPr/{W}pageBreakBefore") is not None
             and p.find(f"{W}pPr/{W}pStyle") is not None
             for p in body.findall(f"{W}p")
             if p.find(f"{W}pPr/{W}pStyle") is not None),
         "структурные заголовки: pStyle + pageBreakBefore"),
    ]
    for good, msg in checks:
        print(f"VERIFY {'OK  ' if good else 'FAIL'}: {msg}", file=sys.stderr)
        ok = ok and good
    return ok


def load_json_arg(val: str | None) -> dict:
    if not val:
        return {}
    if os.path.isfile(val):
        with open(val, "r", encoding="utf-8") as fh:
            data = json.load(fh)
    else:
        data = json.loads(val)
    flat: dict = {}
    for k, v in data.items():
        if isinstance(v, dict):
            flat.update(v)
        else:
            flat[k] = v
    return flat


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="make_games.py",
                                 description="Подборка дидактических/музыкальных игр")
    ap.add_argument("--out", default="games.docx")
    ap.add_argument("--rules", default=None)
    ap.add_argument("--meta", default=None)
    ap.add_argument("--params", default=None)
    ap.add_argument("--no-verify", action="store_true")
    args = ap.parse_args(argv)

    rules = json.loads(json.dumps(DEFAULT_RULES))
    loaded = load_json_arg(args.rules)
    for key, val in loaded.items():
        if isinstance(val, dict) and isinstance(rules.get(key), dict):
            rules[key].update(val)
        else:
            rules[key] = val
    apply_rules(rules)

    meta = load_json_arg(args.meta)
    meta.update(load_json_arg(args.params))
    try:
        rc = build(args.out, meta, do_verify=not args.no_verify)
    except Exception as exc:
        print(f"ОШИБКА сборки: {exc!r}", file=sys.stderr)
        return 4
    print(args.out)
    return rc


if __name__ == "__main__":
    sys.exit(main())
