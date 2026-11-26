#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_karta.py — технологическая карта занятия (педагогический пресет).

Запуск:
    python3 -I scripts/make_karta.py --out karta.docx \
        [--rules presets/pedagogical.json] [--meta meta.json] [--params params.json]

Контент ЦЕЛИКОМ берётся из параметров/--meta; при пустых параметрах генерируется
структура документа с пустыми заполняемыми полями — содержательные данные не выдумываются.

Оформление (§4/§5.8 ТЗ): альбомная A4, поля 20 мм, шапка TNR 14, таблица TNR 12
интервал 1,0, шапка таблицы полужирная + tblHeader, строки-разделители через gridSpan.
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import sys
import zipfile
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

SCRIPT_META = {
    "name": "Технологическая карта занятия",
    "description": "Тех. карта ООД по педагогическому пресету: альбомная A4, блоки "
                   "шапки + таблица «Ход занятия» 7 колонок с этапами 1 / 2 (2.1–2.3) / 3.",
    "category": "Педагогика",
    "version": "1.0",
    "input": "none",
    "output_ext": ".docx",
    "arg_template": ["--out", "{output}", "--rules", "{rules_json}", "--meta", "{meta_json}"],
    "rules_hook": "postprocess",
    "params": [
        {"id": "theme", "label": "Тема занятия", "type": "text", "default": ""},
        {"id": "age_group", "label": "Возрастная группа", "type": "select",
         "options": ["2-3", "3-4", "4-5", "5-6", "6-7"], "default": "3-4"},
        {"id": "duration_min", "label": "Длительность, мин", "type": "int", "default": 15},
        {"id": "goal", "label": "Цель (по Хуторскому)", "type": "textarea", "default": ""},
        {"id": "tasks", "label": "Задачи (об/разв/восп/доп, построчно)", "type": "textarea",
         "default": ""},
        {"id": "educational_area", "label": "Образовательная область", "type": "text",
         "default": ""},
        {"id": "vocab_work", "label": "Словарная работа", "type": "textarea", "default": ""},
        {"id": "planned_result", "label": "Планируемый результат", "type": "textarea",
         "default": ""},
        {"id": "prep_work", "label": "Подготовительная работа", "type": "textarea",
         "default": ""},
        {"id": "materials", "label": "Материалы и оборудование", "type": "textarea",
         "default": ""},
        {"id": "stages", "label": "Этапы хода (JSON-список строк таблицы)", "type": "textarea",
         "default": ""},
    ],
}

# Единицы OOXML; apply_rules пересчитывает модульные переменные из пресета
TWIPS_PER_MM = 56.6929
EMU_PER_CM = 360_000
LINE_UNIT = 240
PT14, PT12 = 28, 24
LINE10, LINE15 = 240, 360
TNR = "Times New Roman"
PAGE_W, PAGE_H = 16838, 11906          # A4 альбомная — ориентация педагогического пресета
MAR = 1134                             # 20 мм со всех сторон
PT_HEAD = PT14                         # шапка документа
PT_TBL = PT12                          # текст таблиц
LINE = LINE10                          # интервал 1,0
BDR = 4

DEFAULT_RULES: dict = {
    "page": {"size": "A4", "orientation": "landscape", "margin_left_mm": 20,
             "margin_right_mm": 20, "margin_top_mm": 20, "margin_bottom_mm": 20},
    "text": {"font": TNR, "size_pt": 14, "line_spacing": 1.0, "first_line_indent_cm": 0},
    "tables": {"font_size_pt": 12, "line_spacing": 1.0, "repeat_header": True,
               "header_bold": True, "border_size_eighths_pt": 4},
}


def apply_rules(rules: dict) -> None:
    """Пересчитать константы оформления из пресета (работает и без --rules)."""
    global PT_HEAD, PT_TBL, LINE, PAGE_W, PAGE_H, MAR, BDR, FONT
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
    tbl = rules.get("tables") or {}
    if tbl.get("font_size_pt"):
        PT_TBL = int(round(float(tbl["font_size_pt"]) * 2))
    if tbl.get("border_size_eighths_pt") is not None:
        BDR = int(tbl["border_size_eighths_pt"])


FONT = TNR
NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"')


def esc(t: object) -> str:
    return escape(str(t), {'"': '&quot;'})


def run(t: str, sz: int | None = None, bold: bool = False, italic: bool = False) -> str:
    """rPr-порядок: rFonts → b → i → sz → szCs (§5 шаблона)."""
    sz = sz if sz is not None else PT_HEAD
    return (f'<w:r><w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
            f'{"<w:b/>" if bold else ""}{"<w:i/>" if italic else ""}'
            f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/></w:rPr>'
            f'<w:t xml:space="preserve">{esc(t)}</w:t></w:r>')


def para(text: str = "", sz: int | None = None, bold: bool = False, align: str = "left",
         line: int | None = None, keep: bool = False) -> str:
    """Порядок pPr: keepNext → spacing → ind → jc; отступ 0 (педагогический пресет)."""
    line = line if line is not None else LINE
    p = ["<w:pPr>"]
    if keep:
        p.append("<w:keepNext/>")
    p.append(f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>')
    p.append('<w:ind w:firstLine="0"/>')
    p.append(f'<w:jc w:val="{align}"/></w:pPr>')
    return f"<w:p>{''.join(p)}{run(text, sz, bold) if text else ''}</w:p>"


def cell(width: int, text: str, sz: int | None = None, bold: bool = False,
         align: str = "left", span: int = 1, valign: str = "center") -> str:
    """Ячейка: tcW (+gridSpan) → vAlign; внутри минимум один w:p (§7.10)."""
    sz = sz if sz is not None else PT_TBL
    tc = f'<w:tc><w:tcPr><w:tcW w:w="{width}" w:type="dxa"/>'
    if span > 1:
        tc += f'<w:gridSpan w:val="{span}"/>'
    tc += f'<w:vAlign w:val="{valign}"/></w:tcPr>'
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


def header_row(col_w: list[int], titles: list[str]) -> str:
    """Шапка: trPr/tblHeader (повтор на каждой странице) + полужирный центр."""
    return ("<w:tr><w:trPr><w:tblHeader/></w:trPr>"
            + "".join(cell(w, t, bold=True, align="center") for w, t in zip(col_w, titles))
            + "</w:tr>")


def row(col_w: list[int], cells_text: list[str], spans: list[int] | None = None) -> str:
    spans = spans or [1] * len(cells_text)
    return "<w:tr>" + "".join(cell(w, t, span=s)
                              for w, t, s in zip(col_w, cells_text, spans)) + "</w:tr>"


# Колонки «Ход занятия» (§5.8): 7 штук, сумма ширин = текстовая полоса альбомной A4
STAGE_COLS = ["№", "Этапы, продолжительность", "Задачи этапа", "Деятельность педагога",
              "Методы, формы, приемы", "Предполагаемая деятельность детей",
              "Планируемые результаты"]


def stage_widths(total: int) -> list[int]:
    weights = [4, 20, 18, 22, 18, 19, 17]  # широкие — деятельность/детям, узкая — №
    w = [int(total * x / sum(weights)) for x in weights]
    w[-1] += total - sum(w)
    return w


def parse_stages(raw: str, total: int) -> list[str]:
    """Этапы: JSON-список строк [кол7] из параметров; пустой — структура-заготовка.

    Заготовка содержит обязательные этапы 1 / 2 (2.1, 2.2, 2.3) / 3 и
    gridSpan-строки-разделители, содержательные ячейки пустые (заполняет педагог).
    """
    cols = stage_widths(total)
    if raw and raw.strip():
        try:
            data = json.loads(raw)
            if isinstance(data, list) and data:
                out = [header_row(cols, STAGE_COLS)]
                for item in data:
                    if item == "---":       # строка-разделитель через gridSpan
                        out.append(row([total], [""], [len(cols)]))
                    elif isinstance(item, list):
                        cells = [str(c) for c in item][:len(cols)]
                        cells += [""] * (len(cols) - len(cells))
                        out.append(row(cols, cells))
                return out
        except json.JSONDecodeError:
            print("ВНИМАНИЕ: stages не распознаны как JSON — вставлена пустая структура",
                  file=sys.stderr)
    # базовая структура по §5.8 без выдуманного содержания
    skeleton = [
        ["1", "Организационно-мотивационный этап, ___ мин", "", "", "", "", ""],
        "---",
        ["2.1", "Постановка проблемы", "", "", "", "", ""],
        ["2.2", "Ознакомление с материалом", "", "", "", "", ""],
        ["2.3", "Практическое решение", "", "", "", "", ""],
        ["3", "Заключительный этап, ___ мин", "", "", "", "", ""],
    ]
    out = [header_row(cols, STAGE_COLS)]
    for item in skeleton:
        if item == "---":
            # строка-разделитель: одна ячейка на всю ширину (gridSpan)
            out.append(row([total], ["2 — Основной этап"], [len(cols)]))
        else:
            out.append(row(cols, item))
    return out


def build(out_path: str, meta: dict, do_verify: bool = True) -> int:
    g = lambda k, d="": str(meta.get(k, d) or d)  # плоская мета → строка
    body: list[str] = []

    # --- шапка документа (TNR 14, центр) ---
    for key in ("ministry", "fsbe", "univ", "kaf"):
        v = g(key)
        if v:
            body.append(para(v, sz=PT12, align="center"))
    body.append(para("ТЕХНОЛОГИЧЕСКАЯ КАРТА ЗАНЯТИЯ", bold=True, align="center", keep=True))
    if g("theme"):
        body.append(para(f"Тема: {g('theme')}", bold=True, align="center"))
    body.append(para(""))

    # --- блоки §5.8: таблица «блок — значение»; пустые значения = заполняемые поля ---
    total = PAGE_W - 2 * MAR
    info_w = [int(total * 0.30), total - int(total * 0.30)]
    rows: list[str] = []
    block_pairs = [
        ("Образовательная область", g("educational_area") or g("oblast")),
        ("Тема", g("theme")),
        ("Возрастная группа", (g("age_group") + " лет") if g("age_group") else ""),
        ("Длительность", (g("duration_min") + " мин") if g("duration_min") else ""),
        ("Цель", g("goal")),
        ("Задачи", g("tasks")),
        ("Словарная работа", g("vocab_work")),
        ("Планируемый результат", g("planned_result")),
        ("Подготовительная работа", g("prep_work")),
        ("Материалы и оборудование", g("materials") or g("equipment")),
    ]
    for label, value in block_pairs:
        rows.append("<w:tr>" + cell(info_w[0], label, bold=True, align="left")
                    + cell(info_w[1], value, align="left") + "</w:tr>")
    body.append(table(info_w, rows))

    body.append(para("Ход занятия", bold=True, align="center", keep=True))
    stage_cols = stage_widths(total)
    body.append(table(stage_cols, parse_stages(g("stages"), sum(stage_cols))))

    # --- одна секция, альбомная; номер внизу по центру (титул нумеруется) ---
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
    # styles: Normal без basedOn; шрифт/кегль/интервал из констант (rules_hook postprocess
    # донастроит выходной файл, но генерация сразу близка к пресету)
    styles = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
              f'<w:styles {NS}><w:docDefaults><w:rPrDefault>'
              f'<w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
              f'<w:sz w:val="{PT_HEAD}"/><w:szCs w:val="{PT_HEAD}"/></w:rPr></w:rPrDefault>'
              f'<w:pPrDefault><w:pPr><w:spacing w:before="0" w:after="0" w:line="{LINE}" '
              'w:lineRule="auto"/></w:pPr></w:pPrDefault></w:docDefaults>'
              f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/>'
              f'<w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
              f'<w:sz w:val="{PT_HEAD}"/><w:szCs w:val="{PT_HEAD}"/></w:rPr></w:style></w:styles>')
    # колонтитул-поле собираем RAW XML (para() экранировал бы <w:fldChar> как текст)
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
    """Самопроверка: zip читается, XML парсятся, секция/поля/шапки таблиц по пресету."""
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
    checks = [
        (sect is not None and list(body)[-1] is sect, "sectPr — последний дочерний body"),
        (sect is not None and int(sect.find(f"{W}pgSz").get(f"{W}w", "0")) >
         int(sect.find(f"{W}pgSz").get(f"{W}h", "0")), "ориентация альбомная"),
    ]
    mar = sect.find(f"{W}pgMar") if sect is not None else None
    checks.append((mar is not None and all(
        int(mar.get(f"{W}{s}", "-1")) == MAR for s in ("top", "left", "right", "bottom")),
        f"поля {MAR} twips (20 мм) со всех сторон"))
    tbls = body.findall(f".//{W}tbl")
    # шапка с повтором обязательна для таблицы «Ход занятия» (последняя в документе);
    # таблица блоков-лейблов шапки — не данных, повтор шапки ей не нужен
    checks.append((bool(tbls) and
                   tbls[-1].find(f"{W}tr/{W}trPr/{W}tblHeader") is not None,
                   "tblHeader в таблице «Ход занятия»"))
    checks.append((all(tc.find(f"{W}p") is not None for tc in body.iter(f"{W}tc")),
                   "каждая ячейка содержит абзац"))
    for good, msg in checks:
        print(f"VERIFY {'OK  ' if good else 'FAIL'}: {msg}", file=sys.stderr)
        ok = ok and good
    return ok


def load_json_arg(val: str | None) -> dict:
    """JSON inline или путь к файлу (GUI передаёт по-разному в режимах)."""
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
    ap = argparse.ArgumentParser(prog="make_karta.py",
                                 description="Технологическая карта занятия (педагогический пресет)")
    ap.add_argument("--out", default="karta.docx", help="выходной .docx")
    ap.add_argument("--rules", default=None, help="пресет правил JSON (путь/inline)")
    ap.add_argument("--meta", default=None, help="мета/профиль: путь или JSON")
    ap.add_argument("--params", default=None, help="параметры формы: путь или JSON")
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
    meta.update(load_json_arg(args.params))  # параметры формы приоритетнее профиля
    try:
        rc = build(args.out, meta, do_verify=not args.no_verify)
    except Exception as exc:
        print(f"ОШИБКА сборки: {exc!r}", file=sys.stderr)
        return 4
    print(args.out)
    return rc


if __name__ == "__main__":
    sys.exit(main())
