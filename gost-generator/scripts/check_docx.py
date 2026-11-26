#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""check_docx.py — автономная проверка .docx на соответствие пресету правил (stdlib).

Запуск:
    python3 -I scripts/check_docx.py file.docx [--rules presets/gost_732.json] [--json]

Скрипт-обёртка для терминала (в приложении основной движок — app.core.docx_inspector).
Закрывает 4 дефекта §11.2 ТЗ:
 1. Шрифт проверяется ФАКТИЧЕСКИЙ по w:rFonts (Arial больше не проходит как TNR).
 2. Основной текст считается ВНЕ таблиц и код-блоков (Consolas/кегль кода исключаются).
 3. Отступ проверяется на конкретное значение (±5 twips), а не «на наличие».
 4. Добавлены проверки: нумерация страниц и старт, титул (titlePg), структурные
    заголовки, подписи рисунков/таблиц, гиперссылки, содержание (TOC), свежесть
    и охват источников — всё по правилам из --rules.

Вывод: человекочитаемый текст (exit 0 — ошибок нет, 1 — есть) или JSON (--json).
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import sys
import zipfile
from xml.etree import ElementTree as ET

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TWIPS_PER_MM = 56.6929
TWIPS_PER_CM = 566.929
LINE_UNIT = 240

INDENT_TOLERANCE = 5     # допуск по абзацному отступу, twips
MARGIN_TOLERANCE = 28    # допуск по полям ≈ 0,5 мм

STRUCTURAL = ["РЕФЕРАТ", "СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ",
              "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ИСТОЧНИКОВ",
              "ПРИЛОЖЕНИЕ"]

DEFAULT_RULES: dict = {
    "page": {"orientation": "portrait", "margin_left_mm": 30, "margin_right_mm": 15,
             "margin_top_mm": 20, "margin_bottom_mm": 20},
    "text": {"font": "Times New Roman", "size_pt": 14, "line_spacing": 1.5,
             "first_line_indent_cm": 1.25, "alignment": "justify"},
    "tables": {"font_size_pt": 12, "line_spacing": 1.0, "repeat_header": True,
               "header_bold": True},
    "code_blocks": {"font": "Consolas", "size_pt": 10},
    "headings": {"structural_uppercase": True, "no_trailing_dot": True,
                 "keep_with_next": True},
    "toc": {"required": True},
    "numbering": {"position": "bottom_center", "title_page_numbered": False,
                  "second_page_shows": 2},
    "references": {"max_source_age_years": 5, "min_cited_percent": 80},
    "figures": {"require_reference_in_text": True},
    "tables_ref": {},
}


def load_rules(path: str | None) -> dict:
    """Читаем пресет: явный путь → имя рядом со скриптом (../presets) → дефолт ГОСТ."""
    rules = json.loads(json.dumps(DEFAULT_RULES))
    candidates = []
    if path:
        candidates = [path, path + ".json",
                      os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   "..", "presets", path + ".json")]
    else:
        candidates = [os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                   "..", "presets", "gost_732.json")]
    for cand in candidates:
        if cand and os.path.isfile(cand):
            with open(cand, "r", encoding="utf-8") as fh:
                user = json.load(fh)
            for key, val in user.items():
                if isinstance(val, dict) and isinstance(rules.get(key), dict):
                    rules[key].update(val)
                else:
                    rules[key] = val
            return rules
    print(f"ВНИМАНИЕ: пресет правил не найден ({path}) — использован встроенный ГОСТ-дефолт",
          file=sys.stderr)
    return rules


def pt_to_hp(v: float) -> int:
    return int(round(float(v) * 2))


def cm_to_tw(v: float) -> int:
    return int(round(float(v) * TWIPS_PER_CM))


def mm_to_tw(v: float) -> int:
    return int(round(float(v) * TWIPS_PER_MM))


def p_text(p) -> str:
    """Весь текст абзаца (для распознавания подписей/структурных заголовков)."""
    return "".join(t.text or "" for t in p.iter(f"{W}t"))


def p_props(p) -> dict:
    ppr = p.find(f"{W}pPr")
    out = {"style": None, "jc": None, "line": None, "first": None,
           "keep": False, "brk": False}
    if ppr is None:
        return out
    st = ppr.find(f"{W}pStyle")
    out["style"] = st.get(f"{W}val") if st is not None else None
    jc = ppr.find(f"{W}jc")
    out["jc"] = jc.get(f"{W}val") if jc is not None else None
    sp = ppr.find(f"{W}spacing")
    out["line"] = int(sp.get(f"{W}line")) if sp is not None and sp.get(f"{W}line") else None
    ind = ppr.find(f"{W}ind")
    out["first"] = int(ind.get(f"{W}firstLine")) if ind is not None and ind.get(f"{W}firstLine") else None
    out["left"] = int(ind.get(f"{W}left")) if ind is not None and ind.get(f"{W}left") else None
    out["keep"] = ppr.find(f"{W}keepNext") is not None
    out["brk"] = ppr.find(f"{W}pageBreakBefore") is not None
    return out


def run_fonts_sz(r) -> tuple[str | None, int | None, bool]:
    """Фактические (из прямого форматирования) шрифт/кегль/bold одного w:r."""
    rpr = r.find(f"{W}rPr")
    font = sz = None
    bold = False
    if rpr is not None:
        rf = rpr.find(f"{W}rFonts")
        font = rf.get(f"{W}ascii") if rf is not None else None
        sel = rpr.find(f"{W}sz")
        sz = int(sel.get(f"{W}val")) if sel is not None else None
        b = rpr.find(f"{W}b")
        bold = b is not None and b.get(f"{W}val", "true") not in ("0", "false", "none")
    return font, sz, bold


def effective_normal(styles_root) -> dict:
    """Из стиля Normal (или docDefaults) берём шрифт/кегль/отступ «по наследству»."""
    out = {"font": None, "sz": None, "first": None, "line": None}
    default = None
    for st in styles_root.findall(f"{W}style"):
        if st.get(f"{W}styleId") == "Normal":
            default = st
            break
    src = default if default is not None else styles_root.find(f"{W}docDefaults")
    if src is None:
        return out
    rpr = src.find(f"{W}rPr")
    if rpr is not None:
        rf = rpr.find(f"{W}rFonts")
        if rf is not None:
            out["font"] = rf.get(f"{W}ascii")
        sel = rpr.find(f"{W}sz")
        if sel is not None:
            out["sz"] = int(sel.get(f"{W}val"))
    ppr = src.find(f"{W}pPr")
    if ppr is not None:
        ind = ppr.find(f"{W}ind")
        if ind is not None and ind.get(f"{W}firstLine"):
            out["first"] = int(ind.get(f"{W}firstLine"))
        sp = ppr.find(f"{W}spacing")
        if sp is not None and sp.get(f"{W}line"):
            out["line"] = int(sp.get(f"{W}line"))
    return out


class Checker:
    """Собирает замечания {code, severity, message, location, hint} по правилам."""

    def __init__(self, doc: ET.Element, styles: ET.Element, rules: dict,
                 parts: dict[str, bytes]) -> None:
        self.doc = doc
        self.styles = styles
        self.rules = rules
        self.parts = parts
        self.issues: list[dict] = []
        self.paras: list[tuple[int, ET.Element, str, dict]] = []  # (№, эл, текст, св-ва)
        self._collect()

    def add(self, code: str, severity: str, message: str, where: str = "", hint: str = "") -> None:
        self.issues.append({"code": code, "severity": severity, "message": message,
                            "location": where, "hint": hint})

    def _collect(self) -> None:
        body = self.doc.find(f"{W}body")
        self.body = body
        # Нумеруем ВСЕ абзацы, но помечаем находящиеся внутри таблиц — они не «текст»
        self.in_table_ids = set()
        for tbl in body.iter(f"{W}tbl"):
            for p in tbl.iter(f"{W}p"):
                self.in_table_ids.add(id(p))
        idx = 0
        for p in body.iter(f"{W}p"):
            idx += 1
            self.paras.append((idx, p, p_text(p), p_props(p)))

    # --- дефект 1+2: фактический шрифт/кегль/интервал основного текста ---
    def _main_runs(self):
        """Runs основного текста: вне таблиц; код-блоки (шрифт/кегль правил кода) исключаем."""
        code_font = (self.rules.get("code_blocks") or {}).get("font", "Consolas")
        code_sz = pt_to_hp((self.rules.get("code_blocks") or {}).get("size_pt", 10))
        for num, p, text, pr in self.paras:
            if id(p) in self.in_table_ids:
                continue
            if pr["style"] and pr["style"].lower().startswith("code"):
                continue
            for r in p.findall(f"{W}r"):
                font, sz, bold = run_fonts_sz(r)
                t = "".join(x.text or "" for x in r.findall(f"{W}t"))
                if not t.strip():
                    continue
                if (font or "") == code_font or (sz == code_sz and (font or "") not in
                        (self.rules["text"]["font"], None, "")):
                    continue  # код — отдельная территория, статистику не портит
                yield num, font, sz, bold

    def check_font_size(self) -> None:
        exp_font = self.rules["text"]["font"]
        exp_sz = pt_to_hp(self.rules["text"]["size_pt"])
        # «наследуется» ли шрифт из Normal: берём фактический из rFonts, если задан прямо
        eff = effective_normal(self.styles)
        fonts: dict[str, int] = {}
        sizes: dict[int, int] = {}
        for _num, font, sz, _b in self._main_runs():
            f = font or eff["font"] or ""
            fonts[f] = fonts.get(f, 0) + 1
            if sz:
                sizes[sz] = sizes.get(sz, 0) + 1
        if fonts:
            actual = max(fonts, key=lambda k: fonts[k])
            if actual != exp_font:
                self.add("FONT_MAIN", "error",
                         f"фактический шрифт основного текста «{actual}», ожидается «{exp_font}»",
                         "основной текст", "задать Times New Roman (или правку пресета)")
        else:
            self.add("FONT_MAIN", "info", "не найдено абзацев основного текста вне таблиц")
        if sizes:
            actual_sz = max(sizes, key=lambda k: sizes[k])
            if actual_sz != exp_sz:
                self.add("SIZE_MAIN", "error",
                         f"кегль основного текста {actual_sz / 2} pt, ожидается {exp_sz / 2} pt")

    def check_spacing_indent(self) -> None:
        exp_line = int(round(float(self.rules["text"]["line_spacing"]) * LINE_UNIT))
        exp_ind = cm_to_tw(self.rules["text"]["first_line_indent_cm"])
        eff = effective_normal(self.styles)
        lines: dict[int, int] = {}
        for num, p, text, pr in self.paras:
            if id(p) in self.in_table_ids or not text.strip():
                continue
            if pr["jc"] in ("center", "right"):
                continue  # титул/подписи живут по своим правилам интервала
            if p.find(f".//{W}instrText") is not None:
                continue
            line = pr["line"] or eff["line"]
            if line:
                lines[line] = lines.get(line, 0) + 1
        if lines:
            got = max(lines, key=lambda k: lines[k])
            if abs(got - exp_line) > 10:
                self.add("LINE_SPACING", "error",
                         f"интервал основного текста {got / LINE_UNIT:.2f}, ожидается "
                         f"{self.rules['text']['line_spacing']}")
        # Дефект 3: значение отступа, а не факт наличия.
        # Учитываем только обычные абзацы основного текста: без стиля-заголовка,
        # не выключенные (центр/право — титул/подписи), не код-блоки, не подписи.
        code_font = (self.rules.get("code_blocks") or {}).get("font", "Consolas")
        bad = 0
        good = 0
        for num, p, text, pr in self.paras:
            if id(p) in self.in_table_ids or not text.strip():
                continue
            if pr["style"]:  # заголовки/служебные — по своим правилам
                continue
            if pr["jc"] in ("center", "right"):
                continue
            if p.find(f".//{W}fldChar") is not None or p.find(f".//{W}instrText") is not None:
                continue  # абзацы с полями (TOC/PAGE) — не текст
            fonts = {run_fonts_sz(r)[0] for r in p.findall(f"{W}r")}
            if fonts and fonts <= {code_font, None}:
                continue  # код-блок: отступа нет по правилу code_blocks.indent=false
            if re.match(r"^[–—-]\s", text) and pr["left"]:
                continue  # маркированный список: выступ слева вместо красной строки
            if pr["left"] and abs(pr["left"] - exp_ind) <= INDENT_TOLERANCE and not pr["first"]:
                continue  # список с отступом = правилу first_line_indent
            first = pr["first"] if pr["first"] is not None else eff["first"]
            if first is None:
                continue
            if abs(first - exp_ind) <= INDENT_TOLERANCE:
                good += 1
            elif first == 0 and re.match(r"^(Таблица|Рисунок)\s", text):
                continue  # подписи по ГОСТ без отступа
            else:
                bad += 1
                if bad <= 3:
                    self.add("INDENT_VALUE", "error",
                             f"абзацный отступ {first} twips ({first / TWIPS_PER_CM:.2f} см), "
                             f"ожидается {exp_ind} ±{INDENT_TOLERANCE}",
                             f"абзац {num}: «{text[:40]}»",
                             f"установить firstLine={exp_ind}")
        if bad > 3:
            self.add("INDENT_VALUE", "error", f"…и ещё {bad - 3} абзацев с неверным отступом")

    def check_page(self) -> None:
        sect = self.body.find(f"{W}sectPr")
        if sect is None:
            self.add("PAGE_SIZE", "error", "sectPr отсутствует — формат страницы не задан")
            return
        r = self.rules.get("page") or {}
        mar = sect.find(f"{W}pgMar")
        want = {"left": mm_to_tw(r.get("margin_left_mm", 30)),
                "right": mm_to_tw(r.get("margin_right_mm", 15)),
                "top": mm_to_tw(r.get("margin_top_mm", 20)),
                "bottom": mm_to_tw(r.get("margin_bottom_mm", 20))}
        if mar is None:
            self.add("PAGE_MARGINS", "error", "pgMar отсутствует")
        else:
            for side, val in want.items():
                got = int(mar.get(f"{W}{side}", "0"))
                if abs(got - val) > MARGIN_TOLERANCE:
                    self.add("PAGE_MARGINS", "error",
                             f"поле {side} = {got / TWIPS_PER_MM:.1f} мм, ожидается "
                             f"{val / TWIPS_PER_MM:.0f} мм")
        pg = sect.find(f"{W}pgSz")
        if pg is not None and str(r.get("orientation", "portrait")) == "portrait":
            w_, h_ = int(pg.get(f"{W}w", "0")), int(pg.get(f"{W}h", "0"))
            if w_ >= h_:
                self.add("PAGE_SIZE", "error", "ориентация альбомная, ожидается книжная")

    def check_numbering(self) -> None:
        """Нумерация: титул без номера, вторая страница = «2», номер внизу по центру."""
        n = self.rules.get("numbering") or {}
        sect = self.body.find(f"{W}sectPr")
        if sect is None:
            return
        pgnum = sect.find(f"{W}pgNumType")
        if pgnum is not None and pgnum.get(f"{W}start") is not None:
            self.add("NUM_START", "error",
                     "задан сброс нумерации (pgNumType start) — страницы могут «перенумероваться»",
                     hint="одна секция + titlePg без сброса старта")
        if n.get("title_page_numbered", False) is False:
            if sect.find(f"{W}titlePg") is None:
                self.add("NUM_START", "error", "титульный не изолирован: отсутствует titlePg")
            else:
                self.add("NUM_START", "info", "titlePg есть; при пустом first-колонтитуле "
                                              "вторая страница показывает «2»")
        ftr = self.parts.get("word/footer1.xml")
        if ftr is None:
            self.add("NUM_POSITION", "error", "колонтитул страниц не найден")
            return
        root = ET.fromstring(ftr)
        has_page = any("PAGE" in (el.text or "") for el in root.iter(f"{W}instrText"))
        center = any(jc.get(f"{W}val") == "center" for jc in root.iter(f"{W}jc"))
        if not has_page:
            self.add("NUM_POSITION", "error", "в нижнем колонтитуле нет поля PAGE")
        if not center:
            self.add("NUM_POSITION", "error", "номер страницы не по центру")

    def check_toc(self) -> None:
        if not (self.rules.get("toc") or {}).get("required", True):
            return
        has_toc = any("TOC" in (el.text or "") for el in self.doc.iter(f"{W}instrText"))
        if not has_toc:
            self.add("TOC_PRESENT", "error", "нет поля автосодержания TOC",
                     hint="страница СОДЕРЖАНИЕ + поле TOC + updateFields")

    def check_structural_headings(self) -> None:
        r = self.rules.get("headings") or {}
        for num, p, text, pr in self.paras:
            t = text.strip()
            up = t.upper()
            is_struct = any(up == s or up.startswith(s) for s in STRUCTURAL)
            if not (is_struct and pr["style"]):
                continue
            if r.get("structural_uppercase", True) and t != up:
                self.add("HEAD_STRUCT", "error", f"структурный заголовок не ПРОПИСНЫМИ: «{t}»",
                         f"абзац {num}")
            if r.get("no_trailing_dot", True) and t.endswith("."):
                self.add("HEAD_STRUCT", "error", f"точка в конце заголовка: «{t}»", f"абзац {num}")
            if pr["jc"] not in ("center",):
                self.add("HEAD_STRUCT", "warning", f"структурный заголовок не по центру: «{t}»",
                         f"абзац {num}")
            if not pr["keep"]:
                self.add("HEAD_KEEP", "warning", f"у заголовка нет keepNext: «{t}»", f"абзац {num}")

    def check_captions(self) -> None:
        for num, p, text, pr in self.paras:
            t = text.strip()
            m = re.match(r"^Таблица\s+(\S+)\s*(.*)$", t)
            if m:
                if "–" not in t and "-" not in t:
                    self.add("TBL_CAPTION", "error", f"название таблицы без тире: «{t[:50]}»",
                             f"абзац {num}")
                if pr["jc"] not in (None, "left", "start"):
                    self.add("TBL_CAPTION", "error",
                             f"название таблицы выроввано «{pr['jc']}», ожидается слева",
                             f"абзац {num}: «{t[:40]}»", "jc=left и firstLine=0")
                if pr["first"]:
                    self.add("TBL_CAPTION", "error", "у названия таблицы есть абзацный отступ",
                             f"абзац {num}", "firstLine=0")
                if not pr["keep"]:
                    self.add("TBL_CAPTION", "warning", "название таблицы не привязано keepNext",
                             f"абзац {num}")
                continue
            m = re.match(r"^Рисунок\s+(\S+)\s*(.*)$", t)
            if m:
                if pr["jc"] != "center":
                    self.add("FIG_CAPTION", "error",
                             f"подпись рисунка не по центру («{pr['jc']}»): «{t[:40]}»",
                             f"абзац {num}")
                if pr["first"]:
                    self.add("FIG_CAPTION", "error", "у подписи рисунка есть абзацный отступ",
                             f"абзац {num}")
                if t.endswith("."):
                    self.add("FIG_CAPTION", "warning", f"точка в конце подписи: «{t[:50]}»")

    def check_tables(self) -> None:
        r = self.rules.get("tables") or {}
        for i, tbl in enumerate(self.body.findall(f"{W}tbl"), 1):
            trs = tbl.findall(f"{W}tr")
            if not trs:
                continue
            # требуем повтор шапки только у таблиц с «шапочной» первой строкой
            # (полностью полужирной) — лейбловые таблицы анкет не касаются
            bolds = [run_fonts_sz(ru)[2] for tc in trs[0].findall(f"{W}tc")
                     for p in tc.iter(f"{W}p") for ru in p.findall(f"{W}r")
                     if p_text(p).strip()]
            # «шапочная» строка — минимум две полужирные ячейки (лейбл+значение в
            # анкетных таблицах не выглядит как шапка данных)
            looks_header = len(bolds) >= 2 and all(bolds)
            trpr = trs[0].find(f"{W}trPr")
            if (looks_header and r.get("repeat_header", True)
                    and (trpr is None or trpr.find(f"{W}tblHeader") is None)):
                self.add("TBL_HEADER", "warning", f"в таблице {i} шапка не повторяется (нет tblHeader)",
                         f"таблица {i}")
            for tc in tbl.iter(f"{W}tc"):
                if tc.find(f"{W}p") is None:
                    self.add("TBL_STRUCT", "error", f"в таблице {i} есть ячейка без абзаца",
                             "Word предложит «исправить» файл")
                    break

    def check_links(self) -> None:
        rels = self.parts.get("word/_rels/document.xml.rels", b"").decode("utf-8", "replace")
        n_hl = len(re.findall(r'Type="[^"]*hyperlink"', rels))
        n_blue = len(re.findall(r'w:color w:val="(?:0563C1|0000FF|1155CC)"',
                                self.parts.get("word/document.xml", b"").decode("utf-8", "replace")))
        if n_blue > n_hl:
            self.add("LINKS", "error",
                     f"синего текста ({n_blue}) больше, чем реальных гиперссылок ({n_hl}) — "
                     "ссылки не кликабельны",
                     hint="использовать Relationship Type hyperlink + TargetMode External")

    def check_references(self) -> None:
        rr = self.rules.get("references") or {}
        max_age = int(rr.get("max_source_age_years", 5) or 0)
        min_cite = int(rr.get("min_cited_percent", 80) or 0)
        # Собираем секцию списка источников
        in_bib = False
        entries: list[tuple[int, str]] = []
        body_all = " ".join(t for _n, _p, t, _pr in self.paras)
        for num, p, text, pr in self.paras:
            up = text.strip().upper()
            if any(up.startswith(s) for s in ("СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
                                              "СПИСОК ИСТОЧНИКОВ", "БИБЛИОГРАФИЧЕСКИЙ СПИСОК")):
                in_bib = True
                continue
            if in_bib:
                m = re.match(r"^\s*(\d+)[.)]\s+(.*)$", text)
                if m:
                    entries.append((int(m.group(1)), m.group(2)))
                elif up and not pr["style"]:
                    in_bib = False  # следующий структурный блок — выходим
        if not entries:
            return
        if max_age > 0:
            year_now = datetime.date.today().year
            for num_e, entry in entries:
                years = [int(y) for y in re.findall(r"\b(?:19|20)\d{2}\b", entry)]
                fresh = [y for y in years if year_now - y <= max_age]
                if years and not fresh:
                    self.add("BIB_AGE", "warning",
                             f"источник {num_e} старше {max_age} лет: «{entry[:60]}»",
                             "список литературы", "заменить на более свежее издание")
        if min_cite > 0:
            cited = {int(n) for n in re.findall(r"\[(\d+)(?:,|\]|$)", body_all)
                     if int(n) <= len(entries) * 2}
            if entries:
                nums = [n for n, _ in entries]
                covered = sum(1 for n in nums if n in cited)
                pct = int(round(100 * covered / len(nums)))
                if pct < min_cite:
                    self.add("BIB_COVERAGE", "warning",
                             f"процитировано {pct}% источников, минимум {min_cite}%")

    def run(self) -> list[dict]:
        for fn in (self.check_page, self.check_font_size, self.check_spacing_indent,
                   self.check_numbering, self.check_toc, self.check_structural_headings,
                   self.check_captions, self.check_tables, self.check_links,
                   self.check_references):
            fn()
        if str((self.rules.get("page") or {}).get("orientation")) == "landscape":
            # педагогический пресет: не путаем — отключаем неуместные ГОСТ-предупреждения
            self.issues = [i for i in self.issues if i["code"] not in ("INDENT_VALUE",)]
        return self.issues


def check_file(path: str, rules: dict) -> tuple[list[dict], str | None]:
    """Возвращает (замечания, ошибка_чтения). Битый zip/XML не роняет скрипт."""
    try:
        with zipfile.ZipFile(path) as z:
            parts = {n: z.read(n) for n in z.namelist()}
    except (zipfile.BadZipFile, OSError) as exc:
        return [], f"файл не читается как DOCX/zip: {exc}"
    try:
        doc = ET.fromstring(parts["word/document.xml"])
        styles = ET.fromstring(parts.get("word/styles.xml", b"<w:styles xmlns:w='http://schemas.openxmlformats.org/wordprocessingml/2006/main'/>"))
    except (ET.ParseError, KeyError) as exc:
        return [], f"битый XML основной части: {exc}"
    return Checker(doc, styles, rules, parts).run(), None


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(prog="check_docx.py",
                                 description="Проверка .docx по пресету правил (stdlib)")
    ap.add_argument("file", help="проверяемый .docx")
    ap.add_argument("--rules", default=None, help="пресет: путь/имя (по умолчанию gost_732)")
    ap.add_argument("--json", action="store_true", help="машинный вывод (конверт §19.4)")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.file):
        msg = f"файл не найден: {args.file}"
        if args.json:
            print(json.dumps({"schema_version": 1, "ok": False, "command": "check",
                              "data": {}, "errors": [{"code": "FILE_NOT_FOUND",
                                                      "message": msg, "path": args.file}],
                              "warnings": []}, ensure_ascii=False))
        else:
            print(f"ОШИБКА: {msg}", file=sys.stderr)
        return 3

    rules = load_rules(args.rules)
    issues, fatal = check_file(args.file, rules)
    if fatal:
        if args.json:
            print(json.dumps({"schema_version": 1, "ok": False, "command": "check",
                              "data": {"file": args.file},
                              "errors": [{"code": "READ_ERROR", "message": fatal}],
                              "warnings": []}, ensure_ascii=False))
        else:
            print(f"ОШИБКА ЧТЕНИЯ: {fatal}", file=sys.stderr)
        return 4

    counts = {"error": 0, "warning": 0, "info": 0}
    for it in issues:
        counts[it["severity"]] = counts.get(it["severity"], 0) + 1

    if args.json:
        print(json.dumps({"schema_version": 1, "ok": counts["error"] == 0,
                          "command": "check",
                          "data": {"file": args.file,
                                   "preset": (rules.get("name") or "gost_732"),
                                   "summary": counts, "issues": issues},
                          "errors": [i for i in issues if i["severity"] == "error"],
                          "warnings": [i for i in issues if i["severity"] == "warning"]},
                         ensure_ascii=False, indent=2))
    else:
        print(f"Проверка: {args.file} (пресет {(rules.get('name') or 'gost_732')})")
        for it in issues:
            loc = f" [{it['location']}]" if it.get("location") else ""
            hint = f" → {it['hint']}" if it.get("hint") else ""
            print(f"  {it['severity'].upper():7} {it['code']:14} {it['message']}{loc}{hint}")
        print(f"Итого: ошибок {counts['error']}, предупреждений {counts['warning']}, "
              f"примечаний {counts['info']}")
    return 1 if counts["error"] else 0


if __name__ == "__main__":
    sys.exit(main())
