#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""mkdocx.py — конвертер Markdown → DOCX по ГОСТ 7.32 (Мультиворк).

Запуск:
    python3 -I scripts/mkdocx.py input.md output.docx '{"fio":"…","theme":"…"}' \
        [--rules presets/gost_732.json] [--meta profile.json] [--no-verify]

Закрывает все 12 обязательных исправлений §11.1 ТЗ:
 1-2. Одна секция + titlePg: вторая физическая страница показывает «2»,
      никакой ручной смены нумерации не ставится; sectPr — только последний
      дочерний элемент body.
 3.  Заголовки получают w:pStyle Heading1..4 (+keepNext) — работает автосодержание.
 4.  Страница СОДЕРЖАНИЕ + поле TOC + settings.xml updateFields.
 5.  Строки «Таблица N – …» распознаются: слева, без отступа, keepNext.
 6.  Структурные заголовки автоматически переводятся в ПРОПИСНЫЕ.
 7.  Стиль Normal не ссылается сам на себя.
 8.  Реквизиты титула (MINISTRY/FSBE/UNIV/KAF/DISC) берутся из меты/правил,
      год по умолчанию — текущий системный.
 9.  Тип работы параметризован (work_type), титул не захардкожен под отчёт.
 10. Отдельная URL-строка «съедается» только при включённом wokwi_link (по умолчанию выкл).
 11. Ссылки <https://…> — настоящие гиперссылки через Relationships.
 12. Все константы оформления читаются из правил (rules_hook: native).
"""

from __future__ import annotations

import argparse
import datetime
import json
import os
import re
import sys
import zipfile
from xml.sax.saxutils import escape

# ---------------------------------------------------------------------------
# Контракт плагина (§8 ТЗ)
# ---------------------------------------------------------------------------
SCRIPT_META = {
    "name": "Markdown → DOCX (ГОСТ 7.32)",
    "description": "Сборка .docx из Markdown по ГОСТ 7.32: титул из профиля, "
                   "автосодержание, таблицы/рисунки/код/ссылки по правилам пресета.",
    "category": "Учебные работы",
    "version": "2.0",
    "input": "file_md",
    "output_ext": ".docx",
    "arg_template": ["{input}", "{output}", "{meta_json}"],
    "rules_hook": "native",
    "params": [
        {"id": "wokwi_link", "label": "Первая URL-строка — ссылка на Wokwi",
         "type": "bool", "default": False},
        {"id": "title_page", "label": "Титульный лист", "type": "bool", "default": True},
        {"id": "toc", "label": "Страница СОДЕРЖАНИЕ с полем TOC", "type": "bool", "default": True},
    ],
}

# ---------------------------------------------------------------------------
# Единицы OOXML (базовые константы; apply_rules пересчитывает модульные переменные)
# 1 мм = 56.6929 twips; 1 см = 566.929 twips; 1 pt = 2 half-points; 1 см = 360000 EMU
# ---------------------------------------------------------------------------
TWIPS_PER_MM = 56.6929
TWIPS_PER_CM = 566.929
EMU_PER_CM = 360_000
LINE_UNIT = 240  # w:line при lineRule="auto": 1,0 = 240

PT14, PT12, PT10 = 28, 24, 20            # w:sz — полупоинты
LINE15, LINE10 = 360, 240                # интервалы 1,5 и 1,0
IND_FIRST = 709                          # 1,25 см абзацный отступ
PAGE_W, PAGE_H = 11906, 16838            # A4 книжная
MAR_L, MAR_R, MAR_TB = 1701, 850, 1134   # 30 / 15 / 20 / 20 мм
TNR = "Times New Roman"
CONSOLAS = "Consolas"

# Актуальные реквизиты титула (прежнее министерство упразднено в 2018 — не использовать)
DEFAULT_MINISTRY = "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ"
DEFAULT_FSBE = ("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ "
                "ВЫСШЕГО ОБРАЗОВАНИЯ")
HYPERLINK_COLOR = "0563C1"  # синий официальный стиля гиперссылок Word

# Структурные элементы, оформляемые ПРОПИСНЫМИ с новой страницы (§5.2)
STRUCTURAL = {
    "РЕФЕРАТ", "СОДЕРЖАНИЕ", "ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "СПИСОК ИСТОЧНИКОВ", "БИБЛИОГРАФИЧЕСКИЙ СПИСОК", "ПРИЛОЖЕНИЕ",
}

# Правила по умолчанию — срезаем из пресета gost_732 (§7.1 ТЗ, ключевые поля)
DEFAULT_RULES: dict = {
    "page": {"size": "A4", "orientation": "portrait", "margin_left_mm": 30,
             "margin_right_mm": 15, "margin_top_mm": 20, "margin_bottom_mm": 20},
    "text": {"font": TNR, "size_pt": 14, "line_spacing": 1.5,
             "first_line_indent_cm": 1.25, "alignment": "justify"},
    "headings": {"structural_uppercase": True, "keep_with_next": True,
                 "multiline_line_spacing": 1.0},
    "figures": {"max_width_cm": 15, "max_height_cm": 18, "dash": "–",
                "caption_template": "Рисунок {n} – {title}"},
    "tables": {"font_size_pt": 12, "line_spacing": 1.0, "border_size_eighths_pt": 4,
               "caption_template": "Таблица {n} – {title}", "repeat_header": True,
               "header_bold": True},
    "code_blocks": {"font": CONSOLAS, "size_pt": 10, "line_spacing": 1.0, "indent": False},
    "toc": {"required": True},
    "numbering": {"position": "bottom_center", "font_size_pt": 12},
}

# --- модульные переменные оформления: всегда читать ИХ, а не PT14 напрямую ---
PT_MAIN = PT14        # кегль основного текста (полупоинты)
PT_TBL = PT12         # кегль текста в таблицах
PT_FTR = PT12         # кегль номера страницы
CODE_SZ = PT10        # кегль кода
LINE_MAIN = LINE15    # межстрочный основного текста
LINE_CODE = LINE10    # межстрочный кода
LINE_TBL = LINE10     # межстрочный в таблицах
IND = IND_FIRST       # абзацный отступ
ALIGN = "both"        # выравнивание основного текста
FONT = TNR
CODE_FONT = CONSOLAS
FIG_MAX_W_CM = 15.0
FIG_MAX_H_CM = 18.0
TOC_REQUIRED = True
BDR = 4               # толщина границ таблиц (в восьмых пункнта)


def apply_rules(rules: dict) -> None:
    """Пересчитать модульные константы из пресета правил (rules_hook: native).

    Вызывается после чтения --rules поверх дефолтов; неизвестные ключи игнорируются,
    отсутствующие — остаются со значениями по умолчанию ГОСТ.
    """
    global PT_MAIN, PT_TBL, PT_FTR, CODE_SZ, LINE_MAIN, LINE_CODE, LINE_TBL
    global IND, ALIGN, FONT, CODE_FONT, PAGE_W, PAGE_H
    global MAR_L, MAR_R, MAR_TB, FIG_MAX_W_CM, FIG_MAX_H_CM, TOC_REQUIRED, BDR

    page = rules.get("page") or {}
    ml = page.get("margin_left_mm")
    mr = page.get("margin_right_mm")
    mt = page.get("margin_top_mm")
    if ml is not None:
        MAR_L = int(round(ml * TWIPS_PER_MM))
    if mr is not None:
        MAR_R = int(round(mr * TWIPS_PER_MM))
    if mt is not None:
        MAR_TB = int(round(mt * TWIPS_PER_MM))
    if str(page.get("orientation", "")).lower() == "landscape":
        PAGE_W, PAGE_H = PAGE_H, PAGE_W

    text = rules.get("text") or {}
    if text.get("size_pt"):
        PT_MAIN = int(round(float(text["size_pt"]) * 2))
    if text.get("font"):
        FONT = str(text["font"])
    if text.get("line_spacing"):
        LINE_MAIN = int(round(float(text["line_spacing"]) * LINE_UNIT))
    if text.get("first_line_indent_cm") is not None:
        IND = int(round(float(text["first_line_indent_cm"]) * TWIPS_PER_CM))
    jc = {"justify": "both", "left": "left", "center": "center"}.get(
        str(text.get("alignment", "justify")).lower())
    if jc:
        ALIGN = jc

    tbl = rules.get("tables") or {}
    if tbl.get("font_size_pt"):
        PT_TBL = int(round(float(tbl["font_size_pt"]) * 2))
    if tbl.get("line_spacing"):
        LINE_TBL = int(round(float(tbl["line_spacing"]) * LINE_UNIT))
    if tbl.get("border_size_eighths_pt") is not None:
        BDR = int(tbl["border_size_eighths_pt"])

    code = rules.get("code_blocks") or {}
    if code.get("font"):
        CODE_FONT = str(code["font"])
    if code.get("size_pt"):
        CODE_SZ = int(round(float(code["size_pt"]) * 2))
    if code.get("line_spacing"):
        LINE_CODE = int(round(float(code["line_spacing"]) * LINE_UNIT))

    fig = rules.get("figures") or {}
    if fig.get("max_width_cm"):
        FIG_MAX_W_CM = float(fig["max_width_cm"])
    if fig.get("max_height_cm"):
        FIG_MAX_H_CM = float(fig["max_height_cm"])

    if (rules.get("toc") or {}).get("required") is not None:
        TOC_REQUIRED = bool(rules["toc"]["required"])
    ftr = (rules.get("numbering") or {}).get("font_size_pt")
    if ftr:
        PT_FTR = int(round(float(ftr) * 2))


NS = ('xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" '
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" '
      'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" '
      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
      'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"')


def esc(t: object) -> str:
    """Экран XML-текста; кавычки тоже (&quot;) — безопаснее для атрибутов (§7.9)."""
    return escape(str(t), {'"': "&quot;"})


def run(t: str, sz: int | None = None, bold: bool = False, italic: bool = False,
        font: str | None = None, color: str | None = None, underline: bool = False) -> str:
    """Один w:r. Порядок детей rPr строго: rFonts → b → i → color → sz → szCs → u (§5)."""
    sz = sz if sz is not None else PT_MAIN
    font = font or FONT
    parts = [f'<w:r><w:rPr><w:rFonts w:ascii="{esc(font)}" w:hAnsi="{esc(font)}" w:cs="{esc(font)}"/>']
    if bold:
        parts.append("<w:b/>")
    if italic:
        parts.append("<w:i/>")
    if color:
        parts.append(f'<w:color w:val="{color}"/>')
    parts.append(f'<w:sz w:val="{sz}"/><w:szCs w:val="{sz}"/>')
    if underline:
        parts.append('<w:u w:val="single"/>')
    parts.append(f'</w:rPr><w:t xml:space="preserve">{esc(t)}</w:t></w:r>')
    return "".join(parts)


def para(runs_xml: str = "", style: str | None = None, keep: bool = False,
         brk: bool = False, align: str | None = None, ind_first: int | None = None,
         ind_left: int | None = None, line: int | None = None) -> str:
    """Абзац w:p. Порядок детей pPr: pStyle → keepNext → pageBreakBefore → spacing → ind → jc.

    ind_first=None — отступ наследуется из стиля Normal (709); 0 — явный сброс
    (подписи таблиц/рисунков, код, шапка титула).
    """
    line = line if line is not None else LINE_MAIN
    p = []
    if style:
        p.append(f'<w:pStyle w:val="{esc(style)}"/>')
    if keep:
        p.append("<w:keepNext/>")
    if brk:
        p.append("<w:pageBreakBefore/>")
    p.append(f'<w:spacing w:before="0" w:after="0" w:line="{line}" w:lineRule="auto"/>')
    if ind_first is not None or ind_left is not None:
        attrs = ""
        if ind_first is not None:
            attrs += f' w:firstLine="{ind_first}"'
        if ind_left is not None:
            attrs += f' w:left="{ind_left}"'
        if not attrs:
            attrs = ' w:firstLine="0"'
        p.append(f"<w:ind{attrs}/>")
    if align:
        p.append(f'<w:jc w:val="{align}"/>')
    return f"<w:p><w:pPr>{''.join(p)}</w:pPr>{runs_xml}</w:p>"


# --- распознавание конструктов markdown ------------------------------------------------
RE_HEAD = re.compile(r"^(#{1,4})\s+(.*)$")
RE_UL = re.compile(r"^\s*[-*]\s+(.*)$")
RE_OL = re.compile(r"^\s*(\d+)[.)]\s+(.*)$")
RE_IMG = re.compile(r"^!\[(.*?)\]\((.*?)\)\s*$")
RE_URL_LINE = re.compile(r"^<?(https?://\S+)>?\s*$")
RE_TBL_CAP = re.compile(r"^(Таблица\s+\S+\s*[–—-])\s*(.*)$")
RE_FIG_CAP = re.compile(r"^(Рисунок\s+\S+\s*[–—-])\s*(.*)$")
RE_INLINE = re.compile(r"(\*\*.+?\*\*|`[^`]+`|<https?://[^>\s]+>)")


def png_size(data: bytes) -> tuple[int, int] | None:
    """Размер PNG из IHDR (смещение 16, два BE uint32) без внешних библиотек."""
    if data[:8] == b"\x89PNG\r\n\x1a\n" and len(data) >= 24:
        w = int.from_bytes(data[16:20], "big")
        h = int.from_bytes(data[20:24], "big")
        return (w, h) if w and h else None
    return None


def jpeg_size(data: bytes) -> tuple[int, int] | None:
    """Размер JPEG: сканируем сегменты SOF0..SOF3 в поисках H/W (2 байта BE)."""
    if data[:2] != b"\xff\xd8":
        return None
    i = 2
    while i < len(data) - 9:
        if data[i] != 0xFF:
            i += 1
            continue
        marker = data[i + 1]
        if marker in (0xC0, 0xC1, 0xC2, 0xC3):
            h = int.from_bytes(data[i + 5:i + 7], "big")
            w = int.from_bytes(data[i + 7:i + 9], "big")
            return (w, h) if w and h else None
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            i += 2
            continue
        seg_len = int.from_bytes(data[i + 2:i + 4], "big")
        i += 2 + max(seg_len, 1)
    return None


def image_target_cm(px_w: int, px_h: int) -> tuple[float, float]:
    """Физический размер рисунка в см: из пикселей (96 dpi), с ограничением по максимумам правил.

    Крошечные изображения (например 1×1 px) растягиваем до читаемых 6 см,
    пропорции сохраняем всегда.
    """
    w_cm = px_w / 96.0 * 2.54
    h_cm = px_h / 96.0 * 2.54
    if max(w_cm, h_cm) < 1.0:
        scale = 6.0 / max(w_cm, h_cm)
        w_cm *= scale
        h_cm *= scale
    scale = min(FIG_MAX_W_CM / w_cm, FIG_MAX_H_CM / h_cm, 1.0)
    return w_cm * scale, h_cm * scale


class Builder:
    """Состояние одной сборки: счётчики нумерации, relationships, части пакета."""

    def __init__(self, meta: dict, opts: dict) -> None:
        self.meta = meta
        self.opts = opts
        self.body: list[str] = []
        self.media: list[tuple[str, bytes, str]] = []  # (имя части, данные, content-type)
        self.doc_rels: list[tuple[str, str, str, str]] = []  # (id, type, target, mode)
        self.img_sizes: dict[str, tuple[int, int]] = {}
        self.n_fig = 0
        self.n_tbl = 0
        self._toc_added = False
        self._table_start_indices: list[int] = []

    # -- relationships ---------------------------------------------------------------
    def add_hyperlink(self, url: str) -> str:
        rid = f"rIdHL{len(self.doc_rels) + 1}"
        self.doc_rels.append((rid,
                              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
                              url, "External"))
        return rid

    def add_image(self, path: str, base_dir: str) -> str | None:
        """Кладёт файл в word/media и возвращает rId; None, если файла нет."""
        full = path if os.path.isabs(path) else os.path.join(base_dir, path)
        if not os.path.isfile(full):
            print(f"ВНИМАНИЕ: изображение не найдено: {full} — вставлен placeholder",
                  file=sys.stderr)
            return None
        with open(full, "rb") as fh:
            data = fh.read()
        size = png_size(data) or jpeg_size(data)
        ext = os.path.splitext(full)[1].lower().lstrip(".") or "png"
        ctype = {"png": "image/png", "jpg": "image/jpeg", "jpeg": "image/jpeg"}.get(ext)
        if ctype is None:
            print(f"ВНИМАНИЕ: неподдерживаемый формат изображения .{ext} — placeholder",
                  file=sys.stderr)
            return None
        name = f"image{len(self.media) + 1}.{ext}"
        self.media.append((name, data, ctype))
        rid = f"rIdImg{len(self.media)}"
        self.doc_rels.append((rid,
                              "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
                              f"media/{name}", ""))
        # размер в пикселях запоминаем для пересчёта EMU при вставке drawing-блока
        self.img_sizes[rid] = size if size else (960, 720)
        return rid

    # -- низкоуровневые конструкции ---------------------------------------------------

    # -- низкоуровневые конструкции ---------------------------------------------------
    def add(self, xml: str) -> None:
        self.body.append(xml)

    def text_para(self, text: str, *, add: bool = True, bold: bool = False,
                  align: str | None = None, ind_first: int | None = None,
                  line: int | None = None, keep: bool = False, brk: bool = False,
                  sz: int | None = None, style: str | None = None) -> str:
        xml = para(self.inline(text, sz=sz), style=style, keep=keep, brk=brk,
                   align=align if align is not None else ALIGN,
                   ind_first=ind_first, line=line)
        if add:
            self.add(xml)
        return xml

    def inline(self, text: str, sz: int | None = None, bold: bool = False) -> str:
        """Разбор **жирного**, `кода` и <https://…> в runs; ссылки — реальные hyperlink."""
        out: list[str] = []
        for part in RE_INLINE.split(text):
            if not part:
                continue
            if part.startswith("**") and part.endswith("**") and len(part) > 4:
                out.append(run(part[2:-2], sz=sz, bold=True))
            elif part.startswith("`") and part.endswith("`") and len(part) > 2:
                out.append(run(part[1:-1], sz=CODE_SZ, font=CODE_FONT))
            elif part.startswith("<http") and part.endswith(">"):
                url = part[1:-1]
                rid = self.add_hyperlink(url)
                out.append(
                    f'<w:hyperlink r:id="{rid}" w:history="1">'
                    + run(url, sz=sz or PT_MAIN, color=HYPERLINK_COLOR, underline=True)
                    + "</w:hyperlink>")
            else:
                out.append(run(part, sz=sz, bold=bold))
        return "".join(out)

    # -- структурные элементы ----------------------------------------------------------
    def heading(self, level: int, text: str) -> None:
        """Заголовок: pStyle обязательно; структурные — ПРОПИСНЫЕ, новая страница, центр."""
        t = text.strip().rstrip(".").strip()
        up = t.upper()
        struct = up in STRUCTURAL or up.startswith("ПРИЛОЖЕНИЕ")
        style = f"Heading{min(level, 4)}"
        if struct:
            if up.split()[0] == "СОДЕРЖАНИЕ":
                self._toc_added = True
            self.add(para(run(up, bold=True), style=style, keep=True, brk=True,
                          align="center", ind_first=0, line=LINE_MAIN))
        else:
            self.add(para(self.inline(t), style=style, keep=True,
                          align="left", ind_first=IND, line=LINE_MAIN))

    def code_block(self, lines: list[str]) -> None:
        """Код: Consolas, кегль правил, интервал 1,0, БЕЗ отступа — нормализатор не тронет."""
        for ln in lines:
            self.add(para(run(ln if ln else " ", sz=CODE_SZ, font=CODE_FONT),
                          align="left", ind_first=0, ind_left=0, line=LINE_CODE))

    def figure(self, alt: str, src: str, base_dir: str) -> None:
        self.n_fig += 1
        rid = self.add_image(src, base_dir)
        if rid is None:
            self.add(para(run(f"[Изображение {src} не найдено и не вставлено]",
                              italic=True), align="center", ind_first=0))
        else:
            px = self.img_sizes.get(rid, (960, 720))
            w_cm, h_cm = image_target_cm(*px)
            cx, cy = int(round(w_cm * EMU_PER_CM)), int(round(h_cm * EMU_PER_CM))
            drawing = (
                f'<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">'
                f'<wp:extent cx="{cx}" cy="{cy}"/>'
                f'<wp:docPr id="{self.n_fig}" name="Рисунок {self.n_fig}"/>'
                '<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">'
                '<pic:pic><pic:nvPicPr>'
                f'<pic:cNvPr id="{self.n_fig}" name="image{self.n_fig}"/><pic:cNvPicPr/>'
                '</pic:nvPicPr><pic:blipFill>'
                f'<a:blip r:embed="{rid}"/><a:stretch><a:fillRect/></a:stretch>'
                '</pic:blipFill><pic:spPr>'
                f'<a:xfrm><a:off x="0" y="0"/><a:ext cx="{cx}" cy="{cy}"/></a:xfrm>'
                '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
                "</pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>")
            self.add(para(drawing, align="center", ind_first=0, line=LINE10))
        # автоподпись ПОД рисунком: по центру, без отступа, тире из правил
        title = alt.strip().rstrip(".") or f"Рисунок {self.n_fig}"
        self.add(para(run(f"Рисунок {self.n_fig} – {title}"),
                      align="center", ind_first=0, line=LINE10, keep=True))

    def table(self, rows: list[list[str]]) -> None:
        """Таблица: шапка bold + tblHeader, границы sz из правил, колонки = ширина текста."""
        n_cols = max(len(r) for r in rows)
        total = PAGE_W - MAR_L - MAR_R
        col_w = [total // n_cols] * n_cols
        col_w[-1] += total - sum(col_w)
        borders = "".join(
            f'<w:{side} w:val="single" w:sz="{BDR}" w:space="0" w:color="000000"/>'
            for side in ("top", "left", "bottom", "right", "insideH", "insideV"))
        xml = ['<w:tbl><w:tblPr>'
               f'<w:tblW w:w="{total}" w:type="dxa"/>'
               f'<w:tblBorders>{borders}</w:tblBorders>'
               '<w:tblLayout w:type="fixed"/>'
               '<w:tblCellMar><w:left w:w="57" w:type="dxa"/>'
               '<w:right w:w="57" w:type="dxa"/></w:tblCellMar>'
               "</w:tblPr><w:tblGrid>"
               + "".join(f'<w:gridCol w:w="{w}"/>' for w in col_w) + "</w:tblGrid>"]
        for r_i, row in enumerate(rows):
            header = r_i == 0
            xml.append("<w:tr>")
            if header:
                # шапка повторяется на каждой странице (§5.4)
                xml.append("<w:trPr><w:tblHeader/></w:trPr>")
            row = row + [""] * (n_cols - len(row))
            for c_i, cell_text in enumerate(row):
                cell_p = para(self.inline(cell_text, sz=PT_TBL, bold=header),
                              align="center" if header else "left",
                              ind_first=0, line=LINE_TBL)
                xml.append(f'<w:tc><w:tcPr><w:tcW w:w="{col_w[c_i]}" w:type="dxa"/>'
                           '<w:vAlign w:val="center"/></w:tcPr>'
                           # каждая tc обязана содержать минимум один p (§7.10)
                           f"{cell_p}</w:tc>")
            xml.append("</w:tr>")
        xml.append("</w:tbl>")
        # после таблицы всегда абзац — иначе Word может потребовать «исправить» файл
        xml.append(para("", ind_first=0, line=LINE_MAIN))
        self.add("".join(xml))

    # -- титульный лист (§7.2-мета) -------------------------------------------------------
    def title_page(self) -> None:
        m = self.meta
        year = m.get("year") or datetime.date.today().year
        work_type = (m.get("work_type") or "").strip()
        if not (work_type or m.get("fio") or m.get("theme")):
            return  # нет данных — не выдумываем содержательный контент

        def line(text: str, bold: bool = False, sz: int | None = None,
                 align: str = "center", space_before: int = 0) -> str:
            ppr = ['<w:spacing w:before="%d" w:after="0" w:line="%d" w:lineRule="auto"/>'
                   % (space_before * 20, LINE10), '<w:ind w:firstLine="0"/>',
                   f'<w:jc w:val="{align}"/>']
            return f"<w:p><w:pPr>{''.join(ppr)}</w:pPr>{run(text, sz=sz, bold=bold)}</w:p>"

        ministry = m.get("ministry") or DEFAULT_MINISTRY
        fsbe = m.get("fsbe") or DEFAULT_FSBE
        univ = m.get("univ") or ""
        kaf = m.get("kaf") or ""
        for t in (ministry, fsbe):
            self.add(line(t, sz=PT12))
        if univ:
            self.add(line(univ, bold=True, sz=PT12))
        if kaf:
            self.add(line(kaf, sz=PT12))
        self.add(line("", space_before=6))
        self.add(line(work_type.upper(), bold=True, sz=PT_MAIN, space_before=6))
        pz_no = str(m.get("pz_no") or "").strip()
        if pz_no and "практическ" in work_type.lower():
            self.add(line(f"по практическому занятию № {pz_no}", sz=PT_MAIN))
        theme = (m.get("theme") or "").strip()
        if theme:
            self.add(line(f"по теме: «{theme}»", bold=True, sz=PT_MAIN, space_before=3))
        disc = m.get("disc") or ""
        if disc:
            self.add(line(f"Дисциплина: {disc}", sz=PT_MAIN, space_before=3))
        self.add(line("", space_before=10))
        fio = (m.get("fio") or "").strip()
        if fio:
            who = m.get("status") or "студент"
            grp = f" группы {m.get('gruppa')}" if m.get("gruppa") else ""
            self.add(line(f"Выполнил(а): {who}{grp} {fio}",
                          align="right", sz=PT_MAIN))
        prep = (m.get("prep") or "").strip()
        if prep:
            self.add(line(f"Проверил: {prep}", align="right", sz=PT_MAIN))
        self.add(line("", space_before=14))
        city = (m.get("city") or "").strip()
        bottom = f"{city}, {year}" if city else str(year)
        self.add(line(bottom, sz=PT_MAIN))

    # -- содержание (поле TOC, §5 «Содержание обязательно») ------------------------------
    def toc_page(self, at_index: int | None = None) -> None:
        """Страница СОДЕРЖАНИЕ + поле TOC; at_index — вставка внутрь тела (после титула)."""
        start = len(self.body) if at_index is None else at_index
        self.body.insert(start, para(run("СОДЕРЖАНИЕ", bold=True), style="Heading1",
                                     keep=True, brk=True, align="center", ind_first=0))
        # Поле: begin → instrText → separate → кэш → end; Word обновит при открытии
        field = (
            '<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
            '<w:r><w:instrText xml:space="preserve"> TOC \\o "1-3" \\h \\z \\u </w:instrText></w:r>'
            '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
            + run("Содержание будет обновлено автоматически при открытии в Word", sz=PT_MAIN)
            + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        self.body.insert(start + 1,
                         para(field, align="left", ind_first=0, line=LINE_MAIN))


# ---------------------------------------------------------------------------
# Сборка пакета DOCX
# ---------------------------------------------------------------------------
def build(md_path: str, out_path: str, meta: dict | None = None,
          rules: dict | None = None, do_verify: bool = True) -> int:
    """Главная точка сборки. Возвращает код возврата процесса."""
    meta = meta or {}
    opts = {"wokwi_link": bool(meta.get("wokwi_link", False))}
    b = Builder(meta, opts)

    with open(md_path, "r", encoding="utf-8") as fh:
        lines = fh.read().splitlines()
    base_dir = os.path.dirname(os.path.abspath(md_path))

    if meta.get("title_page", True) is not False:
        b.title_page()

    title_end = len(b.body)  # Сюда встанет СОДЕРЖАНИЕ: сразу после титула, до контента

    # --- разбор markdown ---
    code_buf: list[str] | None = None
    tbl_buf: list[list[str]] = []
    pending_toc = TOC_REQUIRED and meta.get("toc", True) is not False
    first_url_taken = False
    wokwi_url: str | None = None

    def flush_table() -> None:
        nonlocal tbl_buf
        if tbl_buf:
            b.table(tbl_buf)
            tbl_buf = []

    for raw in lines:
        if code_buf is not None:
            if raw.strip().startswith("```"):
                b.code_block(code_buf)
                code_buf = None
            else:
                code_buf.append(raw)
            continue
        if raw.strip().startswith("```"):
            flush_table()
            code_buf = []
            continue
        if raw.strip().startswith("|"):
            cells = [c.strip() for c in raw.strip().strip("|").split("|")]
            if all(re.fullmatch(r":?-{2,}:?", c) for c in cells if c):
                continue  # строка-разделитель markdown-таблицы
            tbl_buf.append(cells)
            continue
        flush_table()
        stripped = raw.strip()
        if not stripped:
            continue

        m = RE_HEAD.match(stripped)
        if m:
            b.heading(len(m.group(1)), m.group(2))
            continue
        m = RE_IMG.match(stripped)
        if m:
            b.figure(m.group(1), m.group(2), base_dir)
            continue
        m = RE_TBL_CAP.match(stripped)
        if m:
            # «Таблица N – …»: слева, без отступа, keepNext (привязка к таблице)
            b.add(para(run(stripped.rstrip("."), bold=False), keep=True,
                       align="left", ind_first=0))
            b.n_tbl += 1
            continue
        m = RE_URL_LINE.match(stripped)
        if m and opts["wokwi_link"] and not first_url_taken:
            # только при включённом параметре: первая URL — служебная ссылка на Wokwi
            first_url_taken = True
            wokwi_url = m.group(1)
            continue
        m = RE_UL.match(stripped)
        if m:
            b.add(para(run("– ") + b.inline(m.group(1)), align=ALIGN,
                       ind_first=0, ind_left=IND))
            continue
        m = RE_OL.match(stripped)
        if m:
            b.add(para(run(f"{m.group(1)}. ") + b.inline(m.group(2)),
                       align=ALIGN, ind_first=0, ind_left=IND))
            continue
        b.text_para(stripped)

    flush_table()

    if wokwi_url:
        b.add(para(run("Ссылка на схему Wokwi: ") + b.inline(f"<{wokwi_url}>"),
                   align=ALIGN))

    # СОДЕРЖАНИЕ вставляется сразу после титула, если в md своего нет
    if pending_toc and not b._toc_added:
        b.toc_page(at_index=title_end)

    sect = sect_pr_xml()
    document = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f"<w:document {NS}><w:body>{''.join(b.body)}{sect}</w:body></w:document>")
    write_docx(out_path, document, b)

    if do_verify:
        ok = verify(out_path)
        if not ok:
            print("Самопроверка выявила замечания (см. отчёт выше)", file=sys.stderr)
    return 0


def sect_pr_xml() -> str:
    """Одна секция в конце body. titlePg + пустой footer type="first" → титул без
    номера, вторая страница показывает «2» (нумерация непрерывная, без сброса старта)."""
    return ('<w:sectPr>'
            '<w:footerReference w:type="default" r:id="rIdFtrD"/>'
            '<w:footerReference w:type="first" r:id="rIdFtrF"/>'
            f'<w:pgSz w:w="{PAGE_W}" w:h="{PAGE_H}"/>'
            f'<w:pgMar w:top="{MAR_TB}" w:right="{MAR_R}" w:bottom="{MAR_TB}" '
            f'w:left="{MAR_L}" w:header="709" w:footer="680" w:gutter="0"/>'
            '<w:titlePg/>'
            '<w:docGrid w:linePitch="360"/>'
            "</w:sectPr>")


def styles_xml() -> str:
    """docDefaults + Normal (без basedOn!) + Heading1..4 (pStyle для TOC/навигации)."""
    head_base = ('<w:pPr><w:keepNext/><w:spacing w:before="0" w:after="0" '
                 f'w:line="{LINE_MAIN}" w:lineRule="auto"/><w:ind w:firstLine="{IND}"/>'
                 '<w:jc w:val="center"/>'
                 '<w:outlineLvl w:val="{lvl}"/></w:pPr>')
    heads = "".join(
        f'<w:style w:type="paragraph" w:styleId="Heading{i}">'
        f'<w:name w:val="heading {i}"/><w:basedOn w:val="Normal"/>'
        f'<w:next w:val="Normal"/><w:link w:val="Heading{i}Char"/>'
        + head_base.format(lvl=i - 1)
        + f'<w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
        f"<w:b/><w:sz w:val=\"{PT_MAIN}\"/><w:szCs w:val=\"{PT_MAIN}\"/></w:rPr>"
        "</w:style>"
        for i in range(1, 5))
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<w:styles {NS}>'
            "<w:docDefaults><w:rPrDefault>"
            f'<w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
            f'<w:sz w:val="{PT_MAIN}"/><w:szCs w:val="{PT_MAIN}"/></w:rPr>'
            "</w:rPrDefault><w:pPrDefault><w:pPr>"
            f'<w:spacing w:before="0" w:after="0" w:line="{LINE_MAIN}" w:lineRule="auto"/>'
            f'<w:jc w:val="{ALIGN}"/>'
            "</w:pPr></w:pPrDefault></w:docDefaults>"
            f'<w:style w:type="paragraph" w:default="1" w:styleId="Normal">'
            f'<w:name w:val="Normal"/><w:pPr><w:spacing w:before="0" w:after="0" '
            f'w:line="{LINE_MAIN}" w:lineRule="auto"/><w:ind w:firstLine="{IND}"/>'
            f'<w:jc w:val="{ALIGN}"/></w:pPr>'
            f'<w:rPr><w:rFonts w:ascii="{esc(FONT)}" w:hAnsi="{esc(FONT)}" w:cs="{esc(FONT)}"/>'
            f'<w:sz w:val="{PT_MAIN}"/><w:szCs w:val="{PT_MAIN}"/></w:rPr></w:style>'
            + heads + "</w:styles>")


def footer_xml(first_page: bool) -> str:
    """Колонтитул: default — центр с полем PAGE (низ страницы); first — пустой (титул)."""
    if first_page:
        content = para("", align="center", ind_first=0, line=LINE_MAIN)
    else:
        field = ('<w:r><w:fldChar w:fldCharType="begin"/></w:r>'
                 '<w:r><w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>'
                 '<w:r><w:fldChar w:fldCharType="separate"/></w:r>'
                 + run("2", sz=PT_FTR)
                 + '<w:r><w:fldChar w:fldCharType="end"/></w:r>')
        content = para(field, align="center", ind_first=0, line=LINE_MAIN)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f"<w:ftr {NS}>{content}</w:ftr>")


def write_docx(out_path: str, document: str, b: Builder) -> None:
    """Пишет zip-пакет OOXML: части, rels, content-types ( media включается посписочно)."""
    # Default-записи дедуплицируются по расширению: две одноимённые ломают пакет
    ext_map: dict[str, str] = {}
    for n, _d, c in b.media:
        ext_map[os.path.splitext(n)[1].lstrip(".")] = c
    ct_imgs = "".join(f'<Default Extension="{e}" ContentType="{c}"/>'
                      for e, c in sorted(ext_map.items()))
    content_types = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                     '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                     '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                     '<Default Extension="xml" ContentType="application/xml"/>'
                     + ct_imgs +
                     '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
                     '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>'
                     '<Override PartName="/word/settings.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml"/>'
                     '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                     '<Override PartName="/word/footer2.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
                     "</Types>")
    root_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                 '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                 '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
                 "</Relationships>")
    fixed = [
        ("rIdStyles", "styles", "styles.xml", ""),
        ("rIdSettings", "settings", "settings.xml", ""),
        ("rIdFtrD", "footer", "footer1.xml", ""),
        ("rIdFtrF", "footer", "footer2.xml", ""),
    ]
    base = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
    rel_items = "".join(
        f'<Relationship Id="{rid}" Type="{base}{typ}" Target="{tgt}"'
        + (' TargetMode="External"' if mode else "") + "/>"
        for rid, typ, tgt, mode in fixed)
    # Target экранируется: URL с & иначе ломает XML relationship-файла
    rel_items += "".join(
        f'<Relationship Id="{rid}" Type="{esc(typ)}" Target="{esc(tgt)}"'
        + (' TargetMode="External"' if mode else "") + "/>"
        for rid, typ, tgt, mode in b.doc_rels)
    doc_rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + rel_items + "</Relationships>")
    settings = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
                f"<w:settings {NS}><w:updateFields w:val=\"true\"/></w:settings>")

    out_dir = os.path.dirname(os.path.abspath(out_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    with zipfile.ZipFile(out_path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", content_types)
        z.writestr("_rels/.rels", root_rels)
        z.writestr("word/document.xml", document)
        z.writestr("word/_rels/document.xml.rels", doc_rels)
        z.writestr("word/styles.xml", styles_xml())
        z.writestr("word/settings.xml", settings)
        z.writestr("word/footer1.xml", footer_xml(False))
        z.writestr("word/footer2.xml", footer_xml(True))
        for name, data, _c in b.media:
            z.writestr(f"word/media/{name}", data)


# ---------------------------------------------------------------------------
# Самопроверка (§8 ТЗ)
# ---------------------------------------------------------------------------
def verify(path: str) -> bool:
    """Открывает готовый zip и проверяет ключевые параметры оформления. Отчёт в stderr."""
    from xml.etree import ElementTree as ET
    W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
    results: list[tuple[bool, str]] = []
    try:
        with zipfile.ZipFile(path) as z:
            parts = {n: z.read(n) for n in z.namelist()}
    except Exception as exc:  # битый zip — сразу провал
        print(f"VERIFY FAIL: не читается как zip: {exc}", file=sys.stderr)
        return False

    for name in ("word/document.xml", "word/styles.xml", "word/settings.xml",
                 "word/footer1.xml", "word/footer2.xml", "[Content_Types].xml",
                 "_rels/.rels", "word/_rels/document.xml.rels"):
        if name not in parts:
            results.append((False, f"часть отсутствует: {name}"))
            continue
        try:
            ET.fromstring(parts[name])
            results.append((True, f"XML валиден: {name}"))
        except ET.ParseError as exc:
            results.append((False, f"битый XML {name}: {exc}"))

    doc = ET.fromstring(parts["word/document.xml"])
    body = doc.find(f"{W}body")
    sect = body.find(f"{W}sectPr")
    # sectPr обязан быть последним дочерним элементом body
    results.append((sect is not None and list(body)[-1] is sect,
                    "sectPr — последний дочерний body"))
    if sect is not None:
        mar = sect.find(f"{W}pgMar")
        got = {k: int(mar.get(f"{W}{k}", "-1")) for k in ("top", "right", "bottom", "left")} \
            if mar is not None else {}
        want = {"top": MAR_TB, "right": MAR_R, "bottom": MAR_TB, "left": MAR_L}
        results.append((all(got.get(k, -99) == v for k, v in want.items()),
                        f"поля страницы {got} == {want}"))
        results.append((sect.find(f"{W}titlePg") is not None, "titlePg есть"))
        results.append((sect.find(f"{W}pgNumType") is None,
                        "нет сброса нумерации (вторая страница = 2)"))

    ftr = parts.get("word/footer1.xml", b"").decode("utf-8", "replace")
    results.append(("PAGE" in ftr, "footer1 содержит поле PAGE"))
    results.append(('w:jc w:val="center"' in ftr, "номер внизу по центру"))
    sty = parts.get("word/styles.xml", b"").decode("utf-8", "replace")
    results.append((f'<w:sz w:val="{PT_MAIN}"' in sty, f"кегль Normal {PT_MAIN}hp"))
    results.append((f'w:firstLine="{IND}"' in sty, f"отступ Normal {IND}twips"))

    paras = body.findall(f".//{W}p")
    n = len(paras)
    results.append((f'w:val="Heading' in sty, "стили Heading1..4 объявлены"))
    results.append((any(p.find(f"{W}pPr/{W}pStyle") is not None for p in paras),
                    "заголовки имеют pStyle (автосодержание работает)"))
    tbls = body.findall(f".//{W}tbl")
    if tbls:
        ok_h = all(t.find(f"{W}tr/{W}trPr/{W}tblHeader") is not None for t in tbls)
        results.append((ok_h, f"tblHeader во всех {len(tbls)} таблицах"))
    for p in paras:
        text = "".join(t.text or "" for t in p.findall(f".//{W}t"))
        ppr = p.find(f"{W}pPr")
        if re.match(r"^Таблица\s+\S+\s*[–—-]", text):
            ind_el = ppr.find(f"{W}ind") if ppr is not None else None
            jc = ppr.find(f"{W}jc") if ppr is not None else None
            ok = (ind_el is None or ind_el.get(f"{W}firstLine", "0") == "0") \
                and jc is not None and jc.get(f"{W}val") in ("left", "start")
            results.append((ok, f"подпись «{text[:30]}» слева без отступа"))
        elif re.match(r"^Рисунок\s+\S+\s*[–—-]", text):
            jc = ppr.find(f"{W}jc") if ppr is not None else None
            ok = jc is not None and jc.get(f"{W}val") == "center"
            results.append((ok, f"подпись «{text[:30]}» по центру"))
    results.append((b"updateFields" in parts.get("word/settings.xml", b""),
                    "settings.xml: updateFields (TOC обновится)"))

    fails = [msg for ok, msg in results if not ok]
    for ok, msg in results:
        print(f"VERIFY {'OK  ' if ok else 'FAIL'}: {msg}", file=sys.stderr)
    print(f"VERIFY: абзацев {n}, проверок {len(results)}, провалов {len(fails)}",
          file=sys.stderr)
    return not fails


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------
def flatten_meta(raw: dict) -> dict:
    """Профиль (§7.2) приходит вложенным; скриптам доступен плоским словарём."""
    flat: dict = {}
    for section, val in raw.items():
        if isinstance(val, dict):
            for k, v in val.items():
                if isinstance(v, dict):
                    flat.update({f"{k}.{kk}": vv for kk, vv in v.items()})
                else:
                    flat[k] = v
        else:
            flat[section] = val
    return flat


def load_rules_arg(path: str) -> dict:
    """--rules: имя пресета, путь к файлу или inline JSON; поверх DEFAULT_RULES."""
    rules = json.loads(json.dumps(DEFAULT_RULES))  # глубокая копия
    candidates = [path, os.path.join(os.path.dirname(os.path.abspath(__file__)),
                                     "..", "presets", f"{path}.json"), path + ".json"]
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
    try:
        user = json.loads(path)
        for key, val in user.items():
            if isinstance(val, dict) and isinstance(rules.get(key), dict):
                rules[key].update(val)
            else:
                rules[key] = val
    except (json.JSONDecodeError, TypeError):
        print(f"ВНИМАНИЕ: правила не распознаны ({path}) — использован пресет по умолчанию",
              file=sys.stderr)
    return rules


def main(argv: list[str] | None = None) -> int:
    ap = argparse.ArgumentParser(
        prog="mkdocx.py", description="Markdown → DOCX по ГОСТ 7.32 (только stdlib)")
    ap.add_argument("input", help="входной .md файл")
    ap.add_argument("output", help="выходной .docx")
    ap.add_argument("meta_json", nargs="?", default=None,
                    help="JSON меты inline или путь к файлу (профиль/параметры)")
    ap.add_argument("--rules", default=None, help="пресет правил: путь/имя/JSON")
    ap.add_argument("--meta", default=None, help="файл JSON меты (дополняет позиционный)")
    ap.add_argument("--no-verify", action="store_true", help="отключить самопроверку")
    ap.add_argument("--wokwi-link", action="store_true",
                    help="первая отдельная URL-строка — служебная ссылка Wokwi")
    args = ap.parse_args(argv)

    if not os.path.isfile(args.input):
        print(f"ОШИБКА: входной файл не найден: {args.input}", file=sys.stderr)
        return 3

    rules = load_rules_arg(args.rules) if args.rules else \
        json.loads(json.dumps(DEFAULT_RULES))
    apply_rules(rules)

    meta: dict = {}
    if args.meta_json:
        if os.path.isfile(args.meta_json):
            with open(args.meta_json, "r", encoding="utf-8") as fh:
                meta = flatten_meta(json.load(fh))
        else:
            try:
                meta = flatten_meta(json.loads(args.meta_json))
            except json.JSONDecodeError as exc:
                print(f"ОШИБКА: некорректный JSON меты: {exc}", file=sys.stderr)
                return 2
    if args.meta:
        with open(args.meta, "r", encoding="utf-8") as fh:
            meta = {**flatten_meta(json.load(fh)), **meta}
    if args.wokwi_link:
        meta["wokwi_link"] = True

    try:
        rc = build(args.input, args.output, meta,
                   do_verify=not args.no_verify)
    except Exception as exc:  # показываем причину, но не роняем traceback в stdout
        print(f"ОШИБКА сборки: {exc!r}", file=sys.stderr)
        return 4
    print(args.output)  # путь итогового файла — в stdout
    return rc


if __name__ == "__main__":
    sys.exit(main())
