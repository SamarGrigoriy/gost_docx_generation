"""Низкоуровневые OOXML-утилиты: пространства имён, чтение zip-пакета, сериализация.

Модуль — единственное место, где известны тонкости схемы WordprocessingML:

* порядок дочерних элементов в `w:pPr`, `w:rPr`, `w:sectPr`, `w:tblPr`, `w:trPr`,
  `w:tcPr`, `w:settings`. Нарушение порядка = Word предлагает «восстановить файл»
  (п. 7 CHATBOT_TEMPLATE, п. 11.2 ТЗ).
* сериализация `xml.etree` выбрасывает необъявленные префиксы, поэтому список
  `xmlns:*` из исходной части восстанавливается вручную — иначе `mc:Ignorable="w14"`
  начинает ссылаться на несуществующий префикс и пакет становится битым.

Используется и инспектором, и нормализатором, чтобы логика схемы не дублировалась.
"""

from __future__ import annotations

import re
import zipfile
import xml.etree.ElementTree as ET
from typing import Iterable, Iterator, Mapping, MutableMapping, Sequence

__all__ = [
    "DocxError",
    "NS",
    "qn",
    "local_name",
    "read_package",
    "write_package",
    "serialize_part",
    "ensure_child",
    "remove_children",
    "get_int",
    "set_attrs",
    "find_all",
    "find_child",
    "iter_children",
    "text_of",
]

# --- Пространства имён -------------------------------------------------------
W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
M = "http://schemas.openxmlformats.org/officeDocument/2006/math"
V = "urn:schemas-microsoft-com:vml"
O = "urn:schemas-microsoft-com:office:office"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
DGR = "http://schemas.openxmlformats.org/drawingml/2006"

NS: Mapping[str, str] = {
    "w": W,
    "r": R,
    "rel": REL,
    "ct": CT,
    "wp": WP,
    "a": A,
    "pic": PIC,
    "m": M,
    "v": V,
    "o": O,
    "mc": MC,
    "w14": W14,
    "wp14": WP14,
}

# Регистрируем ПРЕЖДЕ чем что-либо парсить: иначе ET перепишет префиксы в ns0/ns1
# и Round-trip через normalize_docx породит лишние отличия (идемпотентность ломается).
for _prefix, _uri in NS.items():
    ET.register_namespace(_prefix, _uri)
# Типовые «мусорные» пространства в реальных документах Word — регистрируем, чтобы
# сохранить их префиксы как есть.
for _prefix, _uri in (
    ("w15", "http://schemas.microsoft.com/office/word/2012/wordml"),
    ("wne", "http://schemas.microsoft.com/office/word/2006/wordml"),
    ("sl", "http://schemas.openxmlformats.org/schemaLibrary/2006/main"),
    ("mcs", "http://schemas.microsoft.com/office/word/2016/wordml/csl"),
):
    ET.register_namespace(_prefix, _uri)


class DocxError(Exception):
    """Битый zip / битый XML / отсутствующая обязательная часть пакета DOCX."""


def qn(name: str) -> str:
    """`"w:p"` → `{uri}p`; уже развернутый `{...}p` и НЕпрефиксные имена (`"cx"`) как есть.

    Атрибуты OOXML бываю без пространства имён (`cx`, `w`, `Id` в relationships) —
    для них qn работает тождественно.
    """
    if name.startswith("{"):
        return name
    if ":" not in name:
        return name
    prefix, _, tag = name.partition(":")
    uri = NS.get(prefix)
    if uri is None:
        raise DocxError(f"Неизвестное пространство имён: {prefix}")
    return f"{{{uri}}}{tag}"


def local_name(tag: str) -> str:
    """Локальное имя тега без пространства имён."""
    return tag.rsplit("}", 1)[-1] if "}" in tag else tag


# --- Порядок детей по схеме --------------------------------------------------
# Полные последовательности CT_PPr / CT_RPr / CT_SectPr / CT_TblPr / CT_TrPr /
# CT_TcPr / CT_Settings. Нужны, чтобы вставка элемента попала в законную позицию.
PPR_ORDER: Sequence[str] = (
    "pStyle", "keepNext", "keepLines", "pageBreakBefore", "framePr", "widowControl",
    "numPr", "suppressLineNumbers", "pBdr", "shd", "tabs", "suppressAutoHyphens",
    "kinsoku", "wordWrap", "overflowPunct", "topLinePunct", "autoSpaceDE",
    "autoSpaceDN", "bidi", "adjustRightInd", "snapToGrid", "spacing", "ind",
    "contextualSpacing", "mirrorIndents", "suppressOverlap", "jc", "textDirection",
    "textAlignment", "textboxTightWrap", "outlineLvl", "divId", "cnfStyle", "rPr",
    "sectPr", "pPrChange",
)
RPR_ORDER: Sequence[str] = (
    "rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike",
    "dstrike", "outline", "shadow", "emboss", "webHidden", "color", "spacing", "w",
    "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd",
    "fitText", "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout",
    "specVanish", "oMath",
)
SECTPR_ORDER: Sequence[str] = (
    "headerReference", "footerReference", "footnotePr", "endnotePr", "type", "pgSz",
    "pgMar", "paperSrc", "pgBorders", "lnNumType", "pgNumType", "cols", "formProt",
    "vAlign", "noEndnote", "titlePg", "textDirection", "bidi", "rtlGutter",
    "docGrid", "printerSettings", "sectPrChange",
)
TBLPR_ORDER: Sequence[str] = (
    "tblStyle", "tblpPr", "tblOverlap", "bidiVisual", "tblStyleRowBandSize",
    "tblStyleColBandSize", "tblW", "jc", "tblCellSpacing", "tblInd", "tblBorders",
    "shd", "tblLayout", "tblCellMar", "tblLook", "tblCaption", "tblDescription",
    "tblPrChange",
)
TRPR_ORDER: Sequence[str] = (
    "cnfStyle", "divId", "gridBefore", "gridAfter", "wBefore", "wAfter", "cantSplit",
    "trHeight", "tblHeader", "tblCellSpacing", "jc", "hidden", "ins", "del",
    "trPrChange",
)
TCPR_ORDER: Sequence[str] = (
    "cnfStyle", "tcW", "gridSpan", "hMerge", "vMerge", "tcBorders", "shd", "noWrap",
    "tcMar", "textDirection", "tcFitText", "vAlign", "hideMark", "headers",
    "tcPrChange",
)
SETTINGS_ORDER: Sequence[str] = (
    "writeProtection", "view", "zoom", "removePersonalInformation",
    "doNotDisplayPageBoundaries", "displayBackgroundShape", "printPostScriptOverText",
    "printFractionalCharacterWidth", "printFormsData", "embedTrueTypeFonts",
    "embedSystemFonts", "saveSubsetFonts", "saveFormsData", "mirrorMargins",
    "alignBordersAndEdges", "bordersDoNotSurroundHeader",
    "bordersDoNotSurroundFooter", "gutterAtTop", "hideSpellingErrors",
    "hideGrammaticalErrors", "activeWritingStyle", "proofState", "formsDesign",
    "attachedTemplate", "linkStyles", "stylePaneFormatFilter", "stylePaneSortMethod",
    "documentType", "mailMerge", "revisionView", "trackChanges", "doNotTrackMoves",
    "doNotTrackFormatting", "documentProtection", "autoFormatOverride",
    "styleLockTheme", "styleLockQFSet", "defaultTabStop", "autoHyphenation",
    "consecutiveHyphenLimit", "hyphenationZone", "doNotHyphenateCaps", "showEnvelope",
    "summaryLength", "clickAndTypeStyle", "defaultTableStyle", "evenAndOddHeaders",
    "bookFoldRevPrinting", "bookFoldPrinting", "bookFoldPrintingSheets",
    "drawingGridHorizontalSpacing", "drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery", "displayVerticalDrawingGridEvery",
    "doNotUseMarginsForDrawingGridOrigin", "drawingGridHorizontalOrigin",
    "drawingGridVerticalOrigin", "doNotShadeFormData", "noPunctuationKerning",
    "characterSpacingControl", "printTwoOnOne", "strictFirstAndLastChars",
    "noLineBreaksAfter", "noLineBreaksBefore", "savePreviewPicture",
    "doNotValidateAgainstSchema", "saveInvalidXml", "ignoreMixedContent",
    "alwaysShowPlaceholderText", "doNotDemarcateInvalidXml", "saveXmlDataOnly",
    "useXSLTWhenSaving", "saveThroughXslt", "showXMLTags", "alwaysMergeEmptyNamespace",
    "updateFields", "hdrShapeDefaults", "footnotePr", "endnotePr", "compat", "docVars",
    "rsids", "mathPr", "attachedSchema", "themeFontLang", "clrSchemeMapping",
    "doNotIncludeSubdocsInStats", "doNotAutoCompressPictures", "forceUpgrade",
    "captions", "readModeInkLockDown", "smartTagType", "shapeDefaults",
    "doNotEmbedSmartTags", "decimalSymbol", "listSeparator",
)

_ORDER_BY_PARENT: Mapping[str, Sequence[str]] = {
    "pPr": PPR_ORDER,
    "rPr": RPR_ORDER,
    "sectPr": SECTPR_ORDER,
    "tblPr": TBLPR_ORDER,
    "trPr": TRPR_ORDER,
    "tcPr": TCPR_ORDER,
    "settings": SETTINGS_ORDER,
}

# Позиция, куда ставим элемент, если порядок родителя неизвестен.
_APPEND = object()


def order_of(parent: ET.Element) -> Sequence[str] | object:
    """Порядок детей для родителя по его локальному имени (или None)."""
    return _ORDER_BY_PARENT.get(local_name(parent.tag), _APPEND)


def set_attrs(el: ET.Element, attrs: Mapping[str, str | int | bool | None]) -> ET.Element:
    """Массовая установка атрибутов; `None` снимает атрибут. Ключи — через `qn()`."""
    for key, value in attrs.items():
        name = key if key.startswith("{") or ":" not in key else qn(key)
        if value is None:
            el.attrib.pop(name, None)
        else:
            el.set(name, str(value))
    return el


def get_int(el: ET.Element | None, attr: str, default: int | None = None) -> int | None:
    """Целочисленный атрибут (OOXML хранит числа строками) с дефолтом."""
    if el is None:
        return default
    raw = el.get(qn(attr)) if not attr.startswith("{") else el.get(attr)
    if raw is None:
        return default
    try:
        return int(float(raw))
    except (TypeError, ValueError):
        return default


def find_child(parent: ET.Element | None, path: str) -> ET.Element | None:
    """Первый потомок по пути из `w:`-имён: `find_child(p, "w:pPr/w:ind")`."""
    if parent is None:
        return None
    return parent.find(_expand_path(path))


def find_all(parent: ET.Element, path: str) -> list[ET.Element]:
    """Все потомки по пути из `w:`-имён."""
    return list(parent.iterfind(_expand_path(path)))


def iter_children(parent: ET.Element | None, local: str) -> Iterator[ET.Element]:
    """Непосредственные дети с указанным локальным именем (без учёта NS)."""
    if parent is None:
        return iter(())
    return (ch for ch in list(parent) if local_name(ch.tag) == local)


def _expand_path(path: str) -> str:
    """`"w:pPr/w:ind"` → `"{uri}pPr/{uri}ind"` для ElementTree."""
    return "/".join(qn(seg) for seg in path.split("/"))


def ensure_child(
    parent: ET.Element,
    tag: str,
    *,
    index: int = 0,
    attrs: Mapping[str, str | int | bool | None] | None = None,
    replace: bool = True,
) -> ET.Element:
    """Создать/найти дочерний элемент в законной позиции схемы.

    `replace=True` убирает одноимённые элементы: в OOXML большинство детей — maxOccurs=1,
    два `w:spacing` Word считает ошибкой схемы.
    """
    name = qn(tag)
    if replace:
        for old in [ch for ch in list(parent) if ch.tag == name]:
            parent.remove(old)
    child = ET.Element(name)
    parent.append(child)
    if attrs:
        set_attrs(child, attrs)
    _reorder(parent, child, index)
    return child


def _reorder(parent: ET.Element, child: ET.Element, index: int) -> None:
    """Переставляет `child` в соответствие со схемой относительно братьев."""
    order = order_of(parent)
    parent.remove(child)
    if order is _APPEND:
        parent.insert(min(index, len(parent)), child)
        return
    name = local_name(child.tag)
    if name not in order:
        parent.append(child)
        return
    my_rank = order.index(name)
    insert_at = len(parent)
    for pos, other in enumerate(list(parent)):
        other_rank = order.index(local_name(other.tag)) if local_name(other.tag) in order else -1
        if other is child:
            continue
        if other_rank > my_rank:
            insert_at = pos
            break
        if other_rank == my_rank and order is not SETTINGS_ORDER:
            # одинаковые элементы (headerReference/footerReference) — держим группу вместе
            continue
    parent.insert(insert_at, child)


def remove_children(parent: ET.Element, tag: str) -> int:
    """Удалить всех детей с указанным тегом; вернуть число удалений."""
    name = qn(tag)
    removed = 0
    for ch in [c for c in list(parent) if c.tag == name]:
        parent.remove(ch)
        removed += 1
    return removed


def has_child(parent: ET.Element | None, tag: str) -> bool:
    """Есть ли у родителя дочерний элемент с таким тегом."""
    if parent is None:
        return False
    name = qn(tag)
    return any(ch.tag == name for ch in list(parent))


def text_of(el: ET.Element | None) -> str:
    """Текстовое содержимое элемента (для `w:instrText`, `w:t`)."""
    return (el.text or "") if el is not None else ""


# --- Чтение/запись пакета ----------------------------------------------------
_XML_DECL_RE = re.compile(rb"^<\?xml[^>]*\?>\s*", re.IGNORECASE)
_XMLNS_RE = re.compile(r'xmlns:([A-Za-z0-9_]+)="([^"]+)"')

# Минимально необходимый набор частей; отсутствие document.xml = не DOCX.
REQUIRED_PARTS = ("[Content_Types].xml", "word/document.xml")


def read_package(path: str) -> tuple[dict[str, bytes], list[str]]:
    """Прочитать весь zip в память.

    Возвращает (части, порядок имён): порядок нужен, чтобы нормализатор писал
    пакет в той же последовательности — иначе diff двух прогонов выглядит «изменением».
    """
    try:
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            parts = {name: zf.read(name) for name in names}
    except zipfile.BadZipFile as exc:  # не zip вообще
        raise DocxError(f"Файл не является DOCX (битый zip): {path}") from exc
    except OSError as exc:
        raise DocxError(f"Не удалось прочитать файл: {path} ({exc})") from exc
    missing = [name for name in REQUIRED_PARTS if name not in parts]
    if missing:
        raise DocxError(f"В пакете нет обязательных частей: {', '.join(missing)}")
    return parts, names


def parse_part(parts: Mapping[str, bytes], name: str) -> ET.Element | None:
    """Распарсить XML-часть; отсутствующая часть = None, битый XML = DocxError."""
    raw = parts.get(name)
    if raw is None:
        return None
    try:
        return ET.fromstring(_XML_DECL_RE.sub(b"", raw, count=1))
    except ET.ParseError as exc:
        raise DocxError(f"Битый XML в части {name}: {exc}") from exc


def serialize_part(root: ET.Element, original: bytes | None) -> bytes:
    """Сериализовать дерево обратно в XML, сохранив объявления пространств имён.

    ET выбрасывает неиспользуемые `xmlns:`, а `mc:Ignorable="w14 wp14"` обязан ссылаться
    только на объявленные префиксы — поэтому дообъявляем их из исходной части.
    """
    body = ET.tostring(root, encoding="unicode")
    if original:
        head = original[:4096].decode("utf-8", "replace")
        declared = set(_XMLNS_RE.findall(body[:4096]))
        extra = [(p, u) for p, u in _XMLNS_RE.findall(head) if (p, u) not in declared]
        if extra:
            inject = " ".join(f'xmlns:{p}="{u}"' for p, u in extra)
            gt = body.index(">")
            body = f"{body[:gt]} {inject}{body[gt:]}"
    decl = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
    return (decl + body).encode("utf-8")


def write_package(
    path: str,
    parts: MutableMapping[str, bytes],
    order: Iterable[str],
    extra_names: Iterable[str] = (),
) -> None:
    """Пересобрать zip. Новые части дописываются в конец, порядок остальных сохраняется."""
    seen: list[str] = []
    for name in list(order) + list(extra_names):
        if name in parts and name not in seen:
            seen.append(name)
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as zf:
        # Фиксированная дата: детерминированный вывод = выполнимая идемпотентность.
        info_stamp = (1980, 1, 1, 0, 0, 0)
        for name in seen:
            zi = zipfile.ZipInfo(name, date_time=info_stamp)
            zi.compress_type = zipfile.ZIP_DEFLATED
            zi.external_attr = 0o600 << 16
            zf.writestr(zi, parts[name])


# --- Рутинные чтения ---------------------------------------------------------
def part_names(parts: Mapping[str, str | bytes], prefix: str) -> list[str]:
    """Имена частей в каталоге (например все `word/footer*.xml`)."""
    return sorted(name for name in parts if name.startswith(prefix))


def parse_rels(parts: Mapping[str, bytes], rels_part: str) -> dict[str, tuple[str, str]]:
    """Relationships: id → (type, target). Целевые пути нормализуются вызывающим кодом."""
    root = parse_part(parts, rels_part)
    out: dict[str, tuple[str, str]] = {}
    if root is None:
        return out
    for rel in list(root):
        rid = rel.get("Id")
        rtype = rel.get("Type")
        target = rel.get("Target")
        if rid and rtype and target:
            out[rid] = (rtype, target)
    return out
