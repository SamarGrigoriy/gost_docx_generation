"""Публичный API нормализации DOCX: `normalize_docx(path, rules, backup=True)`.

Правит файл НА МЕСТЕ (пост-обработка документа, сгенерированного скриптом):

1. читает весь пакет в память (`zipfile` + `xml.etree`, без python-docx/lxml);
2. строит ТУ ЖЕ модель, что и инспектор, — правки и проверки видят документ одинаково;
3. применяет нормы по порядку: `styles.xml` → `w:sectPr` → текст/заголовки/подписи/таблицы
   → нумерация и колонтитулы → `settings.xml` → сноски;
4. сериализует обратно ТОЛЬКО изменённые части;
5. делает бэкап `<name>.orig.docx` (существующий не перезаписывает) и детерминированно
   перезаписывает zip (фиксированная дата записи → повторный прогон даёт байт-в-байт тот же
   XML, то есть идемпотентность проверяема над файлом, а не «на глаз»).

Трогать не положено: `word/media/*`, существующие relationships и моноширинные run'ы
(код-блоки). Новая relationship/Override появляется только если в пакете нет вообще ни
одного подходящего колонтитула — и дописывается строкой в конец файла, чтобы не
переписывать чужие записи.
"""

from __future__ import annotations

import os
import re
import shutil
from typing import Any, Callable, MutableMapping

import xml.etree.ElementTree as ET

from app import constants
from app.core import docx_norm_body as body
from app.core import docx_norm_numbering as numbering
from app.core import docx_xml as dx
from app.core.docx_norm_styles import normalize_styles
from app.core.docx_norm_util import NormCtx
from app.core.docx_parser import DocxParser

__all__ = ["normalize_docx", "backup_path_for"]

RELS_PART = "word/_rels/document.xml.rels"
CONTENT_TYPES_PART = "[Content_Types].xml"

_REL_TYPES = {
    "footer": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer",
    "header": "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header",
}
_CONTENT_TYPES = {
    "footer": "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml",
    "header": "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml",
}
_RID_RE = re.compile(r"^rId(\d+)$")


def backup_path_for(path: str) -> str:
    """Путь резервной копии: `<name>.orig.docx` рядом с оригиналом."""
    base, ext = os.path.splitext(str(path))
    return f"{base}.orig{ext or '.docx'}"


def _make_backup(path: str) -> str:
    """Создать бэкап, если его ещё нет: повторные прогоны не должны затирать оригинал."""
    dst = backup_path_for(path)
    if os.path.exists(dst):
        return dst
    try:
        shutil.copy2(path, dst)
    except OSError as exc:
        raise dx.DocxError(f"Не удалось создать резервную копию {dst}: {exc}") from exc
    return dst


def _next_rid(rels: dict[str, tuple[str, str]]) -> str:
    """Свободный числовой r:id: генераторы именуют id по-разному (rId1, rIdStyles…) —
    берём максимум по числовым и идём дальше, чтобы не столкнуть существующие ссылки."""
    used = set(rels)
    numbers = [int(m.group(1)) for rid in used if (m := _RID_RE.match(rid))]
    candidate = max(numbers, default=0) + 1
    while f"rId{candidate}" in used:
        candidate += 1
    return f"rId{candidate}"


def _append_relationship(parts: MutableMapping[str, bytes], rid: str, rtype: str,
                         target: str) -> None:
    """Дописать одну запись в relationships СТРОКОЙ: сериализация через ElementTree
    переписала бы весь файл и могла изменить чужие записи, а нужен ровно один элемент."""
    item = f'<Relationship Id="{rid}" Type="{rtype}" Target="{target}"/>'
    raw = parts.get(RELS_PART)
    if raw is None:
        parts[RELS_PART] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            f"{item}</Relationships>").encode("utf-8")
        return
    text = raw.decode("utf-8", "replace")
    if "</Relationships>" in text:
        parts[RELS_PART] = text.replace("</Relationships>", f"{item}</Relationships>",
                                        1).encode("utf-8")
    else:  # самозакрывающийся корень: <Relationships ... />
        parts[RELS_PART] = re.sub(r"(?<!/)>\s*$", f">{item}</Relationships>", text,
                                  count=1).encode("utf-8")


def _append_content_type(parts: MutableMapping[str, bytes], part_name: str, ctype: str) -> None:
    """Override новой части в [Content_Types].xml (также строковая вставка)."""
    raw = parts.get(CONTENT_TYPES_PART)
    if raw is None:
        return
    text = raw.decode("utf-8", "replace")
    if f'PartName="/{part_name}"' in text:
        return
    override = f'<Override PartName="/{part_name}" ContentType="{ctype}"/>'
    if "</Types>" in text:
        parts[CONTENT_TYPES_PART] = text.replace("</Types>", f"{override}</Types>",
                                                1).encode("utf-8")


def _make_add_part(model: Any, parts: MutableMapping[str, bytes],
                   order: list[str]) -> Callable[[str, str], str]:
    """Фабрика новых частей колонтитулов: имя → байты → relationship → content-type → r:id."""

    def add_part(kind: str, xml_text: str) -> str:
        taken = {name.rsplit("/", 1)[-1] for name in parts}
        index = 1
        while f"{kind}{index}.xml" in taken:
            index += 1
        name = f"word/{kind}{index}.xml"
        parts[name] = xml_text.encode("utf-8")
        if name not in order:
            order.append(name)  # write_package обходит части в этом порядке
        rid = _next_rid(model.rels)
        _append_relationship(parts, rid, _REL_TYPES[kind], f"{kind}{index}.xml")
        _append_content_type(parts, name, _CONTENT_TYPES[kind])
        model.rels[rid] = (_REL_TYPES[kind], name)
        root = dx.parse_part(parts, name)
        if root is not None:
            (model.footer_parts if kind == "footer" else model.header_parts)[name] = root
        return rid

    return add_part


def _apply_serialized(parts: MutableMapping[str, bytes],
                      dirty: MutableMapping[str, ET.Element]) -> None:
    """Сериализовать изменённые деревья в байты пакета.

    `serialize_part` возвращает объявления `xmlns:*` из исходной части: без них
    `mc:Ignorable="w14"` ссылался бы на необъявленный префикс и Word отверг бы файл.
    """
    for name, root in dirty.items():
        parts[name] = dx.serialize_part(root, parts.get(name))


def normalize_docx(path: str, rules: dict[str, Any], backup: bool = True) -> dict[str, Any]:
    """Привести DOCX к правилам пресета; вернуть список внесённых изменений.

    Args:
        path: файл DOCX — меняется на месте, если `changed == True`.
        rules: словарь пресета (см. `presets/gost_732.json`).
        backup: создавать ли `<name>.orig.docx` перед первой правкой.

    Returns:
        `{"changed": bool, "changes": list[str], "backup": str | None}`.

    Raises:
        DocxError: битый zip/XML, нет обязательных частей, ошибка записи файла.
    """
    parts, order = dx.read_package(str(path))
    model = DocxParser(mono_fonts=constants.MONOSPACE_FONTS).parse_parts(str(path), parts, order)
    ctx = NormCtx(model=model, rules=rules, add_part=_make_add_part(model, parts, order))

    # 1) Стили — раньше всего: инспектор читает реальные свойства по цепочке basedOn,
    #    и без docDefaults/Normal правки текста дали бы cosmetic-эффект.
    normalize_styles(model.styles, rules, ctx)

    # 2) Секции: размер и поля страницы.
    numbering.normalize_sections(ctx)

    # 3) Текст и объекты документа.
    doc_before = len(ctx.changes)
    body.normalize_run_fonts(ctx)
    body.normalize_body_paragraphs(ctx)
    body.normalize_structural_headings(ctx)
    body.normalize_heading_styles(ctx)
    body.normalize_captions(ctx)
    body.normalize_tables(ctx)
    if len(ctx.changes) > doc_before:
        ctx.dirty["word/document.xml"] = model.document

    # 4) Нумерация страниц, колонтитулы, поля, settings.
    numbering.normalize_page_numbering(ctx)
    numbering.insert_toc_field(ctx)
    numbering.ensure_update_fields(ctx)

    # 5) Сноски.
    notes_before = len(ctx.changes)
    body.normalize_footnotes(ctx)
    if len(ctx.changes) > notes_before and model.footnotes is not None:
        ctx.dirty["word/footnotes.xml"] = model.footnotes

    if not ctx.changes:
        return {"changed": False, "changes": [], "backup": None}

    _apply_serialized(parts, ctx.dirty)
    backup_name = _make_backup(str(path)) if backup else None
    tmp = f"{path}.tmp"
    try:
        dx.write_package(tmp, parts, order)
        os.replace(tmp, str(path))
    except OSError as exc:
        if os.path.exists(tmp):
            os.remove(tmp)
        raise dx.DocxError(f"Не удалось записать нормализованный файл {path}: {exc}") from exc
    return {"changed": True, "changes": ctx.changes, "backup": backup_name}
