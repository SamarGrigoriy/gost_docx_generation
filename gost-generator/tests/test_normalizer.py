"""Тесты нормализатора .docx (ТЗ §15, `test_normalizer.py`).

Проверяет ровно то, за что нормализатор в ТЗ и ценят: результат идемпотентен,
файл остаётся валидным OOXML-пакетом, содержимое (текст, картинки, ссылки)
не теряется, а защищённые блоки — код-строки и кегль ячеек — не переформатируются.
"""

from __future__ import annotations

import json
import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from app.core import docx_inspector, docx_normalizer
from tests.docx_fixtures import make_docx

REPO = Path(__file__).resolve().parent.parent
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def preset(name: str) -> dict:
    with open(REPO / "presets" / f"{name}.json", encoding="utf-8") as fh:
        return json.load(fh)


def texts_of(path: Path) -> list[str]:
    """Весь текстовый контент документа по абзацам (для проверки «текст не потерялся»)."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    out = []
    for p in root.iter(f"{{{W_NS}}}p"):
        out.append("".join(t.text or "" for t in p.iter(f"{{{W_NS}}}t")))
    return out


class TestNormalizerBasics(unittest.TestCase):
    """Базовый цикл: дефекты → исправлены, файл валиден, бэкап создан."""

    def test_fixes_defects_and_leaves_valid_file(self) -> None:
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.docx"
            make_docx(path, body_font="Arial", body_indent=283, body_line=240)
            self.assertTrue(docx_inspector.check_docx(path, rules),
                            "фикстура должна быть изначально с нарушениями")
            result = docx_normalizer.normalize_docx(path, rules, backup=True)
            self.assertTrue(result["changed"])
            self.assertTrue(result["changes"], "должен быть человекочитаемый список изменений")
            # Файл открывается как пакет и читается инспектором
            with zipfile.ZipFile(path) as z:
                self.assertIsNone(z.testzip())
                self.assertIn("word/document.xml", z.namelist())
                ET.fromstring(z.read("word/document.xml"))
            self.assertEqual([i.code for i in docx_inspector.check_docx(path, rules)
                              if i.severity == "error"], [])

    def test_backup_created_and_not_overwritten(self) -> None:
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.docx"
            make_docx(path, body_font="Arial")
            first = docx_normalizer.normalize_docx(path, rules, backup=True)
            backup = Path(first["backup"])
            self.assertTrue(backup.exists())
            self.assertEqual(backup.name, "bad.orig.docx")
            snapshot = backup.read_bytes()
            # Второй прогон с другим кеглем: бэкап оригинала не должен перезаписаться
            docx_normalizer.normalize_docx(path, rules, backup=True)
            self.assertEqual(backup.read_bytes(), snapshot,
                             ".orig.docx — страховка исходника, не результат правки")

    def test_no_backup_when_disabled(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.docx"
            make_docx(path, body_font="Arial")
            result = docx_normalizer.normalize_docx(path, preset("gost_732"), backup=False)
            self.assertTrue(result["changed"])
            self.assertFalse((Path(tmp) / "bad.orig.docx").exists())
            self.assertIsNone(result["backup"])


class TestNormalizerIdempotency(unittest.TestCase):
    """Повторный прогон не меняет байты XML (п. 9 ТЗ)."""

    def test_second_run_changes_nothing(self) -> None:
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.docx"
            make_docx(path, body_font="Arial", body_indent=283, body_line=240,
                      table_caption_jc="center")
            docx_normalizer.normalize_docx(path, rules, backup=True)
            with zipfile.ZipFile(path) as z:
                after_first = {n: z.read(n) for n in z.namelist()}
            second = docx_normalizer.normalize_docx(path, rules, backup=True)
            self.assertFalse(second["changed"], f"второй прогон что-то поменял: {second['changes']}")
            with zipfile.ZipFile(path) as z:
                for name in ("word/document.xml", "word/styles.xml"):
                    self.assertEqual(z.read(name), after_first[name],
                                     f"{name} differs on second run")


def rels_map(path: Path) -> dict[str, str]:
    """word/_rels/document.xml.rels как {Id: Type} — срез всех связей пакета."""
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/_rels/document.xml.rels"))
    ns = "{http://schemas.openxmlformats.org/package/2006/relationships}"
    return {el.get("Id"): el.get("Type") for el in root.iter(f"{ns}Relationship")}


class TestNormalizerPreservesContent(unittest.TestCase):
    """Нормализация — про оформление, а не про содержимое."""

    def test_text_and_images_and_rels_survive(self) -> None:
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.docx"
            make_docx(path, body_font="Arial", body_indent=283)
            before_names = set(zipfile.ZipFile(path).namelist())
            before_texts = texts_of(path)
            before_rels = rels_map(path)
            with zipfile.ZipFile(path) as z:
                before_media = {n: z.read(n) for n in before_names
                                if n.startswith("word/media/")}
            docx_normalizer.normalize_docx(path, rules, backup=True)
            after_rels = rels_map(path)
            with zipfile.ZipFile(path) as z:
                after_names = set(z.namelist())
                self.assertEqual(set(before_media),
                                 {n for n in after_names if n.startswith("word/media/")})
                for name, blob in before_media.items():
                    self.assertEqual(z.read(name), blob,
                                     f"{name}: байты картинки изменены")
                self.assertEqual(set(before_rels), set(after_rels),
                                 "набор relationship'ов не должен меняться")
                self.assertIn("relationships/image", after_rels["rIdImg"],
                              "связь с медиа должна остаться типа image")
            self.assertEqual(texts_of(path), before_texts, "текст абзацев изменился")

    def test_code_block_untouched(self) -> None:
        """Consolas-блок — не «текст работы»: шрифт и кегль код-строк остаются."""
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "bad.docx"
            make_docx(path, body_font="Arial", body_indent=283, with_code_block=True)
            with zipfile.ZipFile(path) as z:
                before = z.read("word/document.xml").decode("utf-8")
            code_font = "Consolas"
            self.assertIn(code_font, before)
            docx_normalizer.normalize_docx(path, rules, backup=True)
            with zipfile.ZipFile(path) as z:
                after = z.read("word/document.xml").decode("utf-8")
            self.assertIn(f'w:ascii="{code_font}"', after,
                          "нормализатор не должен перекрашивать код-блоки в TNR")

    def test_clean_file_reports_no_changes(self) -> None:
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "good.docx"
            make_docx(path)
            docx_normalizer.normalize_docx(path, rules, backup=True)
            again = docx_normalizer.normalize_docx(path, rules, backup=True)
            self.assertFalse(again["changed"])


class TestNormalizerOtherPresets(unittest.TestCase):
    """Педагогический пресет: альбомная ориентация и кегль таблиц — свои."""

    def test_pedagogical_target_values(self) -> None:
        rules = preset("pedagogical")
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "karta.docx"
            make_docx(path, body_font="Arial", body_line=360, orient="portrait")
            result = docx_normalizer.normalize_docx(path, rules, backup=True)
            self.assertTrue(result["changed"])
            with zipfile.ZipFile(path) as z:
                xml = z.read("word/document.xml").decode("utf-8")
            self.assertIn('w:orient="landscape"', xml)
            self.assertEqual([i.code for i in docx_inspector.check_docx(path, rules)
                              if i.severity == "error"], [])


if __name__ == "__main__":
    unittest.main()
