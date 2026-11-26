"""Тесты отчётного модуля (ТЗ §19.3 `--report`, §20 — валидируется схемой `report`).

Отчёт читает человек и агент, поэтому проверяется структура, а не только факт
непустой строки: заголовки разделов, корректная markdown-таблица (со строкой
разделителя), экранирование pipe в текстах, русский статус в пакетной сводке
и поведение на нечитаемом файле.
"""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from typing import Any

from app.core import report

REPO = Path(__file__).resolve().parent.parent


def rules() -> dict[str, Any]:
    with open(REPO / "presets" / "gost_732.json", encoding="utf-8") as fh:
        return json.load(fh)


RESULT = {
    "file": "/tmp/отчёт 1.docx",
    "preset": "gost_732",
    "strict": True,
    "read_error": None,
    "summary": {"error": 2, "warning": 1, "info": 3},
    "issues": [
        {"code": "FONT_MAIN", "severity": "error", "message": "шрифт Arial | не тот",
         "location": {"paragraph": 12, "page": 3, "snippet": "текст абзаца"},
         "hint": "задать Times New Roman"},
        {"code": "TBL_CAPTION", "severity": "warning", "message": "подпись по центру",
         "location": {"paragraph": 25, "page": 5, "snippet": "Таблица 1 – Название"},
         "hint": "jc=left"},
        {"code": "MANUAL", "severity": "info", "message": "Проверить вручную: орфография.",
         "location": {"paragraph": None, "page": None, "snippet": ""},
         "hint": "F7 в Word"},
    ],
}


class TestRenderMarkdown(unittest.TestCase):
    def test_sections_and_summary_table(self) -> None:
        text = report.render_markdown(RESULT, rules(), stamp="2026-11-25 10:00")
        self.assertIn("# Отчёт проверки", text)
        self.assertIn("## Сводка", text)
        self.assertIn("|---:|---:|---:|", text, "markdown-таблица без разделителя не рендерится")
        self.assertIn("отчёт 1.docx", text)
        self.assertIn("FONT_MAIN", text)
        self.assertIn("2026-11-25 10:00", text)

    def test_manual_block_present(self) -> None:
        text = report.render_markdown(RESULT, rules())
        self.assertIn("Проверить вручную", text)
        self.assertIn("орфография", text)

    def test_pipe_escaped_in_cells(self) -> None:
        text = report.render_markdown(RESULT, rules())
        row = next((line for line in text.splitlines() if "FONT_MAIN" in line), "")
        self.assertIn("\\|", row, "неэкранированный `|` ломает таблицу")
        self.assertGreaterEqual(row.count("|") - row.count("\\|"), 3)

    def test_read_error_short_circuits(self) -> None:
        broken = dict(RESULT, read_error="не zip", issues=[], summary={"error": 0, "warning": 0,
                                                                       "info": 0})
        text = report.render_markdown(broken, rules())
        self.assertIn("не zip", text)
        self.assertNotIn("## Сводка", text)


class TestRenderText(unittest.TestCase):
    def test_header_and_counts(self) -> None:
        text = report.render_text(RESULT, rules(), stamp="2026-11-25 10:00")
        self.assertIn("ОТЧЁТ ПРОВЕРКИ", text)
        self.assertIn("Режим: strict", text)
        self.assertIn("FONT_MAIN", text)

    def test_soft_mode_label(self) -> None:
        soft = dict(RESULT, strict=False)
        self.assertIn("Режим: мягкий", report.render_text(soft, rules()))

    def test_summary_wording(self) -> None:
        # Формулировка с русскими склонениями: «2 ошибки, 1 предупреждение, 3 рекомендации»
        self.assertEqual(report.format_summary(RESULT["summary"]),
                         "2 ошибки, 1 предупреждение, 3 рекомендации")
        self.assertEqual(report.format_summary({"error": 1, "warning": 0, "info": 0}), "1 ошибка")
        self.assertEqual(report.format_summary({"error": 5, "warning": 0, "info": 0}), "5 ошибок")
        self.assertEqual(report.format_summary({}), "Замечаний нет")


class TestSaveReport(unittest.TestCase):
    def test_markdown_extension_and_content(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            dest = report.save_report(RESULT, rules(), Path(tmp) / "check/report.md")
            self.assertTrue(dest.exists())
            self.assertIn("FONT_MAIN", dest.read_text(encoding="utf-8"))

    def test_extension_added_when_missing(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            dest = report.save_report(RESULT, rules(), Path(tmp) / "otchet", fmt="txt")
            self.assertEqual(dest.suffix, ".txt")

    def test_unknown_format_raises(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(ValueError):
                report.save_report(RESULT, rules(), Path(tmp) / "x", fmt="docx")


class TestBatchReport(unittest.TestCase):
    ITEMS = [
        {"file": "a.docx", "summary": {"error": 0, "warning": 0, "info": 1}, "status": "ok"},
        {"file": "b.docx", "summary": {"error": 3, "warning": 1, "info": 0},
         "status": "violations"},
        {"file": "c.docx", "summary": {}, "read_error": "битый zip", "status": "Ошибка чтения"},
    ]

    def test_batch_markdown_rows(self) -> None:
        text = report.render_batch_markdown(self.ITEMS, stamp="2026-11-25 10:00")
        for name in ("a.docx", "b.docx", "c.docx"):
            self.assertIn(name, text)
        self.assertIn("|---", text)
        self.assertIn("Есть ошибки", text, "статус выводится по-русски (UI/CLI)")

    def test_batch_status_keys_normalized(self) -> None:
        text = report.render_batch_markdown(
            [{"file": "x.docx", "summary": {"error": 1, "warning": 0, "info": 0},
              "status": "violations"}])
        self.assertIn("Есть ошибки", text)

    def test_batch_text_counts(self) -> None:
        text = report.render_batch_text(self.ITEMS)
        self.assertIn("a.docx", text)
        self.assertIn("3", text)

    def test_status_fallback_from_summary(self) -> None:
        self.assertEqual(report._fallback_status({"error": 2, "warning": 0, "info": 0}, None),
                         "Есть ошибки")
        self.assertEqual(report._fallback_status({"error": 0, "warning": 1, "info": 0}, None),
                         "Есть замечания")
        self.assertEqual(report._fallback_status({"error": 0, "warning": 0, "info": 0}, None),
                         "Без замечаний")
        self.assertEqual(report._fallback_status({}, "битый zip"), "Ошибка чтения")

    def test_save_batch_report(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            dest = report.save_batch_report(self.ITEMS, Path(tmp) / "batch.md")
            self.assertTrue(dest.exists())
            self.assertIn("b.docx", dest.read_text(encoding="utf-8"))


if __name__ == "__main__":
    unittest.main()
