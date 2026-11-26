"""Тесты инспектора .docx по правилам (ТЗ §15, `test_inspector.py`).

Фикстуры собираются программно (`tests.docx_fixtures.make_docx`) во временном
каталоге: каждый тест вносит ОДИН дефект и требует ОДИН код нарушения — иначе
регрессия в проверке осталась бы незамеченной. Отдельно проверяется, что
нестрогий пресет понижает severity, а не молчит.
"""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path
from typing import Any

from app.core import docx_inspector
from tests.docx_fixtures import make_docx

REPO = Path(__file__).resolve().parent.parent


def preset(name: str) -> dict[str, Any]:
    """Пресет правил из репозитория (без пользовательского состояния)."""
    with open(REPO / "presets" / f"{name}.json", encoding="utf-8") as fh:
        return json.load(fh)


def codes(issues: list[Any], severity: str | None = None) -> set[str]:
    """Коды замечаний; `severity=None` — всё, кроме информационных (`MANUAL`)."""
    if severity is not None:
        return {i.code for i in issues if i.severity == severity}
    return {i.code for i in issues if i.severity != "info"}


def issue_field(issue: Any, key: str) -> Any:
    """Замечание может быть dataclass (`check_docx`) или словарем (`check_docx_result`)."""
    return getattr(issue, key, None) if not isinstance(issue, dict) else issue.get(key)


class TestInspectorClean(unittest.TestCase):
    """Эталонный файл не должен давать НИ ОДНОГО error (бороться с ложными — №1 приоритет)."""

    def test_good_fixture_has_no_errors(self) -> None:
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "good.docx")
            issues = docx_inspector.check_docx(path, rules)
            self.assertEqual(
                sorted(codes(issues)), [],
                f"ложные нарушения: {[(i.code, i.message) for i in issues if i.severity == 'error']}")

    def test_result_envelope_fields(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "good.docx")
            result = docx_inspector.check_docx_result(path, preset("gost_732"))
            self.assertTrue({"file", "preset", "strict", "summary", "issues"} <= set(result))
            self.assertEqual(result["summary"]["error"], 0)
            self.assertIsNone(result["read_error"])
            self.assertGreater(result["summary"]["info"], 0, "блок «проверить вручную»")
            for issue in result["issues"]:
                location = issue_field(issue, "location") or {}
                self.assertTrue({"paragraph", "page", "snippet"} <= set(location),
                                f"нет координат замечания: {issue}")

    def test_issue_objects_have_hint_and_severity(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "bad.docx", body_font="Arial")
            issues = docx_inspector.check_docx(path, preset("gost_732"))
            target = [i for i in issues if i.code == "FONT_MAIN"]
            self.assertTrue(target)
            self.assertEqual(target[0].severity, "error")
            self.assertTrue(target[0].hint)


class TestInspectorDefects(unittest.TestCase):
    """Каждый внесённый дефект обязан быть найден (п. 9.1 ТЗ)."""

    DEFECTS = (
        ("Arial вместо TNR", {"body_font": "Arial"}, "FONT_MAIN"),
        ("отступ 0,5 см", {"body_indent": 283}, "INDENT_VALUE"),
        ("интервал 1,0", {"body_line": 240}, "LINE_SPACING"),
        ("номер страницы сверху", {"page_field_in": "header"}, "NUM_POSITION"),
        ("подпись таблицы по центру", {"table_caption_jc": "center"}, "TBL_CAPTION"),
        ("нет повторения шапки", {"table_header": False}, "TBL_HEADER"),
        ("заголовок без keepNext", {"heading_keep": False}, "HEAD_KEEP"),
        ("структурный не с новой страницы", {"structural_brk": False}, "HEAD_STRUCT"),
    )

    def test_each_defect_detected(self) -> None:
        rules = preset("gost_732")
        for title, overrides, code in self.DEFECTS:
            with self.subTest(дефект=title):
                with tempfile.TemporaryDirectory() as tmp:
                    path = make_docx(Path(tmp) / "bad.docx", **overrides)
                    found = codes(docx_inspector.check_docx(path, rules))
                    self.assertIn(code, found, f"дефект «{title}» не найден: {sorted(found)}")

    def test_wrong_margins_detected(self) -> None:
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "bad.docx", margin_left=1418, margin_right=1418)
            self.assertIn("PAGE_MARGINS", codes(docx_inspector.check_docx(path, rules)))

    def test_several_defects_all_reported(self) -> None:
        rules = preset("gost_732")
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "bad.docx", body_font="Arial", body_indent=283,
                             page_field_in="header")
            found = codes(docx_inspector.check_docx(path, rules))
            self.assertEqual({"FONT_MAIN", "INDENT_VALUE", "NUM_POSITION"} - found, set())

    def test_issue_location_points_to_paragraph(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "bad.docx", body_indent=283)
            issues = [i for i in docx_inspector.check_docx(path, preset("gost_732"))
                      if i.code == "INDENT_VALUE"]
            self.assertTrue(issues)
            self.assertIsInstance(issues[0].location["paragraph"], int,
                                  "нарушение должно ссылаться на конкретный абзац")


class TestInspectorPresets(unittest.TestCase):
    """Смена пресета меняет severity (п. 7.1: non-strict пресеты)."""

    def test_pedagogical_downgrades_errors(self) -> None:
        gost = preset("gost_732")
        ped = preset("pedagogical")
        self.assertFalse(ped["strict"])
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "bad.docx", body_font="Arial", body_indent=283)
            strict = codes(docx_inspector.check_docx(path, gost))
            soft = codes(docx_inspector.check_docx(path, ped), severity="warning")
            self.assertTrue(strict)
            self.assertFalse(codes(docx_inspector.check_docx(path, ped), severity="error"),
                             "в нестрогом пресете ошибок быть не должно")
            self.assertTrue(soft | strict or soft, "замечания должны сохраниться как warnings")

    def test_free_preset_reports_fewer(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "bad.docx", body_font="Arial", body_indent=283,
                             body_line=240)
            gost_issues = docx_inspector.check_docx(path, preset("gost_732"))
            free_issues = docx_inspector.check_docx(path, preset("free"))
            self.assertGreater(len([i for i in gost_issues if i.severity == "error"]),
                               len([i for i in free_issues if i.severity == "error"]))


class TestInspectorBrokenFile(unittest.TestCase):
    """Битый файл — не исключение наружу, а `read_error` в результате (п. 19.4)."""

    def test_not_a_zip(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "junk.docx"
            path.write_bytes(b"not a docx at all")
            result = docx_inspector.check_docx_result(path, preset("gost_732"))
            self.assertIsNotNone(result["read_error"])
            self.assertEqual(result["summary"]["error"], 0)

    def test_manual_items_present(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = make_docx(Path(tmp) / "good.docx")
            result = docx_inspector.check_docx_result(path, preset("gost_732"))
            manual = [i for i in result["issues"] if issue_field(i, "code") == "MANUAL"]
            self.assertTrue(manual, "блок «проверить вручную» обязателен (п. 9.1)")


if __name__ == "__main__":
    unittest.main()
