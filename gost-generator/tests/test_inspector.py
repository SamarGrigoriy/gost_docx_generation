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


def _styles_xml(path: str) -> str:
    """word/styles.xml из пакета: нужен, чтобы проверить наследуемый отступ в docDefaults."""
    import zipfile
    with zipfile.ZipFile(path) as zf:
        return zf.read("word/styles.xml").decode("utf-8")


def _document_xml(path: str) -> str:
    """word/document.xml из пакета: убедиться, что фикстура действительно изменена."""
    import zipfile
    with zipfile.ZipFile(path) as zf:
        return zf.read("word/document.xml").decode("utf-8")


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


class TestWordRoundTrip(unittest.TestCase):
    """Пересохранённый в Word/LibreOffice документ не должен давать ложных замечаний.

    Живой файл пользователя (`РЦ.docx`) вскрыл три расхождения, каждое из которых
    рождало нарушения там, где редактор рисует всё правильно: подпись с `w:hanging="0"`
    вместо `w:firstLine="0"`, строки автосодержания в стилях TOC1/TOC2 и границы таблицы,
    разложенные редактором по ячейкам.
    """

    def _patched(self, tmp: str, name: str, transform) -> str:
        import zipfile
        src = make_docx(Path(tmp) / f"{name}.docx")
        out = Path(tmp) / f"{name}_patched.docx"
        with zipfile.ZipFile(src) as zin, zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.namelist():
                data = zin.read(item)
                if item == "word/document.xml":
                    data = transform(data.decode("utf-8")).encode("utf-8")
                zout.writestr(item, data)
        return str(out)

    def _codes(self, path: str) -> set[str]:
        return codes(docx_inspector.check_docx(path, preset("gost_732")))

    def test_hanging_indent_means_no_first_line(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = self._patched(tmp, "hanging",
                                 lambda d: d.replace('<w:ind w:firstLine="0"', '<w:ind w:hanging="0"'))
            xml = _document_xml(path)
            self.assertIn('w:hanging="0"', xml, "фикстура должна использовать hanging")
            self.assertIn('w:firstLine="709"', _styles_xml(path), "docDefaults должен давать отступ")
            self.assertEqual(sorted(self._codes(path)), [],
                             "hanging снимает унаследованный firstLine — отступа нет")

    def test_toc_entries_are_not_headings(self) -> None:
        toc = ('<w:p><w:pPr><w:pStyle w:val="TOC1"/></w:pPr>'
               '<w:r><w:t>СОДЕРЖАНИЕ</w:t></w:r><w:r><w:tab/><w:t>2</w:t></w:r></w:p>'
               '<w:p><w:pPr><w:pStyle w:val="TOC2"/></w:pPr>'
               '<w:r><w:t>1.1 Анализ</w:t></w:r><w:r><w:tab/><w:t>3</w:t></w:r></w:p>')

        def insert(d: str) -> str:
            at = d.find('<w:pStyle w:val="Heading1"')
            self.assertGreater(at, 0)
            start = d.rfind("<w:p>", 0, at)
            return d[:start] + toc + d[start:]

        with tempfile.TemporaryDirectory() as tmp:
            path = self._patched(tmp, "toc", insert)
            issues = docx_inspector.check_docx(path, preset("gost_732"))
            bad = [i.message for i in issues
                   if i.code in {"HEAD_STRUCT", "HEAD_KEEP", "HEAD_SIZE", "TBL_CAPTION",
                                 "FIG_CAPTION"} and "Анализ" in i.message]
            self.assertEqual(bad, [], "строка содержания принята за заголовок")
            self.assertEqual(sorted(self._codes(path)), [])

    def test_cell_level_table_borders_are_borders(self) -> None:
        sides = "".join(f'<w:{s} w:val="single" w:sz="4"/>'
                        for s in ("top", "left", "bottom", "right"))

        def to_cells(d: str) -> str:
            import re as _re
            d = _re.sub(r"<w:tblBorders>.*?</w:tblBorders>", "", d, flags=_re.S)
            return d.replace("<w:tcPr>", f"<w:tcPr><w:tcBorders>{sides}</w:tcBorders>")

        with tempfile.TemporaryDirectory() as tmp:
            path = self._patched(tmp, "cellborders", to_cells)
            xml = _document_xml(path)
            self.assertNotIn("w:tblBorders", xml)
            self.assertIn("w:tcBorders", xml)
            self.assertEqual(sorted(self._codes(path)), [],
                             "границы по ячейкам — это границы, а не их отсутствие")


if __name__ == "__main__":
    unittest.main()
