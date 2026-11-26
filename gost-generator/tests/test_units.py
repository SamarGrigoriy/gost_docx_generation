"""Тесты пересчёта единиц OOXML (ТЗ §15, `test_units.py`).

Ловит ошибки округления: часть проверок сравнивает twips с допуском, поэтому
важно, что 1,25 см — это ровно 709, а не 708/710, и что обратный пересчёт
не расходится с прямым больше, чем на допуск.
"""

from __future__ import annotations

import unittest

from app.core import units as u


class TestLengthConversion(unittest.TestCase):
    """мм/см ↔ twips."""

    def test_mm_to_twips_known_values(self) -> None:
        self.assertEqual(u.mm_to_twips(0), 0)
        self.assertEqual(u.mm_to_twips(1), 57)          # 56.6929 → 57
        self.assertEqual(u.mm_to_twips(20), 1134)       # поля педагогического пресета
        self.assertEqual(u.mm_to_twips(30), 1701)       # левое поле ГОСТ
        self.assertEqual(u.mm_to_twips(15), 850)        # правое поле ГОСТ

    def test_cm_to_twips_matches_gost_indent(self) -> None:
        self.assertEqual(u.cm_to_twips(1.25), 709)
        self.assertEqual(u.cm_to_twips(1.5), 850)
        self.assertEqual(u.cm_to_twips(0.5), 283)
        self.assertEqual(u.cm_to_twips(0), 0)

    def test_cm_and_mm_agree(self) -> None:
        for cm in (0.5, 1.25, 1.5, 2.0, 3.0):
            self.assertEqual(u.cm_to_twips(cm), u.mm_to_twips(cm * 10))

    def test_roundtrip_within_one_twip(self) -> None:
        for twips in (283, 709, 850, 1134, 1701):
            self.assertLessEqual(abs(u.mm_to_twips(u.twips_to_mm(twips)) - twips), 1)
            self.assertLessEqual(abs(u.cm_to_twips(u.twips_to_cm(twips)) - twips), 1)


class TestTypography(unittest.TestCase):
    """Кегль и межстрочный интервал."""

    def test_pt_to_half_points(self) -> None:
        self.assertEqual(u.pt_to_half_points(14), 28)
        self.assertEqual(u.pt_to_half_points(12), 24)
        self.assertEqual(u.pt_to_half_points(0.5), 1)

    def test_half_points_roundtrip(self) -> None:
        for pt in (10, 12, 14, 16):
            self.assertEqual(u.half_points_to_pt(u.pt_to_half_points(pt)), float(pt))

    def test_odd_half_points(self) -> None:
        self.assertEqual(u.pt_to_half_points(11.5), 23)
        self.assertEqual(u.half_points_to_pt(23), 11.5)

    def test_spacing_to_line(self) -> None:
        self.assertEqual(u.spacing_to_line(1.0), 240)
        self.assertEqual(u.spacing_to_line(1.15), 276)
        self.assertEqual(u.spacing_to_line(1.5), 360)
        self.assertEqual(u.spacing_to_line(2.0), 480)

    def test_line_to_spacing_roundtrip(self) -> None:
        for spacing in (1.0, 1.15, 1.5, 2.0):
            self.assertAlmostEqual(u.line_to_spacing(u.spacing_to_line(spacing)), spacing,
                                   places=6)


class TestPageGeometry(unittest.TestCase):
    """Формат страницы, EMU и текстовая полоса."""

    def test_a4_portrait_and_landscape(self) -> None:
        self.assertEqual(u.page_size("A4", "portrait"), (11906, 16838))
        self.assertEqual(u.page_size("A4", "landscape"), (16838, 11906))

    def test_unknown_size_falls_back_to_a4(self) -> None:
        self.assertEqual(u.page_size("Letter", "portrait"), u.A4_PORTRAIT)
        self.assertEqual(u.page_size("a4", "landscape"), u.A4_LANDSCAPE)

    def test_cm_to_emu(self) -> None:
        self.assertEqual(u.cm_to_emu(1), 360_000)
        self.assertEqual(u.cm_to_emu(15), 5_400_000)
        self.assertEqual(u.cm_to_emu(0), 0)

    def test_text_width_gost(self) -> None:
        # ГОСТ 7.32: 30 + 15 мм поля → полоса 165 мм = 9355 twips
        self.assertEqual(u.text_width_twips(30, 15), 11906 - 1701 - 850)
        self.assertEqual(u.text_width_twips(30, 15), 9355)

    def test_text_width_landscape(self) -> None:
        width = u.text_width_twips(20, 20, orientation="landscape")
        self.assertEqual(width, 16838 - 2 * u.mm_to_twips(20))


if __name__ == "__main__":
    unittest.main()
