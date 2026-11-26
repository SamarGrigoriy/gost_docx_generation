"""Единицы измерения OOXML и пересчёты.

Причина существования модуля: в OOXML все линейные размеры — twips, кегль —
полупоинты, интервалы — 240-е доли. Разбросанные по коду магические числа делают
правки оформления невозможными, поэтому наружу отдаются только именованные функции.
"""

from __future__ import annotations

TWIPS_PER_MM = 56.6929
TWIPS_PER_CM = 566.929
HALF_POINTS_PER_PT = 2
EMU_PER_CM = 360_000
LINE_UNIT = 240  # w:line при lineRule="auto": 1,0 интервала = 240

A4_PORTRAIT = (11906, 16838)
A4_LANDSCAPE = (16838, 11906)


def mm_to_twips(value_mm: float) -> int:
    """Миллиметры → twips с округлением к целому (OOXML хранит только int)."""
    return int(round(value_mm * TWIPS_PER_MM))


def cm_to_twips(value_cm: float) -> int:
    """Сантиметры → twips."""
    return int(round(value_cm * TWIPS_PER_CM))


def twips_to_mm(value_twips: int) -> float:
    """Обратный пересчёт для отображения значений, прочитанных из DOCX."""
    return value_twips / TWIPS_PER_MM


def twips_to_cm(value_twips: int) -> float:
    """Обратный пересчёт для отображения значений, прочитанных из DOCX."""
    return value_twips / TWIPS_PER_CM


def pt_to_half_points(value_pt: float) -> int:
    """Кегль в пунктах → значение w:sz (полупоинты)."""
    return int(round(value_pt * HALF_POINTS_PER_PT))


def half_points_to_pt(value_half_points: int) -> float:
    """w:sz → пункты."""
    return value_half_points / HALF_POINTS_PER_PT


def spacing_to_line(value_spacing: float) -> int:
    """Межстрочный интервал (1.0/1.15/1.5/2.0) → значение w:line при lineRule="auto"."""
    return int(round(value_spacing * LINE_UNIT))


def line_to_spacing(value_line: int) -> float:
    """w:line → привычный интервал."""
    return value_line / LINE_UNIT


def cm_to_emu(value_cm: float) -> int:
    """Сантиметры → EMU (размеры изображений вDrawing-блоках)."""
    return int(round(value_cm * EMU_PER_CM))


def page_size(size: str = "A4", orientation: str = "portrait") -> tuple[int, int]:
    """Размер страницы в twips: (w, h) с учётом ориентации."""
    if size.upper() != "A4":
        # Другие форматки методичка не требует: считаем A4, чтобы вывод был детерминирован.
        size = "A4"
    width, height = A4_PORTRAIT
    if orientation == "landscape":
        return A4_LANDSCAPE
    return width, height


def text_width_twips(margin_left_mm: float, margin_right_mm: float,
                     size: str = "A4", orientation: str = "portrait") -> int:
    """Ширина текстовой полосы (страница минус поля): суммарная ширина колонок таблиц."""
    page_width, _ = page_size(size, orientation)
    return page_width - mm_to_twips(margin_left_mm) - mm_to_twips(margin_right_mm)
