"""Хранилище пресетов правил оформления: CRUD, валидация, слияние с дефолтами, сводка.

`DEFAULT_RULES` — встроенная копия ГОСТ-пресета: она гарантирует работоспособность
ядра, если каталог `presets/` недоступен (битая сборка, удаление папки). В PyInstaller-
сборке штатные пресеты лежат в `sys._MEIPASS/presets` (распаковываются при старте).
"""

from __future__ import annotations

import copy
import re
import sys
from pathlib import Path
from typing import Any

from app import constants
from app.core import config

# Порядок: section -> (тип, min, max); значения — из ТЗ §5 и methodichki.
_SIZE_RANGE = (8.0, 72.0)
_SPACING_RANGE = (0.5, 3.0)
_MARGIN_RANGE = (5.0, 50.0)  # мм

REQUIRED_SECTIONS: tuple[str, ...] = (
    "page", "text", "headings", "figures", "tables", "formulas", "notes",
    "footnotes", "references", "appendices", "toc", "numbering", "volume",
    "code_blocks",
)

DEFAULT_RULES: dict[str, Any] = {
    "schema_version": 1,
    "name": "ГОСТ 7.32 (Мультиворк)",
    "builtin": True,
    "strict": True,
    "page": {
        "size": "A4",
        "orientation": "portrait",
        "margin_left_mm": 30,
        "margin_right_mm": 15,
        "margin_top_mm": 20,
        "margin_bottom_mm": 20,
        "footer_distance_mm": 12.5,
    },
    "text": {
        "font": "Times New Roman",
        "size_pt": 14,
        "line_spacing": 1.5,
        "first_line_indent_cm": 1.25,
        "alignment": "justify",
        "space_before_pt": 0,
        "space_after_pt": 0,
        "allow_italic_for_terms_only": True,
        "hyphenation": False,
    },
    "headings": {
        "structural_uppercase": True,
        "structural_bold": True,
        "structural_align": "center",
        "structural_new_page": True,
        "structural_no_trailing_dot": True,
        "section_bold": True,
        "section_from_indent": True,
        "section_numbering": "1",
        "subsection_numbering": "1.1",
        "no_trailing_dot": True,
        "multiline_line_spacing": 1.0,
        "keep_with_next": True,
        "min_lines_after": 3,
        "same_size_as_body": True,
    },
    "figures": {
        "caption_template": "Рисунок {n} – {title}",
        "caption_position": "below",
        "caption_align": "center",
        "caption_indent": False,
        "caption_no_trailing_dot": True,
        "numbering": "continuous",
        "dash": "–",
        "max_width_cm": 15,
        "max_height_cm": 18,
        "require_reference_in_text": True,
        "reference_template": "в соответствии с рисунком {n}",
    },
    "tables": {
        "caption_template": "Таблица {n} – {title}",
        "caption_position": "above",
        "caption_align": "left",
        "caption_indent": False,
        "caption_no_trailing_dot": True,
        "continuation_text": "Продолжение таблицы {n}",
        "continuation_align": "right",
        "repeat_header": True,
        "header_bold": True,
        "borders": "single_thin",
        "border_size_eighths_pt": 4,
        "font_size_pt": 12,
        "line_spacing": 1.0,
        "cell_vertical_align": "center",
        "layout": "fixed",
        "allow_row_split": True,
        "require_reference_in_text": True,
    },
    "formulas": {
        "own_line": True,
        "align": "center",
        "blank_line_above_below": True,
        "number_format": "({n})",
        "number_align": "right",
        "explanation_first_word": "где",
        "explanation_no_colon": True,
    },
    "notes": {
        "prefix": "Примечание",
        "dash": "–",
        "from_indent": True,
        "numbered_when_multiple": True,
    },
    "footnotes": {
        "font": "Times New Roman",
        "size_pt": 12,
        "min_allowed_size_pt": 10,
        "superscript_no_space": True,
    },
    "references": {
        "style": "square_brackets",
        "bracket_template": "[{n}, с. {page}]",
        "bibliography_order": "by_appearance",
        "bibliography_numbering": "1.",
        "bibliography_from_indent": True,
        "max_source_age_years": 5,
        "min_cited_percent": 80,
    },
    "appendices": {
        "label": "ПРИЛОЖЕНИЕ {letter}",
        "new_page": True,
        "label_align": "center",
        "title_bold": True,
        "title_align": "center",
        "letters_exclude": ["Ё", "З", "Й", "О", "Ч", "Ъ", "Ы", "Ь"],
    },
    "toc": {"required": True, "auto": True},
    "numbering": {
        "position": "bottom_center",
        "title_page_numbered": False,
        "second_page_shows": 2,
        "font_size_pt": 12,
    },
    "volume": {
        "count_from": "introduction",
        "count_to": "conclusion_end",
        "min_page_fill_percent": 80,
        "min_last_page_fill_percent": 33,
    },
    "code_blocks": {
        "font": "Consolas",
        "size_pt": 10,
        "line_spacing": 1.0,
        "indent": False,
    },
}


class CannotEditBuiltinError(Exception):
    """Попытка перезаписать или удалить встроенный пресет."""


class RulesStore:
    """CRUD пресетов правил. Пользовательские — в `dirs.presets`, встроенные — в `builtin_dir`."""

    def __init__(self, dirs: config.DataDirs, builtin_dir: Path | str | None = None) -> None:
        self.dirs = dirs
        if builtin_dir is not None:
            self.builtin_dir = Path(builtin_dir)
        elif getattr(sys, "frozen", False):
            # В exe штатные пресеты распакованы PyInstaller в _MEIPASS.
            self.builtin_dir = Path(str(sys._MEIPASS)) / constants.DIR_PRESETS
        else:
            self.builtin_dir = Path(dirs.project_dir) / constants.DIR_PRESETS

    # --- внутреннее ---

    def _user_path(self, key: str) -> Path:
        return Path(self.dirs.presets) / f"{key}.json"

    def _builtin_path(self, key: str) -> Path:
        return self.builtin_dir / f"{key}.json"

    def _is_builtin_key(self, key: str) -> bool:
        return self._builtin_path(key).is_file()

    @staticmethod
    def _slug(text: str) -> str:
        cleaned = re.sub(r"[^0-9A-Za-z_-]+", "_", text.strip()).strip("_")
        return cleaned or "custom"

    def _unique_key(self, base: str) -> str:
        key = self._slug(base)
        candidate = key
        index = 2
        while self._user_path(candidate).exists() or self._is_builtin_key(candidate):
            candidate = f"{key}_{index}"
            index += 1
        return candidate

    # --- публичное API ---

    def list(self) -> list[dict[str, Any]]:
        """Список пресетов: элементы {"key","name","builtin","path"}; пользовательские в приоритете."""
        items: dict[str, dict[str, Any]] = {}
        for path in sorted(self.builtin_dir.glob("*.json")) if self.builtin_dir.is_dir() else []:
            data = config.load_json(path, default={})
            name = data.get("name", path.stem) if isinstance(data, dict) else path.stem
            builtin = bool(data.get("builtin", True)) if isinstance(data, dict) else True
            items[path.stem] = {"key": path.stem, "name": name,
                                "builtin": builtin, "path": str(path)}
        user_dir = Path(self.dirs.presets)
        for path in sorted(user_dir.glob("*.json")) if user_dir.is_dir() else []:
            data = config.load_json(path, default={})
            name = data.get("name", path.stem) if isinstance(data, dict) else path.stem
            items[path.stem] = {"key": path.stem, "name": name,
                                "builtin": False, "path": str(path)}
        return list(items.values())

    def load(self, key_or_path: str | Path) -> dict[str, Any]:
        """Загружает пресет по ключу или пути к .json; результат всегда полной структуры."""
        p = Path(key_or_path)
        if p.suffix == ".json" and p.is_file():
            raw = config.load_json(p, default=None)
        else:
            key = str(key_or_path)
            for candidate in (self._user_path(key), self._builtin_path(key)):
                if candidate.is_file():
                    raw = config.load_json(candidate, default=None)
                    break
            else:
                if key == constants.DEFAULT_PRESET_KEY:
                    # Каталог пресетов утрачен — встроенная копия спасает запуск.
                    return copy.deepcopy(DEFAULT_RULES)
                raise FileNotFoundError(f"Пресет не найден: {key_or_path}")
        if not isinstance(raw, dict):
            if str(key_or_path) == constants.DEFAULT_PRESET_KEY:
                return copy.deepcopy(DEFAULT_RULES)
            raise FileNotFoundError(f"Пресет не найден или повреждён: {key_or_path}")
        return merge_with_defaults(config.migrate(raw, "rules", DEFAULT_RULES))

    def save(self, rules: dict[str, Any], key: str | None = None) -> str:
        """Сохраняет пользовательский пресет; под builtin-ключом пишет только через save_as."""
        if key is None:
            key = str(rules.get("key") or rules.get("name") or "custom")
            key = self._unique_key(key)
        if self._is_builtin_key(key):
            raise CannotEditBuiltinError(
                f"Пресет «{key}» встроенный и защищён от записи. "
                f"Используйте «Сохранить как…» — создастся пользовательская копия.")
        return self.save_as(rules, key)

    def save_as(self, rules: dict[str, Any], new_key: str) -> str:
        """Создаёт/перезаписывает пользовательский пресет с указанным ключом (затеняет builtin)."""
        if not str(new_key).strip():
            raise ValueError("Ключ пресета не может быть пустым")
        data = copy.deepcopy(dict(rules))
        data["builtin"] = False
        data["schema_version"] = constants.SCHEMA_VERSION
        data.setdefault("name", str(new_key))
        save_key = self._slug(str(new_key))
        save_json_path = self._user_path(save_key)
        if save_key != str(new_key) and (
            save_json_path.exists() or self._is_builtin_key(save_key)
        ):
            save_key = self._unique_key(str(new_key))
        config.save_json(self._user_path(save_key), data)
        return save_key

    def delete(self, key: str) -> None:
        """Удаляет пользовательский пресет; встроенный — нельзя (только удалить его тень)."""
        user = self._user_path(key)
        if user.is_file():
            user.unlink()
            return
        if self._is_builtin_key(key):
            raise CannotEditBuiltinError(f"Встроенный пресет «{key}» удалить нельзя.")
        raise FileNotFoundError(f"Пресет не найден: {key}")

    def export(self, key: str, dest: Path | str) -> Path:
        """Экспорт пресета в произвольный файл (для передачи/архива)."""
        return config.save_json(Path(dest), self.load(key))

    def import_from(self, src: Path | str) -> str:
        """Импорт пресета из внешнего JSON; всегда создаётся пользовательский ключ."""
        raw = config.load_json(Path(src), default=None)
        if not isinstance(raw, dict):
            raise ValueError(f"Файл не является пресетом правил: {src}")
        return self.save(raw, self._unique_key(Path(src).stem))

    def resolve(self, name_or_path: str | Path) -> dict[str, Any]:
        """Разрешает имя пресета, ключ или путь к .json в полный словарь правил."""
        p = Path(str(name_or_path))
        if p.suffix == ".json" and p.is_file():
            return self.load(p)
        try:
            return self.load(str(name_or_path))
        except FileNotFoundError:
            pass
        wanted = str(name_or_path).strip().casefold()
        for item in self.list():
            if wanted in (item["key"].casefold(), str(item["name"]).casefold()):
                return self.load(item["key"])
        raise FileNotFoundError(f"Пресет не найден: {name_or_path}")


def merge_with_defaults(rules: dict[str, Any]) -> dict[str, Any]:
    """Глубокое слияние: недостающие секции/ключи берутся из DEFAULT_RULES."""
    merged = copy.deepcopy(DEFAULT_RULES)
    _deep_update(merged, rules if isinstance(rules, dict) else {})
    return merged


def _deep_update(base: dict, overlay: dict) -> dict:
    for key, value in overlay.items():
        if isinstance(value, dict) and isinstance(base.get(key), dict):
            _deep_update(base[key], value)
        else:
            base[key] = copy.deepcopy(value)
    return base


def _check_number(errors: list[str], section: dict, key: str,
                  lo: float, hi: float, label: str) -> None:
    value = section.get(key)
    if value is None:
        return
    if isinstance(value, bool) or not isinstance(value, (int, float)):
        errors.append(f"{label}: ожидалось число, получено {type(value).__name__}")
        return
    if not lo <= float(value) <= hi:
        errors.append(f"{label}: значение {value} вне диапазона {lo}..{hi}")


def validate_rules(rules: dict[str, Any]) -> list[str]:
    """Проверяет структуру и диапазоны пресета. Пустой список — пресет корректен."""
    errors: list[str] = []
    if not isinstance(rules, dict):
        return ["Правила должны быть JSON-объектом (dict)"]
    for section in REQUIRED_SECTIONS:
        value = rules.get(section)
        if value is None:
            errors.append(f"Отсутствует обязательная секция «{section}»")
        elif not isinstance(value, dict):
            errors.append(f"Секция «{section}» должна быть объектом")
    page = rules.get("page") if isinstance(rules.get("page"), dict) else {}
    text = rules.get("text") if isinstance(rules.get("text"), dict) else {}
    orientation = page.get("orientation")
    if orientation is not None and orientation not in constants.ORIENTATIONS:
        errors.append(f"Ориентация «{orientation}»: допустимы portrait/landscape")
    alignment = text.get("alignment")
    if alignment is not None and alignment not in constants.ALIGNMENTS:
        errors.append(f"Выравнивание «{alignment}»: допустимы "
                      + "/".join(constants.ALIGNMENTS))
    for key in ("margin_left_mm", "margin_right_mm", "margin_top_mm", "margin_bottom_mm"):
        _check_number(errors, page, key, *_MARGIN_RANGE, label=f"Поле «{key}» (мм)")
    for section_name, keys in (
        ("text", ("size_pt",)),
        ("footnotes", ("size_pt", "min_allowed_size_pt")),
        ("tables", ("font_size_pt",)),
        ("numbering", ("font_size_pt",)),
        ("code_blocks", ("size_pt",)),
    ):
        section = rules.get(section_name)
        if isinstance(section, dict):
            for key in keys:
                _check_number(errors, section, key, *_SIZE_RANGE,
                              label=f"Кегль «{section_name}.{key}» (пт)")
    for section_name, keys in (
        ("text", ("line_spacing",)),
        ("headings", ("multiline_line_spacing",)),
        ("tables", ("line_spacing",)),
        ("code_blocks", ("line_spacing",)),
    ):
        section = rules.get(section_name)
        if isinstance(section, dict):
            for key in keys:
                _check_number(errors, section, key, *_SPACING_RANGE,
                              label=f"Интервал «{section_name}.{key}»")
    indent = text.get("first_line_indent_cm")
    if indent is not None and (isinstance(indent, bool)
                               or not isinstance(indent, (int, float))):
        errors.append("Абзацный отступ (text.first_line_indent_cm) должен быть числом")
    return errors


_ALIGN_RU = {"left": "по левому краю", "center": "по центру",
             "right": "по правому краю", "justify": "по ширине"}
_POSITION_RU = {"above": "над", "below": "под"}


def _num(value: Any) -> str:
    """Число по-русски: 1.5 → «1,5», целые без дробной части."""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    if isinstance(value, float):
        return str(value).replace(".", ",")
    return str(value)


def _flag(rules: dict, section: str, key: str, default: bool = False) -> bool:
    part = rules.get(section)
    return bool(part.get(key, default)) if isinstance(part, dict) else default


def describe_rules(rules: dict[str, Any]) -> str:
    """Краткая человекочитаемая сводка правил (кнопка «Предпросмотр» на вкладке «Правила»)."""
    rules = merge_with_defaults(rules if isinstance(rules, dict) else {})
    page, text = rules["page"], rules["text"]
    headings, figures, tables = rules["headings"], rules["figures"], rules["tables"]
    footnotes, references = rules["footnotes"], rules["references"]
    numbering, volume, code = rules["numbering"], rules["volume"], rules["code_blocks"]
    orient = "книжная" if page.get("orientation") == "portrait" else "альбомная"
    lines = [
        f"Правила: {rules.get('name', 'без названия')} "
        f"[{'строгий режим (ГОСТ)' if rules.get('strict') else 'свободный режим'}]",
        "Страница: {size} {orient}, поля Л{l}/П{r}/В{t}/Н{b} мм, "
        "до колонтитула {footer} мм".format(
            size=page.get("size", "A4"), orient=orient,
            l=_num(page.get("margin_left_mm")), r=_num(page.get("margin_right_mm")),
            t=_num(page.get("margin_top_mm")), b=_num(page.get("margin_bottom_mm")),
            footer=_num(page.get("footer_distance_mm"))),
        "Основной текст: {font} {size} пт, интервал {ls}, отступ {ind} см, "
        "{align}, до/после {sb}/{sa} пт".format(
            font=text.get("font"), size=_num(text.get("size_pt")),
            ls=_num(text.get("line_spacing")), ind=_num(text.get("first_line_indent_cm")),
            align=_ALIGN_RU.get(text.get("alignment"), text.get("alignment")),
            sb=_num(text.get("space_before_pt")), sa=_num(text.get("space_after_pt"))),
        "Заголовки: структурные {case} {weight} {align}{new_page}, без точки; "
        "разделы «{sec}», подразделы «{sub}»; keepWithNext: {keep}; "
        "кегль не больше основного: {same}".format(
            case="ПРОПИСНЫЕ" if _flag(rules, "headings", "structural_uppercase") else "как в тексте",
            weight="полужирные" if _flag(rules, "headings", "structural_bold") else "обычные",
            align=_ALIGN_RU.get(headings.get("structural_align"), "по центру"),
            new_page=", с новой страницы" if _flag(rules, "headings", "structural_new_page") else "",
            sec=headings.get("section_numbering"), sub=headings.get("subsection_numbering"),
            keep="да" if _flag(rules, "headings", "keep_with_next") else "нет",
            same="да" if _flag(rules, "headings", "same_size_as_body") else "нет"),
        "Рисунки: подпись «{tpl}» {pos} {align}, нумерация {num}, "
        "макс. {w}×{h} см, ссылка в тексте: {ref}".format(
            tpl=figures.get("caption_template"),
            pos=_POSITION_RU.get(figures.get("caption_position"), "под"),
            align=_ALIGN_RU.get(figures.get("caption_align"), "по центру"),
            num=figures.get("numbering"), w=_num(figures.get("max_width_cm")),
            h=_num(figures.get("max_height_cm")),
            ref="обязательна" if _flag(rules, "figures", "require_reference_in_text") else "не требуется"),
        "Таблицы: название «{tpl}» {pos} {align}, «{cont}» {cont_align}, "
        "шапка: {repeat}{bold}, кегль {size} пт, интервал {ls}, границы: {borders}".format(
            tpl=tables.get("caption_template"),
            pos=_POSITION_RU.get(tables.get("caption_position"), "над"),
            align=_ALIGN_RU.get(tables.get("caption_align"), "слева"),
            cont=tables.get("continuation_text"),
            cont_align=_ALIGN_RU.get(tables.get("continuation_align"), "справа"),
            repeat="повтор, " if _flag(rules, "tables", "repeat_header") else "",
            bold="полужирная" if _flag(rules, "tables", "header_bold") else "обычная",
            size=_num(tables.get("font_size_pt")), ls=_num(tables.get("line_spacing")),
            borders=tables.get("borders")),
        "Формулы: отдельной строкой {align}, номер {nf} у правого края, "
        "пояснение «{where}» без двоеточия".format(
            align=_ALIGN_RU.get(rules["formulas"].get("align"), "по центру"),
            nf=rules["formulas"].get("number_format"),
            where=rules["formulas"].get("explanation_first_word")),
        "Примечания: «{prefix} – …»{indent}{multi}".format(
            prefix=rules["notes"].get("prefix"),
            indent=" с абзацного отступа" if _flag(rules, "notes", "from_indent") else "",
            multi=", при нескольких — нумерация" if _flag(rules, "notes", "numbered_when_multiple") else ""),
        "Сноски: {font} {size} пт (мин. {min} пт), знак надстрочно без пробела".format(
            font=footnotes.get("font"), size=_num(footnotes.get("size_pt")),
            min=_num(footnotes.get("min_allowed_size_pt"))),
        "Источники: сноски {style}, порядок «{order}», перечисление «{num}», "
        "возраст ≤ {age} лет, процитировано ≥ {cite} %".format(
            style=references.get("style"), order=references.get("bibliography_order"),
            num=references.get("bibliography_numbering"),
            age=_num(references.get("max_source_age_years")),
            cite=_num(references.get("min_cited_percent"))),
        "Приложения: «{label}»{new_page}, буквы без {excl}".format(
            label=rules["appendices"].get("label"),
            new_page=" с новой страницы" if _flag(rules, "appendices", "new_page") else "",
            excl=", ".join(rules["appendices"].get("letters_exclude", []))),
        "Содержание: {req}{auto}".format(
            req="обязательно" if _flag(rules, "toc", "required") else "не обязательно",
            auto=", автосодержание" if _flag(rules, "toc", "auto") else ""),
        "Нумерация страниц: {pos}, титул {title}, вторая страница показывает «{second}», "
        "кегль {size} пт".format(
            pos=numbering.get("position"),
            title="нумеруется" if _flag(rules, "numbering", "title_page_numbered") else "без номера",
            second=_num(numbering.get("second_page_shows")),
            size=_num(numbering.get("font_size_pt"))),
        "Объём: от «{frm}» до «{to}», заполненность ≥ {fill} % "
        "(последняя в разделе ≥ {last} %)".format(
            frm=volume.get("count_from"), to=volume.get("count_to"),
            fill=_num(volume.get("min_page_fill_percent")),
            last=_num(volume.get("min_last_page_fill_percent"))),
        "Блоки кода: {font} {size} пт, интервал {ls}, отступ: {ind}".format(
            font=code.get("font"), size=_num(code.get("size_pt")),
            ls=_num(code.get("line_spacing")),
            ind="да" if _flag(rules, "code_blocks", "indent") else "нет"),
    ]
    return "\n".join(lines)
