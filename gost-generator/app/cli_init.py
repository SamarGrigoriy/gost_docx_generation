"""Текстовые и схемные артефакты для `gost init --agents` и `gost explain` (п. 20).

Содержит: шаблон AGENTS.md, компактные fallback-копии JSON-схем (draft 2020-12,
без внешних $ref), примеры JSON и словарь объяснений кодов нарушений §9.1 и
машинных кодов ошибок CLI. Отдельный модуль — чтобы app/cli.py остался компактным.
"""

from __future__ import annotations

import shutil
from pathlib import Path
from typing import Any

from app import cli_core, constants

SCHEMA_NAMES: tuple[str, ...] = ("rules", "profile", "script-meta", "report", "job")

# --- AGENTS.md (п. 20, п. 19) ------------------------------------------------
# Структура воспроизводит эталон: первый шаг, золотые правила, рецепты,
# таблица пресетов, коды возврата, запрещено, тесты.
AGENTS_MD = """# AGENTS.md — правила работы ИИ-агентов в этом проекте

## Что это
«ГОСТ-Генератор»: GUI + CLI для создания и проверки учебных `.docx` по ГОСТ 7.32 /
методичке «Мультиворк». Одно ядро `app/core`, две оболочки: графическая (`app/ui`)
и headless-CLI (`app/cli`). CLI работает без дисплея и без GUI-зависимостей.

## Первый шаг всегда
```bash
gost context --json      # пресет, скрипты, профили, каталоги, версии
gost schema job          # если будешь формировать задачу
```
Не угадывай параметры — читай `context` и `schema`.

## Золотые правила
1. **Никогда не правь OOXML руками** для исправления оформления. Используй `gost fix`.
2. Все вызовы — с `--json`. **stdout = JSON-конверт, stderr = логи.** Парси только stdout.
3. Проверяй **код возврата**, а не текст: 0 — ок, 1 — есть нарушения, 2 — аргументы,
   3 — файл не найден, 4 — скрипт упал, 5 — таймаут, 6 — конфиг, 7 — нормализация.
4. Перед записью — `--dry-run`, если не уверен в результате. `fix` идемпотентен.
5. Не изобретай содержимое документа. Генерируй только из предоставленного материала.
6. Настройки берутся по приоритету: флаги > GOST_RULES/GOST_PROFILE/GOST_HOME/GOST_OUTPUT >
   gost.json в каталоге проекта > состояние приложения > встроенные дефолты.

## Конверт JSON (п. 19.4)
```json
{
  "schema_version": 1,
  "ok": false,
  "command": "check",
  "data": {},
  "errors": [{"code": "FILE_NOT_FOUND", "message": "...", "path": "..."}],
  "warnings": []
}
```

## Рецепты
```bash
# генерация + нормализация + проверка одной командой
gost job run task.json --json

# то же вручную
gost run mkdocx --input work.md --output out.docx \\
     --rules gost_732 --profile default --normalize --check --json

# проверка
gost check out.docx --rules gost_732 --format json --fail-on error

# автоисправление оформления (идемпотентно, делает бэкап <name>.orig.docx)
gost fix out.docx --rules gost_732 --json

# пакетно
gost batch ./output --glob '*.docx' --check --report summary.md --json

# план без записи файлов
gost run mkdocx --input work.md --output out.docx --dry-run --json

# что за нарушение и как исправить
gost explain TBL_CAPTION
```

## Выбор пресета
| Пресет | Для чего |
|---|---|
| `gost_732` | студенческие работы: реферат, отчёт, лабораторная, курсовая, ВКР |
| `multivork` | то же, базовый вариант методички |
| `pedagogical` | тех.карта, конспект, подборка игр: альбомная ориентация, таблица 12 pt, интервал 1,0 |
| `free` | без жёстких требований, проверки как рекомендации |

Актуальный список: `gost rules list --json`. Не применяй ГОСТ-пресет к педагогическим
документам — получишь ложные нарушения.

## Добавить новый скрипт-генератор
Положи `.py` в каталог `scripts/` и объяви в нём:
```python
SCRIPT_META = {
    "name": "Название", "description": "...", "category": "...",
    "input": "none",            # none | file_md | file_docx | text | folder
    "output_ext": ".docx",
    "arg_template": ["{output}"],
    "rules_hook": "postprocess", # native | constants | postprocess | none
    "params": [{"id": "theme", "label": "Тема", "type": "text", "default": ""}]
}
```
Затем `gost scripts validate <name>` и `gost scripts describe <name> --json`.
Скрипт: только stdlib, без сети, без интерактивного ввода, без запуска оболочек.

## Запрещено
- новые runtime-зависимости (разрешены только GUI-пакет оформления и pillow);
- импорт GUI (`app/ui`) и сторонних GUI-библиотек из `app/core` и `app/cli`;
- docx-библиотеки, lxml, БД, сетевые запросы;
- правка встроенных пресетов (`builtin: true`) — создавай пользовательский;
- запуск через shell-оболочки, динамическое выполнение произвольного кода;
- массовое переформатирование чужого кода.

## Тесты
```bash
python -m unittest discover -s tests -v
```
Новая логика в `core` — обязательно с тестом. Фикстуры `.docx` генерируй программно.
Проверяй идемпотентность нормализатора и валидность JSON по `schemas/`.

## Коммиты
`feat(core): ...`, `fix(cli): ...`, `test: ...`, `docs: ...`.
В описании — что изменилось и как проверено. Решение, не описанное в ТЗ, фиксируй в `DECISIONS.md`.
"""

# --- Компактные fallback-схемы (draft 2020-12, только stdlib-валидаторы) -----

_RULES_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "$id": "gost/schema/rules",
    "title": "Правила оформления (пресет)",
    "type": "object",
    "required": ["page", "text"],
    "properties": {
        "schema_version": {"type": "integer"},
        "name": {"type": "string"},
        "builtin": {"type": "boolean"},
        "strict": {"type": "boolean"},
        "page": {
            "type": "object",
            "required": ["size", "orientation"],
            "properties": {
                "size": {"type": "string"},
                "orientation": {"enum": list(constants.ORIENTATIONS)},
                "margin_top_mm": {"type": "number"},
                "margin_right_mm": {"type": "number"},
                "margin_bottom_mm": {"type": "number"},
                "margin_left_mm": {"type": "number"},
                "footer_distance_mm": {"type": "number"},
            },
        },
        "text": {
            "type": "object",
            "properties": {
                "font": {"type": "string"},
                "size_pt": {"type": "number"},
                "line_spacing": {"type": "number"},
                "first_line_indent_cm": {"type": "number"},
                "alignment": {"enum": list(constants.ALIGNMENTS)},
                "space_before_pt": {"type": "number"},
                "space_after_pt": {"type": "number"},
            },
        },
        "headings": {"type": "object"},
        "figures": {"type": "object"},
        "tables": {"type": "object"},
        "formulas": {"type": "object"},
        "notes": {"type": "object"},
        "footnotes": {"type": "object"},
        "references": {"type": "object"},
        "appendices": {"type": "object"},
        "toc": {"type": "object"},
        "numbering": {"type": "object"},
        "volume": {"type": "object"},
        "code_blocks": {"type": "object"},
    },
}

_PROFILE_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "$id": "gost/schema/profile",
    "title": "Профиль автора/работы",
    "type": "object",
    "required": ["name"],
    "properties": {
        "schema_version": {"type": "integer"},
        "name": {"type": "string"},
        "person": {
            "type": "object",
            "properties": {
                "fio": {"type": "string"},
                "gruppa": {"type": "string"},
                "prep": {"type": "string"},
                "status": {"enum": list(constants.STATUS_CHOICES)},
            },
        },
        "work": {
            "type": "object",
            "properties": {
                "work_type": {"type": "string"},
                "pz_no": {"type": "string"},
                "theme": {"type": "string"},
                "disc": {"type": "string"},
                "city": {"type": "string"},
                "year": {"type": "integer"},
            },
        },
        "institution": {
            "type": "object",
            "properties": {
                "ministry": {"type": "string"},
                "fsbe": {"type": "string"},
                "univ": {"type": "string"},
                "kaf": {"type": "string"},
            },
        },
        "extra": {"type": "object", "additionalProperties": {"type": "string"}},
    },
}

_SCRIPT_META_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "$id": "gost/schema/script-meta",
    "title": "SCRIPT_META скрипта-плагина",
    "type": "object",
    "required": ["name"],
    "properties": {
        "name": {"type": "string"},
        "description": {"type": "string"},
        "category": {"type": "string"},
        "version": {"type": "string"},
        "input": {"enum": list(constants.INPUT_TYPES)},
        "output_ext": {"type": "string"},
        "arg_template": {"type": "array", "items": {"type": "string"}},
        "rules_hook": {"enum": list(constants.RULES_HOOKS)},
        "params": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["id"],
                "properties": {
                    "id": {"type": "string"},
                    "label": {"type": "string"},
                    "type": {"enum": list(constants.PARAM_TYPES)},
                    "default": {},
                    "required": {"type": "boolean"},
                    "options": {"type": "array"},
                },
            },
        },
    },
}

_REPORT_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "$id": "gost/schema/report",
    "title": "Результат проверки (data.check)",
    "type": "object",
    "required": ["file", "summary", "issues"],
    "properties": {
        "file": {"type": "string"},
        "preset": {"type": "string"},
        "strict": {"type": "boolean"},
        "read_error": {"type": ["string", "null"]},
        "summary": {
            "type": "object",
            "required": ["error", "warning", "info"],
            "properties": {
                "error": {"type": "integer"},
                "warning": {"type": "integer"},
                "info": {"type": "integer"},
            },
        },
        "issues": {
            "type": "array",
            "items": {
                "type": "object",
                "required": ["code", "severity", "message"],
                "properties": {
                    "code": {"type": "string"},
                    "severity": {"enum": list(constants.SEVERITIES)},
                    "message": {"type": "string"},
                    "location": {
                        "type": "object",
                        "properties": {
                            "paragraph": {"type": ["integer", "null"]},
                            "page": {"type": ["integer", "null"]},
                            "snippet": {"type": "string"},
                        },
                    },
                    "hint": {"type": "string"},
                },
            },
        },
    },
}

_JOB_SCHEMA: dict[str, Any] = {
    "$schema": "https://json-schema.org/draft/2020-12/schema",
    "$id": "gost/schema/job",
    "title": "Декларативная задача (task.json, п. 19.7)",
    "type": "object",
    "required": ["script"],
    "properties": {
        "schema_version": {"type": "integer"},
        "script": {"type": "string"},
        "input": {"type": "string"},
        "output": {"type": "string"},
        "params": {"type": "object"},
        "rules": {"type": "string"},
        "profile": {"type": "string"},
        "meta": {"type": "object"},
        "post": {
            "type": "object",
            "properties": {
                "normalize": {"type": "boolean"},
                "check": {"type": "boolean"},
                "fail_on": {"enum": ["error", "warning", "never"]},
            },
        },
    },
}

FALLBACK_SCHEMAS: dict[str, dict[str, Any]] = {
    "rules": _RULES_SCHEMA,
    "profile": _PROFILE_SCHEMA,
    "script-meta": _SCRIPT_META_SCHEMA,
    "report": _REPORT_SCHEMA,
    "job": _JOB_SCHEMA,
}

# --- Примеры (examples/, п. 20.1 п.5) ----------------------------------------

EXAMPLES: dict[str, str] = {
    "task.json": """{
  "schema_version": 1,
  "script": "make_karta.py",
  "params": { "theme": "Моя семья", "age_group": "3-4", "duration_min": 15 },
  "rules": "pedagogical",
  "profile": "default",
  "output": "output/karta.docx",
  "post": { "normalize": true, "check": true, "fail_on": "error" }
}
""",
    "run_params.json": """{
  "pz_no": "1",
  "theme": "Пример темы",
  "fio": "Иванов Иван Иванович",
  "gruppa": "ИП-21-1",
  "prep": "Петров П. П."
}
""",
    "profile_custom.json": """{
  "schema_version": 1,
  "name": "Профиль студента",
  "person": { "fio": "Иванов Иван Иванович", "gruppa": "ИП-21-1",
              "prep": "Петров П. П.", "status": "студент" },
  "work": { "work_type": "Реферат", "pz_no": "", "theme": "Пример темы",
            "disc": "Информатика", "city": "Геническ", "year": 2026 },
  "institution": {
    "ministry": "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ",
    "fsbe": "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ",
    "univ": "", "kaf": ""
  },
  "extra": {}
}
""",
    "rules_custom.json": """{
  "schema_version": 1,
  "name": "Мой пресет",
  "strict": false,
  "page": { "size": "A4", "orientation": "portrait",
            "margin_left_mm": 30, "margin_right_mm": 15,
            "margin_top_mm": 20, "margin_bottom_mm": 20 },
  "text": { "font": "Times New Roman", "size_pt": 14, "line_spacing": 1.5,
            "first_line_indent_cm": 1.25, "alignment": "justify" }
}
""",
}

# --- Словарь объяснений (gost explain, п. 9.1 + коды ошибок CLI) -------------

EXPLAIN: dict[str, dict[str, str]] = {
    "PAGE_MARGINS": {
        "title": "Поля страницы не совпадают с правилами",
        "detail": "sectPr/docGrid pgMar отличается от пресета более чем на ±0,5 мм.",
        "fix": "gost fix <файл> --rules <пресет> — проставит поля из секции документа.",
        "source": "ГОСТ 7.32 п. 5.5; методичка «Мультиворк», раздел «Страницы»",
    },
    "PAGE_SIZE": {
        "title": "Формат или ориентация страницы не по правилу",
        "detail": "pgSz в sectPr не соответствует rules.page.size/orientation.",
        "fix": "gost fix; для педагогических работ используй пресет pedagogical (альбомная).",
        "source": "ГОСТ 7.32 п. 5.4",
    },
    "FONT_MAIN": {
        "title": "Основной шрифт не по правилу",
        "detail": "Реальный rFonts преобладающего текста вне таблиц и кода ≠ rules.text.font.",
        "fix": "gost fix перепишет docDefaults/Normal и конфликтующее прямое форматирование.",
        "source": "ГОСТ 7.32 п. 5.6",
    },
    "SIZE_MAIN": {
        "title": "Кегль основного текста не по правилу",
        "detail": "Основной размер шрифта (вне таблиц/кода) ≠ rules.text.size_pt.",
        "fix": "gost fix; таблицы и код-блоки проверяются отдельными правилами.",
        "source": "ГОСТ 7.32 п. 5.6",
    },
    "LINE_SPACING": {
        "title": "Межстрочный интервал не по правилу",
        "detail": "Основной интервал ≠ rules.text.line_spacing (обычно 1,5).",
        "fix": "gost fix установит spacing в стиле Normal.",
        "source": "ГОСТ 7.32 п. 5.7",
    },
    "INDENT_VALUE": {
        "title": "Отступ первой строки не равен правилу",
        "detail": "firstLine ≠ rules (±5 twips); важно значение (обычно 1,25 см ≈ 709 twips), а не факт наличия.",
        "fix": "gost fix проставит firstLine в абзацах основного текста.",
        "source": "ГОСТ 7.32 п. 5.7",
    },
    "ALIGNMENT": {
        "title": "Выравнивание основного текста",
        "detail": "Основной текст не по ширине (justify), если правило требует этого.",
        "fix": "gost fix; в нестрогом режиме — рекомендация.",
        "source": "ГОСТ 7.32 п. 5.7",
    },
    "SPACING_PARA": {
        "title": "Интервалы до/после абзаца ≠ 0",
        "detail": "w:spacing before/after должен быть 0; разбиение пустыми строками запрещено.",
        "fix": "gost fix обнулит before/after в стиле Normal и прямых абзацах.",
        "source": "ГОСТ 7.32 п. 5.7",
    },
    "NUM_POSITION": {
        "title": "Номер страницы не внизу по центру",
        "detail": "В колонтитуле нет поля PAGE внизу по центру (rules.numbering.position).",
        "fix": "gost fix добавит footer с полем PAGE, titlePg для титула.",
        "source": "ГОСТ 7.32 п. 5.5",
    },
    "NUM_START": {
        "title": "Нумерация страниц начинается некорректно",
        "detail": "Титульный лист не нумеруется, но вторая физическая страница должна показывать «2».",
        "fix": "gost fix: одна секция + titlePg (пустой footer первой страницы) или pgNumType start=2.",
        "source": "ГОСТ 7.32 п. 5.5",
    },
    "HEAD_STRUCT": {
        "title": "Структурный элемент оформлен неверно",
        "detail": "РЕФЕРАТ, СОДЕРЖАНИЕ, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ, СПИСОК…, ПРИЛОЖЕНИЕ — прописные, "
                  "полужирные, по центру, с новой страницы, без точки на конце.",
        "fix": "gost fix применит формат к заголовкам из STRUCTURAL_HEADINGS.",
        "source": "ГОСТ 7.32 пп. 5.2, 7.1; методичка §5.2",
    },
    "HEAD_KEEP": {
        "title": "Заголовок может остаться в конце страницы",
        "detail": "У стиля/абзаца заголовка нет keepNext — заголовок отрывается от текста.",
        "fix": "gost fix проставит w:keepNext заголовкам.",
        "source": "ГОСТ 7.32 п. 5.8",
    },
    "HEAD_SIZE": {
        "title": "Кегль заголовка больше основного",
        "detail": "По правилу заголовки не должны превышать rules.text.size_pt.",
        "fix": "gost fix приведует кегль Heading1..4 к правилу.",
        "source": "Методичка «Мультиворк», раздел «Заголовки»",
    },
    "FIG_CAPTION": {
        "title": "Подпись рисунка не по формату",
        "detail": "Нужно «Рисунок N – Название» под изображением, по центру, без отступа и точки.",
        "fix": "Переименуй подпись по шаблону rules.figures.caption_template; gost fix поправит форматирование.",
        "source": "ГОСТ 7.32 п. 5.10",
    },
    "FIG_COUNT": {
        "title": "Изображений больше, чем подписей",
        "detail": "У каждого рисунка должна быть подпись «Рисунок N – …».",
        "fix": "Добавь подписи; декоративные изображения убери или оформи как рисунок.",
        "source": "ГОСТ 7.32 п. 5.10",
    },
    "FIG_ORDER": {
        "title": "Нумерация рисунков со сбоями",
        "detail": "Сквозная нумерация без пропусков и повторов.",
        "fix": "Перенумеруй подписи по порядку появления.",
        "source": "ГОСТ 7.32 п. 5.10",
    },
    "FIG_REF": {
        "title": "Нет ссылки на рисунок в тексте",
        "detail": "Каждый рисунок должен быть упомянут в тексте («… на рисунке 1»/«см. рисунок 1»).",
        "fix": "Добавь ссылку по шаблону rules.figures.reference_template.",
        "source": "ГОСТ 7.32 п. 5.10",
    },
    "TBL_CAPTION": {
        "title": "Название таблицы не над ней или не по формату",
        "detail": "Нужно «Таблица N – Название» над таблицей, слева, без отступа, без точки.",
        "fix": "Оформи подпись по rules.tables.caption_template; gost fix выровняет.",
        "source": "ГОСТ 7.32 п. 5.11",
    },
    "TBL_HEADER": {
        "title": "Шапка таблицы не повторяется/не полужирная",
        "detail": "Первая строка должна иметь tblHeader (повтор на каждой странице) и полужирное начертание.",
        "fix": "gost fix проставит trPr/tblHeader и bold в первой строке.",
        "source": "ГОСТ 7.32 п. 5.11",
    },
    "TBL_BORDERS": {
        "title": "Границы таблицы не тонкие сплошные",
        "detail": "rules.tables.borders = single_thin; заливки и жирные рамки недопустимы.",
        "fix": "gost fix перепишет tblBorders.",
        "source": "Методичка «Мультиворк», раздел «Таблицы»",
    },
    "TBL_ORDER": {
        "title": "Нумерация таблиц не по порядку появления",
        "detail": "Таблицы нумеруются сквозной нумерацией 1, 2, …",
        "fix": "Исправь номера подписей.",
        "source": "ГОСТ 7.32 п. 5.11",
    },
    "TBL_REF": {
        "title": "Нет ссылки на таблицу в тексте",
        "detail": "Каждая таблица должна быть упомянута («в таблице N» / «табл. N»).",
        "fix": "Добавь ссылку в текст перед/после таблицы.",
        "source": "ГОСТ 7.32 п. 5.11",
    },
    "FORMULA": {
        "title": "Оформление формул",
        "detail": "Номер формулы (N) — у правого края; после формулы слово «где» без двоеточия.",
        "fix": "Перенумеруй формулы, убери двоеточие после «где».",
        "source": "ГОСТ 7.32 п. 5.9",
    },
    "FOOTNOTE_FONT": {
        "title": "Шрифт/кегль сносок не по правилу",
        "detail": "Сноски — тем же шрифтом, кегль меньше основного (обычно 10–12 pt).",
        "fix": "Поправь стиль FootnoteText; gost fix затронет стили.",
        "source": "ГОСТ 7.32 п. 5.13",
    },
    "TOC_PRESENT": {
        "title": "Нет автоматического содержания",
        "detail": "Обязательно поле TOC (автосодержание), а не вручную набранный список.",
        "fix": "gost fix назначит заголовкам pStyle HeadingN; в Word: Ссылки → Оглавление. "
              "mkdocx генерирует поле TOC сам.",
        "source": "ГОСТ 7.32 п. 4.2",
    },
    "BIB_ORDER": {
        "title": "Список источников не в порядке упоминания",
        "detail": "rules.references.bibliography_order = by_appearance (или по алфавиту — смотри пресет).",
        "fix": "Переставь записи и номера; обнови ссылки в тексте.",
        "source": "ГОСТ 7.32 п. 5.12",
    },
    "BIB_AGE": {
        "title": "Устаревшие источники",
        "detail": "Год издания ≥ текущий − rules.references.max_source_age_years (обычно 5 лет).",
        "fix": "Замени устаревшие издания более свежими или обоснуй в тексте.",
        "source": "Методичка «Мультиворк», раздел «Источники»",
    },
    "BIB_COVERAGE": {
        "title": "Мало процитированных источников",
        "detail": "Прецитировано ≥ rules.references.min_cited_percent % записей списка.",
        "fix": "Добавь ссылки [N] в текст или убери неиспользуемые записи.",
        "source": "Методичка «Мультиворк», раздел «Источники»",
    },
    "APPENDIX": {
        "title": "Оформление приложений",
        "detail": "ПРИЛОЖЕНИЕ с новой страницы, буквы А, Б, В… (без Ё, З, Й, О, Ч, Ъ, Ы, Ь).",
        "fix": "Перемести приложения на новые страницы, исправь буквы.",
        "source": "ГОСТ 7.32 п. 2.14, 5.14",
    },
    "VOLUME_FILL": {
        "title": "Заполненность страниц",
        "detail": "Эвристика: страницы должны быть заполнены ≥ правила (обычно 80 %).",
        "fix": "Распредели текст; не раздувай документ пустыми страницами.",
        "source": "Методичка «Мультиворк», раздел «Объём»",
    },
    "MANUAL": {
        "title": "Проверить вручную",
        "detail": "Орфография (F7), титульный лист по шаблону вуза, сканы вместо графики, "
                  "оригинальность, выводы по главам.",
        "fix": "Пройди чек-лист блока «Проверить вручную» в отчёте.",
        "source": "Методичка «Мультиворк», приёмочный чек-лист",
    },
    # --- Машинные коды ошибок CLI ---
    "USAGE": {
        "title": "Ошибка использования (код возврата 2)",
        "detail": "Неверные аргументы подкоманды; см. gost <command> --help.",
        "fix": "Проверь имя команды и флаги; список команд: gost --help.",
        "source": "ТЗ п. 19.3, 19.5",
    },
    "FILE_NOT_FOUND": {
        "title": "Файл или каталог не найден (код возврата 3)",
        "detail": "Указанный путь не существует.",
        "fix": "Проверь путь (абсолютный или относительно cwd); кавычки для пробелов и кириллицы.",
        "source": "ТЗ п. 19.5",
    },
    "SCRIPT_NOT_FOUND": {
        "title": "Скрипт-плагин не найден (код возврата 3)",
        "detail": "Имя не совпадает ни с одним скриптом в каталогах scripts/.",
        "fix": "gost scripts list --json — возьми точное имя.",
        "source": "ТЗ п. 19.3",
    },
    "RULES_NOT_FOUND": {
        "title": "Пресет не найден (код возврата 3)",
        "detail": "Нет пресета с таким именем и нет файла по указанному пути.",
        "fix": "gost rules list --json; имя ключа = имя файла без .json.",
        "source": "ТЗ п. 19.3",
    },
    "PROFILE_NOT_FOUND": {
        "title": "Профиль не найден (код возврата 3)",
        "detail": "Нет профиля с таким именем/путём.",
        "fix": "gost profile list --json.",
        "source": "ТЗ п. 19.3",
    },
    "SCRIPT_FAILED": {
        "title": "Скрипт завершился с ошибкой (код возврата 4)",
        "detail": "Дочерний процесс вернул ненулевой код; stderr скрипта — в stderr CLI.",
        "fix": "Смотри вывод запуска; проверь аргументы: gost run SCRIPT --dry-run --json.",
        "source": "ТЗ п. 19.5",
    },
    "TIMEOUT": {
        "title": "Таймаут запуска (код возврата 5)",
        "detail": "Скрипт не завершился за отведённое время и был завершён (terminate → kill).",
        "fix": "Увеличь таймаут в gost.json (script_timeout_s) или упрости задачу.",
        "source": "ТЗ п. 12, 19.5",
    },
    "BAD_CONFIG": {
        "title": "Невалидный конфиг/пресет/профиль (код возврата 6)",
        "detail": "JSON конфигурации повреждён или не проходит схему.",
        "fix": "gost rules validate / gost doctor --json; при битом state.json — gost init.",
        "source": "ТЗ п. 12, 19.5",
    },
    "INVALID_JSON": {
        "title": "Не удалось прочитать JSON (код возврата 6)",
        "detail": "Файл или параметр некорректен: ошибка кодировки (нужен UTF-8) или синтаксиса JSON.",
        "fix": "Проверь файл валидатором; передавай пути с расширением .json.",
        "source": "ТЗ п. 12",
    },
    "JOB_INVALID": {
        "title": "Задача job не проходит схему (код возврата 6)",
        "detail": "task.json не соответствует схеме: gost schema job.",
        "fix": "Сравни с examples/task.json; заполни обязательные поля (script).",
        "source": "ТЗ п. 19.7",
    },
    "NORMALIZE_FAILED": {
        "title": "Нормализация не удалась (код возврата 7)",
        "detail": "Не удалось применить правила к .docx (битый zip/XML, нет прав на запись).",
        "fix": "Проверь, что файл открывается; используй --backup; исходник цел.",
        "source": "ТЗ п. 9, 19.5",
    },
    "VIOLATIONS": {
        "title": "Найдены нарушения (код возврата 1)",
        "detail": "Проверка вернула замечания не ниже порога --fail-on.",
        "fix": "gost explain <код> по каждому нарушению; gost fix <файл> для автоисправления.",
        "source": "ТЗ п. 19.5",
    },
    "CORE_NOT_AVAILABLE": {
        "title": "Модуль ядра недоступен (код возврата 6)",
        "detail": "app.core.* не импортируется: неполная установка или ошибка сборки.",
        "fix": "gost doctor --json покажет, какие модули отсутствуют; восстанови установку.",
        "source": "ТЗ п. 19.1",
    },
    "INTERNAL_ERROR": {
        "title": "Внутренняя ошибка CLI (код возврата 2)",
        "detail": "Необрабатываемая ситуация; traceback — в stderr.",
        "fix": "Сообщи разработчикам; приложи вывод gost doctor --json.",
        "source": "ТЗ п. 12",
    },
}


def explain_entry(code: str) -> dict[str, str]:
    """Найти запись словаря объяснений по коду (без регистра — негарантированно)."""
    entry = EXPLAIN.get(code)
    if entry is None:
        raise cli_core.CliError(
            "NOT_FOUND",
            f"Код {code!r} не найден в словаре объяснений. "
            f"См. gost context --json или schemas/report.json.",
        )
    return dict(entry)


# --- Генерация артефактов (gost init --agents) --------------------------------

_SKELETON_DIRS: tuple[str, ...] = (
    constants.DIR_SCRIPTS, constants.DIR_PRESETS, constants.DIR_PROFILES,
    constants.DIR_OUTPUT, constants.DIR_LOGS,
)


def ensure_skeleton(base: Path) -> list[str]:
    """Создать каталоги scripts/ presets/ profiles/ output/ logs/ в base; вернуть список созданных."""
    created: list[str] = []
    for name in _SKELETON_DIRS:
        d = base / name
        if not d.exists():
            d.mkdir(parents=True, exist_ok=True)
            created.append(str(d))
    return created


def _same_file(src: Path, dest: Path) -> bool:
    try:
        return src.exists() and dest.exists() and src.resolve() == dest.resolve()
    except OSError:
        return False


def _copy_or_write(dest: Path, src: Path | None, payload: str) -> str:
    """Скопировать готовый файл с диска или записать встроенный текст."""
    dest.parent.mkdir(parents=True, exist_ok=True)
    if src is None or (not src.exists() and payload):
        dest.write_text(payload, encoding="utf-8")
    elif not _same_file(src, dest):
        # `gost init` внутри репозитория: источник и приёмник — один файл, копировать нечего.
        shutil.copyfile(src, dest)
    return str(dest)


def write_agents_artifacts(base: Path) -> list[str]:
    """Записать AGENTS.md, schemas/*.json, examples/*.json (+task.json) в base.

    Настоящие схемы и примеры читаются с диска (репозиторий/установка), при их
    отсутствии используются компактные встроенные копии этого модуля.
    Возвращает список записанных путей.
    """
    root = cli_core.repo_root()
    written: list[str] = [
        # с диска — единый источник истины (D12); встроенная копия только для собранного exe.
        _copy_or_write(base / "AGENTS.md", root / "AGENTS.md", AGENTS_MD),
    ]
    on_disk = cli_core.list_schemas()
    for name in SCHEMA_NAMES:
        src = on_disk.get(name)
        payload = "" if src else _dump(FALLBACK_SCHEMAS[name])
        written.append(_copy_or_write(base / constants.DIR_SCHEMAS / f"{name}.json", src, payload))
    examples_src = root / constants.DIR_EXAMPLES
    for fname, payload in EXAMPLES.items():
        src: Path | None = None
        if examples_src.exists():
            cand = examples_src / fname
            if cand.exists():
                src = cand
        written.append(_copy_or_write(base / constants.DIR_EXAMPLES / fname, src, payload))
    # любые дополнительные примеры с диска (кроме уже записанных)
    if examples_src.exists():
        for cand in sorted(examples_src.glob("*.json")):
            dest = base / constants.DIR_EXAMPLES / cand.name
            if dest not in [Path(p) for p in written]:
                written.append(_copy_or_write(dest, cand, ""))
    return written


def _dump(data: Any) -> str:
    """Сериализовать схему в читаемый JSON."""
    import json as _json
    return _json.dumps(data, ensure_ascii=False, indent=2) + "\n"
