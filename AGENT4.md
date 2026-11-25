# AGENTS.md — правила работы ИИ-агентов в этом репозитории

## Что это
«ГОСТ-Генератор»: GUI + CLI для создания и проверки учебных `.docx` по ГОСТ 7.32 /
методичке «Мультиворк». Одно ядро `app/core`, две оболочки: `app/ui` (CustomTkinter) и `app/cli`.

## Первый шаг всегда
```bash
gost context --json      # пресет, скрипты, профили, каталоги, версии
gost schema job          # если будешь формировать задачу
```
Не угадывай параметры — читай `context` и `schema`.

## Золотые правила
1. **Никогда не правь OOXML руками** для исправления оформления. Используй `gost fix`.
2. Все вызовы — с `--json`. **stdout = JSON, stderr = логи.** Парси только stdout.
3. Проверяй **код возврата**, а не текст: 0 — ок, 1 — есть нарушения, 2 — аргументы,
   3 — файл не найден, 4 — скрипт упал, 5 — таймаут, 6 — конфиг, 7 — нормализация.
4. Перед записью — `--dry-run`, если не уверен в результате.
5. Не изобретай содержимое документа. Генерируй только из предоставленного материала.

## Рецепты
```bash
# генерация + нормализация + проверка одной командой
gost job run task.json --json

# то же вручную
gost run mkdocx --input work.md --output out.docx \
     --rules gost_732 --profile default --normalize --check --json

# проверка
gost check out.docx --rules gost_732 --format json --fail-on error

# автоисправление оформления (идемпотентно, делает бэкап)
gost fix out.docx --rules gost_732 --json

# пакетно
gost batch ./output --glob '*.docx' --check --report summary.md --json

# что за нарушение
gost explain TBL_CAPTION
```

## Выбор пресета
| Пресет | Для чего |
|---|---|
| `gost_732` | студенческие работы: реферат, отчёт, лабораторная, курсовая, ВКР |
| `multivork` | то же, базовый вариант методички |
| `pedagogical` | тех.карта, конспект, подборка игр: альбомная ориентация, таблица 12 pt, интервал 1,0 |
| `free` | без жёстких требований, проверки как рекомендации |

Не применяй ГОСТ-пресет к педагогическим документам — получишь ложные нарушения.

## Добавить новый скрипт-генератор
Положи `.py` в `scripts/` и объяви в нём:
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
Скрипт: только stdlib, без сети, без `input()`, без `shell=True`.

## Запрещено
- новые runtime-зависимости (разрешены только `customtkinter`, `pillow`);
- импорт `app/ui` или `customtkinter` из `app/core` и `app/cli`;
- `python-docx`, `lxml`, БД, сетевые запросы;
- правка встроенных пресетов (`builtin: true`) — создавай пользовательский;
- `shell=True`, `eval`, `exec` для пользовательского ввода;
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