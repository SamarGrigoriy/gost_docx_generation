# INTERFACES — внутренний контракт модулей (для согласованной сборки)

Корень проекта: `/home/gsamarin/docx_gen/gost-generator/`
Стиль: stdlib-only, аннотации типов на публичных функциях, docstrings на русском,
без TODO/заглушек, файлы > 600 строк разбивать. `app/core` и `app/cli.py` НЕ импортируют
`app.ui` / `customtkinter`.

Иерархия импортов: ui/cli -> core -> constants/units. Core не знает про Tk.

## app.constants, app.core.units — уже существуют (см. файлы)

## app.core.config (владелец: Agent A)

- `CorruptedConfigError(Exception)` c атрибутом `.path`
- `load_json(path, default=None, backup_on_corrupt=True) -> Any` — битый JSON: бэкап `<name>.corrupt`, возврат default
- `save_json(path, data) -> Path` — atomically-ish, UTF-8, ensure_ascii=False, indent=2, каталоги создаёт
- `migrate(data: dict, kind: str) -> dict` — доведение `schema_version` до `constants.SCHEMA_VERSION` с дозаполнением дефолтов
- `class DataDirs: __init__(self, override: Path|str|None=None)` — атрибуты `home, scripts, presets, profiles, output, logs, state_path, project_dir`; `ensure()`; `is_portable()`
  Порядок (п. 19.6): флаг `GOST_HOME` > `portable.flag` рядом с exe/repo > `./gost.json` > `~/.config/GostGenerator` (Linux) / `%APPDATA%\GostGenerator` (Windows)
- `load_state(dirs) -> dict` / `save_state(dirs, state)` / `default_state(dirs) -> dict` (п. 7.3)

## app.core.rules_store (Agent A)

- `class RulesStore:` `__init__(self, dirs, builtin_dir: Path|None = None)` (builtin_dir=репозитория `presets/`, внутри exe — `sys._MEIPASS/presets`)
  - `list() -> list[dict]` элементы `{"key","name","builtin","path"}` (ключ = имя файла без .json)
  - `load(key_or_path) -> dict` ; `save(rules, key=None) -> str` (builtin → `CannotEditBuiltinError` для записи под тем же ключом; сохраняй как пользовательский)
  - `save_as(rules, new_key) -> str` ; `delete(key)` (builtin → ошибка) ; `export(key, dest)` ; `import_from(src) -> key`
  - `resolve(name_or_path) -> dict` — имя пресета или путь к .json
- `CannotEditBuiltinError(Exception)`
- `DEFAULT_RULES: dict` — встроенная копия ГОСТ-пресета (fallback, если каталога нет)
- `merge_with_defaults(rules) -> dict`
- `validate_rules(rules) -> list[str]` — сообщения о проблемах, пустой список = ок
- `describe_rules(rules) -> str` — человекочитаемая сводка для кнопки «Предпросмотр»

## app.core.profile_store (Agent A)

- `class ProfileStore:` CRUD по той же схеме (`list/load/save/save_as/delete/export/import_from`, `use` — через state)
- `flat_meta(profile) -> dict` — плоский словарь: ключи `person`,`work`,`institution`,`extra` слиты + `"name"`; год по умолчанию = `datetime.date.today().year` (не хардкод)
- `default_profile() -> dict` (п. 7.2)
- `title_header_preview(profile) -> str` — текстовая шапка титульного листа

## app.core.plugin_loader (Agent B)

- `class PluginError(Exception)`
- `parse_meta(source: str) -> dict | None` — через `ast`, без исполнения
- `load_meta_for_path(path) -> dict | None` — SCRIPT_META или sidecar `<name>.meta.json` (sidecar приоритетнее)
- `class PluginInfo` (dataclass): `path, file_name, name, description, category, version, input, output_ext, arg_template, rules_hook, params, has_meta, error`
- `scan_dir(dir, extra_paths=()) -> list[PluginInfo]` — битый скрипт не ломает список (field `error`)
- `validate(path) -> tuple[bool, list[str]]`
- `build_argv(info, *, input_path, output_path, meta: dict, rules: dict, profile: dict, params: dict, tmp_dir: Path) -> list[str]`
  Подстановки `{input} {output} {meta_json} {rules_json} {profile_json} {param_id}`; JSON пишется в tmp_dir.
  Если `arg_template` пустой: `[output]` + (если input есть — input первым) + `json.dumps(params, ensure_ascii=False)` позиционным аргументом.
- `save_sidecar(info, params/…overrides) -> Path`

## app.core.runner (Agent B)

Интерфейсы раннера (п. 3.2, 8.1, 9):
- modes: `internal` (по умолчанию; `[sys.executable, "--internal-run", script, params_json_path]`,
  где params_json = `{"argv": [...], "cwd": ...}`), `external` (`[python_path, script, *argv]`)
- `class RunResult` (dataclass): `argv, exit_code, lines: list[str], duration_s, cancelled, timed_out, error`
  + `ok` property; `stdout_text` property
- `class ScriptRunner:` `__init__(self, timeout_s=120, mode="internal", python_path=None, on_line=None, kill_grace_s=3)`
  - `run_script(info: PluginInfo|dict, *, meta, rules, profile, input_path=None, output_path=None, params=None, cwd=None) -> RunResult`
  - `run_argv(argv: list[str], *, cwd=None) -> RunResult`
  - отмена: `cancel()` — terminate → через 3 с kill; строки читаются в поток и кладутся в `queue.Queue` и `on_line`
  - после успешной генерации пост-хуки вызывает сам модуль вызывающий (раннер только запускает), НО обязан поддерживать `postprocess`-семантику: если rules_hook == "postprocess" и включена нормализация — раннер НЕ вызывает нормализатор напрямую; ответственность за sequence — `run_and_finalize()`.
- `run_and_finalize(...) -> dict` — оркестрация: run → (normalize по rules_hook/флагу) → (check по флагу) → структурированный результат `{"run": RunResult, "normalized": {...}|None, "check": {"summary","issues"}|None, "output": path}` (задел под журнал v2)
- `multiprocessing.freeze_support()` — в `app/main.py`, не в раннере

## app.core.job (Agent B)

- `class JobError(Exception)`
- `validate_job(job: dict) -> list[str]` (схема §19.7)
- `execute_job(job: dict, *, dirs, runner_kwargs=None, dry_run=False) -> dict` — конверт `{"ok", "output", "run", "normalize", "check"}`; ключ `post: {normalize, check, fail_on}`

## app.core.docx_inspector (Agent C)

- `class DocxError(Exception)` — битый zip/XML
- `class Issue` (dataclass): `code, severity, message, location: dict, hint`
  (`location` = `{"paragraph": int|None, "page": int|None, "snippet": str}`)
- `check_docx(path, rules: dict) -> list[Issue]` — strict из `rules["strict"]`; при strict=False error→warning
- `check_docx_result(path, rules) -> dict` — `{"file","preset","strict","summary":{"error":n,"warning":n,"info":n},"issues":[...],"read_error":str|None}`
- Реализовать ВСЕ коды из §9.1 ТЗ (PAGE_MARGINS … MANUAL)
- `_MANUAL_ITEMS: list[tuple[str,str]]` — для блока «Проверить вручную» (код `MANUAL`, severity info)
- Страницы — эвристика: деление объёма текста на ёмкость полосы (из правил), поля `page` могут быть `None`

## app.core.docx_normalizer (Agent C)

- `normalize_docx(path, rules: dict, backup: bool = True) -> dict`
  Результат: `{"changed": bool, "changes": list[str], "backup": str|None}`
- Идемпотентность (повторный прогон → идентичный XML); бэкап `<name>.orig.docx` (один раз, не перезаписывать при повторных запусках, если уже есть)
- Не трогать: Consolas-блоки, содержимое ячеек при отдельном кегле таблиц, media/relationships
- Структурные заголовки — список из `app.constants.STRUCTURAL_HEADINGS`

## app.core.report (Agent C)

- `render_markdown(result, rules) -> str` ; `render_text(result, rules) -> str`
- `save_report(result, rules, dest_path, fmt="md") -> Path`
- `render_batch_markdown(items) -> str`, `items = [{"file", "summary", "status", "result"?}]`
- `export_pdf(path)` — задел v2: docstring + raise NotImplementedError (архитектурный хук, не заглушка фичи v1)
- `_build_parser()` — задел модуля-парсера библиографии (v2) — НЕ реализовывать парсер

## app/cli.py (Agent D)

Команды п. 19.3: `init doctor rules profile scripts run check fix batch job schema explain context`.
- `main(argv: list[str]|None=None) -> int`
- Конверт JSON п. 19.4; stdout только результат, stderr логи; коды возврата п. 19.5;
  приоритет настроек п. 19.6; `--dry-run` для run/fix; без `input()`.
- `gost init --agents` генерирует AGENTS.md, schemas/, examples/ (агенты-файлы — Agent F их подготовит как эталон; cli хранит их тексты в `app/cli_init.py` или аналоге и записывает).

## app/main.py, app/__main__.py (Agent D)

- `main()` — разбор `--internal-run <script> <params_json>` (без GUI!), иначе GUI
- internal-run: читает params_json `{"argv": [...]}`, ставит `sys.argv = [script, *argv]`,
  `runpy.run_path(script, run_name="__main__")`, перехват SystemExit/исключений (traceback в stdout), код возврата
- `multiprocessing.freeze_support()` первой строкой в `if __name__ == "__main__":`
- GUI импортируется лениво (внутри ветки GUI), чтобы `--internal-run` не тянул customtkinter

## app/ui (Agent E)

Строки только из `app/strings.py`; бизнес-логика в core; все операции в потоках через queue.
Точка входа: `class GostApp(ctk.CTk)` в `app_window.py`, `run_gui(state_dirs)` из `app/ui/app_window.py`.

## Контракт скриптов-плагинов (Agent F пишет scripts/*)

CLI mkdocx: `python3 -I scripts/mkdocx.py input.md output.docx '{"fio":...}' [--rules R.json] [--meta meta.json] [--no-verify]`
make_karta/make_games: `--out PATH [--rules R.json] [--meta meta.json]`, `SCRIPT_META` c `input:"none"`,
`rules_hook:"postprocess"`, контент — из `--meta`/параметров с разумными дефолтами (никакого выдуманного контента без параметров — дефолты помечать как примеры).
Каждый скрипт: `verify(path)` (самопроверка, отчёт в stderr, `--no-verify` отключает), путь к файлу в stdout,
exit 0/nonzero, экранирование `escape(t, {'"': '&quot;'})`, `xml:space="preserve"`, порядок детей pPr/rPr по §5 CHATBOT_TEMPLATE.
