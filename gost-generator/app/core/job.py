"""Декларативная задача ``gost job run task.json`` (п. 19.7 ТЗ).

Job — это единая запись «сгенерируй → нормализуй → проверь» для CLI и ИИ-агентов.
Модуль не импортирует UI и работает headless; каталоги берутся из переданного
``dirs`` (DataDirs из app.core.config), а раннер и пост-хуки — из app.core.runner.
"""

from __future__ import annotations

import dataclasses
import datetime as _dt
import json
import re
import sys
from pathlib import Path
from typing import Any

from app import constants
from app.core import plugin_loader, runner

JOB_SCHEMA_KEYS = {"schema_version", "script", "params", "rules", "profile",
                   "output", "input", "post"}
POST_KEYS = {"normalize", "check", "fail_on"}
VALID_FAIL_ON = ("error", "warning", "never")


class JobError(Exception):
    """Задача невалидна или не может быть выполнена (нет скрипта, каталога и т.п.)."""


def validate_job(job: dict) -> list[str]:
    """Проверить словарь задачи по схеме п. 19.7.

    :return: список проблем; пустой список = задача корректна.
    """
    if not isinstance(job, dict):
        return ["Задача должна быть JSON-объектом"]
    problems: list[str] = []

    version = job.get("schema_version")
    if version is not None and (not isinstance(version, int) or isinstance(version, bool)):
        problems.append("schema_version должен быть целым числом")
    elif isinstance(version, int) and version > constants.SCHEMA_VERSION:
        problems.append(f"Неподдерживаемая schema_version {version} (максимум {constants.SCHEMA_VERSION})")

    script = job.get("script")
    if not isinstance(script, str) or not script.strip():
        problems.append("Поле script обязательно и должно быть непустой строкой")

    for key in ("params",):
        if key in job and not isinstance(job[key], dict):
            problems.append(f"Поле {key} должно быть объектом")

    for key in ("rules", "profile", "output", "input"):
        if key in job and not isinstance(job[key], str):
            problems.append(f"Поле {key} должно быть строкой")

    post = job.get("post")
    if post is not None:
        if not isinstance(post, dict):
            problems.append("Поле post должно быть объектом")
        else:
            unknown = set(post) - POST_KEYS
            if unknown:
                problems.append(f"Неизвестные ключи post: {', '.join(sorted(unknown))}")
            for bkey in ("normalize", "check"):
                if bkey in post and not isinstance(post[bkey], bool):
                    problems.append(f"post.{bkey} должен быть булевым")
            if "fail_on" in post and post["fail_on"] not in VALID_FAIL_ON:
                problems.append(f"post.fail_on должен быть одним из: {', '.join(VALID_FAIL_ON)}")

    unknown_top = {k for k in job if k not in JOB_SCHEMA_KEYS}
    if unknown_top:
        problems.append(f"Неизвестные поля задачи: {', '.join(sorted(unknown_top))}")
    return problems


def _safe_component(text: str, fallback: str = "doc") -> str:
    """Очистить строку для использования в имени файла (тема из профиля/параметров)."""
    cleaned = re.sub(r'[\\/:*?"<>|\s]+', "_", text or "").strip("_")
    return cleaned[:80] or fallback


def _script_search_dirs(dirs: Any) -> list[Path]:
    """Каталоги поиска плагинов: данные пользователя, затем штатные (репо или распаковка exe)."""
    result: list[Path] = []
    user_dir = getattr(dirs, "scripts", None) if dirs is not None else None
    result.append(Path(user_dir) if user_dir else Path.cwd() / constants.DIR_SCRIPTS)
    # В собранном exe штатные скрипты распакованы PyInstaller в sys._MEIPASS.
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        result.append(Path(str(meipass)) / constants.DIR_SCRIPTS)
    result.append(Path(__file__).resolve().parent.parent.parent / constants.DIR_SCRIPTS)
    return result


def _resolve_script(job_script: str, dirs: Any) -> Path:
    """Найти скрипт по абсолютному пути или имени (с `.py` и без) в каталогах плагинов."""
    p = Path(job_script)
    if p.is_file():
        return p
    for base in _script_search_dirs(dirs):
        for candidate in (base / job_script, base / f"{job_script}.py"):
            if candidate.is_file():
                return candidate
    raise JobError(f"Скрипт не найден: {job_script}")


def _load_rules(job: dict, dirs: Any) -> dict:
    """Правила по имени пресета или пути; rules_store агента A может ещё отсутствовать."""
    name_or_path = job.get("rules")
    fallback_dir = (getattr(dirs, "presets", None) if dirs is not None
                    else Path.cwd() / constants.DIR_PRESETS)
    try:
        from app.core import rules_store  # ленивый импорт: модуль пишется параллельно

        store = rules_store.RulesStore(dirs)
        return store.resolve(name_or_path) if name_or_path else store.load(constants.DEFAULT_PRESET_KEY)
    except Exception:
        # Fallback на время параллельной сборки: читаем JSON пресета напрямую.
        # Правила — критичны, но отсутствие/битость rules_store не должны валить job.
        if name_or_path and Path(name_or_path).is_file():
            return json.loads(Path(name_or_path).read_text(encoding="utf-8"))
        key = name_or_path or constants.DEFAULT_PRESET_KEY
        preset_file = Path(fallback_dir or ".") / f"{key}.json"
        if preset_file.is_file():
            return json.loads(preset_file.read_text(encoding="utf-8"))
        return {}


def _load_profile(job: dict, dirs: Any) -> tuple[dict, dict]:
    """Профиль метаданных → (сырой dict, плоский meta)."""
    name = job.get("profile") or constants.DEFAULT_PROFILE_KEY
    profiles_dir = (getattr(dirs, "profiles", None) if dirs is not None
                    else Path.cwd() / constants.DIR_PROFILES)
    raw: dict = {}
    try:
        from app.core import profile_store  # ленивый импорт (агент A)

        store = profile_store.ProfileStore(dirs)
        raw = store.load(name)
        try:
            return raw, profile_store.flat_meta(raw)
        except AttributeError:
            pass
    except Exception:
        # Тот же fallback периода параллельной сборки, что и для правил.
        profile_file = Path(profiles_dir or ".") / f"{name}.json"
        if profile_file.is_file():
            raw = json.loads(profile_file.read_text(encoding="utf-8"))
    return raw, flat_meta(raw)


def flat_meta(profile: dict) -> dict:
    """Плоский словарь meta: секции person/work/institution/extra слиты + name (п. 7.2).

    Дублирует контракт profile_store.flat_meta как fallback на период сборки M-этапов;
    если profile_store доступен, используется он (см. _load_profile).
    """
    meta: dict[str, Any] = {"name": profile.get("name", "")}
    for section in ("person", "work", "institution", "extra"):
        block = profile.get(section)
        if isinstance(block, dict):
            meta.update(block)
    return meta


def _default_output(info: plugin_loader.PluginInfo, job: dict, dirs: Any,
                    params: dict, profile_meta: dict) -> Path:
    """Имя вывода по умолчанию: <тема или name>_<YYYY-MM-DD_HHMM><ext> (п. 10.1)."""
    base = _safe_component(str(params.get("theme") or profile_meta.get("theme") or info.name))
    stamp = _dt.datetime.now().strftime("%Y-%m-%d_%H%M")
    out_dir = (getattr(dirs, "output", None) if dirs is not None
               else Path.cwd() / constants.DIR_OUTPUT)
    return Path(out_dir or ".") / f"{base}_{stamp}{info.output_ext or '.docx'}"


def execute_job(job: dict, *, dirs: Any = None, runner_kwargs: dict | None = None,
                dry_run: bool = False) -> dict[str, Any]:
    """Выполнить декларативную задачу целиком; вернуть конверт п. 19.7/INTERFACES.

    Конверт: ``{"ok", "output", "run", "normalize", "check"}``. При ``dry_run=True``
    возвращается план без записи файлов и без запуска процесса.
    """
    problems = validate_job(job)
    if problems:
        raise JobError("Невалидная задача: " + "; ".join(problems))

    script_path = _resolve_script(str(job["script"]), dirs)
    info = plugin_loader.build_info(script_path)
    params = dict(job.get("params") or {})
    rules = _load_rules(job, dirs)
    profile_raw, profile_meta = _load_profile(job, dirs)
    post = dict(job.get("post") or {})
    do_normalize = bool(post.get("normalize", True))
    do_check = bool(post.get("check", False))
    fail_on = str(post.get("fail_on", "error"))

    input_path = params.get("input") or job.get("input")
    output_raw = job.get("output")
    if output_raw:
        output_path = Path(str(output_raw))
        if not output_path.is_absolute():
            if len(output_path.parts) == 1:
                # Только имя файла → кладём в каталог вывода (п. 3.3).
                out_dir = (getattr(dirs, "output", None) if dirs is not None
                           else Path.cwd() / constants.DIR_OUTPUT)
                output_path = Path(out_dir) / output_path
            else:
                # Путь с каталогами (пример п. 19.7: "output/karta.docx") — от корня данных.
                root = getattr(dirs, "home", None) if dirs is not None else Path.cwd()
                output_path = Path(root or ".") / output_path
    else:
        output_path = _default_output(info, job, dirs, params, profile_meta)

    if dry_run:
        # План без побочных эффектов: tmp_dir=None — JSON не пишутся (см. build_argv).
        argv_plan = plugin_loader.build_argv(
            info, input_path=input_path, output_path=output_path, meta=profile_meta,
            rules=rules, profile=profile_raw, params=params, tmp_dir=None)
        return {
            "ok": True,
            "dry_run": True,
            "plan": {
                "script": str(script_path),
                "argv": argv_plan,
                "output": str(output_path),
                "rules_hook": info.rules_hook,
                "post": {"normalize": do_normalize, "check": do_check, "fail_on": fail_on},
            },
        }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    script_runner = runner.ScriptRunner(**(runner_kwargs or {}))
    outcome = runner.run_and_finalize(
        script_runner, info=info, meta=profile_meta, rules=rules, profile=profile_raw,
        input_path=str(input_path) if input_path else None, output_path=output_path,
        params=params, do_normalize=do_normalize, do_check=do_check, fail_on=fail_on)

    run = outcome["run"]
    run_dict = dataclasses.asdict(run) if dataclasses.is_dataclass(run) else run
    if isinstance(run_dict, dict) and isinstance(run, runner.RunResult):
        # asdict теряет свойство ok — восстанавливаем его явно.
        run_dict["ok"] = run.ok

    return {
        "ok": bool(outcome["ok"]),
        "output": outcome["output"],
        # В конверт кладём словарь: RunResult — дата-класс, а JSON-вывод п. 19.4
        # обязан оставаться структурированным, а не repr()-строкой.
        "run": run_dict,
        "normalize": outcome["normalized"],
        "check": outcome["check"],
    }
