"""Команды работы с документами: run / check / fix / batch / job (ТЗ §19.3).

Вся бизнес-логика — в app.core (inspector/normalizer/runner/job/report);
модуль лишь оркестрирует вызовы и упаковывает результат в конверт п. 19.4.
GUI не импортируется, модуль работает headless.
"""

from __future__ import annotations

import argparse
import dataclasses
import json
import sys
from datetime import datetime
from pathlib import Path
from typing import Any

from app import constants
from app import cli_core, cli_init
from app.cli_core import CliError, Result, make_error

def _parse_params(raw: str | None) -> dict[str, Any]:
    """--params JSON|@file → dict; битый JSON — ошибка INVALID_JSON (п. 12)."""
    if not raw:
        return {}
    if raw.startswith("@"):
        data = cli_core.read_json_file(Path(raw[1:]))
    else:
        try:
            data = json.loads(raw)
        except ValueError as exc:
            raise CliError("INVALID_JSON", f"--params: некорректный JSON: {exc}") from exc
    if not isinstance(data, dict):
        raise CliError("INVALID_JSON", "--params должен быть JSON-объектом")
    return data


def _issues_as_dicts(issues: Any) -> list[dict[str, Any]]:
    """Замечания инспектора (Issue|dict) → список словарей для конверта."""
    return [it if isinstance(it, dict) else dataclasses.asdict(it)
            for it in (issues or [])]


def _apply_fail_on(check_data: dict[str, Any], fail_on: str,
                   errors: list[dict[str, Any]], where: str) -> None:
    """Порог --fail-on: при нарушениях добавляет ошибку VIOLATIONS (код возврата 1)."""
    summary = check_data.get("summary") or {}
    n_err = int(summary.get("error", 0))
    n_warn = int(summary.get("warning", 0))
    if (fail_on == "error" and n_err > 0) or (fail_on == "warning" and n_err + n_warn > 0):
        errors.append(make_error(
            "VIOLATIONS",
            f"Нарушения по порогу --fail-on={fail_on}: ошибок {n_err}, "
            f"предупреждений {n_warn}", where))


def _call_report(fn_name: str, *args: Any, **kwargs: Any) -> Any:
    """Вызвать функцию app.core.report, если модуль собран; иначе None."""
    mod = cli_core.core_module("report")
    fn = getattr(mod, fn_name, None) if mod else None
    return fn(*args, **kwargs) if fn else None


def _iso_issues(result: dict[str, Any]) -> dict[str, Any]:
    """Нормализовать результат проверки к контракту data (п. 19.4)."""
    out = dict(result)
    issues = []
    for it in _issues_as_dicts(result.get("issues")):
        loc = dict(it.get("location") or {})
        issues.append({"code": it.get("code"), "severity": it.get("severity"),
                       "message": it.get("message"), "hint": it.get("hint", ""),
                       "location": {"paragraph": loc.get("paragraph"),
                                    "page": loc.get("page"),
                                    "snippet": loc.get("snippet", "")}})
    out["issues"] = issues
    out["summary"] = {sev: int((result.get("summary") or {}).get(sev, 0))
                      for sev in constants.SEVERITIES}
    return out


def _check_one(path: Path, rules: dict[str, Any], preset_key: str,
               strict: bool) -> dict[str, Any]:
    """Проверка одного файла через app.core.docx_inspector (один источник истины)."""
    inspector = cli_core.require_core("docx_inspector", "check_docx_result")
    eff_rules = {**rules, "strict": strict}
    try:
        raw = inspector.check_docx_result(str(path), eff_rules)
    except Exception as exc:  # битый zip/XML — read_error, не падение (п. 10.3)
        raw = {"file": str(path), "preset": preset_key, "strict": strict,
               "summary": {"error": 0, "warning": 0, "info": 0}, "issues": [],
               "read_error": str(exc)}
    result = _iso_issues(dict(raw))
    result.setdefault("file", str(path))
    result.setdefault("preset", preset_key)
    result.setdefault("strict", strict)
    return result


def _normalize_one(path: Path, rules: dict[str, Any], backup: bool) -> dict[str, Any]:
    """Нормализация через core.docx_normalizer; сбой → CliError (код возврата 7)."""
    norm_mod = cli_core.require_core("docx_normalizer", "normalize_docx")
    try:
        res = norm_mod.normalize_docx(str(path), rules, backup=backup)
    except Exception as exc:
        raise CliError("NORMALIZE_FAILED", f"Нормализация не удалась: {exc}", str(path)) from exc
    result = dict(res) if isinstance(res, dict) else {"changed": bool(res), "changes": []}
    if result.get("error"):
        raise CliError("NORMALIZE_FAILED", str(result["error"]), str(path))
    return result


def _check_text(result: dict[str, Any], rules: dict[str, Any], fmt: str) -> str:
    """Текстовое/markdown-представление результата (через core.report, иначе fallback)."""
    if fmt in ("text", "md"):
        rendered = _call_report("render_markdown" if fmt == "md" else "render_text",
                                result, rules)
        if rendered:
            return str(rendered)
    lines = [f"Файл: {result.get('file')}  пресет: {result.get('preset')}  "
             f"strict: {result.get('strict')}"]
    s = result.get("summary") or {}
    lines.append(f"Ошибок: {s.get('error', 0)}  предупреждений: {s.get('warning', 0)}  "
                 f"info: {s.get('info', 0)}")
    for it in result.get("issues") or []:
        loc = (it.get("location") or {}).get("paragraph")
        lines.append(f"[{it.get('severity')}] {it.get('code')} (абз. {loc}): {it.get('message')}")
    return "\n".join(lines)


def _save_report(dest: str, results: list[dict[str, Any]], rules: dict[str, Any],
                 fmt: str) -> None:
    """Сохранить отчёт: core.report, при недоступности — простой markdown в CLI (presentation)."""
    dest_path = Path(dest)
    dest_path.parent.mkdir(parents=True, exist_ok=True)
    if len(results) == 1:
        rendered = _call_report("save_report", results[0], rules, dest_path,
                               "md" if fmt in ("json", "md") else "text")
        if rendered is not None:
            return
    lines = ["# Отчёт проверки", ""]
    for r in results:
        s = r.get("summary") or {}
        lines += [f"## {r.get('file')}", "",
                  f"пресет: {r.get('preset')}  strict: {r.get('strict')}",
                  f"ошибок: {s.get('error', 0)}  предупреждений: {s.get('warning', 0)}", ""]
        for it in r.get("issues") or []:
            lines.append(f"- [{it.get('severity')}] `{it.get('code')}` {it.get('message')}")
        lines.append("")
    dest_path.write_text("\n".join(lines), encoding="utf-8")


def _status_text(status: Any) -> Any:
    """Русская подпись статуса для текстовой сводки (ключи «violations» и т.п.)."""
    label = _call_report("_status_label", str(status or ""))
    return label if label else status


def _batch_table(items: list[dict[str, Any]]) -> str:
    """Markdown-таблица сводки пакета (fallback, если core.report недоступен)."""
    lines = ["| Файл | Ошибок | Предупреждений | Статус |", "|---|---|---|---|"]
    for i in items:
        s = i.get("summary") or {}
        lines.append(f"| {Path(i['file']).name} | {s.get('error', 0)} | "
                     f"{s.get('warning', 0)} | {_status_text(i.get('status'))} |")
    return "\n".join(lines)



# --- run / check / fix / batch / job ----------------------------------------------


def _default_output(dirs: Any, ctx: dict[str, Any], script_name: str) -> Path:
    """Каталог вывода (п. 19.6) и имя <тема или скрипт>_<дата_время>.docx (п. 10.1)."""
    out_dir = Path(cli_core.output_dir(dirs, ctx["cfg"], ctx["state"]))
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M")
    stem = Path(script_name).stem
    try:
        profile = cli_core.load_profile(dirs, ctx["profile_key"])
        theme = str((profile.get("work") or {}).get("theme") or "").strip()
        if theme:
            stem = theme[:60]
    except CliError:
        pass
    safe = "".join(c for c in stem if c not in '\\/:*?"<>|').strip() or "doc"
    out_dir.mkdir(parents=True, exist_ok=True)
    return out_dir / f"{safe}_{stamp}.docx"


def cmd_run(args: argparse.Namespace) -> Result:
    """gost run SCRIPT: генерация плагином + нормализация/проверка (п. 19.3)."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    rules = cli_core.rules_of(ctx)
    path = cli_core.find_script(dirs, args.script)
    params = _parse_params(args.params)
    pl = cli_core.require_core("plugin_loader", "load_meta_for_path", "build_argv")
    runner_mod = cli_core.require_core("runner", "ScriptRunner")
    # PluginInfo даёт и мета-поля, и описание для build_argv; load_meta_for_path — только dict меты
    info = pl.build_info(path) if hasattr(pl, "build_info") else pl.load_meta_for_path(path)
    profile = cli_core.load_profile(dirs, ctx["profile_key"])
    meta = cli_core.flat_meta(profile)
    inp = Path(args.input) if args.input else None
    if inp is not None and not inp.exists():
        raise CliError("FILE_NOT_FOUND", "Входной файл не найден", str(inp))
    out = Path(args.output) if args.output else _default_output(dirs, ctx, args.script)
    info_d = cli_core.plugin_to_dict(info) if info is not None else {}
    hook = info_d.get("rules_hook") or constants.DEFAULT_RULES_HOOK
    do_normalize = args.normalize if args.normalize is not None else (hook == "postprocess")

    if args.dry_run:
        argv = pl.build_argv(info, input_path=str(inp) if inp else None,
                             output_path=str(out), meta=meta, rules=rules,
                             profile=profile, params=params,
                             tmp_dir=Path(dirs.home) / ".gost_tmp")
        data = {"dry_run": True, "script": str(path), "output": str(out),
                "argv": [str(a) for a in argv],
                "normalize": do_normalize, "check": bool(args.check)}
        return Result(data=data, text="План запуска: " + " ".join(map(str, argv)))

    timeout = int(ctx["cfg"].get("script_timeout_s")
                  or ctx["state"].get("script_timeout_s")
                  or constants.DEFAULT_SCRIPT_TIMEOUT_S)
    runner = runner_mod.ScriptRunner(
        timeout_s=timeout,
        mode=ctx["cfg"].get("interpreter_mode") or ctx["state"].get("interpreter_mode") or "internal",
        python_path=ctx["cfg"].get("python_path") or ctx["state"].get("python_path"),
        on_line=cli_core.log)
    try:
        run_res = runner.run_script(info, meta=meta, rules=rules, profile=profile,
                                    input_path=str(inp) if inp else None,
                                    output_path=str(out), params=params)
    except Exception as exc:
        raise CliError("SCRIPT_FAILED", f"Не удалось запустить скрипт: {exc}", str(path)) from exc

    errors: list[dict[str, Any]] = []
    warnings: list[Any] = []
    run_d = {"argv": [str(a) for a in getattr(run_res, "argv", [])],
             "exit_code": getattr(run_res, "exit_code", None),
             "duration_s": getattr(run_res, "duration_s", None),
             "cancelled": bool(getattr(run_res, "cancelled", False)),
             "timed_out": bool(getattr(run_res, "timed_out", False)),
             "error": getattr(run_res, "error", None)}
    data: dict[str, Any] = {"script": str(path), "output": str(out), "run": run_d}
    if run_d["timed_out"]:
        errors.append(make_error("TIMEOUT", f"Скрипт превысил таймаут {timeout} с", str(out)))
    elif not getattr(run_res, "ok", run_d["exit_code"] == 0):
        errors.append(make_error("SCRIPT_FAILED",
                                 f"Скрипт завершился с кодом {run_d['exit_code']}", str(path)))
    elif not out.exists():
        errors.append(make_error("SCRIPT_FAILED",
                                 "Скрипт завершился успешно, но файл не создан", str(out)))
    else:
        if do_normalize:
            data["normalized"] = _normalize_one(out, rules, backup=True)
        if args.check:
            data["check"] = _check_one(out, rules, str(ctx["rules_key"]),
                                       bool(rules.get("strict", True)))
            _apply_fail_on(data["check"], args.fail_on, errors, str(out))
    text = str(out) if not errors else "\n".join(e["message"] for e in errors)
    return Result(data=data, errors=errors, warnings=warnings, text=text)


def cmd_check(args: argparse.Namespace) -> Result:
    """gost check PATH...: проверка документов активным пресетом (п. 19.3)."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    paths = [Path(p) for p in args.paths]
    missing = [p for p in paths if not p.exists()]
    if len(missing) == len(paths):
        raise CliError("FILE_NOT_FOUND", "Файл не найден", str(missing[0]))
    rules = cli_core.rules_of(ctx)
    flag = getattr(args, "strict", None)
    strict = bool(flag) if flag is not None else bool(rules.get("strict", True))
    errors: list[dict[str, Any]] = [make_error("FILE_NOT_FOUND", "Файл не найден", str(p))
                                    for p in missing]
    warnings: list[Any] = []
    results: list[dict[str, Any]] = []
    for p in paths:
        if not p.exists():
            cli_core.log(f"пропуск: нет файла {p}")
            continue
        res = _check_one(p, rules, str(ctx["rules_key"]), strict)
        results.append(res)
        if res.get("read_error"):
            warnings.append(make_error("READ_ERROR", str(res["read_error"]), str(p)))
            continue
        _apply_fail_on(res, args.fail_on, errors, str(p))
    data: Any = results[0] if len(paths) == 1 else {
        "results": results,
        "summary": {sev: sum(r["summary"].get(sev, 0) for r in results)
                    for sev in constants.SEVERITIES},
    }
    if args.report:
        _save_report(args.report, results, rules, args.format)
        if isinstance(data, dict):
            data["report"] = str(args.report)
    if args.format == "json":
        text = json.dumps(data, ensure_ascii=False, indent=2)
    else:
        text = _check_text(results[0], rules, args.format) if len(results) == 1 \
            else "\n\n".join(_check_text(r, rules, args.format) for r in results)
    return Result(data=data, errors=errors, warnings=warnings, text=text)


def cmd_fix(args: argparse.Namespace) -> Result:
    """gost fix PATH...: нормализация оформления; идемпотентна, с бэкапом (п. 19.4)."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    paths = [Path(p) for p in args.paths]
    missing = [p for p in paths if not p.exists()]
    if len(missing) == len(paths):
        raise CliError("FILE_NOT_FOUND", "Файл не найден", str(missing[0]))
    rules = cli_core.rules_of(ctx)
    items: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = [make_error("FILE_NOT_FOUND", "Файл не найден", str(p))
                                    for p in missing]
    for p in paths:
        if not p.exists():
            continue
        if args.dry_run:
            res = _check_one(p, rules, str(ctx["rules_key"]),
                             bool(rules.get("strict", True)))
            plan = sorted({it.get("code", "") for it in res.get("issues") or []
                           if it.get("severity") in ("error", "warning")})
            items.append({"file": str(p), "dry_run": True, "planned_actions": plan,
                          "would_normalize": bool(plan)})
            continue
        cli_core.log(f"fix: {p}")
        try:
            nres = _normalize_one(p, rules, backup=bool(args.backup))
        except CliError as exc:
            errors.append(make_error(exc.code, exc.message, str(p)))
            continue
        items.append({"file": str(p), "dry_run": False, **nres})
    data = items[0] if (len(paths) == 1 and items) else {"items": items}
    text = "\n".join(
        f"{i['file']}: " + (f"план: {', '.join(i.get('planned_actions', [])) or 'нет изменений'}"
                            if i.get("dry_run")
                            else f"изменено: {i.get('changed')}, "
                                 f"{len(i.get('changes', []))} правок")
        for i in items) or "Нет изменений"
    return Result(data=data, errors=errors, text=text)


def cmd_batch(args: argparse.Namespace) -> Result:
    """gost batch DIR: пакетная проверка/исправление каталога (п. 10.3, 19.3)."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    base = Path(args.directory)
    if not base.is_dir():
        raise CliError("DIR_NOT_FOUND", "Каталог не найден", str(base))
    rules = cli_core.rules_of(ctx)
    strict = bool(rules.get("strict", True))
    do_check = bool(args.check) or not args.fix
    files = sorted(base.glob(args.glob))
    items: list[dict[str, Any]] = []
    errors: list[dict[str, Any]] = []
    for f in files:
        entry: dict[str, Any] = {"file": str(f),
                                 "summary": {"error": 0, "warning": 0, "info": 0},
                                 "status": "ok"}
        cli_core.log(f"batch: {f.name}")
        if args.fix:
            try:
                entry["changed"] = bool(_normalize_one(f, rules, backup=True).get("changed"))
            except CliError as exc:
                entry["status"] = "normalize_error"
                errors.append(make_error(exc.code, exc.message, str(f)))
        if do_check:
            res = _check_one(f, rules, str(ctx["rules_key"]), strict)
            entry["summary"] = res["summary"]
            if res.get("read_error"):
                entry["status"] = "read_error"
            elif entry["summary"].get("error"):
                entry["status"] = "violations"
            if args.full:
                entry["issues"] = res["issues"]
        items.append(entry)
    data: dict[str, Any] = {"dir": str(base), "glob": args.glob, "items": items}
    if args.report:
        rendered = _call_report("render_batch_markdown", items)
        dest = Path(args.report)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(str(rendered) if rendered else _batch_table(items), encoding="utf-8")
        data["report"] = str(dest)
    return Result(data=data, errors=errors, text=_batch_table(items))


def cmd_job(args: argparse.Namespace) -> Result:
    """gost job run task.json: генерация+нормализация+проверка одной операцией (п. 19.7)."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    task = Path(args.task)
    if not task.exists():
        raise CliError("FILE_NOT_FOUND", "Файл задачи не найден", str(task))
    job = cli_core.read_json_file(task)
    if not isinstance(job, dict):
        raise CliError("JOB_INVALID", "task.json должен быть объектом JSON", str(task))
    job_mod = cli_core.require_core("job", "validate_job", "execute_job")
    problems = [str(p) for p in job_mod.validate_job(job)]
    if problems:
        raise CliError("JOB_INVALID", "; ".join(problems), str(task))
    job_error = getattr(job_mod, "JobError", CliError)
    try:
        envelope = job_mod.execute_job(job, dirs=dirs)
    except CliError:
        raise
    except job_error as exc:
        raise CliError("JOB_INVALID", str(exc), str(task)) from exc
    except Exception as exc:
        raise CliError("SCRIPT_FAILED", f"Задача не выполнена: {exc}", str(task)) from exc
    data = dict(envelope) if isinstance(envelope, dict) else {"result": envelope}
    errors: list[dict[str, Any]] = []
    run_d = data.get("run")
    run_d = run_d if isinstance(run_d, dict) else (
        dataclasses.asdict(run_d) if dataclasses.is_dataclass(run_d) else {})
    out = str(data.get("output") or "")
    if run_d.get("timed_out"):
        errors.append(make_error("TIMEOUT", "Таймаут скрипта генерации", out))
    elif run_d.get("exit_code") not in (0, None) or run_d.get("error"):
        errors.append(make_error("SCRIPT_FAILED",
                                 f"Скрипт завершился с кодом {run_d.get('exit_code')}", out))
    norm = data.get("normalized") or data.get("normalize") or {}
    if isinstance(norm, dict) and norm.get("error"):
        errors.append(make_error("NORMALIZE_FAILED", str(norm["error"]), out))
    check = data.get("check")
    post = job.get("post") or {}
    if isinstance(check, dict) and check.get("summary"):
        check = _iso_issues(check)
        data["check"] = check
        _apply_fail_on(check, str(post.get("fail_on", "error")), errors, out)
    ok = bool(data.get("ok", True)) and not errors
    return Result(data=data, errors=errors, ok=ok,
                  text=f"job: {'ok' if ok else 'ошибка'}  output={data.get('output')}")


