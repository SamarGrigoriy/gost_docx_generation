"""Обработчики «знающих» подкоманд gost: init/doctor/context/schema/explain/
rules/profile/scripts (ТЗ §19.3). Команды документов — в app/cli_docs.py.

Каждый handler возвращает `cli_core.Result`; бизнес-логика делегируется app.core
(один источник истины, п. 19.1). Модуль не импортирует GUI.
"""

from __future__ import annotations

import argparse
import dataclasses
import json
import shutil
import sys
from pathlib import Path
from typing import Any

from app import constants
from app import cli_core, cli_init
from app.cli_core import CliError, Result, make_error

# --- Настройки и мелкие помощники ----------------------------------------------



# --- init / doctor / context / schema / explain ---------------------------------


def cmd_init(args: argparse.Namespace) -> Result:
    """gost init [--agents] [--dir]: каталоги данных + артефакты для агентов (п. 20)."""
    base = Path(args.dir).expanduser() if args.dir else Path.cwd()
    base.mkdir(parents=True, exist_ok=True)
    created = cli_init.ensure_skeleton(base)
    if args.agents:
        created += cli_init.write_agents_artifacts(base)
    warnings: list[Any] = []
    try:
        cli_core.resolve_dirs().ensure()
    except Exception as exc:  # домашний каталог не должен валить init
        warnings.append(make_error("BAD_CONFIG", f"Каталог данных не создан: {exc}", None))
    return Result(data={"dir": str(base), "agents": bool(args.agents), "created": created},
                  warnings=warnings,
                  text=f"Инициализировано в {base}: создано {len(created)} путей")


def cmd_doctor(args: argparse.Namespace) -> Result:
    """gost doctor: самодиагностика окружения (модули, каталоги, настройки)."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    checks: list[dict[str, Any]] = []

    def add(name: str, status: str, detail: str = "") -> None:
        checks.append({"name": name, "status": status, "detail": detail})

    add("python", "ok", f"{sys.version.split()[0]} @ {sys.executable}")
    for m in cli_core.CORE_MODULES:
        mod = cli_core.core_module(m)
        add(f"core.{m}", "ok" if mod else "warn",
            "импортируется" if mod else "модуль ещё не собран")
    add("dirs", "ok", f"home={dirs.home} portable={getattr(dirs, 'is_portable', lambda: None)()}")
    for attr in ("scripts", "presets", "profiles", "output", "logs"):
        d = Path(getattr(dirs, attr, Path(dirs.home) / attr))
        status, detail = "ok", str(d)
        try:
            d.mkdir(parents=True, exist_ok=True)
            probe = d / ".gost_probe"
            probe.write_text("", encoding="utf-8")
            probe.unlink()
        except OSError as exc:
            status, detail = "error", f"{d}: {exc}"
        add(f"dir:{attr}", status, detail)
    presets = cli_core.list_rules(dirs)
    add("presets", "ok" if presets else "warn",
        f"{len(presets)} шт; активен: {ctx['rules_key']}")
    profiles = cli_core.profiles_list(dirs)
    add("profiles", "ok" if profiles else "warn",
        f"{len(profiles)} шт; активен: {ctx['profile_key']}")
    env = {k: v for k, v in ((constants.ENV_HOME, str(dirs.home)),
                             (constants.ENV_RULES, ctx["rules_key"]),
                             (constants.ENV_PROFILE, ctx["profile_key"])) if v}
    add("env", "ok", json.dumps(env, ensure_ascii=False))
    add("gui", "ok", f"CLI headless; GUI в окружении: {cli_core.gui_status()}")
    errors = [make_error("BAD_CONFIG", c["detail"], c["name"])
              for c in checks if c["status"] == "error"]
    data = {"checks": checks,
            "versions": {"app": constants.APP_VERSION, "schema": constants.SCHEMA_VERSION},
            "python": {"version": sys.version.split()[0], "executable": sys.executable}}
    text = "\n".join(f"[{c['status']:>5}] {c['name']}: {c['detail']}" for c in checks)
    return Result(data=data, errors=errors, text=text)


def cmd_context(args: argparse.Namespace) -> Result:
    """gost context: активный пресет, скрипты с SCRIPT_META, профили, каталоги, версии."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    rules = cli_core.rules_of(ctx)
    pl = cli_core.core_module("plugin_loader")
    scripts: list[dict[str, Any]] = []
    for base in cli_core.scripts_search_dirs(dirs):
        if not Path(base).exists():
            continue
        if pl is not None:
            for info in pl.scan_dir(Path(base)):
                scripts.append(cli_core.plugin_to_dict(info))
        else:
            scripts.extend({"path": str(f), "file_name": f.name, "has_meta": None,
                            "error": "plugin_loader недоступен"}
                           for f in sorted(Path(base).glob("*.py")))
    data = {
        "versions": {"app": constants.APP_VERSION, "schema": constants.SCHEMA_VERSION},
        "python": {"version": sys.version.split()[0], "executable": sys.executable},
        "dirs": {"home": str(dirs.home), "scripts": str(dirs.scripts),
                 "presets": str(dirs.presets), "profiles": str(dirs.profiles),
                 "output": str(dirs.output), "logs": str(dirs.logs)},
        "active_rules_key": ctx["rules_key"],
        "rules": rules,
        "presets": cli_core.list_rules(dirs),
        "active_profile": ctx["profile_key"],
        "profiles": cli_core.profiles_list(dirs),
        "scripts": scripts,
        "env": {constants.ENV_HOME: str(dirs.home),
                constants.ENV_RULES: str(ctx["rules_key"]),
                constants.ENV_PROFILE: str(ctx["profile_key"]),
                constants.ENV_OUTPUT: str(cli_core.output_dir(dirs, ctx["cfg"], ctx["state"]))},
    }
    text = (f"пресет: {ctx['rules_key']}  профиль: {ctx['profile_key']}  "
            f"скриптов: {len(scripts)}  home: {dirs.home}")
    return Result(data=data, text=text)


def cmd_schema(args: argparse.Namespace) -> Result:
    """gost schema rules|profile|script-meta|report|job — JSON-схема (п. 20.1 п.2)."""
    schema = cli_core.get_schema(args.name)
    return Result(data=schema, text=json.dumps(schema, ensure_ascii=False, indent=2))


def cmd_explain(args: argparse.Namespace) -> Result:
    """gost explain CODE — значение нарушения и способ исправления (п. 20.1 п.4)."""
    entry = cli_init.explain_entry(args.code)
    text = (f"{args.code}: {entry['title']}\n{entry['detail']}\n"
            f"Как исправить: {entry['fix']}\nИсточник: {entry['source']}")
    return Result(data={"code": args.code, **entry}, text=text)


# --- rules / profile / scripts ---------------------------------------------------


def cmd_rules(args: argparse.Namespace) -> Result:
    """gost rules list|show|validate|export|import|diff — CRUD пресетов через core."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    store_mod = cli_core.core_module("rules_store")
    sub = args.rsub
    if sub == "list":
        items = cli_core.list_rules(dirs)
        text = "\n".join(
            f"{'*' if i['key'] == ctx['rules_key'] else ' '} {i['key']}"
            f"  {'builtin' if i.get('builtin') else 'user'}  {i.get('path', '')}"
            for i in items) or "Пресеты не найдены"
        return Result(data={"items": items, "active": ctx["rules_key"]}, text=text)
    if sub == "show":
        rules = cli_core.resolve_rules(dirs, args.key)
        describe = getattr(store_mod, "describe_rules", None) if store_mod else None
        try:
            text = str(describe(rules)) if describe else ""
        except Exception:
            text = ""
        return Result(data={"key": args.key, "rules": rules},
                      text=text or json.dumps(rules, ensure_ascii=False, indent=2))
    if sub == "validate":
        key = args.key or ctx["rules_key"]
        rules = cli_core.resolve_rules(dirs, key)
        fn = getattr(store_mod, "validate_rules", None) if store_mod else None
        issues = [str(m) for m in (fn(rules) if fn else [])]
        errors = [make_error("BAD_CONFIG", m, key) for m in issues]
        return Result(data={"key": key, "issues": issues}, errors=errors,
                      text="OK" if not issues else "\n".join(issues))
    if sub == "export":
        rules = cli_core.resolve_rules(dirs, args.key)
        dest = Path(args.dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(json.dumps(rules, ensure_ascii=False, indent=2), encoding="utf-8")
        return Result(data={"key": args.key, "dest": str(dest)}, text=str(dest))
    if sub == "import":
        src = Path(args.src)
        if not src.exists():
            raise CliError("FILE_NOT_FOUND", "Файл пресета не найден", str(src))
        data = cli_core.read_json_file(src)
        if not isinstance(data, dict):
            raise CliError("BAD_CONFIG", "Пресет должен быть JSON-объектом", str(src))
        key = args.key or src.stem
        store_cls = getattr(store_mod, "RulesStore", None) if store_mod else None
        if store_cls is not None:
            store = store_cls(dirs, builtin_dir=cli_core.builtin_presets_dir())
            if args.key:
                # явный ключ важнее имени файла: сохраняем пользовательский пресет
                store.save(data, key=key)
            elif hasattr(store, "import_from"):
                key = str(store.import_from(src))
            else:
                store.save(data, key=key)
        else:
            Path(dirs.presets).mkdir(parents=True, exist_ok=True)
            shutil.copyfile(src, Path(dirs.presets) / f"{key}.json")
        return Result(data={"key": key}, text=key)
    if sub == "diff":
        left = cli_core.resolve_rules(dirs, args.left)
        right = cli_core.resolve_rules(dirs, args.right)
        diff = _diff_dicts(left, right)
        text = "\n".join(f"{d['path']}: {d['left']!r} -> {d['right']!r}" for d in diff) \
            or "Пресеты идентичны"
        return Result(data={"left": args.left, "right": args.right, "diff": diff}, text=text)
    raise CliError("USAGE", f"Неизвестная подкоманда rules: {sub}")


def _diff_dicts(a: dict[str, Any], b: dict[str, Any], prefix: str = "") -> list[dict[str, Any]]:
    """Рекурсивный diff словарей пресетов: [{path, left, right}]."""
    out: list[dict[str, Any]] = []
    for key in sorted(set(a) | set(b)):
        path = f"{prefix}{key}"
        va, vb = a.get(key), b.get(key)
        if isinstance(va, dict) and isinstance(vb, dict):
            out.extend(_diff_dicts(va, vb, path + "."))
        elif va != vb:
            out.append({"path": path, "left": va, "right": vb})
    return out


def cmd_profile(args: argparse.Namespace) -> Result:
    """gost profile list|show|set|use|export|import — профили через core."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    store = cli_core.profile_store(dirs)
    sub = args.psub
    if sub == "list":
        items = cli_core.profiles_list(dirs)
        text = "\n".join(f"{'*' if i['key'] == ctx['profile_key'] else ' '} {i['key']}"
                         for i in items) or "Профили не найдены"
        return Result(data={"items": items, "active": ctx["profile_key"]}, text=text)
    if sub == "show":
        profile = cli_core.load_profile(dirs, args.key)
        flat = cli_core.flat_meta(profile)
        return Result(data={"key": args.key, "profile": profile, "flat": flat},
                      text=json.dumps(flat, ensure_ascii=False, indent=2))
    if sub == "set":
        profile = cli_core.load_profile(dirs, args.key)
        changes = {}
        for pair in args.values:
            if "=" not in pair:
                raise CliError("USAGE", f"Ожидается key=value, получено: {pair!r}")
            k, v = pair.split("=", 1)
            changes[k] = v
        _profile_apply(profile, changes)
        _save_profile(store, dirs, args.key, profile)
        return Result(data={"key": args.key, "changes": changes, "profile": profile},
                      text=f"{args.key}: обновлено {len(changes)} полей")
    if sub == "use":
        cli_core.load_profile(dirs, args.key)  # проверка существования
        state = dict(ctx["state"])
        state["active_profile"] = args.key
        cli_core.save_state(dirs, state)
        return Result(data={"active_profile": args.key}, text=f"Активен профиль: {args.key}")
    if sub == "export":
        profile = cli_core.load_profile(dirs, args.key)
        dest = Path(args.dest)
        dest.parent.mkdir(parents=True, exist_ok=True)
        dest.write_text(json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")
        return Result(data={"key": args.key, "dest": str(dest)}, text=str(dest))
    if sub == "import":
        src = Path(args.src)
        if not src.exists():
            raise CliError("FILE_NOT_FOUND", "Файл профиля не найден", str(src))
        data = cli_core.read_json_file(src)
        if not isinstance(data, dict):
            raise CliError("BAD_CONFIG", "Профиль должен быть JSON-объектом", str(src))
        key = args.key or src.stem
        _save_profile(store, dirs, key, data)
        return Result(data={"key": key}, text=key)
    raise CliError("USAGE", f"Неизвестная подкоманда profile: {sub}")


def _save_profile(store: Any, dirs: Any, key: str, profile: dict[str, Any]) -> None:
    """Сохранить профиль: через ProfileStore, иначе прямой запись в каталог."""
    if store is not None and hasattr(store, "save"):
        try:
            store.save(profile, key=key)
            return
        except CliError:
            raise
        except Exception as exc:
            raise CliError("BAD_CONFIG", f"Не удалось сохранить профиль {key}: {exc}") from exc
    base = Path(dirs.profiles)
    base.mkdir(parents=True, exist_ok=True)
    (base / f"{key}.json").write_text(
        json.dumps(profile, ensure_ascii=False, indent=2), encoding="utf-8")


def _profile_apply(profile: dict[str, Any], changes: dict[str, str]) -> None:
    """Обновить плоские ключи профиля: известная секция — в секцию, иначе в extra."""
    for k, v in changes.items():
        target = None
        for s in ("person", "work", "institution"):
            if isinstance(profile.get(s), dict) and k in profile[s]:
                target = profile[s]
                break
        if target is None:
            target = profile.setdefault("extra", {})
            if not isinstance(target, dict):
                raise CliError("BAD_CONFIG", "Поле extra профиля должно быть объектом")
        target[k] = int(v) if k == "year" else v




def cmd_scripts(args: argparse.Namespace) -> Result:
    """gost scripts list|describe|add|validate — работа с плагинами через core."""
    ctx = cli_core.make_ctx(args)
    dirs = ctx["dirs"]
    pl = cli_core.require_core("plugin_loader", "scan_dir")
    sub = args.ssub
    if sub == "list":
        infos = pl.scan_dir(Path(dirs.scripts),
                            extra_paths=[cli_core.repo_root() / constants.DIR_SCRIPTS])
        items = [cli_core.plugin_to_dict(i) for i in infos]
        text = "\n".join(
            f"{i.get('name') or i.get('file_name')}  [{i.get('input')}]  {i.get('path')}"
            + (f"  ОШИБКА: {i.get('error')}" if i.get("error") else "")
            for i in items) or "Скрипты не найдены"
        return Result(data={"items": items}, text=text)
    if sub == "describe":
        path = cli_core.find_script(dirs, args.name)
        info = pl.load_meta_for_path(path)
        meta = info if isinstance(info, dict) else cli_core.plugin_to_dict(info) if info else None
        return Result(data={"path": str(path), "meta": meta},
                      text=json.dumps(meta, ensure_ascii=False, indent=2)
                      if meta else f"SCRIPT_META не объявлен: {path}")
    if sub == "validate":
        path = Path(args.path)
        if not path.exists():
            raise CliError("FILE_NOT_FOUND", "Скрипт не найден", str(path))
        ok, problems = pl.validate(path)
        problems = [str(p) for p in problems]
        errors = [] if ok else [make_error("BAD_CONFIG", p, str(path)) for p in problems]
        return Result(data={"path": str(path), "valid": bool(ok), "problems": problems},
                      errors=errors, text="OK" if ok else "\n".join(problems))
    if sub == "add":
        path = Path(args.path)
        if not path.exists():
            raise CliError("FILE_NOT_FOUND", "Скрипт не найден", str(path))
        ok, problems = pl.validate(path)
        if not ok:
            raise CliError("BAD_CONFIG",
                           "Скрипт не проходит проверку: " + "; ".join(map(str, problems)),
                           str(path))
        dest_dir = Path(dirs.scripts)
        dest_dir.mkdir(parents=True, exist_ok=True)
        dest = dest_dir / path.name
        shutil.copyfile(path, dest)
        sidecar = path.with_name(path.stem + constants.SIDECAR_SUFFIX)
        if sidecar.exists():
            shutil.copyfile(sidecar, dest_dir / sidecar.name)
        warning = {"code": "SCRIPT_ADDED",
                   "message": "Скрипт выполняется с правами пользователя (п. 12).",
                   "path": str(dest)}
        return Result(data={"path": str(dest)}, warnings=[warning],
                      text=f"Добавлен скрипт: {dest}")
    raise CliError("USAGE", f"Неизвестная подкоманда scripts: {sub}")


