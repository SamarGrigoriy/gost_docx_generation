"""Загрузка плагинов-генераторов (п. 8, 8.2, 8.3 ТЗ).

Ключевое требование безопасности: SCRIPT_META извлекается строго через `ast`
(`ast.parse` + поиск присваивания + `ast.literal_eval`). Код скрипта НИКОГДА не
импортируется и не исполняется при сканировании — пользователь кладёт в каталог
чужие .py-файлы, и их побочные эффекты не должны срабатывать при показе списка.

Приоритет метаданных: sidecar `<name>.meta.json` > SCRIPT_META в коде.
"""

from __future__ import annotations

import ast
import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from app import constants

META_VAR_NAME = "SCRIPT_META"
PLACEHOLDER_RE = re.compile(r"\{([A-Za-z_][A-Za-z0-9_]*)\}")


class PluginError(Exception):
    """Ошибка разбора/валидации скрипта-плагина (битый синтаксис, нечитаемый файл)."""


def parse_meta(source: str) -> dict | None:
    """Извлечь SCRIPT_META из исходного кода без его исполнения.

    Ищем присваивание `SCRIPT_META = <литерал>` на уровне модуля и вычисляем
    правую часть через ast.literal_eval. Никакого import/exec — скрипт с побочными
    эффектами при импорте остаётся безопасным текстом.

    :param source: исходный код .py
    :return: словарь метаданных или None, если SCRIPT_META не объявлен
    :raises PluginError: синтаксическая ошибка или SCRIPT_META не является литералом-словарём
    """
    try:
        tree = ast.parse(source)
    except SyntaxError as exc:
        raise PluginError(f"Синтаксическая ошибка в скрипте: {exc}") from exc

    for node in tree.body:
        targets: list[ast.expr] = []
        value: ast.expr | None = None
        if isinstance(node, ast.Assign):
            targets, value = list(node.targets), node.value
        elif isinstance(node, ast.AnnAssign) and node.value is not None:
            targets, value = [node.target], node.value
        for target in targets:
            if isinstance(target, ast.Name) and target.id == META_VAR_NAME:
                if not isinstance(value, ast.Dict):
                    raise PluginError("SCRIPT_META должен быть литералом-словарём")
                try:
                    meta = ast.literal_eval(value)
                except (ValueError, TypeError, SyntaxError) as exc:
                    raise PluginError(f"SCRIPT_META не является литералом: {exc}") from exc
                if not isinstance(meta, dict):
                    raise PluginError("SCRIPT_META должен быть словарём")
                return meta
    return None


def sidecar_path(script_path: Path | str) -> Path:
    """Путь sidecar-файла описания: <stem>.meta.json рядом со скриптом."""
    p = Path(script_path)
    return p.with_name(p.stem + constants.SIDECAR_SUFFIX)


def load_meta_for_path(path: Path | str) -> dict | None:
    """Загрузить метаду скрипта: sidecar приоритетнее SCRIPT_META в коде.

    Битый sidecar не считается фатальным — в этом случае пробуем прочитать мету из кода.
    """
    p = Path(path)
    side = sidecar_path(p)
    if side.exists():
        try:
            data = json.loads(side.read_text(encoding="utf-8"))
            if isinstance(data, dict):
                return data
        except (OSError, ValueError):
            pass  # sidecar повреждён — fallback на SCRIPT_META из кода
    try:
        source = p.read_text(encoding="utf-8", errors="replace")
    except OSError as exc:
        raise PluginError(f"Не удалось прочитать файл: {exc}") from exc
    return parse_meta(source)


@dataclass
class PluginInfo:
    """Карточка скрипта-плагина для списка на вкладке «Генерация» (п. 8 ТЗ)."""

    path: Path
    file_name: str
    name: str = ""
    description: str = ""
    category: str = "Без категории"
    version: str = ""
    input: str = "none"
    output_ext: str = ".docx"
    arg_template: list[str] = field(default_factory=list)
    rules_hook: str = constants.DEFAULT_RULES_HOOK
    params: list[dict[str, Any]] = field(default_factory=list)
    has_meta: bool = False
    error: str | None = None

    @classmethod
    def from_meta(cls, path: Path | str, meta: dict | None) -> "PluginInfo":
        """Собрать карточку из словаря метаданных (или из имени файла, если меты нет)."""
        p = Path(path)
        info = cls(path=p, file_name=p.name)
        if meta is None:
            # П. 8.2: скрипт без меты показывается по имени файла, has_meta=False.
            info.name = p.stem
            info.has_meta = False
            return info
        info.name = str(meta.get("name") or p.stem)
        info.description = str(meta.get("description") or "")
        info.category = str(meta.get("category") or "Без категории")
        info.version = str(meta.get("version") or "")
        info.input = str(meta.get("input") or "none")
        info.output_ext = str(meta.get("output_ext") or ".docx")
        template = meta.get("arg_template") or []
        info.arg_template = [str(t) for t in template] if isinstance(template, list) else []
        info.rules_hook = str(meta.get("rules_hook") or constants.DEFAULT_RULES_HOOK)
        params = meta.get("params") or []
        if isinstance(params, list):
            info.params = [dict(x) for x in params if isinstance(x, dict)]
        info.has_meta = True
        return info

    def to_meta_dict(self) -> dict[str, Any]:
        """Обратное преобразование: словарь для sidecar-файла (схема п. 8)."""
        return {
            "name": self.name,
            "description": self.description,
            "category": self.category,
            "version": self.version,
            "input": self.input,
            "output_ext": self.output_ext,
            "arg_template": list(self.arg_template),
            "rules_hook": self.rules_hook,
            "params": [dict(p) for p in self.params],
        }


def build_info(path: Path | str) -> PluginInfo:
    """Карточка одного файла. Ошибки не выбрасываются — пишутся в поле `error`."""
    p = Path(path)
    try:
        meta = load_meta_for_path(p)
    except (PluginError, OSError) as exc:
        # П. 8.3: битый скрипт не должен выпадать из списка — показываем с текстом проблемы.
        info = PluginInfo(path=p, file_name=p.name, name=p.stem, has_meta=False)
        info.error = str(exc)
        return info
    return PluginInfo.from_meta(p, meta)


def scan_dir(directory: Path | str, extra_paths: tuple[Path | str, ...] = ()) -> list[PluginInfo]:
    """Просканировать каталог .py-скриптов плюс явно добавленные внешние пути.

    Сортировка по имени файла даёт стабильный порядок в UI между «Обновить» (F5).
    """
    d = Path(directory)
    files: list[Path] = sorted(d.glob("*.py")) if d.is_dir() else []
    for extra in extra_paths:
        ep = Path(extra)
        if ep.is_file() and ep not in files:
            files.append(ep)
    return [build_info(f) for f in files]


def validate(path: Path | str) -> tuple[bool, list[str]]:
    """Проверить скрипт перед добавлением (кнопка «Добавить скрипт…»).

    :return: (ok, список проблем). Пустой список проблем = скрипт пригоден.
    """
    p = Path(path)
    problems: list[str] = []
    if not p.exists():
        return False, [f"Файл не найден: {p}"]

    info = build_info(p)
    if info.error:
        return False, [info.error]
    if not info.has_meta:
        problems.append("SCRIPT_META не объявлен — скрипт будет доступен с ручными аргументами")
        return True, problems

    if info.input not in constants.INPUT_TYPES:
        problems.append(
            f"Неизвестный тип input: {info.input!r} (допустимо: {', '.join(constants.INPUT_TYPES)})"
        )
    if info.rules_hook not in constants.RULES_HOOKS:
        problems.append(
            f"Неизвестный rules_hook: {info.rules_hook!r} "
            f"(допустимо: {', '.join(constants.RULES_HOOKS)})"
        )
    # Проверяем исходную мету: arg_template обязан быть списком строк.
    try:
        meta_raw = load_meta_for_path(p) or {}
    except (PluginError, OSError):
        meta_raw = {}
    template = meta_raw.get("arg_template")
    if template is not None and not isinstance(template, list):
        problems.append("arg_template должен быть списком строк")
    elif isinstance(template, list) and any(not isinstance(t, str) for t in template):
        problems.append("Элементы arg_template должны быть строками")

    for i, param in enumerate(info.params):
        if not param.get("id"):
            problems.append(f"Параметр #{i + 1}: отсутствует id")
        ptype = param.get("type", "text")
        if ptype not in constants.PARAM_TYPES:
            problems.append(
                f"Параметр {param.get('id', f'#{i + 1}')}: неизвестный тип {ptype!r}"
            )
        if ptype == "select" and not param.get("options"):
            problems.append(f"Параметр {param.get('id', f'#{i + 1}')}: select без options")
    return (not problems), problems


def _json_substitute(value: Any) -> str:
    """Строковое представление значения параметра для подстановки в argv."""
    if isinstance(value, str):
        return value
    if isinstance(value, bool):
        # CLI-скрипты обычно ждут lowercase-строки или флаги; json.dumps даёт true/false.
        return json.dumps(value)
    return json.dumps(value, ensure_ascii=False)


def _write_json(tmp_dir: Path, name: str, data: Any) -> str:
    """Записать JSON-артефакт во временный каталог запуска, вернуть путь строкой."""
    tmp_dir.mkdir(parents=True, exist_ok=True)
    target = tmp_dir / name
    target.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return str(target)


def _substitutions(*, input_path: Path | str | None, output_path: Path | str | None,
                   meta: dict, rules: dict, profile: dict, params: dict,
                   tmp_dir: Path | None) -> dict[str, str]:
    """Словарь подстановок arg_template: служебные токены + id параметров (п. 8)."""
    if tmp_dir is not None:
        meta_json = _write_json(tmp_dir, "gost_meta.json", meta)
        rules_json = _write_json(tmp_dir, "gost_rules.json", rules)
        profile_json = _write_json(tmp_dir, "gost_profile.json", profile)
    else:
        # Режим предпросмотра (dry-run): пути показываем, файлы не пишем.
        meta_json = "<tmp>/gost_meta.json"
        rules_json = "<tmp>/gost_rules.json"
        profile_json = "<tmp>/gost_profile.json"

    subs: dict[str, str] = {str(k): _json_substitute(v) for k, v in params.items()}
    subs["input"] = str(input_path) if input_path else ""
    subs["output"] = str(output_path) if output_path else ""
    subs["meta_json"] = meta_json
    subs["rules_json"] = rules_json
    subs["profile_json"] = profile_json
    return subs


def _expand_token(token: str, subs: dict[str, str]) -> str:
    """Заменить в токоне все известные {placeholder}; неизвестные оставляем как есть."""

    def repl(match: re.Match[str]) -> str:
        return subs.get(match.group(1), match.group(0))

    return PLACEHOLDER_RE.sub(repl, token)


def build_argv(info: PluginInfo | dict, *, input_path: Path | str | None = None,
               output_path: Path | str | None = None, meta: dict, rules: dict,
               profile: dict, params: dict | None = None,
               tmp_dir: Path | None = None) -> list[str]:
    """Собрать argv CLI-скрипта по arg_template (п. 8 ТЗ).

    Подстановки: {input} {output} {meta_json} {rules_json} {profile_json} и
    {param_id} для любого параметра. JSON-артефакты пишутся в tmp_dir; при
    tmp_dir=None (dry-run) пути только показываются. При пустом arg_template —
    универсальная форма: [input?, output, json.dumps(params)].
    """
    if isinstance(info, dict):
        meta_like = dict(info)
        path = Path(meta_like.get("path") or "")
        info_obj = PluginInfo.from_meta(path, meta_like if "params" in meta_like else None)
    else:
        info_obj = info
    params = dict(params or {})
    # Дефолты из меты применяются только для параметров, не переданных пользователем.
    for p in info_obj.params:
        pid = str(p.get("id", ""))
        if pid and pid not in params:
            params[pid] = p.get("default", "")

    subs = _substitutions(input_path=input_path, output_path=output_path, meta=meta,
                          rules=rules, profile=profile, params=params, tmp_dir=tmp_dir)

    if info_obj.arg_template:
        return [_expand_token(t, subs) for t in info_obj.arg_template if _expand_token(t, subs) != ""]

    argv: list[str] = []
    if input_path and info_obj.input != "none":
        argv.append(str(input_path))
    if output_path:
        argv.append(str(output_path))
    argv.append(json.dumps(params, ensure_ascii=False))
    return argv


def save_sidecar(info: PluginInfo, **overrides: Any) -> Path:
    """Сохранить описание скрипта в sidecar `<name>.meta.json` (п. 8.2, кнопка «Сохранить описание»).

    overrides позволяет точечно поправить поля (например, rules_hook или arg_template
    для скрипта без меты) без ручного редактирования JSON.
    """
    meta = info.to_meta_dict()
    meta.update(overrides)
    target = sidecar_path(info.path)
    target.parent.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(meta, ensure_ascii=False, indent=2), encoding="utf-8")
    return target
