"""Тонкие адаптеры CLI к ядру app.core и утилиты контрактного вывода (п. 19.4-19.6).

Бизнес-логика (инспектор, нормализатор, раннер, хранилища) живёт только в app.core;
этот модуль НЕ дублирует её, а лишь:
- разрешает каталоги данных с приоритетами п. 19.6 (с деградированным путём,
  когда app.core.config ещё не импортируется);
- читает переменные окружения и gost.json проекта;
- формирует единый JSON-конверт и коды возврата п. 19.5.

Модуль не импортирует GUI (app.ui) и сторонние пакеты — только stdlib и app.constants.
"""

from __future__ import annotations

import dataclasses
import importlib
import importlib.util
import json
import os
import sys
from pathlib import Path
from typing import Any

from app import constants

# --- Ошибки CLI -------------------------------------------------------------


def configure_console() -> None:
    """UTF-8 для stdout/stderr: кириллица в консоли Windows (cp866/cp1251) иначе роняет
    приложение `UnicodeEncodeError` до печати JSON-конверта (п. 19.4)."""
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):  # pragma: no cover - exotic streams
            pass
    _windows_console_utf8()


def _windows_console_utf8() -> None:
    """Кодовая страница 65001 нужна самой консоли: UTF-8 в cmd.exe на cp866 читается
    как «кракозябры». В pipe не вмешиваемся — там приёмник получает ровно UTF-8."""
    if sys.platform != "win32":
        return
    try:
        if not sys.stdout.isatty():
            return
        import ctypes

        ctypes.windll.kernel32.SetConsoleOutputCP(65001)
        ctypes.windll.kernel32.SetConsoleCP(65001)
    except Exception:  # pragma: no cover - отсутствие kernel32/atty
        pass


class CliError(Exception):
    """Ошибка с машинным кодом для конверта JSON (п. 19.4) и кода возврата (п. 19.5)."""

    def __init__(self, code: str, message: str, path: str | None = None) -> None:
        super().__init__(message)
        self.code = code
        self.message = message
        self.path = path


# Маппинг кодов ошибок CLI в коды возврата (п. 19.5).
ERROR_EXIT: dict[str, int] = {
    "VIOLATIONS": constants.EXIT_VIOLATIONS,
    "USAGE": constants.EXIT_USAGE,
    "INTERNAL_ERROR": constants.EXIT_USAGE,
    "FILE_NOT_FOUND": constants.EXIT_NOT_FOUND,
    "DIR_NOT_FOUND": constants.EXIT_NOT_FOUND,
    "NOT_FOUND": constants.EXIT_NOT_FOUND,
    "SCRIPT_NOT_FOUND": constants.EXIT_NOT_FOUND,
    "RULES_NOT_FOUND": constants.EXIT_NOT_FOUND,
    "PROFILE_NOT_FOUND": constants.EXIT_NOT_FOUND,
    "SCRIPT_FAILED": constants.EXIT_SCRIPT_FAILED,
    "TIMEOUT": constants.EXIT_TIMEOUT,
    "BAD_CONFIG": constants.EXIT_BAD_CONFIG,
    "INVALID_JSON": constants.EXIT_BAD_CONFIG,
    "JOB_INVALID": constants.EXIT_BAD_CONFIG,
    "CORE_NOT_AVAILABLE": constants.EXIT_BAD_CONFIG,
    "NORMALIZE_FAILED": constants.EXIT_NORMALIZE_FAILED,
}

# Приоритет: если ошибок несколько, возвращается код с наибольшим приоритетом.
_EXIT_PRIORITY: tuple[str, ...] = (
    "USAGE",
    "INTERNAL_ERROR",
    "CORE_NOT_AVAILABLE",
    "BAD_CONFIG",
    "INVALID_JSON",
    "JOB_INVALID",
    "NORMALIZE_FAILED",
    "TIMEOUT",
    "SCRIPT_FAILED",
    "FILE_NOT_FOUND",
    "DIR_NOT_FOUND",
    "NOT_FOUND",
    "SCRIPT_NOT_FOUND",
    "RULES_NOT_FOUND",
    "PROFILE_NOT_FOUND",
    "VIOLATIONS",
)


def make_error(code: str, message: str, path: str | None = None) -> dict[str, Any]:
    """Собрать элемент errors[] единого конверта."""
    return {"code": code, "message": message, "path": path}


def exit_code_for(errors: list[dict[str, Any]]) -> int:
    """Код возврата по списку ошибок конверта (п. 19.5)."""
    codes = {str(e.get("code", "")) for e in errors}
    for code in _EXIT_PRIORITY:
        if code in codes:
            return ERROR_EXIT[code]
    return constants.EXIT_USAGE if codes else constants.EXIT_OK


def emit_json(command: str, ok: bool, data: Any,
              errors: list[dict[str, Any]], warnings: list[Any],
              stream: Any = None) -> None:
    """Записать единый JSON-конверт (п. 19.4) в stdout.

    :param command: имя подкоманды.
    :param ok: успех операции.
    :param data: полезная нагрузка команды.
    :param errors: список ошибок вида {code, message, path}.
    :param warnings: список предупреждений.
    :param stream: целевой поток (по умолчанию stdout).
    """
    envelope = {
        "schema_version": constants.SCHEMA_VERSION,
        "ok": ok,
        "command": command,
        "data": data,
        "errors": errors,
        "warnings": warnings,
    }
    out = stream if stream is not None else sys.stdout
    out.write(json.dumps(envelope, ensure_ascii=False, indent=2, default=str) + "\n")
    out.flush()


def log(msg: str) -> None:
    """Строка лога/прогресса — всегда в stderr, никогда в stdout (п. 19.4)."""
    sys.stderr.write(msg.rstrip("\n") + "\n")


class Result:
    """Результат команды: части единого конверта (п. 19.4) + текст для stdout."""

    def __init__(self, data: Any = None, errors: list[dict[str, Any]] | None = None,
                 warnings: list[Any] | None = None, ok: bool | None = None,
                 text: str | None = None) -> None:
        self.errors = errors or []
        self.warnings = warnings or []
        self.data = data if data is not None else {}
        self.ok = (not self.errors) if ok is None else ok
        self.text = text


# --- Пути и настройки (п. 19.6) ---------------------------------------------


def repo_root() -> Path:
    """Корень репозитория/установки (каталог, где лежит пакет app)."""
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    return Path(__file__).resolve().parent.parent


def default_home() -> Path:
    """Домашний каталог данных по умолчанию (XDG/APPDATA)."""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or str(Path.home() / "AppData" / "Roaming")
        return Path(base) / constants.CONFIG_DIR_NAME
    return Path.home() / ".config" / constants.CONFIG_DIR_NAME


def read_json_file(path: Path) -> Any:
    """Прочитать JSON; битый файл — ошибка INVALID_JSON (без падения, п. 12)."""
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (OSError, ValueError) as exc:
        raise CliError("INVALID_JSON", f"Не удалось прочитать JSON: {exc}", str(path)) from exc


def load_project_config(cwd: Path | None = None) -> dict[str, Any]:
    """Читает gost.json текущего каталога (слой приоритетов п. 19.6)."""
    base = cwd if cwd is not None else Path.cwd()
    cfg_path = base / constants.PROJECT_CONFIG_FILE
    if not cfg_path.exists():
        return {}
    data = read_json_file(cfg_path)
    return data if isinstance(data, dict) else {}


class FallbackDirs:
    """Деградированное разрешение каталогов, если app.core.config недоступен.

    Повторяет только порядок выбора корня данных (п. 19.6), не логику миграции
    и чтения конфигов: GOST_HOME/override > portable.flag у репозитория >
    home из gost.json > ~/.config/GostGenerator (или %APPDATA%).
    """

    def __init__(self, override: Path | str | None = None) -> None:
        self.project_dir = repo_root()
        home: Path | None = None
        if override:
            home = Path(override)
        elif os.environ.get(constants.ENV_HOME):
            home = Path(os.environ[constants.ENV_HOME])
        elif (self.project_dir / constants.PORTABLE_FLAG).exists():
            home = self.project_dir
        else:
            cfg_home = load_project_config().get("home")
            home = Path(cfg_home) if cfg_home else default_home()
        self.home = home
        self.scripts = self._sub(constants.DIR_SCRIPTS)
        self.presets = self._sub(constants.DIR_PRESETS)
        self.profiles = self._sub(constants.DIR_PROFILES)
        self.output = self._sub(constants.DIR_OUTPUT)
        self.logs = self._sub(constants.DIR_LOGS)
        self.state_path = self.home / constants.STATE_FILE

    def _sub(self, name: str) -> Path:
        """Каталог внутри home; если пуст — используем репозиторий (режим разработки)."""
        local = self.home / name
        if not local.exists() and (self.project_dir / name).exists():
            return self.project_dir / name
        return local

    def ensure(self) -> None:
        """Создать отсутствующие каталоги данных (п. 12)."""
        for d in (self.home, self.scripts, self.presets, self.profiles, self.output, self.logs):
            d.mkdir(parents=True, exist_ok=True)

    def is_portable(self) -> bool:
        """Портативный режим: данные лежат рядом с репозиторием."""
        return self.home == self.project_dir or bool(self.project_dir.joinpath(constants.PORTABLE_FLAG).exists())


def resolve_dirs(override: Path | str | None = None) -> Any:
    """Каталоги данных: делегируем app.core.config.DataDirs, иначе FallbackDirs."""
    mod = core_module("config")
    if mod is not None and hasattr(mod, "DataDirs"):
        try:
            dirs = mod.DataDirs(override)
        except CliError:
            raise
        except Exception as exc:  # битый конфиг — понятная ошибка, не traceback
            raise CliError("BAD_CONFIG", f"Не удалось разобрать каталоги данных: {exc}") from exc
        return dirs
    return FallbackDirs(override)


def load_state(dirs: Any) -> dict[str, Any]:
    """Состояние приложения (активные пресет/профиль): через core, иначе прямое чтение."""
    mod = core_module("config")
    if mod is not None and hasattr(mod, "load_state"):
        try:
            return dict(mod.load_state(dirs))
        except Exception as exc:
            raise CliError("BAD_CONFIG", f"Битый state.json: {exc}") from exc
    state_path = getattr(dirs, "state_path", Path(dirs.home) / constants.STATE_FILE)
    if state_path.exists():
        try:
            data = read_json_file(state_path)
            return data if isinstance(data, dict) else {}
        except CliError:
            return {}
    return {}


def save_state(dirs: Any, state: dict[str, Any]) -> None:
    """Сохранить состояние: через core, иначе атомарная запись JSON."""
    mod = core_module("config")
    if mod is not None and hasattr(mod, "save_state"):
        mod.save_state(dirs, state)
        return
    state_path = Path(getattr(dirs, "state_path", Path(dirs.home) / constants.STATE_FILE))
    state_path.parent.mkdir(parents=True, exist_ok=True)
    tmp = state_path.with_suffix(".json.tmp")
    tmp.write_text(json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8")
    tmp.replace(state_path)


def setting(flag_value: Any, env_name: str, cfg: dict[str, Any], cfg_key: str,
            state: dict[str, Any], state_key: str, default: Any) -> Any:
    """Единая цепочка приоритетов п. 19.6: флаг > env > gost.json > состояние > дефолт."""
    if flag_value:
        return flag_value
    env_value = os.environ.get(env_name)
    if env_value:
        return env_value
    for source, key in ((cfg, cfg_key), (state, state_key)):
        value = source.get(key) if isinstance(source, dict) else None
        if value:
            return value
    return default


def output_dir(dirs: Any, cfg: dict[str, Any], state: dict[str, Any],
               flag_value: Path | str | None = None) -> Path:
    """Каталог вывода с приоритетами п. 19.6 (GOST_OUTPUT и ключ output)."""
    raw = setting(flag_value, constants.ENV_OUTPUT, cfg, "output", state, "output_dir", None)
    return Path(raw) if raw else Path(dirs.output)


# --- Доступ к ядру ----------------------------------------------------------


def core_module(name: str) -> Any:
    """Импортировать app.core.<name>; None, если модуль ещё не собран/не импортируется."""
    try:
        return importlib.import_module(f"app.core.{name}")
    except Exception:
        return None


CORE_MODULES: tuple[str, ...] = (
    "config", "rules_store", "profile_store", "plugin_loader",
    "runner", "job", "docx_inspector", "docx_normalizer", "report",
)


def module_exists(name: str) -> bool:
    """Есть ли модуль в окружении (без импорта — важно для GUI-проверок doctor)."""
    try:
        return importlib.util.find_spec(name) is not None
    except (ImportError, ValueError):
        return False


def gui_status() -> dict[str, bool]:
    """Наличие GUI-библиотек в окружении (только поиск спецификаций, без импорта)."""
    return {
        "tkinter": module_exists("tkinter"),
        "customtkinter": module_exists("customtkinter"),
        "pillow": module_exists("PIL"),
    }


def require_core(name: str, *attrs: str) -> Any:
    """Требовать модуль ядра с нужными атрибутами, иначе CliError CORE_NOT_AVAILABLE."""
    mod = core_module(name)
    missing = [a for a in attrs if mod is None or not hasattr(mod, a)]
    if missing:
        raise CliError(
            "CORE_NOT_AVAILABLE",
            f"Модуль app.core.{name} недоступен (отсутствует: {', '.join(missing)}). "
            f"Проверьте `gost doctor --json`.",
        )
    return mod


def rules_dirs(dirs: Any) -> Path:
    """Каталог пользовательских пресетов (или home, если ещё не создан)."""
    return Path(getattr(dirs, "presets", Path(dirs.home) / constants.DIR_PRESETS))


def builtin_presets_dir() -> Path:
    """Каталог встроенных пресетов в репозитории/сборке."""
    return repo_root() / constants.DIR_PRESETS


def list_rules(dirs: Any) -> list[dict[str, Any]]:
    """Список пресетов: через RulesStore, иначе — ручной обход JSON-файлов."""
    mod = core_module("rules_store")
    if mod is not None and hasattr(mod, "RulesStore"):
        store = mod.RulesStore(dirs, builtin_dir=builtin_presets_dir())
        return list(store.list())
    items: list[dict[str, Any]] = []
    seen: set[str] = set()
    for base, builtin in ((builtin_presets_dir(), True), (rules_dirs(dirs), False)):
        if not base.exists():
            continue
        for f in sorted(base.glob("*.json")):
            if f.name in seen:
                continue
            seen.add(f.name)
            items.append({"key": f.stem, "name": f.stem, "builtin": builtin, "path": str(f)})
    return items


def resolve_rules(dirs: Any, name_or_path: str) -> dict[str, Any]:
    """Загрузить пресет по имени или пути: через RulesStore, иначе прямой поиск файла."""
    mod = core_module("rules_store")
    if mod is not None and hasattr(mod, "RulesStore"):
        store = mod.RulesStore(dirs, builtin_dir=builtin_presets_dir())
        try:
            return dict(store.resolve(name_or_path))
        except FileNotFoundError as exc:
            raise CliError("RULES_NOT_FOUND", f"Пресет не найден: {name_or_path}", str(exc)) from exc
        except CliError:
            raise
        except ValueError as exc:
            raise CliError("BAD_CONFIG", f"Битый пресет {name_or_path}: {exc}") from exc
        except Exception as exc:
            raise CliError("BAD_CONFIG", f"Не удалось загрузить пресет {name_or_path}: {exc}") from exc
    candidate = Path(name_or_path)
    if candidate.suffix == ".json" and candidate.exists():
        data = read_json_file(candidate)
        if not isinstance(data, dict):
            raise CliError("BAD_CONFIG", f"Пресет {name_or_path} — не объект JSON")
        return data
    for base in (rules_dirs(dirs), builtin_presets_dir()):
        f = base / f"{name_or_path}.json"
        if f.exists():
            data = read_json_file(f)
            if not isinstance(data, dict):
                raise CliError("BAD_CONFIG", f"Пресет {name_or_path} — не объект JSON")
            return data
    raise CliError("RULES_NOT_FOUND", f"Пресет не найден: {name_or_path}")


def profiles_list(dirs: Any) -> list[dict[str, Any]]:
    """Список профилей: через ProfileStore, иначе обход JSON в каталоге profiles."""
    mod = core_module("profile_store")
    if mod is not None and hasattr(mod, "ProfileStore"):
        store = mod.ProfileStore(dirs, seed_dir=repo_root() / constants.DIR_PROFILES)
        return list(store.list())
    base = Path(dirs.profiles)
    items: list[dict[str, Any]] = []
    for f in sorted(base.glob("*.json")) if base.exists() else []:
        items.append({"key": f.stem, "name": f.stem, "path": str(f)})
    if not items:
        repo_profiles = repo_root() / constants.DIR_PROFILES
        for f in sorted(repo_profiles.glob("*.json")) if repo_profiles.exists() else []:
            items.append({"key": f.stem, "name": f.stem, "path": str(f)})
    return items


def profile_store(dirs: Any) -> Any:
    """ProfileStore или None, если модуль не собран."""
    mod = core_module("profile_store")
    if mod is not None and hasattr(mod, "ProfileStore"):
        return mod.ProfileStore(dirs, seed_dir=repo_root() / constants.DIR_PROFILES)
    return None


def flat_meta(profile: dict[str, Any]) -> dict[str, Any]:
    """Плоский словарь подстановок для скриптов: через core.profile_store."""
    mod = core_module("profile_store")
    fn = getattr(mod, "flat_meta", None) if mod is not None else None
    if fn is not None:
        return dict(fn(profile))
    return dict(profile)


def load_profile(dirs: Any, key: str) -> dict[str, Any]:
    """Загрузить профиль по ключу/пути с деградированным чтением файла."""
    store = profile_store(dirs)
    if store is not None:
        try:
            return dict(store.load(key))
        except FileNotFoundError as exc:
            raise CliError("PROFILE_NOT_FOUND", f"Профиль не найден: {key}", str(exc)) from exc
        except CliError:
            raise
        except Exception as exc:
            raise CliError("BAD_CONFIG", f"Не удалось загрузить профиль {key}: {exc}") from exc
    candidate = Path(key)
    if candidate.exists():
        return dict(read_json_file(candidate))
    for base in (Path(dirs.profiles), repo_root() / constants.DIR_PROFILES):
        f = base / f"{key}.json"
        if f.exists():
            return dict(read_json_file(f))
    raise CliError("PROFILE_NOT_FOUND", f"Профиль не найден: {key}")


def scripts_search_dirs(dirs: Any) -> list[Path]:
    """Каталоги, где ищутся скрипты-плагины: данные пользователя, затем репозиторий."""
    return [Path(dirs.scripts), repo_root() / constants.DIR_SCRIPTS]


def find_script(dirs: Any, name: str) -> Path:
    """Найти файл скрипта по имени (с .py / без) или пути."""
    p = Path(name)
    if p.exists():
        return p
    for base in scripts_search_dirs(dirs):
        for cand in (base / name, base / f"{name}.py"):
            if cand.exists():
                return cand
    raise CliError("SCRIPT_NOT_FOUND", f"Скрипт не найден: {name}")


def list_schemas() -> dict[str, Path]:
    """Найти настоящие schemas/*.json на диске (инсталляция/репо): имя -> путь."""
    from app import cli_init  # локальный импорт: текст/схемы живут в cli_init

    candidates = [repo_root() / constants.DIR_SCHEMAS, Path.cwd() / constants.DIR_SCHEMAS]
    found: dict[str, Path] = {}
    for base in candidates:
        if not base.exists():
            continue
        for name in cli_init.SCHEMA_NAMES:
            for fname in (f"{name}.json", f"{name}.schema.json"):
                f = base / fname
                if f.exists():
                    found.setdefault(name, f)
    return found


def get_schema(name: str) -> dict[str, Any]:
    """Схема по имени: файл с диска (если есть), иначе компактная встроенная копия."""
    from app import cli_init

    if name not in cli_init.SCHEMA_NAMES:
        raise CliError("USAGE", f"Неизвестная схема: {name}. "
                                f"Доступны: {', '.join(cli_init.SCHEMA_NAMES)}")
    path = list_schemas().get(name)
    if path is not None:
        data = read_json_file(path)
        if isinstance(data, dict):
            return data
        raise CliError("BAD_CONFIG", f"Файл схемы {path} — не объект JSON")
    return cli_init.FALLBACK_SCHEMAS[name]


# --- Контекст настроек команды (п. 19.6) ---------------------------------

def make_ctx(args: argparse.Namespace) -> dict[str, Any]:
    """Контекст настроек с приоритетами п. 19.6: dirs, gost.json, состояние, ключи."""
    dirs = resolve_dirs()
    cfg = load_project_config()
    state = load_state(dirs)
    return {
        "dirs": dirs,
        "cfg": cfg,
        "state": state,
        "rules_key": setting(
            getattr(args, "rules", None), constants.ENV_RULES, cfg, "rules",
            state, "active_preset", constants.DEFAULT_PRESET_KEY),
        "profile_key": setting(
            getattr(args, "profile", None), constants.ENV_PROFILE, cfg, "profile",
            state, "active_profile", constants.DEFAULT_PROFILE_KEY),
    }


def rules_of(ctx: dict[str, Any]) -> dict[str, Any]:
    """Загрузить активный (или указанный) пресет правил."""
    return resolve_rules(ctx["dirs"], ctx["rules_key"])




# --- Конвертация PluginInfo ----------------------------------------------

def plugin_to_dict(info: Any) -> dict[str, Any]:
    """PluginInfo → словарь для JSON-конверта (Path → str)."""
    if dataclasses.is_dataclass(info) and not isinstance(info, type):
        raw = dataclasses.asdict(info)
    elif isinstance(info, dict):
        raw = dict(info)
    else:
        return {"repr": str(info)}
    return {k: (str(v) if isinstance(v, Path) else v) for k, v in raw.items()}
