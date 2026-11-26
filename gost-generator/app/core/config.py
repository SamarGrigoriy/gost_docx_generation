"""Загрузка/сохранение JSON-конфигов, определение каталогов данных, состояние приложения.

Почему здесь, а не в UI: все оболочки (GUI, CLI, тесты) работают с одними и теми же
файлами настроек; битый конфиг не должен ронять приложение (ТЗ §12) — отсюда
бэкап `.corrupt` и возврат дефолтов вместо исключения по умолчанию.
"""

from __future__ import annotations

import copy
import json
import os
import sys
from pathlib import Path
from typing import Any

from app import constants

__all__ = [
    "CorruptedConfigError",
    "load_json",
    "save_json",
    "migrate",
    "DataDirs",
    "load_state",
    "save_state",
    "default_state",
]


class CorruptedConfigError(Exception):
    """Файл конфигурации повреждён (битый JSON). Атрибут `.path` — путь к нему."""

    def __init__(self, path: Path | str) -> None:
        self.path = Path(path)
        super().__init__(f"Повреждённый JSON-файл: {self.path}")


def _corrupt_backup_path(path: Path) -> Path:
    return path.with_name(path.name + ".corrupt")


def load_json(
    path: Path | str,
    default: Any = None,
    backup_on_corrupt: bool = True,
) -> Any:
    """Читаем JSON (UTF-8). Файла нет — `default`. Битый JSON — бэкап `.corrupt` и `default`.

    Если `backup_on_corrupt=False`, вместо тихого возврата дефолта бросается
    `CorruptedConfigError` (для мест, где молча подменить конфиг нельзя).
    """
    p = Path(path)
    try:
        raw = p.read_text(encoding="utf-8")
    except FileNotFoundError:
        return default
    except OSError as exc:
        if not backup_on_corrupt:
            raise CorruptedConfigError(p) from exc
        print(f"Не удалось прочитать {p}: {exc}. Используются значения по умолчанию.",
              file=sys.stderr)
        return default
    try:
        return json.loads(raw)
    except (json.JSONDecodeError, ValueError) as exc:
        if not backup_on_corrupt:
            raise CorruptedConfigError(p) from exc
        backup = _corrupt_backup_path(p)
        try:
            # Переносим битый файл, чтобы следующий старт не жаловался повторно.
            os.replace(p, backup)
            print(f"Файл {p} повреждён ({exc}); копия сохранена в {backup}, "
                  f"используются значения по умолчанию.", file=sys.stderr)
        except OSError:
            print(f"Файл {p} повреждён ({exc}); используются значения по умолчанию.",
                  file=sys.stderr)
        return default


def save_json(path: Path | str, data: Any) -> Path:
    """Атомарная запись JSON: UTF-8, ensure_ascii=False, indent=2, каталоги создаются сами.

    Запись идёт в `<file>.tmp` и затем переименовывается (`os.replace`) — при падении
    процесса посередине записи старый конфиг остаётся целым.
    """
    p = Path(path)
    p.parent.mkdir(parents=True, exist_ok=True)
    tmp = p.with_name(p.name + ".tmp")
    try:
        with tmp.open("w", encoding="utf-8", newline="\n") as fh:
            json.dump(data, fh, ensure_ascii=False, indent=2)
            fh.write("\n")
            fh.flush()
            os.fsync(fh.fileno())
        os.replace(tmp, p)
    except OSError as exc:
        try:
            tmp.unlink()
        except OSError:
            pass
        message = f"Не удалось записать {p}: {exc}. Проверьте права на каталог {p.parent}."
        print(message, file=sys.stderr)
        raise OSError(message) from exc
    return p


def _fill_defaults(data: dict, defaults: dict) -> dict:
    """Рекурсивно дозаполняем отсутствующие ключи из `defaults` (свои значения не трогаем)."""
    for key, value in defaults.items():
        if key not in data or data[key] is None:
            data[key] = copy.deepcopy(value)
        elif isinstance(data[key], dict) and isinstance(value, dict):
            _fill_defaults(data[key], value)
    return data


def _migrate_state(data: dict) -> None:
    """Частные миграции состояния между версиями схемы."""
    # Конфиги до появления schema_version: приводим старые имена ключей к текущим.
    renames = {"timeout": "script_timeout_s", "last_files": "recent_files"}
    for old, new in renames.items():
        if old in data:
            data.setdefault(new, data.pop(old))


def migrate(data: dict, kind: str, defaults: dict | None = None) -> dict:
    """Доводим конфиг до `constants.SCHEMA_VERSION`: дозаполняем дефолты, чиним версию.

    `defaults` — эталонный словарь для данного kind; если не передан, для state
    используется статический шаблон (без путей), иначе — пустой.
    """
    if not isinstance(data, dict):
        raise TypeError("migrate ожидает dict")
    out = copy.deepcopy(data)
    try:
        version = int(out.get("schema_version", 0) or 0)
    except (TypeError, ValueError):
        version = 0
    # Сначала чиним устаревшие ключи, иначе дозаполнение дефолтами закроет миграцию.
    if version < constants.SCHEMA_VERSION and kind == "state":
        _migrate_state(out)
    base = defaults
    if base is None and kind == "state":
        base = state_template()
    if base is not None:
        _fill_defaults(out, base)
    if version < constants.SCHEMA_VERSION and kind == "state":
        _migrate_state(out)
    out["schema_version"] = constants.SCHEMA_VERSION
    return out


def _frozen_exe_dir() -> Path | None:
    if getattr(sys, "frozen", False):  # PyInstaller: рядом с exe ищем portable.flag
        return Path(sys.executable).resolve().parent
    return None


def _default_project_dir() -> Path:
    """Каталог репозитория приложения (…/gost-generator), вычисляется от этого файла."""
    return Path(__file__).resolve().parents[2]


class DataDirs:
    """Каталоги данных приложения. Порядок разрешения дома данных (ТЗ §3.3, §19.6):

    1. `override` (явный путь из настроек/CLI),
    2. переменная окружения `GOST_HOME`,
    3. `portable.flag` рядом с exe / в репозитории / в текущей папке,
    4. `gost.json` проекта с ключом `data_dir`,
    5. `%APPDATA%\\GostGenerator` (Windows) / `~/.config/GostGenerator` (Linux).
    """

    def __init__(
        self,
        override: Path | str | None = None,
        *,
        project_dir: Path | str | None = None,
        env: dict[str, str] | None = None,
    ) -> None:
        self.env: dict[str, str] = dict(os.environ if env is None else env)
        self.project_dir = Path(project_dir) if project_dir else _default_project_dir()
        self.is_frozen = bool(getattr(sys, "frozen", False))
        self.home, self._portable = self._resolve(override)
        self.scripts = self.home / constants.DIR_SCRIPTS
        self.presets = self.home / constants.DIR_PRESETS
        self.profiles = self.home / constants.DIR_PROFILES
        self.output = self.home / constants.DIR_OUTPUT
        self.logs = self.home / constants.DIR_LOGS
        self.state_path = self.home / constants.STATE_FILE

    # --- разрешение путей ---

    def _project_data_dir(self) -> Path | None:
        """Ключ `data_dir` в gost.json текущего проекта (репозиторий или cwd)."""
        for base in (self.project_dir, Path.cwd()):
            cfg = load_json(base / constants.PROJECT_CONFIG_FILE, default=None)
            if isinstance(cfg, dict) and cfg.get("data_dir"):
                value = Path(str(cfg["data_dir"]))
                return value if value.is_absolute() else (base / value)
        return None

    def _has_portable_flag(self, base: Path) -> bool:
        try:
            return (base / constants.PORTABLE_FLAG).is_file()
        except OSError:
            return False

    def _fallback_home(self) -> Path:
        if sys.platform == "win32":
            appdata = self.env.get("APPDATA") or self.env.get("USERPROFILE")
            if appdata:
                return Path(appdata) / constants.CONFIG_DIR_NAME
        return Path(self.env.get("HOME", str(Path.home()))) / ".config" / constants.CONFIG_DIR_NAME

    def _resolve(self, override: Path | str | None) -> tuple[Path, bool]:
        if override is not None:
            return Path(override), True
        env_home = self.env.get(constants.ENV_HOME)
        if env_home:
            return Path(env_home), True
        candidates: list[Path] = []
        exe_dir = _frozen_exe_dir()
        if exe_dir is not None:
            candidates.append(exe_dir)
        candidates.extend([self.project_dir, Path.cwd()])
        for base in candidates:
            if self._has_portable_flag(base):
                return base, True
        data_dir = self._project_data_dir()
        if data_dir is not None:
            return data_dir, True
        return self._fallback_home(), False

    # --- публичное API ---

    def bundled_scripts(self) -> Path | None:
        """Каталог штатных скриптов: распаковка exe (`sys._MEIPASS`) или `scripts/` репозитория.

        Нужен оболочкам, чтобы список генераторов не пустел, когда каталог данных
        пользователя ещё не содержит ни одного `.py`.
        """
        meipass = getattr(sys, "_MEIPASS", None)
        bases: list[Path] = [Path(meipass) if meipass else self.project_dir]
        if meipass:
            bases.append(self.project_dir)
        for base in bases:
            candidate = base / constants.DIR_SCRIPTS
            if candidate.is_dir():
                return candidate
        return None

    def is_portable(self) -> bool:
        """True, если данные лежат рядом с приложением/проектом, а не в каталоге ОС."""
        return self._portable

    def ensure(self) -> "DataDirs":
        """Создаёт все каталоги; при отсутствии прав на запись — внятная ошибка в stderr."""
        try:
            for d in (self.home, self.scripts, self.presets, self.profiles,
                      self.output, self.logs):
                d.mkdir(parents=True, exist_ok=True)
        except OSError as exc:
            message = (
                f"Не удалось создать каталог данных {self.home}: {exc}. "
                f"Укажите доступный каталог переменной {constants.ENV_HOME} или задайте "
                f"права на запись, затем перезапустите приложение."
            )
            print(message, file=sys.stderr)
            raise OSError(message) from exc
        return self

    def __repr__(self) -> str:  # удобно для логов
        return f"DataDirs(home={self.home!s}, portable={self._portable})"


# Ключи состояния по ТЗ §7.3 (кроме путей) — статический шаблон для migrate.
STATE_KEYS: dict[str, Any] = {
    "active_preset": constants.DEFAULT_PRESET_KEY,
    "active_profile": constants.DEFAULT_PROFILE_KEY,
    "theme": constants.DEFAULT_THEME,
    "ui_scale": constants.DEFAULT_UI_SCALE,
    "script_timeout_s": constants.DEFAULT_SCRIPT_TIMEOUT_S,
    "interpreter_mode": "internal",
    "interpreter_path": "",
    "use_external_python": False,
    "recent_files": [],
    "window_geometry": {
        "width": constants.WINDOW_WIDTH,
        "height": constants.WINDOW_HEIGHT,
        "x": None,
        "y": None,
    },
    "normalize_after_generation": True,
}


def state_template() -> dict[str, Any]:
    """Дефолты state.json без привязки к конкретным каталогам."""
    return copy.deepcopy(STATE_KEYS)


def default_state(dirs: DataDirs) -> dict[str, Any]:
    """Свежее состояние по умолчанию: активные пресет/профиль, каталоги, тема и т. д."""
    state = state_template()
    state["schema_version"] = constants.SCHEMA_VERSION
    state["scripts_dir"] = str(dirs.scripts)
    state["output_dir"] = str(dirs.output)
    return state


def load_state(dirs: DataDirs) -> dict[str, Any]:
    """Читает state.json; битый/отсутствующий файл не роняет приложение — возвращает дефолты."""
    raw = load_json(dirs.state_path, default=None)
    if not isinstance(raw, dict):
        return default_state(dirs)
    state = migrate(raw, "state")
    state.setdefault("scripts_dir", str(dirs.scripts))
    state.setdefault("output_dir", str(dirs.output))
    return state


def save_state(dirs: DataDirs, state: dict[str, Any]) -> Path:
    """Сохраняет state.json, нормализуя список последних файлов (не более 10, ТЗ §7.3)."""
    out = dict(state)
    recent = out.get("recent_files")
    if isinstance(recent, list):
        out["recent_files"] = [str(item) for item in recent][: constants.MAX_RECENT_FILES]
    out["schema_version"] = constants.SCHEMA_VERSION
    return save_json(dirs.state_path, out)
