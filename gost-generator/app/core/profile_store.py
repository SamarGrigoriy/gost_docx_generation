"""Хранилище профилей метаданных (студент/преподаватель/вуз) + плоский `meta` для скриптов.

Профиль — единственный источник данных титульника и подстановок; скрипты-плагины
получают его плоским словарём (ТЗ §7.2), поэтому здесь же `flat_meta` и предпросмотр
шапки титульного листа.
"""

from __future__ import annotations

import copy
import datetime
import re
from pathlib import Path
from typing import Any

from app import constants
from app.core import config


def default_profile() -> dict[str, Any]:
    """Профиль по умолчанию (ТЗ §7.2). Год берётся из системной даты, не хардкод."""
    return {
        "schema_version": constants.SCHEMA_VERSION,
        "name": "Основной профиль",
        "person": {"fio": "", "gruppa": "", "prep": "", "status": "студент"},
        "work": {
            "work_type": "Отчёт по практическому занятию",
            "pz_no": "",
            "theme": "",
            "disc": "",
            "city": "",
            "year": datetime.date.today().year,
        },
        "institution": {
            "ministry": constants.DEFAULT_MINISTRY,
            "fsbe": constants.DEFAULT_FSBE,
            "univ": "",
            "kaf": "",
        },
        "extra": {},
    }


class ProfileStore:
    """CRUD профилей в `dirs.profiles`; импорт/экспорт JSON; затенение seed-копий."""

    def __init__(self, dirs: config.DataDirs, seed_dir: Path | str | None = None) -> None:
        self.dirs = dirs
        self.seed_dir = Path(seed_dir) if seed_dir else Path(dirs.project_dir) / constants.DIR_PROFILES

    def _user_path(self, key: str) -> Path:
        return Path(self.dirs.profiles) / f"{key}.json"

    def _seed_path(self, key: str) -> Path:
        return self.seed_dir / f"{key}.json"

    @staticmethod
    def _slug(text: str) -> str:
        cleaned = re.sub(r"[^0-9A-Za-z_-]+", "_", text.strip()).strip("_")
        return cleaned or "profile"

    def _unique_key(self, base: str) -> str:
        key = self._slug(base)
        candidate = key
        index = 2
        while self._user_path(candidate).exists():
            candidate = f"{key}_{index}"
            index += 1
        return candidate

    def list(self) -> list[dict[str, Any]]:
        """Все профили: элементы {"key","name","path"}; пользовательские затеняют seed."""
        items: dict[str, dict[str, Any]] = {}
        for base, source in ((self.seed_dir, "seed"), (Path(self.dirs.profiles), "user")):
            if not base.is_dir():
                continue
            for path in sorted(base.glob("*.json")):
                if path.name.endswith(".corrupt") or path.name.endswith(".tmp"):
                    continue
                data = config.load_json(path, default={})
                name = data.get("name", path.stem) if isinstance(data, dict) else path.stem
                if source == "seed" and path.stem in items:
                    continue
                items[path.stem] = {"key": path.stem, "name": name, "path": str(path)}
        return list(items.values())

    def load(self, key_or_path: str | Path) -> dict[str, Any]:
        """Профиль по ключу или пути; недостающие ключи дозаполняются дефолтами."""
        p = Path(key_or_path)
        if p.suffix == ".json" and p.is_file():
            raw = config.load_json(p, default=None)
        else:
            key = str(key_or_path)
            user = self._user_path(key)
            seed = self._seed_path(key)
            if user.is_file():
                raw = config.load_json(user, default=None)
            elif seed.is_file():
                raw = config.load_json(seed, default=None)
            elif key == constants.DEFAULT_PROFILE_KEY:
                return default_profile()
            else:
                raise FileNotFoundError(f"Профиль не найден: {key_or_path}")
        if not isinstance(raw, dict):
            if str(key_or_path) == constants.DEFAULT_PROFILE_KEY:
                return default_profile()
            raise FileNotFoundError(f"Профиль не найден или повреждён: {key_or_path}")
        return config.migrate(raw, "profile", default_profile())

    def save(self, profile: dict[str, Any], key: str | None = None) -> str:
        """Сохраняет профиль; без ключа — из `name`/`key`, при коллизии добавляется суффикс."""
        if key is None:
            key = str(profile.get("key") or profile.get("name") or "profile")
            key = self._unique_key(key)
        return self.save_as(profile, key)

    def save_as(self, profile: dict[str, Any], new_key: str) -> str:
        """Записывает профиль под конкретным ключом (перезапись допускается — это CRUD)."""
        if not str(new_key).strip():
            raise ValueError("Ключ профиля не может быть пустым")
        data = copy.deepcopy(dict(profile))
        data["schema_version"] = constants.SCHEMA_VERSION
        data.setdefault("name", str(new_key))
        key = self._slug(str(new_key))
        config.save_json(self._user_path(key), data)
        return key

    def delete(self, key: str) -> None:
        """Удаляет пользовательский профиль (seed-копия в репозитории не трогается)."""
        user = self._user_path(key)
        if user.is_file():
            user.unlink()
            return
        if self._seed_path(key).is_file():
            raise PermissionError(
                f"Профиль «{key}» поставляется с приложением; удалите только "
                f"пользовательскую копию в каталоге данных.")
        raise FileNotFoundError(f"Профиль не найден: {key}")

    def export(self, key: str, dest: Path | str) -> Path:
        return config.save_json(Path(dest), self.load(key))

    def import_from(self, src: Path | str) -> str:
        """Импорт профиля из внешнего JSON под уникальным пользовательским ключом."""
        raw = config.load_json(Path(src), default=None)
        if not isinstance(raw, dict) or not (
            {"person", "work", "institution", "extra"} & set(raw)
        ):
            raise ValueError(f"Файл не является профилем метаданных: {src}")
        return self.save(raw, self._unique_key(Path(src).stem))

    def resolve(self, name_or_key: str) -> dict[str, Any]:
        """Ключ или название профиля → полный словарь профиля."""
        try:
            return self.load(name_or_key)
        except FileNotFoundError:
            wanted = str(name_or_key).strip().casefold()
            for item in self.list():
                if wanted == str(item["name"]).casefold():
                    return self.load(item["key"])
            raise FileNotFoundError(f"Профиль не найден: {name_or_key}")


def flat_meta(profile: dict[str, Any]) -> dict[str, Any]:
    """Плоский словарь `meta` для скриптов: person+work+institution+extra + «name».

    Пустые значения сохраняются как "" (скрипт должен различать «не заполнено»),
    отсутствующий/пустой `year` подставляется текущим системным годом.
    """
    meta: dict[str, Any] = {}
    defaults = default_profile()
    for section in ("person", "work", "institution"):
        for key, value in defaults[section].items():
            meta[key] = value
    for section in ("person", "work", "institution", "extra"):
        data = profile.get(section)
        if isinstance(data, dict):
            for key, value in data.items():
                meta[key] = "" if value is None else value
    meta["name"] = profile.get("name", "")
    if meta.get("year") in (None, "", 0):
        meta["year"] = datetime.date.today().year
    return meta


def title_header_preview(profile: dict[str, Any]) -> str:
    """Текст шапки титульного листа — строками, как он будет выглядеть в документе."""
    person = profile.get("person") or {}
    work = profile.get("work") or {}
    inst = profile.get("institution") or {}
    year = work.get("year") or datetime.date.today().year
    lines: list[str] = []
    for key in ("ministry", "fsbe", "univ", "kaf"):
        value = str(inst.get(key, "")).strip()
        if not value:
            continue
        lines.append(value if key != "kaf" else f"Кафедра {value}")
    work_type = str(work.get("work_type", "")).strip()
    if work_type:
        lines.append("")
        lines.append(work_type.upper())
    disc = str(work.get("disc", "")).strip()
    if disc:
        lines.append(f"по дисциплине «{disc}»")
    theme = str(work.get("theme", "")).strip()
    if theme:
        lines.append("")
        lines.append(f"Тема: {theme}")
    pz_no = str(work.get("pz_no", "")).strip()
    if pz_no:
        lines.append(f"Занятие № {pz_no}")
    fio = str(person.get("fio", "")).strip()
    if fio:
        status = str(person.get("status", "")).strip()
        gruppa = str(person.get("gruppa", "")).strip()
        suffix = f", группа {gruppa}" if gruppa else ""
        lines.append("")
        lines.append(f"Выполнил(а): {status} {fio}{suffix}".replace("  ", " "))
    prep = str(person.get("prep", "")).strip()
    if prep:
        lines.append(f"Преподаватель: {prep}")
    city = str(work.get("city", "")).strip()
    lines.append("")
    lines.append(f"{city}, {year}" if city else str(year))
    return "\n".join(lines)
