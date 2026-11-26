"""Тесты app.core.profile_store: CRUD профилей, плоский meta, шапка титула."""

from __future__ import annotations

import datetime
import json
import tempfile
import unittest
from pathlib import Path

from app import constants
from app.core import config
from app.core import profile_store


class DefaultProfileTests(unittest.TestCase):
    def test_year_is_current_not_hardcoded(self) -> None:
        profile = profile_store.default_profile()
        self.assertEqual(profile["work"]["year"], datetime.date.today().year)

    def test_structure_matches_spec_7_2(self) -> None:
        profile = profile_store.default_profile()
        self.assertEqual(profile["name"], "Основной профиль")
        self.assertEqual(profile["person"]["status"], "студент")
        self.assertEqual(profile["work"]["work_type"], "Отчёт по практическому занятию")
        self.assertIn("МИНИСТЕРСТВО НАУКИ", profile["institution"]["ministry"])
        self.assertEqual(profile["extra"], {})
        self.assertTrue(profile_store.default_profile()["institution"]["fsbe"].startswith("ФЕДЕРАЛЬНОЕ"))


class FlatMetaTests(unittest.TestCase):
    def test_flat_keys_and_empty_preserved(self) -> None:
        meta = profile_store.flat_meta(profile_store.default_profile())
        for key in ("fio", "gruppa", "prep", "status", "work_type", "pz_no", "theme",
                    "disc", "city", "year", "ministry", "fsbe", "univ", "kaf", "name"):
            self.assertIn(key, meta)
        self.assertEqual(meta["fio"], "")  # пустые значения сохраняются
        self.assertEqual(meta["year"], datetime.date.today().year)

    def test_extra_keys_flow_into_meta(self) -> None:
        profile = profile_store.default_profile()
        profile["extra"] = {"age_group": "3-4", "duration_min": 15, "week_theme": "Моя семья"}
        meta = profile_store.flat_meta(profile)
        self.assertEqual(meta["age_group"], "3-4")
        self.assertEqual(meta["duration_min"], 15)
        self.assertEqual(meta["week_theme"], "Моя семья")

    def test_missing_year_backfilled_current(self) -> None:
        meta = profile_store.flat_meta({"person": {"fio": "Иванов"}, "work": {"year": ""}})
        self.assertEqual(meta["year"], datetime.date.today().year)


class TitlePreviewTests(unittest.TestCase):
    def test_preview_lines(self) -> None:
        profile = profile_store.default_profile()
        profile["institution"].update({"univ": "МГУ имени Ломоносова", "kaf": "Педагогики"})
        profile["work"].update({"theme": "Водный круговорот", "disc": "Окружающий мир",
                                "city": "Геническ"})
        profile["person"].update({"fio": "Иванов И. И.", "gruppa": "БВ-21",
                                  "prep": "Петрова А. С."})
        preview = profile_store.title_header_preview(profile)
        lines = [ln for ln in preview.split("\n") if ln.strip()]
        self.assertIn("МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ", lines)
        self.assertTrue(lines[1].startswith("ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ"))
        self.assertIn("МГУ имени Ломоносова", lines)
        self.assertIn("Кафедра Педагогики", lines)
        self.assertIn("ОТЧЁТ ПО ПРАКТИЧЕСКОМУ ЗАНЯТИЮ", lines)
        self.assertIn("по дисциплине «Окружающий мир»", lines)
        self.assertIn("Тема: Водный круговорот", lines)
        self.assertTrue(any(ln.startswith("Выполнил(а):") and "Иванов И. И." in ln
                            and "БВ-21" in ln for ln in lines))
        self.assertIn("Преподаватель: Петрова А. С.", lines)
        self.assertEqual(lines[-1], f"Геническ, {datetime.date.today().year}")


class ProfileStoreCrudTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.dirs = config.DataDirs(self.root / "data",
                                    project_dir=self.root, env={}).ensure()
        self.store = profile_store.ProfileStore(self.dirs)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_save_load_roundtrip_utf8(self) -> None:
        profile = profile_store.default_profile()
        profile["person"]["fio"] = "Смирнова Ольга Петровна"
        key = self.store.save(profile, "smirnova")
        loaded = self.store.load(key)
        self.assertEqual(loaded["person"]["fio"], "Смирнова Ольга Петровна")
        raw = (self.dirs.profiles / "smirnova.json").read_text(encoding="utf-8")
        self.assertIn("Смирнова", raw)  # ensure_ascii=False на диске

    def test_profiles_independent(self) -> None:
        a = profile_store.default_profile()
        a["person"]["gruppa"] = "А-1"
        b = profile_store.default_profile()
        b["person"]["gruppa"] = "Б-2"
        self.store.save(a, "a")
        self.store.save(b, "b")
        self.assertEqual(self.store.load("a")["person"]["gruppa"], "А-1")
        self.assertEqual(self.store.load("b")["person"]["gruppa"], "Б-2")
        self.assertEqual({i["key"] for i in self.store.list()}, {"a", "b"})

    def test_save_as_and_delete(self) -> None:
        key = self.store.save_as(profile_store.default_profile(), "копия 1")
        self.assertTrue(key)
        self.assertIn(key, [i["key"] for i in self.store.list()])
        self.store.delete(key)
        with self.assertRaises(FileNotFoundError):
            self.store.load(key)

    def test_load_default_without_files(self) -> None:
        profile = self.store.load("default")
        self.assertEqual(profile["work"]["year"], datetime.date.today().year)

    def test_seed_default_json_used(self) -> None:
        seed = self.root / "profiles"
        seed.mkdir(exist_ok=True)
        legacy = {"schema_version": 0, "name": "Старый", "person": {"fio": "Иванов"}}
        (seed / "default.json").write_text(json.dumps(legacy, ensure_ascii=False),
                                           encoding="utf-8")
        loaded = self.store.load("default")
        self.assertEqual(loaded["person"]["fio"], "Иванов")
        self.assertEqual(loaded["schema_version"], constants.SCHEMA_VERSION)
        # недостающие вложенные ключи дозаполнены дефолтами
        self.assertEqual(loaded["person"]["status"], "студент")

    def test_export_import(self) -> None:
        profile = profile_store.default_profile()
        profile["extra"]["age_group"] = "5-6"
        self.store.save(profile, "vospit")
        dest = self.root / "out.json"
        self.store.export("vospit", dest)
        key = self.store.import_from(dest)
        self.assertEqual(self.store.load(key)["extra"]["age_group"], "5-6")

    def test_import_rejects_non_profile(self) -> None:
        bad = self.root / "bad.json"
        bad.write_text('{"page": {"size": "A4"}}', encoding="utf-8")
        with self.assertRaises(ValueError):
            self.store.import_from(bad)

    def test_corrupt_user_profile_falls_to_default(self) -> None:
        (self.dirs.profiles / "default.json").write_text("{битый", encoding="utf-8")
        profile = self.store.load("default")  # не падает, возвращает дефолт
        self.assertEqual(profile["name"], "Основной профиль")
        self.assertTrue(Path(str(self.dirs.profiles / "default.json") + ".corrupt").exists())

    def test_resolve_by_name(self) -> None:
        profile = profile_store.default_profile()
        profile["name"] = "Иванов И.И."
        self.store.save(profile, "ivanov")
        self.assertEqual(self.store.resolve("иванов и.и.")["person"], profile["person"])


if __name__ == "__main__":
    unittest.main()
