"""Тесты app.core.config: JSON-хранилка, миграции, каталоги данных, состояние."""

from __future__ import annotations

import io
import json
import sys
import tempfile
import unittest
from contextlib import redirect_stderr
from pathlib import Path

from app import constants
from app.core import config


class JsonRoundtripTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_save_creates_dirs_and_utf8(self) -> None:
        target = self.tmp / "a" / "b" / "cfg.json"
        config.save_json(target, {"имя": "Таблица 1 – отдел", "n": 2})
        raw = target.read_text(encoding="utf-8")
        self.assertIn("имя", raw)  # ensure_ascii=False
        self.assertIn('\n  "n": 2', raw)  # indent=2
        self.assertEqual(json.loads(raw)["n"], 2)

    def test_save_leaves_no_tmp_file(self) -> None:
        target = config.save_json(self.tmp / "cfg.json", {"x": 1})
        self.assertFalse(Path(str(target) + ".tmp").exists())

    def test_load_missing_returns_default(self) -> None:
        self.assertEqual(config.load_json(self.tmp / "nope.json", default={"d": 1}),
                         {"d": 1})
        self.assertIsNone(config.load_json(self.tmp / "nope.json"))

    def test_corrupt_json_backed_up_and_default_returned(self) -> None:
        broken = self.tmp / "state.json"
        broken.write_text("{это не json", encoding="utf-8")
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            value = config.load_json(broken, default={"ok": True})
        self.assertEqual(value, {"ok": True})
        self.assertFalse(broken.exists())
        backup = self.tmp / "state.json.corrupt"
        self.assertTrue(backup.is_file())
        self.assertIn("не json", backup.read_text(encoding="utf-8"))
        self.assertIn("повреждён", stderr.getvalue())

    def test_corrupt_strict_raises(self) -> None:
        broken = self.tmp / "rules.json"
        broken.write_text("[1,2", encoding="utf-8")
        with self.assertRaises(config.CorruptedConfigError) as ctx:
            config.load_json(broken, default=[], backup_on_corrupt=False)
        self.assertEqual(Path(ctx.exception.path), broken)


class MigrateTests(unittest.TestCase):
    def test_state_without_version_is_migrated(self) -> None:
        legacy = {"active_preset": "free"}  # старый конфиг без schema_version
        out = config.migrate(legacy, "state")
        self.assertEqual(out["schema_version"], constants.SCHEMA_VERSION)
        self.assertEqual(out["active_preset"], "free")  # своё значение не перезабито
        self.assertEqual(out["script_timeout_s"], constants.DEFAULT_SCRIPT_TIMEOUT_S)
        self.assertEqual(out["ui_scale"], constants.DEFAULT_UI_SCALE)

    def test_legacy_key_renames(self) -> None:
        out = config.migrate({"timeout": 30, "last_files": ["a.docx"]}, "state")
        self.assertEqual(out["script_timeout_s"], 30)
        self.assertEqual(out["recent_files"], ["a.docx"])

    def test_custom_defaults_filled_deeply(self) -> None:
        out = config.migrate({"person": {"fio": "Иван"}}, "profile",
                             defaults=config.state_template() | {
                                 "person": {"fio": "", "gruppa": "БВ-1"},
                             })
        self.assertEqual(out["person"], {"fio": "Иван", "gruppa": "БВ-1"})


class DataDirsTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_override_wins_and_portable(self) -> None:
        dirs = config.DataDirs(self.tmp / "data", project_dir=self.tmp, env={})
        self.assertEqual(dirs.home, self.tmp / "data")
        self.assertTrue(dirs.is_portable())
        dirs.ensure()
        for sub in (dirs.scripts, dirs.presets, dirs.profiles, dirs.output, dirs.logs):
            self.assertTrue(sub.is_dir())
        self.assertEqual(dirs.state_path, dirs.home / "state.json")

    def test_env_gost_home_top_priority(self) -> None:
        (self.tmp / constants.PORTABLE_FLAG).write_text("", encoding="utf-8")
        env = {constants.ENV_HOME: str(self.tmp / "fromenv")}
        dirs = config.DataDirs(project_dir=self.tmp, env=env)
        self.assertEqual(dirs.home, self.tmp / "fromenv")

    def test_portable_flag_next_to_project(self) -> None:
        (self.tmp / constants.PORTABLE_FLAG).write_text("", encoding="utf-8")
        dirs = config.DataDirs(project_dir=self.tmp, env={})
        self.assertEqual(dirs.home, self.tmp)
        self.assertTrue(dirs.is_portable())

    def test_gost_json_data_dir(self) -> None:
        (self.tmp / constants.PROJECT_CONFIG_FILE).write_text(
            json.dumps({"data_dir": "gostdata"}), encoding="utf-8")
        dirs = config.DataDirs(project_dir=self.tmp, env={})
        self.assertEqual(dirs.home, self.tmp / "gostdata")

    def test_fallback_config_dir(self) -> None:
        home = self.tmp / "userhome"
        dirs = config.DataDirs(project_dir=self.tmp / "empty",
                               env={"HOME": str(home), "APPDATA": "", "USERPROFILE": ""})
        self.assertTrue(str(dirs.home).endswith(
            str(Path("GostGenerator")) if sys.platform == "win32"
            else str(Path(".config") / "GostGenerator")))
        self.assertFalse(dirs.is_portable())

    def test_ensure_unwritable_reports_clear_error(self) -> None:
        blocker = self.tmp / "file.txt"
        blocker.write_text("x", encoding="utf-8")
        dirs = config.DataDirs(blocker / "sub", project_dir=self.tmp, env={})
        stderr = io.StringIO()
        with redirect_stderr(stderr), self.assertRaises(OSError):
            dirs.ensure()
        self.assertIn("Не удалось создать каталог данных", stderr.getvalue())


class StateTests(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.dirs = config.DataDirs(Path(self._tmp.name) / "data",
                                    project_dir=Path(self._tmp.name), env={})
        self.dirs.ensure()

    def tearDown(self) -> None:
        self._tmp.cleanup()

    def test_default_state_fields(self) -> None:
        state = config.default_state(self.dirs)
        self.assertEqual(state["script_timeout_s"], 120)
        self.assertEqual(state["active_preset"], constants.DEFAULT_PRESET_KEY)
        self.assertEqual(state["active_profile"], constants.DEFAULT_PROFILE_KEY)
        self.assertEqual(state["theme"], "system")
        self.assertEqual(state["ui_scale"], 100)
        self.assertEqual(state["recent_files"], [])
        self.assertTrue(state["normalize_after_generation"])
        self.assertEqual(state["window_geometry"]["width"], constants.WINDOW_WIDTH)
        self.assertEqual(state["scripts_dir"], str(self.dirs.scripts))

    def test_save_load_roundtrip(self) -> None:
        state = config.default_state(self.dirs)
        state["theme"] = "dark"
        state["recent_files"] = ["/x/отчёт 1.docx", "/x/отчёт 2.docx"]
        config.save_state(self.dirs, state)
        loaded = config.load_state(self.dirs)
        self.assertEqual(loaded["theme"], "dark")
        self.assertEqual(loaded["recent_files"], state["recent_files"])
        # кириллица в файле должна остаться читаемой
        raw = self.dirs.state_path.read_text(encoding="utf-8")
        self.assertIn("отчёт 1.docx", raw)

    def test_recent_files_capped_at_ten(self) -> None:
        state = config.default_state(self.dirs)
        state["recent_files"] = [f"ф{i}.docx" for i in range(25)]
        config.save_state(self.dirs, state)
        loaded = config.load_state(self.dirs)
        self.assertEqual(len(loaded["recent_files"]), constants.MAX_RECENT_FILES)
        self.assertEqual(loaded["recent_files"][0], "ф0.docx")

    def test_load_state_corrupt_falls_back(self) -> None:
        self.dirs.state_path.write_text("{битый", encoding="utf-8")
        stderr = io.StringIO()
        with redirect_stderr(stderr):
            state = config.load_state(self.dirs)
        self.assertEqual(state, config.default_state(self.dirs))
        self.assertTrue(Path(str(self.dirs.state_path) + ".corrupt").exists())


if __name__ == "__main__":
    unittest.main()
