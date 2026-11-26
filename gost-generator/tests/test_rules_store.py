"""Тесты app.core.rules_store: CRUD пресетов, защита builtin, валидация, сводка."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from app import constants
from app.core import config
from app.core import rules_store


class BaseStoreCase(unittest.TestCase):
    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.root = Path(self._tmp.name)
        self.dirs = config.DataDirs(self.root / "data",
                                    project_dir=self.root, env={}).ensure()
        self.builtin = self.root / "builtin_presets"
        self.builtin.mkdir()
        gost = json.loads(json.dumps(rules_store.DEFAULT_RULES))
        (self.builtin / "gost_732.json").write_text(
            json.dumps(gost, ensure_ascii=False), encoding="utf-8")
        free = json.loads(json.dumps(rules_store.DEFAULT_RULES))
        free.update({"name": "Свободный", "builtin": True, "strict": False})
        free["text"].update({"font": "Calibri", "size_pt": 12, "line_spacing": 1.15})
        (self.builtin / "free.json").write_text(
            json.dumps(free, ensure_ascii=False), encoding="utf-8")
        self.store = rules_store.RulesStore(self.dirs, builtin_dir=self.builtin)

    def tearDown(self) -> None:
        self._tmp.cleanup()


class ListLoadTests(BaseStoreCase):
    def test_list_builtins(self) -> None:
        keys = {item["key"]: item for item in self.store.list()}
        self.assertIn("gost_732", keys)
        self.assertIn("free", keys)
        self.assertTrue(keys["free"]["builtin"])
        self.assertEqual(keys["free"]["name"], "Свободный")

    def test_load_by_key_merges_defaults(self) -> None:
        rules = self.store.load("gost_732")
        self.assertEqual(rules["text"]["font"], "Times New Roman")
        self.assertEqual(rules["page"]["margin_left_mm"], 30)
        for section in rules_store.REQUIRED_SECTIONS:
            self.assertIn(section, rules)

    def test_load_repo_preset_matches_default_rules(self) -> None:
        repo_preset = Path(constants.__file__).resolve().parents[1] / "presets" / "gost_732.json"
        if repo_preset.is_file():
            data = json.loads(repo_preset.read_text(encoding="utf-8"))
            self.assertEqual(data, rules_store.DEFAULT_RULES)

    def test_load_missing_raises(self) -> None:
        with self.assertRaises(FileNotFoundError):
            self.store.load("нет_такого")

    def test_load_gost_when_dir_lost_uses_builtin_copy(self) -> None:
        store = rules_store.RulesStore(self.dirs, builtin_dir=self.root / "nowhere")
        self.assertEqual(store.load("gost_732")["text"]["size_pt"], 14)


class CrudTests(BaseStoreCase):
    def test_save_roundtrip(self) -> None:
        rules = self.store.load("gost_732")
        rules["text"]["size_pt"] = 13
        key = self.store.save(rules)
        reloaded = self.store.load(key)
        self.assertEqual(reloaded["text"]["size_pt"], 13)
        item = next(i for i in self.store.list() if i["key"] == key)
        self.assertFalse(item["builtin"])

    def test_save_under_builtin_key_forbidden(self) -> None:
        rules = self.store.load("gost_732")
        with self.assertRaises(rules_store.CannotEditBuiltinError):
            self.store.save(rules, "gost_732")
        # встроенный файл на диске не изменился
        raw = json.loads((self.builtin / "gost_732.json").read_text(encoding="utf-8"))
        self.assertEqual(raw["text"]["size_pt"], 14)

    def test_save_as_creates_user_preset_shadowing_builtin(self) -> None:
        rules = self.store.load("gost_732")
        rules["text"]["font"] = "Arial"
        key = self.store.save_as(rules, "gost_732")
        self.assertEqual(key, "gost_732")
        self.assertEqual(self.store.load("gost_732")["text"]["font"], "Arial")
        items = [i for i in self.store.list() if i["key"] == "gost_732"]
        self.assertEqual(len(items), 1)
        self.assertFalse(items[0]["builtin"])
        # удаление тени — снова виден встроенный
        self.store.delete("gost_732")
        self.assertEqual(self.store.load("gost_732")["text"]["font"], "Times New Roman")

    def test_delete_builtin_forbidden(self) -> None:
        with self.assertRaises(rules_store.CannotEditBuiltinError):
            self.store.delete("free")

    def test_export_import(self) -> None:
        dest = self.root / "exported.json"
        self.store.export("free", dest)
        key = self.store.import_from(dest)
        self.assertEqual(self.store.load(key)["text"]["font"], "Calibri")

    def test_resolve_by_name_path_key(self) -> None:
        self.assertEqual(self.store.resolve("Свободный")["strict"], False)
        self.assertEqual(self.store.resolve("gost_732")["text"]["size_pt"], 14)
        path = self.root / "file_preset.json"
        self.store.export("free", path)
        self.assertEqual(self.store.resolve(path)["name"], "Свободный")


class ValidationTests(unittest.TestCase):
    def test_default_rules_valid(self) -> None:
        self.assertEqual(rules_store.validate_rules(rules_store.DEFAULT_RULES), [])

    def test_missing_sections_reported(self) -> None:
        errors = rules_store.validate_rules({"page": {}, "text": {}})
        joined = "\n".join(errors)
        for section in ("headings", "figures", "tables", "toc", "code_blocks"):
            self.assertIn(section, joined)

    def test_ranges_and_enums(self) -> None:
        rules = json.loads(json.dumps(rules_store.DEFAULT_RULES))
        rules["text"]["size_pt"] = 80
        rules["text"]["line_spacing"] = 4
        rules["page"]["margin_left_mm"] = 2
        rules["page"]["orientation"] = "diagonal"
        rules["text"]["alignment"] = "centerleft"
        errors = "\n".join(rules_store.validate_rules(rules))
        self.assertIn("size_pt", errors)
        self.assertIn("line_spacing", errors)
        self.assertIn("margin_left_mm", errors)
        self.assertIn("Ориентация", errors)
        self.assertIn("Выравнивание", errors)

    def test_merge_with_defaults_adds_missing(self) -> None:
        merged = rules_store.merge_with_defaults({"text": {"size_pt": 12}})
        self.assertEqual(merged["text"]["size_pt"], 12)
        self.assertEqual(merged["text"]["font"], "Times New Roman")
        self.assertIn("appendices", merged)


class DescribeTests(BaseStoreCase):
    def test_describe_covers_all_groups(self) -> None:
        text = rules_store.describe_rules(self.store.load("gost_732"))
        for token in ("Страница:", "Основной текст:", "Заголовки:", "Рисунки:",
                      "Таблицы:", "Формулы:", "Примечания:", "Сноски:",
                      "Источники:", "Приложения:", "Содержание:",
                      "Нумерация страниц:", "Объём:", "Блоки кода:"):
            self.assertIn(token, text)
        self.assertIn("Times New Roman 14 пт", text)
        self.assertIn("1,5", text)  # десятичная запятая по-русски
        self.assertIn("строгий режим", text)

    def test_describe_free_preset_mentions_svoobodny(self) -> None:
        text = rules_store.describe_rules(self.store.load("free"))
        self.assertIn("свободный режим", text)
        self.assertIn("Calibri", text)


if __name__ == "__main__":
    unittest.main()
