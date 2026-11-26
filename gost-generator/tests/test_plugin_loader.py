"""Тесты plugin_loader: ast-парсинг SCRIPT_META без исполнения, sidecar, build_argv, validate."""

from __future__ import annotations

import json
import tempfile
import unittest
from pathlib import Path

from app import constants
from app.core import plugin_loader as pl

META_OK = {
    "name": "Технологическая карта ООД",
    "description": "Генерация тех.карты занятия",
    "category": "Педагогика",
    "version": "1.0",
    "input": "none",
    "output_ext": ".docx",
    "arg_template": ["--out", "{output}", "--meta", "{meta_json}",
                     "--rules", "{rules_json}", "--profile", "{profile_json}",
                     "--theme", "{theme}"],
    "rules_hook": "postprocess",
    "params": [
        {"id": "theme", "label": "Тема", "type": "text", "default": "Моя семья"},
        {"id": "duration_min", "label": "Длительность", "type": "int", "default": 15},
    ],
}


def source_with_meta(meta: dict | None = None, tail: str = "print('executed')") -> str:
    """Сгенерировать исходник скрипта с SCRIPT_META-литералом и хвостом-побочным эффектом."""
    literal = repr(meta if meta is not None else META_OK)
    return f"SCRIPT_META = {literal}\n{tail}\n"


class ParseMetaTests(unittest.TestCase):
    """Парсинг меты строго через ast: никакой интерпретации кода скрипта (п. 8.3)."""

    def test_parse_meta_literal(self) -> None:
        meta = pl.parse_meta(source_with_meta())
        self.assertIsNotNone(meta)
        assert meta is not None
        self.assertEqual(meta["name"], META_OK["name"])
        self.assertEqual([p["id"] for p in meta["params"]], ["theme", "duration_min"])

    def test_parse_meta_no_execution(self) -> None:
        """Скрипт, пишущий маркер при импорте, не должен оставить маркер после parse_meta."""
        with tempfile.TemporaryDirectory() as td:
            marker = Path(td) / "marker.txt"
            source = source_with_meta(tail=f"open(r'{marker}', 'w').write('boom')")
            meta = pl.parse_meta(source)
            self.assertIsNotNone(meta)
            self.assertFalse(marker.exists(), "код скрипта не должен исполняться при разборе меты")

    def test_scan_dir_does_not_execute(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            marker = Path(td) / "marker2.txt"
            script = Path(td) / "gen.py"
            script.write_text(source_with_meta(tail=f"open(r'{marker}', 'w').write('boom')"),
                              encoding="utf-8")
            infos = pl.scan_dir(td)
            self.assertEqual(len(infos), 1)
            self.assertTrue(infos[0].has_meta)
            self.assertFalse(marker.exists())

    def test_script_without_meta(self) -> None:
        self.assertIsNone(pl.parse_meta("import sys\nprint(sys.argv)\n"))
        with tempfile.TemporaryDirectory() as td:
            script = Path(td) / "novice.py"
            script.write_text("print('hi')", encoding="utf-8")
            info = pl.build_info(script)
            self.assertFalse(info.has_meta)
            self.assertIsNone(info.error)
            # П. 8.2: показываем по имени файла с пометкой «без метаданных».
            self.assertEqual(info.name, "novice")
            self.assertEqual(info.rules_hook, constants.DEFAULT_RULES_HOOK)

    def test_broken_syntax_kept_in_list(self) -> None:
        """Битый синтаксис: скрипт остаётся в списке с текстом проблемы в поле error."""
        with tempfile.TemporaryDirectory() as td:
            broken = Path(td) / "broken.py"
            broken.write_text("def ( : : )", encoding="utf-8")
            good = Path(td) / "good.py"
            good.write_text(source_with_meta(tail=""), encoding="utf-8")
            infos = pl.scan_dir(td)
            self.assertEqual(len(infos), 2)
            by_name = {i.file_name: i for i in infos}
            self.assertIsNotNone(by_name["broken.py"].error)
            self.assertFalse(by_name["broken.py"].has_meta)
            self.assertEqual(by_name["broken.py"].name, "broken")
            self.assertTrue(by_name["good.py"].has_meta)

    def test_sidecar_priority(self) -> None:
        """Sidecar <name>.meta.json приоритетнее SCRIPT_META в коде (п. 8.2)."""
        with tempfile.TemporaryDirectory() as td:
            script = Path(td) / "gen.py"
            script.write_text(source_with_meta(), encoding="utf-8")
            side = pl.sidecar_path(script)
            side_meta = dict(META_OK, name="Версия из sidecar")
            side.write_text(json.dumps(side_meta, ensure_ascii=False), encoding="utf-8")
            info = pl.build_info(script)
            self.assertEqual(info.name, "Версия из sidecar")
            self.assertTrue(info.has_meta)

    def test_scan_dir_extra_paths(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            d = Path(td) / "scripts"
            d.mkdir()
            (d / "a.py").write_text(source_with_meta(tail=""), encoding="utf-8")
            outside = Path(td) / "Внешний скрипт.py"  # пробелы и кириллица в имени
            outside.write_text("print('outer')", encoding="utf-8")
            infos = pl.scan_dir(d, extra_paths=(outside,))
            self.assertEqual({i.file_name for i in infos}, {"a.py", "Внешний скрипт.py"})


class ArgvTests(unittest.TestCase):
    """build_argv: подстановки контракта п. 8 и универсальная форма для скриптов без меты."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.tmp = Path(self._tmp.name)
        self.info = pl.PluginInfo.from_meta(self.tmp / "gen.py", json.loads(json.dumps(META_OK)))
        self.addCleanup(self._tmp.cleanup)

    def test_template_substitutions(self) -> None:
        meta = {"fio": "Иванов И.И.", "theme": "Моя семья"}
        rules = {"text": {"size_pt": 14}}
        profile = {"name": "Профиль"}
        argv = pl.build_argv(
            self.info, output_path=self.tmp / "out.docx", meta=meta, rules=rules,
            profile=profile, params={"theme": "Прогулка в лес"}, tmp_dir=self.tmp / "args")
        self.assertEqual(argv[0], "--out")
        self.assertEqual(argv[1], str(self.tmp / "out.docx"))
        self.assertEqual(argv[2], "--meta")
        meta_file = Path(argv[3])
        self.assertEqual(json.loads(meta_file.read_text(encoding="utf-8")), meta)
        self.assertEqual(json.loads(Path(argv[5]).read_text(encoding="utf-8")), rules)
        # Профильный JSON — сырой профиль (плоский meta уходит отдельным файлом).
        self.assertEqual(json.loads(Path(argv[7]).read_text(encoding="utf-8")), profile)
        self.assertEqual(argv[-2], "--theme")
        self.assertEqual(argv[-1], "Прогулка в лес")

    def test_template_param_default(self) -> None:
        """Параметр не передан — подставляется default из SCRIPT_META.params."""
        argv = pl.build_argv(self.info, output_path=self.tmp / "o.docx", meta={}, rules={},
                             profile={}, params={}, tmp_dir=self.tmp / "args")
        self.assertEqual(argv[-1], "Моя семья")

    def test_template_unknown_placeholder_kept(self) -> None:
        info = pl.PluginInfo(path=self.tmp / "x.py", file_name="x.py",
                             arg_template=["--size={unknown_param}", "{output}"])
        argv = pl.build_argv(info, output_path=Path("o.docx"), meta={}, rules={},
                             profile={}, params={}, tmp_dir=self.tmp / "args")
        self.assertEqual(argv, ["--size={unknown_param}", "o.docx"])

    def test_template_bool_and_number_params(self) -> None:
        info = pl.PluginInfo(path=self.tmp / "x.py", file_name="x.py",
                             arg_template=["{flag}", "{count}", "{duration_min}"],
                             params=[{"id": "duration_min", "type": "int", "default": 15}])
        argv = pl.build_argv(info, meta={}, rules={}, profile={},
                             params={"flag": True, "count": 3}, tmp_dir=self.tmp / "args")
        self.assertEqual(argv, ["true", "3", "15"])

    def test_universal_form_without_template(self) -> None:
        """Пустой arg_template: input первым (если есть), затем output, затем JSON параметров."""
        info = pl.PluginInfo(path=self.tmp / "u.py", file_name="u.py", input="file_md")
        src = self.tmp / "in.md"
        src.write_text("# заголовок", encoding="utf-8")
        argv = pl.build_argv(info, input_path=src, output_path=self.tmp / "out.docx",
                             meta={}, rules={}, profile={}, params={"a": 1},
                             tmp_dir=self.tmp / "args")
        self.assertEqual(argv[0], str(src))
        self.assertEqual(argv[1], str(self.tmp / "out.docx"))
        self.assertEqual(json.loads(argv[2]), {"a": 1})
        self.assertEqual(len(argv), 3)

    def test_universal_form_input_none(self) -> None:
        """input=none — входной файл в argv не попадает (п. 8, тип none)."""
        info = pl.PluginInfo(path=self.tmp / "u.py", file_name="u.py", input="none")
        argv = pl.build_argv(info, input_path=self.tmp / "ignored.md",
                             output_path=self.tmp / "out.docx", meta={}, rules={},
                             profile={}, params={}, tmp_dir=self.tmp / "args")
        self.assertEqual(len(argv), 2)
        self.assertEqual(argv[0], str(self.tmp / "out.docx"))
        self.assertEqual(json.loads(argv[1]), {})


class ValidateTests(unittest.TestCase):
    """validate(): синтаксис, неизвестные input/rules_hook/тип параметра, arg_template."""

    def _validate_source(self, source: str) -> tuple[bool, list[str]]:
        with tempfile.TemporaryDirectory() as td:
            script = Path(td) / "s.py"
            script.write_text(source, encoding="utf-8")
            return pl.validate(script)

    def test_valid(self) -> None:
        ok, problems = self._validate_source(source_with_meta(tail=""))
        self.assertTrue(ok, problems)
        self.assertEqual(problems, [])

    def test_broken_syntax(self) -> None:
        ok, problems = self._validate_source("def ( :")
        self.assertFalse(ok)
        self.assertTrue(any("Синтакс" in p for p in problems))

    def test_no_meta_is_usable_with_note(self) -> None:
        ok, problems = self._validate_source("print('x')")
        self.assertTrue(ok)  # п. 8.2: без меты можно запускать с ручными аргументами
        self.assertTrue(any("SCRIPT_META" in p for p in problems))

    def test_unknown_input_and_hook(self) -> None:
        bad = dict(META_OK, input="ftp", rules_hook="magic")
        ok, problems = self._validate_source(source_with_meta(bad, tail=""))
        self.assertFalse(ok)
        self.assertTrue(any("input" in p for p in problems))
        self.assertTrue(any("rules_hook" in p for p in problems))

    def test_unknown_param_type_and_bad_template(self) -> None:
        bad = dict(META_OK,
                   params=[{"id": "x", "type": "money"}],
                   arg_template="--out {output}")  # не список
        ok, problems = self._validate_source(source_with_meta(bad, tail=""))
        self.assertFalse(ok)
        self.assertTrue(any("тип" in p for p in problems))
        self.assertTrue(any("arg_template" in p for p in problems))


class SidecarTests(unittest.TestCase):
    """save_sidecar: кнопка «Сохранить описание» (п. 8.2)."""

    def test_save_sidecar_and_reload(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            script = Path(td) / "novice.py"
            script.write_text("import sys\nprint(sys.argv)\n", encoding="utf-8")
            info = pl.build_info(script)
            self.assertFalse(info.has_meta)
            path = pl.save_sidecar(info, arg_template=["{input}", "{output}"],
                                   rules_hook="none", category="Учебные работы")
            self.assertEqual(path.name, "novice.meta.json")
            reloaded = pl.build_info(script)
            self.assertTrue(reloaded.has_meta)
            self.assertEqual(reloaded.arg_template, ["{input}", "{output}"])
            self.assertEqual(reloaded.rules_hook, "none")
            self.assertEqual(reloaded.category, "Учебные работы")

    def test_corrupt_sidecar_falls_back_to_code(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            script = Path(td) / "gen.py"
            script.write_text(source_with_meta(), encoding="utf-8")
            pl.sidecar_path(script).write_text("{битый json", encoding="utf-8")
            info = pl.build_info(script)
            self.assertTrue(info.has_meta)
            self.assertEqual(info.name, META_OK["name"])


if __name__ == "__main__":
    unittest.main()
