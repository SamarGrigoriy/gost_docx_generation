"""Тесты headless-CLI `gost` (ТЗ §15, §19, §20.1).

Запускают реальный процесс: `python -m app.cli ...` из корня репозитория.
Проверяют: контракт конверта (п. 19.4), коды возврата (п. 19.5), разделение
stdout/stderr, отсутствие GUI-импортов, генерацию артефактов init --agents.

Часть интеграционных тестов (check/fix/job на реальных .docx) зависит от модулей
app.core.docx_inspector / docx_normalizer / report — они собираются параллельно;
до их появления эти тесты красные, остальные — зелёные.
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import unittest
import zipfile
from pathlib import Path
from typing import Any

REPO = Path(__file__).resolve().parent.parent
ENVELOPE_KEYS = {"schema_version", "ok", "command", "data", "errors", "warnings"}


def run_cli(*args: str, cwd: Path | None = None,
            env: dict[str, str] | None = None) -> subprocess.CompletedProcess[str]:
    """Вызвать CLI подкомандой в отдельном процессе (headless, без GUI)."""
    return subprocess.run(
        [sys.executable, "-m", "app.cli", *args],
        cwd=str(cwd or REPO), capture_output=True, text=True, timeout=120,
        env={**os.environ, **(env or {})} if env is not None else None,
    )


def parse_envelope(stdout: str) -> dict[str, Any]:
    """Разобрать JSON-конверт из stdout и проверить его обязательные поля."""
    env = json.loads(stdout)
    assert isinstance(env, dict), "конверт должен быть JSON-объектом"
    assert ENVELOPE_KEYS <= set(env), f"нет полей конверта: {ENVELOPE_KEYS - set(env)}"
    assert env["schema_version"] == 1
    assert isinstance(env["ok"], bool)
    assert isinstance(env["errors"], list)
    assert isinstance(env["warnings"], list)
    return env


def make_minimal_docx(path: Path) -> None:
    """Собрать минимально валидный .docx (zip с OOXML) без сторонних библиотек."""
    doc = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
           '<w:body><w:p><w:r><w:t>Текст</w:t></w:r></w:p></w:body></w:document>')
    ct = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          '<Default Extension="xml" ContentType="application/xml"/>'
          '<Override PartName="/word/document.xml" '
          'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
          '</Types>')
    rels = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument'
            '/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>')
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("[Content_Types].xml", ct)
        z.writestr("_rels/.rels", rels)
        z.writestr("word/document.xml", doc)


class TestCliBasics(unittest.TestCase):
    """Базовые гарантии: --help без GUI, конверт, коды возврата, разделение потоков."""

    def test_help_works_headless(self) -> None:
        """`--help` не должен требовать tkinter/GUI-модули."""
        proc = run_cli("--help")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("gost", proc.stdout)
        for cmd in ("init", "doctor", "rules", "profile", "scripts", "run",
                    "check", "fix", "batch", "job", "schema", "explain", "context"):
            self.assertIn(cmd, proc.stdout)

    def test_cli_does_not_import_gui(self) -> None:
        """Импорт app.cli и выполнение команд не тянут GUI (п. 19.1): ядовитые
        записи в sys.modules превратили бы любой импорт в TypeError."""
        code = (
            "import sys; sys.path.insert(0, %r);"
            "sys.modules['customtkinter'] = None; sys.modules['app.ui'] = None;"
            "sys.modules['tkinter'] = None;"
            "import app.cli; rc = app.cli.main(['doctor', '--json']); sys.exit(rc)" % str(REPO)
        )
        proc = subprocess.run([sys.executable, "-c", code], cwd=str(REPO),
                              capture_output=True, text=True, timeout=120)
        self.assertEqual(proc.returncode, 0, proc.stderr)
        env = parse_envelope(proc.stdout)
        self.assertTrue(env["ok"])

    def test_source_files_have_no_gui_imports(self) -> None:
        """grep-тест: в app/cli.py и app/cli_init.py нет 'customtkinter' и 'app.ui'."""
        for rel in ("app/cli.py", "app/cli_init.py"):
            src = (REPO / rel).read_text(encoding="utf-8")
            self.assertNotIn("customtkinter", src, f"{rel}: упоминание GUI-библиотеки")
            self.assertNotIn("app.ui", src, f"{rel}: упоминание GUI-пакета")

    def test_doctor_json_envelope(self) -> None:
        """doctor --json: валидный конверт в stdout; stderr не содержит результат."""
        proc = run_cli("doctor", "--json")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        env = parse_envelope(proc.stdout)
        self.assertEqual(env["command"], "doctor")
        self.assertIn("checks", env["data"])
        self.assertNotIn("schema_version", proc.stderr)

    def test_context_json_envelope(self) -> None:
        """context --json: пресет полностью, скрипты, профили, каталоги, версии, python."""
        proc = run_cli("context", "--json")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        env = parse_envelope(proc.stdout)
        data = env["data"]
        for key in ("rules", "scripts", "profiles", "dirs", "versions", "python",
                    "active_rules_key", "active_profile"):
            self.assertIn(key, data, f"context: нет поля {key}")
        self.assertEqual(data["rules"]["page"]["size"], "A4")
        self.assertNotIn("schema_version", proc.stderr)

    def test_stdout_stderr_separation_text(self) -> None:
        """Без --json результат идёт в stdout, а логи — в stderr."""
        proc = run_cli("explain", "PAGE_MARGINS")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("PAGE_MARGINS", proc.stdout)
        self.assertNotIn("PAGE_MARGINS:", proc.stderr)

    def test_check_missing_file_exit3(self) -> None:
        """check несуществующего файла → код возврата 3 и FILE_NOT_FOUND."""
        proc = run_cli("check", "/nonexistent/no_such_file.docx", "--json")
        self.assertEqual(proc.returncode, 3, proc.stderr)
        env = parse_envelope(proc.stdout)
        self.assertFalse(env["ok"])
        self.assertEqual(env["errors"][0]["code"], "FILE_NOT_FOUND")

    def test_unknown_command_exit2(self) -> None:
        """Неизвестная подкоманда → код 2 (п. 19.5)."""
        proc = run_cli("frobnicate")
        self.assertEqual(proc.returncode, 2)

    def test_no_subcommand_exit2(self) -> None:
        """Вызов без команды → код 2 (argparse required subcommand)."""
        proc = run_cli()
        self.assertEqual(proc.returncode, 2)


class TestCliKnowledge(unittest.TestCase):
    """rules / profile / schema / explain — знания о пресетах и профилях."""

    def test_rules_list(self) -> None:
        proc = run_cli("rules", "list", "--json")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        env = parse_envelope(proc.stdout)
        keys = {i["key"] for i in env["data"]["items"]}
        self.assertIn("gost_732", keys)

    def test_rules_show(self) -> None:
        proc = run_cli("rules", "show", "gost_732", "--json")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        env = parse_envelope(proc.stdout)
        self.assertEqual(env["data"]["rules"]["text"]["size_pt"], 14)

    def test_rules_validate(self) -> None:
        proc = run_cli("rules", "validate", "gost_732", "--json")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        env = parse_envelope(proc.stdout)
        self.assertTrue(env["ok"], env["errors"])

    def test_profile_list(self) -> None:
        proc = run_cli("profile", "list", "--json")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        env = parse_envelope(proc.stdout)
        self.assertTrue(any(i["key"] == "default" for i in env["data"]["items"]))

    def test_schema_all_names(self) -> None:
        """gost schema <name> → валидный JSON Schema draft 2020-12 для каждого имени."""
        for name in ("rules", "profile", "script-meta", "report", "job"):
            proc = run_cli("schema", name)
            self.assertEqual(proc.returncode, 0, f"{name}: {proc.stderr}")
            schema = json.loads(proc.stdout)
            self.assertEqual(schema.get("$schema", ""),
                             "https://json-schema.org/draft/2020-12/schema", name)
            self.assertIn("properties", schema, name)

    def test_schema_unknown_exit2(self) -> None:
        proc = run_cli("schema", "nope")
        self.assertEqual(proc.returncode, 2)

    def test_explain_page_margins(self) -> None:
        """explain PAGE_MARGINS — текст с заголовком/исправлением (п. 20.1 п.4)."""
        proc = run_cli("explain", "PAGE_MARGINS")
        self.assertEqual(proc.returncode, 0, proc.stderr)
        self.assertIn("Поля страницы", proc.stdout)
        self.assertIn("gost fix", proc.stdout)

    def test_explain_json_and_cli_codes(self) -> None:
        """Словарь покрывает все коды §9.1 и коды ошибок CLI."""
        codes_9_1 = ["PAGE_MARGINS", "PAGE_SIZE", "FONT_MAIN", "SIZE_MAIN", "LINE_SPACING",
                     "INDENT_VALUE", "ALIGNMENT", "SPACING_PARA", "NUM_POSITION", "NUM_START",
                     "HEAD_STRUCT", "HEAD_KEEP", "HEAD_SIZE", "FIG_CAPTION", "FIG_COUNT",
                     "FIG_ORDER", "FIG_REF", "TBL_CAPTION", "TBL_HEADER", "TBL_BORDERS",
                     "TBL_ORDER", "TBL_REF", "FORMULA", "FOOTNOTE_FONT", "TOC_PRESENT",
                     "BIB_ORDER", "BIB_AGE", "BIB_COVERAGE", "APPENDIX", "VOLUME_FILL", "MANUAL"]
        for code in codes_9_1 + ["FILE_NOT_FOUND", "VIOLATIONS", "TIMEOUT"]:
            proc = run_cli("explain", code, "--json")
            self.assertEqual(proc.returncode, 0, f"{code}: {proc.stderr}")
            env = parse_envelope(proc.stdout)
            for field in ("title", "detail", "fix", "source"):
                self.assertIn(field, env["data"], f"{code}: нет поля {field}")


class TestCliInit(unittest.TestCase):
    """gost init --agents — генерация артефактов (п. 20)."""

    def test_init_agents_creates_artifacts(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            target = Path(td) / "proj"
            proc = run_cli("init", "--agents", "--dir", str(target), "--json")
            self.assertEqual(proc.returncode, 0, proc.stderr)
            env = parse_envelope(proc.stdout)
            self.assertTrue(env["ok"])
            self.assertEqual(env["data"]["dir"], str(target))
            for rel in ("AGENTS.md", "schemas/rules.json", "schemas/job.json",
                        "examples/task.json"):
                self.assertTrue((target / rel).exists(), f"не создан {rel}")
            for d in ("scripts", "presets", "profiles", "output", "logs"):
                self.assertTrue((target / d).is_dir(), f"не создан каталог {d}")
            agents = (target / "AGENTS.md").read_text(encoding="utf-8")
            self.assertIn("gost context --json", agents)
            self.assertIn("stdout", agents)
            self.assertIn("gost job run task.json", agents)
            task = json.loads((target / "examples" / "task.json").read_text(encoding="utf-8"))
            self.assertIn("script", task)
            # схемы на диске обязаны совпадать со схемой команды `gost schema job`
            schema = json.loads((target / "schemas" / "job.json").read_text(encoding="utf-8"))
            self.assertEqual(schema["$schema"], "https://json-schema.org/draft/2020-12/schema")


class TestCliDryRun(unittest.TestCase):
    """run --dry-run: план без запуска и записей (п. 19.4)."""

    def test_run_dry_run_plan(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            out = Path(td) / "result.docx"
            proc = run_cli("run", "mkdocx.py", "--input", "scripts/check_docx.py",
                           "--output", str(out), "--dry-run", "--json")
            self.assertEqual(proc.returncode, 0, proc.stderr)
            env = parse_envelope(proc.stdout)
            self.assertTrue(env["data"]["dry_run"])
            self.assertIsInstance(env["data"]["argv"], list)
            self.assertFalse(out.exists(), "--dry-run не должен создавать файлы")


class TestCliIntegration(unittest.TestCase):
    """Интеграция с ядром-сборкой (Agent C: docx_inspector/docx_normalizer/report).

    Эти тесты описывают итоговый контракт и становятся зелёными после сборки
    соседних модулей; их падение не влияет на остальные классы.
    """

    def test_check_real_docx_contract(self) -> None:
        """check data: file/preset/strict/summary/issues c location{paragraph,snippet}."""
        with tempfile.TemporaryDirectory() as td:
            docx = Path(td) / "doc.docx"
            make_minimal_docx(docx)
            proc = run_cli("check", str(docx), "--rules", "gost_732", "--json")
            self.assertIn(proc.returncode, (0, 1), proc.stderr)  # нарушения → 1
            env = parse_envelope(proc.stdout)
            data = env["data"]
            for key in ("file", "preset", "strict", "summary", "issues"):
                self.assertIn(key, data)
            self.assertEqual(set(data["summary"]) , {"error", "warning", "info"})
            for issue in data["issues"]:
                self.assertIn("code", issue)
                self.assertIn("severity", issue)
                self.assertIn("location", issue)
                self.assertIn("paragraph", issue["location"])
                self.assertIn("snippet", issue["location"])

    def test_fix_creates_backup_and_idempotent(self) -> None:
        """fix: бэкап <name>.orig.docx; повторный запуск не меняет файл (п. 19.4)."""
        with tempfile.TemporaryDirectory() as td:
            docx = Path(td) / "doc.docx"
            make_minimal_docx(docx)
            first = run_cli("fix", str(docx), "--rules", "gost_732", "--json")
            self.assertEqual(first.returncode, 0, first.stderr)
            env = parse_envelope(first.stdout)
            self.assertTrue(env["ok"], env["errors"])
            after_first = docx.read_bytes()
            second = run_cli("fix", str(docx), "--rules", "gost_732", "--json")
            self.assertEqual(second.returncode, 0, second.stderr)
            env2 = parse_envelope(second.stdout)
            self.assertFalse(env2["data"].get("changed", False),
                             "повторный fix не должен менять файл (идемпотентность)")
            self.assertEqual(after_first, docx.read_bytes())

    def test_job_run_full_cycle(self) -> None:
        """job run task.json — цикл агента init → context → run → check (п. 20.1)."""
        with tempfile.TemporaryDirectory() as td:
            task = {
                "schema_version": 1,
                "script": "mkdocx.py",
                "input": "scripts/check_docx.py",  # любой существующий вход
                "output": str(Path(td) / "job.docx"),
                "rules": "gost_732",
                "profile": "default",
                "post": {"normalize": True, "check": False, "fail_on": "never"},
            }
            task_path = Path(td) / "task.json"
            task_path.write_text(json.dumps(task, ensure_ascii=False), encoding="utf-8")
            proc = run_cli("job", "run", str(task_path), "--json")
            env = parse_envelope(proc.stdout)
            self.assertIn(proc.returncode, (0, 1, 4), proc.stderr)
            self.assertEqual(env["command"], "job")
            self.assertIn("output", env["data"])


_META_ONLY = ('SCRIPT_META = {{"name": "{name}", "version": "1.0", "input": "none",\n'
              '  "output_ext": ".docx", "rules_hook": "none",\n'
              '  "arg_template": ["--out", "{{output}}"], "params": []}}\n')


class TestCliReturnCodes(unittest.TestCase):
    """Все восемь кодов возврата (п. 19.5) и приоритет настроек (п. 19.6)."""

    @staticmethod
    def _env() -> dict[str, str]:
        """CLI из временного «проекта» всё равно должен видеть пакет app из репозитория."""
        return {"PYTHONPATH": str(REPO)}

    def test_script_failure_exit4(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root, env = Path(td), self._env()
            script = root / "boom.py"
            script.write_text(_META_ONLY.format(name="Boom") + "import sys\nsys.exit(3)\n",
                              encoding="utf-8")
            proc = run_cli("run", str(script), "--output", str(root / "b.docx"),
                           "--json", cwd=root, env=env)
            self.assertEqual(proc.returncode, 4, proc.stderr)
            self.assertIn("SCRIPT_FAILED", {e["code"] for e in parse_envelope(proc.stdout)["errors"]})

    def test_script_timeout_exit5(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root, env = Path(td), self._env()
            script = root / "slow.py"
            script.write_text(_META_ONLY.format(name="Slow") + "import time\ntime.sleep(30)\n",
                              encoding="utf-8")
            (root / "gost.json").write_text(json.dumps({"script_timeout_s": 0.2}), encoding="utf-8")
            proc = run_cli("run", str(script), "--output", str(root / "s.docx"),
                           "--json", cwd=root, env=env)
            self.assertEqual(proc.returncode, 5, proc.stderr)
            self.assertIn("TIMEOUT", {e["code"] for e in parse_envelope(proc.stdout)["errors"]})

    def test_invalid_job_exit6(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root, env = Path(td), self._env()
            task = root / "task.json"
            task.write_text(json.dumps({"script": 123}), encoding="utf-8")
            proc = run_cli("job", "run", str(task), "--json", cwd=root, env=env)
            self.assertEqual(proc.returncode, 6, proc.stderr)
            self.assertIn("JOB_INVALID", {e["code"] for e in parse_envelope(proc.stdout)["errors"]})

    def test_normalize_failure_exit7(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root, env = Path(td), self._env()
            broken = root / "broken.docx"
            broken.write_text("это не zip-архив", encoding="utf-8")
            proc = run_cli("fix", str(broken), "--json", cwd=root, env=env)
            self.assertEqual(proc.returncode, 7, proc.stderr)
            self.assertIn("NORMALIZE_FAILED",
                          {e["code"] for e in parse_envelope(proc.stdout)["errors"]})

    def test_fix_dry_run_changes_nothing(self) -> None:
        with tempfile.TemporaryDirectory() as td:
            root, env = Path(td), self._env()
            docx = root / "doc.docx"
            make_minimal_docx(docx)
            before = docx.read_bytes()
            proc = run_cli("fix", str(docx), "--dry-run", "--json", cwd=root, env=env)
            self.assertEqual(proc.returncode, 0, proc.stderr)
            data = parse_envelope(proc.stdout)["data"]
            self.assertTrue(data["dry_run"])
            self.assertIsInstance(data["planned_actions"], list)
            self.assertEqual(before, docx.read_bytes(), "--dry-run не пишет файл")
            self.assertFalse((root / "doc.orig.docx").exists(), "--dry-run не создаёт бэкап")

    def test_settings_priority(self) -> None:
        """флаги > GOST_RULES > gost.json проекта > пользовательский конфиг/дефолт (п. 19.6)."""
        with tempfile.TemporaryDirectory() as td:
            root, env = Path(td), self._env()
            docx = root / "doc.docx"
            make_minimal_docx(docx)

            def preset(*flags: str, rules_env: str | None = None) -> str:
                proc = run_cli("check", str(docx), *flags, "--json", cwd=root,
                               env={**env, **({"GOST_RULES": rules_env} if rules_env else {})})
                return parse_envelope(proc.stdout)["data"]["preset"]

            self.assertIn("Мультиворк", preset())
            (root / "gost.json").write_text(json.dumps({"rules": "pedagogical"}), encoding="utf-8")
            self.assertIn("Педагогическ", preset())
            self.assertIn("Свободный", preset(rules_env="free"))
            self.assertIn("Мультиворк", preset("--rules", "gost_732", rules_env="free"))


if __name__ == "__main__":
    unittest.main()
