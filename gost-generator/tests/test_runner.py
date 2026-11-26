"""Тесты runner + main(--internal-run) без GUI: коды возврата, отмена, таймаут, режимы.

Фикстуры-скрипты генерируются кодом теста во временном каталоге (ТЗ п. 15:
фикстуры не хранятся в репозитории).
"""

from __future__ import annotations

import json
import os
import subprocess
import sys
import tempfile
import threading
import unittest
from pathlib import Path

from app.core import plugin_loader as pl
from app.core import runner as rn

REPO_ROOT = Path(__file__).resolve().parents[1]

SCRIPT_ECHO = """\
import sys
print("STARTED", flush=True)
print("argv:" + "|".join(sys.argv[1:]), flush=True)
print("ГОСТ-генератор: кириллица в выводе", flush=True)
"""

SCRIPT_EXIT1 = """\
import sys
print("before-exit", flush=True)
sys.exit(1)
"""

SCRIPT_EXIT3 = """\
import sys
sys.exit(3)
"""

SCRIPT_SLEEP = """\
import time
print("STARTED", flush=True)
time.sleep(300)
"""

SCRIPT_RAISE = """\
print("before-raise", flush=True)
raise ValueError("намеренный сбой теста")
"""

SCRIPT_MAKE_DOCX = """\
import sys
print("STARTED", flush=True)
with open(sys.argv[1], "w", encoding="utf-8") as fh:
    fh.write("fake docx")
print("DONE", flush=True)
"""


class ScriptFixtures(unittest.TestCase):
    """Базовый класс: временный каталог + запись фикстуры."""

    def setUp(self) -> None:
        self._tmp = tempfile.TemporaryDirectory()
        self.addCleanup(self._tmp.cleanup)
        self.tmp = Path(self._tmp.name)

    def write_script(self, name: str, source: str) -> Path:
        path = self.tmp / name
        path.write_text(source, encoding="utf-8")
        return path


class ExternalModeTests(ScriptFixtures):
    """Внешний режим: [python_path, script, *argv] (п. 3.2.5)."""

    def test_exit_zero_and_lines(self) -> None:
        script = self.write_script("echo.py", SCRIPT_ECHO)
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable, timeout_s=30)
        result = runner.run_argv([sys.executable, str(script), "aa", "bb"])
        self.assertTrue(result.ok, result.stdout_text)
        self.assertEqual(result.exit_code, 0)
        self.assertIn("STARTED", result.lines)
        self.assertIn("argv:aa|bb", result.lines)
        self.assertIn("ГОСТ-генератор: кириллица в выводе", result.stdout_text)
        self.assertEqual(result.argv[0], sys.executable)
        self.assertGreaterEqual(result.duration_s, 0.0)

    def test_on_line_streaming(self) -> None:
        script = self.write_script("echo.py", SCRIPT_ECHO)
        seen: list[str] = []
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable,
                                 timeout_s=30, on_line=seen.append)
        result = runner.run_argv([sys.executable, str(script)])
        self.assertTrue(result.ok)
        self.assertEqual(seen, result.lines)

    def test_exit_code_nonzero(self) -> None:
        script = self.write_script("fail.py", SCRIPT_EXIT1)
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable, timeout_s=30)
        result = runner.run_argv([sys.executable, str(script)])
        self.assertEqual(result.exit_code, 1)
        self.assertFalse(result.ok)
        self.assertIn("before-exit", result.lines)

    def test_traceback_captured_in_lines(self) -> None:
        script = self.write_script("raise.py", SCRIPT_RAISE)
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable, timeout_s=30)
        result = runner.run_argv([sys.executable, str(script)])
        self.assertFalse(result.ok)
        self.assertIn("Traceback", result.stdout_text)
        self.assertIn("ValueError", result.stdout_text)

    def test_timeout_kills_process(self) -> None:
        script = self.write_script("sleep.py", SCRIPT_SLEEP)
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable, timeout_s=1)
        result = runner.run_argv([sys.executable, str(script)])
        self.assertTrue(result.timed_out)
        self.assertFalse(result.ok)
        self.assertLess(result.duration_s, 10.0)

    def test_cancel_terminates_long_process(self) -> None:
        script = self.write_script("sleep.py", SCRIPT_SLEEP)
        started = threading.Event()
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable,
                                 timeout_s=30, on_line=lambda line: started.set() if line == "STARTED" else None)
        box: dict[str, rn.RunResult] = {}

        def target() -> None:
            box["res"] = runner.run_argv([sys.executable, str(script)])

        thread = threading.Thread(target=target, daemon=True)
        thread.start()
        self.assertTrue(started.wait(timeout=15), "процесс не вывел STARTED")
        runner.cancel()
        thread.join(timeout=15)
        self.assertFalse(thread.is_alive(), "run_argv не завершился после cancel()")
        result = box["res"]
        self.assertTrue(result.cancelled)
        self.assertFalse(result.ok)

    def test_run_script_builds_argv_from_template(self) -> None:
        script = self.write_script("echo.py", SCRIPT_ECHO)
        info = pl.PluginInfo(path=script, file_name=script.name,
                             arg_template=["{output}", "--theme={theme}"])
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable, timeout_s=30)
        result = runner.run_script(info, meta={}, rules={}, profile={},
                                   output_path=self.tmp / "out.docx",
                                   params={"theme": "Зима"}, tmp_dir=self.tmp / "run")
        self.assertTrue(result.ok, result.stdout_text)
        self.assertIn("argv:" + str(self.tmp / "out.docx") + "|--theme=Зима", result.lines)


class InternalModeTests(ScriptFixtures):
    """Внутренний режим через app.main --internal-run (п. 3.2)."""

    def test_main_module_internal_run(self) -> None:
        script = self.write_script("echo.py", SCRIPT_ECHO)
        params = self.tmp / "params.json"
        params.write_text(json.dumps({"argv": ["X", "Y"], "cwd": str(self.tmp)}),
                          encoding="utf-8")
        proc = subprocess.run(
            [sys.executable, "-m", "app.main", "--internal-run", str(script), str(params)],
            cwd=str(REPO_ROOT), capture_output=True, text=True, encoding="utf-8",
            errors="replace", timeout=60)
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("STARTED", proc.stdout)
        self.assertIn("argv:X|Y", proc.stdout)

    def test_main_package_internal_run_and_exit_codes(self) -> None:
        """`python -m app` — вторая половина той же точки входа; код возврата скрипта передаётся наружу."""
        fail = self.write_script("exit3.py", SCRIPT_EXIT3)
        params = self.tmp / "p.json"
        params.write_text(json.dumps({"argv": [], "cwd": None}), encoding="utf-8")
        proc = subprocess.run(
            [sys.executable, "-m", "app", "--internal-run", str(fail), str(params)],
            cwd=str(REPO_ROOT), capture_output=True, text=True, timeout=60,
            encoding="utf-8", errors="replace")
        self.assertEqual(proc.returncode, 3, proc.stdout + proc.stderr)

        boom = self.write_script("raise.py", SCRIPT_RAISE)
        proc = subprocess.run(
            [sys.executable, "-m", "app", "--internal-run", str(boom), str(params)],
            cwd=str(REPO_ROOT), capture_output=True, text=True, timeout=60,
            encoding="utf-8", errors="replace")
        self.assertEqual(proc.returncode, 1)
        self.assertIn("Traceback", proc.stdout)

    def test_gui_not_imported_in_internal_mode(self) -> None:
        """Ветка --internal-run не тянет GUI: проверим импортом app.ui.app_window в скрипте-маркере."""
        probe = self.write_script(
            "probe.py",
            "import sys\nprint('ui_loaded:' + str('app.ui.app_window' in sys.modules or 'customtkinter' in sys.modules))\n")
        params = self.tmp / "p.json"
        params.write_text(json.dumps({"argv": [], "cwd": None}), encoding="utf-8")
        proc = subprocess.run(
            [sys.executable, "-m", "app.main", "--internal-run", str(probe), str(params)],
            cwd=str(REPO_ROOT), capture_output=True, text=True, timeout=60,
            encoding="utf-8", errors="replace")
        self.assertEqual(proc.returncode, 0, proc.stdout + proc.stderr)
        self.assertIn("ui_loaded:False", proc.stdout)

    def test_runner_internal_dev_mode(self) -> None:
        """ScriptRunner(mode='internal') в режиме разработки строит [python, -m, app.main, --internal-run, ...]."""
        script = self.write_script("echo.py", SCRIPT_ECHO)
        info = pl.PluginInfo(path=script, file_name=script.name, arg_template=["{output}"])
        runner = rn.ScriptRunner(mode="internal", timeout_s=60)
        result = runner.run_script(info, meta={}, rules={}, profile={},
                                   output_path=self.tmp / "o.docx", params={},
                                   tmp_dir=self.tmp / "run")
        self.assertTrue(result.ok, result.stdout_text)
        self.assertEqual(result.argv[:4],
                         [sys.executable, "-m", "app.main", "--internal-run"])
        # Windows-temp может прийти 8.3-именем (RUNNER~1) — сравниваем нормализованно.
        self.assertEqual(os.path.normcase(str(Path(result.argv[4]).resolve())),
                         os.path.normcase(str(script.resolve())))
        payload = json.loads(Path(result.argv[5]).read_text(encoding="utf-8"))
        self.assertEqual(payload["argv"][0], str(self.tmp / "o.docx"))
        self.assertIn("argv:" + str(self.tmp / "o.docx"), result.lines)

    def test_runner_internal_nonzero_exit(self) -> None:
        script = self.write_script("exit3.py", SCRIPT_EXIT3)
        info = pl.PluginInfo(path=script, file_name=script.name, arg_template=["{output}"])
        runner = rn.ScriptRunner(mode="internal", timeout_s=60)
        result = runner.run_script(info, meta={}, rules={}, profile={},
                                   output_path=self.tmp / "o.docx", params={},
                                   tmp_dir=self.tmp / "run")
        self.assertEqual(result.exit_code, 3)
        self.assertFalse(result.ok)


class RunAndFinalizeTests(ScriptFixtures):
    """Оркестрация run → normalize → check: lazy-импорты переживают отсутствие модулей C."""

    def test_finalize_skips_missing_core_modules(self) -> None:
        try:
            import app.core.docx_normalizer  # noqa: F401
            has_normalizer = True
        except ImportError:
            has_normalizer = False
        try:
            import app.core.docx_inspector  # noqa: F401
            has_inspector = True
        except ImportError:
            has_inspector = False

        script = self.write_script("mk.py", SCRIPT_MAKE_DOCX)
        info = pl.PluginInfo(path=script, file_name=script.name, arg_template=["{output}"],
                             rules_hook="postprocess", has_meta=True)
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable, timeout_s=30)
        out = self.tmp / "gen.docx"
        outcome = rn.run_and_finalize(runner, info=info, meta={}, rules={}, profile={},
                                      output_path=out, params={})
        self.assertTrue(outcome["run"].ok, outcome["run"].stdout_text)
        self.assertEqual(outcome["output"], str(out))
        self.assertTrue(out.exists())
        if not has_normalizer:
            self.assertEqual(outcome["normalized"], {"skipped": "normalizer unavailable"})
        else:  # pragma: no cover — модуль появится позже, проверяем лишь непадение
            self.assertIsInstance(outcome["normalized"], dict)
        if not has_inspector:
            self.assertEqual(outcome["check"], {"skipped": "inspector unavailable"})
        else:  # pragma: no cover
            self.assertIsInstance(outcome["check"], dict)

    def test_finalize_no_postprocess_hook(self) -> None:
        """rules_hook=none — нормализатор не вызывается даже при do_normalize=True."""
        script = self.write_script("mk.py", SCRIPT_MAKE_DOCX)
        info = pl.PluginInfo(path=script, file_name=script.name, arg_template=["{output}"],
                             rules_hook="none", has_meta=True)
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable, timeout_s=30)
        outcome = rn.run_and_finalize(runner, info=info, meta={}, rules={}, profile={},
                                      output_path=self.tmp / "x.docx", params={},
                                      do_normalize=True, do_check=False,
                                      fail_on="never")
        self.assertTrue(outcome["ok"])
        self.assertIsNone(outcome["normalized"])
        self.assertIsNone(outcome["check"])

    def test_finalize_failed_run_short_circuits(self) -> None:
        script = self.write_script("fail.py", SCRIPT_EXIT1)
        info = pl.PluginInfo(path=script, file_name=script.name, arg_template=["{output}"],
                             rules_hook="postprocess", has_meta=True)
        runner = rn.ScriptRunner(mode="external", python_path=sys.executable, timeout_s=30)
        outcome = rn.run_and_finalize(runner, info=info, meta={}, rules={}, profile={},
                                      output_path=self.tmp / "x.docx", params={})
        self.assertFalse(outcome["ok"])
        self.assertIsNone(outcome["normalized"])
        self.assertIsNone(outcome["check"])


if __name__ == "__main__":
    unittest.main()
