"""Запуск скриптов-плагинов: subprocess, потоковое чтение, отмена, таймаут (п. 3.2, 9, 12 ТЗ).

Два режима интерпретатора:
- ``internal`` — скрипт запускает само приложение: внутри PyInstaller-exe
  ``sys.executable`` — это сам exe, поэтому команда
  ``[sys.executable, "--internal-run", script, params_json]`` работает и в собранном
  виде, и (через ``-m app.main``) в режиме разработки;
- ``external`` — внешний Python (для скриптов со сторонними зависимостями):
  ``[python_path, script, *argv]``.

UI никогда не блокируется: строки вывода читаются в отдельном потоке и публикуются
в ``queue.Queue`` + колбэк ``on_line``. Никакого ``shell=True``.
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import tempfile
import threading
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable

from app import constants
from app.core import plugin_loader

# Корень репозитория (app/core/runner.py -> вверх на 2 уровня): нужен, чтобы в режиме
# разработки дочерний процесс `-m app.main` гарантированно находил пакет app,
# независимо от cwd запуска.
PROJECT_ROOT = str(Path(__file__).resolve().parents[2])

MODE_INTERNAL = "internal"
MODE_EXTERNAL = "external"
INTERNAL_RUN_FLAG = "--internal-run"


@dataclass
class RunResult:
    """Структурированный результат запуска (задел под журнал запусков v2, п. 13)."""

    argv: list[str] = field(default_factory=list)
    exit_code: int | None = None
    lines: list[str] = field(default_factory=list)
    duration_s: float = 0.0
    cancelled: bool = False
    timed_out: bool = False
    error: str | None = None

    @property
    def ok(self) -> bool:
        """Успех = нулевой код возврата и ни отмены, ни таймаута, ни сбоя запуска."""
        return (self.exit_code == 0 and not self.cancelled
                and not self.timed_out and self.error is None)

    @property
    def stdout_text(self) -> str:
        """Весь перехваченный вывод (stdout+stderr) одним текстом."""
        return "\n".join(self.lines)


class ScriptRunner:
    """Один раннер = одна последовательная сессия запуска; потокобезопасная отмена."""

    def __init__(self, timeout_s: float = constants.DEFAULT_SCRIPT_TIMEOUT_S,
                 mode: str = MODE_INTERNAL, python_path: str | None = None,
                 on_line: Callable[[str], None] | None = None,
                 kill_grace_s: float = constants.KILL_GRACE_PERIOD_S) -> None:
        self.timeout_s = timeout_s
        self.mode = mode
        self.python_path = python_path
        self.on_line = on_line
        self.kill_grace_s = kill_grace_s
        self.lines_queue: queue.Queue[str] = queue.Queue()
        self._proc: subprocess.Popen[str] | None = None
        self._cancel_requested = False

    # ------------------------------------------------------------------ отмена

    def cancel(self) -> None:
        """Остановить текущий процесс: terminate → kill через kill_grace_s (п. 12)."""
        self._cancel_requested = True
        proc = self._proc
        if proc is None or proc.poll() is not None:
            return
        try:
            proc.terminate()  # POSIX: SIGTERM процессу; Windows: TerminateProcess-подготовку
            try:
                proc.wait(timeout=self.kill_grace_s)
                return
            except subprocess.TimeoutExpired:
                pass
            proc.kill()
        except OSError:
            # Процесс уже мог завершиться между poll() и terminate() — это не ошибка.
            pass

    # ------------------------------------------------------------- низкоуровневый запуск

    def _publish(self, line: str) -> None:
        """Отдать строку в очередь (для UI-цикла) и колбэку (для логов/тестов)."""
        self.lines_queue.put(line)
        if self.on_line is not None:
            try:
                self.on_line(line)
            except Exception:  # колбэк UI не должен ронять читателя вывода
                pass

    def run_argv(self, argv: list[str], *, cwd: str | Path | None = None) -> RunResult:
        """Выполнить готовую команду, возвращая структурированный результат.

        stdout и stderr объединяются (stderr=STDOUT), кодировка UTF-8 с errors="replace":
        скрипты печатают кириллицу, и битый байт не должен ронять лог-поток.
        """
        result = RunResult(argv=list(argv))
        if not argv:
            result.error = "Пустая команда запуска"
            return result

        creationflags = subprocess.CREATE_NO_WINDOW if os.name == "nt" else 0
        env = os.environ.copy()
        # Для `-m app.main` в режиме разработки: PYTHONPATH всегда видит корень репозитория.
        env["PYTHONPATH"] = (PROJECT_ROOT + os.pathsep + env["PYTHONPATH"]
                             if env.get("PYTHONPATH") else PROJECT_ROOT)
        start = time.monotonic()
        self._cancel_requested = False
        try:
            proc = subprocess.Popen(
                argv,
                cwd=str(cwd) if cwd else None,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=creationflags,
                env=env,
            )
        except OSError as exc:
            result.error = f"Не удалось запустить процесс: {exc}"
            result.duration_s = time.monotonic() - start
            return result

        self._proc = proc
        lines: list[str] = []
        lines_lock = threading.Lock()

        def reader() -> None:
            # Построчное чтение в отдельном потоке: UI получает лог в реальном времени,
            # а главный поток ждёт завершения и может принять cancel() в любой момент.
            assert proc.stdout is not None
            try:
                for raw in proc.stdout:
                    line = raw.rstrip("\r\n")
                    with lines_lock:
                        lines.append(line)
                    self._publish(line)
            except (OSError, ValueError):
                pass  # поток закрылся при завершении процесса — штатная ситуация

        thread = threading.Thread(target=reader, daemon=True, name="gost-runner-io")
        thread.start()
        try:
            result.exit_code = proc.wait(timeout=self.timeout_s)
        except subprocess.TimeoutExpired:
            result.timed_out = True
            proc.kill()
            try:
                result.exit_code = proc.wait(timeout=self.kill_grace_s)
            except subprocess.TimeoutExpired:  # pragma: no cover — зомби-процесс вне нашего контроля
                result.exit_code = None
        thread.join(timeout=2.0)
        try:
            proc.stdout.close()  # иначе unittest ResourceWarning на незакрытом канале
        except OSError:
            pass
        with lines_lock:
            result.lines = list(lines)
        result.cancelled = self._cancel_requested
        result.duration_s = time.monotonic() - start
        self._proc = None
        return result

    # ------------------------------------------------------------- высокоуровневый запуск

    def _build_command(self, script: Path, argv: list[str], params_json_path: Path,
                       cwd: str | Path | None) -> list[str]:
        """Собрать команду по режиму (п. 3.2): internal — через --internal-run, external — напрямую."""
        if self.mode == MODE_EXTERNAL:
            python = self.python_path or sys.executable
            return [python, str(script), *argv]
        if getattr(sys, "frozen", False):
            # Внутри собранного exe sys.executable — сам GostGenerator.exe.
            return [sys.executable, INTERNAL_RUN_FLAG, str(script), str(params_json_path)]
        # Режим разработки: тот же служебный путь, но через модуль app.main.
        return [sys.executable, "-m", "app.main", INTERNAL_RUN_FLAG,
                str(script), str(params_json_path)]

    def run_script(self, info: plugin_loader.PluginInfo | dict, *, meta: dict,
                   rules: dict, profile: dict, input_path: Path | str | None = None,
                   output_path: Path | str | None = None, params: dict | None = None,
                   cwd: str | Path | None = None,
                   tmp_dir: Path | None = None) -> RunResult:
        """Запустить скрипт-плагин с построением argv по его arg_template.

        Все временные JSON (meta/rules/profile/params для internal-режима) пишутся
        в tmp_dir (по умолчанию — временный каталог), чтобы не засорять вывод.
        """
        if isinstance(info, dict):
            info_obj = plugin_loader.PluginInfo.from_meta(Path(str(info.get("path", ""))), info)
        else:
            info_obj = info
        result_stub = RunResult(argv=[])
        if not info_obj.path:
            result_stub.error = "Не указан путь к скрипту"
            return result_stub
        if not Path(info_obj.path).exists():
            result_stub.error = f"Скрипт не найден: {info_obj.path}"
            return result_stub

        tmp = Path(tmp_dir) if tmp_dir else Path(tempfile.mkdtemp(prefix="gost_run_"))
        tmp.mkdir(parents=True, exist_ok=True)

        cli_argv = plugin_loader.build_argv(
            info_obj, input_path=input_path, output_path=output_path,
            meta=meta, rules=rules, profile=profile, params=params or {}, tmp_dir=tmp)

        if self.mode == MODE_INTERNAL:
            # Внутренний режим: точка входа читает argv из JSON, см. app/main.py.
            params_json = tmp / "gost_params.json"
            payload = {"argv": cli_argv, "cwd": str(cwd) if cwd else None}
            params_json.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
        else:
            params_json = tmp / "unused.json"
        command = self._build_command(Path(info_obj.path), cli_argv, params_json, cwd)
        return self.run_argv(command, cwd=cwd)


# ----------------------------------------------------------------------- оркестрация


def run_and_finalize(runner: ScriptRunner, *, info: plugin_loader.PluginInfo | dict,
                     meta: dict, rules: dict, profile: dict,
                     input_path: Path | str | None = None,
                     output_path: Path | str | None = None,
                     params: dict | None = None, cwd: str | Path | None = None,
                     do_normalize: bool = True, do_check: bool = True,
                     fail_on: str = "error") -> dict[str, Any]:
    """Полный цикл: запуск → (нормализация по rules_hook) → (проверка) → результат.

    Ответственность за последовательность пост-хуков — здесь (п. 9 ТЗ): сам
    ``run_script`` только запускает процесс. docx_normalizer/docx_inspector
    импортируются лениво: на этапе параллельной сборки модулей их ещё может не быть —
    это фиксируется в результате как ``{"skipped": ...}``, а не падением.
    """
    run_res = runner.run_script(info, meta=meta, rules=rules, profile=profile,
                                input_path=input_path, output_path=output_path,
                                params=params, cwd=cwd)
    out_path = Path(output_path) if output_path else None
    outcome: dict[str, Any] = {
        "run": run_res,
        "normalized": None,
        "check": None,
        "output": str(out_path) if out_path else None,
        "ok": run_res.ok,
    }
    if not run_res.ok:
        return outcome

    info_obj = (plugin_loader.PluginInfo.from_meta(Path(str(info.get("path", ""))), info)
                if isinstance(info, dict) else info)

    # П. 9: нормализатор запускаем для rules_hook == postprocess при включённой нормализации.
    if do_normalize and info_obj.rules_hook == "postprocess" and out_path and out_path.exists():
        try:
            from app.core.docx_normalizer import normalize_docx  # ленивый импорт (агент C)
        except ImportError:
            outcome["normalized"] = {"skipped": "normalizer unavailable"}
        else:
            try:
                outcome["normalized"] = normalize_docx(out_path, rules)
            except Exception as exc:  # нормализация не должна ронять цикл
                outcome["normalized"] = {"error": str(exc)}
                outcome["ok"] = False

    if do_check and out_path and out_path.exists():
        try:
            from app.core.docx_inspector import check_docx_result  # ленивый импорт (агент C)
        except ImportError:
            outcome["check"] = {"skipped": "inspector unavailable"}
        else:
            try:
                outcome["check"] = check_docx_result(out_path, rules)
            except Exception as exc:
                outcome["check"] = {"error": str(exc)}
                outcome["ok"] = False

    # fail_on: порог, при котором наличие замечаний помечает операцию неуспешной.
    check = outcome["check"]
    if check and "summary" in check and fail_on != "never":
        summary = check["summary"]
        violated = summary.get("error", 0) > 0
        if fail_on == "warning":
            violated = violated or summary.get("warning", 0) > 0
        if violated:
            outcome["ok"] = False
    return outcome
