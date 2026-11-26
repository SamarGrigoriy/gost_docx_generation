"""Исполняющий микс вкладки «Генерация» (запуск/отмена/результат).

Вынесено из generator_tab по конвенции «файлы > 600 строк разбивать»
(INTERFACES.md). Вся тяжёлая работа — вызовы `app.core.runner/run_and_finalize`,
`docx_normalizer`, `docx_inspector`; последовательность пост-операций инкапсулирует
core, UI лишь собирает ввод и отображает конверт результата.
"""

from __future__ import annotations

import inspect
import os
import threading
from typing import Any

from app import constants as C
from app.core import docx_inspector, docx_normalizer, runner
from app.strings import t
from app.ui import dialogs


class GeneratorRunMixin:
    """Методы запуска для GeneratorTab (общие атрибуты: app/selected/_busy/...)."""

    # ------------------------------------------------------------------
    # Запуск (поток + очередь, п. 10.1/12)
    # ------------------------------------------------------------------
    def generate(self) -> None:
        if self._busy:
            self.log(t("gen_running"))
            return
        errors = self.validate()
        if errors:
            self.errors_lbl.configure(text="\n".join(errors))
            return
        out = self.out_row.get()
        try:
            os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
        except OSError as exc:
            self.errors_lbl.configure(text=t("gen_err_output_dir") + f": {exc}")
            return
        if os.path.exists(out) and not dialogs.confirm_overwrite(self, out):
            return
        info = self.selected
        meta = self._flat_meta() if self.use_profile.get() else {}
        kwargs = {
            "meta": meta,
            "rules": self.app.rules,
            "profile": self.app.profile,
            "input_path": self.uni_input.get() or None,
            "output_path": out,
            "params": self.collect_params(),
        }
        self._busy = True
        self.run_btn.configure(state="disabled")
        self.set_result_badge("", "muted")
        self.app.log.set_running(True, t("gen_running"))
        self._thread = threading.Thread(
            target=self.app.safe(self._worker), args=(info, kwargs), daemon=True)
        self._thread.start()

    def _worker(self, info: Any, kwargs: dict) -> None:
        self._runner = runner.ScriptRunner(
            on_line=self._on_line, kill_grace_s=C.KILL_GRACE_PERIOD_S,
            **self.app.runner_kwargs())
        try:
            result = self._run_and_finalize(info, kwargs)
        except Exception as exc:
            result = {"run": None, "error": exc}
        self.app.post(lambda: self._on_done(result, kwargs["output_path"]))

    def _on_line(self, line: Any) -> None:
        """Колбэк раннера: str или (kind, text) — оба варианта в журнал."""
        if isinstance(line, tuple) and len(line) == 2:
            self.app.log.log(str(line[1]), kind=str(line[0]))
        else:
            self.app.log.log(str(line))

    def _run_and_finalize(self, info: Any, kwargs: dict) -> dict:
        """Оркестрация run→normalize→check инкапсулирована в core (run_and_finalize).

        Сигнатура контрактная (`run_and_finalize(...) -> dict`): диспатч по первым
        параметрам — method/функция с runner-аргументом/без. Если модуль ещё не
        предоставляет функцию (например, старая копия), выполняется эквивалентная
        последовательность вызовов core — без «логики» в UI.
        """
        post = {"normalize": bool(self.cb_apply_rules.get()),
                "check": bool(self.cb_check_after.get()), "fail_on": "error"}
        fn = getattr(runner, "run_and_finalize", None)
        if fn is not None:
            try:
                first = list(inspect.signature(fn).parameters)
            except (TypeError, ValueError):
                first = []
            args = (self._runner, info) if first and first[0] in ("runner", "self") else (info,)
            try:
                return fn(*args, post=post, **kwargs)
            except TypeError as exc:
                if "post" not in str(exc):
                    raise
        result: dict[str, Any] = {"run": self._runner.run_script(info, **kwargs),
                                  "normalized": None, "check": None,
                                  "output": kwargs.get("output_path")}
        out = kwargs.get("output_path")
        if result["run"].ok and post["normalize"] and out and os.path.exists(out):
            result["normalized"] = docx_normalizer.normalize_docx(
                out, self.app.rules, backup=True)
        if result["run"].ok and post["check"] and out and os.path.exists(out):
            result["check"] = docx_inspector.check_docx_result(out, self.app.rules)
        return result

    def cancel(self) -> None:
        if self._runner is not None and self._busy:
            self._runner.cancel()

    # ------------------------------------------------------------------
    # Отображение результата
    # ------------------------------------------------------------------
    def _on_done(self, result: dict, output_path: str) -> None:
        self._busy = False
        self.app.log.set_running(False)
        self._runner_ref_keep = self._runner  # раннер живёт до отмены следующего запуска
        run = result.get("run")
        if run is None:
            self._finish_failed(result.get("error"))
            return
        if getattr(run, "timed_out", False):
            self.log(t("log_timeout", timeout=self.app.runner_kwargs()["timeout_s"]))
        if getattr(run, "cancelled", False):
            self.set_result_badge(t("gen_cancelled"), "warning")
            self.log(t("gen_cancelled"))
            self.run_btn.configure(state="normal")
            return
        if not getattr(run, "ok", getattr(run, "exit_code", 1) == 0):
            message = t("gen_failed", code=getattr(run, "exit_code", "?"))
            self.log(message)
            self.set_result_badge(message, kind="error")
            # П. 15.5: падение скрипта показывается диалогом с хвостом вывода —
            # молча повесить красный бейдж при закрытой панели журнала недостаточно.
            tail = "\n".join(str(line) for line in (getattr(run, "lines", None) or [])[-12:])
            dialogs.show_info(self, message, tail)
            self.run_btn.configure(state="normal")
            return
        if not os.path.exists(output_path):
            self.log(t("gen_output_missing", path=output_path))
            self.set_result_badge(t("gen_output_missing", path="…"), kind="error")
            self.run_btn.configure(state="normal")
            return

        self.app.add_recent(output_path)
        normalized = result.get("normalized") or result.get("normalize")
        if normalized:
            changes = normalized.get("changes") or []
            self.log(t("gen_normalized", n=len(changes)) if changes
                     else t("gen_normalized_none"))
            for ch in changes:
                self.log(f"  · {ch}")
            if normalized.get("backup"):
                self.log(t("gen_norm_backup", path=normalized["backup"]))
        check = result.get("check")
        n_issues = 0
        if check:
            summary = check.get("summary") or {}
            n_issues = int(summary.get("error", 0)) + int(summary.get("warning", 0)) \
                + int(summary.get("info", 0))
            self.log(t("chk_summary", **{k: summary.get(k, 0)
                                         for k in ("error", "warning", "info")}))
        self.set_result_badge(
            t("gen_done_ok") if n_issues == 0 else t("gen_done_issues", n=n_issues),
            kind="ok" if n_issues == 0 else "warning")
        if self.cb_open_after.get():
            err = dialogs.open_path(output_path)
            if err:
                self.log(t("open_error", path=output_path, why=err))
        self.run_btn.configure(state="normal")
        self._validate_ui()

    def _finish_failed(self, exc: Any) -> None:
        self.set_result_badge(t("err_generic", why=exc or "?"), kind="error")
        self.log(t("err_generic", why=str(exc or "")))
        self.run_btn.configure(state="normal")
