"""Вкладка «Проверка» (п. 10.3).

Движок замечаний — `app.core.docx_inspector`, отчёты — `app.core.report`,
исправление — `app.core.docx_normalizer`. Вкладка только запускает core в
рабочем потоке и отображает результат; пакетный режим не падает на битых .docx
(read_error/DocxError → статус «Ошибка чтения»).
"""

from __future__ import annotations

import glob
import os
import threading
from typing import Any

import customtkinter as ctk

from app import constants as C
from app.core import docx_inspector, docx_normalizer, report
from app.strings import t
from app.ui import dialogs
from app.ui.widgets import Badge, FilePickerRow, IssueTable, SectionAccordion

# Метрики вкладки.
BTN_W = 150
ROW_H = 34
CELL_PADY = 8
BATCH_ROWS_H = 160
ROW_RADIUS = 6
BATCH_COL_W = {"file": 340, "err": 90, "warn": 140, "status": 150}
FIX_ROUNDS = 2  # проверка → нормализация → повторная проверка (п. 10.3)
BATCH_GLOB = "*.docx"


def _manual_items() -> list:
    """Чек-лист «Проверить вручную» из inspector (код MANUAL, п. 5.9/9.1)."""
    items = getattr(docx_inspector, "_MANUAL_ITEMS", None) or \
        getattr(docx_inspector, "MANUAL_ITEMS", None) or []
    return list(items)


def _summary_of(result: dict) -> dict:
    summary = result.get("summary") or {}
    return {k: int(summary.get(k, 0)) for k in C.SEVERITIES}


def _as_list(result: Any, key: str) -> list:
    value = result.get(key) if isinstance(result, dict) else getattr(result, key, [])
    return list(value or [])


class CheckerTab(ctk.CTkFrame):
    def __init__(self, master: Any, app: Any):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self._busy = False
        self._cancelled = False
        self._last_result: dict | None = None      # одиночная проверка
        self._batch_items: list[dict] = []          # пакетная сводка

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)
        self._build_controls()
        self._build_results()
        self._build_manual()

    # ------------------------------------------------------------------
    # Управление
    # ------------------------------------------------------------------
    def _build_controls(self) -> None:
        top = ctk.CTkFrame(self)
        top.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 0))
        top.grid_columnconfigure(1, weight=1)

        self.mode_var = ctk.StringVar(master=self, value=t("chk_mode_file"))
        for col, label_key in enumerate(("chk_mode_file", "chk_mode_folder")):
            ctk.CTkRadioButton(top, text=t(label_key), variable=self.mode_var,
                               value=t(label_key), command=self._on_mode_change).grid(
                row=0, column=col, padx=8, pady=8, sticky="w")

        self.file_row = FilePickerRow(top, label=t("chk_file"), kind="file",
                                      filetypes=dialogs.FILETYPES_DOCX,
                                      title=t("chk_pick_file"))
        self.file_row.grid(row=1, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
        self.folder_row = FilePickerRow(top, label=t("chk_folder"), kind="folder",
                                        title=t("chk_pick_folder"))
        self.folder_row.grid(row=2, column=0, columnspan=2, sticky="ew", padx=8, pady=3)
        self.folder_row.grid_remove()

        buttons = ctk.CTkFrame(self, fg_color="transparent")
        buttons.grid(row=1, column=0, sticky="ew", padx=8, pady=6)
        self.check_btn = ctk.CTkButton(buttons, text=t("btn_check"), width=BTN_W,
                                       command=self.app.safe(self.run_check))
        self.check_btn.pack(side="left", padx=3)
        ctk.CTkButton(buttons, text=t("chk_fix"), width=BTN_W, fg_color="gray40",
                      command=self.app.safe(self.run_fix)).pack(side="left", padx=3)
        ctk.CTkButton(buttons, text=t("chk_save_md"), width=BTN_W, fg_color="gray40",
                      command=self.app.safe(lambda: self.save_report("md"))).pack(
            side="left", padx=3)
        ctk.CTkButton(buttons, text=t("chk_save_txt"), width=BTN_W, fg_color="gray40",
                      command=self.app.safe(lambda: self.save_report("txt"))).pack(
            side="left", padx=3)
        ctk.CTkButton(buttons, text=t("chk_save_summary"), width=BTN_W, fg_color="gray40",
                      command=self.app.safe(self.save_summary)).pack(side="left", padx=3)
        ctk.CTkButton(buttons, text=t("chk_open_file"), width=BTN_W, fg_color="gray40",
                      command=self.app.safe(self.open_file)).pack(side="left", padx=3)

    def _on_mode_change(self) -> None:
        folder = self.folder_mode
        self.file_row.grid() if not folder else self.file_row.grid_remove()
        self.folder_row.grid_remove() if not folder else self.folder_row.grid()
        # Пустая пакетная таблица в одиночном режиме только съедает место замечаний.
        self.batch_frame.grid() if folder else self.batch_frame.grid_remove()

    # ------------------------------------------------------------------
    # Результаты: сводка + таблица замечаний + пакетная таблица
    # ------------------------------------------------------------------
    def _build_results(self) -> None:
        bar = ctk.CTkFrame(self, fg_color="transparent")
        bar.grid(row=2, column=0, sticky="ew", padx=8)
        self.badges: dict[str, Badge] = {}
        for sev in C.SEVERITIES:
            badge = Badge(bar, t(f"severity_{sev}") + ": 0", kind=sev)
            badge.pack(side="left", padx=4)
            self.badges[sev] = badge
        self.total_badge = Badge(bar, t("chk_empty"), kind="muted")
        self.total_badge.pack(side="right", padx=4)
        # Прогресс в ряду сводки: в тесном ряду кнопок он обрезался.
        self.progress = ctk.CTkProgressBar(bar, width=BTN_W)
        self.progress.pack(side="right", padx=8)
        self.progress.set(0)

        self.issue_table = IssueTable(self, on_detail=self.app.safe(self._show_detail))
        self.issue_table.grid(row=3, column=0, sticky="nsew", padx=8, pady=4)

        self.batch_frame = ctk.CTkFrame(self)
        self.batch_frame.grid(row=4, column=0, sticky="nsew", padx=8, pady=(0, 8))
        self.batch_frame.grid_columnconfigure(0, weight=1)
        ctk.CTkLabel(self.batch_frame, text=t("chk_mode_folder"), anchor="w",
                     font=ctk.CTkFont(weight="bold")).grid(row=0, column=0, sticky="w", padx=6)
        header = ctk.CTkFrame(self.batch_frame, fg_color="transparent")
        header.grid(row=1, column=0, sticky="ew")
        for col, (key, width) in enumerate((
                ("chk_col_file", BATCH_COL_W["file"]), ("chk_col_errors", BATCH_COL_W["err"]),
                ("chk_col_warnings", BATCH_COL_W["warn"]),
                ("chk_col_status", BATCH_COL_W["status"]))):
            ctk.CTkLabel(header, text=t(key), width=width, anchor="w").grid(row=0, column=col,
                                                                            padx=6)
        self.batch_rows = ctk.CTkScrollableFrame(self.batch_frame, label_text="", height=BATCH_ROWS_H)
        self.batch_rows.grid(row=2, column=0, sticky="nsew", padx=4, pady=4)
        self.batch_frame.grid_rowconfigure(2, weight=1)
        # Замечания важнее пустой пакетной таблицы: отдаём им вдвое больше места.
        self.grid_rowconfigure(3, weight=2)
        self.grid_rowconfigure(4, weight=1)
        self._on_mode_change()

    # ------------------------------------------------------------------
    # Блок «Проверить вручную» (п. 10.3)
    # ------------------------------------------------------------------
    def _build_manual(self) -> None:
        self.manual_acc = SectionAccordion(self, t("chk_manual"), expanded=False)
        self.manual_acc.grid(row=5, column=0, sticky="ew", padx=8, pady=(0, 8))
        ctk.CTkLabel(self.manual_acc.body, text=t("chk_manual_intro"), anchor="w",
                     justify="left").pack(anchor="w", pady=2)
        for _, description in _manual_items():
            ctk.CTkCheckBox(self.manual_acc.body, text=str(description)).pack(
                anchor="w", pady=2)

    # ------------------------------------------------------------------
    # Запуск проверок
    # ------------------------------------------------------------------
    @property
    def folder_mode(self) -> bool:
        return self.mode_var.get() == t("chk_mode_folder")

    def run_check(self) -> None:
        self._start(batch=False, fix=False)

    def run_fix(self) -> None:
        self._start(batch=False, fix=True)

    def cancel(self) -> None:
        self._cancelled = True  # флаг останавливает перебор пакетного режима

    def _start(self, batch: bool, fix: bool) -> None:
        if self._busy:
            return
        if batch or self.folder_mode:
            folder = self.folder_row.get()
            files = sorted(glob.glob(os.path.join(folder, BATCH_GLOB))) if folder else []
            if not files:
                self.app.log.log_sys(t("gen_no_scripts"))
                return
            self._run_batch(files)
            return
        path = self.file_row.get()
        if not path or not os.path.exists(path):
            self.app.log.log_sys(t("gen_err_input_missing"))
            return
        self._busy = True
        self._cancelled = False
        self.check_btn.configure(state="disabled")
        self.app.log.set_running(True, t("chk_running"))

        def worker() -> None:
            try:
                out = self._check_single_worker(path, fix)
            except Exception as exc:  # показываем в диалоге, не роняем UI
                out = {"error": exc}
            self.app.post(lambda: self._on_single_done(out))

        threading.Thread(target=worker, daemon=True).start()

    def _check_single_worker(self, path: str, fix: bool) -> dict:
        rules = self.app.rules
        result = docx_inspector.check_docx_result(path, rules)
        before = _summary_of(result)["error"] if not result.get("read_error") else -1
        normalize_info: dict | None = None
        if fix and not result.get("read_error"):
            for round_no in range(1, FIX_ROUNDS + 1):
                errors = _summary_of(result)["error"]
                if errors == 0:
                    break
                self.app.log.log(t("chk_fix_round", i=round_no))
                normalize_info = docx_normalizer.normalize_docx(path, rules, backup=True)
                result = docx_inspector.check_docx_result(path, rules)
        return {"result": result, "path": path, "errors_before": before,
                "normalize": normalize_info}

    def _on_single_done(self, out: dict) -> None:
        self._busy = False
        self.app.log.set_running(False)
        self.check_btn.configure(state="normal")
        if out.get("error"):
            dialogs.exception_dialog(self, t("err_generic", why=out["error"]), repr(out["error"]))
            return
        result = out["result"]
        self._last_result = result
        self._batch_items = []
        self._render_issue_result(result)
        if out.get("normalize"):
            changes = out["normalize"].get("changes") or []
            self.app.log.log_sys(t("chk_fix_done", before=out.get("errors_before", 0),
                                   after=_summary_of(result)["error"]))
            for ch in changes:
                self.app.log.log(f"  · {ch}")
        self.app.add_recent(result.get("file") or "")

    def _render_issue_result(self, result: dict) -> None:
        summary = _summary_of(result)
        for sev, badge in self.badges.items():
            badge.set_text(f"{t(f'severity_{sev}')}: {summary.get(sev, 0)}", sev)
        if result.get("read_error"):
            self.total_badge.set_text(t("chk_read_error"), kind="error")
            self.issue_table.clear()
            self.app.log.log_sys(t("chk_read_error") + f": {result['read_error']}")
            return
        total = sum(summary.values())
        self.total_badge.set_text(t("chk_summary", **summary),
                                  kind="ok" if total == 0 else "warning")
        self.issue_table.set_issues(_as_list(result, "issues"))

    # ------------------------------------------------------------------
    # Пакетный режим
    # ------------------------------------------------------------------
    def _run_batch(self, files: list[str]) -> None:
        self._busy = True
        self._cancelled = False
        self.check_btn.configure(state="disabled")
        self.app.log.set_running(True, t("chk_running"))
        self.app.log.log_sys(t("chk_files_found", n=len(files)))

        def worker() -> None:
            rules = self.app.rules
            items: list[dict] = []
            for i, path in enumerate(files):
                if self._cancelled:
                    break
                try:
                    res = docx_inspector.check_docx_result(path, rules)
                    summary = _summary_of(res)
                    if res.get("read_error"):
                        status = t("chk_read_error")
                    elif summary["error"] or summary["warning"]:
                        status = t("chk_status_issues")
                    else:
                        status = t("chk_status_ok")
                    items.append({"file": os.path.basename(path), "path": path,
                                  "summary": summary, "status": status, "result": res})
                except docx_inspector.DocxError as exc:
                    items.append({"file": os.path.basename(path), "path": path,
                                  "summary": {s: 0 for s in C.SEVERITIES},
                                  "status": t("chk_read_error"), "result": None,
                                  "error": str(exc)})
                except Exception as exc:
                    items.append({"file": os.path.basename(path), "path": path,
                                  "summary": {s: 0 for s in C.SEVERITIES},
                                  "status": t("chk_read_error"), "result": None,
                                  "error": repr(exc)})
                progress = (i + 1) / len(files)
                self.app.post(lambda p=progress: self.progress.set(p))
            self.app.post(lambda: self._on_batch_done(items))

        threading.Thread(target=worker, daemon=True).start()

    def _on_batch_done(self, items: list[dict]) -> None:
        self._busy = False
        self._batch_items = items
        self.app.log.set_running(False)
        self.check_btn.configure(state="normal")
        self.progress.set(1.0)
        for child in self.batch_rows.winfo_children():
            child.destroy()
        agg = {s: 0 for s in C.SEVERITIES}
        for r, item in enumerate(items):
            agg["error"] += item["summary"].get("error", 0)
            agg["warning"] += item["summary"].get("warning", 0)
            agg["info"] += item["summary"].get("info", 0)
            frame = ctk.CTkFrame(self.batch_rows, corner_radius=ROW_RADIUS)
            frame.grid(row=r, column=0, sticky="ew", pady=2)
            cells = (
                (item["file"], BATCH_COL_W["file"], None),
                (str(item["summary"].get("error", 0)), BATCH_COL_W["err"], None),
                (str(item["summary"].get("warning", 0)), BATCH_COL_W["warn"], None),
            )
            for col, (text_, width, _x) in enumerate(cells):
                ctk.CTkLabel(frame, text=text_, width=width, anchor="w").grid(
                    row=0, column=col, padx=6, pady=CELL_PADY)
            kind = "error" if item["status"] == t("chk_read_error") else (
                "warning" if item["status"] == t("chk_status_issues") else "ok")
            badge = Badge(frame, item["status"], kind=kind)
            badge.grid(row=0, column=3, padx=6, pady=CELL_PADY)
            for w in (frame, *frame.winfo_children()):
                w.bind("<Double-Button-1>",
                       lambda _e, it=item: self._open_batch_detail(it))
        self.total_badge.set_text(t("chk_summary", **agg),
                                  kind="ok" if agg["error"] == 0 else "warning")
        self.issue_table.clear()

    def _open_batch_detail(self, item: dict) -> None:
        """Двойной клик по строке пакета — детальный отчёт файла в таблице замечаний."""
        result = item.get("result")
        if result is None:
            self.app.log.log_sys(t("chk_read_error") + f": {item.get('error', '')}")
            return
        self._last_result = result
        self._render_issue_result(result)

    def _show_detail(self, issue: dict) -> None:
        location = issue.get("location") or {}
        snippet = str(location.get("snippet") or "")
        text = (f"{issue.get('code')} [{issue.get('severity')}]\n{issue.get('message')}\n\n"
                f"{t('chk_issue_hint')}: {issue.get('hint') or '—'}\n\n"
                f"{t('chk_issue_snippet')}:\n{snippet or '—'}")
        dialogs.show_info(self, t("chk_detail"), text)

    # ------------------------------------------------------------------
    # Отчёты (core.report)
    # ------------------------------------------------------------------
    def save_report(self, fmt: str) -> None:
        if not self._last_result:
            self.app.log.log_sys(t("chk_empty"))
            return
        ext = C.REPORT_MD_EXT if fmt == "md" else C.REPORT_TXT_EXT
        name = os.path.splitext(os.path.basename(str(self._last_result.get("file")
                                                  or "report"))) [0] + ext
        dest = dialogs.choose_save_path(self, t("chk_save_md"),
                                        defaultextension=ext, initialfile=name,
                                        filetypes=((t("filetypes_all"), "*" + ext),),
                                        initialdir=self.app.output_dir())
        if not dest:
            return
        path = report.save_report(self._last_result, self.app.rules, dest, fmt=fmt)
        self.app.log.log_sys(t("chk_report_saved", path=path))

    def save_summary(self) -> None:
        if not self._batch_items:
            self.app.log.log_sys(t("chk_empty"))
            return
        dest = dialogs.choose_save_path(self, t("chk_save_summary"),
                                        defaultextension=C.REPORT_MD_EXT,
                                        initialfile="summary" + C.REPORT_MD_EXT,
                                        filetypes=dialogs.FILETYPES_ALL,
                                        initialdir=self.app.output_dir())
        if not dest:
            return
        text = report.render_batch_markdown([
            {k: v for k, v in item.items() if k != "path"} for item in self._batch_items])
        with open(dest, "w", encoding="utf-8", errors="replace") as fh:
            fh.write(text)
        self.app.log.log_sys(t("chk_summary_saved", path=dest))

    def open_file(self) -> None:
        path = self.file_row.get() or (self._last_result or {}).get("file") or ""
        if not path:
            return
        err = dialogs.open_path(path)
        self.app.log.log_sys(t("open_error", path=path, why=err) if err
                             else t("open_done", path=path))
