"""Лог-панель (п. 10): журнал реального времени поверх очереди из рабочих потоков.

Потоки никогда не трогают Tk-виджеты: они кладут строки в `queue.Queue`,
а панель вычитает её циклом `after()` в главном потоке (п. 12 «UI не фризится»).
"""

from __future__ import annotations

import queue
import sys
import tkinter as tk
from typing import Any, Callable

import customtkinter as ctk

from app import constants as C
from app.strings import t

# Метрики панели — именованные константы слоя UI.
POLL_INTERVAL_MS = 80          # частота вычитки очереди в главном потоке
STDERR_TAG = "stderr"
STDOUT_TAG = "stdout"
SYSTEM_TAG = "system"
STDERR_COLOR = "#e8a0a0"
STDOUT_COLOR = "#c8c8c8"
SYSTEM_COLOR = "#8fb8e8"
COLLAPSED_HEIGHT = 36          # полоска-заголовок в свёрнутом виде
MAX_LINES_KEEP = 5000          # защита от разрастания Text при шумных скриптах
TRIM_STEP = 1000               # сколько строк отрезаем разом при переполнении
BTN_W = 110
BAR_H = 4
LOG_FONT_SIZE = 12
PROGRESS_W = 140
PANEL_CHROME = 46          # полоса заголовка + отступы вокруг textbox
TEXT_MIN_HEIGHT = 48       # ниже двух-трёх строк журнал читать нельзя

_MONO_FAMILIES = {
    "win32": "Consolas",
    "darwin": "Menlo",
}


def mono_font(size: int) -> ctk.CTkFont:
    """Моноширинный шрифт журнала: семейство берётся по ОС, кегль — параметр."""
    family = _MONO_FAMILIES.get(sys.platform, "DejaVu Sans Mono")
    try:
        return ctk.CTkFont(family=family, size=size)
    except Exception:  # Tk может не знать семейство — fallback системным
        return ctk.CTkFont(size=size)


class LogPanel(ctk.CTkFrame):
    """Нижняя панель журнала: textbox + кнопки + прогресс + «Отменить».

    `session_sink(line, kind)` — колбэк GostApp для записи сессии в файл (п. 9).
    """

    def __init__(self, master: Any, height: int = C.LOG_PANEL_HEIGHT,
                 on_cancel: Callable[[], None] | None = None):
        super().__init__(master)
        self._queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self._running = False
        self._collapsed = False
        self._full_height = height
        # CTkTextbox по умолчанию просит 200 px и раздувает панель независимо от `height`;
        # без этого на низком экране журнал отнимал место у нижних кнопок сайдбара.
        self._text_height = max(TEXT_MIN_HEIGHT, height - PANEL_CHROME)
        self._session_sink: Callable[[str, str], None] | None = None
        self._on_cancel = on_cancel
        self._dirty_lines = 0

        self.header = ctk.CTkFrame(self, fg_color="transparent")
        self.header.pack(fill="x", padx=6, pady=(4, 0))
        ctk.CTkLabel(self.header, text=t("log_title"),
                     font=ctk.CTkFont(weight="bold")).pack(side="left")
        self.status_lbl = ctk.CTkLabel(self.header, text=t("status_idle"), anchor="e")
        self.status_lbl.pack(side="right", padx=6)
        self.progress = ctk.CTkProgressBar(self.header, width=PROGRESS_W, height=BAR_H,
                                           mode="indeterminate")
        self.progress.pack(side="right", padx=6)
        self.progress.stop()
        self.cancel_btn = ctk.CTkButton(self.header, text=t("log_cancel"), width=BTN_W // 2,
                                        fg_color="#8a3b3b", state="disabled",
                                        command=self._cancel_clicked)
        self.cancel_btn.pack(side="right", padx=6)
        btns = ctk.CTkFrame(self.header, fg_color="transparent")
        btns.pack(side="left", padx=12)
        for key, cmd in (("log_copy", self.copy_all), ("log_clear", self.clear),
                         ("log_save", self.save_to_file)):
            ctk.CTkButton(btns, text=t(key), width=BTN_W // 2, fg_color="gray30",
                          command=cmd).pack(side="left", padx=3)

        self.textbox = ctk.CTkTextbox(self, height=self._text_height, wrap="none",
                                      font=mono_font(LOG_FONT_SIZE),
                                      activate_scrollbars=True)
        self.textbox.pack(fill="both", expand=True, padx=6, pady=(2, 6))
        self._configure_tags()
        self.textbox.configure(state="disabled")

        self._poll()

    # -- внешний API -----------------------------------------------------------
    def set_session_sink(self, sink: Callable[[str, str], None] | None) -> None:
        self._session_sink = sink

    def set_cancel_callback(self, cb: Callable[[], None] | None) -> None:
        self._on_cancel = cb

    def log(self, text: str, kind: str = "stdout") -> None:
        """Вызывается из ЛЮБОГО потока: только постановка в очередь."""
        self._queue.put((kind, text))

    def log_sys(self, text: str) -> None:
        self.log(text, kind=SYSTEM_TAG)

    def get_text(self) -> str:
        return self.textbox.get("1.0", "end").rstrip("\n")

    def set_running(self, running: bool, status: str | None = None) -> None:
        # Поток-безопасно: планируем изменение UI через очередь-мост.
        self._running = running
        self._queue.put(("__ui__", "running" if running else "idle"))
        if status:
            self.set_status(status)

    def set_status(self, status: str) -> None:
        self._queue.put(("__ui__", f"status:{status}"))

    # -- действия кнопок ---------------------------------------------------------
    def copy_all(self) -> None:
        toplevel = self.winfo_toplevel()
        toplevel.clipboard_clear()
        toplevel.clipboard_append(self.get_text())
        self.log_sys(t("log_copied"))

    def clear(self) -> None:
        self.textbox.configure(state="normal")
        self.textbox.delete("1.0", "end")
        self.textbox.configure(state="disabled")

    def save_to_file(self) -> None:
        from app.ui import dialogs  # лениво: диалоги тянут filedialog
        path = dialogs.choose_save_path(
            self.winfo_toplevel(), t("log_save"), defaultextension=".log",
            initialfile="gost_log.log", filetypes=(("LOG", "*.log"),) + dialogs.FILETYPES_ALL)
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8", errors="replace") as fh:
                fh.write(self.get_text())
            self.log_sys(t("log_saved", path=path))
        except OSError as exc:
            self.log_sys(t("open_error", path=path, why=exc))

    def toggle_collapse(self) -> None:
        """Ctrl+L: сворачивание панели до полоски заголовка (п. 10)."""
        self._collapsed = not self._collapsed
        if self._collapsed:
            self.textbox.pack_forget()
            self.configure(height=COLLAPSED_HEIGHT)
            self.pack_propagate(False)
        else:
            self.configure(height=self._full_height)
            self.pack_propagate(True)
            self.textbox.pack(fill="both", expand=True, padx=6, pady=(2, 6))

    # -- внутреннее --------------------------------------------------------------
    def _configure_tags(self) -> None:
        inner = getattr(self.textbox, "textbox", None)  # tk.Text внутри CTkTextbox
        if inner is None:
            return
        try:
            inner.tag_configure(STDERR_TAG, foreground=STDERR_COLOR)
            inner.tag_configure(STDOUT_TAG, foreground=STDOUT_COLOR)
            inner.tag_configure(SYSTEM_TAG, foreground=SYSTEM_COLOR)
        except tk.TclError:
            pass  # нет поддержки тегов — пишем монохромно, не роняем UI

    def _cancel_clicked(self) -> None:
        if self._on_cancel:
            self._on_cancel()
        self.cancel_btn.configure(state="disabled")

    def _poll(self) -> None:
        """Вычитка очереди порциями в главном потоке; повтор через POLL_INTERVAL_MS."""
        batch: list[tuple[str, str]] = []
        try:
            while True:
                batch.append(self._queue.get_nowait())
        except queue.Empty:
            pass
        if batch:
            self._append_batch(batch)
        self.after(POLL_INTERVAL_MS, self._poll)

    def _append_batch(self, batch: list[tuple[str, str]]) -> None:
        self.textbox.configure(state="normal")
        # тегированная вставка идёт через внутренний tk.Text (CTkTextbox.insert
        # не гарантирует проброс tags во всех версиях CustomTkinter)
        inner = getattr(self.textbox, "textbox", None)
        try:
            for kind, line in batch:
                if kind == "__ui__":
                    self._apply_ui_command(line)
                    continue
                marker = t("log_stderr_mark") if kind == "stderr" else ""
                tag = STDERR_TAG if kind == "stderr" else (
                    SYSTEM_TAG if kind in (SYSTEM_TAG, "system") else STDOUT_TAG)
                payload = marker + line.rstrip("\n") + "\n"
                if inner is not None:
                    try:
                        inner.insert("end", payload, (tag,))
                    except tk.TclError:
                        self.textbox.insert("end", payload)
                else:
                    self.textbox.insert("end", payload)
                self._dirty_lines += 1
                if self._session_sink:
                    try:
                        self._session_sink(line, kind)
                    except OSError:
                        self._session_sink = None  # диск/прав не хватает — пишем только в UI
            self._trim_if_needed()
            self.textbox.see("end")  # автоскролл
        finally:
            self.textbox.configure(state="disabled")

    def _apply_ui_command(self, payload: str) -> None:
        if payload in ("running", "idle"):
            running = payload == "running"
            if running:
                self.progress.start()
                self.status_lbl.configure(text=t("status_running"))
                self.cancel_btn.configure(state="normal")
            else:
                self.progress.stop()
                self.status_lbl.configure(text=t("status_idle"))
                self.cancel_btn.configure(state="disabled")
        elif payload.startswith("status:"):
            self.status_lbl.configure(text=payload[len("status:"):])

    def _trim_if_needed(self) -> None:
        if self._dirty_lines < MAX_LINES_KEEP:
            return
        self._dirty_lines = 0
        inner = getattr(self.textbox, "textbox", None)
        if inner is None:
            return  # без прямого доступа к tk.Text индекс не прочитать — не рискуем
        try:
            total = int(inner.index("end").split(".")[0])
            if total > MAX_LINES_KEEP:
                self.textbox.delete("1.0", f"{total - MAX_LINES_KEEP + TRIM_STEP}.0")
        except (tk.TclError, ValueError):
            pass
