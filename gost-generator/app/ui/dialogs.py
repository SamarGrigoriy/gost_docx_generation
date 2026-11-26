"""Диалоги и выбор файлов (п. 10.5, 12 ТЗ). Модуль не импортирует widgets —
иначе получится цикл (widgets.FilePickerRow вызывает отсюда pick_file)."""

from __future__ import annotations

import os
import platform
import shutil
import subprocess
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from typing import Any, Callable

import customtkinter as ctk

from app import constants as C
from app.strings import t

# Метрики оформления диалогов — именованные, не магические числа в теле.
DLG_WIDTH = 520
SETTINGS_WIDTH = 640
SETTINGS_HEIGHT = 560
ABOUT_WIDTH = 560
ABOUT_HEIGHT = 420
WELCOME_WIDTH = 620
WELCOME_HEIGHT = 420
EXC_WIDTH = 760
EXC_HEIGHT = 520
WELCOME_TITLE_SIZE = 18
ROW_PAD = 8
BTN_W = 150

FILETYPES_DOCX = ((t("filetypes_docx"), "*.docx"),)
FILETYPES_PY = ((t("filetypes_py"), "*.py"),)
FILETYPES_JSON = ((t("filetypes_json"), "*.json"),)
FILETYPES_ALL = ((t("filetypes_all"), "*.*"),)
FILETYPES_DOCX_ALL = FILETYPES_DOCX + FILETYPES_ALL
# Сколько ждать ответа системного просмотрщика: дольше — значит окно уже открывается.
OPEN_VERIFY_TIMEOUT_S = 1.5


# ---------------------------------------------------------------------------
# Абстракция «источник файла» — задел под drag&drop v2 (п. 13 ТЗ).
#
# В v1 единственный источник — системный диалог (Tk DnD на CustomTkinter без
# сторонних библиотек ненадёжен); решение зафиксировано в DECISIONS.md
# (п. 10.3 ТЗ: «иначе — только диалог, зафиксировать в DECISIONS.md»).
# Для v2 достаточно реализовать FileSource c тем же API (TkinterDnD,
# глобальный drop-target) и назначить `active_file_source` — вкладки менять не
# придётся.
# ---------------------------------------------------------------------------
class FileSource:
    """Интерфейс источника пути к файлу/папке."""

    def pick_file(self, parent: Any, title: str, filetypes: tuple = FILETYPES_DOCX_ALL,
                  initialdir: str | None = None) -> str | None:
        raise NotImplementedError

    def pick_folder(self, parent: Any, title: str, initialdir: str | None = None) -> str | None:
        raise NotImplementedError

    def save_as(self, parent: Any, title: str, defaultextension: str = ".docx",
                initialfile: str = "", filetypes: tuple = FILETYPES_DOCX_ALL,
                initialdir: str | None = None) -> str | None:
        raise NotImplementedError


class DialogFileSource(FileSource):
    """v1: только системные диалоги tkinter.filedialog (модальные поверх parent)."""

    @staticmethod
    def _top(parent: Any) -> Any:
        return parent.winfo_toplevel() if parent is not None else None

    def pick_file(self, parent, title, filetypes=FILETYPES_DOCX_ALL, initialdir=None):
        return filedialog.askopenfilename(
            parent=self._top(parent), title=title, filetypes=list(filetypes),
            initialdir=initialdir or os.getcwd()) or None

    def pick_folder(self, parent, title, initialdir=None):
        return filedialog.askdirectory(
            parent=self._top(parent), title=title, initialdir=initialdir or os.getcwd()) or None

    def save_as(self, parent, title, defaultextension=".docx", initialfile="",
                filetypes=FILETYPES_DOCX_ALL, initialdir=None):
        return filedialog.asksaveasfilename(
            parent=self._top(parent), title=title, defaultextension=defaultextension,
            initialfile=initialfile, filetypes=list(filetypes),
            initialdir=initialdir or os.getcwd()) or None


active_file_source: FileSource = DialogFileSource()


def pick_file(parent: Any, title: str, filetypes: tuple = FILETYPES_DOCX_ALL,
              initialdir: str | None = None) -> str | None:
    return active_file_source.pick_file(parent, title, filetypes, initialdir)


def pick_folder(parent: Any, title: str, initialdir: str | None = None) -> str | None:
    return active_file_source.pick_folder(parent, title, initialdir)


def choose_save_path(parent: Any, title: str, defaultextension: str = ".docx",
                     initialfile: str = "", filetypes: tuple = FILETYPES_DOCX_ALL,
                     initialdir: str | None = None) -> str | None:
    return active_file_source.save_as(parent, title, defaultextension, initialfile,
                                      filetypes, initialdir)


# ---------------------------------------------------------------------------
# Служебные модальные базы
# ---------------------------------------------------------------------------
class _ModalDialog(ctk.CTkToplevel):
    """Центрируемое модальное окно-заготовка (без блокировки главного цикла)."""

    def __init__(self, master: ctk.CTk | None, title: str, width: int, height: int):
        super().__init__(master)
        self.title(title)
        self.geometry(f"{width}x{height}")
        self.resizable(False, False)
        self.protocol("WM_DELETE_WINDOW", self.destroy)
        if master is not None:
            self.transient(master)
            x = master.winfo_rootx() + max((master.winfo_width() - width) // 2, 0)
            y = master.winfo_rooty() + max((master.winfo_height() - height) // 3, 0)
            self.geometry(f"+{x}+{y}")
        try:
            self.grab_set()
        except tk.TclError:
            pass
        self.bind("<Escape>", lambda _e: self.destroy())


def ask_text(parent: ctk.CTk | None, title: str, prompt: str,
             initial: str = "") -> str | None:
    """Однострочный ввод (новый ключ пресета, переименование профиля)."""

    class _Ask(_ModalDialog):
        def __init__(self):
            super().__init__(parent, title, DLG_WIDTH, 190)
            self.result: str | None = None
            lbl = ctk.CTkLabel(self, text=prompt, wraplength=DLG_WIDTH - 60)
            lbl.grid(row=0, column=0, padx=ROW_PAD, pady=(16, 4), sticky="w")
            self.entry = ctk.CTkEntry(self, width=DLG_WIDTH - 60)
            self.entry.grid(row=1, column=0, padx=ROW_PAD, pady=4)
            self.entry.insert(0, initial)
            self.entry.focus_set()
            btns = ctk.CTkFrame(self, fg_color="transparent")
            btns.grid(row=2, column=0, pady=12)
            ctk.CTkButton(btns, text=t("dlg_ok"), width=BTN_W // 2,
                          command=self._ok).grid(row=0, column=0, padx=4)
            ctk.CTkButton(btns, text=t("dlg_cancel"), width=BTN_W // 2,
                          fg_color="gray40", command=self.destroy).grid(row=0, column=1, padx=4)
            self.bind("<Return>", lambda _e: self._ok())

        def _ok(self) -> None:
            value = self.entry.get().strip()
            self.result = value or None
            self.destroy()

    dlg = _Ask()
    dlg.wait_window()
    return dlg.result


# ---------------------------------------------------------------------------
# Подтверждения и предупреждения (п. 10.1, 12)
# ---------------------------------------------------------------------------
def confirm(parent: Any, title: str, message: str) -> bool:
    return messagebox.askyesno(title, message, parent=parent)


def confirm_overwrite(parent: Any, path: str) -> bool:
    """Предупреждение перезаписи существующего файла (п. 10.1)."""
    return confirm(parent, t("app_title"), t("confirm_overwrite", path=path))


def warn_external_script(parent: Any) -> bool:
    """«Скрипт выполняется с правами пользователя» (п. 12)."""
    return confirm(parent, t("warn_external_title"), t("warn_external_text"))


def show_info(parent: Any, title: str, message: str) -> None:
    messagebox.showinfo(title, message, parent=parent)


# ---------------------------------------------------------------------------
# Диалог исключения с кнопкой «Скопировать traceback» (п. 12)
# ---------------------------------------------------------------------------
def exception_dialog(master: ctk.CTk | None, title_text: str, traceback_text: str) -> None:
    class _Exc(_ModalDialog):
        def __init__(self):
            super().__init__(master, t("dlg_exception_title"), EXC_WIDTH, EXC_HEIGHT)
            head = ctk.CTkLabel(self, text=f"{title_text}\n\n{t('exc_intro')}",
                                wraplength=EXC_WIDTH - 40, justify="left")
            head.pack(padx=12, pady=(12, 4), anchor="w")
            box = ctk.CTkTextbox(self, wrap="word")
            box.pack(fill="both", expand=True, padx=12, pady=4)
            box.insert("1.0", traceback_text)
            box.configure(state="disabled")
            bar = ctk.CTkFrame(self, fg_color="transparent")
            bar.pack(fill="x", padx=12, pady=10)
            ctk.CTkButton(bar, text=t("dlg_exc_copy"), width=BTN_W,
                          command=self._copy).pack(side="left")
            ctk.CTkButton(bar, text=t("dlg_close"), width=BTN_W // 2, fg_color="gray40",
                          command=self.destroy).pack(side="right")

        def _copy(self) -> None:
            toplevel = self.winfo_toplevel()
            toplevel.clipboard_clear()
            toplevel.clipboard_append(traceback_text)

    _Exc()


# ---------------------------------------------------------------------------
# «О программе» (п. 10.5)
# ---------------------------------------------------------------------------
def about_dialog(master: ctk.CTk | None, datadir: str, portable: bool) -> None:
    try:  # версия CustomTkinter берётся из пакета, не хардкодится
        ctk_version = getattr(ctk, "__version__", "?")
    except Exception:  # pragma: no cover
        ctk_version = "?"
    text = t(
        "about_text",
        title=t("app_title"),
        version=C.APP_VERSION,
        py=f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}",
        ctk=ctk_version,
        datadir=datadir,
        portable="включён" if portable else "выключен",
    )

    class _About(_ModalDialog):
        def __init__(self):
            super().__init__(master, t("dlg_about_title"), ABOUT_WIDTH, ABOUT_HEIGHT)
            box = ctk.CTkTextbox(self, wrap="word")
            box.pack(fill="both", expand=True, padx=12, pady=12)
            box.insert("1.0", text)
            box.configure(state="disabled")
            ctk.CTkButton(self, text=t("dlg_close"), width=BTN_W // 2,
                          command=self.destroy).pack(pady=(0, 12))

    _About()


# ---------------------------------------------------------------------------
# Приветствие первого запуска: 3 шага (п. 12)
# ---------------------------------------------------------------------------
def welcome_dialog(master: ctk.CTk | None, on_done: Callable[[bool], None]) -> None:
    """on_done(dont_show_again) — вызывается при закрытии."""

    class _Welcome(_ModalDialog):
        def __init__(self):
            super().__init__(master, t("dlg_welcome_title"), WELCOME_WIDTH, WELCOME_HEIGHT)
            wrap = WELCOME_WIDTH - 60
            title = ctk.CTkLabel(self, text=t("dlg_welcome_title"),
                                 font=ctk.CTkFont(size=WELCOME_TITLE_SIZE, weight="bold"))
            title.pack(pady=(16, 8))
            for key in ("welcome_step1", "welcome_step2", "welcome_step3"):
                ctk.CTkLabel(self, text=t(key), wraplength=wrap, justify="left").pack(
                    padx=24, pady=4, anchor="w")
            self.dont_show = ctk.CTkCheckBox(self, text=t("welcome_dont_show"))
            self.dont_show.pack(pady=12)
            ctk.CTkButton(self, text=t("dlg_ok"), width=BTN_W,
                          command=self._close).pack(pady=(0, 16))
            self.bind("<Destroy>", self._on_destroy, add="+")
            self._fired = False

        def _close(self) -> None:
            self.destroy()

        def _on_destroy(self, _event: tk.Event) -> None:
            if not self._fired:
                self._fired = True
                on_done(bool(self.dont_show.get_variable().get()
                             if hasattr(self.dont_show, "get_variable")
                             else self.dont_show.get()))

    _Welcome()


# ---------------------------------------------------------------------------
# Диалог «Настройки» (п. 10.5)
# ---------------------------------------------------------------------------
_INTERP_INTERNAL = t("set_interp_internal")
_INTERP_EXTERNAL = t("set_interp_external")


def settings_dialog(master: ctk.CTk, state: dict, dirs: Any,
                    on_apply: Callable[[dict], None]) -> None:
    """Правит копию state; on_apply(new_state) — когда пользователь жмёт «ОК».

    «Сбросить настройки» — только с подтверждением; portable-флаг меняется
    с предупреждением о перезапуске (файл-маркер создаёт/удаляет core).
    """

    class _Settings(_ModalDialog):
        def __init__(self):
            super().__init__(master, t("dlg_settings_title"), SETTINGS_WIDTH, SETTINGS_HEIGHT)
            self.vars: dict[str, Any] = {}
            body = ctk.CTkScrollableFrame(self)
            body.pack(fill="both", expand=True, padx=8, pady=4)
            row = 0

            def add_dir(label_key: str, skey: str, current: str) -> None:
                nonlocal row
                label = t(label_key)
                ctk.CTkLabel(body, text=label, anchor="w").grid(
                    row=row, column=0, sticky="w", padx=6, pady=ROW_PAD)
                entry = ctk.CTkEntry(body)
                entry.insert(0, current)
                entry.grid(row=row, column=1, sticky="ew", padx=6, pady=ROW_PAD)
                ctk.CTkButton(body, text=t("btn_browse"), width=BTN_W // 2,
                              command=lambda e=entry, title=label: self._browse(e, title)).grid(
                    row=row, column=2, padx=6)
                self.vars[skey] = entry
                row += 1

            def add_choice(label_key: str, skey: str, values: tuple | list,
                           current: str | int) -> None:
                nonlocal row
                ctk.CTkLabel(body, text=t(label_key), anchor="w").grid(
                    row=row, column=0, sticky="w", padx=6, pady=ROW_PAD)
                menu = ctk.CTkOptionMenu(body, values=[str(v) for v in values])
                menu.set(str(current))
                menu.grid(row=row, column=1, sticky="w", padx=6, pady=ROW_PAD)
                self.vars[skey] = menu
                row += 1

            def add_entry(label_key: str, skey: str, current: Any) -> None:
                nonlocal row
                ctk.CTkLabel(body, text=t(label_key), anchor="w").grid(
                    row=row, column=0, sticky="w", padx=6, pady=ROW_PAD)
                entry = ctk.CTkEntry(body)
                entry.insert(0, str(current))
                entry.grid(row=row, column=1, sticky="ew", padx=6, pady=ROW_PAD)
                self.vars[skey] = entry
                row += 1

            body.grid_columnconfigure(1, weight=1)

            add_dir("set_scripts_dir", "scripts_dir", str(state.get("scripts_dir") or dirs.scripts))
            add_dir("set_output_dir", "output_dir", str(state.get("output_dir") or dirs.output))
            mode_key = "external" if state.get("interpreter_mode") == "external" else "internal"
            add_choice("set_interp_mode", "interpreter_mode",
                       (_INTERP_INTERNAL, _INTERP_EXTERNAL),
                       _INTERP_INTERNAL if mode_key == "internal" else _INTERP_EXTERNAL)
            add_entry("set_interp_path", "interpreter_path",
                      state.get("interpreter_path") or sys.executable)
            add_entry("set_timeout", "script_timeout_s",
                      state.get("script_timeout_s") or C.DEFAULT_SCRIPT_TIMEOUT_S)
            add_choice("set_scale", "ui_scale", C.UI_SCALES, state.get("ui_scale") or C.DEFAULT_UI_SCALE)
            add_choice("lbl_theme", "theme", C.UI_THEMES, state.get("theme") or C.DEFAULT_THEME)

            self.portable_var = tk.BooleanVar(
                master=self, value=bool(state.get("portable") or dirs.is_portable()))
            ctk.CTkSwitch(body, text=t("set_portable"), variable=self.portable_var).grid(
                row=row, column=0, columnspan=3, sticky="w", padx=6, pady=ROW_PAD)
            row += 1

            ctk.CTkButton(body, text=t("dlg_open_datadir"), width=BTN_W,
                          command=lambda: self._open_datadir(dirs)).grid(
                row=row, column=0, sticky="w", padx=6, pady=ROW_PAD)
            ctk.CTkButton(body, text=t("dlg_reset"), width=BTN_W, fg_color="#8a3b3b",
                          command=self._reset).grid(row=row, column=1, sticky="w", padx=6)
            row += 1

            bar = ctk.CTkFrame(self, fg_color="transparent")
            bar.pack(fill="x", padx=12, pady=10)
            ctk.CTkButton(bar, text=t("dlg_ok"), width=BTN_W, command=self._apply).pack(side="right")
            ctk.CTkButton(bar, text=t("dlg_cancel"), width=BTN_W, fg_color="gray40",
                          command=self.destroy).pack(side="right", padx=6)

        def _open_datadir(self, dirs: Any) -> None:
            err = open_path(str(dirs.home))
            if err:
                show_info(self, t("dlg_settings_title"), t("open_error", path=str(dirs.home),
                                                          why=err))

        def _browse(self, entry: ctk.CTkEntry, title: str) -> None:
            path = pick_folder(self, title)
            if path:
                entry.delete(0, tk.END)
                entry.insert(0, path)

        def _reset(self) -> None:
            if not confirm(self, t("dlg_settings_title"), t("dlg_reset_confirm")):
                return
            try:
                from app.core import config  # лениво: только при сбросе
                fresh = config.default_state(dirs)
            except Exception:
                fresh = {}
            self._fill(fresh)
            on_apply(fresh)
            self.destroy()

        def _fill(self, values: dict) -> None:
            for key, widget in self.vars.items():
                if key in values:
                    try:
                        widget.delete(0, tk.END)
                    except tk.TclError:
                        pass
                    if isinstance(widget, ctk.CTkOptionMenu):
                        current = values[key]
                        if key == "interpreter_mode":
                            current = _INTERP_EXTERNAL if current == "external" else _INTERP_INTERNAL
                        widget.set(str(current))
                    else:
                        widget.insert(0, str(values[key]))

        def _apply(self) -> None:
            new_state = dict(state)
            try:
                for key, widget in self.vars.items():
                    if isinstance(widget, ctk.CTkOptionMenu):
                        raw = widget.get()
                        if key == "interpreter_mode":
                            new_state["interpreter_mode"] = (
                                "external" if raw == _INTERP_EXTERNAL else "internal")
                        elif key == "ui_scale":
                            new_state["ui_scale"] = int(raw)
                        else:
                            new_state[key] = raw
                    else:
                        raw = widget.get().strip()
                        if key == "script_timeout_s":
                            new_state["script_timeout_s"] = int(float(raw or C.DEFAULT_SCRIPT_TIMEOUT_S))
                        else:
                            new_state[key] = raw
                new_state["portable"] = bool(self.portable_var.get())
            except ValueError:
                show_info(self, t("dlg_settings_title"), t("gen_err_params", details=t("set_timeout")))
                return
            on_apply(new_state)
            self.destroy()

    _Settings()


# ---------------------------------------------------------------------------
# Открытие файла/папки системным просмотрщиком (без shell=True, п. 12)
# ---------------------------------------------------------------------------
def open_path(path: str | os.PathLike[str]) -> str | None:
    """Открыть файл/каталог родным способом ОС. Возвращает текст ошибки или None.

    Запуск проверяется: `Popen` без кода возврата означает только «процесс стартовал»,
    а пользователь при этом видит ничего — просмотрщик мог отказать (нет ассоциации,
    права, DISPLAY). Поэтому ждём завершения с коротким лимитом: молча долгий процесс —
    успех, быстрый ненулевой код — ошибка с текстом из stderr.
    """
    target = str(path)
    if not os.path.exists(target):
        return t("open_missing", path=target)
    try:
        system = platform.system()
        if system == "Windows":
            os.startfile(target)  # type: ignore[attr-defined]  # есть только на Windows
            return None
        cmd = "open" if system == "Darwin" else "xdg-open"
        binary = shutil.which(cmd)
        if binary is None:
            return t("open_no_handler", cmd=cmd)
        proc = subprocess.Popen([binary, target], stdin=subprocess.DEVNULL,
                                stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
                                text=True, encoding="utf-8", errors="replace")
        try:
            try:
                rc = proc.wait(timeout=OPEN_VERIFY_TIMEOUT_S)
            except subprocess.TimeoutExpired:
                return None  # процесс живёт сам: окно просмотрщика открывается
            if rc == 0:
                return None
            try:
                why = (proc.stderr.read() if proc.stderr else "").strip()
            except Exception:
                why = ""
            return t("open_failed", code=rc, why=why or t("open_failed_plain"))
        finally:
            # Ручку конвейера надо закрыть обязательно: на Windows открытый fd блокирует
            # удаление каталога, а в тесту ResourceWarning — виден как утечка.
            for pipe in (proc.stdout, proc.stderr):
                if pipe is not None:
                    try:
                        pipe.close()
                    except OSError:
                        pass
    except Exception as exc:  # нет ассоциаций/прав — понятное сообщение, не падение
        return str(exc)


def portable_flag_path(dirs: Any) -> str:
    """Путь маркера portable рядом с исполняемым/репозиторием (п. 19.6)."""
    base = getattr(sys, "_MEIPASS", None)
    anchor = os.path.dirname(sys.executable) if getattr(sys, "frozen", False) else (
        str(dirs.project_dir) if base is None else base)
    return os.path.join(anchor, C.PORTABLE_FLAG)
