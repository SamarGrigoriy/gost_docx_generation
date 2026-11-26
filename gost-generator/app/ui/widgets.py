"""Переиспользуемые виджеты вкладки: поля параметров, аккордеоны, таблицы замечаний.

Правило слоя: виджеты только отображают/собирают значения; любые пересчёты и
проверки — в `app.core`. Строки — только через `app.strings.t`.
"""

from __future__ import annotations

import dataclasses
import datetime as dt
import tkinter as tk
from typing import Any, Callable

import customtkinter as ctk

from app import constants as C
from app.strings import T as T_KEYS, t
from app.ui import dialogs

# Метрики оформления (единственное место «чисел UI» слоя виджетов).
FIELD_LABEL_W = 180
FIELD_ENTRY_W = 320
TEXTAREA_H = 110
DELETE_BTN_W = 40
BADGE_PADX = 6
BADGE_PADY = 2
BADGE_RADIUS = 8
ACCENT_WIDTH = 4
ISSUE_COL_W = {"sev": 92, "code": 118, "msg": 380, "where": 150, "how": 300}
ISSUE_ROWS_H = 180  # минимальная высота таблицы: пустая таблица не должна схлопываться в 0
SECTION_HEADER_TEXT = ("gray15", "gray92")  # на прозрачном фоне CTk сам не подберёт цвет
ROW_SPACING = 3
SECTION_RADIUS = 8
ROW_RADIUS = 6
SECTION_BORDER = 1
KV_KEY_W = 160
KV_VAL_W = 260
PICKER_ENTRY_W = 340
PICKER_BTN_W = 96
ERROR_TEXT_COLOR = "#e57373"
PLACEHOLDER_COLOR = "gray55"
DATE_PLACEHOLDER = "ГГГГ-ММ-ДД"

# Цвета бейжей severity/состояний; одинаковы во всём UI.
SEVERITY_COLORS = {
    "error": "#c0392b",
    "warning": "#c88a2b",
    "info": "#3d6fa8",
    "ok": "#2e7d4f",
    "muted": "#5b5b66",
}
SEVERITY_TEXT_COLORS = {
    "error": "#ffffff",
    "warning": "#141414",
    "info": "#ffffff",
    "ok": "#ffffff",
    "muted": "#e8e8e8",
}


TOOLTIP_DELAY_MS = 400
TOOLTIP_OFFSET_PX = 6
TOOLTIP_BG = "#ffffe1"


def attach_tooltip(widget: Any, text: str) -> None:
    """Появляющаяся подсказка для виджета.

    В customtkinter 6.0 аргумент `tooltip_text` у базовых виджетов удалён, а Tk
    не всплывает события от дочернего canvas'а, поэтому вешаем обработчики и на
    виджет, и на его детей.
    """
    if not text:
        return
    state: dict[str, Any] = {"job": None, "win": None}

    def hide(_event: Any = None) -> None:
        if state["job"] is not None:
            widget.after_cancel(state["job"])
            state["job"] = None
        if state["win"] is not None:
            try:
                state["win"].destroy()
            except tk.TclError:
                pass
            state["win"] = None

    def show() -> None:
        state["job"] = None
        if state["win"] is not None or not widget.winfo_exists():
            return
        try:
            win = tk.Toplevel(widget)
            win.overrideredirect(True)
            win.attributes("-topmost", True)
            tk.Label(win, text=text, bg=TOOLTIP_BG, relief="solid", bd=1,
                     padx=6, pady=2, justify="left").pack()
            win.update_idletasks()
            x = widget.winfo_rootx() + TOOLTIP_OFFSET_PX
            y = widget.winfo_rooty() + widget.winfo_height() + TOOLTIP_OFFSET_PX
            win.geometry(f"+{x}+{y}")
            state["win"] = win
        except tk.TclError:
            state["win"] = None

    def on_enter(_event: Any = None) -> None:
        if state["job"] is None and widget.winfo_exists():
            state["job"] = widget.after(TOOLTIP_DELAY_MS, show)

    for target in (widget, *widget.winfo_children()):
        target.bind("<Enter>", on_enter, add="+")
        target.bind("<Leave>", hide, add="+")
        target.bind("<Destroy>", hide, add="+")


def with_tip(widget: Any, text: str | None) -> Any:
    """Присоединить подсказку и вернуть виджет — чтобы не разбивать цепочку конструктора."""
    attach_tooltip(widget, str(text or ""))
    return widget


def issue_to_dict(issue: Any) -> dict:
    """Issue (dataclass) или dict → dict с ключами code/severity/message/location/hint."""
    if dataclasses.is_dataclass(issue) and not isinstance(issue, type):
        row = dataclasses.asdict(issue)
    elif isinstance(issue, dict):
        row = dict(issue)
    else:  # внешний объект с атрибутами
        row = {f: getattr(issue, f, None) for f in
               ("code", "severity", "message", "location", "hint")}
    row.setdefault("code", "")
    row.setdefault("severity", "info")
    row.setdefault("message", "")
    row.setdefault("location", {})
    row.setdefault("hint", "")
    return row


def format_location(location: dict) -> str:
    """Человекочитаемое «Где»: абзац/страница-оценка (номера — эвристика inspector)."""
    parts: list[str] = []
    para = location.get("paragraph") if isinstance(location, dict) else None
    page = location.get("page") if isinstance(location, dict) else None
    if para is not None:
        parts.append(t("loc_paragraph", n=para))
    if page is not None:
        parts.append(t("loc_page", n=page))
    return ", ".join(parts) if parts else t("loc_unknown")


# ---------------------------------------------------------------------------
# Badge
# ---------------------------------------------------------------------------
class Badge(ctk.CTkLabel):
    """Цветной бейдж: severity замечания, статус, пометки скрипта."""

    def __init__(self, master: Any, text: str, kind: str = "muted", **kwargs: Any):
        bg = SEVERITY_COLORS.get(kind, SEVERITY_COLORS["muted"])
        fg = SEVERITY_TEXT_COLORS.get(kind, "#ffffff")
        tooltip = str(kwargs.pop("tooltip_text", "") or "")
        super().__init__(
            master, text=text, corner_radius=BADGE_RADIUS, fg_color=bg, text_color=fg,
            padx=BADGE_PADX, pady=BADGE_PADY, **kwargs)
        attach_tooltip(self, tooltip)
        self._kind = kind

    @classmethod
    def for_severity(cls, master: Any, severity: str) -> "Badge":
        kind = severity if severity in SEVERITY_COLORS else "muted"
        label_key = f"severity_{severity}"
        label = t(label_key) if label_key in T_KEYS else severity
        return cls(master, label, kind=kind)

    def set_kind(self, kind: str) -> None:
        self._kind = kind
        self.configure(fg_color=SEVERITY_COLORS.get(kind, SEVERITY_COLORS["muted"]),
                       text_color=SEVERITY_TEXT_COLORS.get(kind, "#ffffff"))

    def set_text(self, text: str, kind: str | None = None) -> None:
        self.configure(text=text)
        if kind:
            self.set_kind(kind)


# ---------------------------------------------------------------------------
# SectionAccordion
# ---------------------------------------------------------------------------
class SectionAccordion(ctk.CTkFrame):
    """Раскрывающаяся секция: заголовок-кнопка с акцентной полосой + тело.

    Тело — обычный CTkFrame; вкладки наполняют его через `.body` (grid/pack).
    """

    def __init__(self, master: Any, title: str, tooltip: str | None = None,
                 expanded: bool = False, on_toggle: Callable[[bool], None] | None = None):
        super().__init__(master, corner_radius=SECTION_RADIUS, border_width=SECTION_BORDER)
        self._title = title
        self._expanded = expanded
        self._on_toggle = on_toggle
        header_style = {"fg_color": "transparent", "hover_color": "gray25",
                        "text_color": SECTION_HEADER_TEXT}
        self.header = with_tip(ctk.CTkButton(
            self, text="", anchor="w", font=ctk.CTkFont(weight="bold"),
            command=self.toggle, **header_style), tooltip)
        self.header.pack(fill="x")
        self.body = ctk.CTkFrame(self, fg_color="transparent")
        self._update()

    def set_expanded(self, value: bool) -> None:
        if value == self._expanded:
            return
        self._expanded = value
        self._update()
        if self._on_toggle:
            self._on_toggle(value)

    def toggle(self) -> None:
        self.set_expanded(not self._expanded)

    def set_title(self, title: str) -> None:
        self._title = title
        self._update()

    def _update(self) -> None:
        arrow = "▾" if self._expanded else "▸"
        self.header.configure(text=f"{arrow}  {self._title}")
        if self._expanded:
            self.body.pack(fill="both", expand=True, padx=10, pady=(0, 8))
        else:
            self.body.pack_forget()


# ---------------------------------------------------------------------------
# FilePickerRow
# ---------------------------------------------------------------------------
class FilePickerRow(ctk.CTkFrame):
    """Поле пути + «Обзор…»; источник пути — абстракция dialogs.FileSource (v1 — диалог)."""

    def __init__(self, master: Any, label: str = "", kind: str = "file",
                 initial: str = "", filetypes: tuple | None = None,
                 title: str | None = None, on_change: Callable[[], None] | None = None,
                 tooltip: str | None = None):
        super().__init__(master, fg_color="transparent")
        self.kind = kind
        self.filetypes = filetypes or (dialogs.FILETYPES_DOCX_ALL if kind == "file" else None)
        self.title = title or label
        self.on_change = on_change
        if label:
            ctk.CTkLabel(self, text=label, width=FIELD_LABEL_W, anchor="w").pack(side="left")
        self.entry = with_tip(ctk.CTkEntry(self, width=PICKER_ENTRY_W), tooltip)
        self.entry.insert(0, initial)
        self.entry.pack(side="left", fill="x", expand=True)
        ctk.CTkButton(self, text=t("btn_browse"), width=PICKER_BTN_W,
                      command=self._browse).pack(side="left", padx=(6, 0))
        self.entry.bind("<KeyRelease>", self._fire, add="+")

    def _browse(self) -> None:
        parent = self.winfo_toplevel()
        if self.kind == "folder":
            path = dialogs.pick_folder(parent, self.title)
        elif self.kind == "file_save":
            path = dialogs.choose_save_path(parent, self.title,
                                            filetypes=self.filetypes or dialogs.FILETYPES_DOCX_ALL)
        else:
            path = dialogs.pick_file(parent, self.title,
                                     filetypes=self.filetypes or dialogs.FILETYPES_ALL)
        if path:
            self.set(path)

    def _fire(self, _event: Any = None) -> None:
        if self.on_change:
            self.on_change()

    def get(self) -> str:
        return self.entry.get().strip()

    def set(self, value: str) -> None:
        self.entry.delete(0, tk.END)
        self.entry.insert(0, value)
        self._fire()

    def set_state(self, enabled: bool) -> None:
        self.entry.configure(state="normal" if enabled else "disabled")


# ---------------------------------------------------------------------------
# ParamField
# ---------------------------------------------------------------------------
class ParamField(ctk.CTkFrame):
    """Поле параметра из SCRIPT_META.params: spec = {id,label,type,default,options}.

    Типы — constants.PARAM_TYPES. Инлайн-валидация: `validate()` возвращает
    текст ошибки или None; ошибка подсвечивается подписью под полем.
    """

    def __init__(self, master: Any, spec: dict, on_change: Callable[[], None] | None = None):
        super().__init__(master, fg_color="transparent")
        self.spec = spec
        self.param_id: str = str(spec.get("id") or spec.get("label") or "param")
        self.type: str = spec.get("type") or "text"
        self._on_change = on_change
        self._widgets: list[Any] = []

        ctk.CTkLabel(self, text=str(spec.get("label") or self.param_id),
                     width=FIELD_LABEL_W, anchor="w").grid(row=0, column=0, sticky="nw",
                                                           padx=(0, 6), pady=2)
        editor_holder = ctk.CTkFrame(self, fg_color="transparent")
        editor_holder.grid(row=0, column=1, sticky="ew")
        self._editor = self._build_editor(editor_holder)
        # Построенный редактор надо ещё разместить, иначе поле остаётся пустой подписью.
        self._editor.pack(fill="x", expand=True)
        self.error_lbl = ctk.CTkLabel(self, text="", anchor="w",
                                      text_color=ERROR_TEXT_COLOR)
        self.error_lbl.grid(row=1, column=1, sticky="w")
        self.grid_columnconfigure(1, weight=1)
        self.set(spec.get("default"))

    # -- построение редактора ------------------------------------------------
    def _build_editor(self, holder: ctk.CTkFrame) -> Any:
        tip = str(self.spec.get("tooltip") or "")
        ptype = self.type
        if ptype == "textarea":
            box = with_tip(ctk.CTkTextbox(holder, height=TEXTAREA_H, width=FIELD_ENTRY_W), tip)
            box.bind("<KeyRelease>", self._change, add="+")
            self._widgets.append(box)
            return box
        if ptype == "bool":
            var = tk.BooleanVar(master=holder, value=False)
            sw = with_tip(ctk.CTkSwitch(holder, text="", variable=var, command=self._change), tip)
            self._widgets.append(sw)
            return sw
        if ptype == "select":
            options = [str(o) for o in (self.spec.get("options") or [])] or [""]
            menu = with_tip(ctk.CTkOptionMenu(holder, values=options, width=FIELD_ENTRY_W,
                                              command=self._change), tip)
            self._widgets.append(menu)
            return menu
        if ptype in ("file", "folder"):
            row = FilePickerRow(holder, kind=ptype, title=str(self.spec.get("label") or ""),
                                on_change=self._change, tooltip=tip,
                                filetypes=dialogs.FILETYPES_ALL)
            self._widgets.append(row)
            return row
        entry = with_tip(ctk.CTkEntry(holder, width=FIELD_ENTRY_W), tip)
        if ptype == "date":
            entry.configure(placeholder_text=DATE_PLACEHOLDER)
        entry.bind("<KeyRelease>", self._change, add="+")
        self._widgets.append(entry)
        return entry

    # -- значения -------------------------------------------------------------
    def get(self) -> Any:
        ed = self._editor
        if self.type == "bool":
            return bool(ed.get_variable().get()) if hasattr(ed, "get_variable") else bool(ed.get())
        if self.type == "textarea":
            return ed.get("1.0", "end").strip()
        if self.type in ("file", "folder"):
            return ed.get()
        if self.type == "select":
            return ed.get()
        return ed.get().strip()

    def set(self, value: Any) -> None:
        ed = self._editor
        if self.type == "bool":
            truthy = bool(value) if value is not None else False
            if hasattr(ed, "select"):
                (ed.select if truthy else ed.deselect)()
            return
        text = "" if value is None else str(value)
        if self.type == "textarea":
            ed.delete("1.0", "end")
            if text:
                ed.insert("1.0", text)
            return
        if self.type in ("file", "folder"):
            ed.set(text)
            return
        if self.type == "select":
            ed.set(text or ed.cget("values")[0])
            return
        ed.delete(0, tk.END)
        ed.insert(0, text)

    # -- валидация --------------------------------------------------------------
    def validate(self) -> str | None:
        """Инлайн-валидация значения; None — значение корректно."""
        value = self.get()
        required = bool(self.spec.get("required"))
        ptype = self.type if self.type in C.PARAM_TYPES else "text"
        if required and not str(value):
            return t("gen_err_required", label=str(self.spec.get("label") or self.param_id))
        if str(value) == "" and not required:
            self._show_error(None)
            return None
        try:
            value_s = str(value)
            if ptype == "int":
                int(value_s.replace(" ", ""))
            elif ptype == "float":
                float(value_s.replace(",", ".").replace(" ", ""))
            elif ptype == "date":
                dt.date.fromisoformat(value_s) if "-" in value_s else \
                    dt.datetime.strptime(value_s, "%d.%m.%Y").date()
            elif ptype in ("file", "folder") and value_s:
                import os
                if not os.path.exists(value_s):
                    return t("gen_err_input_missing")
        except (ValueError, TypeError):
            return t("gen_err_params", details=self.spec.get("label") or self.param_id)
        self._show_error(None)
        return None

    def _show_error(self, message: str | None) -> None:
        self.error_lbl.configure(text=message or "")

    def set_enabled(self, enabled: bool) -> None:
        state = "normal" if enabled else "disabled"
        for w in self._widgets:
            try:
                w.configure(state=state)
            except tk.TclError:
                pass

    def _change(self, _event: Any = None) -> None:
        self.validate()
        if self._on_change:
            self._on_change()


# ---------------------------------------------------------------------------
# KVEditor — редактируемая таблица ключ-значение (profile.extra, п. 10.4)
# ---------------------------------------------------------------------------
class KVEditor(ctk.CTkScrollableFrame):
    def __init__(self, master: Any, on_change: Callable[[], None] | None = None,
                 height: int = 0):
        kwargs: dict[str, Any] = {"label_text": ""}
        if height:
            kwargs["height"] = height
        super().__init__(master, **kwargs)
        self.on_change = on_change
        self._rows: list[dict[str, Any]] = []
        self._placeholder = ctk.CTkLabel(self, text=t("kv_empty"))
        self._placeholder.grid(row=0, column=0, pady=8)
        self._row_counter = 1  # строка 0 зарезервирована под плейсхолдер
        self.add_btn = ctk.CTkButton(self, text=t("prof_add_field"), width=KV_KEY_W // 4,
                                     fg_color="gray40", command=self.add_empty)
        self.add_btn.grid(row=self._row_counter, column=0, sticky="w", pady=4)

    def _reposition_add_btn(self) -> None:
        self.add_btn.grid(row=self._row_counter, column=0, sticky="w", pady=4)

    def set(self, data: dict) -> None:
        for row in self._rows:
            row["frame"].destroy()
        self._rows.clear()
        self._row_counter = 1
        self._placeholder.grid_forget()
        for key, value in (data or {}).items():
            self._add_row(str(key), "" if value is None else str(value))
        if not data:
            self._add_row("", "")
        self._reposition_add_btn()

    def get(self) -> dict:
        out: dict[str, str] = {}
        for row in self._rows:
            key = row["key"].get().strip()
            if key:
                out[key] = row["val"].get().strip()
        return out

    def add_empty(self) -> None:
        self._add_row("", "")
        self._placeholder.grid_forget()
        self._reposition_add_btn()

    def _add_row(self, key: str, value: str) -> None:
        frame = ctk.CTkFrame(self, fg_color="transparent")
        frame.grid(row=self._row_counter, column=0, sticky="ew", pady=ROW_SPACING)
        self.grid_columnconfigure(0, weight=1)
        k_entry = ctk.CTkEntry(frame, width=KV_KEY_W, placeholder_text=t("prof_key"))
        k_entry.insert(0, key)
        k_entry.grid(row=0, column=0, padx=(0, 4))
        v_entry = ctk.CTkEntry(frame, width=KV_VAL_W, placeholder_text=t("prof_value"))
        v_entry.insert(0, value)
        v_entry.grid(row=0, column=1, padx=(0, 4))
        row_ref = {"frame": frame, "key": k_entry, "val": v_entry}
        btn = ctk.CTkButton(frame, text=t("prof_del_field"), width=DELETE_BTN_W,
                            fg_color="gray40",
                            command=lambda r=row_ref: self._remove_row(r))
        btn.grid(row=0, column=2)
        for e in (k_entry, v_entry):
            e.bind("<KeyRelease>", lambda _ev: self._fire(), add="+")
        self._rows.append(row_ref)
        self._row_counter += 1

    def _remove_row(self, row_ref: dict) -> None:
        row_ref["frame"].destroy()
        if row_ref in self._rows:
            self._rows.remove(row_ref)
        if not self._rows:
            self._placeholder.grid()
        self._reposition_add_btn()
        self._fire()

    def _fire(self) -> None:
        if self.on_change:
            self.on_change()


# ---------------------------------------------------------------------------
# IssueTable — таблица замечаний inspector (п. 10.3)
# ---------------------------------------------------------------------------
class IssueTable(ctk.CTkFrame):
    """Прокручиваемая таблица Severity|Код|Замечание|Где|Как исправить.

    Фильтры по severity, поиск подстроки, двойной клик → on_detail(issue_dict).
    Реальная логика проверки — в core; здесь только отображение.
    """

    def __init__(self, master: Any, on_detail: Callable[[dict], None] | None = None):
        super().__init__(master)
        self.on_detail = on_detail
        self._all: list[dict] = []

        toolbar = ctk.CTkFrame(self, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=4, pady=(4, 0))
        self._filters: dict[str, ctk.CTkCheckBox] = {}
        for sev in C.SEVERITIES:
            var = tk.BooleanVar(master=self, value=True)
            cb = ctk.CTkCheckBox(toolbar, text=t(f"flt_{sev}"), variable=var,
                                 command=self._rebuild)
            cb.pack(side="left", padx=6)
            self._filters[sev] = cb
        self.search = ctk.CTkEntry(toolbar, width=FIELD_ENTRY_W,
                                   placeholder_text=t("flt_search"))
        self.search.pack(side="right", padx=6)
        self.search.bind("<KeyRelease>", lambda _e: self._rebuild(), add="+")

        header = ctk.CTkFrame(self, fg_color="transparent")
        header.grid(row=1, column=0, sticky="ew")
        for col, (key, w) in enumerate(
                (("col_severity", ISSUE_COL_W["sev"]), ("col_code", ISSUE_COL_W["code"]),
                 ("col_message", ISSUE_COL_W["msg"]), ("col_where", ISSUE_COL_W["where"]),
                 ("col_how", ISSUE_COL_W["how"]))):
            ctk.CTkLabel(header, text=t(key), width=w, anchor="w",
                         font=ctk.CTkFont(weight="bold")).grid(row=0, column=col, padx=6)

        self.rows_frame = ctk.CTkScrollableFrame(self, label_text="", height=ISSUE_ROWS_H)
        self.rows_frame.grid(row=2, column=0, sticky="nsew", padx=2, pady=4)
        self.grid_rowconfigure(2, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._row_widgets: list[ctk.CTkFrame] = []

    # -- публичный API ---------------------------------------------------------
    def set_issues(self, issues: list[Any]) -> None:
        self._all = [issue_to_dict(i) for i in issues]
        self._all.sort(key=lambda i: C.SEVERITY_ORDER.get(i.get("severity", "info"), 9))
        self._rebuild()

    def clear(self) -> None:
        self._all = []
        self._rebuild()

    def visible_count(self) -> int:
        return len(self._all_rows_filtered())

    # -- внутреннее -------------------------------------------------------------
    def _all_rows_filtered(self) -> list[dict]:
        wanted = {sev for sev, cb in self._filters.items() if self._cb_on(cb)}
        needle = self.search.get().strip().casefold()
        rows = []
        for issue in self._all:
            if issue.get("severity") not in wanted:
                continue
            if needle:
                haystack = " ".join(str(issue.get(k, "")) for k in
                                    ("message", "code", "hint")).casefold()
                if needle not in haystack:
                    continue
            rows.append(issue)
        return rows

    @staticmethod
    def _cb_on(cb: ctk.CTkCheckBox) -> bool:
        var = cb.get_variable() if hasattr(cb, "get_variable") else None
        return bool(var.get()) if var is not None else bool(cb.get())

    def _rebuild(self) -> None:
        for w in self._row_widgets:
            w.destroy()
        self._row_widgets.clear()
        for r, issue in enumerate(self._all_rows_filtered()):
            frame = ctk.CTkFrame(self.rows_frame, corner_radius=ROW_RADIUS)
            frame.grid(row=r, column=0, sticky="ew", pady=ROW_SPACING)
            sev = str(issue.get("severity", "info"))
            badge = Badge(frame, sev, kind=sev if sev in SEVERITY_COLORS else "muted",
                          width=ISSUE_COL_W["sev"] - BADGE_PADX * 2)
            badge.grid(row=0, column=0, padx=6, pady=4)
            cells = (
                (str(issue.get("code", "")), ISSUE_COL_W["code"]),
                (str(issue.get("message", "")), ISSUE_COL_W["msg"]),
                (format_location(issue.get("location") or {}), ISSUE_COL_W["where"]),
                (str(issue.get("hint", "")), ISSUE_COL_W["how"]),
            )
            for col, (text_, width_) in enumerate(cells, start=1):
                ctk.CTkLabel(frame, text=text_, width=width_, anchor="w",
                             justify="left").grid(row=0, column=col, padx=6, pady=4)
            frame.bind("<Double-Button-1>", lambda _e, i=issue: self._open_detail(i))
            for child in frame.winfo_children():
                child.bind("<Double-Button-1>", lambda _e, i=issue: self._open_detail(i))
            self._row_widgets.append(frame)
        self.rows_frame.grid_columnconfigure(0, weight=1)

    def _open_detail(self, issue: dict) -> None:
        if self.on_detail:
            self.on_detail(issue)
