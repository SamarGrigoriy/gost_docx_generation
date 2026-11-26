"""Вкладка «Профиль» (п. 10.4).

CRUD профилей через `app.core.profile_store`, предпросмотр шапки —
`title_header_preview`, плоская подстановка — `flat_meta`. Год по умолчанию
берётся из системной даты (через core/`datetime`), не хардкодится.
"""

from __future__ import annotations

import datetime as dt
import json
import tkinter as tk
from typing import Any

import customtkinter as ctk

from app import constants as C
from app.core import profile_store, rules_store
from app.strings import t
from app.ui import dialogs
from app.ui.widgets import Badge, KVEditor, SectionAccordion

# Метрики вкладки.
LIST_W = 240
LABEL_W = 170
EDIT_W = 320
BTN_W = 120
SAVE_BTN_W = 150
ACTIVE_DOT_W = 16
ROW_H = 30
PREVIEW_H = 200

_CUSTOM_WORK_TYPE = "Свой вариант"


class ProfileTab(ctk.CTkFrame):
    def __init__(self, master: Any, app: Any):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self._key: str = ""
        self._base: dict = {}
        self._dirty = False
        self._fields: dict[tuple, dict] = {}
        self._list_buttons: dict[str, ctk.CTkButton] = {}

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_list()
        self._build_form()
        app.bind_observer(self._on_app_event)
        self.reload_list(app.app_state.get("active_profile") or C.DEFAULT_PROFILE_KEY)

    # ------------------------------------------------------------------
    # Список профилей
    # ------------------------------------------------------------------
    def _build_list(self) -> None:
        frame = ctk.CTkFrame(self, width=LIST_W)
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        frame.grid_columnconfigure(0, weight=1)
        frame.grid_rowconfigure(0, weight=1)
        self.list_frame = ctk.CTkScrollableFrame(frame, label_text="")
        self.list_frame.grid(row=0, column=0, sticky="nsew", padx=4, pady=4)
        self.list_frame.grid_columnconfigure(0, weight=1)

        btns = ctk.CTkFrame(frame, fg_color="transparent")
        btns.grid(row=1, column=0, padx=4, pady=4)
        actions = (
            ("prof_new", self.new_profile),
            ("prof_duplicate", self.duplicate_profile),
            ("prof_rename", self.rename_profile),
            ("prof_delete", self.delete_profile),
            ("prof_import", self.import_json),
            ("prof_export", self.export_json),
            ("prof_make_active", self.make_active),
        )
        for i, (key, cmd) in enumerate(actions):
            ctk.CTkButton(btns, text=t(key), width=BTN_W, height=ROW_H // 2 * 2,
                          fg_color="gray40",
                          command=self.app.safe(cmd)).grid(
                row=i // 2, column=i % 2, padx=3, pady=3)

    def reload_list(self, select_key: str | None = None) -> None:
        for child in self.list_frame.winfo_children():
            child.destroy()
        self._list_buttons.clear()
        try:
            items = self.app.profile_store.list()
        except Exception as exc:
            self.app.log.log_sys(t("err_generic", why=exc))
            return
        active_key = self.app.app_state.get("active_profile") or C.DEFAULT_PROFILE_KEY
        for i, item in enumerate(items):
            key = item["key"]
            label = item.get("name") or key
            btn = ctk.CTkButton(self.list_frame, text=label, anchor="w", height=ROW_H,
                                fg_color="gray30" if key == active_key else "gray22",
                                command=lambda k=key: self.select_profile(k))
            btn.grid(row=i, column=0, sticky="ew", pady=2)
            self._list_buttons[key] = btn
            if key == active_key:
                ctk.CTkLabel(self.list_frame, text="●", width=ACTIVE_DOT_W).grid(
                    row=i, column=1, padx=2)
        if select_key and select_key in self._list_buttons and select_key != self._key:
            self.select_profile(select_key)

    def _refresh_active_marks(self) -> None:
        active_key = self.app.app_state.get("active_profile") or C.DEFAULT_PROFILE_KEY
        for key, btn in self._list_buttons.items():
            btn.configure(fg_color="gray30" if key == active_key else "gray22")

    # ------------------------------------------------------------------
    # Форма
    # ------------------------------------------------------------------
    def _build_form(self) -> None:
        scroll = ctk.CTkScrollableFrame(self, label_text="")
        scroll.grid(row=0, column=1, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)
        self._scroll = scroll
        self._row = 0

        self.acc_person = self._section("prof_section_person")
        self._add_field(self.acc_person, ("person", "fio"), "prof_fio", "text")
        self._add_field(self.acc_person, ("person", "gruppa"), "prof_gruppa", "text")
        self._add_field(self.acc_person, ("person", "prep"), "prof_prep", "text")
        self._add_field(self.acc_person, ("person", "status"), "prof_status", "select",
                        options=C.STATUS_CHOICES)

        self.acc_work = self._section("prof_section_work")
        self._add_field(self.acc_work, ("work", "work_type"), "prof_work_type", "select",
                        options=C.WORK_TYPES)
        self._add_field(self.acc_work, ("work", "pz_no"), "prof_pz_no", "text")
        self._add_field(self.acc_work, ("work", "theme"), "prof_theme", "text")
        self._add_field(self.acc_work, ("work", "disc"), "prof_disc", "text")
        self._add_field(self.acc_work, ("work", "city"), "prof_city", "text")
        self._add_field(self.acc_work, ("work", "year"), "prof_year", "num")
        self.custom_work_entry = ctk.CTkEntry(self.acc_work, width=EDIT_W,
                                              placeholder_text=t("prof_work_type_custom"))
        self.custom_work_entry.grid_forget()
        self.custom_work_entry.bind("<KeyRelease>", lambda _e: self._on_edit(), add="+")
        self.custom_work_entry.bind("<FocusOut>", lambda _e: self._sync_custom_work(),
                                    add="+")

        self.acc_inst = self._section("prof_section_institution")
        self._add_field(self.acc_inst, ("institution", "ministry"), "prof_ministry", "text")
        self._add_field(self.acc_inst, ("institution", "fsbe"), "prof_fsbe", "text")
        self._add_field(self.acc_inst, ("institution", "univ"), "prof_univ", "text")
        self._add_field(self.acc_inst, ("institution", "kaf"), "prof_kaf", "text")

        self.acc_extra = self._section("prof_section_extra")
        self.kv = KVEditor(self.acc_extra.body, on_change=self._on_edit)
        self.kv.pack(fill="x")

        self.acc_preview = self._section("prof_preview_title", expanded=True)
        self.preview_box = ctk.CTkTextbox(self.acc_preview.body, height=PREVIEW_H,
                                          wrap="word")
        self.preview_box.pack(fill="x")
        self.preview_box.configure(state="disabled")

        bar = ctk.CTkFrame(scroll, fg_color="transparent")
        bar.grid(row=self._row, column=0, sticky="ew", padx=4, pady=8)
        ctk.CTkButton(bar, text=t("prof_save"), width=SAVE_BTN_W,
                      command=self.app.safe(self.save_current)).pack(side="left", padx=4)
        ctk.CTkButton(bar, text=t("prof_revert"), width=BTN_W, fg_color="gray40",
                      command=self.app.safe(self.revert)).pack(side="left", padx=4)
        ctk.CTkButton(bar, text=t("prof_check_meta"), width=BTN_W, fg_color="gray40",
                      command=self.app.safe(self.show_flat_meta)).pack(side="left", padx=4)
        self.dirty_badge = Badge(bar, "", kind="warning")
        self.dirty_badge.pack(side="right", padx=4)
        self.dirty_badge.pack_forget()  # пустой бейдж — это беспричинная цветная плашка

    def _section(self, title_key: str, expanded: bool = True) -> SectionAccordion:
        acc = SectionAccordion(self._scroll, t(title_key), expanded=expanded)
        acc.grid(row=self._row, column=0, sticky="ew", pady=4)
        self._row += 1
        return acc

    def _add_field(self, acc: SectionAccordion, path: tuple, label_key: str,
                   ftype: str, options: tuple = ()) -> None:
        holder = ctk.CTkFrame(acc.body, fg_color="transparent")
        holder.pack(fill="x", pady=3)
        ctk.CTkLabel(holder, text=t(label_key), width=LABEL_W, anchor="w").pack(side="left")
        widget: Any
        if ftype == "select":
            widget = ctk.CTkOptionMenu(holder, values=list(options), width=EDIT_W // 2,
                                       command=lambda _v: self._on_select(path))
        else:
            widget = ctk.CTkEntry(holder, width=EDIT_W)
            widget.bind("<KeyRelease>", lambda _e: self._on_edit(), add="+")
        # Без pack виджет создаётся, но в строке не показывается.
        widget.pack(side="left", padx=(6, 0))
        self._fields[path] = {"widget": widget, "label_key": label_key, "type": ftype}

    # ------------------------------------------------------------------
    # Значения <-> виджеты
    # ------------------------------------------------------------------
    def select_profile(self, key: str, force: bool = False) -> None:
        if key == self._key and not force:
            return
        if self._dirty and not dialogs.confirm(self, t("tab_profile"),
                                               t("dlg_reset_confirm")):
            return
        try:
            self._base = self.app.profile_store.load(key)
        except Exception as exc:
            dialogs.exception_dialog(self, t("err_generic", why=exc), repr(exc))
            return
        self._key = key
        self._fill_form()
        self._mark_dirty(False)
        self._refresh_active_marks()

    def _fill_form(self) -> None:
        for path, ref in self._fields.items():
            value = self._base
            for part in path:
                value = value.get(part, {}) if isinstance(value, dict) else {}
            widget = ref["widget"]
            if isinstance(widget, ctk.CTkOptionMenu):
                widget.set(str(value) if value else widget.cget("values")[0])
            else:
                widget.delete(0, tk.END)
                text = "" if value == "" or value is None else str(value)
                if path == ("work", "year") and not text:
                    text = str(dt.date.today().year)  # год по умолчанию — текущий (п. 10.4)
                widget.insert(0, text)
        self.kv.set(self._base.get("extra") or {})
        self._sync_custom_work()
        self._update_preview()

    def _collect(self) -> dict:
        """Собрать профиль из формы, сохранив незнакомые ключи исходного JSON."""
        profile = json.loads(json.dumps(self._base, ensure_ascii=False))  # глубокая копия
        for path, ref in self._fields.items():
            widget = ref["widget"]
            if isinstance(widget, ctk.CTkOptionMenu):
                value = widget.get()
            else:
                value = widget.get().strip()
                if path == ("work", "year") and value.isdigit():
                    value = int(value)
            if path == ("work", "work_type") and value == _CUSTOM_WORK_TYPE:
                value = self.custom_work_entry.get().strip() or _CUSTOM_WORK_TYPE
            node = profile
            for part in path[:-1]:
                node = node.setdefault(part, {})
            node[path[-1]] = value
        profile["extra"] = self.kv.get()
        return profile

    def _on_edit(self) -> None:
        self._mark_dirty(True)
        self._update_preview()

    def _on_select(self, path: tuple) -> None:
        if path == ("work", "work_type"):
            self._sync_custom_work()
        self._on_edit()

    def _sync_custom_work(self) -> None:
        ref = self._fields.get(("work", "work_type"))
        selected = ref["widget"].get() if ref else ""
        if selected == _CUSTOM_WORK_TYPE:
            self.custom_work_entry.grid(in_=ref["widget"].master, row=1, column=0,
                                        sticky="w", pady=(2, 0))
            if not self.custom_work_entry.get():
                current = (self._base.get("work") or {}).get("work_type", "")
                if current and current not in C.WORK_TYPES:
                    self.custom_work_entry.insert(0, current)
        else:
            self.custom_work_entry.grid_forget()
        self._mark_dirty(True)
        self._update_preview()

    def _mark_dirty(self, dirty: bool) -> None:
        self._dirty = dirty
        self.dirty_badge.set_text(t("prof_unsaved") if dirty else "", "warning")
        self.dirty_badge.pack(side="right", padx=4) if dirty else self.dirty_badge.pack_forget()

    def _update_preview(self) -> None:
        try:
            text = profile_store.title_header_preview(self._collect())
        except Exception:
            text = ""  # превью не должно ронять вкладку на полудиных данных
        self.preview_box.configure(state="normal")
        self.preview_box.delete("1.0", "end")
        self.preview_box.insert("1.0", text)
        self.preview_box.configure(state="disabled")

    # ------------------------------------------------------------------
    # CRUD
    # ------------------------------------------------------------------
    def save_current(self) -> None:
        """Ctrl+S с вкладки; сохраняет активный или выбранный профиль."""
        if not self._key and not self._fields:
            return
        profile = self._collect()
        try:
            saved_key = self.app.profile_store.save(profile, key=self._key)
        except rules_store.CannotEditBuiltinError:
            new_key = dialogs.ask_text(self, t("prof_new"), t("ask_profile_name"),
                                       initial=f"{self._key}_custom")
            if not new_key:
                return
            profile["name"] = profile.get("name") or new_key
            saved_key = self.app.profile_store.save_as(profile, new_key)
            self._key = str(saved_key)
        self._base = profile
        self._mark_dirty(False)
        self.reload_list(str(saved_key))
        if str(saved_key) == self.app.app_state.get("active_profile"):
            self.app.set_profile(profile)
        self.app.log.log_sys(t("prof_saved", name=profile.get("name") or saved_key))

    def revert(self) -> None:
        self.select_profile(self._key, force=True)

    def new_profile(self) -> None:
        name = dialogs.ask_text(self, t("prof_new"), t("ask_profile_name"))
        if not name:
            return
        profile = profile_store.default_profile()
        profile["name"] = name
        key = self.app.profile_store.save(profile)
        self.app.log.log_sys(t("prof_created", key=key))
        self.reload_list(str(key))
        self.select_profile(str(key), force=True)

    def duplicate_profile(self) -> None:
        if not self._key:
            return
        source = self._collect()
        new_key = f"{self._key}_copy"
        profile = dict(source)
        profile["name"] = f"{source.get('name', self._key)} (копия)"
        saved = self.app.profile_store.save_as(profile, new_key)
        self.reload_list(str(saved))

    def rename_profile(self) -> None:
        if not self._key:
            return
        new_name = dialogs.ask_text(self, t("prof_rename"), t("ask_rename"),
                                    initial=self._base.get("name", self._key))
        if not new_name:
            return
        profile = self._collect()
        profile["name"] = new_name
        self.app.profile_store.save(profile, key=self._key)
        self._base = profile
        self._mark_dirty(False)
        self.reload_list(self._key)

    def delete_profile(self) -> None:
        if not self._key:
            return
        active = self.app.app_state.get("active_profile")
        if self._key == C.DEFAULT_PROFILE_KEY or self._key == active:
            dialogs.show_info(self, t("prof_delete"), t("rules_delete_builtin"))
            return
        self.app.profile_store.delete(self._key)
        self.app.log.log_sys(t("prof_deleted", name=self._base.get("name") or self._key))
        self._key = ""
        self._base = {}
        self.reload_list(active)

    def make_active(self) -> None:
        if not self._key:
            return
        if self._dirty:
            self.save_current()
        self.app.reload_profile(self._key)
        self.app.log.log_sys(t("prof_active_switched",
                               name=self._base.get("name") or self._key))
        self.reload_list(self._key)

    def import_json(self) -> None:
        path = dialogs.pick_file(self, t("prof_import"), filetypes=dialogs.FILETYPES_JSON)
        if not path:
            return
        key = self.app.profile_store.import_from(path)
        self.app.log.log_sys(t("prof_imported", key=key))
        self.reload_list(str(key))

    def export_json(self) -> None:
        if not self._key:
            return
        dest = dialogs.choose_save_path(self, t("prof_export"), defaultextension=".json",
                                        initialfile=f"{self._key}.json",
                                        filetypes=dialogs.FILETYPES_JSON,
                                        initialdir=self.app.output_dir())
        if not dest:
            return
        self.app.profile_store.export(self._key, dest)
        self.app.log.log_sys(t("prof_exported", path=dest))

    def show_flat_meta(self) -> None:
        meta = profile_store.flat_meta(self._collect())
        self.app.log.log_sys(json.dumps(meta, ensure_ascii=False, indent=2))
        self.app.log.log_sys(t("prof_meta_shown"))

    # ------------------------------------------------------------------
    def _on_app_event(self, topic: str) -> None:
        if topic == "profile" and not self._dirty:
            key = self.app.app_state.get("active_profile") or C.DEFAULT_PROFILE_KEY
            if key != self._key:
                self.reload_list(key)

    def cancel(self) -> None:
        pass  # вкладка не запускает долгих процессов
