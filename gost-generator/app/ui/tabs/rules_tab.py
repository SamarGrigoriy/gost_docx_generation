"""Вкладка «Правила оформления» (п. 10.2).

Аккордеоны секций по группам правил, строгий режим через validate_rules,
CRUD пресетов через RulesStore, нормализация готового .docx через core.
Пересчёты единиц (мм/см/пт↔twips) выполняет core; UI оперирует человекочитаемыми
значениями из JSON-пресета как есть.
"""

from __future__ import annotations

import copy
import os
import threading
import tkinter as tk
from typing import Any

import customtkinter as ctk

from app import constants as C
from app.core import docx_inspector, docx_normalizer, rules_store
from app.strings import t
from app.ui import dialogs
from app.ui.widgets import Badge, SectionAccordion, with_tip

# Метрики вкладки.
LABEL_W = 260
EDIT_W = 200
TOP_BTN_W = 140
WRAP = 560
POLL_MS = 200
PRESET_LABEL_W = 90
PRESET_MENU_W = 260
ROW_H = 32

# Схемы секций: path — путь в JSON пресета; options — кортеж из constants.
# tip — подсказка «откуда требование» (пункт методички).
SECTION_SPECS: list[dict] = [
    {"title": "sec_page", "tip": "tip_page", "fields": [
        {"path": ("page", "size"), "label": "f_page_size", "type": "select",
         "options": C.PAGE_SIZES, "tip": "tip_page"},
        {"path": ("page", "orientation"), "label": "f_orientation", "type": "select",
         "options": C.ORIENTATIONS, "tip": "tip_page"},
        {"path": ("page", "margin_left_mm"), "label": "f_margin_left", "type": "num", "tip": "tip_margins"},
        {"path": ("page", "margin_right_mm"), "label": "f_margin_right", "type": "num", "tip": "tip_margins"},
        {"path": ("page", "margin_top_mm"), "label": "f_margin_top", "type": "num", "tip": "tip_margins"},
        {"path": ("page", "margin_bottom_mm"), "label": "f_margin_bottom", "type": "num", "tip": "tip_margins"},
        {"path": ("page", "footer_distance_mm"), "label": "f_footer_distance", "type": "num",
         "tip": "tip_footer"},
    ]},
    {"title": "sec_text", "tip": "tip_text", "fields": [
        {"path": ("text", "font"), "label": "f_font", "type": "select",
         "options": C.FONT_CHOICES, "tip": "tip_font"},
        {"path": ("text", "size_pt"), "label": "f_size_pt", "type": "num", "tip": "tip_text"},
        {"path": ("text", "line_spacing"), "label": "f_line_spacing", "type": "select",
         "options": tuple(str(s) for s in C.LINE_SPACING_CHOICES), "tip": "tip_spacing",
         "cast": float},
        {"path": ("text", "first_line_indent_cm"), "label": "f_indent_cm", "type": "num",
         "tip": "tip_text"},
        {"path": ("text", "alignment"), "label": "f_alignment", "type": "select",
         "options": C.ALIGNMENTS, "tip": "tip_text"},
        {"path": ("text", "space_before_pt"), "label": "f_space_before", "type": "num",
         "tip": "tip_text"},
        {"path": ("text", "space_after_pt"), "label": "f_space_after", "type": "num",
         "tip": "tip_text"},
        {"path": ("text", "allow_italic_for_terms_only"), "label": "f_allow_italic",
         "type": "bool", "tip": "tip_text"},
        {"path": ("text", "hyphenation"), "label": "f_hyphenation", "type": "bool",
         "tip": "tip_text"},
    ]},
    {"title": "sec_headings", "tip": "tip_headings", "fields": []},  # заполняет _HEADINGS_SPECS
    {"title": "sec_figures", "tip": "tip_figures", "fields": [
        {"path": ("figures", "caption_template"), "label": "f_fig_template", "type": "text",
         "tip": "tip_figures"},
        {"path": ("figures", "caption_position"), "label": "f_fig_position", "type": "select",
         "options": ("below", "above"), "tip": "tip_figures"},
        {"path": ("figures", "caption_align"), "label": "f_fig_align", "type": "select",
         "options": C.ALIGNMENTS, "tip": "tip_figures"},
        {"path": ("figures", "caption_indent"), "label": "f_fig_indent", "type": "bool",
         "tip": "tip_figures"},
        {"path": ("figures", "caption_no_trailing_dot"), "label": "f_fig_no_dot",
         "type": "bool", "tip": "tip_figures"},
        {"path": ("figures", "numbering"), "label": "f_fig_numbering", "type": "select",
         "options": ("continuous", "per_section"), "tip": "tip_figures"},
        {"path": ("figures", "max_width_cm"), "label": "f_fig_max_width", "type": "num",
         "tip": "tip_figures"},
        {"path": ("figures", "max_height_cm"), "label": "f_fig_max_height", "type": "num",
         "tip": "tip_figures"},
        {"path": ("figures", "require_reference_in_text"), "label": "f_fig_require_ref",
         "type": "bool", "tip": "tip_figures"},
    ]},
    {"title": "sec_tables", "tip": "tip_tables", "fields": [
        {"path": ("tables", "caption_template"), "label": "f_tbl_template", "type": "text",
         "tip": "tip_tables"},
        {"path": ("tables", "caption_position"), "label": "f_tbl_position", "type": "select",
         "options": ("above", "below"), "tip": "tip_tables"},
        {"path": ("tables", "caption_align"), "label": "f_tbl_align", "type": "select",
         "options": C.ALIGNMENTS, "tip": "tip_tables"},
        {"path": ("tables", "caption_indent"), "label": "f_tbl_indent", "type": "bool",
         "tip": "tip_tables"},
        {"path": ("tables", "continuation_text"), "label": "f_tbl_continuation",
         "type": "text", "tip": "tip_tables"},
        {"path": ("tables", "continuation_align"), "label": "f_tbl_cont_align",
         "type": "select", "options": C.ALIGNMENTS, "tip": "tip_tables"},
        {"path": ("tables", "repeat_header"), "label": "f_tbl_repeat_header", "type": "bool",
         "tip": "tip_tables"},
        {"path": ("tables", "header_bold"), "label": "f_tbl_header_bold", "type": "bool",
         "tip": "tip_tables"},
        {"path": ("tables", "borders"), "label": "f_tbl_borders", "type": "select",
         "options": ("single_thin", "none", "grid"), "tip": "tip_tables"},
        {"path": ("tables", "border_size_eighths_pt"), "label": "f_tbl_border_size",
         "type": "num", "tip": "tip_tables"},
        {"path": ("tables", "font_size_pt"), "label": "f_tbl_font_size", "type": "num",
         "tip": "tip_tables"},
        {"path": ("tables", "line_spacing"), "label": "f_tbl_line_spacing", "type": "num",
         "tip": "tip_tables"},
        {"path": ("tables", "allow_row_split"), "label": "f_tbl_allow_split", "type": "bool",
         "tip": "tip_tables"},
        {"path": ("tables", "require_reference_in_text"), "label": "f_tbl_require_ref",
         "type": "bool", "tip": "tip_tables"},
    ]},
    {"title": "sec_formulas_notes", "tip": "tip_formulas", "fields": [
        {"path": ("formulas", "own_line"), "label": "f_form_own_line", "type": "bool",
         "tip": "tip_formulas"},
        {"path": ("formulas", "align"), "label": "f_form_align", "type": "select",
         "options": C.ALIGNMENTS, "tip": "tip_formulas"},
        {"path": ("formulas", "blank_line_above_below"), "label": "f_form_blank",
         "type": "bool", "tip": "tip_formulas"},
        {"path": ("formulas", "number_format"), "label": "f_form_num_format", "type": "text",
         "tip": "tip_formulas"},
        {"path": ("formulas", "number_align"), "label": "f_form_num_align", "type": "select",
         "options": C.ALIGNMENTS, "tip": "tip_formulas"},
        {"path": ("formulas", "explanation_first_word"), "label": "f_form_expl_word",
         "type": "text", "tip": "tip_formulas"},
        {"path": ("formulas", "explanation_no_colon"), "label": "f_form_expl_no_colon",
         "type": "bool", "tip": "tip_formulas"},
        {"path": ("notes", "prefix"), "label": "f_note_prefix", "type": "text",
         "tip": "tip_formulas"},
        {"path": ("notes", "dash"), "label": "f_note_dash", "type": "text",
         "tip": "tip_formulas"},
        {"path": ("notes", "from_indent"), "label": "f_note_from_indent", "type": "bool",
         "tip": "tip_formulas"},
        {"path": ("notes", "numbered_when_multiple"), "label": "f_note_numbered",
         "type": "bool", "tip": "tip_formulas"},
        {"path": ("footnotes", "font"), "label": "f_fn_font", "type": "select",
         "options": C.FONT_CHOICES, "tip": "tip_formulas"},
        {"path": ("footnotes", "size_pt"), "label": "f_fn_size", "type": "num",
         "tip": "tip_formulas"},
        {"path": ("footnotes", "min_allowed_size_pt"), "label": "f_fn_min_size",
         "type": "num", "tip": "tip_formulas"},
    ]},
    {"title": "sec_references", "tip": "tip_references", "fields": [
        {"path": ("references", "style"), "label": "f_ref_style", "type": "select",
         "options": ("square_brackets", "superscript", "author_year"), "tip": "tip_references"},
        {"path": ("references", "bibliography_order"), "label": "f_ref_order", "type": "select",
         "options": ("by_appearance", "alphabetical"), "tip": "tip_references"},
        {"path": ("references", "max_source_age_years"), "label": "f_ref_max_age",
         "type": "num", "tip": "tip_references"},
        {"path": ("references", "min_cited_percent"), "label": "f_ref_min_cited",
         "type": "num", "tip": "tip_references"},
    ]},
    {"title": "sec_misc", "tip": "tip_misc", "fields": [
        {"path": ("appendices", "label"), "label": "f_app_label", "type": "text", "tip": "tip_misc"},
        {"path": ("appendices", "new_page"), "label": "f_app_new_page", "type": "bool",
         "tip": "tip_misc"},
        {"path": ("toc", "required"), "label": "f_toc_required", "type": "bool", "tip": "tip_misc"},
        {"path": ("toc", "auto"), "label": "f_toc_auto", "type": "bool", "tip": "tip_misc"},
        {"path": ("numbering", "position"), "label": "f_num_position", "type": "select",
         "options": ("bottom_center", "top_center", "bottom_right"), "tip": "tip_misc"},
        {"path": ("numbering", "title_page_numbered"), "label": "f_num_title", "type": "bool",
         "tip": "tip_misc"},
        {"path": ("numbering", "font_size_pt"), "label": "f_num_size", "type": "num",
         "tip": "tip_misc"},
        {"path": ("volume", "min_page_fill_percent"), "label": "f_vol_min_fill", "type": "num",
         "tip": "tip_misc"},
        {"path": ("volume", "min_last_page_fill_percent"), "label": "f_vol_min_last",
         "type": "num", "tip": "tip_misc"},
        {"path": ("code_blocks", "font"), "label": "f_code_font", "type": "select",
         "options": C.MONOSPACE_FONTS, "tip": "tip_misc"},
        {"path": ("code_blocks", "size_pt"), "label": "f_code_size", "type": "num",
         "tip": "tip_misc"},
        {"path": ("code_blocks", "line_spacing"), "label": "f_code_spacing", "type": "num",
         "tip": "tip_misc"},
        {"path": ("code_blocks", "indent"), "label": "f_code_indent", "type": "bool",
         "tip": "tip_misc"},
    ]},
]

# Секция «Заголовки»: генерируется из flat-спецификации (все флаги §7.1).
_HEADINGS_SPECS: list[dict] = [
    {"path": ("headings", "structural_uppercase"), "label": "f_struct_upper", "type": "bool",
     "tip": "tip_headings"},
    {"path": ("headings", "structural_bold"), "label": "f_struct_bold", "type": "bool",
     "tip": "tip_headings"},
    {"path": ("headings", "structural_align"), "label": "f_struct_align", "type": "select",
     "options": C.ALIGNMENTS, "tip": "tip_headings"},
    {"path": ("headings", "structural_new_page"), "label": "f_struct_new_page", "type": "bool",
     "tip": "tip_headings"},
    {"path": ("headings", "structural_no_trailing_dot"), "label": "f_struct_no_dot",
     "type": "bool", "tip": "tip_headings"},
    {"path": ("headings", "section_bold"), "label": "f_section_bold", "type": "bool",
     "tip": "tip_headings"},
    {"path": ("headings", "section_from_indent"), "label": "f_section_from_indent",
     "type": "bool", "tip": "tip_headings"},
    {"path": ("headings", "section_numbering"), "label": "f_section_numbering", "type": "select",
     "options": ("1", "none"), "tip": "tip_headings"},
    {"path": ("headings", "subsection_numbering"), "label": "f_subsection_numbering",
     "type": "select", "options": ("1.1", "none"), "tip": "tip_headings"},
    {"path": ("headings", "no_trailing_dot"), "label": "f_head_no_dot", "type": "bool",
     "tip": "tip_headings"},
    {"path": ("headings", "multiline_line_spacing"), "label": "f_multiline_spacing",
     "type": "select", "options": tuple(str(s) for s in C.LINE_SPACING_CHOICES),
     "cast": float, "tip": "tip_headings"},
    {"path": ("headings", "keep_with_next"), "label": "f_keep_next", "type": "bool",
     "tip": "tip_headings"},
    {"path": ("headings", "min_lines_after"), "label": "f_min_lines_after", "type": "num",
     "tip": "tip_headings"},
    {"path": ("headings", "same_size_as_body"), "label": "f_same_size", "type": "bool",
     "tip": "tip_headings"},
]


def _get_path(data: dict, path: tuple) -> Any:
    node: Any = data
    for key in path:
        if not isinstance(node, dict) or key not in node:
            return None
        node = node[key]
    return node


def _set_path(data: dict, path: tuple, value: Any) -> None:
    node = data
    for key in path[:-1]:
        node = node.setdefault(key, {})
    node[path[-1]] = value


def _builtin_value(path: tuple) -> Any:
    return _get_path(rules_store.DEFAULT_RULES, path)


class RulesTab(ctk.CTkFrame):
    def __init__(self, master: Any, app: Any):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self._rules: dict = {}
        self._key: str = ""
        self._dirty = False
        self._widgets: dict[tuple, dict] = {}
        self._busy = False
        self._presets_display: dict[str, str] = {}

        self.grid_rowconfigure(3, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self._build_top()
        self._build_body()
        app.bind_observer(self._on_app_event)
        self.load_preset(app.app_state.get("active_preset") or C.DEFAULT_PRESET_KEY)

    # ------------------------------------------------------------------
    # Верхняя панель
    # ------------------------------------------------------------------
    def _build_top(self) -> None:
        top = ctk.CTkFrame(self)
        top.grid(row=0, column=0, sticky="ew", padx=8, pady=(8, 0))
        ctk.CTkLabel(top, text=t("rules_preset"), width=PRESET_LABEL_W, anchor="w").grid(
            row=0, column=0, padx=8, pady=8)
        self.preset_menu = ctk.CTkOptionMenu(top, values=["…"], width=PRESET_MENU_W,
                                             command=self._on_preset_selected)
        self.preset_menu.grid(row=0, column=1, padx=4, pady=8)
        actions = (
            ("rules_save", self.save_current),
            ("rules_save_as", self.save_as),
            ("rules_reset", self.reset_to_builtin),
            ("rules_delete", self.delete_preset),
            ("rules_import", self.import_json),
            ("rules_export", self.export_json),
            ("rules_preview", self.preview),
            ("rules_normalize_btn", self.normalize_docx),
        )
        # 8 кнопок в один ряд не влезают даже при 1200 px — разбиваем на два ряда.
        for i, group in enumerate((actions[:4], actions[4:])):
            row_frame = ctk.CTkFrame(self, fg_color="transparent")
            row_frame.grid(row=1 + i, column=0, sticky="ew", padx=8)
            for key, cmd in group:
                ctk.CTkButton(row_frame, text=t(key), width=TOP_BTN_W, height=ROW_H,
                              fg_color="gray40" if key != "rules_save" else None,
                              command=self.app.safe(cmd)).pack(side="left", padx=3, pady=4)
        self.strict_switch = with_tip(
            ctk.CTkSwitch(row_frame, text=t("rules_strict"),
                          command=self._on_strict_toggle), t("rules_strict_tip"))
        self.strict_switch.pack(side="left", padx=12)
        self.dirty_badge = Badge(self, "", kind="warning")
        # Пустой бейдж — это цветная плашка без текста; показываем его только когда есть что сказать.
        self.dirty_badge.grid(row=2, column=0, sticky="ne", padx=8)
        self.dirty_badge.grid_remove()

    def _build_body(self) -> None:
        self.scroll = ctk.CTkScrollableFrame(self, label_text="")
        self.scroll.grid(row=3, column=0, sticky="nsew", padx=8, pady=6)
        self.scroll.grid_columnconfigure(0, weight=1)
        for i, section in enumerate(self._all_sections()):
            acc = SectionAccordion(self.scroll, t(section["title"]),
                                   tooltip=t(section.get("tip") or "tip_page"),
                                   expanded=(i == 0))
            acc.grid(row=i, column=0, sticky="ew", pady=4)
            for j, spec in enumerate(section["fields"]):
                self._build_field(acc.body, spec, row=j)

    @staticmethod
    def _all_sections() -> list[dict]:
        sections = []
        for section in SECTION_SPECS:
            if section["title"] == "sec_headings":
                sections.append({"title": "sec_headings", "tip": "tip_headings",
                                 "fields": _HEADINGS_SPECS})
            else:
                sections.append(section)
        return sections

    def _build_field(self, body: ctk.CTkFrame, spec: dict, row: int) -> None:
        path = tuple(spec["path"])
        label = t(spec["label"])
        tip = t(spec.get("tip") or "tip_page")
        ptype = spec["type"]
        holder = ctk.CTkFrame(body, fg_color="transparent")
        holder.grid(row=row, column=0, sticky="ew", pady=2)
        body.grid_columnconfigure(0, weight=1)
        with_tip(ctk.CTkLabel(holder, text=label, width=LABEL_W, anchor="w"),
                 tip).grid(row=0, column=0, padx=(0, 6))
        widget: Any
        if ptype == "bool":
            widget = with_tip(ctk.CTkSwitch(holder, text=""), tip)
            widget.configure(command=lambda p=path, w=widget: self._on_change(p, w))
        elif ptype == "select":
            values = [str(v) for v in spec.get("options") or ()] or [""]
            widget = with_tip(ctk.CTkOptionMenu(holder, values=values, width=EDIT_W), tip)
            widget.configure(command=lambda _v, p=path, w=widget: self._on_change(p, w))
        else:
            widget = with_tip(ctk.CTkEntry(holder, width=EDIT_W), tip)
            widget.bind("<FocusOut>", lambda _e, p=path, w=widget: self._on_change(p, w),
                        add="+")
            widget.bind("<Return>", lambda _e, p=path, w=widget: self._on_change(p, w),
                        add="+")
        self._widgets[path] = {"spec": spec, "widget": widget}
        # Виджет значения — вторая колонка строки; без grid он создаётся, но не виден.
        holder.grid_columnconfigure(1, weight=1)
        widget.grid(row=0, column=1, sticky="w")

    # ------------------------------------------------------------------
    # Значения <-> виджеты
    # ------------------------------------------------------------------
    def _fill_widgets(self) -> None:
        for path, ref in self._widgets.items():
            value = _get_path(self._rules, path)
            widget = ref["widget"]
            if isinstance(widget, ctk.CTkSwitch):
                (widget.select if value else widget.deselect)()
            elif isinstance(widget, ctk.CTkOptionMenu):
                widget.set(str(value) if value is not None else "")
            else:
                widget.delete(0, tk.END)
                if value is not None:
                    widget.insert(0, str(value))

    def _widget_value(self, ref: dict) -> Any:
        widget = ref["widget"]
        spec = ref["spec"]
        if isinstance(widget, ctk.CTkSwitch):
            var = widget.get_variable() if hasattr(widget, "get_variable") else None
            return bool(var.get()) if var is not None else bool(widget.get())
        raw = widget.get()
        if spec["type"] == "select" and spec.get("cast") is float:
            return float(str(raw).replace(",", "."))
        if spec["type"] == "num":
            text = str(raw).replace(",", ".").strip()
            return float(text) if "." in text else int(text or 0)
        return raw

    def _on_change(self, path: tuple, _widget: Any) -> None:
        if self._busy:
            return
        ref = self._widgets[path]
        try:
            value = self._widget_value(ref)
        except ValueError:
            value = _get_path(self._rules, path)
        candidate = copy.deepcopy(self._rules)
        _set_path(candidate, path, value)
        problems = rules_store.validate_rules(candidate)
        if self.strict_on() and problems:
            fallback = _builtin_value(path)
            _set_path(self._rules, path, fallback if fallback is not None
                      else _get_path(self._rules, path))
            self._busy = True
            self._fill_widgets()
            self._busy = False
            self.app.log.log_sys(t("rules_rejected", why=problems[0]))
        else:
            self._rules = candidate
            if problems:
                self.app.log.log_sys(t("rules_warnings", details="; ".join(problems)))
        self._mark_dirty(True)
        self.app.set_rules(self._rules, notify=False)

    def _on_strict_toggle(self) -> None:
        strict = bool(self.strict_switch.get_variable().get()
                      if hasattr(self.strict_switch, "get_variable")
                      else self.strict_switch.get())
        self._rules["strict"] = strict
        if strict:
            candidate = copy.deepcopy(self._rules)
            problems = rules_store.validate_rules(candidate)
            if problems:
                # «приведение значений к пресету» (п. 10.2): спорные поля — к встроенным
                for path in self._widgets:
                    _set_path(self._rules, path, _builtin_value(path))
                self._rules["strict"] = True
                self._busy = True
                self._fill_widgets()
                self._busy = False
                self.app.log.log_sys(t("rules_rejected", why="; ".join(problems[:2])))
            self._lock_beyond_preset(True)
        else:
            self._lock_beyond_preset(False)
        self._mark_dirty(True)
        self.app.set_rules(self._rules, notify=False)
        self.app.update_status()

    def _lock_beyond_preset(self, locked: bool) -> None:
        """Строгий режим: free-form правки, ведущие за границы пресета, блокируются
        через disabled-состояния числовых полей (значения остаются валидными)."""
        for ref in self._widgets.values():
            widget = ref["widget"]
            if isinstance(widget, ctk.CTkEntry):
                widget.configure(state="disabled" if locked else "normal")

    def strict_on(self) -> bool:
        var = self.strict_switch.get_variable() if hasattr(self.strict_switch, "get_variable") \
            else None
        return bool(var.get()) if var is not None else bool(self.strict_switch.get())

    def _mark_dirty(self, dirty: bool) -> None:
        self._dirty = dirty
        if dirty:
            self.dirty_badge.set_text(t("rules_unsaved"), "warning")
            self.dirty_badge.grid()
        else:
            self.dirty_badge.set_text("", "muted")
            self.dirty_badge.grid_remove()

    # ------------------------------------------------------------------
    # Пресеты: CRUD
    # ------------------------------------------------------------------
    def load_presets_menu(self, select_key: str | None = None) -> None:
        items = self.app.rules_store.list()
        displays = []
        self._presets_display = {}
        for item in items:
            name = item.get("name") or item.get("key")
            disp = t("rules_builtin_suffix", name=name) if item.get("builtin") else name
            self._presets_display[disp] = item["key"]
            displays.append(disp)
        self.preset_menu.configure(values=displays or ["…"])
        if select_key:
            for disp, key in self._presets_display.items():
                if key == select_key:
                    self.preset_menu.set(disp)
                    break

    def load_preset(self, key: str) -> None:
        self.load_presets_menu(key)
        self._key = key
        self._rules = copy.deepcopy(self.app.rules or rules_store.DEFAULT_RULES)
        self._fill_widgets()
        self.strict_switch.select() if self._rules.get("strict") else self.strict_switch.deselect()
        self._on_strict_visual_state()
        self._mark_dirty(False)

    def _on_strict_visual_state(self) -> None:
        self._lock_beyond_preset(self.strict_on())

    def _on_preset_selected(self, display: str) -> None:
        key = self._presets_display.get(display)
        if not key or key == self._key:
            return
        if self._dirty and not dialogs.confirm(
                self, t("rules_preset"), t("dlg_reset_confirm")):
            self.load_presets_menu(self._key)
            return
        self.app.reload_rules(key)

    def _on_app_event(self, topic: str) -> None:
        if topic == "rules" and self._rules is not self.app.rules:
            self.load_preset(self.app.app_state.get("active_preset") or C.DEFAULT_PRESET_KEY)

    def save_current(self) -> None:
        """Ctrl+S / «Сохранить». Встроенный пресет не перезаписывается (AGENTS.md)."""
        try:
            saved_key = self.app.rules_store.save(self._rules, key=self._key)
        except rules_store.CannotEditBuiltinError:
            new_key = dialogs.ask_text(self, t("rules_save_as"), t("ask_preset_key"),
                                       initial=f"{self._key}_custom")
            if not new_key:
                return
            saved_key = self.app.rules_store.save_as(self._rules, new_key)
            self._key = str(saved_key)
        problems = rules_store.validate_rules(self._rules)
        if problems:
            self.app.log.log_sys(t("rules_warnings", details="; ".join(problems)))
        self.app.app_state["active_preset"] = self._key
        self.app.set_rules(self._rules)
        self.app.update_status()
        self.app.save_state()
        self.load_presets_menu(self._key)
        self._mark_dirty(False)
        self.app.log.log_sys(t("rules_saved", key=self._key))

    def save_as(self) -> None:
        new_key = dialogs.ask_text(self, t("rules_save_as"), t("ask_preset_key"),
                                   initial=f"{self._key}_copy")
        if not new_key:
            return
        self.app.rules_store.save_as(self._rules, new_key)
        self._key = new_key
        self.load_presets_menu(new_key)
        self._mark_dirty(False)
        self.app.log.log_sys(t("rules_saved", key=new_key))

    def reset_to_builtin(self) -> None:
        self._rules = copy.deepcopy(rules_store.DEFAULT_RULES)
        self._fill_widgets()
        self._mark_dirty(True)
        self.app.log.log_sys(t("rules_reset_done"))

    def delete_preset(self) -> None:
        items = {i["key"]: i for i in self.app.rules_store.list()}
        item = items.get(self._key)
        if item is None or item.get("builtin"):
            dialogs.show_info(self, t("rules_delete"), t("rules_delete_builtin"))
            return
        self.app.rules_store.delete(self._key)
        self.app.log.log_sys(t("rules_deleted", key=self._key))
        self._key = C.DEFAULT_PRESET_KEY
        self.app.reload_rules(self._key)

    def import_json(self) -> None:
        path = dialogs.pick_file(self, t("rules_import"), filetypes=dialogs.FILETYPES_JSON)
        if not path:
            return
        key = self.app.rules_store.import_from(path)
        self.app.log.log_sys(t("rules_imported", key=key))
        self.app.reload_rules(str(key))

    def export_json(self) -> None:
        dest = dialogs.choose_save_path(self, t("rules_export"),
                                        defaultextension=".json",
                                        initialfile=f"{self._key}.json",
                                        filetypes=dialogs.FILETYPES_JSON,
                                        initialdir=self.app.output_dir())
        if not dest:
            return
        self.app.rules_store.export(self._key, dest)
        self.app.log.log_sys(t("rules_exported", path=dest))

    # ------------------------------------------------------------------
    # Предпросмотр и нормализация готового файла
    # ------------------------------------------------------------------
    def preview(self) -> None:
        text = rules_store.describe_rules(self._rules)
        self.app.log.log_sys(text)

    def normalize_docx(self) -> None:
        """«Применить к готовому .docx…»: normalize_docx в потоке, изменения в лог."""
        if self._busy:
            return
        path = dialogs.pick_file(self, t("norm_pick"), filetypes=dialogs.FILETYPES_DOCX)
        if not path:
            return
        self._busy = True
        self.app.log.set_running(True, t("chk_running"))

        def worker() -> None:
            try:
                res = docx_normalizer.normalize_docx(path, self._rules, backup=True)
            except docx_inspector.DocxError as exc:
                res = {"error": str(exc)}
            except Exception as exc:
                res = {"error": repr(exc)}
            self.app.post(lambda: self._on_normalized(path, res))

        threading.Thread(target=worker, daemon=True).start()

    def cancel(self) -> None:
        pass  # нормализация — короткая синхронная операция core, потока-раннера нет

    def _on_normalized(self, path: str, res: dict) -> None:
        self._busy = False
        self.app.log.set_running(False)
        if res.get("error"):
            self.app.log.log_sys(t("err_generic", why=res["error"]))
            return
        changes = res.get("changes") or []
        if not res.get("changed"):
            self.app.log.log_sys(t("gen_normalized_none"))
        else:
            self.app.log.log_sys(t("gen_normalized", n=len(changes)))
            for ch in changes:
                self.app.log.log(f"  · {ch}")
        if res.get("backup"):
            self.app.log.log_sys(t("gen_norm_backup", path=res["backup"]))
            if os.path.exists(str(res["backup"])):
                self.app.log.log_sys(t("norm_backup_exists", path=res["backup"]))
