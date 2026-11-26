"""Вкладка «Генерация» (п. 10.1, 8.2, 8.3).

Список скриптов, динамическая форма из SCRIPT_META, универсальные поля для
скриптов без меты, пред- и пост-операции. Вся исполнительная логика — в
`app.core.runner/plugin_loader/docx_*`; вкладка только собирает ввод, запускает
в рабочем потоке и отображает результат через очередь журнала.
"""

from __future__ import annotations

import datetime as dt
import os
import shlex
from typing import Any

import customtkinter as ctk

from app import constants as C
from app.core import plugin_loader, profile_store
from app.strings import t
from app.ui import dialogs
from app.ui.widgets import Badge, FilePickerRow, ParamField, SectionAccordion, with_tip
from app.ui.tabs.generator_run import GeneratorRunMixin

# Метрики компоновки вкладки.
LIST_W = 320
BTN_H = 32
SMALL_BTN_W = 130
BTN_H_SM = 28
HOOK_LABEL_W = 220
HOOK_MENU_W = 180
RUN_BTN_W = 260
WRAP = 480
NAME_TIME_FORMAT = "%Y-%m-%d_%H%M"  # автоимя выхода (п. 10.1)
PARAMS_TITLE_KEY = "gen_params"
_UNSAFE_FILENAME_CHARS = '<>:"/\\|?*'


def sanitize_filename(value: str) -> str:
    """Тема/название → безопасная основа имени файла (кириллица сохраняется)."""
    cleaned = "".join("_" if ch in _UNSAFE_FILENAME_CHARS else ch for ch in value).strip().strip(".")
    cleaned = cleaned.replace("\n", " ")[:80]
    return cleaned or "document"


class GeneratorTab(GeneratorRunMixin, ctk.CTkFrame):
    def __init__(self, master: Any, app: Any):
        super().__init__(master, fg_color="transparent")
        self.app = app
        self.selected: Any | None = None
        self._param_fields: dict[str, ParamField] = {}
        self._runner: Any | None = None
        self._thread: threading.Thread | None = None
        self._busy = False
        self._extra_paths: list[str] = []
        self._rules_touched = False

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_left()
        self._build_right()

        app.bind_observer(self._on_app_event)
        self.refresh_scripts(notify=False)

    # ------------------------------------------------------------------
    # Левая колонка: поиск + список скриптов
    # ------------------------------------------------------------------
    def _build_left(self) -> None:
        frame = ctk.CTkFrame(self, width=LIST_W)
        frame.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        frame.grid_propagate(False)
        frame.grid_columnconfigure(0, weight=1)

        self.search = ctk.CTkEntry(frame, placeholder_text=t("gen_search"))
        self.search.grid(row=0, column=0, padx=8, pady=6, sticky="ew")
        self.search.bind("<KeyRelease>", lambda _e: self._rebuild_list(), add="+")

        btns = ctk.CTkFrame(frame, fg_color="transparent")
        btns.grid(row=1, column=0, sticky="ew", padx=4)
        ctk.CTkButton(btns, text=t("gen_refresh"), height=BTN_H_SM, width=SMALL_BTN_W,
                      command=self.app.safe(self.refresh_scripts)).pack(side="left", padx=4)
        ctk.CTkButton(btns, text=t("gen_add_script"), height=BTN_H_SM, width=SMALL_BTN_W,
                      command=self.app.safe(self.add_script)).pack(side="left", padx=4)
        # Третья кнопка не влезает в колонку 320 px — она на отдельном ряду.
        btns2 = ctk.CTkFrame(frame, fg_color="transparent")
        btns2.grid(row=2, column=0, sticky="ew", padx=4)
        ctk.CTkButton(btns2, text=t("gen_open_scripts_dir"), height=BTN_H_SM,
                      fg_color="gray40",
                      command=self.app.safe(self.open_scripts_dir)).pack(side="left", padx=4)

        self.list_frame = ctk.CTkScrollableFrame(frame, label_text="")
        self.list_frame.grid(row=3, column=0, sticky="nsew", padx=4, pady=4)
        frame.grid_rowconfigure(3, weight=1)
        self.list_frame.grid_columnconfigure(0, weight=1)

    # ------------------------------------------------------------------
    # Правая колонка
    # ------------------------------------------------------------------
    def _build_right(self) -> None:
        scroll = ctk.CTkScrollableFrame(self, label_text="")
        scroll.grid(row=0, column=1, sticky="nsew")
        scroll.grid_columnconfigure(0, weight=1)
        self.right = scroll

        self.info_name = ctk.CTkLabel(scroll, text=t("gen_select_script"), anchor="w",
                                      font=ctk.CTkFont(weight="bold"))
        self.info_name.grid(row=0, column=0, sticky="w", padx=8, pady=(8, 0))
        self.info_desc = ctk.CTkLabel(scroll, text="", anchor="w", justify="left",
                                      wraplength=WRAP)
        self.info_desc.grid(row=1, column=0, sticky="w", padx=8)
        self.info_meta_line = ctk.CTkLabel(scroll, text="", anchor="w", wraplength=WRAP)
        self.info_meta_line.grid(row=2, column=0, sticky="w", padx=8)
        self.hook_badge = Badge(scroll, "rules_hook", kind="muted")
        self.hook_badge.grid(row=3, column=0, sticky="w", padx=8, pady=4)

        # -- форма параметров -------------------------------------------------
        self.params_acc = SectionAccordion(scroll, t(PARAMS_TITLE_KEY), expanded=True)
        self.params_acc.grid(row=4, column=0, sticky="ew", padx=8, pady=4)
        self.params_box = ctk.CTkFrame(self.params_acc.body, fg_color="transparent")
        self.params_box.pack(fill="x")
        self.params_hint = ctk.CTkLabel(self.params_acc.body, text="", anchor="w",
                                        justify="left", wraplength=WRAP)
        self.params_hint.pack(fill="x")

        # -- универсальные поля для скриптов без меты (п. 8.2) -----------------
        self.uni_acc = SectionAccordion(scroll, t("gen_universal"), expanded=False)
        self.uni_acc.grid(row=5, column=0, sticky="ew", padx=8, pady=4)
        ub = self.uni_acc.body
        self.uni_input = FilePickerRow(ub, label=t("gen_input"), kind="file",
                                       on_change=self._validate_ui)
        self.uni_input.pack(fill="x", pady=3)
        self.uni_args = ctk.CTkEntry(ub, placeholder_text=t("gen_extra_args"))
        self.uni_args.pack(fill="x", pady=3)
        self.uni_args.bind("<KeyRelease>", lambda _e: self._validate_ui(), add="+")
        hook_row = ctk.CTkFrame(ub, fg_color="transparent")
        hook_row.pack(fill="x", pady=3)
        ctk.CTkLabel(hook_row, text=t("gen_rules_hook"), width=HOOK_LABEL_W, anchor="w").pack(side="left")
        self.hook_menu = ctk.CTkOptionMenu(hook_row, values=list(C.RULES_HOOKS), width=HOOK_MENU_W)
        self.hook_menu.set(C.DEFAULT_RULES_HOOK)
        self.hook_menu.pack(side="left")
        ctk.CTkButton(ub, text=t("gen_save_desc"), width=SMALL_BTN_W, fg_color="gray40",
                      command=self.app.safe(self.save_description)).pack(anchor="w", pady=6)

        # -- метаданные ---------------------------------------------------------
        self.meta_acc = SectionAccordion(scroll, t("gen_meta_block"), expanded=False)
        self.meta_acc.grid(row=6, column=0, sticky="ew", padx=8, pady=4)
        mb = self.meta_acc.body
        self.use_profile = ctk.CTkCheckBox(mb, text=t("gen_meta_from_profile"))
        self.use_profile.select()
        self.use_profile.pack(anchor="w", pady=3)
        self.meta_summary = ctk.CTkLabel(mb, text="", anchor="w", justify="left",
                                         wraplength=WRAP)
        self.meta_summary.pack(anchor="w", pady=3)
        ctk.CTkButton(mb, text=t("gen_meta_edit"), width=SMALL_BTN_W, fg_color="gray40",
                      command=lambda: self.app.show_tab("profile")).pack(anchor="w", pady=3)

        # -- выходной файл -------------------------------------------------------
        self.out_row = FilePickerRow(scroll, label=t("gen_output"), kind="file_save",
                                     on_change=self._validate_ui)
        self.out_row.grid(row=7, column=0, sticky="ew", padx=8, pady=4)

        # -- пост-операции ---------------------------------------------------------
        cb_row = ctk.CTkFrame(scroll, fg_color="transparent")
        cb_row.grid(row=8, column=0, sticky="w", padx=8)
        self.cb_apply_rules = ctk.CTkCheckBox(cb_row, text=t("gen_check_apply_rules"),
                                              command=self._on_post_toggle)
        self.cb_check_after = ctk.CTkCheckBox(cb_row, text=t("gen_check_after"),
                                              command=self._on_post_toggle)
        self.cb_open_after = ctk.CTkCheckBox(cb_row, text=t("gen_check_open"),
                                             command=self._on_post_toggle)
        self.cb_check_after.select()
        for cb in (self.cb_apply_rules, self.cb_check_after, self.cb_open_after):
            cb.pack(anchor="w", pady=2)

        # -- действия ----------------------------------------------------------------
        actions = ctk.CTkFrame(scroll, fg_color="transparent")
        actions.grid(row=9, column=0, sticky="ew", padx=8, pady=8)
        self.run_btn = ctk.CTkButton(actions, text=t("btn_generate"), height=BTN_H + BTN_H_SM // 4,
                                     width=RUN_BTN_W,
                                     command=self.app.safe(self.generate))
        self.run_btn.pack(side="left")
        ctk.CTkButton(actions, text=t("log_cancel"), width=SMALL_BTN_W, fg_color="#8a3b3b",
                      command=self.app.safe(self.cancel)).pack(side="left", padx=8)
        ctk.CTkButton(actions, text=t("btn_open_output"), width=SMALL_BTN_W, fg_color="gray40",
                      command=self.app.safe(self.open_output_dir)).pack(side="left")
        self.result_badge = Badge(actions, "", kind="muted")
        self._result_badge_packed = False
        self.errors_lbl = ctk.CTkLabel(scroll, text="", anchor="w", justify="left",
                                       text_color="#e57373", wraplength=WRAP)
        self.errors_lbl.grid(row=10, column=0, sticky="w", padx=8, pady=(0, 8))

    # ------------------------------------------------------------------
    # Список скриптов
    # ------------------------------------------------------------------
    def refresh_scripts(self, notify: bool = True) -> None:
        self.app.reload_plugins(notify=False, extra_paths=tuple(self._extra_paths))
        if notify:
            self.app.notify("plugins")
        self._rebuild_list()
        self.log(t("gen_scanned", n=len(self.app.plugins)))

    def log(self, message: str) -> None:
        self.app.log.log_sys(message)

    def _rebuild_list(self) -> None:
        for child in self.list_frame.winfo_children():
            child.destroy()
        needle = self.search.get().strip().casefold()
        visible = [
            p for p in self.app.plugins
            if not needle or needle in (p.name or p.file_name or "").casefold()
            or needle in (p.file_name or "").casefold()
        ]
        if not visible:
            ctk.CTkLabel(self.list_frame, text=t("gen_no_scripts")).grid(row=0, column=0, pady=8)
            return
        groups: dict[str, list[Any]] = {}
        for info in visible:
            groups.setdefault(info.category or t("gen_no_category"), []).append(info)
        row = 0
        for category in sorted(groups):
            ctk.CTkLabel(self.list_frame, text=category, anchor="w",
                         font=ctk.CTkFont(weight="bold")).grid(
                row=row, column=0, sticky="w", padx=4, pady=(6, 0))
            row += 1
            for info in sorted(groups[category], key=lambda i: i.name or i.file_name):
                self._list_row(row, info)
                row += 1

    def _list_row(self, row: int, info: Any) -> None:
        frame = ctk.CTkFrame(self.list_frame, fg_color="transparent")
        frame.grid(row=row, column=0, sticky="ew", pady=2)
        frame.grid_columnconfigure(0, weight=1)
        label = info.name or info.file_name
        btn = with_tip(
            ctk.CTkButton(frame, text=label, anchor="w", height=BTN_H,
                          fg_color="gray35" if self.selected is info else "gray25",
                          command=lambda i=info: self.select(i)), label)
        btn.grid(row=0, column=0, sticky="ew")
        # Бейджи под названием: иначе на 320 px остаётся ~8 знаков имени.
        badges = ctk.CTkFrame(frame, fg_color="transparent")
        badges.grid(row=1, column=0, sticky="w", pady=(2, 0))
        if getattr(info, "error", None):
            Badge(badges, t("gen_broken_mark"), kind="error",
                  tooltip_text=str(info.error)).pack(side="left", padx=(0, 3))
        elif not getattr(info, "has_meta", True):
            Badge(badges, t("gen_no_meta_mark"), kind="muted").pack(side="left", padx=(0, 3))
        if getattr(info, "rules_hook", None) == "postprocess" and info.has_meta:
            Badge(badges, "postprocess", kind="info").pack(side="left")

    def select(self, info: Any) -> None:
        self.selected = info
        self._rebuild_list()
        self._render_info()
        self._rebuild_form()
        self._reset_output_name()
        self._refresh_meta_summary()
        self._apply_default_post_flags()
        self._validate_ui()

    def _render_info(self) -> None:
        info = self.selected
        if info is None:
            self.info_name.configure(text=t("gen_select_script"))
            self.info_desc.configure(text="")
            self.info_meta_line.configure(text="")
            self.hook_badge.configure(text="")
            return
        self.info_name.configure(text=info.name or info.file_name)
        self.info_desc.configure(text=info.description or "")
        self.info_meta_line.configure(text=(
            f"{t('gen_version', value=info.version or '?')}\n"
            f"{t('gen_path', value=info.path)}"
            + (f"\n{info.error}" if getattr(info, "error", None) else "")))
        hook = info.rules_hook or (self.hook_menu.get() if not info.has_meta else "")
        kind = {"native": "ok", "constants": "info", "postprocess": "warning",
                "none": "muted"}.get(hook or "", "muted")
        self.hook_badge.set_text(f"rules_hook: {hook}" if hook else "", kind)

    # ------------------------------------------------------------------
    # Форма параметров
    # ------------------------------------------------------------------
    def _rebuild_form(self) -> None:
        for child in self.params_box.winfo_children():
            child.destroy()
        self._param_fields.clear()
        info = self.selected
        if info is None:
            self.params_hint.configure(text=t("gen_select_script"))
            self.uni_acc.set_expanded(False)
            return
        specs = list(getattr(info, "params", None) or [])
        meta = self._flat_meta()
        for i, spec in enumerate(specs):
            field = ParamField(self.params_box, spec, on_change=self._on_param_change)
            if self.use_profile.get() and field.param_id in meta and specs:
                preset = meta[field.param_id]
                if preset not in (None, ""):
                    field.set(preset)
            field.grid(row=i, column=0, sticky="ew", pady=3)
            self.params_box.grid_columnconfigure(0, weight=1)
            self._param_fields[field.param_id] = field
        has_input = info.has_meta and (info.input or "none") != "none"
        show_input = has_input or not info.has_meta
        self.uni_input.pack_forget()
        if show_input:
            self.uni_input.pack(fill="x", pady=3, before=self.uni_args)
        self.uni_acc.set_expanded(not info.has_meta)
        self.params_hint.configure(text="" if specs or not info.has_meta
                                   else t("gen_no_meta_mark"))

    def _on_param_change(self) -> None:
        self._sync_output_name()
        self._validate_ui()

    def _on_post_toggle(self) -> None:
        self._rules_touched = True
        self._validate_ui()

    # ------------------------------------------------------------------
    # Метаданные / профиль
    # ------------------------------------------------------------------
    def _flat_meta(self) -> dict:
        try:
            return profile_store.flat_meta(self.app.profile)
        except Exception:
            return {}

    def _refresh_meta_summary(self) -> None:
        meta = self._flat_meta()
        if not meta:
            self.meta_summary.configure(text=t("gen_meta_summary_empty"))
            return
        pairs = "; ".join(f"{k}={v}" for k, v in sorted(meta.items()) if v not in (None, ""))
        self.meta_summary.configure(text=pairs or t("gen_meta_summary_empty"))

    def _apply_default_post_flags(self) -> None:
        """Чекбокс нормализации включается при strict (п. 10.1), если юзер не трогал."""
        if not self._rules_touched:
            if self.app.rules.get("strict"):
                self.cb_apply_rules.select()
            else:
                self.cb_apply_rules.deselect()

    def _on_app_event(self, topic: str) -> None:
        if topic == "plugins":
            self._rebuild_list()
        elif topic == "profile":
            self._refresh_meta_summary()
            if self.selected is not None:
                self._reset_output_name()
        elif topic == "rules":
            self._apply_default_post_flags()
            self._render_info()

    # ------------------------------------------------------------------
    # Выходной файл (автоимя `<тема|name>_<дата_время>.docx`, п. 10.1)
    # ------------------------------------------------------------------
    def _default_output_name(self) -> str:
        info = self.selected
        meta = self._flat_meta()
        # Тема из формы важнее профиля: тот же порядок, что у job._default_output (п. 19.7).
        theme = self._param_fields.get("theme")
        base = ((theme.get() if theme else "") or meta.get("theme") or meta.get("name")
                or (info.name if info else "")
                or (info.file_name if info else "") or "document")
        ext = (info.output_ext if info and info.output_ext else ".docx")
        stamp = dt.datetime.now().strftime(NAME_TIME_FORMAT)
        return f"{sanitize_filename(str(base))}_{stamp}{ext}"

    def _reset_output_name(self) -> None:
        name = os.path.join(self.app.output_dir(), self._default_output_name())
        self._auto_output = name
        self.out_row.set(name)

    def _sync_output_name(self) -> None:
        """Пересобрать автоимя, пока пользователь не вписал свой путь в поле вывода."""
        if self.selected is None:
            return
        if self.out_row.get() != getattr(self, "_auto_output", None):
            return
        self._reset_output_name()

    def collect_params(self) -> dict:
        params = {pid: f.get() for pid, f in self._param_fields.items()}
        info = self.selected
        if info is not None and not info.has_meta:
            extra = self.uni_args.get().strip()
            if extra:
                # Токены из свободной строки едут скрипту внутри JSON-параметров:
                # контракт build_argv (п. INTERFACES) не имеет отдельного канала argv.
                try:
                    params["_extra_args"] = shlex.split(extra)
                except ValueError:
                    params["_extra_args"] = extra.split()
            params["_rules_hook"] = self.hook_menu.get()
        return params

    # ------------------------------------------------------------------
    # Сохранение описания безметового скрипта (sidecar, п. 8.2)
    # ------------------------------------------------------------------
    def save_description(self) -> None:
        info = self.selected
        if info is None:
            return
        overrides = {
            "name": info.name or os.path.splitext(info.file_name)[0],
            "description": info.description or "",
            "category": info.category or t("gen_no_category"),
            "version": info.version or "1.0",
            "input": "none" if self.uni_input.get() == "" and not info.has_meta
            else (info.input or "file_docx"),
            "output_ext": info.output_ext or ".docx",
            "rules_hook": self.hook_menu.get() if not info.has_meta
            else (info.rules_hook or C.DEFAULT_RULES_HOOK),
            "params": [f.spec for f in self._param_fields.values()] or list(info.params or []),
        }
        path = plugin_loader.save_sidecar(info, overrides)
        self.log(t("gen_desc_saved", path=path))
        self.refresh_scripts()

    # ------------------------------------------------------------------
    # Добавление/обновление скриптов
    # ------------------------------------------------------------------
    def add_script(self) -> None:
        path = dialogs.pick_file(self, t("gen_add_script"), filetypes=dialogs.FILETYPES_PY)
        if not path:
            return
        if not dialogs.warn_external_script(self):
            return
        inside = os.path.dirname(os.path.abspath(path)) == os.path.abspath(
            self.app.scripts_dir())
        if not inside and dialogs.confirm(
                self, t("gen_add_script"), t("gen_add_copy", path=self.app.scripts_dir())):
            import shutil
            try:
                os.makedirs(self.app.scripts_dir(), exist_ok=True)
                shutil.copy2(path, os.path.join(self.app.scripts_dir(), os.path.basename(path)))
            except OSError as exc:
                self.log(t("gen_add_copy_error", why=exc))
                return
            self.log(t("gen_script_copied", path=self.app.scripts_dir()))
            self.refresh_scripts()
            return
        self._extra_paths.append(path)
        self.refresh_scripts()

    def set_result_badge(self, text: str, kind: str = "muted") -> None:
        """Пустой бейдж не рисуется: Tk показывает вместо него серую заглушку."""
        if text and not self._result_badge_packed:
            self.result_badge.pack(side="left", padx=8)
            self._result_badge_packed = True
        elif not text and self._result_badge_packed:
            self.result_badge.pack_forget()
            self._result_badge_packed = False
        self.result_badge.set_text(text, kind)

    def open_scripts_dir(self) -> None:
        err = dialogs.open_path(self.app.scripts_dir())
        if err:
            self.log(t("open_error", path=self.app.scripts_dir(), why=err))

    def open_output_dir(self) -> None:
        err = dialogs.open_path(self.app.output_dir())
        if err:
            self.log(t("open_error", path=self.app.output_dir(), why=err))

    # ------------------------------------------------------------------
    # Валидация до запуска (инлайн, п. 10.1)
    # ------------------------------------------------------------------
    def validate(self) -> list[str]:
        errors: list[str] = []
        info = self.selected
        if info is None:
            return [t("gen_select_script")]
        if getattr(info, "error", None):
            return [f"{t('gen_broken_mark')}: {info.error}"]
        for field in self._param_fields.values():
            err = field.validate()
            if err:
                errors.append(err)
        needs_input = (info.has_meta and (info.input or "none") in
                       ("file_md", "file_docx", "folder")) or not info.has_meta
        input_path = self.uni_input.get()
        if needs_input:
            if not input_path:
                errors.append(t("gen_err_required", label=t("gen_input")))
            elif not os.path.exists(input_path):
                errors.append(t("gen_err_input_missing"))
        out = self.out_row.get()
        if not out:
            errors.append(t("gen_err_required", label=t("gen_output")))
        else:
            out_dir = os.path.dirname(os.path.abspath(out))
            if not os.path.isdir(out_dir):
                parent = os.path.dirname(out_dir)
                if not (os.path.isdir(parent) and os.access(parent, os.W_OK)):
                    errors.append(t("gen_err_output_dir"))
            elif not os.access(out_dir, os.W_OK):
                errors.append(t("gen_err_output_dir"))
            want_ext = info.output_ext or ".docx"
            if not out.lower().endswith(want_ext.lower()):
                errors.append(t("gen_err_ext", ext=want_ext))
        try:
            from app.core import job
            problems = job.validate_job(self._job_dict())
            errors.extend(t("gen_err_job", details=x) for x in problems)
        except ImportError:
            pass  # job-модуль факультативен для UI-пути
        return errors

    def _job_dict(self) -> dict:
        """Конверт задачи по схеме §19.7 — используется и для валидации, и для запуска."""
        info = self.selected
        job = {
            "schema_version": C.SCHEMA_VERSION,
            "script": info.file_name if info else "",
            "params": self.collect_params(),
            "rules": self.app.app_state.get("active_preset") or C.DEFAULT_PRESET_KEY,
            "profile": self.app.app_state.get("active_profile") or C.DEFAULT_PROFILE_KEY,
            "output": self.out_row.get(),
            "post": {
                "normalize": bool(self.cb_apply_rules.get()),
                "check": bool(self.cb_check_after.get()),
                "fail_on": "error",
            },
        }
        # Пустой input не передаём вовсе: null не проходит валидацию схемы (п. 19.7).
        if self.uni_input.get():
            job["input"] = self.uni_input.get()
        return job

    def _validate_ui(self) -> None:
        if self._busy:
            return
        errors = self.validate()
        self.errors_lbl.configure(text="\n".join(errors))
        self.run_btn.configure(state="disabled" if errors else "normal")

