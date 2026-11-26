"""Главное окно (п. 10): сайдбар, вкладки, лог-панель, статус-строка, горячие клавиши.

Бизнес-логика целиком в `app.core`; здесь — компоновка, маршрутизация событий и
мост «поток → UI» через очереди. `customtkinter` импортируется только в слое UI.
"""

from __future__ import annotations

import datetime as dt
import glob
import os
import queue
import sys
import threading
import traceback
from collections.abc import Mapping
from typing import Any, Callable

import customtkinter as ctk

from app import constants as C
from app.strings import t
from app.core import config, plugin_loader, profile_store, rules_store
from app.ui import dialogs
from app.ui.log_panel import LogPanel
from app.ui.tabs.checker_tab import CheckerTab
from app.ui.tabs.generator_tab import GeneratorTab
from app.ui.tabs.profile_tab import ProfileTab
from app.ui.tabs.rules_tab import RulesTab

# Метрики компоновки окна — именованные (габариты окна берутся из constants).
SIDEBAR_WIDTH = 200
SIDEBAR_BTN_H = 40
SIDEBAR_PAD = 10
STATUS_BAR_H = 28
LOG_SPACING = 4
SELECTED_BTN_BG = "gray25"
FLAT_RADIUS = 0  # панельные фреймы без скругления
LOGO_FALLBACK_SIZE = 24
LOGO_MAX_H = 96  # иначе квадратный логотип съедает половину боковой панели
BOTTOM_BTN_H = 30
SCREEN_MARGIN_W = 60   # запас на рамку и тень окна со стороны WM
SCREEN_MARGIN_H = 90   # панель задач сверху/снизу + заголовок окна
POSITION_MARGIN = 40   # запас для позиции окна: нижний край не должен съезжать с экрана
SIDEBAR_MIN_HEIGHT = 480  # логотип + 4 вкладки + «Настройки» + «О программе» + «Тема»
LOG_PANEL_MIN_HEIGHT = 120
LOG_PANEL_CHROME = 58     # заголовок с кнопками + отступы вокруг текстового поля
LOGO_SMALL_H = 56         # логотип на низком экране (1366x768 и меньше)
LOGO_FULL_SCREEN_H = 800

TAB_ORDER = ("generator", "rules", "checker", "profile")
_TAB_TITLE_KEYS = {
    "generator": "tab_generator",
    "rules": "tab_rules",
    "checker": "tab_checker",
    "profile": "tab_profile",
}
THEME_LABELS = {m: f"{t('lbl_theme')}: {m}" for m in C.UI_THEMES}


class GostApp(ctk.CTk):
    """Оболочка приложения: владеет состоянием, сторами, вкладками и журналом."""

    def __init__(self, dirs: Any):
        super().__init__()
        self.dirs = dirs
        dirs.ensure()

        try:
            self.app_state: dict = config.load_state(dirs)
        except Exception as exc:  # битый state.json: core делает бэкап, UI предупреждает
            self.app_state = config.default_state(dirs)
            dialogs.exception_dialog(self, t("err_generic", why=exc), traceback.format_exc())

        self._ui_queue: queue.Queue[Callable[[], None]] = queue.Queue()
        self._observers: list[Callable[[str], None]] = []
        self.active_tab = "generator"

        self.rules_store = rules_store.RulesStore(dirs)
        self.profile_store = profile_store.ProfileStore(dirs)
        self.rules: dict = {}
        self.profile: dict = {}
        self.plugins: list = []

        self._apply_look_and_feel()
        self.title(t("app_title"))
        self.geometry(self._geometry_string(self.app_state.get("window_geometry"),
                                            self._screen_default_w(),
                                            self._screen_default_h()))
        self.minsize(C.WINDOW_MIN_WIDTH, C.WINDOW_MIN_HEIGHT)
        self.protocol("WM_DELETE_WINDOW", self._on_close)

        # Данные грузятся до вкладок: вкладки читают app.rules/profile/plugins при сборке.
        self.reload_rules(self.app_state.get("active_preset") or C.DEFAULT_PRESET_KEY,
                          notify=False)
        self.reload_profile(self.app_state.get("active_profile") or C.DEFAULT_PROFILE_KEY,
                            notify=False)
        self.reload_plugins(notify=False)

        self._build_layout()
        self._bind_hotkeys()

        self._open_session_log()
        self._install_exception_hooks()
        self._poll_ui_queue()

        if not self.app_state.get("welcome_shown"):
            def _done(dont_show: bool) -> None:
                self.app_state["welcome_shown"] = True
                self.save_state()
            dialogs.welcome_dialog(self, _done)

        self.log.log_sys(t("status_version", version=C.APP_VERSION) + f" · {dirs.home}")

    # ------------------------------------------------------------------
    # Оформление: тема/масштаб (п. 10)
    # ------------------------------------------------------------------
    def _apply_look_and_feel(self) -> None:
        theme = self.app_state.get("theme") or C.DEFAULT_THEME
        if theme not in C.UI_THEMES:
            theme = C.DEFAULT_THEME
        ctk.set_appearance_mode(theme)
        scale_pct = int(self.app_state.get("ui_scale") or C.DEFAULT_UI_SCALE)
        if scale_pct not in C.UI_SCALES:
            scale_pct = C.DEFAULT_UI_SCALE
        factor = scale_pct / 100.0
        ctk.set_widget_scaling(factor)
        ctk.set_window_scaling(factor)
        self._theme = theme
        self._scale_pct = scale_pct

    def cycle_theme(self) -> None:
        idx = C.UI_THEMES.index(self._theme) if self._theme in C.UI_THEMES else 0
        self._theme = C.UI_THEMES[(idx + 1) % len(C.UI_THEMES)]
        self.app_state["theme"] = self._theme
        ctk.set_appearance_mode(self._theme)
        if hasattr(self, "_theme_btn"):
            self._theme_btn.configure(text=THEME_LABELS[self._theme])
        self.save_state()

    # ------------------------------------------------------------------
    # Компоновка
    # ------------------------------------------------------------------
    def _build_layout(self) -> None:
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)

        sidebar = ctk.CTkFrame(self, width=SIDEBAR_WIDTH, corner_radius=FLAT_RADIUS)
        sidebar.grid(row=0, column=0, sticky="ns")
        sidebar.grid_propagate(False)
        sidebar.grid_columnconfigure(0, weight=1)

        logo_box = ctk.CTkFrame(sidebar, fg_color="transparent")
        logo_box.grid(row=0, column=0, pady=(SIDEBAR_PAD * 2, SIDEBAR_PAD * 2),
                      padx=SIDEBAR_PAD, sticky="ew")
        self._build_logo(logo_box)

        self._tab_btns: dict[str, ctk.CTkButton] = {}
        for i, name in enumerate(TAB_ORDER, start=1):
            btn = ctk.CTkButton(
                sidebar, text=t(_TAB_TITLE_KEYS[name]), anchor="w",
                height=SIDEBAR_BTN_H, command=lambda n=name: self.show_tab(n))
            btn.grid(row=i, column=0, padx=SIDEBAR_PAD, pady=LOG_SPACING, sticky="ew")
            self._tab_btns[name] = btn
        self._btn_default_fg = self._tab_btns["generator"].cget("fg_color")

        bottom = ctk.CTkFrame(sidebar, fg_color="transparent")
        bottom.grid(row=len(TAB_ORDER) + 1, column=0, pady=SIDEBAR_PAD,
                    padx=SIDEBAR_PAD, sticky="s")
        # Строка-«пружина» прижимает «Настройки/О программе/Тема» к низу панели.
        sidebar.grid_rowconfigure(len(TAB_ORDER) + 1, weight=1)
        # Три кнопки в столбик: подписи «Настройки», «О программе», «Тема: …» должны
        # помещаться целиком, поэтому две в ряд недопустимы (200 px сайдбара не делятся).
        ctk.CTkButton(bottom, text=t("btn_settings"), anchor="w", height=BOTTOM_BTN_H,
                      command=self.safe(self.open_settings)).pack(fill="x", pady=LOG_SPACING)
        ctk.CTkButton(bottom, text=t("btn_about"), anchor="w", height=BOTTOM_BTN_H,
                      command=self.safe(self.open_about)).pack(fill="x", pady=LOG_SPACING)
        self._theme_btn = ctk.CTkButton(bottom, text=THEME_LABELS.get(self._theme, t("lbl_theme")),
                                        anchor="w", height=BOTTOM_BTN_H,
                                        fg_color="gray35",
                                        command=self.safe(self.cycle_theme))
        self._theme_btn.pack(fill="x", pady=LOG_SPACING)

        self.content = ctk.CTkFrame(self, fg_color="transparent")
        self.content.grid(row=0, column=1, sticky="nsew", padx=(0, SIDEBAR_PAD))
        self.content.grid_columnconfigure(0, weight=1)
        self.content.grid_rowconfigure(0, weight=1)

        # Лог-панель создаётся до вкладок: вкладки пишут в журнал при своей сборке.
        self.log = LogPanel(self, height=self._screen_log_height(),
                            on_cancel=self.safe(self.cancel_active))
        self.log.grid(row=1, column=0, columnspan=2,
                      sticky="ew", padx=SIDEBAR_PAD, pady=LOG_SPACING)

        self.tabs: dict[str, Any] = {
            "generator": GeneratorTab(self.content, self),
            "rules": RulesTab(self.content, self),
            "checker": CheckerTab(self.content, self),
            "profile": ProfileTab(self.content, self),
        }
        for tab in self.tabs.values():
            tab.grid(row=0, column=0, sticky="nsew")

        self.statusbar = ctk.CTkFrame(self, height=STATUS_BAR_H, corner_radius=FLAT_RADIUS)
        self.statusbar.grid(row=2, column=0, columnspan=2, sticky="ew")
        self.status_labels: dict[str, ctk.CTkLabel] = {}
        for key in ("preset", "profile", "output"):
            lbl = ctk.CTkLabel(self.statusbar, text="", anchor="w")
            lbl.pack(side="left", padx=SIDEBAR_PAD)
            self.status_labels[key] = lbl
        ver = ctk.CTkLabel(self.statusbar, text=t("status_version", version=C.APP_VERSION),
                           anchor="e")
        ver.pack(side="right", padx=SIDEBAR_PAD)
        self.update_status()
        self.show_tab("generator")

    def _build_logo(self, box: ctk.CTkFrame) -> None:
        """Логотип assets/logo.png через PIL; нет файла/PIL — текстовая заглушка."""
        candidates = [
            os.path.join(getattr(sys, "_MEIPASS", ""), "assets", "logo.png"),
            os.path.join(str(self.dirs.project_dir), "assets", "logo.png"),
        ]
        try:
            from PIL import Image  # pillow разрешён AGENTS.md
            for path in candidates:
                if path and os.path.isfile(path):
                    img = Image.open(path)
                    ratio = img.height / max(img.width, 1)
                    # На низком экране логотип уступает место журналу и нижним кнопкам.
                    max_h = LOGO_MAX_H if self.winfo_screenheight() >= LOGO_FULL_SCREEN_H \
                        else LOGO_SMALL_H
                    w = min(SIDEBAR_WIDTH - SIDEBAR_PAD * 4, int(max_h / max(ratio, 0.01)))
                    self._logo_img = ctk.CTkImage(light_image=img, dark_image=img,
                                                  size=(w, int(w * ratio)))
                    ctk.CTkLabel(box, image=self._logo_img, text="").pack()
                    ctk.CTkLabel(box, text=t("app_title")).pack()
                    return
        except Exception:
            pass  # отсутствует PIL или файл — допустимая текстовая заглушка
        ctk.CTkLabel(box, text=t("logo_placeholder"),
                     font=ctk.CTkFont(size=LOGO_FALLBACK_SIZE, weight="bold")).pack()
        ctk.CTkLabel(box, text=t("app_title")).pack()

    # ------------------------------------------------------------------
    # Навигация
    # ------------------------------------------------------------------
    def show_tab(self, name: str) -> None:
        if name not in self.tabs:
            return
        self.active_tab = name
        self.tabs[name].tkraise()
        for key, btn in self._tab_btns.items():
            btn.configure(fg_color=SELECTED_BTN_BG if key == name else self._btn_default_fg)

    def bind_observer(self, cb: Callable[[str], None]) -> None:
        """Уведомление вкладок об изменении общих данных (rules/profile/plugins)."""
        self._observers.append(cb)

    def notify(self, topic: str) -> None:
        """Публичный вход в _notify для вкладок."""
        self._notify(topic)

    def _notify(self, topic: str) -> None:
        for cb in list(self._observers):
            try:
                cb(topic)
            except Exception as exc:
                self.log.log(f"traceback: {traceback.format_exc()}")
                dialogs.exception_dialog(self, t("err_generic", why=exc), traceback.format_exc())

    # ------------------------------------------------------------------
    # Данные (тонкие обёртки над core; вызывается только из главного потока)
    # ------------------------------------------------------------------
    def reload_plugins(self, notify: bool = True, extra_paths: tuple = ()) -> None:
        scripts_dir = self.app_state.get("scripts_dir") or self.dirs.scripts
        paths = list(extra_paths)
        # Штатные генераторы лежат в репозитории/exe, а не в каталоге данных — показываем их тоже.
        bundled = self.dirs.bundled_scripts()
        if bundled is not None:
            paths.extend(sorted(bundled.glob("*.py")))
        try:
            self.plugins = plugin_loader.scan_dir(scripts_dir, extra_paths=tuple(paths))
        except Exception as exc:
            self.plugins = []
            self.log.log_sys(t("err_generic", why=exc))
        if notify:
            self._notify("plugins")

    def reload_rules(self, key: str, notify: bool = True) -> None:
        try:
            self.rules = self.rules_store.load(key)
            self.app_state["active_preset"] = key
        except FileNotFoundError:
            self.rules = dict(rules_store.DEFAULT_RULES)
            self.app_state["active_preset"] = C.DEFAULT_PRESET_KEY
            self.log.log_sys(t("rules_reset_done"))
        except Exception as exc:
            self.rules = dict(rules_store.DEFAULT_RULES)
            dialogs.exception_dialog(self, t("err_generic", why=exc), traceback.format_exc())
        self.update_status()
        if notify:
            self._notify("rules")
        self.save_state()

    def reload_profile(self, key: str, notify: bool = True) -> None:
        try:
            self.profile = self.profile_store.load(key)
            self.app_state["active_profile"] = key
        except FileNotFoundError:
            self.profile = profile_store.default_profile()
            self.app_state["active_profile"] = C.DEFAULT_PROFILE_KEY
        except Exception as exc:
            self.profile = profile_store.default_profile()
            dialogs.exception_dialog(self, t("err_generic", why=exc), traceback.format_exc())
        self.update_status()
        if notify:
            self._notify("profile")
        self.save_state()

    def set_rules(self, rules: dict, notify: bool = True) -> None:
        """Обновить активные правила после правок вкладки «Правила»."""
        self.rules = rules
        if notify:
            self._notify("rules")

    def set_profile(self, profile: dict, notify: bool = True) -> None:
        self.profile = profile
        if notify:
            self._notify("profile")

    def output_dir(self) -> str:
        return str(self.app_state.get("output_dir") or self.dirs.output)

    def add_recent(self, path: str) -> None:
        """Недавние файлы в state (п. 7.3), не больше MAX_RECENT_FILES."""
        recent = [p for p in self.app_state.get("recent_files") or [] if p != path]
        recent.insert(0, path)
        self.app_state["recent_files"] = recent[: C.MAX_RECENT_FILES]
        self.save_state()

    def scripts_dir(self) -> str:
        return str(self.app_state.get("scripts_dir") or self.dirs.scripts)

    def runner_kwargs(self) -> dict:
        """Параметры ScriptRunner из настроек (п. 3.2, 10.5)."""
        mode = "external" if self.app_state.get("interpreter_mode") == "external" else "internal"
        return {
            "timeout_s": int(self.app_state.get("script_timeout_s") or C.DEFAULT_SCRIPT_TIMEOUT_S),
            "mode": mode,
            "python_path": self.app_state.get("interpreter_path") if mode == "external" else None,
        }

    # ------------------------------------------------------------------
    # Статус-строка
    # ------------------------------------------------------------------
    def update_status(self) -> None:
        if not hasattr(self, "status_labels"):
            return  # статус-строка ещё не собрана (данные грузятся до layout)
        preset = self.app_state.get("active_preset") or C.DEFAULT_PRESET_KEY
        prof = self.profile.get("name") or self.app_state.get("active_profile") or C.DEFAULT_PROFILE_KEY
        self.status_labels["preset"].configure(
            text=t("status_preset", value=f"{preset}{' · strict' if self.rules.get('strict') else ''}"))
        self.status_labels["profile"].configure(text=t("status_profile", value=prof))
        self.status_labels["output"].configure(text=t("status_output", value=self.output_dir()))

    # ------------------------------------------------------------------
    # Мост «поток → UI» и безопасность колбэков
    # ------------------------------------------------------------------
    def post(self, fn: Callable[[], None]) -> None:
        """Выполнить fn в главном потоке (безопасно из рабочих потоков)."""
        self._ui_queue.put(fn)

    def _poll_ui_queue(self) -> None:
        try:
            while True:
                fn = self._ui_queue.get_nowait()
                try:
                    fn()
                except Exception as exc:
                    self.log.log(f"traceback: {traceback.format_exc()}")
                    dialogs.exception_dialog(self, t("err_generic", why=exc),
                                             traceback.format_exc())
        except queue.Empty:
            pass
        self.after(UI_QUEUE_POLL_MS, self._poll_ui_queue)

    def safe(self, fn: Callable) -> Callable:
        """Обёртка колбэка: исключение → журнал + диалог, UI не падает (п. 12)."""

        def wrapper(*call_args: Any, **call_kwargs: Any) -> None:
            try:
                fn(*call_args, **call_kwargs)
            except Exception as exc:
                text = traceback.format_exc()
                self.log.log(f"traceback: {text}")
                # Диалог — Tk-виджет: из рабочего потока (Thread(target=safe(...)))
                # его показывать нельзя, поэтому ставим его в UI-очередь.
                self.post(lambda: dialogs.exception_dialog(
                    self, t("err_generic", why=exc), text))

        return wrapper

    # ------------------------------------------------------------------
    # Горячие клавиши (п. 10)
    # ------------------------------------------------------------------
    def _bind_hotkeys(self) -> None:
        self.bind_all("<Control-g>", lambda _e: self.safe(self.tabs["generator"].generate)())
        self.bind_all("<Control-k>", lambda _e: self.safe(self.hotkey_check)())
        self.bind_all("<Control-s>", lambda _e: self.safe(self.hotkey_save)())
        self.bind_all("<F5>", lambda _e: self.safe(self.tabs["generator"].refresh_scripts)())
        self.bind_all("<Control-l>", lambda _e: self.safe(self.toggle_log)())
        self.bind_all("<Escape>", lambda _e: self.safe(self.cancel_active)())

    def hotkey_check(self) -> None:
        self.show_tab("checker")
        self.tabs["checker"].run_check()

    def hotkey_save(self) -> None:
        tab = self.tabs.get(self.active_tab)
        saver = getattr(tab, "save_current", None)
        if callable(saver):
            saver()
        else:
            self.save_state()

    def toggle_log(self) -> None:
        self.log.toggle_collapse()

    def cancel_active(self) -> None:
        tab = self.tabs.get(self.active_tab)
        canceller = getattr(tab, "cancel", None)
        if callable(canceller):
            canceller()

    # ------------------------------------------------------------------
    # Настройки / о программе
    # ------------------------------------------------------------------
    def open_settings(self) -> None:
        dialogs.settings_dialog(self, self.app_state, self.dirs, self.safe(self._apply_settings))

    def _apply_settings(self, new_state: dict) -> None:
        old_portable = bool(self.app_state.get("portable"))
        self.app_state.update(new_state)
        self._apply_look_and_feel()
        self.save_state()
        self.reload_plugins()
        self.update_status()
        self.log.log_sys(t("set_saved"))
        if bool(self.app_state.get("portable")) != old_portable:
            self._toggle_portable_flag(bool(self.app_state.get("portable")))
            dialogs.show_info(self, t("dlg_settings_title"), t("set_portable_restart"))

    def _toggle_portable_flag(self, enabled: bool) -> None:
        """Файл-маркер portable рядом с exe/репозиторием (иерархия п. 19.6)."""
        flag = dialogs.portable_flag_path(self.dirs)
        try:
            if enabled:
                with open(flag, "w", encoding="utf-8") as fh:
                    fh.write("1")
            elif os.path.exists(flag):
                os.remove(flag)
        except OSError as exc:  # нет прав — понятная строка в журнал, не падение
            self.log.log_sys(t("err_generic", why=exc))

    def open_about(self) -> None:
        dialogs.about_dialog(self, str(self.dirs.home), bool(self.dirs.is_portable()))

    # ------------------------------------------------------------------
    # Состояние / сессионный лог / закрытие
    # ------------------------------------------------------------------
    def _screen_default_w(self) -> int:
        return max(C.WINDOW_MIN_WIDTH,
                   min(C.WINDOW_WIDTH, self.winfo_screenwidth() - SCREEN_MARGIN_W))

    def _screen_default_h(self) -> int:
        """Размер по умолчанию, вписанный в экран: 800 px не влезают в 1366x768 с панелью."""
        return max(C.WINDOW_MIN_HEIGHT,
                   min(C.WINDOW_HEIGHT, self.winfo_screenheight() - SCREEN_MARGIN_H))

    def _screen_log_height(self) -> int:
        """Журнал не должен отрезать нижние кнопки сайдбара на низком экране."""
        room = self._screen_default_h() - SIDEBAR_MIN_HEIGHT - STATUS_BAR_H - LOG_PANEL_CHROME
        return max(LOG_PANEL_MIN_HEIGHT, min(C.LOG_PANEL_HEIGHT, room))

    def _geometry_string(self, value: Any, default_w: int, default_h: int) -> str:
        """`window_geometry` из state.json (словарь схемы п. 7.3) → строка Tk `WxH+X+Y`."""
        if isinstance(value, str) and value:      # старое состояние могло хранить строку
            return value
        if not isinstance(value, Mapping):
            return f"{default_w}x{default_h}"
        width = int(value.get("width") or default_w)
        height = int(value.get("height") or default_h)
        # Окно ещё не успело отобразиться (см. _geometry_dict) — сохранялись 1x1:
        # такой размер не воспроизводим, берём размер по умолчанию.
        if width < C.WINDOW_MIN_WIDTH or height < C.WINDOW_MIN_HEIGHT:
            return f"{default_w}x{default_h}"
        width = min(width, self.winfo_screenwidth() - SCREEN_MARGIN_W)
        height = min(height, self.winfo_screenheight() - SCREEN_MARGIN_H)
        x, y = value.get("x"), value.get("y")
        if x is None or y is None:
            return f"{width}x{height}"
        # Сохранённая позиция могла увести нижний край за экран — вместе с панелью журнала.
        x = max(0, min(int(x), self.winfo_screenwidth() - width - POSITION_MARGIN))
        y = max(0, min(int(y), self.winfo_screenheight() - height - POSITION_MARGIN))
        return f"{width}x{height}{x:+d}{y:+d}"

    def _geometry_dict(self) -> dict[str, Any]:
        """Текущая геометрия окна в словарь схемы (п. 7.3).

        `wm geometry` до вывода окна на экран отдаёт запрошенные «1x1+0+0», поэтому
        берём фактические размеры; если окно ещё не отображено — храним прежнее значение.
        """
        stored = self.app_state.get("window_geometry")
        if not self.winfo_ismapped():
            return dict(stored) if isinstance(stored, Mapping) else {
                "width": C.WINDOW_WIDTH, "height": C.WINDOW_HEIGHT, "x": None, "y": None}
        return {"width": self.winfo_width(), "height": self.winfo_height(),
                "x": self.winfo_rootx(), "y": self.winfo_rooty()}

    def save_state(self) -> None:
        try:
            self.app_state["window_geometry"] = self._geometry_dict()
            config.save_state(self.dirs, self.app_state)
        except Exception as exc:
            self.log.log_sys(t("err_generic", why=exc))

    def _open_session_log(self) -> None:
        stamp = dt.datetime.now().strftime(C.LOG_TIMESTAMP_FORMAT)
        path = os.path.join(str(self.dirs.logs), f"{stamp}.log")
        try:
            os.makedirs(os.path.dirname(path), exist_ok=True)
            self._session_fh = open(path, "w", encoding="utf-8", errors="replace")
            self._session_path = path
            self.log.set_session_sink(self._write_session_line)
        except OSError:
            self._session_fh = None
            self._session_path = None

    def _write_session_line(self, line: str, kind: str) -> None:
        if self._session_fh is None:
            return
        prefix = "[stderr] " if kind == "stderr" else "[stdout] "
        self._session_fh.write(prefix + line.rstrip("\n") + "\n")
        try:
            # Без flush строки лежат в буфере, и падение приложения теряет весь журнал.
            self._session_fh.flush()
        except OSError:
            pass

    def _rotate_logs(self) -> None:
        """Ротация сессионных логов: держим LOG_KEEP_COUNT старых—новых файлов (п. 12)."""
        try:
            files = sorted(glob.glob(os.path.join(str(self.dirs.logs), "*.log")))
            for old in files[:-C.LOG_KEEP_COUNT] if len(files) > C.LOG_KEEP_COUNT else []:
                os.remove(old)
        except OSError:
            pass

    def _install_exception_hooks(self) -> None:
        def hook(exc_type: Any, exc: BaseException, tb: Any) -> None:
            text = "".join(traceback.format_exception(exc_type, exc, tb))
            self.post(lambda: self.log.log(f"traceback: {text}", kind="stderr"))
            self.post(lambda: dialogs.exception_dialog(self, str(exc), text))

        sys.excepthook = hook

        def thread_hook(args: threading.ExceptHookArgs) -> None:
            text = "".join(traceback.format_exception(
                args.exc_type, args.exc_value, args.exc_traceback))
            self.post(lambda: self.log.log(f"traceback(thread): {text}", kind="stderr"))

        threading.excepthook = thread_hook

    def report_callback_exception(self, exc: Any, val: Any = None,
                                  tb: Any = None) -> None:  # хук tk.Tk
        """Перехват исключений внутри Tk-колбэков (не молчим, п. 12).

        Сигнатура Tk менялась: в новых версиях приходит один экземпляр,
        поэтому форматируем защитливо от любого количества аргументов.
        """
        args = [a for a in (exc, val, tb) if a is not None]
        try:
            if isinstance(exc, BaseException):
                text = "".join(traceback.format_exception(
                    type(exc), exc, tb if tb is not None else exc.__traceback__))
            else:
                text = "".join(traceback.format_exception(*args))
        except TypeError:
            text = str(exc)
        self.log.log(f"traceback: {text}", kind="stderr")
        dialogs.exception_dialog(self, str(val or exc), text)

    def _on_close(self) -> None:
        self.save_state()
        if self._session_fh is not None:
            try:
                self._session_fh.flush()
                self._session_fh.close()
            except OSError:
                pass
        self._rotate_logs()
        self.destroy()


UI_QUEUE_POLL_MS = 60  # частота вычитки потока-моста в главном потоке


def run_gui(dirs: Any) -> None:
    """Точка входа GUI (вызывается из app/main.py после проверки --internal-run)."""
    app = GostApp(dirs)
    app.mainloop()
