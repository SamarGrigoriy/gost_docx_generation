"""Дымовые проверки GUI (п. 15.1): окно собирается, вкладки открываются, колбэки работают.

Пропускаются без X-дисплея. Класс регрессий, который они ловят: `app.safe()`
не пробрасывал аргументы, а вкладка «Генерация» зовёт рабочий поток через
`Thread(target=self.app.safe(self._worker), args=(...))` — из-за этого кнопка
«Сгенерировать» навечно оставалась в состоянии «занято».
"""

from __future__ import annotations

import tempfile
import unittest


def _make_app(home: str):
    """Собрать GostApp на временном каталоге данных; None, если GUI недоступен."""
    try:
        import customtkinter as ctk  # noqa: F401

        from app.core import config
        from app.ui import app_window
    except Exception as exc:  # pragma: no cover - окружение без GUI
        raise unittest.SkipTest(f"GUI недоступен: {exc}")
    try:
        return app_window.GostApp(config.DataDirs(override=home))
    except Exception as exc:
        raise unittest.SkipTest(f"Нет дисплея для Tk: {exc}")


def _pump(win, rounds: int = 20) -> None:
    """Прокрутить события: журнал и UI-очередь обновляются только в главном цикле."""
    for _ in range(rounds):
        win.update()


def _drain(win, seconds: float = 1.5) -> None:
    """То же во времени: строки журнала приходят через after-очередь с шагом 60 мс,
    без пауз между update() они не успевают вставиться в текстбокс."""
    import time

    end = time.time() + seconds
    while time.time() < end:
        win.update()
        time.sleep(0.02)


class UiSmokeTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls._tmp = tempfile.TemporaryDirectory()
        cls.win = _make_app(cls._tmp.name)

    @classmethod
    def tearDownClass(cls) -> None:
        cls.win.destroy()
        cls._tmp.cleanup()

    def test_all_tabs_open(self) -> None:
        from app.ui import app_window

        for name in app_window.TAB_ORDER:
            self.win.show_tab(name)
            self.assertEqual(self.win.active_tab, name)

    def test_safe_forwards_arguments(self) -> None:
        received: list[tuple] = []
        self.win.safe(lambda *a, **k: received.append((a, k)))("info", key=1)
        self.assertEqual(received, [(("info",), {"key": 1})])

    def test_safe_reports_errors_without_raising(self) -> None:
        def boom() -> None:
            raise ValueError("нет данных")

        self.win.safe(boom)()  # не должно бросать в вызывающий код
        _pump(self.win)
        self.assertIn("нет данных", self.win.log.get_text())

    def test_generator_select_enables_run(self) -> None:
        tab = self.win.tabs["generator"]
        info = next((p for p in self.win.plugins if "make_games" in (p.file_name or "")), None)
        self.assertIsNotNone(info, "штатный make_games должен быть в списке скриптов")
        tab.select(info)
        self.assertEqual(tab.validate(), [])
        self.assertEqual(str(tab.run_btn.cget("state")), "normal")

    def test_output_name_follows_theme_param(self) -> None:
        tab = self.win.tabs["generator"]
        info = next((p for p in self.win.plugins if "make_games" in (p.file_name or "")), None)
        tab.select(info)
        tab._param_fields["theme"].set("Математическая мозаика")
        tab._on_param_change()
        self.assertIn("Математическая мозаика", tab.out_row.get())


    def test_sidebar_bottom_buttons_are_not_clipped(self) -> None:
        """Регрессия приёмки §15.1: на экране 768 px журнал отнимал место у «Тема».

        Кнопки «Настройки/О программе/Тема» прижаты к низу сайдбара; если высота
        строки меньше запрошенной, Tk их не рисует вовсе — молча, без ошибки.
        """
        _pump(self.win)
        sidebar = self.win._tab_btns["generator"].master
        btn_bottom = self.win._theme_btn.winfo_rooty() + self.win._theme_btn.winfo_height()
        self.assertEqual(str(self.win._theme_btn.winfo_ismapped()), "1", "кнопка темы не видна")
        self.assertLessEqual(btn_bottom, sidebar.winfo_rooty() + sidebar.winfo_height(),
                             "нижние кнопки сайдбара не влезают в окно")


    def test_mouse_wheel_over_foreign_widget_does_not_crash(self) -> None:
        """CustomTkinter 6.0 на CPython 3.14 отдаёт `event.widget` строкой — патч `app.ui`."""
        import types

        import customtkinter as ctk

        frame = ctk.CTkScrollableFrame(ctk.CTkToplevel(self.win))
        frame.pack()
        _pump(self.win, 3)
        foreign = types.SimpleNamespace(widget="!toplevel!frame!button", num=4, delta=-120)
        frame._mouse_wheel_all(foreign)  # не должно бросать AttributeError
        own = types.SimpleNamespace(widget=str(frame._parent_canvas) + "!inner", num=4, delta=-120)
        self.assertTrue(frame._check_if_valid_scroll(own.widget))
        frame.master.destroy()

    def test_add_script_copies_into_scripts_dir(self) -> None:
        """«Добавить скрипт…»: подтверждение → копирование в каталог скриптов → список обновлён."""
        import os
        import tempfile

        from app.ui import dialogs

        tab = self.win.tabs["generator"]
        src_dir = tempfile.mkdtemp()
        src = os.path.join(src_dir, "qa_added.py")
        with open(src, "w", encoding="utf-8") as fh:
            fh.write("print('qa')\n")
        picked: list = []
        originals = (dialogs.pick_file, dialogs.warn_external_script, dialogs.confirm)
        dialogs.pick_file = lambda *a, **k: src
        dialogs.warn_external_script = lambda *a, **k: True
        dialogs.confirm = lambda *a, **k: True
        try:
            tab.add_script()
            _pump(self.win, 5)
            copied = os.path.join(str(self.win.scripts_dir()), "qa_added.py")
            picked.append(copied)
            self.assertTrue(os.path.isfile(copied), "скрипт не скопирован в каталог скриптов")
            self.assertIn("qa_added.py", {p.file_name for p in self.win.plugins})
        finally:
            (dialogs.pick_file, dialogs.warn_external_script,
             dialogs.confirm) = originals
            for path in picked:
                if os.path.isfile(path):
                    os.remove(path)
            import shutil

            shutil.rmtree(src_dir, ignore_errors=True)
        tab.refresh_scripts(notify=False)

    def test_open_scripts_dir_reports_outcome_in_journal(self) -> None:
        """Кнопка «Открыть папку скриптов» отчитывается в журнале в обоих исходах.

        Без строки в журнале клик выглядит мёртвым: просмотрщик мог переиспользовать своё
        окно (нового не появляется) или отказаться открывать каталог — пользователю нечего
        читать на экране. Подменяем `xdg-open` на скрипт с нужным кодом возврата.
        """
        import os
        import shutil
        import sys
        import tempfile

        from app.strings import t
        from app.ui import dialogs

        if not sys.platform.startswith("linux"):
            self.skipTest("механизм xdg-open проверяется на Linux")
        tab = self.win.tabs["generator"]
        scripts = str(self.win.scripts_dir())
        old_path = os.environ["PATH"]
        dirs: list[str] = []

        def fake_handler(rc: int, note: str) -> str:
            d = tempfile.mkdtemp()
            dirs.append(d)
            script = os.path.join(d, "xdg-open")
            with open(script, "w", encoding="utf-8") as fh:
                fh.write(f"#!/bin/sh\nprintf '%s' '{note}' >&2\nexit {rc}\n")
            os.chmod(script, 0o755)
            return d

        try:
            os.environ["PATH"] = fake_handler(1, "нет метода") + os.pathsep + old_path
            self.win.log.clear()
            tab.open_scripts_dir()
            _drain(self.win)
            text = self.win.log.get_text()
            self.assertIn(t("open_error", path=scripts,
                            why=t("open_failed", code=1, why="нет метода")), text)

            os.environ["PATH"] = fake_handler(0, "") + os.pathsep + old_path
            self.win.log.clear()
            tab.open_scripts_dir()
            _drain(self.win)
            self.assertIn(t("open_done", path=scripts), self.win.log.get_text())
        finally:
            os.environ["PATH"] = old_path
            for d in dirs:
                shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
