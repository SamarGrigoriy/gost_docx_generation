"""UI-слой. Импорт пакета ставит совместимость-патчи над сторонними виджетами.

Патчи нужны до сборки любого окна: `app.ui.*` импортируются из `app.ui.app_window`,
а пакет инициализируется первым.
"""

from __future__ import annotations


def _patch_ctk_scrollable_wheel() -> None:
    """CustomTkinter 6.0 + CPython 3.14: `event.widget` бывает строкой-путём Tk.

    `_mouse_wheel_all` привязан через `bind_all`, то есть слышит колесо мыши во всём
    приложении — включая нативный диалог выбора файла, где у виджетов нет python-обёртки.
    Штатная проверка лезет в `widget.master` и падает с AttributeError на `str`.
    """
    try:
        import customtkinter as ctk
    except Exception:
        return  # GUI-пакет недоступен (headless) — патч не нужен
    cls = getattr(ctk, "CTkScrollableFrame", None)
    if cls is None or getattr(cls, "_gost_wheel_patched", False):
        return
    original = cls._check_if_valid_scroll

    def _check_if_valid_scroll(self, widget):  # noqa: ANN001 - метод класса CTk
        if isinstance(widget, str):
            path = str(self._parent_canvas)
            return widget == path or widget.startswith(path + "!")
        return original(self, widget)

    cls._check_if_valid_scroll = _check_if_valid_scroll
    cls._gost_wheel_patched = True


_patch_ctk_scrollable_wheel()
