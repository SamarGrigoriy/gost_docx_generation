# -*- mode: python ; coding: utf-8 -*-
# Два EXE из одного кода (ТЗ §14.3, §19.2): GUI без консоли + CLI с консолью.
# Причина (см. ТЗ §19.2): --windowed exe не имеет stdout, ИИ-агент не увидит JSON-вывод.
from PyInstaller.utils.hooks import copy_metadata, collect_data_files, collect_submodules

import customtkinter

# Ядро подключается к CLI динамически (`importlib.import_module("app.core." + name)` в
# app/cli_core.py), поэтому граф статических импортов PyInstaller его не видит и в onefile
# архив кладёт только app.cli_core: `gost.exe check` падал с CORE_NOT_AVAILABLE.
from app.cli_core import CORE_MODULES  # noqa: E402

datas = [
    ("assets", "assets"),
    ("presets", "presets"),
    ("profiles", "profiles"),
    ("scripts", "scripts"),
    ("schemas", "schemas"),
    ("examples", "examples"),
    ("AGENTS.md", "."),
]
# Темы CustomTkinter — типовая проблема PyInstaller: .json темы лежат внутри пакета и
# без явного сбора не попадают в onefile-архив.
datas += collect_data_files("customtkinter")
datas += copy_metadata("customtkinter")

hiddenimports = [
    "customtkinter",
    "PIL._tkinter_finder",
] + [f"app.core.{name}" for name in CORE_MODULES] + collect_submodules("customtkinter")

common = dict(
    pathex=["."],
    hiddenimports=hiddenimports,
    hookspath=[],
    datas=datas,
    noarchive=False,
)

gui = Analysis(
    ["app/main.py"],
    name="GostGenerator",
    **{k: v for k, v in common.items()},
)
pyz_gui = PYZ(gui.pure)
exe_gui = EXE(
    pyz_gui,
    gui.scripts,
    gui.binaries,
    gui.datas,
    [],
    name="GostGenerator",
    debug=False,
    strip=False,
    upx=True,
    console=False,
    icon="assets/logo.ico",
)

cli = Analysis(
    ["app/cli.py"],
    name="gost",
    **{k: v for k, v in common.items()},
)
pyz_cli = PYZ(cli.pure)
exe_cli = EXE(
    pyz_cli,
    cli.scripts,
    cli.binaries,
    cli.datas,
    [],
    name="gost",
    debug=False,
    strip=False,
    upx=True,
    console=True,
    icon=None,
)
