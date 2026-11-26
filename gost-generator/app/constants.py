"""Общие константы приложения: версии, пути по умолчанию, значения правил по умолчанию.

Всё, что ранее было магическими числами в UI и ядре, собрано здесь, чтобы
критерии приёмки (агенты читают только AGENTS.md) оставались проверяемыми.
"""

from __future__ import annotations

APP_NAME = "GostGenerator"
APP_TITLE = "ГОСТ-Генератор"
APP_VERSION = "1.0.0"
SCHEMA_VERSION = 1

# --- Каталоги данных (относительные имена внутри GOST_HOME) ---
DIR_SCRIPTS = "scripts"
DIR_PRESETS = "presets"
DIR_PROFILES = "profiles"
DIR_OUTPUT = "output"
DIR_LOGS = "logs"
DIR_SCHEMAS = "schemas"
DIR_EXAMPLES = "examples"
PORTABLE_FLAG = "portable.flag"
STATE_FILE = "state.json"
PROJECT_CONFIG_FILE = "gost.json"

CONFIG_DIR_NAME = "GostGenerator"

# --- UI ---
WINDOW_WIDTH = 1200
WINDOW_HEIGHT = 800
WINDOW_MIN_WIDTH = 1000
WINDOW_MIN_HEIGHT = 680
LOG_PANEL_HEIGHT = 200
UI_SCALES = (100, 110, 125)
UI_THEMES = ("system", "dark", "light")
DEFAULT_THEME = "system"
DEFAULT_UI_SCALE = 100

# --- Раннер ---
DEFAULT_SCRIPT_TIMEOUT_S = 120
KILL_GRACE_PERIOD_S = 3
MAX_RECENT_FILES = 10
RECENT_COMMANDS = 20

# --- Логи сессий ---
LOG_FILE_PATTERN = "logs/{stamp}.log"
LOG_TIMESTAMP_FORMAT = "%Y-%m-%d_%H%M%S"
LOG_KEEP_COUNT = 30

# --- Правила по умолчанию (совпадают с presets/gost_732.json) ---
DEFAULT_PRESET_KEY = "gost_732"
DEFAULT_PROFILE_KEY = "default"
DEFAULT_FONT = "Times New Roman"
DEFAULT_SIZE_PT = 14
DEFAULT_LINE_SPACING = 1.5
DEFAULT_INDENT_CM = 1.25
DEFAULT_ALIGN = "justify"
MARGIN_TOLERANCE_MM = 0.5
INDENT_TOLERANCE_TWIPS = 5

FONT_CHOICES = ("Times New Roman", "Arial", "PT Sans", "Calibri")
LINE_SPACING_CHOICES = (1.0, 1.15, 1.5, 2.0)
ALIGNMENTS = ("left", "center", "right", "justify")
ORIENTATIONS = ("portrait", "landscape")
PAGE_SIZES = ("A4",)

# --- Проверка (docx_inspector) ---
SEVERITIES = ("error", "warning", "info")
SEVERITY_ORDER = {"error": 0, "warning": 1, "info": 2}
MONOSPACE_FONTS = ("Consolas", "Courier New", "Courier", "Liberation Mono")

# Структурные элементы, оформляемые ПРОПИСНЫМИ с новой страницы (п. 5.2 методички).
STRUCTURAL_HEADINGS = (
    "РЕФЕРАТ",
    "СОДЕРЖАНИЕ",
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "СПИСОК ИСТОЧНИКОВ",
    "ПРИЛОЖЕНИЕ",
)

# Буквы кириллицы для приложений: Ё, З, Й, О, Ч, Ъ, Ы, Ь исключены (ГОСТ 7.32, п. 2.14).
APPENDIX_LETTERS = tuple(
    ch for ch in "АБВГДЕЖЗИКЛМНОПРСТУФХЦЧШЩЭЮЯ" if ch not in "ЁЗЙОЧЪЫЬ"
)

# --- Профиль ---
WORK_TYPES = (
    "Реферат",
    "Отчёт по практическому занятию",
    "Лабораторная работа",
    "Курсовая работа",
    "ВКР",
    "Контрольная",
    "Технологическая карта",
    "Конспект занятия",
    "Подборка игр",
    "Свой вариант",
)
STATUS_CHOICES = ("студент", "преподаватель", "методист")
DEFAULT_MINISTRY = "МИНИСТЕРСТВО НАУКИ И ВЫСШЕГО ОБРАЗОВАНИЯ РОССИЙСКОЙ ФЕДЕРАЦИИ"
DEFAULT_FSBE = (
    "ФЕДЕРАЛЬНОЕ ГОСУДАРСТВЕННОЕ БЮДЖЕТНОЕ ОБРАЗОВАТЕЛЬНОЕ УЧРЕЖДЕНИЕ ВЫСШЕГО ОБРАЗОВАНИЯ"
)

# --- Плагины ---
INPUT_TYPES = ("none", "file_md", "file_docx", "text", "folder")
RULES_HOOKS = ("native", "constants", "postprocess", "none")
DEFAULT_RULES_HOOK = "postprocess"
PARAM_TYPES = ("text", "textarea", "int", "float", "bool", "select", "file", "folder", "date")
SIDECAR_SUFFIX = ".meta.json"

# --- Выход отчётов ---
REPORT_MD_EXT = ".md"
REPORT_TXT_EXT = ".txt"

# --- Переменные окружения CLI (п. 19.6) ---
ENV_HOME = "GOST_HOME"
ENV_RULES = "GOST_RULES"
ENV_PROFILE = "GOST_PROFILE"
ENV_OUTPUT = "GOST_OUTPUT"

# --- Коды возврата CLI (п. 19.5) ---
EXIT_OK = 0
EXIT_VIOLATIONS = 1
EXIT_USAGE = 2
EXIT_NOT_FOUND = 3
EXIT_SCRIPT_FAILED = 4
EXIT_TIMEOUT = 5
EXIT_BAD_CONFIG = 6
EXIT_NORMALIZE_FAILED = 7
