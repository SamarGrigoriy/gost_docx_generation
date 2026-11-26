# ГОСТ-Генератор

Генерация и проверка учебных документов `.docx` по ГОСТ 7.32 и методичке «Мультиворк»:
графическое приложение (CustomTkinter) и командный интерфейс для человека и ИИ-агента.

Требования (ТЗ) и порядок работы над ними — в [PROMT.md](PROMT.md),
правила для агентов — в [AGENT4.md](AGENT4.md) и
[gost-generator/AGENTS.md](gost-generator/AGENTS.md), шаблон чат-бота —
в [CHATBOT_TEMPLATE.md](CHATBOT_TEMPLATE.md).

## Раскладка репозитория

| Путь | Содержимое |
|------|-----------|
| `gost-generator/app/` | ядро (правила, инспектор, нормализатор, CLI) и GUI |
| `gost-generator/scripts/` | скрипты-генераторы (плагины с `SCRIPT_META`) |
| `gost-generator/tests/` | 169 автотестов (`unittest`) |
| `gost-generator/schemas/`, `examples/` | контракты JSON для агента (ТЗ §19) |
| `gost-generator/packaging/` | запуск из исходников и пункт меню для Linux |
| `gost-generator/CHANGELOG.md`, `DECISIONS.md` | что сделано и почему решено так |

## Установка (Linux)

```sh
cd gost-generator
python3 -m venv --system-site-packages venv     # tkinter берётся из системы
./venv/bin/pip install -r requirements.txt
./venv/bin/python -m app                         # GUI
./venv/bin/python -m app.cli doctor --format json  # CLI
```

Пункт меню «ГОСТ-Генератор» (без прав суперпользователя, в `~/.local`):

```sh
make desktop            # или sh packaging/install-linux-desktop.sh
```

После установки приложение появляется в меню «Офис» и запускается командой
`~/.local/bin/gost-generator`.

## Установка (Windows)

Скачайте `GostGenerator-windows.zip` из последнего релиза или из артефакта workflow
[`build-windows`](../../actions/workflows/build-windows.yml), распакуйте в любую папку и
запустите `GostGenerator.exe` (GUI) или `gost.exe` (CLI). Python устанавливать не нужно:
архив portable — маркер `portable.flag` держит данные (`scripts/`, `presets/`, `profiles/`,
`output/`) рядом с программой.

## Сборка `.exe` под Windows

PyInstaller собирает Windows-исполняемые файлы только на Windows (локальная машина — Linux,
поэтому сборка закреплена за CI, см. `DECISIONS.md` D36).

1. В GitHub Actions: вкладка **Actions → build-windows → Run workflow** (`workflow_dispatch`).
   Результат — артефакт `GostGenerator-windows.zip`.
2. На теге `v*` тот же workflow дополнительно выкладывает архив в релиз.
3. На машине с Windows локально:

   ```powershell
   cd gost-generator
   py -3.11 -m venv venv
   .\venv\Scripts\Activate.ps1
   pip install -r requirements-dev.txt
   pyinstaller build.spec --noconfirm
   # dist\GostGenerator.exe (GUI) и dist\gost.exe (CLI)
   ```

CI-прогон проверяет собранные `.exe` без установленного Python: `gost.exe doctor --json`
(конверт §19.4), `gost.exe check … --fail-on error` (коды возврата §19.5) и запуск GUI-окна.

## Проверки

```sh
cd gost-generator
./venv/bin/python -m unittest discover -s tests -v   # 169 тестов
make lint                                            # compileall
```

GUI-тесты пропускаются, если дисплея нет; на Windows-раннере они выполняются полностью.
