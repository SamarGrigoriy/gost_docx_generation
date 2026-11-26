#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""make_fixtures.py — генерация тестовых фикстур (скрипт, не тест).

Запуск из корня репозитория:
    python3 tests/fixtures/make_fixtures.py

Создаёт в tests/fixtures/:
  пример.example.md          — эталонный markdown-вход для mkdocx
  diag.png                   — минимальная PNG-картинка для теста встраивания
  fixtures_good.docx         — mkdocx по эталонному md (ожидается 0 ошибок)
  fixtures_broken.docx       — тот же файл с deliberate-дефектами (мини-редактор OOXML
                               допустим ТОЛЬКО для тестовой фикстуры):
                               Arial вместо TNR, поля 25 мм, отступ 0,5 см,
                               интервал 1,0, номер страницы сверху,
                               подпись таблицы по центру
  fixtures_karta.docx        — make_karta с демонстрационной метой
  fixtures_games.docx        — make_games с демонстрационной метой
"""

from __future__ import annotations

import base64
import os
import subprocess
import sys
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
ROOT = os.path.abspath(os.path.join(HERE, "..", ".."))
SCRIPTS = os.path.join(ROOT, "scripts")
PRESET = os.path.join(ROOT, "presets", "gost_732.json")
PEDAGOGIC = os.path.join(ROOT, "presets", "pedagogical.json")

# Эталонный markdown: заголовки всех уровней, списки, таблица с подписью,
# код-блок, рисунок, гиперссылка, список источников со ссылками в тексте.
# СОДЕРЖАНИЕ намеренно нет: mkdocx сам вставляет страницу с полем TOC
# (стрим тест проверяет именно автогенерацию содержания).
EXAMPLE_MD = """# РЕФЕРАТ

Краткое изложение работы по теме «Тестовая тема» приведено в таблице 1 и на рисунке 1.

## 1. Теоретические основы

Текст раздела с **выделением**, `кодом` и ссылкой <https://example.ru/doc?A=1&b=2>.

- первый пункт списка;
- второй пункт списка.

Таблица 1 – Характеристики системы

| Параметр | Значение | Ед. |
|---|---|---|
| Ширина | 15 | см |
| Масса | 2 & 3 | кг |

Данные приведены в таблице 1 [1, с. 25].

### 1.1 Подраздел

Текст подраздела, продолжение [2, с. 10].

```python
def hello():
    return "мир"
```

# 2. Практическая часть

![Схема взаимодействия](diag.png)

Структура показана в соответствии с рисунком 1.

# ЗАКЛЮЧЕНИЕ

Работа выполнена, цели достигнуты.

# СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ

1. Иванов И. И. Основы документооборота. Москва : Наука, 2024. 320 с.
2. Петров П. П. Автоматизация отчётности. Санкт-Петербург : Питер, 2023. 180 с.
"""

META_GOOD = ('{"fio":"Иванов Иван Иванович","gruppa":"ИТ-21",'
             '"prep":"Петров П. П.","status":"студент",'
             '"work_type":"Отчёт по практическому занятию","pz_no":"1",'
             '"theme":"Тестовая тема","disc":"Информатика","city":"Геническ",'
             '"univ":"ГБУ ВО «Керченский военно-морской технологический университет»",'
             '"kaf":"Кафедра информационных технологий"}')

# 1×1 px PNG — минимальный валидный растр, чтобы не держать бинарь в репо
PNG_B64 = ("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")


def run(cmd: list[str]) -> None:
    """Запуск дочернего скрипта списком аргументов (без shell), диагностика — в лог."""
    print("$", " ".join(cmd))
    proc = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                          errors="replace", timeout=120)
    if proc.returncode != 0:
        sys.stderr.write(proc.stdout + proc.stderr)
        raise SystemExit(f"команда завершилась кодом {proc.returncode}")


def make_broken(src: str, dst: str) -> None:
    """Мини-редактор OOXML: вносит 6 фикстурированных дефектов (только для теста!)."""
    with zipfile.ZipFile(src) as z:
        items = {n: z.read(n) for n in z.namelist()}
    doc = items["word/document.xml"].decode("utf-8")
    sty = items["word/styles.xml"].decode("utf-8")
    # дефект 1: Arial вместо Times New Roman (фактические rFonts)
    doc = doc.replace('w:ascii="Times New Roman"', 'w:ascii="Arial"')
    doc = doc.replace('w:hAnsi="Times New Roman"', 'w:hAnsi="Arial"')
    sty = sty.replace('w:ascii="Times New Roman"', 'w:ascii="Arial"')
    sty = sty.replace('w:hAnsi="Times New Roman"', 'w:hAnsi="Arial"')
    # дефект 2: поля 25 мм вместо 30/15
    doc = doc.replace('w:left="1701"', 'w:left="1418"').replace('w:right="850"', 'w:right="1418"')
    # дефект 3: интервал 1,0 вместо 1,5 в унаследованных свойствах
    sty = sty.replace('w:line="360"', 'w:line="240"')
    doc = doc.replace('w:line="360"', 'w:line="240"')
    # дефект 4: номер страницы — в верхнем колонтитуле вместо нижнего
    doc, items = move_page_number_to_header(doc, items)
    items["word/document.xml"] = doc.encode("utf-8")
    # дефект 5: отступ 0,5 см в унаследованном стиле Normal
    sty = sty.replace('w:firstLine="709"', 'w:firstLine="284"')
    items["word/styles.xml"] = sty.encode("utf-8")
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in items.items():
            z.writestr(n, d)


def move_page_number_to_header(doc: str, items: dict[str, bytes]) -> tuple[str, dict[str, bytes]]:
    """Перенести поле PAGE из footer1.xml в новый header1.xml (дефект NUM_POSITION)."""
    ftr = items.get("word/footer1.xml", b"").decode("utf-8")
    if "<w:instrText" not in ftr and "fldChar" not in ftr:
        raise SystemExit("в footer1.xml нет поля PAGE — дефект номера сверху не внести")
    header = ftr.replace("<w:ftr ", "<w:hdr ", 1).replace("</w:ftr>", "</w:hdr>")
    items["word/header1.xml"] = header.encode("utf-8")
    for name in ("word/footer1.xml", "word/footer2.xml"):
        xml = items.get(name, b"").decode("utf-8")
        items[name] = _drop_page_runs(xml).encode("utf-8")
    ct = items["[Content_Types].xml"].decode("utf-8")
    ct = ct.replace(
        '<Override PartName="/word/footer1.xml"',
        '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats'
        '-officedocument.wordprocessingml.header+xml"/><Override PartName="/word/footer1.xml"', 1)
    items["[Content_Types].xml"] = ct.encode("utf-8")
    rels = items["word/_rels/document.xml.rels"].decode("utf-8")
    items["word/_rels/document.xml.rels"] = rels.replace(
        '<Relationship Id="rIdFtrD"',
        '<Relationship Id="rIdHdrD" Type="http://schemas.openxmlformats.org/officeDocument'
        '/2006/relationships/header" Target="header1.xml"/>'
        '<Relationship Id="rIdFtrD"', 1).encode("utf-8")
    if "<w:headerReference" not in doc:
        doc = doc.replace("<w:footerReference",
                          '<w:headerReference w:type="default" r:id="rIdHdrD"/>'
                          "<w:footerReference", 1)
    return doc, items


def _drop_page_runs(xml: str) -> str:
    """Удалить из колонтитула run'ы поля (fldChar/instrText) — номер переехал наверх."""
    out = []
    pos = 0
    while True:
        start = xml.find("<w:r>", pos)
        if start < 0:
            out.append(xml[pos:])
            break
        end = xml.find("</w:r>", start)
        end = len(xml) if end < 0 else end + len("</w:r>")
        chunk = xml[start:end]
        if "fldChar" not in chunk and "instrText" not in chunk:
            out.append(xml[pos:start])
            out.append(chunk)
        else:
            out.append(xml[pos:start])
        pos = end
    return "".join(out)


def center_table_caption(path: str) -> None:
    """Правка фикстуры (дефект 6): jc у абзаца «Таблица 1 –» → center."""
    with zipfile.ZipFile(path) as z:
        items = {n: z.read(n) for n in z.namelist()}
    doc = items["word/document.xml"].decode("utf-8")
    marker = "Таблица 1 – Характеристики системы"
    idx = doc.find(marker)
    if idx < 0:
        raise SystemExit("в фикстуре не найдена подпись таблицы — правка не внесена")
    start = doc.rfind("<w:p>", 0, idx)
    end = doc.find("</w:p>", idx) + len("</w:p>")
    para_xml = doc[start:end].replace('<w:jc w:val="left"/>', '<w:jc w:val="center"/>')
    doc = doc[:start] + para_xml + doc[end:]
    items["word/document.xml"] = doc.encode("utf-8")
    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        for n, d in items.items():
            z.writestr(n, d)


def main() -> int:
    os.makedirs(HERE, exist_ok=True)
    with open(os.path.join(HERE, "пример.example.md"), "w", encoding="utf-8") as fh:
        fh.write(EXAMPLE_MD)
    with open(os.path.join(HERE, "diag.png"), "wb") as fh:
        fh.write(base64.b64decode(PNG_B64))

    py = sys.executable
    good = os.path.join(HERE, "fixtures_good.docx")
    broken = os.path.join(HERE, "fixtures_broken.docx")

    run([py, "-I", os.path.join(SCRIPTS, "mkdocx.py"),
         os.path.join(HERE, "пример.example.md"), good, META_GOOD,
         "--rules", PRESET, "--no-verify"])
    make_broken(good, broken)
    center_table_caption(broken)

    run([py, "-I", os.path.join(SCRIPTS, "make_karta.py"),
         "--out", os.path.join(HERE, "fixtures_karta.docx"),
         "--rules", PEDAGOGIC, "--no-verify",
         "--meta", '{"theme":"Моя семья","age_group":"4-5","duration_min":20,'
                   '"goal":"Формирование представлений о семье",'
                   '"tasks":"Об.: называть членов семьи.\\nРазв.: связная речь.",'
                   '"materials":"Семейные фотографии"}'])
    run([py, "-I", os.path.join(SCRIPTS, "make_games.py"),
         "--out", os.path.join(HERE, "fixtures_games.docx"),
         "--rules", PEDAGOGIC, "--no-verify",
         "--meta", '{"theme":"Осень","age_group":"5-6",'
                   '"games":"[\\"Листопад|Развитие слуха и движений|'
                   'Дети бегут под бубен|Не задевать друг друга\\",'
                   '\\\"Тише едешь|Быстрая реакция|Ведущий ловит|Стоять на месте\\"]",'
                   '"intro":"Дидактические игры актуальны в старшей группе."}'])

    # валидность: каждый .docx — читаемый zip с парсящимся document.xml
    from xml.etree import ElementTree as ET
    for name in ("fixtures_good.docx", "fixtures_broken.docx",
                 "fixtures_karta.docx", "fixtures_games.docx"):
        p = os.path.join(HERE, name)
        with zipfile.ZipFile(p) as z:
            ET.fromstring(z.read("word/document.xml"))
        print(f"OK zip+xml: {name} ({os.path.getsize(p)} bytes)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
