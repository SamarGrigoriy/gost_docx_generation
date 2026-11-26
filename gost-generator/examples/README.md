# Примеры команд CLI (по одному на команду)

Все команды печатают JSON в stdout при `--json`; логи — в stderr.
Коды возврата: 0 ок, 1 нарушения, 2 аргументы, 3 файл не найден, 4 скрипт упал,
5 таймаут, 6 конфиг, 7 нормализация (см. AGENTS.md).

```bash
# init — подготовить рабочую область + агентные артефакты
gost init --agents --dir ./work

# doctor — самодиагностика окружения
gost doctor --json

# rules list / show / validate / export / import / diff
gost rules list --json
gost rules show gost_732 --json
gost rules validate ./presets/gost_732.json --json
gost rules export gost_732 ./out/gost.json
gost rules import ./out/gost.json my_copy
gost rules diff gost_732 my_copy --json

# profile list / show / set / use / export / import
gost profile list --json
gost profile show default --json
gost profile set default work.theme="Сортировка массивов"
gost profile use default
gost profile export default ./out/profile.json
gost profile import ./examples/meta.json import_meta

# scripts list / describe / add / validate
gost scripts list --json
gost scripts describe mkdocx --json
gost scripts add ./my_plugin.py --copy
gost scripts validate mkdocx

# run — генерация (+нормализация+проверка) с параметрами из файла примера
gost run mkdocx --input examples/report.md --output output/report.docx \
     --params @examples/meta.json --rules gost_732 --normalize --check --json

# check — проверка готового файла, отчёт в markdown
gost check output/report.docx --rules gost_732 --format md --report output/report.md

# fix — автоисправление оформления (идемпотентно, бэкап .orig.docx)
gost fix output/report.docx --rules gost_732 --json

# batch — пакетная проверка папки
gost batch ./output --glob '*.docx' --check --report output/summary.md --json

# job run — декларативная задача целиком
gost job run examples/task.json --json

# schema — схемы данных
gost schema rules
gost schema job

# explain — расшифровка кода нарушения
gost explain TBL_CAPTION

# context — полная сводка для агента (первый вызов)
gost context --json
```
