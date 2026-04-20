# SEOParser Viki

`SEOParser Viki` - Windows-приложение на Python для работы с XMLRiver и SERPRiver. Программа предоставляет графический интерфейс для проверки позиций в Google и Yandex, получения данных Wordstat и анализа домена в выдаче SERP.

## Возможности

- Проверка позиций через XMLRiver для Google и Yandex.
- Получение данных Wordstat через XMLRiver Wordstat New.
- Проверка наличия домена в выдаче через SERPRiver.
- Импорт запросов из `CSV` и `XLSX`.
- Экспорт результатов в `CSV` и `XLSX`.
- Работа со справочниками регионов, языков и доменов из `data/references`.
- Сохранение настроек API и лимитов в локальный файл `.env`.
- Serpriver документация API - https://serpriver.ru/docs-api/
- XMLRIVER документация API - https://xmlriver.com/apidoc/

## Стек и зависимости

Основные runtime-зависимости перечислены в [requirements.txt](requirements.txt):

- `aiohttp`
- `loguru`
- `openpyxl`
- `pandas`
- `pydantic`
- `pydantic-settings`
- `python-dotenv`
- `tenacity`
- `ttkbootstrap`

Для сборки portable-версии используется отдельный файл [requirements-build.txt](requirements-build.txt), который добавляет `PyInstaller`.

## Требования

- Windows 10/11.
- Python 3.12+ для запуска из исходников.
- Доступ к API XMLRiver и SERPRiver.
- Локальные справочники в каталоге `data/references`.

## Структура проекта

```text
SEOParser_Viki/
  app/
    services/
    utils/
  data/
    references/
      countries.xlsx
      domains.xlsx
      geo.csv
      langs.xlsx
      yandex_geo.csv
      .gitkeep
  .env.example
  LICENSE
  main.py
  README.md
  requirements.txt
  requirements-build.txt
  run.cmd
  run.ps1
```

Примечания:

- Файл `.env` не хранится в репозитории и создаётся локально.
- Каталог `data/references` должен присутствовать и рядом с исходниками, и рядом с portable-сборкой.
- Файл `data/references/lr.csv` сейчас не используется кодом и не обязателен для работы приложения.

## Настройка окружения

1. Создайте виртуальное окружение:

```powershell
python -m venv .venv
```

2. Активируйте его:

```powershell
.venv\Scripts\Activate.ps1
```

3. Установите зависимости:

```powershell
pip install -r requirements.txt
```

4. Создайте файл `.env` на основе шаблона:

```powershell
Copy-Item .env.example .env
```

## Запуск приложения

Из исходников:

```powershell
python main.py
```

Или через вспомогательные скрипты, если используется локальная `.venv`:

```powershell
.\run.ps1
```

```cmd
run.cmd
```

## Использование

1. Запустите приложение.
2. Введите API-ключи XMLRiver и SERPRiver или сохраните их в `.env`.
3. Выберите нужный раздел интерфейса:
   - XMLRiver для Google, Yandex и Wordstat.
   - SERPRiver для поиска домена в выдаче.
4. Введите запросы вручную или импортируйте их из `CSV/XLSX`.
5. Запустите обработку и при необходимости экспортируйте результаты.

## Форматы ввода и вывода

Вход:

- `CSV`: используется первый столбец, разделитель определяется автоматически.
- `XLSX`: используется первый лист и первый столбец.

Выход:

- `CSV` в `UTF-8 with BOM`, удобный для Excel.
- `XLSX` через `openpyxl`.

## Справочники

Приложение ожидает следующие файлы в `data/references`:

- `geo.csv` - локации Google.
- `countries.xlsx` - страны Google.
- `langs.xlsx` - языки Google.
- `domains.xlsx` - домены Google.
- `yandex_geo.csv` - регионы Yandex.

При отсутствии этих файлов приложение запустится, но часть функций будет заблокирована, а пользователь увидит сообщения об ошибках загрузки справочников.

## Логи и локальные файлы

- Лог приложения сохраняется в `parser.log` рядом с исполняемым файлом или рядом с исходниками.
- Настройки сохраняются в локальный `.env` рядом с приложением.
- Build-артефакты `build/` и `dist/` не должны коммититься в репозиторий.

## Portable `.exe`

Для Windows рекомендуется собирать приложение в режиме `PyInstaller onedir`. Этот режим соответствует текущей архитектуре проекта: приложение читает `.env`, записывает `parser.log` и ожидает каталог `data/references` рядом с `.exe`.

### Подготовка к сборке

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install -r requirements-build.txt
```

### Сборка

Используйте один из скриптов:

```powershell
.\build_exe.ps1
```

```cmd
build_exe.cmd
```

После сборки ожидаемая структура дистрибутива:

```text
dist/
  SEOParser_Viki/
    SEOParser_Viki.exe
    _internal/
    data/
      references/
        countries.xlsx
        domains.xlsx
        geo.csv
        langs.xlsx
        yandex_geo.csv
    .env.example
```

Что нужно сделать перед передачей пользователю:

- Скопировать `.env.example` в `.env`.
- Заполнить API-ключи.
- Убедиться, что каталог `data/references` лежит рядом с `SEOParser_Viki.exe`.


## Типовые проблемы

- `Python not found` при запуске `run.ps1` или `run.cmd`.
  Значит, не создано локальное окружение `.venv` или в нём не установлен Python.
- Ошибки импорта `CSV/XLSX`.
  Проверьте, что в файле есть данные в первом столбце и не превышен `IMPORT_ROW_LIMIT`.
- Пустые списки регионов и локаций.
  Проверьте наличие файлов в `data/references`.
- Ошибки API.
  Проверьте корректность `XMLRIVER_USER_ID`, `XMLRIVER_API_KEY` и `SERPRIVER_API_KEY`.

## Лицензия

Проект распространяется по лицензии MIT. См. [LICENSE](LICENSE).
