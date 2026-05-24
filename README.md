# SEOParser Viki

`SEOParser Viki` - Windows-приложение с графическим интерфейсом для SEO-задач через XMLRiver, SERPRiver и Букварикс.

## Возможности

- Проверка позиций в Google и Yandex через XMLRiver.
- Пакетная обработка XMLRiver SERP-запросов по 10 ключей за раз по умолчанию. Если передать 100 ключей, программа отправит 10 последовательных пачек.
- Трекер позиций домена в выдаче Yandex XMLRiver.
- Сбор фраз по домену и проверка позиций собранных фраз.
- Получение Wordstat-данных через XMLRiver Wordstat New.
- Проверка домена через SERPRiver.
- Получение данных Букварикс по словам и доменам.
- Импорт запросов из `CSV` и `XLSX`.
- Экспорт результатов в `CSV` и `XLSX`.
- Справочники регионов, стран, языков и доменов из `data/references`.
- Сохранение API-ключей и настроек в локальный файл `.env`.

Документация API:

- XMLRiver: <https://xmlriver.com/apidoc/>
- SERPRiver: <https://serpriver.ru/docs-api/>

## Готовая portable-версия для пользователя

Это основной вариант для обычного пользователя.

1. Скачайте архив готовой portable-сборки.
2. Распакуйте архив в любую папку, например `C:\Apps\SEOParser_Viki`.
3. Запустите `SEOParser_Viki.exe`.
4. Введите API-ключи во вкладках программы и нажмите `Сохранить`.

Portable-сборка уже содержит Python и runtime-библиотеки внутри папки приложения. Пользователю не нужно устанавливать Python, создавать `.venv` или ставить зависимости через `pip`.

В portable-папке важны следующие файлы и каталоги:

- `SEOParser_Viki.exe` - запуск приложения.
- `.env` - локальные API-ключи и настройки. Создаётся программой или вручную из `.env.example`.
- `.env.example` - шаблон настроек.
- `parser.log` - лог работы приложения.
- `data/references` - справочники для регионов, стран, языков и доменов.
- `VERSION.txt` - версия portable-сборки.
- `UPDATE.md`, `update_portable.cmd`, `update_portable.ps1` - обновление старой portable-папки.

## Обновление portable-версии

Для обновления старой portable-папки используйте файлы из новой сборки:

1. Закройте старую программу.
2. Распакуйте новую portable-сборку в отдельную временную папку.
3. Запустите из новой папки:

```cmd
update_portable.cmd "C:\Path\To\Old\SEOParser_Viki"
```

Скрипт обновит приложение и внутренние runtime-файлы, но сохранит пользовательские `.env`, `parser.log` и существующую папку `data`.

Если нужно заменить справочники из новой сборки, используйте PowerShell-вариант:

```powershell
.\update_portable.ps1 -TargetPath "C:\Path\To\Old\SEOParser_Viki" -UpdateReferences
```

Подробная инструкция находится в [UPDATE.md](UPDATE.md).

## Запуск из исходников для разработчика

Требования:

- Windows 10/11.
- Python 3.12 или новее.
- Доступ к API XMLRiver, SERPRiver или Букварикс, если нужны соответствующие функции.

Подготовка:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
Copy-Item .env.example .env
```

Запуск:

```powershell
python main.py
```

Или через вспомогательный скрипт:

```powershell
.\run.ps1
```

Для `cmd.exe`:

```cmd
run.cmd
```

Если при запуске `run.ps1` или `run.cmd` появляется `Python not found`, значит локальное окружение `.venv` не создано или зависимости установлены не в эту папку.

## Сборка portable `.exe`

Сборка выполняется через PyInstaller в режиме `onedir`.

Подготовьте окружение:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pip install -r requirements-build.txt
```

Запустите сборку:

```powershell
.\build_exe.ps1
```

Или из `cmd.exe`:

```cmd
build_exe.cmd
```

Готовая portable-папка будет создана здесь:

```text
dist/
  SEOParser_Viki/
    SEOParser_Viki.exe
    _internal/
    data/
      references/
    .env.example
    VERSION.txt
    UPDATE.md
    update_portable.cmd
    update_portable.ps1
```

Перед передачей пользователю упакуйте папку `dist/SEOParser_Viki` в архив. Пользователь должен распаковать её и запустить `SEOParser_Viki.exe`.

Файл `SEOParser_Viki.spec` является локальным артефактом PyInstaller. Он может содержать абсолютные пути конкретного компьютера, поэтому не используйте его как основную инструкцию сборки. Основной способ сборки - `build_exe.ps1` или `build_exe.cmd`.

## Настройки `.env`

Пример находится в [.env.example](.env.example).

Основные параметры:

```env
XMLRIVER_USER_ID=
XMLRIVER_API_KEY=
SERPRIVER_API_KEY=
BUKVARIX_API_KEY=free
REQUEST_CONNECT_TIMEOUT=5
REQUEST_READ_TIMEOUT=30
XMLRIVER_MAX_CONCURRENCY=10
SERPRIVER_MAX_CONCURRENCY=10
IMPORT_ROW_LIMIT=10000
```

`XMLRIVER_MAX_CONCURRENCY` управляет размером пачки XMLRiver SERP-запросов. Значение `10` означает, что программа отправляет до 10 запросов одновременно, ждёт завершения этой пачки и только потом отправляет следующие 10.

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
      lr.csv
      yandex_geo.csv
  tests/
  .env.example
  build_exe.cmd
  build_exe.ps1
  main.py
  README.md
  requirements.txt
  requirements-build.txt
  run.cmd
  run.ps1
  update_portable.cmd
  update_portable.ps1
```

`data/references` должен быть рядом с исходниками при запуске из кода и рядом с `SEOParser_Viki.exe` в portable-сборке. Если справочников нет, приложение запустится, но часть полей и функций будет недоступна.

## Импорт и экспорт

Входные файлы:

- `CSV` - используется первый столбец, разделитель определяется автоматически.
- `XLSX` - используется первый лист и первый столбец.

Выходные файлы:

- `CSV` сохраняется в `UTF-8 with BOM`, чтобы Excel корректно открывал русский текст.
- `XLSX` создаётся через `openpyxl`.

## Локальные файлы и артефакты

- `.env` хранит локальные ключи и не должен попадать в репозиторий.
- `parser.log` хранит лог приложения и не нужен в коммитах.
- `build/` и `dist/` создаются при сборке и не должны коммититься.
- `.venv/` является локальным окружением разработчика и не входит в portable-сборку.

## Проверка тестов

Проект использует стандартный `unittest`:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

`pytest` не требуется для обычной проверки, если он отдельно не установлен в окружение разработчика.

## Типовые проблемы

- `Python not found` при запуске `run.ps1` или `run.cmd`: создайте `.venv` и установите зависимости.
- Пустые списки регионов или стран: проверьте наличие файлов в `data/references`.
- Ошибки API: проверьте `XMLRIVER_USER_ID`, `XMLRIVER_API_KEY`, `SERPRIVER_API_KEY` и баланс сервиса.
- Ошибка занятости поисковых ботов XMLRiver при большом списке ключей: проверьте, что используется свежая версия программы и `XMLRIVER_MAX_CONCURRENCY` не выше допустимого лимита аккаунта.

## Лицензия

Проект распространяется по лицензии MIT. См. [LICENSE](LICENSE).
