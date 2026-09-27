# SteamLibraryAnalyzer

PowerShell tool for analyzing public Steam libraries and exporting detailed statistics to Excel.

SteamLibraryAnalyzer can analyze one or two Steam profiles, compare their libraries, calculate playtime statistics, estimate the current Steam Store value of a library, group games by genres and franchises, and generate a short player profile in Russian.

---

## English

### Features

The script can collect and calculate:

- Total number of games
- Total playtime
- Playtime during the last 2 weeks
- Current estimated library value
- Estimated value without current discounts
- Games that were never launched
- Games with less than 1 hour
- Games with 1–10 hours
- Games with 10–50 hours
- Games with 50–100 hours
- Games with 100+ hours
- Abandoned games
- Favorite genres
- Total playtime by genre
- Game franchises
- Number of games in each franchise
- Total playtime by franchise
- Common games between two Steam accounts
- Games owned only by the first account
- Games owned only by the second account
- Library compatibility percentage
- Automatic short player profile

The result is exported to:

```text
Steam_Analysis.xlsx
```

### Excel sheets

The generated workbook contains:

- `Сводка` — general profile statistics
- `Все игры` — full game library
- `Мертвые игры` — games with 0 minutes played
- `Заброшенные` — games with significant playtime that have not been launched for a long time
- `Жанры` — genre statistics
- `Серии` — franchise statistics
- `Структура` — playtime distribution
- `Совместимость` — comparison of two profiles
- `Общие и различия` — common games and games owned by only one profile

### Requirements

- Windows 10 or Windows 11
- Windows PowerShell 5.1 or newer
- Internet connection
- Steam Web API Key
- Public Steam profile
- Public `Game details`

Microsoft Excel itself is not required to generate the `.xlsx` file.

The script uses the PowerShell module:

```powershell
ImportExcel
```

If the module is missing, the script will try to install it automatically.

### Steam privacy settings

The Steam profile must be public.

Open:

```text
Steam Profile
→ Edit Profile
→ Privacy Settings
```

Set:

```text
My profile = Public
Game details = Public
```

If `Game details` is private, the script will not be able to read the game library.

### Steam Web API Key

A Steam Web API Key is required.

Get one here:

https://steamcommunity.com/dev/apikey

For the domain field, you can use:

```text
localhost
```

The API key is entered when the script starts.

Do not publish your Steam Web API Key or hardcode it into the public version of the script.

### Quick start

1. Download the repository or the release archive.
2. Extract all files into one folder.
3. Run:

```text
START_Steam_Analyzer.bat
```

4. Enter your Steam Web API Key.
5. Enter the first Steam profile URL.
6. Enter the second profile URL or press Enter to skip it.
7. Wait until the analysis is complete.
8. Open:

```text
Steam_Analysis.xlsx
```

### Supported Steam profile links

Vanity URL:

```text
https://steamcommunity.com/id/ExampleName/
```

SteamID64 profile URL:

```text
https://steamcommunity.com/profiles/7656119XXXXXXXXXX/
```

### Manual PowerShell start

```powershell
powershell -ExecutionPolicy Bypass -File .\SteamLibraryAnalyzer.ps1
```

If Windows blocks the downloaded files:

```powershell
Unblock-File .\SteamLibraryAnalyzer.ps1
Unblock-File .\START_Steam_Analyzer.bat
```

### Manual ImportExcel installation

If automatic installation fails:

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

Then run the analyzer again.

### Library value

The library value is estimated from current Steam Store data.

The report can include:

- Current price
- Price before the current discount
- Free games
- Games for which Steam returned pricing information

This is **not the real amount of money spent by the account owner**.

The script cannot determine:

- The original purchase price
- Historical discounts
- Bundle prices
- Gifted games
- Activated Steam keys
- Purchases from third-party stores
- Games received for free

The value should therefore be treated only as an estimate based on current Steam Store data.

### Dead games

By default, a game is considered dead if it has:

```text
0 minutes played
```

The setting can be changed in the script:

```powershell
$DeadGameMaxMinutes = 0
```

### Abandoned games

By default, a game is considered abandoned if it has at least:

```text
20 hours played
```

and has not been launched for:

```text
365 days
```

Settings:

```powershell
$AbandonedMinHours = 20
$AbandonedDays = 365
```

For example, to use 6 months:

```powershell
$AbandonedDays = 180
```

### Franchise detection

Some franchises are detected by matching game names.

Examples:

- STAR WARS
- Call of Duty
- Need for Speed
- Grand Theft Auto
- Assassin's Creed
- Resident Evil
- Half-Life
- Portal
- Fallout
- The Elder Scrolls
- Far Cry
- Battlefield
- Forza
- DOOM
- Wolfenstein
- Metro
- Tom Clancy's
- Counter-Strike

The list can be edited inside:

```powershell
$FranchisePatterns
```

### Cache

The script creates:

```text
steam_store_cache.json
```

This cache stores Steam Store metadata and makes repeated scans much faster.

If the data appears outdated, delete the cache file and run the script again.

### Generated files

After running the analyzer, these files may appear:

```text
Steam_Analysis.xlsx
steam_store_cache.json
```

They normally should not be committed to GitHub.

Recommended `.gitignore`:

```gitignore
Steam_Analysis.xlsx
steam_store_cache.json
*.xlsx
.env
```

### Limitations

Some data may be unavailable when:

- The Steam profile is private
- `Game details` is private
- A game has been removed from Steam
- A Store page is unavailable
- Steam temporarily rejects requests
- A game has no current price
- Steam changes its API behavior

Large libraries may take some time to analyze on the first run.

Repeated scans are usually much faster because Store information is cached locally.

### Security

Never publish your Steam Web API Key.

If a key is accidentally committed to a public repository, revoke or replace it.

### License

This project is licensed under the MIT License.

See:

```text
LICENSE
```

for details.

### Author

**Andrey Moiseyenka**

GitHub: `Ovose`

---

# Русская версия

SteamLibraryAnalyzer — PowerShell-скрипт для анализа публичных библиотек Steam и экспорта подробной статистики в Excel.

Скрипт может анализировать один или два Steam-профиля, сравнивать их библиотеки, считать игровое время, оценивать текущую стоимость библиотеки по данным Steam Store, группировать игры по жанрам и сериям и создавать краткий портрет игрока.

## Возможности

Скрипт собирает и рассчитывает:

- Общее количество игр
- Общее количество часов
- Часы за последние 2 недели
- Текущую примерную стоимость библиотеки
- Стоимость библиотеки без текущих скидок
- Игры, которые ни разу не запускались
- Игры с менее чем 1 часом
- Игры с 1–10 часами
- Игры с 10–50 часами
- Игры с 50–100 часами
- Игры со 100+ часами
- Заброшенные игры
- Любимые жанры
- Общее количество часов по жанрам
- Игровые серии
- Количество игр в каждой серии
- Общее количество часов по сериям
- Общие игры двух Steam-аккаунтов
- Игры, которые есть только у первого аккаунта
- Игры, которые есть только у второго аккаунта
- Процент совпадения библиотек
- Автоматический краткий портрет игрока

Результат сохраняется в:

```text
Steam_Analysis.xlsx
```

## Листы Excel

В итоговом файле создаются:

- `Сводка` — общая статистика профиля
- `Все игры` — полная библиотека
- `Мертвые игры` — игры с 0 минутами
- `Заброшенные` — игры, в которые много играли, но давно не запускали
- `Жанры` — статистика по жанрам
- `Серии` — статистика по игровым сериям
- `Структура` — распределение игр по количеству часов
- `Совместимость` — сравнение двух профилей
- `Общие и различия` — общие игры и игры, которые есть только у одного профиля

## Требования

- Windows 10 или Windows 11
- Windows PowerShell 5.1 или новее
- Интернет
- Steam Web API Key
- Публичный Steam-профиль
- Открытый `Game details`

Microsoft Excel для создания `.xlsx` не обязателен.

Скрипт использует PowerShell-модуль:

```powershell
ImportExcel
```

Если модуль отсутствует, скрипт попробует установить его автоматически.

## Настройки приватности Steam

Steam-профиль должен быть открытым.

Открой:

```text
Steam Profile
→ Edit Profile
→ Privacy Settings
```

Установи:

```text
My profile = Public
Game details = Public
```

Если `Game details` закрыт, скрипт не сможет получить библиотеку игр.

## Steam Web API Key

Для работы нужен Steam Web API Key.

Получить его можно здесь:

https://steamcommunity.com/dev/apikey

В поле Domain Name можно указать:

```text
localhost
```

API key вводится при запуске.

Не публикуй свой Steam Web API Key и не прописывай личный ключ прямо в публичной версии скрипта.

## Быстрый запуск

1. Скачай репозиторий или архив релиза.
2. Распакуй все файлы в одну папку.
3. Запусти:

```text
START_Steam_Analyzer.bat
```

4. Введи Steam Web API Key.
5. Вставь ссылку на первый Steam-профиль.
6. Вставь ссылку на второй профиль или нажми Enter, чтобы пропустить.
7. Дождись завершения анализа.
8. Открой:

```text
Steam_Analysis.xlsx
```

## Поддерживаемые ссылки Steam

Обычная ссылка:

```text
https://steamcommunity.com/id/ExampleName/
```

Ссылка через SteamID64:

```text
https://steamcommunity.com/profiles/7656119XXXXXXXXXX/
```

## Ручной запуск через PowerShell

```powershell
powershell -ExecutionPolicy Bypass -File .\SteamLibraryAnalyzer.ps1
```

Если Windows заблокировал скачанные файлы:

```powershell
Unblock-File .\SteamLibraryAnalyzer.ps1
Unblock-File .\START_Steam_Analyzer.bat
```

## Установка ImportExcel вручную

Если автоматическая установка не сработала:

```powershell
Install-Module ImportExcel -Scope CurrentUser -Force
```

После этого снова запусти анализатор.

## Стоимость библиотеки

Стоимость библиотеки рассчитывается по текущим данным Steam Store.

В отчёте могут использоваться:

- Текущая цена
- Цена без текущей скидки
- Бесплатные игры
- Игры, для которых Steam смог вернуть цену

Это **не точная сумма денег, которую владелец аккаунта реально потратил**.

Скрипт не может определить:

- Реальную цену покупки
- Старые скидки
- Стоимость bundle
- Подаренные игры
- Игры, активированные ключом
- Покупки в сторонних магазинах
- Игры, полученные бесплатно

Поэтому стоимость библиотеки — это только примерная оценка по текущим данным Steam Store.

## Мёртвые игры

По умолчанию мёртвой считается игра с:

```text
0 минут игрового времени
```

Настройка:

```powershell
$DeadGameMaxMinutes = 0
```

## Заброшенные игры

По умолчанию игра считается заброшенной, если в ней наиграно минимум:

```text
20 часов
```

и она не запускалась:

```text
365 дней
```

Настройки:

```powershell
$AbandonedMinHours = 20
$AbandonedDays = 365
```

Например, для 6 месяцев:

```powershell
$AbandonedDays = 180
```

## Определение игровых серий

Некоторые серии определяются по названию игры.

Например:

- STAR WARS
- Call of Duty
- Need for Speed
- Grand Theft Auto
- Assassin's Creed
- Resident Evil
- Half-Life
- Portal
- Fallout
- The Elder Scrolls
- Far Cry
- Battlefield
- Forza
- DOOM
- Wolfenstein
- Metro
- Tom Clancy's
- Counter-Strike

Список можно изменить в переменной:

```powershell
$FranchisePatterns
```

## Кэш

Скрипт создаёт:

```text
steam_store_cache.json
```

В нём сохраняются данные Steam Store.

Это значительно ускоряет повторные запуски.

Если кажется, что данные устарели, удали файл кэша и снова запусти скрипт.

## Создаваемые файлы

После запуска могут появиться:

```text
Steam_Analysis.xlsx
steam_store_cache.json
```

Обычно их не стоит загружать в GitHub.

Рекомендуемый `.gitignore`:

```gitignore
Steam_Analysis.xlsx
steam_store_cache.json
*.xlsx
.env
```

## Ограничения

Некоторые данные могут отсутствовать, если:

- Steam-профиль приватный
- `Game details` закрыт
- Игра удалена из Steam
- Страница игры недоступна
- Steam временно отклоняет запросы
- У игры нет актуальной цены
- Steam изменил работу API

Первый анализ большой библиотеки может занять некоторое время.

Повторные запуски обычно работают намного быстрее благодаря локальному кэшу.

## Безопасность

Никогда не публикуй свой Steam Web API Key.

Если ключ случайно попал в публичный репозиторий, отзови или замени его.

## Лицензия

Проект распространяется под лицензией MIT.

Подробности находятся в файле:

```text
LICENSE
```

## Автор

**Andrey Moiseyenka**

GitHub: `Ovose`
