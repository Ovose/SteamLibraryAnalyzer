STEAM LIBRARY ANALYZER

БЫСТРЫЙ ЗАПУСК
1. Распакуйте всю папку из ZIP.
2. Запустите START_Steam_Analyzer.bat.
3. Если ImportExcel отсутствует, скрипт попробует установить его сам.
4. Введите Steam Web API key.
5. Вставьте ссылку на Steam-профиль.
6. Для второго профиля вставьте вторую ссылку или нажмите Enter.
7. Результат появится рядом со скриптом: Steam_Analysis.xlsx

ВАЖНО ДЛЯ STEAM
Профиль должен быть открыт:
Profile -> Edit Profile -> Privacy Settings
My profile = Public
Game details = Public

API KEY
https://steamcommunity.com/dev/apikey

РУЧНОЙ ЗАПУСК
Откройте PowerShell в папке со скриптом и выполните:

powershell -ExecutionPolicy Bypass -File .\SteamLibraryAnalyzer.ps1

ЕСЛИ IMPORTEXCEL НЕ УСТАНОВИЛСЯ АВТОМАТИЧЕСКИ
Install-Module ImportExcel -Scope CurrentUser -Force

ЕСЛИ WINDOWS ЗАБЛОКИРОВАЛ СКАЧАННЫЕ ФАЙЛЫ
Unblock-File .\SteamLibraryAnalyzer.ps1
Unblock-File .\START_Steam_Analyzer.bat

После этого снова запустите START_Steam_Analyzer.bat.
