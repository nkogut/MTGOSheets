REM Syncs daily MTGO card prices and updates tracker spreadsheets
REM Schedule this with a few minutes delay from login
REM Update PYEXE to point to your python executable if needed

setlocal enabledelayedexpansion
set "DIR=%~dp0"
set "PYTHONPATH=%DIR%\src"
set "PYEXE=%DIR%venv\Scripts\python.exe"

set "LOG_DIR=%DIR%\logs\%DATE:~10,4%_%DATE:~4,2%"
set "LOG_FILE=%LOG_DIR%\update_prices_log_%DATE:~10,4%_%DATE:~4,2%_%DATE:~7,2%.txt"

if not exist "%LOG_DIR%" mkdir "%LOG_DIR%"

REM Only run if there is no log for today yet
if not exist "%LOG_FILE%" (
    set "START_TIME=%TIME%"
    echo Begin updating >> "%LOG_FILE%"
    echo Start time: !START_TIME! >> "%LOG_FILE%"

    "%PYEXE%" -u -m mtgo_sheets.updater >> "%LOG_FILE%" 2>&1
    "%PYEXE%" -u -m mtgo_sheets.tracker >> "%LOG_FILE%" 2>&1

    echo Updating complete >> "%LOG_FILE%"
    set "END_TIME=!TIME!"
    echo End time: !END_TIME! >> "%LOG_FILE%"
)
