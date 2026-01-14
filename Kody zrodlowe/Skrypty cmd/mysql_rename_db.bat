@echo off
setlocal enabledelayedexpansion

rem ---
rem MySQL Database Rename/Clone Script
rem Description: Performs a dump of an existing database and restores it under a new name.
rem ---

rem Configuration
set "CUR_PATH=%~dp0"
set "CONFIG_FILE=%CUR_PATH%db_config.set"

rem Initialization of paths
set "DB_DIR="
set "DB_HOST="

if not exist "%CONFIG_FILE%" goto :help
(
 set /p DB_DIR=
 set /p DB_HOST=
) < "%CONFIG_FILE%"

if "%DB_DIR%"=="" goto :help

set "BIN_PATH=%CUR_PATH%%DB_DIR%\bin\"
set "SQL_BACKUP_PATH=%CUR_PATH%sql_dumps\"

rem Ensure workspace directories exist
if not exist "%SQL_BACKUP_PATH%" (
    mkdir "%SQL_BACKUP_PATH%"
    echo [INFO] Created directory for SQL dumps.
)

echo.
echo =======================================================
echo   DATABASE UTILITY: Rename / Clone MySQL Database
echo =======================================================
echo.

rem Get arguments or prompt user
set "OLD_DB=%1"
set "NEW_DB=%2"
set "DB_PASS=%3"

if "%OLD_DB%"=="" set /p OLD_DB="Enter source database name: "
if "%NEW_DB%"=="" set /p NEW_DB="Enter target database name: "
if "%DB_PASS%"=="" set /p DB_PASS="Enter database password: "

echo [PROCESS] Exporting source database: %OLD_DB%...
"%BIN_PATH%mysqldump" -h %DB_HOST% -u root -p%DB_PASS% %OLD_DB% > "%SQL_BACKUP_PATH%%OLD_DB%.sql"

if %ERRORLEVEL% NEQ 0 (
    echo [ERROR] Export failed. Please check credentials and database name.
    pause
    exit /b %ERRORLEVEL%
)

echo [PROCESS] Creating target database: %NEW_DB%...
"%BIN_PATH%mysql" -h %DB_HOST% -u root -p%DB_PASS% --execute="CREATE DATABASE IF NOT EXISTS %NEW_DB% CHARACTER SET utf8 COLLATE utf8_general_ci;"

echo [PROCESS] Importing data to: %NEW_DB%...
"%BIN_PATH%mysql" -h %DB_HOST% -u root -p%DB_PASS% %NEW_DB% < "%SQL_BACKUP_PATH%%OLD_DB%.sql"

if %ERRORLEVEL% EQU 0 (
    echo.
    echo =======================================================
    echo   SUCCESS: Database cloned to %NEW_DB%
    echo =======================================================
) else (
    echo [ERROR] Import failed.
)

goto :exit

:help
echo.
echo Usage: %~nx0 [source_db] [target_db] [password]
echo.
echo Requirements:
echo  - A 'db_config.set' file in the script directory containing:
echo    Line 1: Relative path to MySQL bin folder
echo    Line 2: Database host address (e.g., localhost)
echo.
pause

:exit
endlocal