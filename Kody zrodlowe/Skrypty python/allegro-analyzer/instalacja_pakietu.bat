@echo off

REM Sprawdzenie, czy podano parametr
if "%1"=="" (
    echo Nie podano nazwy pakietu do zainstalowania.
    echo Uzycie: %0 nazwa_pakietu
    exit /b 1
)

REM Instalacja pakietu
python -m pip install %1
