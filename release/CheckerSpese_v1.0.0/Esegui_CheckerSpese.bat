@echo off
chcp 65001 >nul
color 0A
title Checker Spese - Avvio

echo.
echo ═══════════════════════════════════════════════════════════════
echo   CHECKER SPESE - Bot per pulizia dati Excel
echo ═══════════════════════════════════════════════════════════════
echo.

:: Controlla se l'eseguibile esiste
if exist "CheckerSpese.exe" (
    echo [✓] Eseguibile trovato: CheckerSpese.exe
    echo.
    echo [INFO] Esecuzione dell'applicazione...
    echo.
    start "" "CheckerSpese.exe"
    timeout /t 2 >nul
    exit
)

:: Se non c'è l'eseguibile, prova con Python
echo [!] Eseguibile CheckerSpese.exe non trovato
echo [INFO] Tentativo di esecuzione tramite Python...
echo.

:: Verifica Python
python --version >nul 2>&1
if errorlevel 1 (
    color 0C
    echo [✗] ERRORE: Python non è installato!
    echo.
    echo Per utilizzare questo programma hai bisogno di:
    echo   1. Scaricare Python da: https://www.python.org/downloads/
    echo   2. Durante l'installazione, selezionare "Add Python to PATH"
    echo   3. Riavviare il computer
    echo.
    pause
    exit /b 1
)

echo [✓] Python installato correttamente
python --version

:: Verifica tkinter
echo [INFO] Verifica modulo tkinter...
python -c "import tkinter" >nul 2>&1
if errorlevel 1 (
    color 0E
    echo [!] ATTENZIONE: Modulo tkinter non disponibile!
    echo.
    echo Questo è raro su Windows. Soluzioni:
    echo   1. Reinstalla Python assicurandoti di selezionare "tcl/tk and IDLE"
    echo   2. Oppure usa il Launcher di Python per installare tkinter
    echo.
    pause
    exit /b 1
)

echo [✓] Modulo tkinter disponibile

:: Verifica dipendenze
echo [INFO] Verifica dipendenze...
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    color 0E
    echo [!] ATTENZIONE: Dipendenze mancanti!
    echo.
    echo Installazione in corso...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [✗] ERRORE: Impossibile installare le dipendenze
        echo Verifica la connessione internet e riprova.
        pause
        exit /b 1
    )
    echo [✓] Dipendenze installate
)

echo [✓] Tutte le dipendenze sono installate
echo.

:: Controlla se ci sono file Excel
set found_excel=0
for %%f in (*.xlsx *.xlsm) do (
    if not "%%f"=="errori.xlsx" (
        if not "%%~nf"=="clean_*" (
            set found_excel=1
        )
    )
)

if %found_excel%==0 (
    color 0E
    echo [!] ATTENZIONE: Nessun file Excel (.xlsx/.xlsm) trovato in questa cartella!
    echo.
    echo Posiziona il file Excel da processare nella stessa cartella di questo script.
    echo.
    pause
    exit /b 1
)

:: Esegui il programma
echo ═══════════════════════════════════════════════════════════════
echo   Avvio del programma...
echo ═══════════════════════════════════════════════════════════════
echo.

python checker_spese.py

if errorlevel 1 (
    color 0C
    echo.
    echo [✗] Si è verificato un errore durante l'esecuzione.
    echo.
    pause
    exit /b 1
)

echo.
echo ═══════════════════════════════════════════════════════════════
echo   Processo completato!
echo ═══════════════════════════════════════════════════════════════
echo.
pause
