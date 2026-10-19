@echo off
setlocal enabledelayedexpansion

:MENU
cls
echo ==========================================
echo   CONCILIACAO MERCADO LIVRE / TOTVS
echo ==========================================
echo Qual canal voce quer processar?
echo   1) FULL
echo   2) COLETAS ES
echo   3) MOVEIS
echo   4) ELETRODOMESTICO
echo   5) CENTER
echo   6) COLETAS CO05
echo   7) PRIME
echo   8) WEB
echo.
set /p CANAL_NUM="Escolha o canal (1 a 8): "

if "%CANAL_NUM%"=="1" set CANAL=FULL
if "%CANAL_NUM%"=="2" set CANAL=ES
if "%CANAL_NUM%"=="3" set CANAL=MOVEIS
if "%CANAL_NUM%"=="4" set CANAL=ELETRO
if "%CANAL_NUM%"=="5" set CANAL=CENTER
if "%CANAL_NUM%"=="6" set CANAL=CO05
if "%CANAL_NUM%"=="7" set CANAL=PRIME
if "%CANAL_NUM%"=="8" set CANAL=WEB

if not defined CANAL (
    echo Opcao invalida.
    pause
    goto MENU
)

echo.
echo Como deseja processar?
echo   1) Data unica (dd.mm)
echo   2) Intervalo de datas (dd.mm a dd.mm)
echo.
set /p MODO="Escolha (1 ou 2): "

if "%MODO%"=="1" goto DATA_UNICA
if "%MODO%"=="2" goto INTERVALO
echo Opcao invalida.
pause
goto MENU

:DATA_UNICA
set /p DATA="Digite a data (dd.mm): "
cd /d "%~dp0"
"%~dp0..\.venv\Scripts\python.exe" interativo.py %CANAL% %DATA% > "%TEMP%\conciliacao_out.txt" 2>&1
type "%TEMP%\conciliacao_out.txt"
goto MOSTRAR_RESUMO_UNICO

:INTERVALO
set /p DATA_INI="Data inicial (dd.mm): "
set /p DATA_FIM="Data final   (dd.mm): "
cd /d "%~dp0"
"%~dp0..\.venv\Scripts\python.exe" interativo.py %CANAL% --data-inicio %DATA_INI% --data-fim %DATA_FIM% > "%TEMP%\conciliacao_out.txt" 2>&1
type "%TEMP%\conciliacao_out.txt"
goto MOSTRAR_RESUMO_INTERVALO

:MOSTRAR_RESUMO_UNICO
echo.
echo ==========================================
echo   RESUMO DA CONCILIACAO
echo ==========================================
for /f "tokens=1-6 delims=|" %%A in ('findstr /R /C:"^RESUMO|" "%TEMP%\conciliacao_out.txt"') do (
    echo Canal:        %%B
    echo Data:         %%C
    echo H a M:        %%D
    echo VERIFICAR:    %%E
    echo Calendario:   %%F
)
echo ==========================================
pause
goto MENU

:MOSTRAR_RESUMO_INTERVALO
echo.
echo ==========================================
echo   RESUMO DO INTERVALO
echo ==========================================
for /f "tokens=1-6 delims=|" %%A in ('findstr /R /C:"^RESUMO|" "%TEMP%\conciliacao_out.txt"') do (
    echo Canal: %%B   Data: %%C   H:M=%%D   Verificar=%%E   Calendario=%%F
)
echo ==========================================
pause
goto MENU