@echo off
REM Cria as pastas locais (C:) que o processo precisa. Rodar uma vez, antes
REM de usar koncili.py / rodar_canal.py pela primeira vez.
REM Nao cria nada na rede (X:) -- isso precisa estar mapeado separadamente,
REM ver LEIA-ME.txt.

echo Criando pastas locais...

if not exist "C:\KIT CONCILIACAO MKP" mkdir "C:\KIT CONCILIACAO MKP"
if not exist "C:\TEMP" mkdir "C:\TEMP"

if not exist "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\FULL"          mkdir "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\FULL"
if not exist "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\COLETAS ES"    mkdir "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\COLETAS ES"
if not exist "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\MOVEIS"        mkdir "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\MOVEIS"
if not exist "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\CENTER"        mkdir "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\CENTER"
if not exist "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\COLETAS CO05"  mkdir "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\COLETAS CO05"
if not exist "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\PRIME"         mkdir "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\PRIME"
if not exist "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\WEB"           mkdir "C:\RODAR TOTVS\PROCESSAR\MERCADO LIVRE\WEB"
if not exist "C:\RODAR TOTVS\ELETRODOMESTICO"                      mkdir "C:\RODAR TOTVS\ELETRODOMESTICO"

echo.
echo Feito. Falta:
echo   1) Mapear a unidade X: (rede) -- pedir caminho/acesso ao Everton/TI.
echo   2) Colocar o TOTVS_2022.xlsx em "C:\KIT CONCILIACAO MKP\" (pedir ao Everton).
echo   3) Colocar o(s) PTI do dia em C:\TEMP\ antes de rodar o Para Processar.
pause
