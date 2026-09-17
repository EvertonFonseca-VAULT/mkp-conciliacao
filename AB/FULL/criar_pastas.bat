@echo off
REM Cria a estrutura de pastas que o processo de "processar" (escolher_full.bat)
REM espera encontrar em %USERPROFILE%\Desktop\Projetos -- roda 1x por maquina,
REM antes de copiar os arquivos do processo pra dentro. Nao baixa nem copia
REM nada sozinho, so garante que as pastas existem (mkdir e idempotente aqui:
REM se ja existir, so avisa e segue).

set "BASE=%USERPROFILE%\Desktop\Projetos"

for %%P in (
    "%BASE%\Mkp\AB\FULL"
    "%BASE%\Mkp\AB\ES"
    "%BASE%\Mkp\AB\MOVEIS"
    "%BASE%\Screaper Mercado"
) do (
    if not exist "%%~P" (
        mkdir "%%~P"
        echo Criada: %%~P
    ) else (
        echo Ja existia: %%~P
    )
)

echo.
echo ============================================================
echo  Pastas prontas. Falta copiar os arquivos pra dentro delas:
echo ============================================================
echo.
echo  "%BASE%\Mkp\AB\FULL":
echo    - escolher_full.py
echo    - rodar_full.py
echo    - processar_full.py
echo    - modelo_ab_full.py
echo    - tipagem_full.py
echo    - requirements.txt
echo.
echo  "%BASE%\Mkp\AB\ES"  (so se for usar COLETAS ES):
echo    - escolher_es.py
echo    - rodar_es.py
echo    - processar_es.py
echo    - modelo_ab_es.py
echo    - tipagem_es.py
echo.
echo  "%BASE%\Mkp\AB\MOVEIS":
echo    - modelo_ab.py
echo.
echo  "%BASE%\Screaper Mercado":
echo    - Pagina.py
echo.
echo  O launcher de 2 cliques (processar_ab.bat -- pergunta FULL ou ES e
echo  continua sozinho) fica em C:\RODAR TOTVS -- copie ele pra la.
echo.
echo ============================================================
echo  Depois de copiar tudo:
echo ============================================================
echo    1^) pip install -r "%BASE%\Mkp\AB\FULL\requirements.txt"
echo    2^) Confirme que a maquina tem acesso a:
echo         - X:\109_GC_CTR_FINANCEIRO CONTAS A RECEBER\... (rede)
echo         - C:\KIT CONCILIACAO MKP\TOTVS_2022.xlsx
echo         - C:\TEMP\PTI FULL.csv  (e C:\TEMP\PTI ES.csv, se for usar ES)
echo    Sem esses 2 arquivos (dado, nao codigo), o processo trava tentando
echo    abrir arquivo que nao existe.
echo.
pause
