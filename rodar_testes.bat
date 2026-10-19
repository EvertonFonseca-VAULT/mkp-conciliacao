@echo off
REM Autoteste de cada modulo: 100% offline, nao mexe em X:/C:\RODAR TOTVS nem
REM precisa de login no Koncili. Serve pra confirmar que a instalacao (Python
REM + requirements.txt) esta funcionando antes de rodar com dado real.
REM Rodar deste mesmo lugar (nao precisa "cd" em cada pasta).

echo ===== koncili.py =====
python KONCILI\koncili.py --teste
if errorlevel 1 goto ERRO

echo ===== koncili_shopee.py =====
python KONCILI\koncili_shopee.py --teste
if errorlevel 1 goto ERRO

echo ===== rodar_canal.py =====
python PARA_PROCESSAR\rodar_canal.py --teste
if errorlevel 1 goto ERRO

echo ===== rodar_es.py =====
python PARA_PROCESSAR\ES\rodar_es.py --teste
if errorlevel 1 goto ERRO

echo ===== escolher_full.py =====
python PARA_PROCESSAR\FULL\escolher_full.py --self-test
if errorlevel 1 goto ERRO

echo ===== escolher_es.py =====
python PARA_PROCESSAR\ES\escolher_es.py --self-test
if errorlevel 1 goto ERRO

echo ===== empilhar_canais.py =====
python EMPILHAR\empilhar_canais.py --teste
if errorlevel 1 goto ERRO

echo ===== dashboard.py (CONCILIACAO) =====
python CONCILIACAO\dashboard.py
if errorlevel 1 goto ERRO

echo ===== cofre_senhas.py (cofre de senhas do Windows) =====
python COMUM\cofre_senhas.py --teste
if errorlevel 1 goto ERRO

echo.
echo TUDO OK.
pause
exit /b 0

:ERRO
echo.
echo FALHOU um dos testes acima -- ver mensagem de erro.
pause
exit /b 1
