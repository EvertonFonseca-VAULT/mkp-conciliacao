@echo off
REM Caminho fixo (nao %~dp0), mesmo padrao do iniciar_conciliacao.bat (AB/RB)
REM -- precisa do projeto clonado nesse caminho (por usuario).
cd /d "%USERPROFILE%\Desktop\Projetos\Mkp\AB\FULL"
set /p DATA="Data a processar (DD.MM.AAAA): "
py rodar_full.py --data %DATA%
pause
