@echo off
REM Caminho fixo (nao %~dp0) -- assim esse .bat pode ficar em qualquer
REM lugar (ex. na pasta de saida no X:), sem precisar levar os modulos
REM .py junto. Precisa do projeto clonado nesse caminho (por usuario).
cd /d "%USERPROFILE%\Desktop\Projetos\Mkp\AB\RB"
py interativo.py
pause
