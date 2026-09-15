@echo off
REM Lancador fino -- a logica real esta no .ps1 do lado (mais facil de
REM manter: evita aspas aninhadas de PowerShell dentro de .bat).
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0abrir_chrome_debug.ps1"
