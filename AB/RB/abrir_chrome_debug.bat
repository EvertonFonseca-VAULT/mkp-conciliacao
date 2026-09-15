@echo off
REM Abre o Chrome com debug remoto (porta 9222) pro scraper ML/MP, so se
REM ainda nao estiver aberto. Depois de abrir, faca login manual em
REM mercadopago.com.br e mercadolivre.com.br (ver Documentacao Screaper MP
REM e ML.md) antes de rodar o scraper.

curl -s -m 2 http://localhost:9222/json/version >nul 2>&1
if %errorlevel%==0 (
    echo Chrome com debug remoto ja esta aberto na porta 9222.
) else (
    echo Abrindo Chrome com debug remoto...
    start "" "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="C:\ChromeDebugProfile"
    echo.
    echo Chrome aberto. Faca login manual no Mercado Pago e no Mercado Livre
    echo antes de rodar o scraper (deixe uma aba de cada logada).
)
pause
