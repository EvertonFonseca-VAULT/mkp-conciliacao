# Abre o Chrome com debug remoto (porta 9222) pro scraper ML/MP, num perfil
# dedicado (C:\ChromeDebugProfile) separado do Chrome normal do dia a dia.
#
# Por que precisa mais que so abrir o chrome.exe: o Chrome so aceita
# --remote-debugging-port/--user-data-dir na hora que o PROCESSO daquele
# perfil especifico comeca. Se sobrou um chrome.exe travado usando esse
# mesmo perfil de debug (janela fechada mas processo/lock ainda vivo -- ex.
# depois de um crash), o novo chrome.exe so repassa o comando pro antigo e
# ignora os parametros, sem abrir nada visivel. So mata o(s) processo(s)
# do PERFIL DE DEBUG (nunca o Chrome normal do usuario, que fica intacto).

$ErrorActionPreference = "SilentlyContinue"
$perfil = "C:\ChromeDebugProfile"

function DebugAberto {
    try {
        $r = Invoke-WebRequest -Uri "http://localhost:9222/json/version" -TimeoutSec 2 -UseBasicParsing
        return $r.StatusCode -eq 200
    } catch {
        return $false
    }
}

if (DebugAberto) {
    Write-Host "Chrome com debug remoto ja esta aberto na porta 9222."
} else {
    Write-Host "Limpando instancia presa do perfil de debug (se houver)..."
    Get-CimInstance Win32_Process -Filter "Name='chrome.exe'" |
        Where-Object { $_.CommandLine -like "*$perfil*" } |
        ForEach-Object { Stop-Process -Id $_.ProcessId -Force }

    $lock = Join-Path $perfil "SingletonLock"
    if (Test-Path $lock) { Remove-Item $lock -Force }

    Write-Host "Abrindo Chrome com debug remoto..."
    Start-Process "C:\Program Files\Google\Chrome\Application\chrome.exe" -ArgumentList @(
        "--remote-debugging-port=9222",
        "--user-data-dir=$perfil",
        "https://www.mercadopago.com.br"
    )

    Start-Sleep -Seconds 2
    if (DebugAberto) {
        Write-Host "Chrome aberto com debug remoto (porta 9222 respondendo)."
    } else {
        Write-Host "Aviso: o Chrome abriu mas a porta 9222 ainda nao respondeu -- espera mais um pouco e roda de novo."
    }
    Write-Host ""
    Write-Host "Faca login manual no Mercado Pago e no Mercado Livre antes de rodar o scraper (deixe uma aba de cada logada)."
}

Read-Host "Pressione Enter pra fechar"
