<#
.SYNOPSIS
Abre o Microsoft Edge com debug port (9223) para que o script Python conecte via CDP.

.DESCRIPTION
Lança o Edge instalado no sistema com --remote-debugging-port=9223 e
um user-data-dir dedicado em C:\temp\edge_debug. Após abrir, vc loga
manualmente no TJRJ e no confirmeonline. Daí roda o script Python.

NOTA: usamos Edge porque o antivírus corporativo costuma bloquear Chrome
com flag de debug port (técnica também usada por malware).
#>

$EdgePaths = @(
    "${env:ProgramFiles(x86)}\Microsoft\Edge\Application\msedge.exe",
    "${env:ProgramFiles}\Microsoft\Edge\Application\msedge.exe",
    "${env:LocalAppData}\Microsoft\Edge\Application\msedge.exe"
)

$EdgeExe = $EdgePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $EdgeExe) {
    Write-Host "ERRO: Edge nao encontrado nos caminhos padrao." -ForegroundColor Red
    $EdgePaths | ForEach-Object { Write-Host "  $_" }
    exit 1
}

$UserDataDir = "C:\temp\edge_debug"
$DebugPort = 9223

if (-not (Test-Path $UserDataDir)) {
    New-Item -ItemType Directory -Path $UserDataDir -Force | Out-Null
    Write-Host "Pasta de perfil criada: $UserDataDir" -ForegroundColor Green
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Iniciando Edge com debug port $DebugPort" -ForegroundColor Cyan
Write-Host " Edge:   $EdgeExe" -ForegroundColor Cyan
Write-Host " Perfil: $UserDataDir" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Espere o Edge abrir (janela separada do seu Edge normal)" -ForegroundColor Yellow
Write-Host "  2. Pule o setup inicial se aparecer ('Skip' / 'Iniciar sem fazer login')" -ForegroundColor Yellow
Write-Host "  3. Faca login no TJRJ (resolva o codigo verificador)" -ForegroundColor Yellow
Write-Host "  4. Faca login no confirmeonline.com.br" -ForegroundColor Yellow
Write-Host "  5. NAO feche essa janela do Edge" -ForegroundColor Yellow
Write-Host "  6. Em outro PowerShell, rode o script Python" -ForegroundColor Yellow
Write-Host ""

Start-Process -FilePath $EdgeExe -ArgumentList @(
    "--remote-debugging-port=$DebugPort",
    "--user-data-dir=$UserDataDir"
)

Write-Host "Edge lancado em background. Esperando debug port abrir..." -ForegroundColor Cyan
$ok = $false
for ($i = 0; $i -lt 15; $i++) {
    Start-Sleep -Seconds 1
    try {
        $null = Invoke-WebRequest -Uri "http://127.0.0.1:$DebugPort/json/version" -UseBasicParsing -TimeoutSec 2 -ErrorAction Stop
        Write-Host "OK! Debug port respondendo em http://127.0.0.1:$DebugPort" -ForegroundColor Green
        $ok = $true
        break
    } catch {}
}
if (-not $ok) {
    Write-Host "AVISO: Debug port nao respondeu apos 15s. Verifique o Edge." -ForegroundColor Yellow
}
