<#
.SYNOPSIS
Abre o Chrome com debug port para que o script Python possa se conectar via CDP.

.DESCRIPTION
Lança o Chrome instalado no sistema com --remote-debugging-port=9222 e
um user-data-dir dedicado em C:\temp\chrome_debug. Após abrir, vc loga
manualmente no TJRJ e no confirmeonline. Daí roda o script Python.
#>

$ChromePaths = @(
    "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "${env:LocalAppData}\Google\Chrome\Application\chrome.exe"
)

$ChromeExe = $ChromePaths | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $ChromeExe) {
    Write-Host "ERRO: Chrome não encontrado nos caminhos padrão." -ForegroundColor Red
    Write-Host "Caminhos verificados:" -ForegroundColor Yellow
    $ChromePaths | ForEach-Object { Write-Host "  $_" }
    exit 1
}

$UserDataDir = "C:\temp\chrome_debug"
$DebugPort = 9222

if (-not (Test-Path $UserDataDir)) {
    New-Item -ItemType Directory -Path $UserDataDir -Force | Out-Null
    Write-Host "Pasta de perfil criada: $UserDataDir" -ForegroundColor Green
}

Write-Host ""
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host " Iniciando Chrome com debug port $DebugPort" -ForegroundColor Cyan
Write-Host " Chrome: $ChromeExe" -ForegroundColor Cyan
Write-Host " Perfil: $UserDataDir" -ForegroundColor Cyan
Write-Host "==================================================" -ForegroundColor Cyan
Write-Host ""
Write-Host "  1. Espere o Chrome abrir" -ForegroundColor Yellow
Write-Host "  2. Faça login no TJRJ (resolva o código verificador)" -ForegroundColor Yellow
Write-Host "  3. Faça login no confirmeonline (se ainda nao logou)" -ForegroundColor Yellow
Write-Host "  4. NÃO feche essa janela do Chrome" -ForegroundColor Yellow
Write-Host "  5. Em outro PowerShell, rode o script Python" -ForegroundColor Yellow
Write-Host ""

& $ChromeExe --remote-debugging-port=$DebugPort --user-data-dir=$UserDataDir
