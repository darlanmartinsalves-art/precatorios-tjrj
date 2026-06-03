# Consulta saldo atualizado de precatórios no TJRJ e preenche coluna H da planilha
# API: GET https://www3.tjrj.jus.br/PortalConhecimento/api/precatorios/consultaPosicao?numeroPrecatorio=XXXX

param(
    [string]$ArquivoEntrada = "C:\Users\DARLANMARTINS\Downloads\Precatórios 2027.xlsx",
    [string]$ArquivoSaida   = "C:\Users\DARLANMARTINS\Downloads\Precatórios 2027 - Atualizado.xlsx",
    [int]   $MaxParalelo    = 15
)

Add-Type -AssemblyName System.IO.Compression.FileSystem
Add-Type -AssemblyName System.IO.Compression

# ── 1. LER XLSX ──────────────────────────────────────────────────────────────
Write-Host "Lendo planilha..." -ForegroundColor Cyan
$zip = [System.IO.Compression.ZipFile]::OpenRead($ArquivoEntrada)

$ssEntry  = $zip.GetEntry("xl/sharedStrings.xml")
$ssReader = New-Object System.IO.StreamReader($ssEntry.Open())
$ssXml    = $ssReader.ReadToEnd(); $ssReader.Close()

$shEntry  = $zip.GetEntry("xl/worksheets/sheet1.xml")
$shReader = New-Object System.IO.StreamReader($shEntry.Open())
$shXml    = $shReader.ReadToEnd(); $shReader.Close()

$stEntry  = $zip.GetEntry("xl/styles.xml")
$stReader = New-Object System.IO.StreamReader($stEntry.Open())
$stXml    = $stReader.ReadToEnd(); $stReader.Close()

$zip.Dispose()

# Decodificar sharedStrings
$strings = @()
[regex]::Matches($ssXml, '<t[^>]*>([^<]*)</t>') | ForEach-Object { $strings += $_.Groups[1].Value }

# ── 2. EXTRAIR NÚMEROS DE PRECATÓRIO ────────────────────────────────────────
Write-Host "Extraindo números de precatório..." -ForegroundColor Cyan
$rowMatches = [regex]::Matches($shXml, '<row r="(\d+)"[^>]*>(.*?)</row>')

$precatorios = [System.Collections.Generic.List[PSCustomObject]]::new()
foreach ($rm in $rowMatches) {
    $rowNum  = [int]$rm.Groups[1].Value
    if ($rowNum -lt 2) { continue }
    $rowBody = $rm.Groups[2].Value
    $bCell   = [regex]::Match($rowBody, '<c r="B' + $rowNum + '"[^>]*>.*?<v>(\d+)</v>.*?</c>')
    if (-not $bCell.Success) { continue }
    $numPrec = $strings[[int]$bCell.Groups[1].Value]
    if ($numPrec -match '^\d{4}\.\d') {
        $precatorios.Add([PSCustomObject]@{ Row = $rowNum; Numero = $numPrec })
    }
}
$total = $precatorios.Count
Write-Host "Total de precatórios: $total" -ForegroundColor Green

# ── 3. CONSULTAR API EM PARALELO ────────────────────────────────────────────
$baseUrl    = "https://www3.tjrj.jus.br/PortalConhecimento/api/precatorios/consultaPosicao"
$resultados = [System.Collections.Hashtable]::Synchronized(@{})
$contador   = [System.Collections.Hashtable]::Synchronized(@{ N = 0 })

$scriptBlock = {
    param($numero, $baseUrl, $resultados, $contador)
    try {
        $resp = Invoke-WebRequest -Uri "${baseUrl}?numeroPrecatorio=${numero}" -UseBasicParsing -TimeoutSec 30 -ErrorAction Stop
        $json = $resp.Content | ConvertFrom-Json
        $resultados[$numero] = $json.Saldo
    } catch {
        $resultados[$numero] = "ERRO"
    }
    [void][System.Threading.Interlocked]::Increment([ref]$contador["N"])
}

$pool = [RunspaceFactory]::CreateRunspacePool(1, $MaxParalelo)
$pool.Open()
$jobs = [System.Collections.Generic.List[PSCustomObject]]::new()

Write-Host "Consultando API ($MaxParalelo em paralelo)..." -ForegroundColor Cyan
$sw = [System.Diagnostics.Stopwatch]::StartNew()

foreach ($p in $precatorios) {
    $ps = [PowerShell]::Create()
    $ps.RunspacePool = $pool
    [void]$ps.AddScript($scriptBlock)
    [void]$ps.AddArgument($p.Numero)
    [void]$ps.AddArgument($baseUrl)
    [void]$ps.AddArgument($resultados)
    [void]$ps.AddArgument($contador)
    $jobs.Add([PSCustomObject]@{ PS = $ps; Handle = $ps.BeginInvoke() })
}

# Progresso
while ($contador["N"] -lt $total) {
    $done = $contador["N"]
    $pct  = [Math]::Round(($done / $total) * 100, 1)
    $rate = if ($sw.Elapsed.TotalSeconds -gt 0) { [Math]::Round($done / $sw.Elapsed.TotalSeconds, 1) } else { 1 }
    $rem  = if ($rate -gt 0) { [int](($total - $done) / $rate) } else { 0 }
    Write-Progress -Activity "Consultando TJRJ" -Status "$done / $total ($pct%)  |  ${rate}/s  |  restam ${rem}s" -PercentComplete $pct
    Start-Sleep -Milliseconds 800
}
Write-Progress -Completed -Activity "Consultando TJRJ"

foreach ($j in $jobs) { try { $j.PS.EndInvoke($j.Handle) } catch {} ; $j.PS.Dispose() }
$pool.Close(); $pool.Dispose()
$sw.Stop()

$ok    = ($resultados.Values | Where-Object { $_ -isnot [string] }).Count
$erros = ($resultados.Values | Where-Object { $_ -eq "ERRO" }).Count
Write-Host "Consultas concluídas em $([int]$sw.Elapsed.TotalSeconds)s  |  OK: $ok  |  Erros: $erros" -ForegroundColor Green

# ── 4. ATUALIZAR XML DO SHEET ───────────────────────────────────────────────
Write-Host "Atualizando XML da planilha..." -ForegroundColor Cyan

# Adicionar estilo de moeda (numFmtId=4 = #,##0.00) ao styles.xml
# Contar xfs atuais e acrescentar um
$countMatch = [regex]::Match($stXml, 'cellXfs count="(\d+)"')
$oldCount   = [int]$countMatch.Groups[1].Value
$newCount   = $oldCount + 1
$estiloMoeda = $oldCount  # novo índice (0-based)

$novoXf = '<xf borderId="0" fillId="0" fontId="2" numFmtId="4" xfId="0" applyAlignment="1" applyFont="1" applyNumberFormat="1"><alignment readingOrder="0"/></xf>'
$stXml  = $stXml -replace ('cellXfs count="' + $oldCount + '"'), ('cellXfs count="' + $newCount + '"')
$stXml  = $stXml -replace '</cellXfs>', ($novoXf + '</cellXfs>')

# Reconstruir sheet XML processando linha a linha
$sbNovo = New-Object System.Text.StringBuilder
$pos    = 0

foreach ($rm in $rowMatches) {
    # Texto antes desta linha
    [void]$sbNovo.Append($shXml.Substring($pos, $rm.Index - $pos))

    $rowNum  = [int]$rm.Groups[1].Value
    $rowFull = $rm.Value
    $rowBody = $rm.Groups[2].Value

    # Verificar se temos saldo para esta linha
    $numPrec = $null
    $saldo   = $null
    $bCell   = [regex]::Match($rowBody, '<c r="B' + $rowNum + '"[^>]*>.*?<v>(\d+)</v>.*?</c>')
    if ($bCell.Success) {
        $numPrec = $strings[[int]$bCell.Groups[1].Value]
        if ($numPrec -and $resultados.ContainsKey($numPrec) -and $resultados[$numPrec] -isnot [string]) {
            $saldo = $resultados[$numPrec]
        }
    }

    if ($null -ne $saldo) {
        # Montar célula H
        $celH = '<c r="H' + $rowNum + '" s="' + $estiloMoeda + '"><v>' + $saldo + '</v></c>'

        # Verificar se já existe célula H nesta linha
        $hExiste = [regex]::Match($rowBody, '<c r="H' + $rowNum + '"[^/]*/?>(?:<v>[^<]*</v></c>|</c>)?')
        if ($hExiste.Success) {
            # Substituir a célula H existente
            $rowFull = $rowFull.Replace($hExiste.Value, $celH)
        } else {
            # Inserir a célula H antes do fechamento da tag </row>
            # Encontrar a posição correta: após G ou antes de I/J/...
            $insertBefore = $null
            foreach ($col in @('I','J','K','L','M','N')) {
                $m = [regex]::Match($rowBody, '<c r="' + $col + $rowNum + '"')
                if ($m.Success) { $insertBefore = '<c r="' + $col + $rowNum + '"'; break }
            }

            if ($insertBefore) {
                $rowFull = $rowFull.Replace($insertBefore, $celH + $insertBefore)
            } else {
                $rowFull = $rowFull.Replace('</row>', $celH + '</row>')
            }
        }
    }

    [void]$sbNovo.Append($rowFull)
    $pos = $rm.Index + $rm.Length
}

# Restante do XML após as linhas
[void]$sbNovo.Append($shXml.Substring($pos))
$shXmlNovo = $sbNovo.ToString()

# ── 5. SALVAR NOVO XLSX ─────────────────────────────────────────────────────
Write-Host "Salvando $ArquivoSaida ..." -ForegroundColor Cyan
Copy-Item $ArquivoEntrada $ArquivoSaida -Force

$zipOut = [System.IO.Compression.ZipFile]::Open($ArquivoSaida, [System.IO.Compression.ZipArchiveMode]::Update)

$e = $zipOut.GetEntry("xl/worksheets/sheet1.xml"); $e.Delete()
$w = New-Object System.IO.StreamWriter(($zipOut.CreateEntry("xl/worksheets/sheet1.xml")).Open(), [System.Text.Encoding]::UTF8)
$w.Write($shXmlNovo); $w.Close()

$e = $zipOut.GetEntry("xl/styles.xml"); $e.Delete()
$w = New-Object System.IO.StreamWriter(($zipOut.CreateEntry("xl/styles.xml")).Open(), [System.Text.Encoding]::UTF8)
$w.Write($stXml); $w.Close()

$zipOut.Dispose()

# ── 6. RELATÓRIO FINAL ──────────────────────────────────────────────────────
Write-Host ""
Write-Host "==========================================" -ForegroundColor Green
Write-Host " CONCLUIDO!" -ForegroundColor Green
Write-Host " Arquivo: $ArquivoSaida" -ForegroundColor Green
Write-Host " Saldos preenchidos : $ok" -ForegroundColor Green
if ($erros -gt 0) {
    Write-Host " Erros de consulta  : $erros" -ForegroundColor Yellow
    $logPath = "C:\Users\DARLANMARTINS\Downloads\erros_precatorios.csv"
    $precatorios | Where-Object { $resultados[$_.Numero] -eq "ERRO" } |
        Select-Object Row, Numero |
        Export-Csv -Path $logPath -NoTypeInformation -Encoding UTF8
    Write-Host " Log de erros salvo : $logPath" -ForegroundColor Yellow
}
Write-Host "==========================================" -ForegroundColor Green
