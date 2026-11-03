# Gera teste.html, roda a bateria no Chrome headless e mostra PASS/FAIL.
# Uso:  powershell -ExecutionPolicy Bypass -File .\rodar_testes.ps1

$ErrorActionPreference = "Stop"
$raiz = Split-Path -Parent $MyInvocation.MyCommand.Path
$py = $env:DSH_PYTHON
if (-not $py) { $py = "python" }

$gerador = Join-Path $raiz "analise\gerar_teste.py"
$pagina  = Join-Path $raiz "app\teste.html"
$dump    = Join-Path $raiz "app\_dom.txt"

$chrome = @(
  "C:\Program Files\Google\Chrome\Application\chrome.exe",
  "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
  "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"
) | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1

if (-not $chrome) { Write-Error "Chrome ou Edge nao encontrado."; exit 1 }
if (-not (Test-Path $gerador)) { Write-Error "Gerador ausente: $gerador"; exit 1 }

& $py $gerador | Out-Null
if (-not (Test-Path $pagina)) { Write-Error "teste.html nao foi gerado."; exit 1 }

if (Test-Path $dump) { Remove-Item $dump -Force }
$tmp = Join-Path $env:TEMP ("chr_" + [guid]::NewGuid().ToString("N").Substring(0,8))
$url = "file:///" + ($pagina -replace '\\','/')

$p = Start-Process $chrome -ArgumentList @(
  '--headless=new','--no-sandbox','--disable-gpu',
  "--user-data-dir=$tmp",'--virtual-time-budget=3000','--dump-dom',$url
) -RedirectStandardOutput $dump -PassThru -WindowStyle Hidden

if (-not $p.WaitForExit(30000)) { $p.Kill(); Write-Host "TIMEOUT (processo encerrado)" -ForegroundColor Yellow }

if (-not (Test-Path $dump)) { Write-Error "sem saida do navegador."; exit 1 }
$dom = Get-Content $dump -Raw -Encoding UTF8
$i = $dom.IndexOf('id="RES"')
if ($i -lt 0) { Write-Error "bloco de resultados nao encontrado."; exit 1 }

$trecho = $dom.Substring($i, [Math]::Min(4000, $dom.Length - $i))
$trecho = $trecho -replace '</pre>[\s\S]*',''
$trecho = $trecho -replace 'id="RES">',''

$pass = 0; $fail = 0
foreach ($linha in ($trecho -split ' \|\|\| ')) {
  $t = $linha.Trim()
  if ($t -like 'PASS*') { $pass++; Write-Host $t -ForegroundColor Green }
  elseif ($t -like 'FAIL*') { $fail++; Write-Host $t -ForegroundColor Red }
  elseif ($t) { Write-Host $t -ForegroundColor Yellow }
}

Write-Host ""
Write-Host ("RESULTADO: {0} PASS, {1} FAIL" -f $pass, $fail) -ForegroundColor Cyan
if ($fail -gt 0) { exit 1 }
