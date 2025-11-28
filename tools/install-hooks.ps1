param(
  [switch]$Force
)

# 1) Ubicación del repo
$ErrorActionPreference = 'Stop'
$repo = (& git rev-parse --show-toplevel) 2>$null
if (-not $repo) { throw "No se pudo resolver la raíz del repositorio. Ejecuta este script desde el repo." }

# 2) Asegurar el .ps1 del hook (desde el sample)
$hooksSrcDir = Join-Path $repo 'hooks'
$precommitPs1 = Join-Path $hooksSrcDir 'pre-commit.ps1'
$precommitSample = Join-Path $hooksSrcDir 'pre-commit.sample'

if (-not (Test-Path $precommitPs1)) {
  if (Test-Path $precommitSample) {
    Copy-Item $precommitSample $precommitPs1 -Force:$Force
    Write-Host "Copiado hooks/pre-commit.sample -> hooks/pre-commit.ps1"
  } else {
    throw "No se encontró hooks/pre-commit.sample. Crea tu script en hooks/pre-commit.ps1"
  }
}

# 3) Crear wrappers en .git/hooks
$gitHooks = Join-Path $repo '.git/hooks'
New-Item -ItemType Directory -Path $gitHooks -Force | Out-Null

$bashWrapper = @'
#!/usr/bin/env bash
set -euo pipefail
repo_root="$(git rev-parse --show-toplevel)"
if command -v pwsh >/dev/null 2>&1; then
  pwsh -NoProfile -ExecutionPolicy Bypass -File "$repo_root/hooks/pre-commit.ps1"
else
  powershell -NoProfile -ExecutionPolicy Bypass -File "$repo_root/hooks/pre-commit.ps1"
fi
'@

$cmdWrapper = @'
@echo off
setlocal
for /f "delims=" %%i in ('git rev-parse --show-toplevel') do set repo=%%i
powershell -NoProfile -ExecutionPolicy Bypass -File "%repo%\hooks\pre-commit.ps1"
exit /b %ERRORLEVEL%
'@

$bashPath = Join-Path $gitHooks 'pre-commit'
$cmdPath  = Join-Path $gitHooks 'pre-commit.cmd'

Set-Content -Path $bashPath -Value $bashWrapper -NoNewline -Encoding ascii
Set-Content -Path $cmdPath  -Value $cmdWrapper  -NoNewline -Encoding ascii

Write-Host "Hook instalado en: $gitHooks"
Write-Host "Listo. Probá hacer un commit; si hay binarios/secretos bloqueados, el hook lo va a indicar."