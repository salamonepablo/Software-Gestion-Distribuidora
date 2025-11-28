#!/usr/bin/env pwsh
# Hook pre-commit: bloquea binarios/secretos no deseados
# Copiar a .git/hooks/pre-commit para activarlo

$patterns = @('*.exe','*.dll','*.ocx','*.pfx','*.crt','*.key')
$staged = git diff --cached --name-only
$matches = @()
foreach ($pat in $patterns) {
  $regex = '^' + ([regex]::Escape($pat).Replace('\*','.*')) + '$'
  $matches += ($staged | Where-Object { $_ -match $regex })
}
$matches = $matches | Sort-Object -Unique
if ($matches.Count -gt 0) {
  Write-Host "ERROR: Archivos binarios/secretos detectados en staging:" -ForegroundColor Red
  $matches | ForEach-Object { Write-Host "  - $_" -ForegroundColor Yellow }
  Write-Host "Use 'git rm --cached <archivo>' o ejecuta .\\tools\\Clean-Binaries.ps1" -ForegroundColor Red
  exit 1
}
exit 0
