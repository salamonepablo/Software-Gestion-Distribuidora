<#!
Script: Clean-Binaries.ps1
Descripción: Elimina del índice (sin borrar local) ejecutables, librerías, certificados y otros binarios que no deben versionarse.
Uso:
  powershell -ExecutionPolicy Bypass -File .\tools\Clean-Binaries.ps1
Luego revisar con: git status
#>
Param([switch]$DryRun)

$patterns = @('*.exe','*.dll','*.ocx','*.pfx','*.crt','*.key')
$tracked = @()
foreach ($p in $patterns) {
  $files = git ls-files $p 2>$null
  if ($files) { $tracked += $files }
}

if (-not $tracked) { Write-Host "No hay binarios/secretos rastreados." -ForegroundColor Green; exit 0 }

Write-Host "Encontrados $(($tracked).Count) archivos rastreados a limpiar:" -ForegroundColor Yellow
$tracked | ForEach-Object { Write-Host "  - $_" }

if ($DryRun) { Write-Host "DryRun: no se ejecuta git rm." -ForegroundColor Cyan; exit 0 }

$tracked | ForEach-Object { git rm --cached $_ }

Write-Host "Limpieza completada. Ejecuta 'git add .' si se modificó .gitignore y luego commit." -ForegroundColor Green
