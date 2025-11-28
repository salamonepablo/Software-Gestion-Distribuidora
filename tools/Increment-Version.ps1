Param(
    [string]$ProjectFile = "./SPCSI.vbp",
    [ValidateSet("revision","minor","major")][string]$Increment = "revision",
    [string]$SetVersion,
    [switch]$UpdateExeName
)

<#!
Script: Increment-Version.ps1
Descripción: Incrementa o fija la versión (MAJOR.MINOR.PATCH) en archivo .vbp.
- MAJOR -> MajorVer
- MINOR -> MinorVer
- PATCH -> RevisionVer

Uso típico:
  powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1
  powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -Increment minor
  powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -SetVersion 2.0.0
  powershell -ExecutionPolicy Bypass -File .\tools\Increment-Version.ps1 -SetVersion 1.4.3 -UpdateExeName

Si -UpdateExeName se indica, renombra ExeName32 para incluir sufijo _vX_Y_Z
#>

if (-not (Test-Path $ProjectFile)) {
    Write-Error "No se encontró el archivo de proyecto: $ProjectFile"; exit 1
}

$content = Get-Content -Raw -Path $ProjectFile -Encoding Default

# Extraer valores actuales
$major = [int](([regex]::Match($content, "(?m)^MajorVer=(\d+)").Groups[1].Value)  )
$minor = [int](([regex]::Match($content, "(?m)^MinorVer=(\d+)").Groups[1].Value)  )
$patch = [int](([regex]::Match($content, "(?m)^RevisionVer=(\d+)").Groups[1].Value)  )

$oldVersion = "$major.$minor.$patch"

if ($SetVersion) {
    if ($SetVersion -notmatch '^(\d+)\.(\d+)\.(\d+)$') { Write-Error "Formato inválido en -SetVersion. Use MAJOR.MINOR.PATCH"; exit 1 }
    $parts = $SetVersion.Split('.')
    $major = [int]$parts[0]
    $minor = [int]$parts[1]
    $patch = [int]$parts[2]
}
else {
    switch ($Increment) {
        'major'   { $major += 1; $minor = 0; $patch = 0 }
        'minor'   { $minor += 1; $patch = 0 }
        'revision'{ $patch += 1 }
    }
}

$newVersion = "$major.$minor.$patch"

# Reemplazar valores en contenido
$content = [regex]::Replace($content, "(?m)^MajorVer=\d+", "MajorVer=$major")
$content = [regex]::Replace($content, "(?m)^MinorVer=\d+", "MinorVer=$minor")
$content = [regex]::Replace($content, "(?m)^RevisionVer=\d+", "RevisionVer=$patch")

if ($UpdateExeName) {
    # Buscar ExeName32="Nombre.exe" y actualizar sufijo
    $exeMatch = [regex]::Match($content, 'ExeName32="([A-Za-z0-9_\-]+)\.exe"')
    if ($exeMatch.Success) {
        $baseName = $exeMatch.Groups[1].Value -replace '_v\d+_\d+_\d+$',''  # quitar version previa si existe
        $newExe = "${baseName}_v$($major)_$($minor)_$($patch).exe"
        $content = [regex]::Replace($content, 'ExeName32="[A-Za-z0-9_\-]+\.exe"', 'ExeName32="' + $newExe + '"')
    }
}

# Guardar (mantener encoding ANSI por compatibilidad VB6)
[System.IO.File]::WriteAllText((Resolve-Path $ProjectFile), $content, (New-Object System.Text.ASCIIEncoding))

Write-Host "Versión anterior: $oldVersion" -ForegroundColor Yellow
Write-Host "Nueva versión:    $newVersion" -ForegroundColor Green
if ($UpdateExeName) { Write-Host "ExeName32 actualizado (si coincidía)" -ForegroundColor Cyan }

# Sugerencia de tag
Write-Host "Ejecute: git commit -am 'Bump version to v$newVersion'; git tag -a v$newVersion -m 'Release v$newVersion'" -ForegroundColor Magenta
