Param(
    [string]$Root = '.',
    [string]$DbName = 'DB_SPC_SI.mdb',
    [switch]$FailOnWarnings
)

# Verifica que no existan .exe dentro del repo (excepto quizá en dist/ si se decide en el futuro)
$rootPath = Resolve-Path $Root
$exeFiles = Get-ChildItem -Path $rootPath -Recurse -Include *.exe -ErrorAction SilentlyContinue | Where-Object { $_.FullName -notmatch '\\dist\\' }

# Verifica presencia de la base principal
$dbFiles = Get-ChildItem -Path $rootPath -Recurse -Include $DbName -ErrorAction SilentlyContinue

# Formatos fuente esperados
$sourcePatterns = @('*.vbp','*.bas','*.frm')
$missingSource = @()
foreach ($pat in $sourcePatterns) {
    if (-not (Get-ChildItem -Path $rootPath -Recurse -Include $pat -ErrorAction SilentlyContinue)) {
        $missingSource += $pat
    }
}

Write-Host "--- Verificación Estado Repo ---" -ForegroundColor Cyan
Write-Host "Raíz: $rootPath" -ForegroundColor DarkGray

if ($exeFiles) {
    Write-Warning "Se encontraron ejecutables en el repo (no deberían subirse):"
    $exeFiles | ForEach-Object { Write-Host "  - "$_.FullName -ForegroundColor Yellow }
} else {
    Write-Host "OK: Sin .exe fuera de 'dist/'." -ForegroundColor Green
}

if ($dbFiles) {
    Write-Host "OK: Base principal '$DbName' presente." -ForegroundColor Green
} else {
    Write-Warning "No se encontró '$DbName'. Verifique que no esté ignorada por error."
}

if ($missingSource.Count -gt 0) {
    Write-Warning "Faltan archivos fuente esperados: $($missingSource -join ', ')"
} else {
    Write-Host "OK: Tipos de código fuente presentes (.vbp/.bas/.frm)." -ForegroundColor Green
}

# Sugerencias
Write-Host "--- Sugerencias ---" -ForegroundColor Cyan
if ($exeFiles) { Write-Host "Ejecute 'git rm --cached <archivo.exe>' para retirarlos y luego commit." -ForegroundColor Magenta }
Write-Host "Para generar nueva versión: .\\tools\\Increment-Version.ps1" -ForegroundColor Magenta
Write-Host "Para crear tag: git tag -a vX.Y.Z -m 'Release vX.Y.Z'" -ForegroundColor Magenta

if ($FailOnWarnings -and ($exeFiles -or -not $dbFiles -or $missingSource.Count -gt 0)) {
    Write-Error "Fallando por advertencias detectadas."; exit 1
}
