<#
.SYNOPSIS
    Checks that CrearQR writes the QR image where the print forms read it.

.DESCRIPTION
    FormImpriir.frm and FormImprimirNC.frm load the QR from App.Path\QRs\. This
    test extracts the real CrearQR from VariablesPublicas.bas, translates it to
    VBScript and runs it in 32-bit MSScriptControl against a mocked App and a
    mocked FEAFIPLib.Qr whose RutaLibreria differs from App.Path. The mock
    Generar only logs its arguments: no DLL, network, database or printer is used.

    Each case prints a "CASE" line with the resulting ArchivoQR and Generar
    payload so the output of two source versions can be compared.

    Run with 32-bit Windows PowerShell (ScriptControl is 32-bit only):
    C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\tests\Test-QrPath.ps1
#>
param([string]$SourcePath)
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
if (-not $SourcePath) { $SourcePath = Join-Path $repoRoot 'VariablesPublicas.bas' }
$ansi = [System.Text.Encoding]::GetEncoding(1252)
$retailPath = 'C:\Trabajos Activos\SPC-Retail'
$libraryPath = 'C:\Trabajos Activos\SPC-Core'

$failures = New-Object System.Collections.Generic.List[string]
function Assert-True([bool]$condition, [string]$message) {
    if ($condition) { Write-Host "PASS $message" } else { Write-Host "FAIL $message"; $failures.Add($message) }
}
function Assert-Equal($expected, $actual, [string]$message) {
    Assert-True ([string]$expected -ceq [string]$actual) "$message (expected '$expected', got '$actual')"
}

$module = $ansi.GetString([System.IO.File]::ReadAllBytes($SourcePath))
$crearQr = [regex]::Match($module, '(?ms)^Public Sub CrearQR\(.*?^End Sub\r?$').Value
Assert-True ($crearQr -ne '') 'CrearQR found in source'
if ($crearQr -eq '') { exit 1 }

# VB6 -> VBScript for the subset used by CrearQR.
$code = $crearQr -replace '(?m)^(\s*)Public\s+', '$1'
$code = $code -replace 'New FEAFIPLib\.Qr', 'New MockQr'
$code = $code -replace '\sAs\s+[A-Za-z_][\w.]*', ''
$code = $code -replace '(\d)#', '$1'
$code = $code -replace '\bMsgBox\b', 'MockMsgBox'

$mock = @'
Dim MockLibraryPath, GenerarLog, MsgLog, Ver
SetLocale 1033 ' payload log uses "." as decimal separator
Class MockAppClass
    Public Path
End Class
Dim App
Set App = New MockAppClass
Class MockQr
    Public ArchivoQR, ErrorDesc
    Public Property Get RutaLibreria
        RutaLibreria = MockLibraryPath
    End Property
    Public Function Generar(v, f, c, p, t, n, i, m, z, td, nd, tc, ca)
        GenerarLog = ArchivoQR & "|" & Join(Array(v, f, c, p, t, n, i, m, z, td, nd, tc, ca), "|")
        Generar = True
    End Function
End Class
Sub MockMsgBox(text)
    MsgLog = MsgLog & text & ";"
End Sub
'@

$sc = New-Object -ComObject MSScriptControl.ScriptControl
$sc.Language = 'VBScript'
$sc.AddCode($mock + "`r`n" + $code)

function Invoke-CrearQr([string]$appPath, [string]$libPath, [string]$callArgs) {
    $sc.ExecuteStatement("App.Path = `"$appPath`" : MockLibraryPath = `"$libPath`" : GenerarLog = `"`" : MsgLog = `"`"")
    $sc.ExecuteStatement("Call CrearQR($callArgs)")
    return [string]$sc.Eval('GenerarLog'), [string]$sc.Eval('MsgLog')
}

# Supported comprobantes: factura A/B (FormImpriir) and nota de credito A/B (FormImprimirNC).
$cases = @(
    @{ Name = 'FB 11944'; Tipo = 6; Prefix = 'FB'; Nro = 11944; Importe = '12345.67'; Doc = '0' },
    @{ Name = 'FA 3021'; Tipo = 1; Prefix = 'FA'; Nro = 3021; Importe = '1000.5'; Doc = '27929007862' },
    @{ Name = 'NCA 120'; Tipo = 3; Prefix = 'NCA'; Nro = 120; Importe = '250'; Doc = '30708432543' },
    @{ Name = 'NCB 845'; Tipo = 8; Prefix = 'NCB'; Nro = 845; Importe = '99.99'; Doc = '0' }
)
$scenarios = @(
    @{ Name = 'library differs'; Lib = $libraryPath },
    @{ Name = 'library same'; Lib = $retailPath }
)

foreach ($case in $cases) {
    foreach ($scenario in $scenarios) {
        $callArgs = "`"28/10/2026`", CDbl(30708432543), 4, $($case.Tipo), CDbl($($case.Nro)), CDbl($($case.Importe)), `"PES`", 1, 80, CDbl($($case.Doc)), `"A`", CDbl(75123456789012)"
        $log, $msg = Invoke-CrearQr $retailPath $scenario.Lib $callArgs
        Write-Host "CASE $($case.Name) / $($scenario.Name): $log"
        $expectedFile = "$retailPath\QRs\qr_$($case.Prefix)_4_$($case.Nro).jpg"
        $expectedPayload = "1|28/10/2026|30708432543|4|$($case.Tipo)|$($case.Nro)|$($case.Importe)|PES|1|80|$($case.Doc)|E|75123456789012"
        Assert-Equal "$expectedFile|$expectedPayload" $log "$($case.Name) / $($scenario.Name) writes QR under App.Path with unchanged payload"
        Assert-Equal '' $msg "$($case.Name) / $($scenario.Name) shows no error"
    }
}

if ($failures.Count -gt 0) {
    Write-Host "`n$($failures.Count) check(s) failed."
    exit 1
}
Write-Host "`nAll checks passed."
exit 0
