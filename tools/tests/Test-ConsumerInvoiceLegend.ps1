<#
.SYNOPSIS
    Checks the consumer-final legend on type B invoices printed by FormImpriir.frm
    and reprinted from Buscar factura (FormVerFacturas.frm, FormVerFacturaCtaCte.frm).

.DESCRIPTION
    Extracts IsConsumerFinalCustomer and ImprimirLeyendaConsumidorFinal from each
    VB6 form, runs them in 32-bit MSScriptControl against a simulated Printer, and
    checks the exact legend, the print conditions, the restored printer state, the
    call sites after each QR in GenerarFEB/ImprimirFEB, the A/B button dispatch,
    identical helper copies across forms and the ANSI/CRLF encoding.

    Run with 32-bit Windows PowerShell (ScriptControl is 32-bit only):
    C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\tests\Test-ConsumerInvoiceLegend.ps1
#>
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$ansi = [System.Text.Encoding]::GetEncoding(1252)

# «Orientación al Consumidor Provincia de Buenos Aires 148»
$expectedLegend = [string][char]0xAB + 'Orientaci' + [char]0xF3 + 'n al Consumidor Provincia de Buenos Aires 148' + [char]0xBB

# Expected active call sites per form: PDF routine (original + duplicado) and paper routine.
$forms = @(
    @{ File = 'FormImpriir.frm';          PdfButton = 'cmdGenerarPDF'; PrintButton = 'cmdImprimir';   PdfCalls = 2; PaperCalls = 1 },
    @{ File = 'FormVerFacturas.frm';      PdfButton = 'cmdGenerarFE';  PrintButton = 'cmdImprimirFE'; PdfCalls = 2; PaperCalls = 2 },
    @{ File = 'FormVerFacturaCtaCte.frm'; PdfButton = 'cmdGenerarFE';  PrintButton = 'cmdImprimirFE'; PdfCalls = 2; PaperCalls = 2 }
)

$failures = New-Object System.Collections.Generic.List[string]
function Assert-True([bool]$condition, [string]$message) {
    if ($condition) { Write-Host "PASS $message" } else { Write-Host "FAIL $message"; $failures.Add($message) }
}

function Get-Block([string]$text, [string]$pattern) {
    $match = [regex]::Match($text, $pattern, 'Singleline, IgnoreCase')
    if ($match.Success) { return $match.Value }
    return $null
}

$mock = @'
Class MockPrinter
    Public FontName, FontSize, FontBold, FontItalic, ForeColor, CurrentX, CurrentY, Calls, LastText, LastState
    Private Sub Class_Initialize()
        FontName = "Courier New": FontSize = 12: FontBold = True: FontItalic = True: ForeColor = vbWhite
        CurrentX = 0: CurrentY = 0: Calls = 0: LastText = "": LastState = ""
    End Sub
    Public Sub PrintText(t)
        Calls = Calls + 1
        LastText = t
        LastState = FontName & "|" & FontSize & "|" & B(FontBold) & "|" & B(FontItalic) & "|" & ForeColor & "|" & CurrentX & "|" & CurrentY
    End Sub
    Public Function State()
        State = FontName & "|" & FontSize & "|" & B(FontBold) & "|" & B(FontItalic) & "|" & ForeColor
    End Function
End Class
Function B(v)
    If v Then B = "True" Else B = "False"
End Function
Dim Printer
Function RunCase(tipo, condicion, razon)
    Set Printer = New MockPrinter
    ImprimirLeyendaConsumidorFinal tipo, condicion, razon
    RunCase = Printer.Calls & "#" & Printer.LastText & "#" & Printer.LastState & "#" & Printer.State()
End Function
'@

$printedState = 'Arial|8|False|False|0|15|274'
$restoredState = 'Courier New|12|True|True|16777215'

$cases = @(
    @{ Name = 'B + CF prints the legend';                  Args = @('B', 'CF', 'JUAN PEREZ');        Prints = $true },
    @{ Name = 'B + CONSUMIDOR FINAL name prints the legend'; Args = @('B', 'RI', ' consumidor final '); Prints = $true },
    @{ Name = 'B + lowercase cf prints the legend';        Args = @('B', ' cf ', '');                Prints = $true },
    @{ Name = 'B + Responsable Inscripto does not print';  Args = @('B', 'RI', 'EMPRESA SA');        Prints = $false },
    @{ Name = 'B + Monotributo does not print';            Args = @('B', 'MO', 'EMPRESA SA');        Prints = $false },
    @{ Name = 'A + CF does not print';                     Args = @('A', 'CF', 'CONSUMIDOR FINAL');  Prints = $false },
    @{ Name = 'B + Null customer data does not print';     Args = @('B', [DBNull]::Value, [DBNull]::Value); Prints = $false }
)

function Get-ActiveLines([string]$source, [string]$procedure) {
    $block = Get-Block $source ("Private Sub $procedure\(\).*?\r\nEnd Sub")
    if ($null -eq $block) { return $null }
    return @($block -split "`r`n" | Where-Object { $_.TrimStart() -notmatch "^'" })
}

function Get-ActiveCallCount([string]$source, [string]$procedure) {
    $lines = Get-ActiveLines $source $procedure
    if ($null -eq $lines) { return -1 }
    $count = 0
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match 'ImprimirLeyendaConsumidorFinal\s*\(?\s*tFacturaC!tipofactura') {
            $previous = ($lines[[Math]::Max(0, $i - 6)..($i - 1)] -join "`n")
            if ($previous -match 'Printer\.PaintPicture PicQR\.Picture') { $count++ }
        }
    }
    return $count
}

function Get-Dispatch([string]$source, [string]$procedure) {
    $lines = Get-ActiveLines $source $procedure
    if ($null -eq $lines) { return $null }
    return (($lines | Where-Object { $_ -match '^\s*(If |Call )' } | ForEach-Object { $_.Trim() }) -join ';')
}

$referenceHelpers = $null

foreach ($form in $forms) {
    $name = $form.File
    $bytes = [System.IO.File]::ReadAllBytes((Join-Path $repoRoot $name))
    $source = $ansi.GetString($bytes)

    # --- Encoding ----------------------------------------------------------
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    Assert-True (-not $hasBom) "${name}: no UTF-8 BOM"
    $bareLf = [regex]::Matches($source, '(?<!\r)\n').Count
    Assert-True ($bareLf -eq 0) "${name}: CRLF only"
    Assert-True ($source.Contains('"' + $expectedLegend + '"')) "${name}: legend literal is stored in Windows-1252 with angle quotes"

    # --- Helpers -----------------------------------------------------------
    $functionBlock = Get-Block $source 'Private Function IsConsumerFinalCustomer\(.*?End Function'
    $subBlock = Get-Block $source 'Private Sub ImprimirLeyendaConsumidorFinal\(.*?End Sub'
    Assert-True ($null -ne $functionBlock) "${name}: IsConsumerFinalCustomer exists"
    Assert-True ($null -ne $subBlock) "${name}: ImprimirLeyendaConsumidorFinal exists"

    if ($null -ne $functionBlock -and $null -ne $subBlock) {
        $helpers = $functionBlock + "`r`n" + $subBlock
        if ($null -eq $referenceHelpers) {
            $referenceHelpers = $helpers
        } else {
            Assert-True ($helpers -ceq $referenceHelpers) "${name}: helpers are identical to FormImpriir.frm"
        }

        $vbs = $helpers
        $vbs = $vbs -replace '(?m)^\s*Private\s+', ''
        $vbs = $vbs -replace 'ByVal\s+', ''
        $vbs = $vbs -replace '\s+As\s+\w+', ''
        $vbs = $vbs -replace 'Printer\.Print\s', 'Printer.PrintText '

        $sc = New-Object -ComObject MSScriptControl.ScriptControl
        $sc.Language = 'VBScript'
        $sc.AddCode($mock + "`r`n" + $vbs)

        foreach ($case in $cases) {
            $result = [string]$sc.Run('RunCase', $case.Args[0], $case.Args[1], $case.Args[2])
            $parts = $result.Split('#')
            if ($case.Prints) {
                Assert-True ($parts[0] -eq '1' -and $parts[1] -ceq $expectedLegend) "${name}: $($case.Name): exact text once"
                Assert-True ($parts[2] -eq $printedState) "${name}: $($case.Name): Arial 8 regular black at x15 y274 (got $($parts[2]))"
                Assert-True ($parts[3] -eq $restoredState) "${name}: $($case.Name): font and color restored (got $($parts[3]))"
            } else {
                Assert-True ($parts[0] -eq '0') "${name}: $($case.Name)"
                Assert-True ($parts[3] -eq $restoredState) "${name}: $($case.Name): printer state untouched"
            }
        }
    }

    # --- Call sites --------------------------------------------------------
    $pdf = Get-ActiveCallCount $source 'GenerarFEB'
    $paper = Get-ActiveCallCount $source 'ImprimirFEB'
    $pdfA = Get-ActiveCallCount $source 'GenerarFE'
    $paperA = Get-ActiveCallCount $source 'ImprimirFE'
    Assert-True ($pdf -eq $form.PdfCalls) "${name}: GenerarFEB (PDF) calls the legend after the QR on original and duplicado (got $pdf)"
    Assert-True ($paper -eq $form.PaperCalls) "${name}: ImprimirFEB (paper) calls the legend after the QR on each active page (expected $($form.PaperCalls), got $paper)"
    Assert-True ($pdfA -eq 0) "${name}: GenerarFE (type A PDF) does not call the legend (got $pdfA)"
    Assert-True ($paperA -eq 0) "${name}: ImprimirFE (type A paper) does not call the legend (got $paperA)"
    $expectedTotal = $form.PdfCalls + $form.PaperCalls
    $totalCalls = [regex]::Matches($source, '(?m)^[ \t]*(Call |If .*Then Call )?ImprimirLeyendaConsumidorFinal\s*\(?\s*tFacturaC').Count
    Assert-True ($totalCalls -eq $expectedTotal) "${name}: exactly $expectedTotal active call sites (got $totalCalls)"
    $guarded = [regex]::Matches($source, '(?m)^[ \t]*If Not tClientes\.NoMatch Then Call ImprimirLeyendaConsumidorFinal\(tFacturaC!tipofactura, tClientes!condicionIva, tClientes!RazonSocial\)\r$').Count
    Assert-True ($guarded -eq $expectedTotal) "${name}: every call is guarded by tClientes.NoMatch (got $guarded)"

    # --- Button dispatch ---------------------------------------------------
    $pdfDispatch = Get-Dispatch $source "$($form.PdfButton)_Click"
    $printDispatch = Get-Dispatch $source "$($form.PrintButton)_Click"
    Assert-True ($pdfDispatch -eq 'If TextTipoFactura.text = "A" Then;Call GenerarFE;If TextTipoFactura.text = "B" Then;Call GenerarFEB') "${name}: $($form.PdfButton)_Click sends A to GenerarFE and B to GenerarFEB"
    Assert-True ($printDispatch -eq 'If TextTipoFactura.text = "A" Then;Call ImprimirFE;If TextTipoFactura.text = "B" Then;Call ImprimirFEB') "${name}: $($form.PrintButton)_Click sends A to ImprimirFE and B to ImprimirFEB"
}

if ($failures.Count -gt 0) {
    Write-Host "`n$($failures.Count) check(s) failed."
    exit 1
}
Write-Host "`nAll checks passed."
exit 0
