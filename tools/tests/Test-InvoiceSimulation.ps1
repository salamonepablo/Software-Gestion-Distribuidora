<#
.SYNOPSIS
    Checks the IDE-only invoice simulation mode (facturacion de prueba sin ARCA).

.DESCRIPTION
    Extracts the simulation procedures from VariablesPublicas.bas (guard, enable,
    SimulateInvoiceAuthorization, IsInvoiceSimulationCae, the real ActualizarCAE
    and the start of FacturaElectronicaSPC up to the ARCA code) and the activation
    procedures from FormFactura.frm, translates them to VBScript and runs them in
    32-bit MSScriptControl against a mocked App, Debug.Assert, MsgBox, ARCA call
    and an in-memory DAO database. No real database, printer or ARCA is used.

    Also checks the form designer flags, the PRUEBA marker on every active CAE
    print in FormImpriir.frm and FormImprimirNC.frm, and the ANSI/CRLF encoding.

    Run with 32-bit Windows PowerShell (ScriptControl is 32-bit only):
    C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\tests\Test-InvoiceSimulation.ps1
#>
$ErrorActionPreference = 'Stop'

$repoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
$ansi = [System.Text.Encoding]::GetEncoding(1252)
$retailPath = 'C:\Trabajos Activos\SPC-Retail'
$simulationCae = '99999999999999'
$marker = 'PRUEBA - SIN VALIDEZ FISCAL'
$confirmText = 'Activar facturacion SIN ARCA sobre la copia local? Se modificaran stock y cuenta corriente. Para salir del modo de prueba debe detener y reiniciar el proyecto.'
$debitNoteMessage = 'El modo de prueba no autoriza notas de debito: solo facturas y notas de credito A y B. El comprobante local puede haber quedado grabado sin CAE.'
$unavailableText = 'El modo de prueba requiere ejecutar este proyecto desde el IDE en C:\Trabajos Activos\SPC-Retail.'

$failures = New-Object System.Collections.Generic.List[string]
function Assert-True([bool]$condition, [string]$message) {
    if ($condition) { Write-Host "PASS $message" } else { Write-Host "FAIL $message"; $failures.Add($message) }
}
function Assert-Equal($expected, $actual, [string]$message) {
    Assert-True ([string]$expected -ceq [string]$actual) "$message (expected '$expected', got '$actual')"
}

function Read-Source([string]$name) {
    $path = Join-Path $repoRoot $name
    $bytes = [System.IO.File]::ReadAllBytes($path)
    $text = $ansi.GetString($bytes)
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    Assert-True (-not $hasBom) "$name has no UTF-8 BOM"
    Assert-True ([regex]::Matches($text, '(?<!\r)\n').Count -eq 0) "$name uses CRLF only"
    return $text
}

function Get-Block([string]$text, [string]$pattern) {
    $match = [regex]::Match($text, $pattern, 'Singleline')
    if ($match.Success) { return $match.Value }
    return $null
}

# VB6 -> VBScript for the subset used by the simulation procedures.
function ConvertTo-Vbs([string]$code) {
    $code = $code -replace '(?m)^(\s*)(Private|Public)\s+', '$1'
    $code = $code -replace '(?m)^\s*On Error GoTo \w+\s*$', ''
    $code = $code -replace '(?m)^\w+:\s*$', ''
    $code = $code -replace 'Debug\.Assert\s+', 'If MockIde Then Call '
    $code = $code -replace 'Optional\s+', ''
    $code = $code -replace '\sAs\s+[A-Za-z_][\w.]*', ''
    $code = $code -replace '(\d)#', '$1'
    $code = $code -replace '(\w+)!(\w+)', '$1.Item("$2")'
    $code = $code -replace '\bMsgBox\b', 'MockMsgBox'
    $code = $code -replace '\bMe\.', 'MockForm.'
    return $code
}

$module = Read-Source 'VariablesPublicas.bas'
$factura = Read-Source 'FormFactura.frm'
$impriir = Read-Source 'FormImpriir.frm'
$imprimirNC = Read-Source 'FormImprimirNC.frm'

# --- Module structure ------------------------------------------------------
$declarations = @(
    [regex]::Match($module, '(?m)^Private invoiceSimulationEnabled As Boolean\r$').Value,
    [regex]::Match($module, '(?m)^Public Const InvoiceSimulationCae As Double = 99999999999999#\r$').Value
)
Assert-True ($declarations[0] -ne '') 'invoiceSimulationEnabled is a private module flag'
Assert-True ($declarations[1] -ne '') 'InvoiceSimulationCae constant is 99999999999999'

$names = @(
    'Private Function DetectInvoiceIde\(',
    'Public Function CanSimulateInvoices\(',
    'Public Function IsInvoiceSimulationEnabled\(',
    'Public Function EnableInvoiceSimulation\(',
    'Private Function SimulateInvoiceAuthorization\(',
    'Public Function IsInvoiceSimulationCae\(',
    'Public Function ActualizarCAE\('
)
$blocks = @()
foreach ($name in $names) {
    $block = Get-Block $module ($name + '.*?\r\nEnd Function')
    $label = $name.Substring(0, $name.Length - 2)
    Assert-True ($null -ne $block) "$label exists"
    $blocks += $block
}

$facturaElectronica = Get-Block $module 'Public Function FacturaElectronicaSPC\(.*?\r\nEnd Function'
$beforeArca = $null
if ($null -ne $facturaElectronica) {
    $beforeArca = Get-Block $facturaElectronica '^.*?(?=\r\n[^\r\n]*On Error GoTo CapErr)'
    $simulationAt = $facturaElectronica.IndexOf('SimulateInvoiceAuthorization(TipoComp, CbteDesde)')
    $arcaAt = $facturaElectronica.IndexOf('wsfev1')
    Assert-True ($simulationAt -gt 0 -and $simulationAt -lt $arcaAt -and $null -ne $beforeArca -and $beforeArca.Contains('Exit Function')) 'FacturaElectronicaSPC exits to the simulation before any ARCA code'
}

$simulateBlock = $blocks[4]
if ($null -ne $simulateBlock) {
    $handler = Get-Block $simulateBlock 'Failed:.*'
    Assert-True ($null -ne $handler -and $handler.Contains('MsgBox "No se pudo simular la autorizacion: " & Err.Description') -and -not $handler.Contains('= True')) 'SimulateInvoiceAuthorization error handler warns and returns False'
    Assert-True ($simulateBlock.Contains('Call ActualizarCAE(documentTable, documentTypeCode, documentNumber, InvoiceSimulationCae,')) 'Simulation writes the CAE through the existing ActualizarCAE'
}

# --- Form structure --------------------------------------------------------
$formHeader = Get-Block $factura '^.*?(?=\r\n   Begin )'
Assert-True ($null -ne $formHeader -and $formHeader -match "\r\n   KeyPreview      =   -1  'True\r\n") 'FormFactura has KeyPreview enabled'
$button = Get-Block $factura '\r\n( +)Begin VB\.CommandButton cmdInvoiceSimulation\s*\r\n.*?\r\n\1End\r\n'
Assert-True ($null -ne $button -and $button -match "Visible\s+=\s+0\s+'False" -and $button -match "TabStop\s+=\s+0\s+'False") 'cmdInvoiceSimulation is hidden and outside the tab order'

# Layout: the button must sit in free space, inside its container and clear of every
# other control or frame, including design-hidden ones the form may show at runtime.
function Get-FormControls([string]$text) {
    $designer = Get-Block $text '^.*?\r\nAttribute VB_Name'
    $stack = New-Object System.Collections.ArrayList
    $controls = New-Object System.Collections.ArrayList
    foreach ($line in ($designer -split "`r`n")) {
        if ($line -match '^\s*Begin (\S+) (\S+)') {
            [void]$stack.Add(@{ Type = $Matches[1]; Name = $Matches[2]; Left = 0; Top = 0; Width = 0; Height = 0; Visible = $true; Path = @($stack | ForEach-Object { $_ }) })
        } elseif ($line -match '^\s*End$' -and $stack.Count -gt 0) {
            $node = $stack[$stack.Count - 1]
            $stack.RemoveAt($stack.Count - 1)
            $absLeft = $node.Left; $absTop = $node.Top
            foreach ($ancestor in $node.Path) { if ($ancestor.Type -ne 'VB.Form') { $absLeft += $ancestor.Left; $absTop += $ancestor.Top } }
            $node.AbsLeft = $absLeft; $node.AbsTop = $absTop
            $node.Parent = if ($node.Path.Count -gt 0) { $node.Path[$node.Path.Count - 1] } else { $null }
            [void]$controls.Add($node)
        } elseif ($stack.Count -gt 0 -and $line -match '^\s*(Left|Top|Width|Height)\s*=\s*(-?\d+)') {
            $stack[$stack.Count - 1][$Matches[1]] = [int]$Matches[2]
        } elseif ($stack.Count -gt 0 -and $line -match "^\s*Visible\s*=\s*0\s") {
            $stack[$stack.Count - 1].Visible = $false
        }
    }
    return $controls
}
$controls = Get-FormControls $factura
$simulationButton = $controls | Where-Object { $_.Name -eq 'cmdInvoiceSimulation' }
Assert-True ($null -ne $simulationButton) 'cmdInvoiceSimulation is in the designer layout'
if ($null -ne $simulationButton) {
    $parent = $simulationButton.Parent
    $inside = $parent.Type -eq 'VB.Form' -or ($simulationButton.Left -ge 0 -and $simulationButton.Top -ge 0 -and
        $simulationButton.Left + $simulationButton.Width -le $parent.Width -and $simulationButton.Top + $simulationButton.Height -le $parent.Height)
    Assert-True $inside "cmdInvoiceSimulation fits inside $($parent.Name)"
    $ancestors = @($simulationButton.Path | ForEach-Object { $_.Name })
    $overlaps = @()
    foreach ($control in $controls) {
        if ($control.Type -eq 'VB.Form' -or $control.Name -eq 'cmdInvoiceSimulation' -or $ancestors -contains $control.Name) { continue }
        if ($simulationButton.AbsLeft -lt $control.AbsLeft + $control.Width -and $control.AbsLeft -lt $simulationButton.AbsLeft + $simulationButton.Width -and
            $simulationButton.AbsTop -lt $control.AbsTop + $control.Height -and $control.AbsTop -lt $simulationButton.AbsTop + $simulationButton.Height) {
            $overlaps += $control.Name
        }
    }
    Assert-True ($overlaps.Count -eq 0) "cmdInvoiceSimulation overlaps no other control (got: $($overlaps -join ', '))"
    $frame4Inputs = @($controls | Where-Object { @('TextApellidoNombre', 'txtOrdenCompra', 'Label29') -contains $_.Name })
    Assert-True ($frame4Inputs.Count -eq 3) 'Frame4 customer controls are present for the layout check'
}

$formNames = @('Private Sub Form_KeyDown\(', 'Private Sub cmdInvoiceSimulation_Click\(', 'Private Sub UpdateInvoiceSimulationNotice\(')
$formBlocks = @()
foreach ($name in $formNames) {
    $block = Get-Block $factura ($name + '.*?\r\nEnd Sub')
    $label = $name.Substring(0, $name.Length - 2)
    Assert-True ($null -ne $block) "$label exists in FormFactura"
    $formBlocks += $block
}
$loadPrefix = Get-Block $factura 'Private Sub Form_Load\(\)\r\n.*?\r\n    UpdateInvoiceSimulationNotice\r\n'
Assert-True ($null -ne $loadPrefix -and -not $loadPrefix.Contains('FormFactura.Height')) 'Form_Load shows the simulation state before its existing setup'
Assert-True ($factura.Contains("Private Sub Form_Load()`r`n    If CanSimulateInvoices() Then`r`n        cmdInvoiceSimulation.Visible = True`r`n        cmdInvoiceSimulation.ZOrder 0`r`n    End If`r`n    UpdateInvoiceSimulationNotice`r`n`r`n    FormFactura.Height = 10200")) 'Form_Load keeps the existing setup right after the simulation lines'

# --- Runtime ---------------------------------------------------------------
$mock = @'
Const vbKeyF12 = 123
Const vbShiftMask = 1
Const vbCtrlMask = 2
Const dbOpenTable = 1
Dim MockIde, App, BaseSPC, MsgLog, MsgButtons, MsgAnswer, AfipCalls, AfipResult, MockForm, cmdInvoiceSimulation
Class MockApp
    Public Path
End Class
Class MockButton
    Public Visible, Caption, Enabled, ZOrderCalls
    Private Sub Class_Initialize()
        Visible = False: Caption = "Activar facturacion de prueba": Enabled = True: ZOrderCalls = 0
    End Sub
    Public Sub ZOrder(position)
        If position = 0 Then ZOrderCalls = ZOrderCalls + 1
    End Sub
End Class
Class MockFormState
    Public Caption
End Class
Class MockDatabase
    Public Records, Opens, Writes, FailOpen
    Private Sub Class_Initialize()
        Set Records = CreateObject("Scripting.Dictionary")
        Opens = 0: Writes = 0: FailOpen = False
    End Sub
    Public Function OpenRecordset(tableName, openType)
        If FailOpen Then Err.Raise 3024, , "No se pudo abrir la base"
        Opens = Opens + 1
        Dim rs
        Set rs = New MockRecordset
        Set rs.Owner = Me
        rs.TableName = tableName
        Set OpenRecordset = rs
    End Function
End Class
Class MockRecordset
    Public Owner, TableName, Index, NoMatch, CurrentKey, Editing, Closed
    Public Sub Seek(comparison, tipo, numero)
        CurrentKey = TableName & "|" & tipo & "|" & CStr(numero)
        NoMatch = Not Owner.Records.Exists(CurrentKey)
    End Sub
    Public Property Get Item(fieldName)
        Item = Owner.Records(CurrentKey)(fieldName)
    End Property
    Public Property Let Item(fieldName, value)
        If Not Editing Then Err.Raise 3020, , "Update without Edit"
        Owner.Records(CurrentKey)(fieldName) = value
    End Property
    Public Sub Edit()
        Editing = True
    End Sub
    Public Sub Update()
        Editing = False
        Owner.Writes = Owner.Writes + 1
    End Sub
    Public Sub Close()
        Closed = True
    End Sub
End Class
' VB6 Boolean defaults to False and prints in English; VBScript may be Empty and localized.
Function B(value)
    If value Then B = "True" Else B = "False"
End Function
Function Format(value, pattern)
    If LCase(pattern) = "yyyymmdd" Then
        Format = Year(value) & Right("0" & Month(value), 2) & Right("0" & Day(value), 2)
    Else
        Format = value
    End If
End Function
Function MockMsgBox(prompt, buttons, title)
    MsgLog = MsgLog & title & ">" & prompt & ";"
    MsgButtons = buttons
    MockMsgBox = MsgAnswer
End Function
Function MockAfip(tipo, numero)
    AfipCalls = AfipCalls + 1
    MockAfip = AfipResult
End Function
Sub Reset(ide, path)
    MockIde = ide
    Set App = New MockApp
    App.Path = path
    Set BaseSPC = New MockDatabase
    Set MockForm = New MockFormState
    MockForm.Caption = "Facturas"
    Set cmdInvoiceSimulation = New MockButton
    MsgLog = "": MsgButtons = 0: MsgAnswer = vbNo: AfipCalls = 0: AfipResult = True
    invoiceSimulationEnabled = False
End Sub
Sub AddDocument(tableName, tipo, numero)
    Dim fields
    Set fields = CreateObject("Scripting.Dictionary")
    fields("CAE") = 0
    fields("FechaVC") = "sin CAE"
    BaseSPC.Records.Add tableName & "|" & tipo & "|" & CStr(numero), fields
End Sub
Function Field(tableName, tipo, numero, fieldName)
    Field = CStr(BaseSPC.Records(tableName & "|" & tipo & "|" & CStr(numero))(fieldName))
End Function
Function Invoice(tipoComp, numero)
    Invoice = FacturaElectronicaSPC(4, 99, 0, tipoComp, numero, numero, "20261029", 121, 100, "PES", 1, 5, 100, 21, 2, "IIBB", 100, 0, 0, 0, 0, 0, "")
End Function
Function Outcome(result)
    Outcome = B(result) & "|afip=" & AfipCalls & "|opens=" & BaseSPC.Opens & "|writes=" & BaseSPC.Writes & "|enabled=" & B(IsInvoiceSimulationEnabled())
End Function
Function PressKey(keyCode, shift)
    Dim code
    code = keyCode
    Form_KeyDown code, shift
    PressKey = code & "|" & B(cmdInvoiceSimulation.Visible) & "|" & cmdInvoiceSimulation.ZOrderCalls & "|" & MsgLog
End Function
Function FormState()
    FormState = MockForm.Caption & "|" & B(cmdInvoiceSimulation.Visible) & "|" & cmdInvoiceSimulation.Caption & "|" & B(cmdInvoiceSimulation.Enabled) & "|" & B(IsInvoiceSimulationEnabled())
End Function
'@

$vbsReady = ($blocks -notcontains $null) -and ($formBlocks -notcontains $null) -and $null -ne $beforeArca -and $null -ne $loadPrefix -and $declarations -notcontains ''
Assert-True $vbsReady 'All simulation procedures are available for runtime checks'

if ($vbsReady) {
    # VB6 jumps to Failed: on any error; emulate it by running the body under a wrapper.
    $simulateBody = $simulateBlock -replace '(?s)\r\nFailed:.*?(?=\r\nEnd Function)', ''
    $simulateBody = $simulateBody -replace 'Function SimulateInvoiceAuthorization\(', 'Function SimulateInvoiceAuthorizationBody('
    $simulateBody = $simulateBody -replace 'SimulateInvoiceAuthorization = True', 'SimulateInvoiceAuthorizationBody = True'
    $simulateWrapper = @'
Function SimulateInvoiceAuthorization(documentType, documentNumber)
    Dim description
    On Error Resume Next
    SimulateInvoiceAuthorization = SimulateInvoiceAuthorizationBody(documentType, documentNumber)
    If Err.Number <> 0 Then
        description = Err.Description
        Err.Clear
        SimulateInvoiceAuthorization = False
        MockMsgBox "No se pudo simular la autorizacion: " & description, vbExclamation, "FACTURACION DE PRUEBA"
    End If
End Function
'@
    $moduleCode = @(
        'Dim invoiceSimulationEnabled',
        'Const InvoiceSimulationCae = 99999999999999',
        $blocks[0], $blocks[1], $blocks[2], $blocks[3], $simulateBody, $simulateWrapper, $blocks[5], $blocks[6],
        ($beforeArca + "`r`n    FacturaElectronicaSPC = MockAfip(TipoComp, CbteDesde)`r`nEnd Function"),
        $formBlocks[0], $formBlocks[1], $formBlocks[2],
        ($loadPrefix + 'End Sub')
    ) -join "`r`n"

    $sc = New-Object -ComObject MSScriptControl.ScriptControl
    $sc.Language = 'VBScript'
    $sc.AllowUI = $false
    $sc.AddCode($mock + "`r`n" + (ConvertTo-Vbs $moduleCode))
    function Run([string]$procedure) {
        switch ($args.Count) {
            0 { return [string]$sc.Run($procedure) }
            1 { return [string]$sc.Run($procedure, $args[0]) }
            2 { return [string]$sc.Run($procedure, $args[0], $args[1]) }
            3 { return [string]$sc.Run($procedure, $args[0], $args[1], $args[2]) }
            4 { return [string]$sc.Run($procedure, $args[0], $args[1], $args[2], $args[3]) }
        }
    }
    function Eval([string]$expression) { return [string]$sc.Eval($expression) }

    $expectedExpiry = (Get-Date).Date.AddDays(10).ToString('dd/MM/yyyy', [System.Globalization.CultureInfo]::InvariantCulture)

    # Guard
    Run 'Reset' $true $retailPath | Out-Null
    Assert-Equal 'True' (Eval 'B(CanSimulateInvoices())') 'IDE in the Retail folder can simulate'
    Run 'Reset' $true 'c:\trabajos activos\spc-retail' | Out-Null
    Assert-Equal 'True' (Eval 'B(CanSimulateInvoices())') 'Folder comparison ignores letter case'
    foreach ($case in @(
        @{ Ide = $false; Path = $retailPath; Name = 'Compiled executable' },
        @{ Ide = $true; Path = 'C:\Trabajos Activos\SPC-Core'; Name = 'IDE in the Core folder' },
        @{ Ide = $true; Path = 'C:\SPC-Retail'; Name = 'IDE in an installed folder' },
        @{ Ide = $true; Path = 'C:\Trabajos Activos\SPC-Retail\Copia'; Name = 'IDE in a Retail subfolder' }
    )) {
        Run 'Reset' $case.Ide $case.Path | Out-Null
        Assert-Equal 'False' (Eval 'B(CanSimulateInvoices())') "$($case.Name) cannot simulate"
        Assert-Equal 'False' (Eval 'B(EnableInvoiceSimulation())') "$($case.Name) cannot enable simulation"
        Assert-Equal 'True|afip=1|opens=0|writes=0|enabled=False' (Run 'Outcome' (Eval 'Invoice(6, 11940)')) "$($case.Name) keeps calling ARCA"
    }

    # Disabled mode keeps the ARCA path and its result
    foreach ($afipResult in @($true, $false)) {
        Run 'Reset' $true $retailPath | Out-Null
        $sc.ExecuteStatement("AfipResult = $afipResult")
        Run 'AddDocument' 'FacturaC' 'B' 11940 | Out-Null
        Assert-Equal "$afipResult|afip=1|opens=0|writes=0|enabled=False" (Run 'Outcome' (Eval 'Invoice(6, 11940)')) "Disabled mode calls ARCA once and returns its result ($afipResult)"
        Assert-Equal '0' (Run 'Field' 'FacturaC' 'B' 11940 'CAE') 'Disabled mode leaves the CAE untouched'
    }

    # Enabled mode bypasses ARCA and stores the fake CAE
    foreach ($case in @(
        @{ Tipo = 1; Table = 'FacturaC'; Letter = 'A' },
        @{ Tipo = 6; Table = 'FacturaC'; Letter = 'B' },
        @{ Tipo = 3; Table = 'NotaCreditoC'; Letter = 'A' },
        @{ Tipo = 8; Table = 'NotaCreditoC'; Letter = 'B' }
    )) {
        Run 'Reset' $true $retailPath | Out-Null
        Run 'AddDocument' $case.Table $case.Letter 500 | Out-Null
        Run 'AddDocument' $case.Table $case.Letter 501 | Out-Null
        Assert-Equal 'True' (Eval 'B(EnableInvoiceSimulation())') "Type $($case.Tipo): enable in the Retail IDE"
        Assert-Equal 'True|afip=0|opens=2|writes=1|enabled=True' (Run 'Outcome' (Eval "Invoice($($case.Tipo), 500)")) "Type $($case.Tipo): authorized without ARCA, one write"
        Assert-Equal $simulationCae (Run 'Field' $case.Table $case.Letter 500 'CAE') "Type $($case.Tipo): $($case.Table) $($case.Letter) 500 gets the fake CAE"
        Assert-Equal $expectedExpiry (Run 'Field' $case.Table $case.Letter 500 'FechaVC') "Type $($case.Tipo): CAE expiry is today + 10 days"
        Assert-Equal '0|sin CAE' ((Run 'Field' $case.Table $case.Letter 501 'CAE') + '|' + (Run 'Field' $case.Table $case.Letter 501 'FechaVC')) "Type $($case.Tipo): other documents untouched"
        Assert-Equal '' (Eval 'MsgLog') "Type $($case.Tipo): no message box"
    }

    # Enabled mode fails closed
    $prefix = 'FACTURACION DE PRUEBA>No se pudo simular la autorizacion: '
    foreach ($case in @(
        @{ Name = 'Nota de debito A'; Tipo = 2; Setup = 'AddDocument "NotaDebitoC", "A", 700'; Opens = 0; Message = $debitNoteMessage },
        @{ Name = 'Nota de debito B'; Tipo = 7; Setup = 'AddDocument "NotaDebitoC", "B", 700'; Opens = 0; Message = $debitNoteMessage },
        @{ Name = 'Missing invoice'; Tipo = 6; Setup = 'AddDocument "FacturaC", "A", 700'; Opens = 1; Message = 'No se encontro el comprobante para asignar el CAE de prueba.' },
        @{ Name = 'Guard lost after enabling'; Tipo = 6; Setup = 'AddDocument "FacturaC", "B", 700 : MockIde = False'; Opens = 0; Message = 'Entorno de prueba no disponible.' },
        @{ Name = 'Database error'; Tipo = 6; Setup = 'AddDocument "FacturaC", "B", 700 : BaseSPC.FailOpen = True'; Opens = 0; Message = 'No se pudo abrir la base' }
    )) {
        Run 'Reset' $true $retailPath | Out-Null
        Assert-Equal 'True' (Eval 'B(EnableInvoiceSimulation())') "$($case.Name): enable in the Retail IDE"
        $sc.ExecuteStatement($case.Setup)
        Assert-Equal "False|afip=0|opens=$($case.Opens)|writes=0|enabled=True" (Run 'Outcome' (Eval "Invoice($($case.Tipo), 700)")) "$($case.Name): rejected without ARCA or writes"
        Assert-Equal ($prefix + $case.Message + ';') (Eval 'MsgLog') "$($case.Name): warns the user"
    }

    # FormFactura activation
    Run 'Reset' $true $retailPath | Out-Null
    Run 'Form_Load' | Out-Null
    Assert-Equal 'Facturas|True|Activar facturacion de prueba|True|False' (Run 'FormState') 'Form_Load in the Retail IDE shows the activation button'
    Run 'Reset' $false $retailPath | Out-Null
    Run 'Form_Load' | Out-Null
    Assert-Equal 'Facturas|False|Activar facturacion de prueba|True|False' (Run 'FormState') 'Form_Load in the executable keeps the button hidden'

    Run 'Reset' $true $retailPath | Out-Null
    Assert-Equal '0|True|1|' (Run 'PressKey' 123 3) 'Ctrl+Shift+F12 in the Retail IDE shows the button and consumes the key'
    Run 'Reset' $false $retailPath | Out-Null
    Assert-Equal "0|False|0|Modo de prueba no disponible>$unavailableText;" (Run 'PressKey' 123 3) 'Ctrl+Shift+F12 outside the IDE explains why and keeps the button hidden'
    foreach ($key in @(@{ Code = 123; Shift = 2; Name = 'Ctrl+F12' }, @{ Code = 123; Shift = 0; Name = 'F12' }, @{ Code = 122; Shift = 3; Name = 'Ctrl+Shift+F11' })) {
        Run 'Reset' $true $retailPath | Out-Null
        Assert-Equal "$($key.Code)|False|0|" (Run 'PressKey' $key.Code $key.Shift) "$($key.Name) passes through untouched"
    }

    Run 'Reset' $true $retailPath | Out-Null
    $sc.ExecuteStatement('MsgAnswer = vbNo')
    Run 'cmdInvoiceSimulation_Click' | Out-Null
    Assert-Equal "FACTURACION DE PRUEBA>$confirmText;" (Eval 'MsgLog') 'Activation asks for confirmation with the local database warning'
    Assert-Equal '308' (Eval 'MsgButtons') 'Confirmation is Yes/No, exclamation, default No'
    Assert-Equal 'Facturas|False|Activar facturacion de prueba|True|False' (Run 'FormState') 'Answering No keeps simulation off'

    $sc.ExecuteStatement('MsgAnswer = vbYes : MsgLog = ""')
    Run 'cmdInvoiceSimulation_Click' | Out-Null
    Assert-Equal "FACTURAS - $marker|True|$marker|False|True" (Run 'FormState') 'Answering Yes enables simulation and marks the form'
    $sc.ExecuteStatement('MsgLog = ""')
    Run 'cmdInvoiceSimulation_Click' | Out-Null
    Assert-Equal '' (Eval 'MsgLog') 'Clicking again while enabled asks nothing'
    Run 'Form_Load' | Out-Null
    Assert-Equal "FACTURAS - $marker|True|$marker|False|True" (Run 'FormState') 'Reopening FormFactura keeps the simulation marker'

    Run 'Reset' $false $retailPath | Out-Null
    $sc.ExecuteStatement('MsgAnswer = vbYes')
    Run 'cmdInvoiceSimulation_Click' | Out-Null
    Assert-Equal '|Facturas|False|Activar facturacion de prueba|True|False' ((Eval 'MsgLog') + '|' + (Run 'FormState')) 'Click outside the IDE does nothing'

    # Print marker helper
    foreach ($case in @(
        @{ Value = '99999999999999'; Expected = 'True' },
        @{ Value = '"99999999999999"'; Expected = 'True' },
        @{ Value = '75123456789012'; Expected = 'False' },
        @{ Value = '0'; Expected = 'False' },
        @{ Value = 'Null'; Expected = 'False' },
        @{ Value = '""'; Expected = 'False' },
        @{ Value = 'Empty'; Expected = 'False' }
    )) {
        Assert-Equal $case.Expected (Eval "B(IsInvoiceSimulationCae($($case.Value)))") "IsInvoiceSimulationCae($($case.Value))"
    }
}

# --- Print marker call sites -----------------------------------------------
function Test-CaeMarker([string]$name, [string]$text, [string]$recordset, [int]$expected, [int]$commented) {
    $wrapped = [regex]::Matches($text, "\r\n( +)If IsInvoiceSimulationCae\($recordset!CAE\) Then\r\n\1    Printer\.Print `"$marker`"\r\n\1Else\r\n\1    Printer\.Print `"C\.A\.E: `" & $recordset!CAE\r\n\1End If\r\n").Count
    $active = [regex]::Matches($text, "(?m)^[ \t]*Printer\.Print `"C\.A\.E: `"").Count
    $inactive = [regex]::Matches($text, "(?m)^'[ \t]*Printer\.Print `"C\.A\.E: `"").Count
    Assert-Equal $expected $wrapped "$name wraps every active CAE print with the PRUEBA marker"
    Assert-Equal $expected $active "$name active CAE prints"
    Assert-Equal $commented $inactive "$name commented CAE prints stay commented"
}
Test-CaeMarker 'FormImpriir.frm' $impriir 'tFacturaC' 6 2
Test-CaeMarker 'FormImprimirNC.frm' $imprimirNC 'tNotaCreditoC' 8 0

if ($failures.Count -gt 0) {
    Write-Host "`n$($failures.Count) check(s) failed."
    exit 1
}
Write-Host "`nAll checks passed."
exit 0
