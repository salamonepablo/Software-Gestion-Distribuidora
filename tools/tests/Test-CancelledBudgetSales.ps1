# Run with 32-bit Windows PowerShell; only disposable DAO fixtures are written.
$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 4) { throw 'Use 32-bit Windows PowerShell.' }
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$source = [Text.Encoding]::GetEncoding(1252).GetString([IO.File]::ReadAllBytes((Join-Path $root 'FormListadoVentas.frm')))
$liquidar = [regex]::Match($source, '(?s)Private Sub cmdLiquidar_Click\(\).*?End Sub').Value
$expressions = [regex]::Matches($liquidar, '(?m)^\s*vsql2 = (.+)\r?$')
if ($expressions.Count -ne 4) { throw 'Expected four active budget query expressions.' }
$vb = New-Object -ComObject MSScriptControl.ScriptControl
$vb.Language = 'VBScript'
$helper = [regex]::Match($source, '(?s)Private Function SQLVentasConNC\(.*?End Function').Value
$vb.AddCode(($helper -replace 'Private Function', 'Function' -replace 'ByVal ', '' -replace ' As String', '' -replace ' As Long', '' -replace 'Mid\$', 'Mid'))
$dao = New-Object -ComObject DAO.DBEngine.36
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('SPC-CancelledSales-' + [Guid]::NewGuid().ToString('N') + '.mdb')
$db = $dao.CreateDatabase($fixture, ';LANGID=0x0409;CP=1252;COUNTRY=0', 32)
Write-Output "Disposable fixture: $fixture"
function Exec([string]$sql) { $db.Execute($sql, 128) }
function Assert($actual, $expected, [string]$label) {
    if ([Math]::Abs([double]$actual - [double]$expected) -gt 0.000001) { throw "$label expected=$expected actual=$actual" }
}
function Adapt([string]$text) {
    $text = $text -replace 'cmbVendedores\(0\)\.text', 'Vendedor' -replace 'cmbVendedores\(1\)\.text', 'Vendedor'
    $text = $text -replace 'cmbProductos\(1\)\.text', 'Producto' -replace 'Val\(cmbCliente\.text\)', 'Cliente' -replace 'cmbCliente\.text', 'ClienteTexto'
    return $text
}
function Filters([string]$product='*', [string]$client='Todos', [string]$seller='GP', [string]$from='9/1/2026', [string]$to='9/30/2026') {
    $vb.ExecuteStatement('FechaDesde="' + $from + '": FechaHasta="' + $to + '": Vendedor="' + $seller + '": Producto="' + $product + '": ClienteTexto="' + $client + '": Cliente=' + $(if ($client -eq 'Todos') { '0' } else { $client }))
}
function QueryTotals([int]$branch) {
    $sql = $vb.Eval((Adapt $expressions[$branch].Groups[1].Value.Trim()))
    $rs = $db.OpenRecordset($sql, 4)
    $qty=0.0; $money=0.0
    try { while (-not $rs.EOF) { $qty += $rs.Fields.Item('cantidad').Value; $money += $rs.Fields.Item('totalLinea').Value; $rs.MoveNext() } }
    finally { $rs.Close() }
    return @($qty,$money)
}
# Execute actual Liquidar aggregation, translating VB6 syntax/UI only, not arithmetic.
$body = $liquidar.Substring(0, $liquidar.IndexOf('CapturaErrores:')) + "End Sub"
$body = Adapt $body
$body = $body -replace 'Private Sub cmdLiquidar_Click', 'Sub Liquidar'
$body = $body -replace ' As (Double|Boolean)', '' -replace 'Call SeteoGrilla', '' -replace 'On Error GoTo CapturaErrores', ''
$body = $body -replace 'FechaDesde = Format\(txtFechaDesde.text, "m/d/yyyy"\)', '' -replace 'FechaHasta = Format\(txtFechaHasta.text, "m/d/yyyy"\)', ''
$body = $body -replace 'OptionL1.Value', 'L1' -replace 'OptionL2.Value', 'L2' -replace 'OptionAll.Value', 'Todos'
$body = $body -replace 'lblTotal.Visible = True', '' -replace 'lblTotal.Caption = .*', 'CaptionQty = totalProductos'
$body = $body -replace 'txtImporteTotal.text = .*', 'Amount = GrandTotal'
$body = [regex]::Replace($body, '(\w+)!(\w+)', '$1.Fields("$2").Value')
$vb.AddObject('BaseSPC', $db)
$vb.AddCode(@'
Const dbOpenSnapshot = 4, dbOpenDynaset = 2, dbOpenTable = 1
Dim FechaDesde, FechaHasta, Vendedor, Producto, Cliente, ClienteTexto
Dim L1, L2, Todos, Cant, IdCodProd, IdCodProd2, GridQty, GridAmount, GridRows, CaptionQty, Amount
Dim qventasL1, qVentasL2, tVA, vSQL1, vsql2
Sub LlenarGrilla(p, q, m)
    GridQty = GridQty + q
    GridAmount = GridAmount + m
    GridRows = GridRows + 1
End Sub
Function BuscarDescProd(p)
    BuscarDescProd = p
End Function
'@)
$vb.AddCode($body)
function Aggregate([bool]$all=$false) {
    $vb.ExecuteStatement('L1=False: L2=' + $(if ($all) {'False'} else {'True'}) + ': Todos=' + $(if ($all) {'True'} else {'False'}) + ': Cant=0: IdCodProd="": IdCodProd2="": GridQty=0: GridAmount=0: GridRows=0: CaptionQty=-999: Amount=-999')
    $vb.Run('Liquidar') | Out-Null
    Assert ($vb.Eval('CaptionQty')) ($vb.Eval('GridQty')) 'Caption equals displayed grid quantity'
    Assert ($vb.Eval('Amount')) ($vb.Eval('GridAmount')) 'Amount equals displayed grid amount'
}
try {
    Exec 'CREATE TABLE PresupuestoC (NroPresu LONG, FechaPresu DATETIME, CodCliente LONG, CodVendedor TEXT(20), Anulado TEXT(10))'
    Exec 'CREATE TABLE PresupuestoD (NroPresu LONG, CodProd TEXT(20), cantidad DOUBLE, totalLinea CURRENCY)'
    $db.CreateQueryDef('qListadoVentasPres', 'SELECT C.FechaPresu,C.CodCliente,C.CodVendedor,C.Anulado,D.CodProd,D.cantidad,D.totalLinea FROM PresupuestoC C INNER JOIN PresupuestoD D ON C.NroPresu=D.NroPresu') | Out-Null
    Exec 'CREATE TABLE FacFixture (FechaFactura DATETIME, CodCliente LONG, CodVendedor TEXT(20), IDCodProd TEXT(20), cantidad DOUBLE, totalLinea CURRENCY, PorcentajeIVA DOUBLE)'
    $db.CreateQueryDef('qListadoVentasFac', 'SELECT * FROM FacFixture') | Out-Null
    Exec 'CREATE TABLE NotaCreditoC (TipoNotaCredito TEXT(2), NroNotaCredito LONG, FechaNotaCredito DATETIME, CodCliente LONG, CodVendedor TEXT(20), PorcentajeIVA DOUBLE)'
    Exec 'CREATE TABLE NotaCreditoD (TipoNotaCredito TEXT(2), NroNotaCredito LONG, IDCodProd TEXT(20), Cantidad DOUBLE, TotalLinea CURRENCY)'
    Exec 'CREATE TABLE tAuxiliarVentas (IdProducto TEXT(20) CONSTRAINT PrimaryKey PRIMARY KEY, Descripcion TEXT(30), cantidad DOUBLE, Importe CURRENCY)'
    Exec "INSERT INTO PresupuestoC VALUES (1,#9/1/2026#,34,'GP',NULL)"
    Exec "INSERT INTO PresupuestoC VALUES (2,#9/30/2026#,34,'GP',NULL)"
    Exec "INSERT INTO PresupuestoC VALUES (212944,#9/28/2026#,34,'GP','si')"
    Exec "INSERT INTO PresupuestoD VALUES (1,'65',1,106400)"
    Exec "INSERT INTO PresupuestoD VALUES (1,'65',1,106400)"
    Exec "INSERT INTO PresupuestoD VALUES (2,'75',3,300)"
    Exec "INSERT INTO PresupuestoD VALUES (212944,'65',1,106400)"
    # Collect both intended RED failures rather than stopping after the first.
    $failures = @()
    Filters
    try { $r=QueryTotals 0; Assert $r[0] 5 'Cancelled header excluded, NULL retained' } catch { $failures += $_.Exception.Message }
    Filters '*' 'Todos' 'GP' '9/1/2026' '9/1/2026'
    try { Aggregate; Assert ($vb.Eval('GridQty')) 2 'Duplicate active details' } catch { $failures += $_.Exception.Message }
    if ($failures.Count) { throw ($failures -join '; ') }
    foreach ($row in @(@(3,'8/31/2026',34,'GP'),@(4,'10/1/2026',34,'GP'),@(5,'9/15/2026',99,'OTHER'),@(6,'9/15/2026',99,'GP'))) {
        Exec "INSERT INTO PresupuestoC VALUES ($($row[0]),#$($row[1])#,$($row[2]),'$($row[3])',NULL)"
        Exec "INSERT INTO PresupuestoD VALUES ($($row[0]),'65',10,1000)"
    }
    for ($branch=0; $branch -lt 4; $branch++) {
        $client = if ($branch % 2) {'34'} else {'Todos'}
        Filters '*' $client
        $r=QueryTotals $branch
        Assert $r[0] $(if ($client -eq 'Todos') {15} else {5}) "Four query branches $branch / client"
        Assert $r[1] $(if ($client -eq 'Todos') {214100} else {213100}) "Unchanged budget amount $branch"
        Filters '65' $client
        $r=QueryTotals $branch; Assert $r[0] $(if ($client -eq 'Todos') {12} else {2}) 'Product filter'
        Filters '*' $client 'NONE'; $r=QueryTotals $branch; Assert $r[0] 0 'Empty seller'
        Filters '*' $client 'GP' '9/28/2026' '9/28/2026'; $r=QueryTotals $branch; Assert $r[0] 0 'Cancelled-only query'
    }
    Filters '*' '99' 'OTHER'; $r=QueryTotals 1
    Assert $r[0] 10 'Matching alternate client and vendor'
    Filters '65' 'Todos' '*'; $r=QueryTotals 0
    Assert $r[0] 22 'Wildcard vendor includes both active clients, excludes cancellation'
    Filters; Aggregate
    Assert ($vb.Eval('GridQty')) 15 'All-client L2 aggregation'
    Assert ($vb.Eval('Amount')) 214100 'All-client L2 amount'
    Filters '*' '34'; Aggregate
    Assert ($vb.Eval('GridQty')) 5 'Grouped L2 quantity'
    Assert ($vb.Eval('GridRows')) 2 'Grouped L2 rows'
    Assert ($vb.Eval('Amount')) 213100 'L2 amount'
    Filters '*' '34' 'GP' '9/28/2026' '9/28/2026'; Aggregate
    Assert ($vb.Eval('GridRows')) 0 'Cancelled-only L2 has no phantom row'
    Filters '*' '34' 'NONE'; Aggregate
    Assert ($vb.Eval('GridRows')) 0 'Empty L2 has no phantom row'
    Exec "INSERT INTO FacFixture VALUES (#9/7/2026#,34,'GP','65',4,100,21)"
    Exec "INSERT INTO NotaCreditoC VALUES ('A',1,#9/7/2026#,34,'GP',21)"
    Exec "INSERT INTO NotaCreditoD VALUES ('A',1,'65',1,50)"
    Exec "INSERT INTO NotaCreditoD VALUES ('A',1,'NC_ONLY',2,-10)"
    Filters '*' '34'; Aggregate $true
    Assert ($vb.Eval('GridQty')) 6 'Todos L1 plus L2 minus NC units'
    Assert ($vb.Eval('Amount')) 213172.6 'Todos unchanged signed NC / budget amounts'
    Assert ($vb.Eval('GridRows')) 3 'Todos overlap / L2-only / NC-only product groups'
    Filters; Aggregate $true
    Assert ($vb.Eval('GridQty')) 16 'All-client Todos aggregation'
    Assert ($vb.Eval('Amount')) 214172.6 'All-client Todos amount'
    Filters '65' '34' 'GP' '9/7/2026' '9/28/2026'; Aggregate $true
    Assert ($vb.Eval('GridQty')) 3 'Todos nonempty L1 with cancelled-only L2'
    Assert ($vb.Eval('Amount')) 60.5 'Todos preserves L1 amount when L2 empty'
    Write-Output 'PASS: actual four SQL expressions and Liquidar aggregation; NULL/si, duplicates/groups, boundaries/client/vendor/product, cancelled-only/empty L2, caption/grid, Todos mixed and signed NC.'
} finally { $db.Close() }
