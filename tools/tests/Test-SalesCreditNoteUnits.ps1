# Run with Windows SysWOW64 PowerShell: DAO 3.6 and ScriptControl are 32-bit.
$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 4) { throw 'Use 32-bit Windows PowerShell.' }
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$sourcePath = Join-Path $root 'FormListadoVentas.frm'
$encoding = [Text.Encoding]::GetEncoding(1252)
$bytes = [IO.File]::ReadAllBytes($sourcePath)
$source = $encoding.GetString($bytes)
if ($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) { throw 'VB6 source must not have a UTF-8 BOM.' }
if ($source -match '(?<!\r)\n|\r(?!\n)') { throw 'VB6 source must use CRLF exclusively.' }
if (-not [Linq.Enumerable]::SequenceEqual([byte[]]$bytes, [byte[]]$encoding.GetBytes($source))) { throw 'Windows-1252 byte roundtrip failed.' }
# Roundtrip alone also accepts UTF-8 bytes decoded as ANSI; check an existing accented label.
if (-not $source.Contains('Liquidaci' + [char]0xF3 + 'n Total:')) { throw 'Existing ANSI accented label was corrupted or re-encoded.' }
Write-Output 'PASS: VB6 Windows-1252 roundtrip, no UTF-8 BOM, CRLF exclusively.'
$liquidar = [regex]::Match($source, '(?s)Private Sub cmdLiquidar_Click\(\).*?End Sub').Value
$expressions = [regex]::Matches($liquidar, '(?m)^\s*vSQL1 = (.+)\r?$')
if ($expressions.Count -ne 4) { throw 'Expected four active invoice filter expressions.' }
$helper = [regex]::Match($source, '(?s)Private Function SQLVentasConNC\(.*?End Function').Value
$vb = New-Object -ComObject MSScriptControl.ScriptControl
$vb.Language = 'VBScript'
if ($helper) {
    # Execute the production body; translate only VB6 declarations and Mid$ spelling.
    $body = $helper -replace 'Private Function', 'Function' -replace 'ByVal ', '' -replace ' As String', '' -replace ' As Long', '' -replace 'Mid\$', 'Mid'
    $vb.AddCode($body)
    $calls = [regex]::Matches($liquidar, 'OpenRecordset\(SQLVentasConNC\(vSQL1\), dbOpenSnapshot\)')
    if ($calls.Count -ne 2) { throw 'Both active L1 consumers must use the helper.' }
}
$dao = New-Object -ComObject DAO.DBEngine.36
# Deliberately retained for inspection; never open or copy the working MDB.
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('SPC-SalesUnits-' + [Guid]::NewGuid().ToString('N') + '.mdb')
$db = $dao.CreateDatabase($fixture, ';LANGID=0x0409;CP=1252;COUNTRY=0', 32)
Write-Output "Disposable fixture: $fixture"
function Exec([string]$sql) { $db.Execute($sql, 128) }
function Assert($actual, $expected, [string]$label) {
    if ($actual -ne $expected) { throw "$label expected=$expected actual=$actual" }
}
function AssertMoney([double]$actual, [double]$expected, [string]$label) {
    if ([Math]::Abs($actual - $expected) -gt 0.0000001) { throw "$label expected=$expected actual=$actual" }
}
function Header([string]$table, [string]$type, [int]$number, [string]$date, [int]$client=34, [string]$seller='GP', [double]$iva=21) {
    if ($table -eq 'FacturaC') { Exec "INSERT INTO $table VALUES ('$type',$number,#$date#,$client,'$seller')" }
    else { Exec "INSERT INTO $table VALUES ('$type',$number,#$date#,$client,'$seller',$iva)" }
}
function Detail([string]$table, [string]$type, [int]$number, [string]$product, [double]$units, [double]$money=100, [double]$iva=21) {
    if ($table -eq 'FacturaD') { Exec "INSERT INTO $table VALUES ('$type',$number,'$product',$units,$money,$iva)" }
    else { Exec "INSERT INTO $table VALUES ('$type',$number,'$product',$units,$money,999999,30)" }
}
function Query([int]$branch, [string]$seller='GP', [string]$product='*', [int]$client=34, [string]$from='9/1/2026', [string]$to='9/30/2026', [switch]$InvoiceOnly) {
    $vb.ExecuteStatement('FechaDesde = "' + $from + '": FechaHasta = "' + $to + '": Vendedor = "' + $seller + '": Producto = "' + $product + '": Cliente = ' + $client)
    $expr = $expressions[$branch].Groups[1].Value.Trim() -replace 'cmbVendedores\(0\)\.text', 'Vendedor' -replace 'cmbProductos\(1\)\.text', 'Producto' -replace 'Val\(cmbCliente\.text\)', 'Cliente'
    $sql = $vb.Eval($expr)
    if ($helper -and -not $InvoiceOnly) { $sql = $vb.Run('SQLVentasConNC', $sql) }
    return [string]$sql
}
function Rows([string]$sql) {
    $rs = $db.OpenRecordset($sql, 4)
    $groups = @{}
    try {
        while (-not $rs.EOF) {
            $p = [string]$rs.Fields.Item('IDCodProd').Value
            if (-not $groups.ContainsKey($p)) { $groups[$p] = @(0.0, 0.0) }
            $groups[$p][0] += [double]$rs.Fields.Item('cantidad').Value
            $money = [double]$rs.Fields.Item('totalLinea').Value
            $iva = [double]$rs.Fields.Item('PorcentajeIVA').Value
            $groups[$p][1] += $money + ($money * $iva / 100)
            $rs.MoveNext()
        }
    } finally { $rs.Close() }
    return ,$groups
}
try {
    Exec 'CREATE TABLE FacturaC (TipoFactura TEXT(2), NroFactura LONG, FechaFactura DATETIME, CodCliente LONG, CodVendedor TEXT(20))'
    Exec 'CREATE TABLE FacturaD (TipoFactura TEXT(2), NroFactura LONG, IDCodProd TEXT(20), Cantidad DOUBLE, totalLinea CURRENCY, PorcentajeIVA DOUBLE)'
    Exec 'CREATE TABLE NotaCreditoC (TipoNotaCredito TEXT(2), NroNotaCredito LONG, FechaNotaCredito DATETIME, CodCliente LONG, CodVendedor TEXT(20), PorcentajeIVA DOUBLE)'
    Exec 'CREATE TABLE NotaCreditoD (TipoNotaCredito TEXT(2), NroNotaCredito LONG, IDCodProd TEXT(20), Cantidad DOUBLE, TotalLinea CURRENCY, PrecioUnitario CURRENCY, PorcentajeDescuento DOUBLE)'
    $db.CreateQueryDef('qListadoVentasFac', 'SELECT C.FechaFactura,C.CodCliente,C.CodVendedor,D.IDCodProd,D.Cantidad,D.totalLinea,D.PorcentajeIVA FROM FacturaC C INNER JOIN FacturaD D ON C.TipoFactura=D.TipoFactura AND C.NroFactura=D.NroFactura') | Out-Null
    Header 'FacturaC' 'A' 1 '9/7/2026'
    Header 'NotaCreditoC' 'A' 741 '9/7/2026'
    $products = @('45FKA','65','75','ON75')
    $original = @(2,4,2,2)
    for ($i=0; $i -lt 4; $i++) {
        Detail 'FacturaD' 'A' 1 $products[$i] $original[$i]
        Detail 'NotaCreditoD' 'A' 741 $products[$i] 1
    }
    for ($branch=0; $branch -lt 4; $branch++) {
        $result = Rows (Query $branch)
        $sum = 0
        for ($i=0; $i -lt 4; $i++) {
            Assert $result[$products[$i]][0] ($original[$i]-1) "September branch $branch product $($products[$i])"
            AssertMoney $result[$products[$i]][1] 0 "Invoice minus recorded NC plus IVA branch $branch"
            $sum += $result[$products[$i]][0]
        }
        Assert $sum 6 "September total branch $branch"
    }
    Write-Output 'PASS: September 1/3/1/1 total 6; all four L1/Todos filter branches; positive recorded NC money reduces sales.'
    # Same NC number, different type: only the exact composite pair may match.
    Header 'NotaCreditoC' 'B' 741 '9/8/2026' 99 'OTHER'
    Detail 'NotaCreditoD' 'B' 741 'COLLISION' 9
    Detail 'NotaCreditoD' 'C' 741 'ORPHAN' 10
    # NC-only product, fractional units, duplicate details (UNION ALL), date boundaries.
    Header 'NotaCreditoC' 'A' 2 '9/1/2026'
    Header 'NotaCreditoC' 'A' 3 '9/30/2026'
    Header 'NotaCreditoC' 'A' 4 '8/31/2026'
    Header 'NotaCreditoC' 'A' 5 '10/1/2026'
    Detail 'NotaCreditoD' 'A' 2 'NC_ONLY' 0.5
    Detail 'NotaCreditoD' 'A' 2 'NC_ONLY' 0.5
    Detail 'NotaCreditoD' 'A' 3 'NC_ONLY' 2
    Detail 'NotaCreditoD' 'A' 4 'OUTSIDE' 50
    Detail 'NotaCreditoD' 'A' 5 'OUTSIDE' 50
    for ($branch=0; $branch -lt 4; $branch++) {
        $r = Rows (Query $branch)
        Assert $r['NC_ONLY'][0] -3 "NC-only/boundaries branch $branch"
        AssertMoney $r['NC_ONLY'][1] -363 'NC-only recorded money plus IVA (three details, not quantity repricing)'
        foreach ($excluded in @('COLLISION','ORPHAN','OUTSIDE')) { Assert $r.ContainsKey($excluded) $false "Exclude $excluded" }
        $r = Rows (Query $branch 'GP' 'NC_*')
        Assert $r.Count 1 'Product LIKE filter'
        Assert $r['NC_ONLY'][0] -3 'Fractional duplicate details'
    }
    $r = Rows (Query 1 'OTHER' '*' 99)
    Assert $r['COLLISION'][0] -9 'NC composite/type/client/seller filter'
    Assert $r.Count 1 'Specific client/seller'
    $r = Rows (Query 0 '*' '*')
    Assert $r['COLLISION'][0] -9 'Todos client and seller wildcard'
    $r = Rows (Query 1 '*' '*' 34)
    Assert $r.ContainsKey('COLLISION') $false 'Client filter with wildcard seller'
    $r = Rows (Query 1 'NONE' '*')
    Assert $r.Count 0 'Empty result'
    Header 'FacturaC' 'B' 1 '9/30/2026'
    Detail 'FacturaD' 'B' 1 'MONEY' 1 -17.4321 10.5
    Header 'FacturaC' 'A' 6 '9/1/2026' 99 'OTHER'
    Detail 'FacturaD' 'A' 6 'MONEY' 2 31.2345 0
    for ($branch=0; $branch -lt 4; $branch++) {
        foreach ($seller in @('GP','*','OTHER')) {
            foreach ($product in @('*','M*','NC_*')) {
                $before = Rows (Query $branch $seller $product -InvoiceOnly)
                $after = Rows (Query $branch $seller $product)
                foreach ($p in $before.Keys) {
                    $credit = if ($products -contains $p) { 121 } else { 0 }
                    AssertMoney $after[$p][1] ($before[$p][1] - $credit) "Invoice arithmetic with NC netting $branch/$seller/$product/$p"
                }
                foreach ($p in $after.Keys) {
                    if (-not $before.ContainsKey($p)) {
                        $credit = if ($p -eq 'NC_ONLY') { 363 } else { 121 }
                        AssertMoney $after[$p][1] (-$credit) "NC-only money $p"
                    }
                }
            }
        }
    }
    Write-Output 'PASS: invoice arithmetic preserved including negative/fractional money, varied VAT, and invoice composite-number collision; filtered NC money netted.'
    # Signed detail must be reversed, not normalized with Abs; IVA is on the NC header.
    Header 'NotaCreditoC' 'A' 7 '9/7/2026' 34 'GP' 10.5
    Detail 'NotaCreditoD' 'A' 7 'SIGNED' 1 -17.4321
    Header 'NotaCreditoC' 'A' 8 '9/7/2026' 34 'GP' 0
    Detail 'NotaCreditoD' 'A' 8 'RECORDED' 2 31.2345
    for ($branch=0; $branch -lt 4; $branch++) {
        $r = Rows (Query $branch)
        AssertMoney $r['SIGNED'][1] 19.2624705 'Negative recorded NC detail reversed with header IVA'
        AssertMoney $r['RECORDED'][1] -31.2345 'Recorded discounted total, no unit-price or discount recalculation'
        Assert $r['SIGNED'][0] -1 'Signed-money detail still subtracts units'
    }
    # Recorded A-741 details: gross differs from the rounded header by 0.0018.
    Header 'FacturaC' 'A' 9 '9/7/2026' 99 'CASE'
    Header 'NotaCreditoC' 'A' 9 '9/7/2026' 99 'CASE'
    $recorded = @(76942.15,76942.15,92561.98,53223.14)
    for ($i=0; $i -lt 4; $i++) {
        Detail 'FacturaD' 'A' 9 $products[$i] $original[$i] 234749.995 0
        Detail 'NotaCreditoD' 'A' 9 $products[$i] 1 $recorded[$i]
    }
    $r = Rows (Query 3 'CASE' '*' 99)
    $gross = 0.0; $units = 0.0
    foreach ($p in $products) { $gross += $r[$p][1]; $units += $r[$p][0] }
    AssertMoney $gross 576399.9818 'Recorded September net without premature rounding'
    AssertMoney ([Math]::Round($gross,2)) 576399.98 'September display total'
    Assert $units 6 'Recorded September net units'
    Write-Output 'PASS: signed NC detail, recorded discount/header IVA, September rounded net 576399.98 (unrounded 576399.9818).'
    # Exercise actual L2 SQL with NULL active cancellation status.
    Exec 'CREATE TABLE PresFixture (FechaPresu DATETIME, CodVendedor TEXT(20), CodProd TEXT(20), CodCliente LONG, cantidad DOUBLE, totalLinea CURRENCY, Anulado TEXT(10))'
    $db.CreateQueryDef('qListadoVentasPres', 'SELECT * FROM PresFixture') | Out-Null
    Exec "INSERT INTO PresFixture VALUES (#9/7/2026#,'GP','45FKA',34,3,50,NULL)"
    Exec "INSERT INTO PresFixture VALUES (#9/7/2026#,'GP','L2_ONLY',34,2,70,NULL)"
    $l2expr = [regex]::Matches($liquidar, '(?m)^\s*vsql2 = (.+)\r?$')[3].Groups[1].Value.Trim() -replace 'cmbVendedores\(0\)\.text', 'Vendedor' -replace 'cmbProductos\(1\)\.text', 'Producto' -replace 'Val\(cmbCliente\.text\)', 'Cliente'
    $vb.ExecuteStatement('FechaDesde="9/1/2026": FechaHasta="9/30/2026": Vendedor="GP": Producto="*": Cliente=34')
    $l2 = $db.OpenRecordset($vb.Eval($l2expr), 4)
    $combined = Rows (Query 3)
    $l2units = 0; $l2money = 0
    while (-not $l2.EOF) {
        $p = [string]$l2.Fields.Item('CodProd').Value
        if (-not $combined.ContainsKey($p)) { $combined[$p] = @(0.0,0.0) }
        $units = [double]$l2.Fields.Item('cantidad').Value
        $money = [double]$l2.Fields.Item('totalLinea').Value
        $combined[$p][0] += $units; $combined[$p][1] += $money
        $l2units += $units; $l2money += $money
        $l2.MoveNext()
    }
    $l2.Close()
    Assert $l2units 5 'L2 source units unchanged'
    Assert $l2money 120 'L2 money unchanged'
    Assert $combined['45FKA'][0] 4 'Combined overlapping product'
    AssertMoney $combined['45FKA'][1] 50 'Combined invoice minus NC plus unchanged budget money'
    Assert $combined['NC_ONLY'][0] -3 'Combined NC-only product'
    Assert $combined['L2_ONLY'][0] 2 'Combined L2-only product'
    Write-Output 'PASS: composite joins, NC-only, fractional/duplicate units, date endpoints/exclusions, client/seller/product filters, empty source, L2 source and combined aggregation.'
} finally {
    $db.Close()
}
