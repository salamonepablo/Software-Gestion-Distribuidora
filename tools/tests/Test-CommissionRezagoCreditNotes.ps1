# Run in 32-bit Windows PowerShell. Only the retained disposable MDB is written.
$ErrorActionPreference = 'Stop'
if ([IntPtr]::Size -ne 4) { throw 'Use 32-bit Windows PowerShell.' }
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$encoding = [Text.Encoding]::GetEncoding(1252)
$bytes = [IO.File]::ReadAllBytes((Join-Path $root 'FormLiqComisiones.frm'))
$source = $encoding.GetString($bytes)
if ($bytes[0] -eq 239 -and $bytes[1] -eq 187 -and $bytes[2] -eq 191) { throw 'UTF-8 BOM in VB6 source.' }
if ($source -match '(?<!\r)\n|\r(?!\n)') { throw 'VB6 source requires CRLF.' }
if (-not [Linq.Enumerable]::SequenceEqual([byte[]]$bytes,[byte[]]$encoding.GetBytes($source))) { throw 'ANSI roundtrip failed.' }
if (-not $source.Contains('Liquidaci' + [char]0xF3 + 'n Total:')) { throw 'ANSI accented label corrupted.' }
Write-Output 'PASS: ANSI bytes, accented label, CRLF, no BOM.'
$liquidar = [regex]::Match($source,'(?s)Private Sub cmdLiquidar_Click\(\).*?End Sub').Value
$helper = [regex]::Match($source,'(?s)Private Function SQLComisionesConRezago\(.*?End Function').Value
$vb = New-Object -ComObject MSScriptControl.ScriptControl
$vb.Language = 'VBScript'
if ($helper) {
    $vb.AddCode(($helper -replace 'Private Function','Function' -replace 'ByVal ','' -replace ' As String',''))
    if (-not $liquidar.Contains('SQLComisionesConRezago(FechaDesde, FechaHasta, cmbVendedores(0).text)')) { throw 'UI must consume production helper.' }
    if ($liquidar.Contains('qComisiones.MoveFirst')) { throw 'Empty snapshot must not MoveFirst.' }
    if (-not $liquidar.Contains('If qComisiones.EOF Then')) { throw 'Missing local empty-result guard.' }
    if ($liquidar.Contains('[PagoC.NroPago]') -or $liquidar.Contains('[qLiqComisiones.Comision]')) { throw 'UI must consume stable aliases.' }
}
$dao = New-Object -ComObject DAO.DBEngine.36
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('SPC-CommissionRezago-' + [Guid]::NewGuid().ToString('N') + '.mdb')
$db = $dao.CreateDatabase($fixture,';LANGID=0x0409;CP=1252;COUNTRY=0',32)
Write-Output "Disposable fixture retained: $fixture"
function Exec([string]$sql) { $db.Execute($sql,128) }
function Assert($actual,$expected,[string]$label) {
    if ($actual -ne $expected) { throw "$label expected=$expected actual=$actual" }
}
function Money([double]$actual,[double]$expected,[string]$label) {
    if ([Math]::Abs($actual-$expected) -gt 0.0000001) { throw "$label expected=$expected actual=$actual" }
}
function Header([string]$type,[int]$number,[double]$gross=100,[int]$client=2608,[string]$date='9/9/2026',[string]$iva='21',[string]$iibb='0') {
    Exec "INSERT INTO NotaCreditoC VALUES ('$type',$number,#$date#,$client,'OLD',$gross,$iva,$iibb,25,999)"
}
function Detail([string]$type,[int]$number,[string]$code='REZAGO',[double]$amount=100) {
    $literal = if ($null -eq $code -or $code -eq '<NULL>') { 'NULL' } else { "'$code'" }
    Exec "INSERT INTO NotaCreditoD VALUES ('$type',$number,$literal,$amount,50,999,999)"
}
function Query([string]$seller='700',[string]$from='9/1/2026',[string]$to='9/30/2026') {
    if ($helper) { return [string]$vb.Run('SQLComisionesConRezago',$from,$to,$seller) }
    # Pre-change behavior, evaluated from the actual UI expression, for meaningful RED.
    $expr = [regex]::Match($liquidar,'(?m)^\s*vSQL = (.+)\r?$').Groups[1].Value.Trim() -replace 'cmbVendedores\(0\)\.text','Vendedor'
    $vb.ExecuteStatement('FechaDesde="' + $from + '": FechaHasta="' + $to + '": Vendedor="' + $seller + '"')
    return [string]$vb.Eval($expr)
}
function Rows([string]$sql) {
    $rs = $db.OpenRecordset($sql,4)
    $result = @()
    try {
        while (-not $rs.EOF) {
            if ($helper) {
                $doc = [string]$rs.Fields.Item('Documento').Value
                $rate = [double]$rs.Fields.Item('TasaComision').Value
            } else {
                $doc = [string]$rs.Fields.Item('PagoC.NroPago').Value
                $rate = [double]$rs.Fields.Item('qLiqComisiones.Comision').Value
            }
            $result += [pscustomobject]@{ Doc=$doc; Amount=[double]$rs.Fields.Item('ImportePago').Value; Rate=$rate; Method=[string]$rs.Fields.Item('FormaPago').Value; Client=[string]$rs.Fields.Item('RazonSocial').Value; Date=$rs.Fields.Item('FechaPago').Value }
            $rs.MoveNext()
        }
    } finally { $rs.Close() }
    return $result
}
try {
    Exec 'CREATE TABLE Empleados (Legajo TEXT(20), Comision DOUBLE)'
    Exec 'CREATE TABLE Clientes (IDCliente LONG, RazonSocial TEXT(50), Vendedor TEXT(20), PorcentajeComision DOUBLE)'
    Exec 'CREATE TABLE PagoC (NroPago LONG, FechaPago DATETIME, TotalAbonado CURRENCY, IDCliente LONG, Sucursal TEXT(2))'
    Exec 'CREATE TABLE PagoD (NroPago LONG, ImportePago CURRENCY, FormaPago TEXT(20), Sucursal TEXT(2))'
    Exec 'CREATE TABLE NotaCreditoC (TipoNotaCredito TEXT(2), NroNotaCredito LONG, FechaNotaCredito DATETIME, CodCliente LONG, CodVendedor TEXT(20), TotalNotaCredito CURRENCY, PorcentajeIVA DOUBLE, AlicuotaIIBB DOUBLE, PorcentajeDesc DOUBLE, totalIva CURRENCY)'
    Exec 'CREATE TABLE NotaCreditoD (TipoNotaCredito TEXT(2), NroNotaCredito LONG, IDCodProd TEXT(20), TotalLinea CURRENCY, PorcentajeDescuento DOUBLE, PrecioUnitario CURRENCY, Cantidad DOUBLE)'
    $original = 'SELECT PagoC.NroPago,PagoC.FechaPago,PagoC.TotalAbonado,Clientes.RazonSocial,Empleados.Legajo,IIF(Clientes.PorcentajeComision Is Null,Empleados.Comision,Clientes.PorcentajeComision) AS Comision,PagoD.ImportePago,PagoD.FormaPago,* FROM (Empleados INNER JOIN (Clientes INNER JOIN PagoC ON Clientes.IDCliente=PagoC.IDCliente) ON Empleados.Legajo=Clientes.Vendedor) INNER JOIN PagoD ON PagoC.NroPago=PagoD.NroPago ORDER BY PagoC.FechaPago,PagoD.FormaPago'
    $db.CreateQueryDef('qLiqComisiones',$original) | Out-Null
    Exec "INSERT INTO Empleados VALUES ('700',5)"
    Exec "INSERT INTO Empleados VALUES ('800',3)"
    Exec "INSERT INTO Clientes VALUES (2608,'A742 customer','700',2)"
    Exec "INSERT INTO Clientes VALUES (2,'Fallback','700',NULL)"
    Exec "INSERT INTO Clientes VALUES (3,'Zero','700',0)"
    Exec "INSERT INTO Clientes VALUES (4,'Other seller','800',4)"
    Header 'A' 742 735999.97
    Detail 'A' 742
    $r = @(Rows (Query))
    Assert $r.Count 1 'Missing all-REZAGO NC A742 collection row'
    Assert $r[0].Doc 'NC A-742' 'Composite NC identifier'
    Assert $r[0].Method 'Rezago' 'Collection method'
    Money $r[0].Amount 735999.97 'Recorded header gross, not net/repriced detail'
    Money $r[0].Rate 2 'Current customer override'
    $arithmetic = [regex]::Match($liquidar,'(?m)^\s*LiqLinea = (\(qComisiones!ImportePago.+)\r?$').Groups[1].Value.Trim()
    $arithmetic = $arithmetic -replace 'qComisiones!ImportePago','Amount' -replace 'qComisiones!TasaComision','Rate'
    $vb.ExecuteStatement('Amount=735999.97: Rate=2')
    Money ($vb.Eval($arithmetic)) 14719.9994 'Actual UI arithmetic, no premature rounding'
    Money ([Math]::Round($vb.Eval($arithmetic),2)) 14720 'A742 displayed commission'
    if (-not $liquidar.Contains('FG1.text = FormatCurrency(LiqLinea, 2)')) { throw 'UI must display two-decimal commission.' }
    if (-not $liquidar.Contains('FG1.Clear') -or -not $liquidar.Contains('txtImporteTotal.text = FormatCurrency(0, 2)')) { throw 'Empty run must clear prior grid/total.' }
    Detail 'A' 742
    Assert (@(Rows (Query)).Count) 1 'Multiple REZAGO details do not multiply header'
    # Ordinary, null, blank, unknown, non-exact codes, header with no detail.
    $codes = @('NORMAL','<NULL>','','UNKNOWN','REZAGO_EXTRA','rezago','REZAGO ')
    for ($i=0; $i -lt $codes.Count; $i++) {
        Header 'A' (10+$i)
        Detail 'A' (10+$i) $codes[$i]
        Header 'A' (30+$i)
        Detail 'A' (30+$i) $codes[$i]
    }
    Header 'A' 50
    Detail 'B' 50 # Orphan detail must not qualify the A header.
    Assert (@(Rows (Query)).Count) 1 'Exclude ordinary/inexact/no-detail/orphan documents'
    Header 'B' 742 12
    Detail 'B' 742
    Detail 'C' 742 'NORMAL' # Unrelated composite orphan must not contaminate A/B.
    $r = @(Rows (Query))
    Assert $r.Count 2 'Composite collision: A and B independently eligible'
    Money (($r | Where-Object Doc -eq 'NC B-742').Amount) 12 'B header not A header'
    Header 'A' 60 -17.4321 2 '9/1/2026'; Detail 'A' 60
    Header 'A' 61 100 3 '9/30/2026'; Detail 'A' 61
    Header 'A' 62 100 4; Detail 'A' 62
    Header 'A' 63 100 2608 '8/31/2026'; Detail 'A' 63
    Header 'A' 64 100 2608 '10/1/2026'; Detail 'A' 64
    $r = @(Rows (Query))
    Assert $r.Count 4 'Inclusive endpoints, other seller and outside dates excluded'
    $signed = $r | Where-Object Doc -eq 'NC A-60'
    Money $signed.Amount -17.4321 'Preserve recorded negative sign, never Abs'
    Money $signed.Rate 5 'Null customer override uses employee rate'
    Money (($r | Where-Object Doc -eq 'NC A-61').Rate) 0 'Zero customer override preserved'
    Assert (@(Rows (Query '800')).Count) 1 'Selected current seller, not original NC vendor'
    Exec "UPDATE Clientes SET Vendedor='800' WHERE IDCliente=2608"
    Assert (@(Rows (Query)).Count) 2 'Reassignment removes NC from old current seller'
    Assert (@(Rows (Query '800')).Count) 3 'Reassignment moves NC to new seller'
    Exec "UPDATE Clientes SET Vendedor='700' WHERE IDCliente=2608"
    Assert (@(Rows (Query 'NONE')).Count) 0 'Empty selected seller'
    Assert (@(Rows (Query '700' '1/1/2000' '1/2/2000')).Count) 0 'Empty selected dates'
    Write-Output 'PASS: NC-only, A742 gross/2%/14720, duplicate details, ordinary exclusions, exact code, composite/orphan, signed gross, current seller/reassignment, null/zero rates, date bounds, empty.'
    # Preserve the saved query including its known number-only cross-branch join.
    Exec "INSERT INTO PagoC VALUES (1,#9/1/2026#,100,2608,'L1')"
    Exec "INSERT INTO PagoC VALUES (1,#9/30/2026#,200,2,'L2')"
    Exec "INSERT INTO PagoD VALUES (1,25,'Efectivo','L1')"
    Exec "INSERT INTO PagoD VALUES (1,-10.1234,'Cheque','L2')"
    Exec "INSERT INTO PagoC VALUES (2,#8/31/2026#,300,2608,'L1')"
    Exec "INSERT INTO PagoD VALUES (2,300,'Efectivo','L1')"
    Exec "INSERT INTO PagoC VALUES (3,#9/9/2026#,50,4,'L1')"
    Exec "INSERT INTO PagoD VALUES (3,50,'Efectivo','L1')"
    $baseline = $db.OpenRecordset("SELECT * FROM qLiqComisiones WHERE FechaPago>=#9/1/2026# AND FechaPago<=#9/30/2026# AND Legajo='700' ORDER BY FormaPago,FechaPago",4)
    $expected = @()
    while (-not $baseline.EOF) {
        $expected += ('{0}|{1}|{2}|{3}|{4}|{5}' -f $baseline.Fields.Item('PagoC.NroPago').Value,$baseline.Fields.Item('FechaPago').Value,$baseline.Fields.Item('ImportePago').Value,$baseline.Fields.Item('FormaPago').Value,$baseline.Fields.Item('RazonSocial').Value,$baseline.Fields.Item('qLiqComisiones.Comision').Value)
        $baseline.MoveNext()
    }
    $baseline.Close()
    $payments = @(Rows (Query) | Where-Object Method -ne 'Rezago')
    Assert $payments.Count 4 'Existing cross-branch join remains unchanged'
    for ($i=0; $i -lt $payments.Count; $i++) {
        $p = $payments[$i]
        Assert ('{0}|{1}|{2}|{3}|{4}|{5}' -f $p.Doc,$p.Date,$p.Amount,$p.Method,$p.Client,$p.Rate) $expected[$i] 'Original payment row/order/rate/amount unchanged'
    }
    Assert (@(Rows (Query)).Count) 8 'Payments plus NC rows, no deduplication'
    Write-Output 'PASS: original payment rows, ordering, signed amount/rates, and latent cross-branch join preserved; union adds NC without multiplying headers.'
    Header 'A' 90 9999 2608 '9/9/2026' '21' '3'
    Detail 'A' 90 'REZAGO' 80 # Already discounted; stored price/quantity/discount deliberately disagree.
    Detail 'A' 90 'BATTERY' 500
    $r = @(Rows (Query))
    $mixed = @($r | Where-Object Doc -eq 'NC A-90')
    Assert $mixed.Count 1 'Mixed REZAGO plus battery must qualify once'
    Money $mixed[0].Amount 99.2 'Type A recorded discounted net plus IVA/IIBB; no manual header allocation'
    Detail 'A' 90 'REZAGO' -20
    Detail 'A' 90 '<NULL>' 900
    Detail 'A' 90 '' 900
    Detail 'A' 90 'rezago' 900
    Detail 'A' 90 'REZAGO ' 900
    $mixed = @(Rows (Query) | Where-Object Doc -eq 'NC A-90')
    Assert $mixed.Count 1 'Multiple signed REZAGO lines aggregate once'
    Money $mixed[0].Amount 74.4 'Signed net sum; null/blank/inexact lines excluded'
    Header 'B' 90 9999 2608 '9/9/2026' '21' '3'
    Detail 'B' 90 'REZAGO' 121
    Detail 'B' 90 'BATTERY' 605
    Header 'C' 90 9999
    Detail 'C' 90 'REZAGO' -121
    Detail 'C' 90 'NORMAL' 605
    Header 'A' 91 9999 2 '9/1/2026' 'NULL' 'NULL'
    Detail 'A' 91 'REZAGO' -10
    Detail 'A' 91 'NORMAL' 500
    Header 'A' 92 9999 3 '9/30/2026' '10.5' '2.5'
    Detail 'A' 92 'REZAGO' 100
    Detail 'A' 92 'NORMAL' 500
    Header 'A' 93 9999 4; Detail 'A' 93; Detail 'A' 93 'NORMAL'
    Header 'A' 94 9999 2608 '8/31/2026'; Detail 'A' 94; Detail 'A' 94 'NORMAL'
    Header 'A' 95 9999 2608 '10/1/2026'; Detail 'A' 95; Detail 'A' 95 'NORMAL'
    $r = @(Rows (Query))
    Assert $r.Count 13 'Five mixed rows added; selected seller/date bounds retained'
    Money (($r | Where-Object Doc -eq 'NC B-90').Amount) 121 'Non-A gross detail: no duplicate IVA/IIBB despite populated header rates/IVA'
    Money (($r | Where-Object Doc -eq 'NC C-90').Amount) -121 'Non-A signed gross, composite identity'
    Money (($r | Where-Object Doc -eq 'NC A-91').Amount) -10 'Null tax rates mean no additional tax'
    Money (($r | Where-Object Doc -eq 'NC A-91').Rate) 5 'Mixed current employee fallback'
    Money (($r | Where-Object Doc -eq 'NC A-92').Amount) 113 'Alternate applicable Type A IVA/IIBB rates'
    Money (($r | Where-Object Doc -eq 'NC A-92').Rate) 0 'Mixed zero customer rate'
    Exec "UPDATE Clientes SET Vendedor='800' WHERE IDCliente=2608"
    Assert (@(Rows (Query '800') | Where-Object Doc -eq 'NC A-90').Count) 1 'Mixed follows reassigned current seller'
    Assert (@(Rows (Query) | Where-Object Doc -eq 'NC A-90').Count) 0 'Mixed removed from former seller'
    Write-Output 'PASS: mixed discounted signed exact REZAGO aggregates; Type A IVA/IIBB/null rates; non-A gross no double tax; composite types; current seller/rates/date bounds; no repricing or discounts reapplied.'
} finally { $db.Close() }
