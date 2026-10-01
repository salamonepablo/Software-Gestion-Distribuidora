Attribute VB_Name = "DeliveryReassignment"
Option Explicit

Public Function ReassignDelivery(ByVal databasePath As String, ByVal branchNumber As Long, ByVal deliveryNumber As Long, ByVal invoiceType As String, ByVal invoiceNumber As Long, ByVal customerNumber As Long, ByVal creditType As String, ByVal creditNumber As Long, ByRef failureReason As String) As Boolean
    Dim workspace As DAO.Workspace
    Dim database As DAO.Database
    Dim delivery As DAO.Recordset
    Dim oldInvoice As DAO.Recordset
    Dim newInvoice As DAO.Recordset
    Dim credit As DAO.Recordset
    Dim audit As DAO.Recordset
    Dim definition As DAO.TableDef
    Dim hasAudit As Boolean
    Dim transactionStarted As Boolean
    Dim sourceType As String
    Dim sourceNumber As Long
    Dim errorText As String
    failureReason = ""
    On Error GoTo Failed
    If (invoiceType <> "A" And invoiceType <> "B") Or (creditType <> "A" And creditType <> "B") Then Err.Raise vbObjectError + 2200, , "Tipo de comprobante no valido."
    Set workspace = DBEngine.CreateWorkspace("DeliveryReuse", "Admin", "", dbUseJet)
    Set database = workspace.OpenDatabase(databasePath)
    For Each definition In database.TableDefs
        If definition.Name = "DeliveryReassignmentHistory" Then hasAudit = True
    Next
    workspace.BeginTrans
    transactionStarted = True
    If Not hasAudit Then
        database.Execute "CREATE TABLE DeliveryReassignmentHistory (Id COUNTER CONSTRAINT PK_DeliveryReassignment PRIMARY KEY, BranchNumber LONG, DeliveryNumber LONG, SourceInvoiceType TEXT(1), SourceInvoiceNumber LONG, TargetInvoiceType TEXT(1), TargetInvoiceNumber LONG, CreditType TEXT(1), CreditNumber LONG, CustomerNumber LONG, CreatedAt DATETIME)", dbFailOnError
        database.Execute "CREATE UNIQUE INDEX UX_DeliveryReassignmentCredit ON DeliveryReassignmentHistory (CreditType, CreditNumber)", dbFailOnError
    End If
    Set delivery = database.OpenRecordset("SELECT * FROM RemitoC WHERE IdSucursal=" & branchNumber & " AND NroRemito=" & deliveryNumber, dbOpenDynaset)
    If delivery.EOF Then Err.Raise vbObjectError + 2201, , "No se encontro el remito."
    If delivery!CodCliente <> customerNumber Then Err.Raise vbObjectError + 2202, , "El remito pertenece a otro cliente."
    sourceType = Trim$(delivery!TipoFactura & "")
    sourceNumber = Val(delivery!NroFactura & "")
    If sourceType <> "A" And sourceType <> "B" Then Err.Raise vbObjectError + 2203, , "No se encontro una factura original valida."
    If sourceType = invoiceType And sourceNumber = invoiceNumber Then Err.Raise vbObjectError + 2204, , "El remito ya esta asociado a esta factura."
    Set oldInvoice = database.OpenRecordset("SELECT * FROM FacturaC WHERE TipoFactura='" & sourceType & "' AND NroFactura=" & sourceNumber, dbOpenSnapshot)
    Set newInvoice = database.OpenRecordset("SELECT * FROM FacturaC WHERE TipoFactura='" & invoiceType & "' AND NroFactura=" & invoiceNumber, dbOpenDynaset)
    Set credit = database.OpenRecordset("SELECT * FROM NotaCreditoC WHERE TipoNotaCredito='" & creditType & "' AND NroNotaCredito=" & creditNumber, dbOpenSnapshot)
    If oldInvoice.EOF Or newInvoice.EOF Or credit.EOF Then Err.Raise vbObjectError + 2205, , "Falta la factura original, la nueva o la nota de credito."
    If oldInvoice!CodCliente <> customerNumber Or newInvoice!CodCliente <> customerNumber Or credit!CodCliente <> customerNumber Then Err.Raise vbObjectError + 2206, , "Los comprobantes deben pertenecer al mismo cliente."
    If creditType <> sourceType Then Err.Raise vbObjectError + 2207, , "La nota de credito debe tener la misma letra que la factura original."
    If Len(Trim$(credit!CAE & "")) = 0 Or Val(credit!CAE & "") = 0 Then Err.Raise vbObjectError + 2208, , "La nota de credito no tiene CAE."
    If CCur(credit!TotalNotaCredito) <> CCur(oldInvoice!TotalFactura) Then Err.Raise vbObjectError + 2209, , "La nota de credito no es por el total de la factura original."
    If CDate(credit!FechaNotaCredito) < CDate(oldInvoice!FechaFactura) Then Err.Raise vbObjectError + 2210, , "La nota de credito es anterior a la factura original."
    If Val(newInvoice!NroRemito & "") <> 0 Then
        If Val(newInvoice!NroRemito & "") <> deliveryNumber Or Val(newInvoice!IdSucursal & "") <> branchNumber Then Err.Raise vbObjectError + 2211, , "La nueva factura ya tiene otro remito asociado."
    End If
    Set audit = database.OpenRecordset("SELECT * FROM DeliveryReassignmentHistory WHERE CreditType='" & creditType & "' AND CreditNumber=" & creditNumber, dbOpenDynaset)
    If Not audit.EOF Then Err.Raise vbObjectError + 2212, , "Esta nota de credito ya fue usada para reasociar un remito."
    If Not DeliveryLinesMatch(database, branchNumber, deliveryNumber, invoiceType, invoiceNumber) Then Err.Raise vbObjectError + 2213, , "Los productos y cantidades de la nueva factura no coinciden con el remito original."
    audit.AddNew
    audit!BranchNumber = branchNumber
    audit!DeliveryNumber = deliveryNumber
    audit!SourceInvoiceType = sourceType
    audit!SourceInvoiceNumber = sourceNumber
    audit!TargetInvoiceType = invoiceType
    audit!TargetInvoiceNumber = invoiceNumber
    audit!CreditType = creditType
    audit!CreditNumber = creditNumber
    audit!CustomerNumber = customerNumber
    audit!CreatedAt = Now
    audit.Update
    delivery.Edit
    delivery!TipoFactura = invoiceType
    delivery!NroFactura = invoiceNumber
    delivery.Update
    newInvoice.Edit
    newInvoice!IdSucursal = branchNumber
    newInvoice!NroRemito = deliveryNumber
    newInvoice.Update
    workspace.CommitTrans
    transactionStarted = False
    ReassignDelivery = True
CleanUp:
    On Error Resume Next
    If Not audit Is Nothing Then audit.Close
    If Not credit Is Nothing Then credit.Close
    If Not newInvoice Is Nothing Then newInvoice.Close
    If Not oldInvoice Is Nothing Then oldInvoice.Close
    If Not delivery Is Nothing Then delivery.Close
    If Not database Is Nothing Then database.Close
    If Not workspace Is Nothing Then workspace.Close
    Exit Function
Failed:
    errorText = Err.Description
    On Error Resume Next
    If transactionStarted Then workspace.Rollback
    failureReason = errorText
    GoTo CleanUp
End Function

Private Function DeliveryLinesMatch(ByVal database As DAO.Database, ByVal branchNumber As Long, ByVal deliveryNumber As Long, ByVal invoiceType As String, ByVal invoiceNumber As Long) As Boolean
    Dim deliveryLines As DAO.Recordset
    Dim invoiceLines As DAO.Recordset
    Set deliveryLines = database.OpenRecordset("SELECT IdCodProd, Sum(cantidad) AS Quantity FROM RemitoD WHERE IdSucursal=" & branchNumber & " AND NroRemito=" & deliveryNumber & " GROUP BY IdCodProd ORDER BY IdCodProd", dbOpenSnapshot)
    Set invoiceLines = database.OpenRecordset("SELECT IdCodProd, Sum(cantidad) AS Quantity FROM FacturaD WHERE TipoFactura='" & invoiceType & "' AND NroFactura=" & invoiceNumber & " GROUP BY IdCodProd ORDER BY IdCodProd", dbOpenSnapshot)
    If deliveryLines.EOF Or invoiceLines.EOF Then GoTo Done
    Do While Not deliveryLines.EOF And Not invoiceLines.EOF
        If CStr(deliveryLines!IdCodProd) <> CStr(invoiceLines!IdCodProd) Then GoTo Done
        If Abs(CDbl(deliveryLines!Quantity) - CDbl(invoiceLines!Quantity)) > 0.000001 Then GoTo Done
        deliveryLines.MoveNext
        invoiceLines.MoveNext
    Loop
    DeliveryLinesMatch = deliveryLines.EOF And invoiceLines.EOF
Done:
    deliveryLines.Close
    invoiceLines.Close
End Function
