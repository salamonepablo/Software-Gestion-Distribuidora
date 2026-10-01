Attribute VB_Name = "CreditAssociation"
Option Explicit

Public Function ResolveCreditAssociation(ByVal databasePath As String, ByVal selectedInvoiceNumber As String, ByVal invoiceType As String, ByVal customerNumber As Long, ByRef associatedType As Long, ByRef associatedNumber As Double, ByRef associatedDate As String, ByRef failureReason As String) As Boolean
    Dim database As DAO.Database
    Dim invoice As DAO.Recordset
    Dim invoiceNumber As Long
    On Error GoTo Failed
    associatedType = 0
    associatedNumber = 0
    associatedDate = ""
    failureReason = ""
    selectedInvoiceNumber = Trim$(selectedInvoiceNumber)
    If Len(selectedInvoiceNumber) = 0 Then
        ResolveCreditAssociation = True
        Exit Function
    End If
    If invoiceType <> "A" And invoiceType <> "B" Then Err.Raise vbObjectError + 2310, , "Tipo de factura asociada no valido."
    If Not IsNumeric(selectedInvoiceNumber) Then Err.Raise vbObjectError + 2311, , "Numero de factura asociada no valido."
    If CDbl(selectedInvoiceNumber) < 1 Or CDbl(selectedInvoiceNumber) > 2147483647# Then Err.Raise vbObjectError + 2311, , "Numero de factura asociada no valido."
    If CDbl(selectedInvoiceNumber) <> Fix(CDbl(selectedInvoiceNumber)) Then Err.Raise vbObjectError + 2311, , "Numero de factura asociada no valido."
    invoiceNumber = CLng(selectedInvoiceNumber)
    Set database = DBEngine.OpenDatabase(databasePath, False, True)
    Set invoice = database.OpenRecordset("SELECT NroFactura, TipoFactura, FechaFactura, CodCliente FROM FacturaC WHERE TipoFactura='" & invoiceType & "' AND NroFactura=" & CStr(invoiceNumber), dbOpenSnapshot)
    If invoice.EOF Then Err.Raise vbObjectError + 2312, , "La factura asociada no existe."
    If CLng(invoice!CodCliente) <> customerNumber Then Err.Raise vbObjectError + 2313, , "La factura asociada pertenece a otro cliente."
    If invoiceType = "A" Then associatedType = 1 Else associatedType = 6
    associatedNumber = CDbl(invoice!NroFactura)
    associatedDate = Format$(invoice!FechaFactura, "yyyymmdd")
    ResolveCreditAssociation = True
CleanUp:
    On Error Resume Next
    If Not invoice Is Nothing Then invoice.Close
    If Not database Is Nothing Then database.Close
    Exit Function
Failed:
    failureReason = Err.Description
    associatedType = 0
    associatedNumber = 0
    associatedDate = ""
    Resume CleanUp
End Function
