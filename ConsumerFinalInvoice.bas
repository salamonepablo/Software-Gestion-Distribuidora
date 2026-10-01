Attribute VB_Name = "ConsumerFinalInvoice"
Option Explicit

Public Const ConsumerFinalIdentificationLimit As Currency = 10000000@

Public Function IsConsumerFinalCustomer(ByVal taxCondition As Variant, ByVal customerName As Variant) As Boolean
    Dim conditionText As String
    Dim nameText As String
    If Not IsNull(taxCondition) Then conditionText = UCase$(Trim$(CStr(taxCondition)))
    If Not IsNull(customerName) Then nameText = UCase$(Trim$(CStr(customerName)))
    IsConsumerFinalCustomer = (conditionText = "CF" Or nameText = "CONSUMIDOR FINAL")
End Function

Public Function ResolveInvoiceRecipient(ByVal invoiceType As String, ByVal taxCondition As Variant, ByVal customerName As Variant, ByVal documentText As String, ByVal amount As Currency, ByRef documentType As Long, ByRef documentNumber As Double, ByRef errorText As String) As Boolean
    Dim normalized As String
    errorText = ""
    normalized = Trim$(documentText)

    If UCase$(Trim$(invoiceType)) = "B" And IsConsumerFinalCustomer(taxCondition, customerName) Then
        If amount < ConsumerFinalIdentificationLimit Then
            documentType = 99
            documentNumber = 0
            ResolveInvoiceRecipient = True
            Exit Function
        End If
        If IsValidCuit(normalized) Then
            documentType = 80
        ElseIf IsValidDni(normalized) Then
            documentType = 96
        Else
            errorText = "Para una factura B a consumidor final de $10.000.000 o mas, ingrese un CUIT o DNI valido antes de guardar."
            Exit Function
        End If
        documentNumber = CDbl(normalized)
        ResolveInvoiceRecipient = True
        Exit Function
    End If

    If normalized <> "" Then
        documentType = 80
        documentNumber = CDbl(normalized)
    Else
        documentType = 96
        documentNumber = 11111111
    End If
    ResolveInvoiceRecipient = True
End Function

Public Function IsValidDni(ByVal value As String) As Boolean
    If Len(value) <> 7 And Len(value) <> 8 Then Exit Function
    If Not value Like String$(Len(value), "#") Then Exit Function
    IsValidDni = (CDbl(value) > 0)
End Function

Public Function IsValidCuit(ByVal value As String) As Boolean
    Dim weights As Variant
    Dim index As Long
    Dim checksum As Long
    Dim checkDigit As Long
    If Len(value) <> 11 Then Exit Function
    If Not value Like String$(11, "#") Then Exit Function
    weights = Array(5, 4, 3, 2, 7, 6, 5, 4, 3, 2)
    For index = 1 To 10
        checksum = checksum + CLng(Mid$(value, index, 1)) * weights(index - 1)
    Next index
    checkDigit = 11 - (checksum Mod 11)
    If checkDigit = 11 Then checkDigit = 0
    If checkDigit = 10 Then checkDigit = 9
    IsValidCuit = (checkDigit = CLng(Right$(value, 1)))
End Function

Public Function EnsureInvoiceRecipientFields(ByVal database As DAO.Database, ByRef errorText As String) As Boolean
    On Error GoTo Failed
    Dim field As DAO.Field
    Dim hasType As Boolean
    Dim hasNumber As Boolean
    database.TableDefs.Refresh
    For Each field In database.TableDefs("FacturaC").Fields
        If field.Name = "RecipientDocType" Then hasType = True
        If field.Name = "RecipientDocNumber" Then hasNumber = True
    Next field
    If Not hasType Then database.Execute "ALTER TABLE FacturaC ADD COLUMN RecipientDocType LONG", dbFailOnError
    If Not hasNumber Then database.Execute "ALTER TABLE FacturaC ADD COLUMN RecipientDocNumber DOUBLE", dbFailOnError
    EnsureInvoiceRecipientFields = True
    Exit Function
Failed:
    errorText = Err.Description
End Function

Public Sub PrintedInvoiceRecipient(ByVal invoice As DAO.Recordset, ByVal customerId As Long, ByRef documentType As Long, ByRef documentNumber As Double)
    On Error GoTo LegacyInvoice
    If Not IsNull(invoice.Fields("RecipientDocType").Value) And Not IsNull(invoice.Fields("RecipientDocNumber").Value) Then
        documentType = CLng(invoice.Fields("RecipientDocType").Value)
        documentNumber = CDbl(invoice.Fields("RecipientDocNumber").Value)
        Exit Sub
    End If
LegacyInvoice:
    documentType = 80
    documentNumber = CUITCliente(customerId)
End Sub
