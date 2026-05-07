Option Explicit

' ============================================================
' FormOrdenPago - plantilla manual con persistencia minima
' ============================================================

Private Const ID_TABLA_ORDENPAGO As String = "OrdenPago"
Private Const SEP_COL As String = "|"
Private Const SEP_ROW As String = vbCrLf

Private m_Cargando As Boolean

' Columnas grillas (0-based)
Private Enum eColDeuda
    colDeudaConcepto = 0
    colDeudaImporte = 1
End Enum

Private Enum eColCheques
    colChequeBanco = 0
    colChequeNumero = 1
    colChequeFecha = 2
    colChequeImporte = 3
End Enum

Private Enum eColOtros
    colOtroConcepto = 0
    colOtroImporte = 1
End Enum

Private Sub Form_Load()
    On Error GoTo EH

    m_Cargando = True

    EnsureDatabaseReady
    EnsureSchemaOrdenPago

    InitGrids
    LimpiarFormulario False

    txtFecha.Text = Format$(Date, "dd/mm/yyyy")
    txtNroOrden.Text = CStr(GetSiguienteNumeroOrden())
    CargarComboReimpresion

    RecalcularTodo
    m_Cargando = False
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error en Form_Load: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdNuevo_Click()
    On Error GoTo EH
    LimpiarFormulario True
    Exit Sub
EH:
    MsgBox "Error en Nuevo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

' ------------------------------------------------------------
' Inicializacion UI
' ------------------------------------------------------------
Private Sub InitGrids()
    On Error GoTo EH

    ' Deuda
    With grdDeuda
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colDeudaConcepto) = "Concepto"
        .TextMatrix(0, colDeudaImporte) = "Importe"
        .ColWidth(colDeudaConcepto) = 3000
        .ColWidth(colDeudaImporte) = 1500
    End With

    ' Cheques
    With grdCheques
        .Rows = 2
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colChequeBanco) = "Banco"
        .TextMatrix(0, colChequeNumero) = "Numero"
        .TextMatrix(0, colChequeFecha) = "Fecha"
        .TextMatrix(0, colChequeImporte) = "Importe"
        .ColWidth(colChequeBanco) = 2200
        .ColWidth(colChequeNumero) = 1600
        .ColWidth(colChequeFecha) = 1300
        .ColWidth(colChequeImporte) = 1400
    End With

    ' Otros
    With grdOtros
        .Rows = 2
        .Cols = 2
        .FixedRows = 1
        .FixedCols = 0
        .TextMatrix(0, colOtroConcepto) = "Concepto"
        .TextMatrix(0, colOtroImporte) = "Importe"
        .ColWidth(colOtroConcepto) = 3000
        .ColWidth(colOtroImporte) = 1500
    End With

    Exit Sub
EH:
    MsgBox "Error inicializando grillas: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub LimpiarFormulario(ByVal nuevoNumero As Boolean)
    On Error GoTo EH

    m_Cargando = True

    txtProveedor.Text = ""
    txtFecha.Text = Format$(Date, "dd/mm/yyyy")
    txtEfectivo.Text = "0,00"
    SetRetIIBBText "0,00"

    txtSubDeuda.Text = "0,00"
    txtSubCheques.Text = "0,00"
    txtSubOtros.Text = "0,00"
    txtTotalPago.Text = "0,00"
    txtSaldo.Text = "0,00"
    txtImporteLetras.Text = ""

    ResetGridRows grdDeuda
    ResetGridRows grdCheques
    ResetGridRows grdOtros

    If nuevoNumero Then
        txtNroOrden.Text = CStr(GetSiguienteNumeroOrden())
    End If

    m_Cargando = False
    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = False
    MsgBox "Error limpiando formulario: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub ResetGridRows(ByRef g As MSFlexGrid)
    g.Rows = 2
    LimpiarFila g, 1
End Sub

Private Sub LimpiarFila(ByRef g As MSFlexGrid, ByVal rowIndex As Integer)
    Dim c As Integer
    If rowIndex < 1 Or rowIndex > g.Rows - 1 Then Exit Sub
    For c = 0 To g.Cols - 1
        g.TextMatrix(rowIndex, c) = ""
    Next c
End Sub

' ------------------------------------------------------------
' Botones agregar/eliminar filas
' ------------------------------------------------------------
Private Sub cmdAddDeuda_Click()
    AddRow grdDeuda
End Sub

Private Sub cmdDelDeuda_Click()
    DelCurrentRow grdDeuda
    RecalcularTodo
End Sub

Private Sub cmdAddCheque_Click()
    AddRow grdCheques
End Sub

Private Sub cmdDelCheque_Click()
    DelCurrentRow grdCheques
    RecalcularTodo
End Sub

Private Sub cmdAddOtro_Click()
    AddRow grdOtros
End Sub

Private Sub cmdDelOtro_Click()
    DelCurrentRow grdOtros
    RecalcularTodo
End Sub

Private Sub AddRow(ByRef g As MSFlexGrid)
    On Error GoTo EH
    g.Rows = g.Rows + 1
    g.Row = g.Rows - 1
    g.Col = 0
    Exit Sub
EH:
    MsgBox "Error agregando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub DelCurrentRow(ByRef g As MSFlexGrid)
    On Error GoTo EH

    If g.Row < 1 Then Exit Sub
    If g.Rows <= 2 Then
        LimpiarFila g, 1
        Exit Sub
    End If

    g.RemoveItem g.Row
    If g.Row > g.Rows - 1 Then g.Row = g.Rows - 1
    Exit Sub

EH:
    MsgBox "Error eliminando fila: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Edicion manual de celdas (MSFlexGrid no editable nativo)
' Doble click -> InputBox
' ------------------------------------------------------------
Private Sub grdDeuda_DblClick()
    EditarCeldaGrid grdDeuda
End Sub

Private Sub grdCheques_DblClick()
    EditarCeldaGrid grdCheques
End Sub

Private Sub grdOtros_DblClick()
    EditarCeldaGrid grdOtros
End Sub

Private Sub EditarCeldaGrid(ByRef g As MSFlexGrid)
    On Error GoTo EH

    Dim r As Integer, c As Integer
    Dim v As String

    r = g.Row
    c = g.Col

    If r < 1 Then Exit Sub

    v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
    g.TextMatrix(r, c) = Trim$(v)

    RecalcularTodo
    Exit Sub

EH:
    MsgBox "Error editando celda: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Recalculo automatico (sin bloquear edicion manual)
' ------------------------------------------------------------
Private Sub cmdRecalcular_Click()
    RecalcularTodo
End Sub

Private Sub cmdRestaurarAuto_Click()
    RecalcularTodo
End Sub

Private Sub txtEfectivo_Change()
    If m_Cargando Then Exit Sub
End Sub

Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
End Sub

Private Sub RecalcularTodo()
    On Error GoTo EH

    Dim subDeuda As Currency
    Dim subCheques As Currency
    Dim subOtros As Currency
    Dim efectivo As Currency
    Dim retIIBB As Currency
    Dim totalPago As Currency
    Dim saldo As Currency

    subDeuda = SumarColumna(grdDeuda, colDeudaImporte)
    subCheques = SumarColumna(grdCheques, colChequeImporte)
    subOtros = SumarColumna(grdOtros, colOtroImporte)

    efectivo = ParseCurrency(txtEfectivo.Text)
    retIIBB = ParseCurrency(GetRetIIBBText())

    totalPago = subCheques + subOtros + efectivo + retIIBB
    saldo = subDeuda - totalPago

    m_Cargando = True
    txtSubDeuda.Text = FormatMoney(subDeuda)
    txtSubCheques.Text = FormatMoney(subCheques)
    txtSubOtros.Text = FormatMoney(subOtros)
    txtTotalPago.Text = FormatMoney(totalPago)
    txtSaldo.Text = FormatMoney(saldo)
    txtImporteLetras.Text = NumeroALetrasSimple(totalPago)
    m_Cargando = False

    Exit Sub
EH:
    m_Cargando = False
    MsgBox "Error en recalculo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function SumarColumna(ByRef g As MSFlexGrid, ByVal colImporte As Integer) As Currency
    Dim r As Integer
    Dim t As Currency
    t = 0

    For r = 1 To g.Rows - 1
        t = t + ParseCurrency(g.TextMatrix(r, colImporte))
    Next r

    SumarColumna = t
End Function

' ------------------------------------------------------------
' Guardado
' ------------------------------------------------------------
Private Sub cmdGuardar_Click()
    On Error GoTo EH

    Dim nroOrden As Long
    Dim rs As DAO.Recordset
    Dim deudaTxt As String, chequesTxt As String, otrosTxt As String

    EnsureDatabaseReady

    nroOrden = CLng(Val(NzS(txtNroOrden.Text)))

    If nroOrden <= 0 Then
        MsgBox "Numero de orden invalido.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    deudaTxt = SerializarGrid(grdDeuda)
    chequesTxt = SerializarGrid(grdCheques)
    otrosTxt = SerializarGrid(grdOtros)

    Set rs = BaseSPC.OpenRecordset("SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden), dbOpenDynaset)
    If rs.EOF Then
        rs.AddNew
        rs!nroOrden = nroOrden
    Else
        rs.Edit
    End If

    rs!Fecha = ParseDateOrToday(txtFecha.Text)
    rs!Proveedor = NzS(txtProveedor.Text)

    rs!subDeuda = ParseCurrency(txtSubDeuda.Text)
    rs!subCheques = ParseCurrency(txtSubCheques.Text)
    rs!subOtros = ParseCurrency(txtSubOtros.Text)
    rs!efectivo = ParseCurrency(txtEfectivo.Text)
    rs!retIIBB = ParseCurrency(GetRetIIBBText())
    rs!totalPago = ParseCurrency(txtTotalPago.Text)
    rs!saldo = ParseCurrency(txtSaldo.Text)
    rs!ImporteLetras = NzS(txtImporteLetras.Text)

    rs!DetalleDeuda = deudaTxt
    rs!DetalleCheques = chequesTxt
    rs!DetalleOtros = otrosTxt

    rs!FechaAlta = Now
    rs.Update
    rs.Close
    Set rs = Nothing

    ActualizarCorrelativoOrden nroOrden
    CargarComboReimpresion

    MsgBox "Orden guardada correctamente.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then
        rs.Close
        Set rs = Nothing
    End If
    MsgBox "Error guardando orden: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Reimpresion / Carga
' ------------------------------------------------------------
Private Sub CargarComboReimpresion()
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim texto As String

    EnsureDatabaseReady

    cboBuscarOrden.Clear

    Set rs = BaseSPC.OpenRecordset("SELECT NroOrden, Fecha, Proveedor FROM OrdenPago ORDER BY NroOrden DESC", dbOpenSnapshot)
    Do While Not rs.EOF
        texto = CStr(NzL(rs!nroOrden)) & " - " & Format$(NzD(rs!Fecha), "dd/mm/yyyy") & " - " & NzS(rs!Proveedor)
        cboBuscarOrden.AddItem texto
        cboBuscarOrden.ItemData(cboBuscarOrden.NewIndex) = CLng(NzL(rs!nroOrden))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    Exit Sub
EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error cargando combo de ordenes: " & CStr(Err.Number) & " - " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdCargarOrden_Click()
    On Error GoTo EH

    If cboBuscarOrden.ListIndex < 0 Then
        MsgBox "Seleccione una orden.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    CargarOrdenPorNumero CLng(cboBuscarOrden.ItemData(cboBuscarOrden.ListIndex))
    Exit Sub

EH:
    MsgBox "Error cargando orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub CargarOrdenPorNumero(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim rs As DAO.Recordset
    Dim sql As String

    EnsureDatabaseReady

    sql = "SELECT * FROM OrdenPago WHERE NroOrden=" & CStr(nroOrden)
    Set rs = BaseSPC.OpenRecordset(sql, dbOpenSnapshot)

    If rs.EOF Then
        rs.Close
        Set rs = Nothing
        MsgBox "No se encontro la orden seleccionada.", vbExclamation, "Orden de Pago"
        Exit Sub
    End If

    m_Cargando = True

    txtNroOrden.Text = CStr(NzL(rs!nroOrden))
    txtFecha.Text = Format$(NzD(rs!Fecha), "dd/mm/yyyy")
    txtProveedor.Text = NzS(rs!Proveedor)

    txtSubDeuda.Text = FormatMoney(NzC(rs!subDeuda))
    txtSubCheques.Text = FormatMoney(NzC(rs!subCheques))
    txtSubOtros.Text = FormatMoney(NzC(rs!subOtros))
    txtEfectivo.Text = FormatMoney(NzC(rs!efectivo))
    SetRetIIBBText FormatMoney(NzC(rs!retIIBB))
    txtTotalPago.Text = FormatMoney(NzC(rs!totalPago))
    txtSaldo.Text = FormatMoney(NzC(rs!saldo))
    txtImporteLetras.Text = NzS(rs!ImporteLetras)

    DeserializarGrid grdDeuda, NzS(rs!DetalleDeuda)
    DeserializarGrid grdCheques, NzS(rs!DetalleCheques)
    DeserializarGrid grdOtros, NzS(rs!DetalleOtros)

    m_Cargando = False

    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    m_Cargando = False
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "Error al cargar orden: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Sub cmdReimprimir_Click()
    cmdImprimir_Click
End Sub

' ------------------------------------------------------------
' Impresion (stub funcional)
' ------------------------------------------------------------
Private Sub cmdImprimir_Click()
    On Error GoTo EH

    Printer.Print "==========================================="
    Printer.Print "ORDEN DE PAGO Nro: " & NzS(txtNroOrden.Text)
    Printer.Print "Fecha: " & NzS(txtFecha.Text)
    Printer.Print "Proveedor: " & NzS(txtProveedor.Text)
    Printer.Print "-------------------------------------------"
    Printer.Print "SubTotal Deuda:   " & NzS(txtSubDeuda.Text)
    Printer.Print "SubTotal Cheques: " & NzS(txtSubCheques.Text)
    Printer.Print "SubTotal Otros:   " & NzS(txtSubOtros.Text)
    Printer.Print "Efectivo:         " & NzS(txtEfectivo.Text)
    Printer.Print "Ret IIBB:         " & NzS(GetRetIIBBText())
    Printer.Print "TOTAL PAGO:       " & NzS(txtTotalPago.Text)
    Printer.Print "SALDO:            " & NzS(txtSaldo.Text)
    Printer.Print "Importe en letras: " & NzS(txtImporteLetras.Text)
    Printer.Print "==========================================="
    Printer.EndDoc

    MsgBox "Impresion enviada (stub).", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    MsgBox "Error al imprimir: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

' ------------------------------------------------------------
' Numeracion correlativa
' ------------------------------------------------------------
Private Function GetSiguienteNumeroOrden() As Long
    On Error GoTo EH

    Dim idSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim ultimo As Long

    EnsureDatabaseReady

    idSucursal = GetIdSucursalActual()

    sql = "SELECT UltimoNumero FROM UltimosNumeros WHERE IdSucursal=" & CStr(idSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!idSucursal = idSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = 0
        rs.Update
        ultimo = 0
    Else
        ultimo = NzL(rs!UltimoNumero)
    End If

    rs.Close
    Set rs = Nothing

    GetSiguienteNumeroOrden = ultimo + 1
    Exit Function

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    GetSiguienteNumeroOrden = 1
End Function

Private Sub ActualizarCorrelativoOrden(ByVal nroOrden As Long)
    On Error GoTo EH

    Dim idSucursal As Long
    Dim rs As DAO.Recordset
    Dim sql As String

    EnsureDatabaseReady

    idSucursal = GetIdSucursalActual()
    sql = "SELECT * FROM UltimosNumeros WHERE IdSucursal=" & CStr(idSucursal) & _
          " AND IDTabla='" & ID_TABLA_ORDENPAGO & "'"

    Set rs = BaseSPC.OpenRecordset(sql, dbOpenDynaset)

    If rs.EOF Then
        rs.AddNew
        rs!idSucursal = idSucursal
        rs!IDTabla = ID_TABLA_ORDENPAGO
        rs!UltimoNumero = nroOrden
    Else
        If NzL(rs!UltimoNumero) < nroOrden Then
            rs.Edit
            rs!UltimoNumero = nroOrden
        Else
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If

    rs.Update
    rs.Close
    Set rs = Nothing
    Exit Sub

EH:
    On Error Resume Next
    If Not rs Is Nothing Then rs.Close
    MsgBox "No se pudo actualizar correlativo: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function GetIdSucursalActual() As Long
    GetIdSucursalActual = 1
End Function

' ------------------------------------------------------------
' Serializacion detalles
' ------------------------------------------------------------
Private Function SerializarGrid(ByRef g As MSFlexGrid) As String
    Dim r As Integer, c As Integer
    Dim linea As String
    Dim salida As String

    salida = ""
    For r = 1 To g.Rows - 1
        If FilaTieneDatos(g, r) Then
            linea = ""
            For c = 0 To g.Cols - 1
                If c > 0 Then linea = linea & SEP_COL
                linea = linea & EscaparDato(g.TextMatrix(r, c))
            Next c
            If Len(salida) > 0 Then salida = salida & SEP_ROW
            salida = salida & linea
        End If
    Next r

    SerializarGrid = salida
End Function

Private Sub DeserializarGrid(ByRef g As MSFlexGrid, ByVal dataText As String)
    On Error GoTo EH

    Dim rowsArr() As String
    Dim colsArr() As String
    Dim r As Long, c As Long
    Dim targetRow As Integer

    ResetGridRows g

    If Trim$(dataText) = "" Then Exit Sub

    rowsArr = Split(dataText, SEP_ROW)
    g.Rows = UBound(rowsArr) + 2

    targetRow = 1
    For r = LBound(rowsArr) To UBound(rowsArr)
        colsArr = Split(rowsArr(r), SEP_COL)
        For c = 0 To g.Cols - 1
            If c <= UBound(colsArr) Then
                g.TextMatrix(targetRow, c) = DesescaparDato(colsArr(c))
            Else
                g.TextMatrix(targetRow, c) = ""
            End If
        Next c
        targetRow = targetRow + 1
    Next r

    Exit Sub
EH:
    MsgBox "Error deserializando grilla: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function FilaTieneDatos(ByRef g As MSFlexGrid, ByVal rowIndex As Integer) As Boolean
    Dim c As Integer
    For c = 0 To g.Cols - 1
        If Trim$(g.TextMatrix(rowIndex, c)) <> "" Then
            FilaTieneDatos = True
            Exit Function
        End If
    Next c
    FilaTieneDatos = False
End Function

Private Function EscaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\", "\\")
    t = Replace(t, SEP_COL, "\p")
    t = Replace(t, vbCr, "")
    t = Replace(t, vbLf, "\n")
    EscaparDato = t
End Function

Private Function DesescaparDato(ByVal s As String) As String
    Dim t As String
    t = Replace(s, "\n", vbLf)
    t = Replace(t, "\p", SEP_COL)
    t = Replace(t, "\\", "\")
    DesescaparDato = t
End Function

' ------------------------------------------------------------
' Schema fallback
' ------------------------------------------------------------
Private Sub EnsureSchemaOrdenPago()
    On Error GoTo EH

    EnsureDatabaseReady

    If Not TablaExiste("OrdenPago") Then
        BaseSPC.Execute _
            "CREATE TABLE OrdenPago (" & _
            "NroOrden LONG CONSTRAINT PK_OrdenPago PRIMARY KEY, " & _
            "Fecha DATETIME, " & _
            "Proveedor TEXT(150), " & _
            "SubDeuda CURRENCY, " & _
            "SubCheques CURRENCY, " & _
            "SubOtros CURRENCY, " & _
            "Efectivo CURRENCY, " & _
            "RetIIBB CURRENCY, " & _
            "TotalPago CURRENCY, " & _
            "Saldo CURRENCY, " & _
            "ImporteLetras MEMO, " & _
            "DetalleDeuda MEMO, " & _
            "DetalleCheques MEMO, " & _
            "DetalleOtros MEMO, " & _
            "FechaAlta DATETIME" & _
            ")"
    End If

    Exit Sub
EH:
    MsgBox "Error creando/verificando tabla OrdenPago: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub

Private Function TablaExiste(ByVal nombreTabla As String) As Boolean
    On Error GoTo EH
    Dim tdf As DAO.TableDef
    For Each tdf In BaseSPC.TableDefs
        If StrComp(tdf.Name, nombreTabla, vbTextCompare) = 0 Then
            TablaExiste = True
            Exit Function
        End If
    Next tdf
    TablaExiste = False
    Exit Function
EH:
    TablaExiste = False
End Function

' ------------------------------------------------------------
' Utilitarios parse/format
' ------------------------------------------------------------
Private Function ParseCurrency(ByVal s As String) As Currency
    Dim t As String
    t = Trim$(s)
    If t = "" Then
        ParseCurrency = 0
        Exit Function
    End If

    t = Replace(t, ".", "")
    t = Replace(t, ",", ".")
    ParseCurrency = CCur(Val(t))
End Function

Private Function FormatMoney(ByVal v As Currency) As String
    FormatMoney = Format$(v, "0.00")
    FormatMoney = Replace(FormatMoney, ".", ",")
End Function

Private Function NzS(ByVal v As Variant) As String
    If IsNull(v) Then
        NzS = ""
    Else
        NzS = Trim$(CStr(v))
    End If
End Function

Private Function NzL(ByVal v As Variant) As Long
    If IsNull(v) Or v = "" Then
        NzL = 0
    Else
        NzL = CLng(v)
    End If
End Function

Private Function NzC(ByVal v As Variant) As Currency
    If IsNull(v) Or v = "" Then
        NzC = 0
    Else
        NzC = CCur(v)
    End If
End Function

Private Function NzD(ByVal v As Variant) As Date
    If IsNull(v) Or v = "" Then
        NzD = Date
    Else
        NzD = CDate(v)
    End If
End Function

Private Function ParseDateOrToday(ByVal s As String) As Date
    On Error GoTo EH
    ParseDateOrToday = CDate(s)
    Exit Function
EH:
    ParseDateOrToday = Date
End Function

Private Function NumeroALetrasSimple(ByVal importe As Currency) As String
    NumeroALetrasSimple = "SON PESOS " & FormatMoney(importe)
End Function

' ------------------------------------------------------------
' Helpers de compatibilidad (control opcional txtRetIIBB)
' ------------------------------------------------------------
Private Sub EnsureDatabaseReady()
    On Error GoTo EH

    If BaseSPC Is Nothing Then
        Set BaseSPC = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    End If
    Exit Sub
EH:
    MsgBox "No se pudo abrir la base de datos: " & Err.Description, vbCritical, "Orden de Pago"
End Sub

Private Function TieneControl(ByVal nombreControl As String) As Boolean
    On Error GoTo EH
    Dim ctl As Control
    Set ctl = Me.Controls(nombreControl)
    TieneControl = Not (ctl Is Nothing)
    Exit Function
EH:
    TieneControl = False
End Function

Private Function GetRetIIBBText() As String
    On Error GoTo EH
    If TieneControl("txtRetIIBB") Then
        GetRetIIBBText = CStr(Me.Controls("txtRetIIBB").Text)
    Else
        GetRetIIBBText = "0,00"
    End If
    Exit Function
EH:
    GetRetIIBBText = "0,00"
End Function

Private Sub SetRetIIBBText(ByVal valor As String)
    On Error Resume Next
    If TieneControl("txtRetIIBB") Then
        Me.Controls("txtRetIIBB").Text = valor
    End If
End Sub

¡Excelente avance! 🙌  
Perfecto, hacemos esos 4 ajustes.

Te paso **solo cambios puntuales** para pegar (adaptá nombres que te queden distinto, porque vos cambiaste algunos).

---

### 1) Enter como Tab en el formulario

**Propiedad en diseño:**
- `FormOrdenPago.KeyPreview = True`

**Código:**
```vb
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Dim c As Control
        Set c = Me.ActiveControl

        If Not c Is Nothing Then
            Select Case TypeName(c)
                Case "MSFlexGrid", "MshFlexGrid", "VSFlexGrid"
                    ' En grilla se maneja aparte
                Case Else
                    KeyCode = 0
                    SendKeys "{TAB}"
            End Select
        End If
    End If
End Sub
```

---

### 2) En grillas, editar con Enter (en vez de doble click)

```vb
Private Sub grdDeuda_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdDeuda
    End If
End Sub

Private Sub grdCheques_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdCheques
    End If
End Sub

Private Sub grdOtros_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        EditarCeldaGrid grdOtros
    End If
End Sub
```

> Como ya tenés `EditarCeldaGrid`, esto te reutiliza lo mismo del doble click.

---

### 3) `cmdRestaurarAuto_Click` debe limpiar pantalla y dejar nuevo número

**Reemplazá tu evento actual por:**
```vb
Private Sub cmdRestaurarAuto_Click()
    On Error GoTo EH

    LimpiarFormulario True   ' True => limpia + pide siguiente numero
    txtProveedor.SetFocus
    Exit Sub

EH:
    MsgBox "Error al restaurar: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
```

---

### 4) Nuevo botón para impresión “parecida a la foto”

Agregá un botón (por ejemplo `cmdImprimirFormato`) y pegá este evento.  
(usa nombres de tus controles reales; abajo puse los más comunes de tu código).

```vb
Private Sub cmdImprimir_Click()
    On Error GoTo EH

    Dim y As Single
    Dim i As Integer
    Dim imp As String

    Printer.ScaleMode = vbTwips
    Printer.FontName = "Courier New"
    Printer.FontSize = 9

    y = 400

    Printer.CurrentX = 1200: Printer.CurrentY = y
    Printer.FontBold = True
    Printer.Print "PAGO A PROVEEDORES"
    Printer.FontBold = False

    y = y + 350
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Proveedor: " & txtProveedor.Text
    Printer.CurrentX = 6500: Printer.CurrentY = y: Printer.Print "Fecha: " & txtFecha.Text

    y = y + 350
    Printer.Line (300, y)-(10500, y)
    y = y + 150

    ' DEUDA
    Printer.FontBold = True
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "DEUDA"
    Printer.FontBold = False
    y = y + 220
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Detalle"
    Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print "Importe"
    y = y + 180
    Printer.Line (300, y)-(5000, y)
    y = y + 120

    For i = 1 To grdDeuda.Rows - 1
        If Trim$(grdDeuda.TextMatrix(i, 0)) <> "" Then
            Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print Left$(grdDeuda.TextMatrix(i, 0), 30)
            imp = grdDeuda.TextMatrix(i, 1) ' ajusta indice si cambia
            Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print imp
            y = y + 200
        End If
    Next i

    y = y + 100
    Printer.CurrentX = 300:  Printer.CurrentY = y: Printer.Print "Total Deuda:"
    Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print txtSubDeuda.Text

    ' CHEQUES
    y = y + 350
    Printer.FontBold = True
    Printer.CurrentX = 5600: Printer.CurrentY = y - 350: Printer.Print "CHEQUES"
    Printer.FontBold = False
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Nro"
    Printer.CurrentX = 7000: Printer.CurrentY = y: Printer.Print "Banco"
    Printer.CurrentX = 8200: Printer.CurrentY = y: Printer.Print "Vto"
    Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print "Importe"
    y = y + 180
    Printer.Line (5600, y)-(10500, y)
    y = y + 120

    For i = 1 To grdCheques.Rows - 1
        If Trim$(grdCheques.TextMatrix(i, 0)) <> "" Or Trim$(grdCheques.TextMatrix(i, 1)) <> "" Then
            Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 1) ' nro
            Printer.CurrentX = 7000: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 0) ' banco
            Printer.CurrentX = 8200: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 2) ' vto
            Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print grdCheques.TextMatrix(i, 3) ' importe
            y = y + 200
        End If
    Next i

    y = y + 100
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Subtotal Cheques:"
    Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print txtSubCheques.Text

    ' OTROS + TOTALES
    y = y + 300
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Subtotal Otros: " & txtSubOtros.Text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Efectivo: " & txtEfectivo.Text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Retencion IIBB: " & GetRetIIBBText()
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Total Pago: " & txtTotalPago.Text
    y = y + 220
    Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Saldo: " & txtSaldo.Text

    y = y + 350
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "Recibi la suma de: " & txtImporteLetras.Text

    y = y + 600
    Printer.CurrentX = 300: Printer.CurrentY = y: Printer.Print "Firma: ____________________________"

    Printer.EndDoc
    MsgBox "Impresion enviada.", vbInformation, "Orden de Pago"
    Exit Sub

EH:
    MsgBox "Error al imprimir formato: " & Err.Description, vbExclamation, "Orden de Pago"
End Sub
```

---

Si querés, en el próximo mensaje te paso una **versión 2 del formato de impresión** todavía más parecida al papel (con “cajas”/líneas y mejor alineación), ya ajustada a los **índices reales** de tus grillas.

Perfecto, ahí está el problema: el `Change` de `txtEfectivo` no estaba recalculando.

Pegá **esto** reemplazando tus eventos:

```vb
Private Sub txtEfectivo_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub
```

```vb
Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub
```

Opcional (recomendado): formatear `txtEfectivo` al salir del campo:

```vb
Private Sub txtEfectivo_LostFocus()
    Dim v As Double

    If m_Cargando Then Exit Sub

    On Error GoTo EH
    If Trim$(txtEfectivo.Text) = "" Then
        v = 0
    Else
        v = CDbl(txtEfectivo.Text)
    End If

    m_Cargando = True
    txtEfectivo.Text = Format$(v, "0.00")
    m_Cargando = False

    RecalcularTodo
    Exit Sub

EH:
    m_Cargando = True
    txtEfectivo.Text = "0.00"
    m_Cargando = False
    RecalcularTodo
End Sub

```

Prueba rápida:
1) cargá deuda,  
2) escribí importe en `txtEfectivo`,  
3) verificá que `txtTotalPago` sube y `txtSaldo` baja en el momento.  

Si querés, después te lo adapto para aceptar coma decimal argentina (`0,00`) sin problemas.

¡Quedó muy bien tu avance! 👏  
Para cerrar lo de formato europeo en **todo** (textbox + grilla + impresión), pegá estos bloques.

---

### 1) Reemplazá tus utilitarios `ParseCurrency` y `FormatMoney` por estos

```vb
Private Function ParseCurrency(ByVal s As String) As Currency
    Dim t As String
    Dim pDot As Long, pCom As Long, decSep As String
    Dim intPart As String, decPart As String
    Dim sign As Double
    
    t = Trim$(s)
    If Len(t) = 0 Then
        ParseCurrency = 0
        Exit Function
    End If
    
    t = Replace$(t, " ", "")
    t = Replace$(t, "$", "")
    
    sign = 1
    If Left$(t, 1) = "-" Then
        sign = -1
        t = Mid$(t, 2)
    ElseIf Left$(t, 1) = "+" Then
        t = Mid$(t, 2)
    End If
    
    pDot = InStrRev(t, ".")
    pCom = InStrRev(t, ",")
    
    If pDot > 0 And pCom > 0 Then
        If pDot > pCom Then
            decSep = "."
        Else
            decSep = ","
        End If
    ElseIf pDot > 0 Then
        If Len(t) - pDot = 3 Then
            decSep = ""    ' miles
        Else
            decSep = "."   ' decimal
        End If
    ElseIf pCom > 0 Then
        If Len(t) - pCom = 3 Then
            decSep = ""    ' miles
        Else
            decSep = ","   ' decimal
        End If
    Else
        decSep = ""
    End If
    
    If decSep = "" Then
        t = Replace$(t, ".", "")
        t = Replace$(t, ",", "")
        If Len(t) = 0 Then t = "0"
        ParseCurrency = CCur(sign * Val(t))
        Exit Function
    End If
    
    If decSep = "." Then
        intPart = Left$(t, InStrRev(t, ".") - 1)
        decPart = Mid$(t, InStrRev(t, ".") + 1)
    Else
        intPart = Left$(t, InStrRev(t, ",") - 1)
        decPart = Mid$(t, InStrRev(t, ",") + 1)
    End If
    
    intPart = Replace$(intPart, ".", "")
    intPart = Replace$(intPart, ",", "")
    If Len(intPart) = 0 Then intPart = "0"
    
    If Len(decPart) = 0 Then
        decPart = "00"
    ElseIf Len(decPart) = 1 Then
        decPart = decPart & "0"
    ElseIf Len(decPart) > 2 Then
        decPart = Left$(decPart, 2)
    End If
    
    ParseCurrency = CCur(sign * Val(intPart & "." & decPart))
End Function

Private Function GroupThousands(ByVal digits As String) As String
    Dim i As Long, out As String, cnt As Long
    
    out = ""
    cnt = 0
    
    For i = Len(digits) To 1 Step -1
        out = Mid$(digits, i, 1) & out
        cnt = cnt + 1
        If cnt = 3 And i > 1 Then
            out = "." & out
            cnt = 0
        End If
    Next i
    
    GroupThousands = out
End Function

Private Function FormatMoney(ByVal v As Currency) As String
    Dim isNeg As Boolean
    Dim absV As Currency
    Dim cents As Long
    Dim intPart As Long
    Dim decPart As Long
    Dim intTxt As String
    
    isNeg = (v < 0)
    absV = Abs(v)
    
    cents = CLng((absV * 100) + 0.5)
    intPart = cents \ 100
    decPart = cents Mod 100
    
    intTxt = GroupThousands(CStr(intPart))
    
    If isNeg Then
        FormatMoney = "-" & intTxt & "," & Right$("0" & CStr(decPart), 2)
    Else
        FormatMoney = intTxt & "," & Right$("0" & CStr(decPart), 2)
    End If
End Function

Private Sub FormatImporteTextBox(ByRef txt As TextBox)
    txt.Text = FormatMoney(ParseCurrency(txt.Text))
End Sub
```

---

### 2) Corregí `txtRetIIBB_Change` (en tu código quedó sin recalcular)

```vb
Private Sub txtRetIIBB_Change()
    If m_Cargando Then Exit Sub
    RecalcularTodo
End Sub
```

---

### 3) LostFocus para todos los importes

```vb
Private Sub txtEfectivo_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtEfectivo
    RecalcularTodo
End Sub

Private Sub txtSubDeuda_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubDeuda
End Sub

Private Sub txtSubCheques_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubCheques
End Sub

Private Sub txtSubOtros_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSubOtros
End Sub

Private Sub txtTotalPago_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtTotalPago
End Sub

Private Sub txtSaldo_LostFocus()
    If m_Cargando Then Exit Sub
    FormatImporteTextBox txtSaldo
End Sub

Private Sub txtRetIIBB_LostFocus()
    If m_Cargando Then Exit Sub
    If TieneControl("txtRetIIBB") Then
        Me.Controls("txtRetIIBB").Text = FormatMoney(ParseCurrency(Me.Controls("txtRetIIBB").Text))
    End If
    RecalcularTodo
End Sub
```

---

### 4) Formatear importes en grilla al editar celda

En tu `EditarCeldaGrid`, reemplazá la parte final por esto:

```vb
v = InputBox$("Editar valor:", "Editar celda", g.TextMatrix(r, c))
v = Trim$(v)

If g.Name = "grdDeuda" And c = colDeudaImporte Then
    g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
ElseIf g.Name = "grdCheques" And c = colChequeImporte Then
    g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
ElseIf g.Name = "grdOtros" And c = colOtroImporte Then
    g.TextMatrix(r, c) = FormatMoney(ParseCurrency(v))
Else
    g.TextMatrix(r, c) = v
End If
```

---

### 5) Impresión: siempre formatear con función

Donde imprimís importes, usá patrón:

```vb
Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.Text))
```

Ejemplos concretos para reemplazar en `cmdImprimir_Click`:

```vb
Printer.CurrentX = 3800: Printer.CurrentY = y: Printer.Print FormatMoney(ParseCurrency(txtSubDeuda.Text))
...
Printer.CurrentX = 9400: Printer.CurrentY = y: Printer.Print FormatMoney(ParseCurrency(txtSubCheques.Text))
...
Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Efectivo: " & FormatMoney(ParseCurrency(txtEfectivo.Text))
Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Retencion IIBB: " & FormatMoney(ParseCurrency(GetRetIIBBText()))
Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Total Pago: " & FormatMoney(ParseCurrency(txtTotalPago.Text))
Printer.CurrentX = 5600: Printer.CurrentY = y: Printer.Print "Saldo: " & FormatMoney(ParseCurrency(txtSaldo.Text))
```

Y para grillas:
```vb
impi = FormatMoney(ParseCurrency(grdDeuda.TextMatrix(i, 1)))
...
Printer.Print FormatMoney(ParseCurrency(grdCheques.TextMatrix(i, 3)))
```

---

Si querés, en el próximo paso te paso **tu `cmdImprimir_Click` ya reescrito completo** con estos formatos aplicados línea por línea (copiar/pegar directo).