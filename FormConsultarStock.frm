VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormConsultarStock 
   Caption         =   "CONSULTA DE STOCK"
   ClientHeight    =   6075
   ClientLeft      =   225
   ClientTop       =   555
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   ScaleHeight     =   6075
   ScaleWidth      =   11730
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   11415
      Begin VB.CommandButton Imprime 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   6
         Top             =   4440
         Width           =   4575
      End
      Begin VB.CommandButton btnSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5760
         TabIndex        =   5
         Top             =   4440
         Width           =   4575
      End
      Begin VB.ComboBox cmbOrigen 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton cmdVerStock 
         Caption         =   "&Ver Stock"
         Height          =   495
         Left            =   5520
         TabIndex        =   1
         Top             =   720
         Width           =   1815
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2775
         Left            =   360
         TabIndex        =   3
         Top             =   1440
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   4895
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   0
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Depósito a Consultar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   600
         TabIndex        =   4
         Top             =   360
         Width           =   1770
      End
   End
End
Attribute VB_Name = "FormConsultarStock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstProductos As DAO.Recordset
 Dim cantidadProducto As Integer

Public Function BuscarDescProd(IdCodProd As String) As String

    Set tP = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    tP.Index = "PrimaryKey"
    tP.Seek "=", IdCodProd
    
    If Not tP.NoMatch Then
        BuscarDescProd = tP!Descripcion
    End If
    
    tP.Close

End Function

Private Sub DesHagoStock()
    'Sumo el Stock en Depósito Destino
        Set tS = BaseSPC.OpenRecordset("Stock", dbOpenTable)
        
        tS.Index = "PrimaryKey"
        tS.MoveFirst
        
        'Resto el Stock en Depósito Origen
          tS.Seek "=", CodProd, IdDepoOrigen
            
        If Not tS.NoMatch Then
            tS.Edit
                tS!CodProd = CodProd
                tS!IDDEPOSITO = IdDepoOrigen
                tS!cantidad = tS.cantidad - FormatNumber(Cant, 2)
                tS!FechaUM = Format(Date, "DD/MM/YYYY")
            tS.Update
        End If
    
    'Sumo el Stock en Depósito Destino
        tS.Seek "=", CodProd, IdDepoDestino
              
        'Si tiene stock de este producto
            If Not tS.NoMatch Then
                'CantIni = tSotck!Stock
                tS.Edit
                    tS.CodProd = CodProd
                    tS.IDDEPOSITO = IdDepoDestino
                    tS.cantidad = tS!cantidad + FormatNumber(Cant, 2)
                    tS.FechaUM = Format(Date, "DD/MM/YYYY")
                tS.Update
        'Si no tiene stock de este producto
             Else
                tS.AddNew
                    tS!CodProd = CodProd
                    tS!IDDEPOSITO = IdDepoDestino
                    tS!cantidad = FormatNumber(Cant, 2)
                    tS!FechaUM = Format(Date, "DD/MM/YYYY")
                tS.Update
            End If

End Sub

Private Sub DisabledAll()
    
    txtIdMov.Enabled = False
    txtFecha.Enabled = False
    cmbOrigen.Enabled = False
    cmbDestino.Enabled = False
    FG1.Enabled = False
    btnEliminar.Enabled = False
    btnModificar.Enabled = False
    btnGrabar.Enabled = True

End Sub

Private Sub EnabledAll()
    
    txtIdMov.Enabled = True
    txtFecha.Enabled = True
    cmbOrigen.Enabled = True
    cmbDestino.Enabled = True
    FG1.Enabled = True
    btnEliminar.Enabled = True
    btnModificar.Enabled = True
    If btnGrabar.Enabled = False Then btnGrabar.Enabled = True
    
End Sub

Private Sub Mostrar()

 'Seteo la grilla
    FG1.Rows = 2
    FG1.Row = 0
    
    FG1.Col = 0
    FG1.ColWidth(0) = 500
    FG1.Text = "Item"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 1800
    FG1.Text = "Código Producto"
    
    FG1.Col = 2
    FG1.ColWidth(2) = 6000
    FG1.Text = "Descripción"
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1800
    FG1.Text = "Cantidad"

    txtIdMov.Text = tMovIntStockC!IdMovIntStock
    txtFecha.Text = tMovIntStockC!FechaMovInt
    
    Set tDepositos = BaseSPC.OpenRecordset("Depositos", dbOpenTable)
    tDepositos.Index = "PrimaryKey"
    tDepositos.Seek "=", tMovIntStockC!IdDepoOrigen
    
    If Not tDepositos.NoMatch Then
        cmbOrigen.Text = tDepositos!Descripcion
    End If
    
    tDepositos.Seek "=", tMovIntStockC!IdDepoDest
    
    If Not tDepositos.NoMatch Then
        cmbDestino.Text = tDepositos!Descripcion
    End If
       
    'tDepositos.Close
        
    vSQL = "SELECT * FROM MovIntStockD WHERE IdMovInt =" & tMovIntStockC!IdMovIntStock & " ORDER BY ItemMov"
    
    Set tMovIntStockD = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    tMovIntStockD.MoveFirst
    
    FG1.Row = 1
    While Not tMovIntStockD.EOF
        FG1.Col = 0
        FG1.Text = tMovIntStockD!ItemMov
        FG1.Col = 1
        FG1.Text = tMovIntStockD!IdCodProd
        FG1.Col = 2
        FG1.Text = BuscarDescProd(tMovIntStockD!IdCodProd)
        FG1.Col = 3
        FG1.Text = tMovIntStockD!cantidad
    
        tMovIntStockD.MoveNext
        FG1.Col = 0
        FG1.Rows = FG1.Rows + 1
        FG1.Row = FG1.Row + 1
    Wend
    
    tMovIntStockD.Close
    
    Call DisabledAll
    
End Sub

Private Sub SeteoGrilla()
    
    FG1.Cols = 4
    FG1.Row = 0
    
    FG1.Col = 0
    FG1.ColWidth(0) = 2800
    FG1.Text = "Producto"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 3800
    FG1.Text = "Descripcion"
    
    FG1.Col = 2
    FG1.ColWidth(2) = 1800
    FG1.Text = "Cantidad"
    
    FG1.Col = 3
    FG1.ColWidth(2) = 2000
    FG1.Text = "Fecha UM"


End Sub

Private Sub btnAdelante_Click()

    On Error GoTo CapturaErrores
        FG1.Clear
        
        If Not tMovIntStockC.EOF Then
            tMovIntStockC.MoveNext
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "Ultimo Registro !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "Ultimo Registro !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select


End Sub

Private Sub btnAtras_Click()
    
    On Error GoTo CapturaErrores
        FG1.Clear
        
        If Not tMovIntStockC.BOF Then
            tMovIntStockC.MovePrevious
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "Primer Registro !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "Primer Registro !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select

End Sub

Private Sub btnBuscar_Click()

    If txtIdMov.Text <> "" Then
        Set tMovIntStockC = BaseSPC.OpenRecordset("MovIntStockC", dbOpenTable)
        tMovIntStockC.Index = "PrimaryKey"
        tMovIntStockC.Seek "=", txtIdMov.Text
    
        If Not tMovIntStockC.NoMatch Then
            Call Mostrar
         Else
           MsgBox "Registro Buscado Inexistente !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
        End If
        
    Else
        MsgBox "Debe Ingresar El Nro de Movimiento a Buscar !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
        txtIdMov.SetFocus
    End If

End Sub

Private Sub btnEliminar_Click()

    Rta = MsgBox("¿ Seguro Elimina el Registro Actual ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")

    If Rta = 1 Then
        
        Set tStock = BaseSPC.OpenRecordset("Stock", dbOpenTable)
    
        'Guardo los codigos de los depositos
            tDepositos.Index = "IndDesc"
        
            tDepositos.Seek "=", cmbOrigen.Text
            If Not tDepositos.NoMatch Then IdDepoOrigen = tDepositos!IDDEPOSITO
        
            tDepositos.Seek "=", cmbDestino.Text
            If Not tDepositos.NoMatch Then IdDepoDestino = tDepositos!IDDEPOSITO
            
            
            
            'Call ActualizarStock(CodProd, IdDepoOrigen, IdDepoDestino, Cant)
        
            tMovIntStockC.Delete
        Call EnabledAll
        Call LimpiarPantalla
    End If

End Sub

Private Sub btnGrabar_Click()

 On Error GoTo CapturaErrores
 
 Rta = MsgBox("¿ Seguro Genera Nuevo Registro ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
 If Rta = 1 Then
    Dim Filas As Integer
    Dim Columna As Integer
    
    Set tMovIntStockC = BaseSPC.OpenRecordset("MovIntStockC", dbOpenTable)
    Set tMovIntStockD = BaseSPC.OpenRecordset("MovIntStockD", dbOpenTable)
    Set tUltimosNumeros = BaseSPC.OpenRecordset("UltimosNumeros", dbOpenTable)
    Set tStock = BaseSPC.OpenRecordset("Stock", dbOpenTable)
    
    'Guardo los codigos de los depositos
        tDepositos.Index = "IndDesc"
    
        tDepositos.Seek "=", cmbOrigen.Text
        If Not tDepositos.NoMatch Then IdDepoOrigen = tDepositos!IDDEPOSITO
    
        tDepositos.Seek "=", cmbDestino.Text
        If Not tDepositos.NoMatch Then IdDepoDestino = tDepositos!IDDEPOSITO
        
    'Grabo La Cabecera y el Detalle de los movimientos
        tMovIntStockC.AddNew
            tMovIntStockC!IdMovIntStock = txtIdMov.Text
            tMovIntStockC!FechaMovInt = txtFecha.Text
            tMovIntStockC!IdDepoOrigen = IdDepoOrigen
            tMovIntStockC!IdDepoDest = IdDepoDestino
        tMovIntStockC.Update
    

            Filas = FG1.Rows
            Columnas = FG1.Cols
            For I = 1 To Filas
                FG1.Row = I
                FG1.Col = 1
                If FG1.Text = "" Then Exit For
                tMovIntStockD.AddNew
                    FG1.Col = 0
                    tMovIntStockD!IdMovInt = txtIdMov.Text
                    tMovIntStockD!ItemMov = FG1.Text
                    FG1.Col = 1
                    tMovIntStockD!IdCodProd = FG1.Text
                    CodProd = FG1.Text
                    FG1.Col = 3
                    tMovIntStockD!cantidad = FG1.Text
                    Cant = FG1.Text
                    
                  'Actualizo Stock, Origen y Destino
                    'Call ActualizarStock(CodProd, IdDepoOrigen, IdDepoDestino, Cant)
                tMovIntStockD.Update
            Next I
        
        'Actualizar Ultimos Numeros
            tUltimosNumeros.Index = "PrimaryKey"
            tUltimosNumeros.MoveFirst
            tUltimosNumeros.Seek "=", "tMovIntStockC"
            
            If Not tUltimosNumeros.NoMatch Then
                tUltimosNumeros.Edit
                    tUltimosNumeros!UltimoNumero = tUltimosNumeros!UltimoNumero + 1
                    txtIdMov.Text = txtIdMov.Text + 1
                tUltimosNumeros.Update
            End If
            
            Call LimpiarPantalla
            txtIdMov.SetFocus
  Else
    txtIdMov.SetFocus
 End If

CapturaErrores:
    Select Case Err
        Case 3022
            MsgBox "Nro de Movimiento Interno de Stock YA EXISTE por favor Verifique !!!", vbCritical + vbDefaultButton1, "SPC - INFO DEL SISTEMA"
            txtIdMov.SetFocus
    End Select
    
End Sub

Private Sub LimpiarPantalla()
    FG1.Clear
    FG1.Rows = 2
End Sub

Private Sub btnLimpiar_Click()

    Call LimpiarPantalla
    Call SeteoGrilla
    Call EnabledAll
    txtIdMov.SetFocus
    
End Sub

Private Sub btnPrimero_Click()

'    On Error GoTo CapturaErrores
        FG1.Clear
        
        Set tMovIntStockC = BaseSPC.OpenRecordset("MovIntStockC", dbOpenTable)
        
        If Not tMovIntStockC.EOF Then
            tMovIntStockC.MoveFirst
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If

 'tMovIntStockC.Close
        
CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select

End Sub

Private Sub btnSalir_Click()

    Unload Me
    
End Sub

Private Sub btnUltimo_Click()

    On Error GoTo CapturaErrores
        FG1.Clear
        
        Set tMovIntStockC = BaseSPC.OpenRecordset("MovIntStockC", dbOpenTable)
        
        If Not tMovIntStockC.BOF Then
            tMovIntStockC.MoveLast
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If

'        tMovIntStockC.Close
        
CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select

End Sub

Private Sub cmbDestino_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub


Private Sub cmbDestino_LostFocus()
    
    tDepositos.Index = "IndDesc"
    tDepositos.MoveFirst
    tDepositos.Seek "=", cmbDestino.Text
    
    If Not tDepositos.NoMatch Then
        codDepositoDest = tDepositos!IDDEPOSITO
        codVendedorDest = tDepositos!VendedorAsociado
     Else
        A = MsgBox("Depósito Inexistente", vbCritical, "ERROR")
    End If
    
    FG1.Col = 0
    FG1.Row = 1
    FG1.SetFocus
    
    'MsgBox (codDepositoDest)
    'MsgBox (codVendedorDest)

End Sub




Private Sub btnSalir_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub cmbOrigen_GotFocus()
    cmbOrigen.SelLength = Len(cmbOrigen.Text)
End Sub

Private Sub cmbOrigen_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If


End Sub


Private Sub cmbOrigen_LostFocus()

    tDepositos.Index = "IndDesc"
    tDepositos.MoveFirst
    tDepositos.Seek "=", cmbOrigen.Text
    
    If Not tDepositos.NoMatch Then
        codDeposito = tDepositos!IDDEPOSITO
        codVendedor = tDepositos!VendedorAsociado
     Else
        A = MsgBox("Depósito Inexistente", vbCritical, "ERROR")
    End If
    
   ' MsgBox (codDeposito)
   ' MsgBox (codVendedor)

End Sub


Private Sub cmdVerStock_Click()

    Set tDepositos = BaseSPC.OpenRecordset("Depositos", dbOpenTable)
    
    tDepositos.Index = "IndDesc"
    
    tDepositos.Seek "=", cmbOrigen.Text
    
    If Not tDepositos.NoMatch Then IDDEPOSITO = tDepositos!IDDEPOSITO
    
    vSQL = ("SELECT * FROM Stock WHERE IDDeposito='" & IDDEPOSITO & "' ORDER BY CodProd")
    
    Set tStock = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)

    tStock.MoveFirst
    
    FG1.Rows = 2
    FG1.Row = 1
    
    While Not tStock.EOF
        FG1.Col = 0
        FG1.Text = tStock!CodProd
        FG1.Col = 1
        FG1.Text = BuscarDescProd(tStock!CodProd)
        FG1.Col = 2
        FG1.Text = tStock!cantidad
        FG1.Col = 3
        FG1.Text = tStock!FechaUM
        
        tStock.MoveNext
        FG1.Rows = FG1.Rows + 1
        FG1.Row = FG1.Row + 1
    Wend

End Sub

Private Sub cmdVerStock_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Imprime_Click()
    Dim objPrinterFlex As PrinterFlex
    Set objPrinterFlex = New PrinterFlex
    
    With objPrinterFlex
      
      'Asignamos los valores de los encabezados, el pie de página, el color_
      'del texto y el tamaño de la fuente
        
        'texto de los encabezdos y el pir de pagina
        .TextEncabezado1 = "MOVIMIENTOS DE STOCK"
            
            'Elegimos el tipo de revisión
                    Deposito = cmbOrigen.Text
                    'Pie = "Total de Revisiones AB: " & txtRev_AB.Text
                    Pie = "Desarrollado por SPC Software Integral"
        
        .TextEncabezado2 = Deposito
        '& Chr(10) & "Desde el " & txtDesde.Text & " al " & txtHasta.Text
                
        'CGrid.Row = 1
        'CGrid.Col = 10
        'Anio = CGrid.Text
        'CGrid.Col = 11
        'Periodo = CGrid.Text
        
        .TextPiePagina = Pie
               
        'Colores de la fuentes
        '.ColorPiePagina = txtPiePagina.ForeColor
        '.ColorEncabezado1 = txtEncabezado1.ForeColor
        '.ColorEncabezado2 = txtEncabezado2.ForeColor
        
        'Tamaños de las fuentes
        .SizeEncabezado1 = 12
        .SizeEncabezado2 = 10
        .SizePiePagina = 8
        .AjustarColumnas = False
      
        .Orientacion = Vertical
        'Imprimimos pasando el nombre del FlexGrid a imprimir
        .ImprimirFlexGrid FG1
    End With
    
    'Call objPrinterFlex.ImprimirFlexGrid(CGrid)

End Sub



Private Sub FG1_GotFocus()
    
   linea = 1
   FG1.Col = 0
   FG1.Row = 1
   FG1.Text = linea
   FG1.Col = 1

End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    Dim cantidad As Integer
    Dim Posicion As Integer
        
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    Set tStock = db.OpenRecordset("Stock", dbOpenTable)
    
    tStock.Index = "PrimaryKey"
            
    If KeyAscii >= 32 And KeyAscii <= 127 Then
        FG1.Text = FG1.Text & Chr(KeyAscii)
    End If

    Select Case KeyAscii
       Case 9, 13, 39
            FG1.Col = 1
            codigoprodMA = UCase(FG1.Text)
            FG1.Text = codigoprodMA
                   
            Dim busca1 As String, busca2 As String
            busca1 = RTrim(LTrim(codigoprodMA))
            busca2 = busca1 + "z"
                                     
            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
            codigoProdTabla = rstProductos.Fields!CodProd
            
            'If codigoProdTabla <> RTrim(LTrim(CodigoProdMA)) Then
            
             If rstProductos.NoMatch Then
                 mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                 codigoprod = ""
                 FG1.Col = 2
                 FG1.Text = ""
                 FG1.Col = 3
             Else
                 Call muestrodatosproductos
                 FG1.Col = FG1.Col + 1
             End If
               
            '**** Verifico Stock y Actualizo Cantidad
            If FG1.Col = 3 And FG1.Text <> "" Then
                 If KeyAscii = 13 Or KeyAscii = 9 Or KeyAscii = 39 Then
                     tStock.Seek "=", codigoprodMA, codDeposito
                     If tStock.NoMatch Then
                          MsgBox "Stock Inexistente en Almacén Seleccionado", vbCritical, "ERROR"
                          FG1.Col = 3
                          FG1.CellBackColor = QBColor(4)
                          FG1.CellFontBold = True
                          FG1.CellForeColor = QBColor(7)
                        Else
                           cantidad = FG1.Text
                           If tStock!cantidad < cantidad Then
                              MsgBox "Stock Insuficiente", vbCritical, "ERROR"
                              FG1.Col = 3
                              FG1.CellBackColor = QBColor(4)
                              FG1.CellFontBold = True
                              FG1.CellForeColor = QBColor(7)
                           End If
                     End If
                 End If
                 FG1.Rows = FG1.Rows + 1
                 Posicion = FG1.Rows
                 FG1.Row = Posicion - 1
                 FG1.Col = 0
                 linea = linea + 1
                 FG1.Text = linea
                 FG1.Col = 1
             End If
       
       Case vbKeyBack
            
            If Len(FG1) >= 1 Then
               FG1 = Left$(FG1, Len(FG1) - 1)
            Else
                KeyAscii = 0
            End If
           
     End Select
       
        
    codigoprod = ""
    'tStock.Close
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub muestrodatosproductos()

    FG1.Col = 2
    FG1.Text = rstProductos.Fields!Descripcion
           
End Sub

Private Sub Form_Load()

 'Seteo la grilla
    Call SeteoGrilla
    
  'Abro Base de Datos
    'Seteo la captura de errores de no hay registros en el archivo
        On Error GoTo CapturaErrores
        
        Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        
        'Tabla Depositos
            Set tDepositos = BaseSPC.OpenRecordset("Depositos", dbOpenTable)
            
            tDepositos.Index = "PrimaryKey"
            tDepositos.MoveFirst
        
           'Lleno combo Origen y Destino
                While Not tDepositos.EOF
                    cmbOrigen.AddItem tDepositos!Descripcion
                    tDepositos.MoveNext
                Wend
        
CapturaErrores:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub



Private Sub txtFecha_GotFocus()

    Dim Largo As Integer
    Largo = Len(txtFecha.Text)
    txtFecha.SelLength = Largo

End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub

Private Sub txtFecha_LostFocus()

        If Not IsDate(txtFecha.Text) Then
            MsgBox "Formato de Fecha Incorrecto", vbCritical, "ERROR !"
            txtFecha.Text = Format(Date, "DD/MM/YYYY")
        End If

End Sub

Private Sub txtIdMov_GotFocus()
    
    Dim Largo As Integer
    
    Largo = Len(txtIdMov.Text)
    txtIdMov.SelLength = Largo

End Sub


Private Sub txtIdMov_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub Imprime_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub
