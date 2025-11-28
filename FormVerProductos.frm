VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormVerProductos 
   Caption         =   "Listado Productos"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8595
   ScaleWidth      =   12405
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   11655
      Begin VB.TextBox txt12Cuotas 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   8160
         TabIndex        =   18
         Text            =   "15"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txt6Cuotas 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   6960
         TabIndex        =   12
         Text            =   "15"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txt3Cuotas 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   5760
         TabIndex        =   11
         Text            =   "8"
         Top             =   480
         Width           =   735
      End
      Begin VB.TextBox txt2Cuotas 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         TabIndex        =   10
         Text            =   "0"
         Top             =   480
         Width           =   735
      End
      Begin VB.CommandButton cmdInteres 
         Caption         =   "Intereses"
         Height          =   495
         Left            =   9480
         TabIndex        =   14
         Top             =   240
         Width           =   1815
      End
      Begin VB.ComboBox cmbLista 
         Height          =   315
         Left            =   2280
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "12 Cuotas"
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
         Left            =   8040
         TabIndex        =   20
         Top             =   240
         Width           =   870
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "6 Cuotas"
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
         Left            =   6840
         TabIndex        =   16
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "3 Cuotas"
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
         Left            =   5640
         TabIndex        =   15
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Contado"
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
         Left            =   4440
         TabIndex        =   13
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Elija Lista de Precios:"
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
         Left            =   360
         TabIndex        =   8
         Top             =   360
         Width           =   1860
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   7200
      Width           =   11775
      Begin VB.CommandButton cmdPrintList 
         Caption         =   "&Imprimir Lista"
         Height          =   615
         Left            =   7920
         TabIndex        =   17
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonAgregar 
         Caption         =   "&Agregar"
         Height          =   615
         Left            =   4080
         TabIndex        =   6
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   10080
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonEliminar 
         Caption         =   "&Eliminar"
         Height          =   615
         Left            =   6000
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonGuardar 
         Caption         =   "&Guardar"
         Height          =   630
         Index           =   0
         Left            =   2280
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton BotonMostrarProductos 
         Caption         =   "&Mostrar Productos"
         Height          =   630
         Index           =   1
         Left            =   480
         TabIndex        =   2
         Top             =   240
         Width           =   1230
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   10398
      _Version        =   393216
      Rows            =   10
      Cols            =   5
      FixedCols       =   0
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
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
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   75
   End
End
Attribute VB_Name = "FormVerProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MuestroDatosCuotas(Cuotas)

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("select all * from Productos order by CodProd")
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    Call titulos
      
    If rstProductos.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Productos", vbCritical, "Final de la busqueda")
     
    End If
     
    Select Case Cuotas
    
        Case 2
            If Len(txt2Cuotas.Text) = 1 Then
                Interes = "1,0" & txt2Cuotas.Text
                Interes = CDbl(Interes)
             Else
                Interes = "1," & txt2Cuotas.Text
                Interes = CDbl(Interes)
            End If
        Case 3
            If Len(txt3Cuotas.Text) = 1 Then
                Interes = "1,0" & txt3Cuotas.Text
                Interes = CDbl(Interes)
             Else
                Interes = "1," & txt3Cuotas.Text
                Interes = CDbl(Interes)
            End If
        Case 6
            If Len(txt6Cuotas.Text) = 1 Then
                Interes = "1,0" & txt6Cuotas.Text
                Interes = CDbl(Interes)
             Else
                Interes = "1," & txt6Cuotas.Text
                Interes = CDbl(Interes)
            End If
            
        Case 12
            If Len(txt12Cuotas.Text) = 1 Then
                Interes = "1,0" & txt12Cuotas.Text
                Interes = CDbl(Interes)
             Else
                Interes = "1," & txt12Cuotas.Text
                Interes = CDbl(Interes)
            End If
            
    End Select
          
    linea2 = 1
    rstProductos.MoveFirst
    Do While Not rstProductos.EOF
        MSFlexGrid1.AddItem " "
        MSFlexGrid1.Row = linea2
       
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = rstProductos.Fields!CodProd
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = rstProductos.Fields!Descripcion
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = rstProductos.Fields!UnidadMedida
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = Format((rstProductos.Fields!PrecioUnitarioPresupuesto * Interes), "#00.00")
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = Format((rstProductos.Fields!PrecioUnitarioFactura * Interes), "#00.00")
            linea2 = linea2 + 1
      
       rstProductos.MoveNext
       
    Loop



End Sub

Private Sub titulos()

    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(0) = flexAlignCenteright
    MSFlexGrid1.Text = "Codigo"
    MSFlexGrid1.ColWidth(0) = 1200
    
        
    MSFlexGrid1.Col = 1
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(1) = flexAlignCenterright
    MSFlexGrid1.Text = "Descripcion"
    MSFlexGrid1.ColWidth(1) = 5000
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
    MSFlexGrid1.Text = "UM"
    MSFlexGrid1.ColWidth(2) = 800
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(3) = flexAlignCenterCenter
    MSFlexGrid1.Text = "Precio Presupues."
    MSFlexGrid1.ColWidth(3) = 1800
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.CellFontBold = True
    MSFlexGrid1.ColAlignment(4) = flexAlignCenterCenter
    MSFlexGrid1.Text = "Precio Factura"
    MSFlexGrid1.ColWidth(4) = 1800
 
 End Sub


Private Sub BotonAgregar_Click()

    FormProductos.Show

End Sub

Private Sub BotonEliminar_Click()

     On Error GoTo CapErr
     
     ruta = App.Path & "\DB_SPC_SI.mdb"
     Set db = DBEngine.OpenDatabase(ruta)
     Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
   
    tProductos.Index = "PrimaryKey"
    
    MSFlexGrid1.Col = 0
   ' MsgBox (MSFlexGrid1.Text)
            
    tProductos.Seek "=", MSFlexGrid1.Text
        
    M = MsgBox("¿Seguro desea eliminar el registro?", vbOKCancel, "INFO DEL SISTEMA")
        
    If M = 1 Then
        If Not tProductos.NoMatch Then
            tProductos.Delete
        End If
        Call muestrodatos
    End If

CapErr:
    Select Case Err
        Case 3021
          M = MsgBox("No se puede Eliminar, Producto con Movimientos", vbCritical, "INFO DEL SISTEMA")
          Exit Sub
    End Select

End Sub

Private Sub BotonGuardar_Click(Index As Integer)

    ruta = App.Path & "\DB_SPC_SI.mdb"
     Set db = DBEngine.OpenDatabase(ruta)
     Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
        
       MSFlexGrid1.Col = 0
       MSFlexGrid1.Row = 1
      Filas = MSFlexGrid1.Rows
      linea = 1
      Do While linea < Filas
          
           MSFlexGrid1.Row = linea
           MSFlexGrid1.Col = 0
          If MSFlexGrid1.Text <> " " Then
    
             MSFlexGrid1.Col = 0
             codigoprod = MSFlexGrid1.Text
        
             Dim busca1 As String, busca2 As String
             busca1 = RTrim(LTrim(codigoprod))
             busca2 = busca1 + "z"
        
             rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
             
             rstProductos.Edit
             MSFlexGrid1.Col = 0
             rstProductos.Fields!CodProd = MSFlexGrid1.Text
             
             MSFlexGrid1.Col = 1
             rstProductos.Fields!Descripcion = MSFlexGrid1.Text
                        
             MSFlexGrid1.Col = 2
             rstProductos.Fields!UnidadMedida = MSFlexGrid1.Text
                            
             MSFlexGrid1.Col = 3
             rstProductos.Fields!PrecioUnitarioPresupuesto = Format(MSFlexGrid1.Text, "#00.00")
                            
             MSFlexGrid1.Col = 4
             rstProductos.Fields!PrecioUnitarioFactura = Format(MSFlexGrid1.Text, "#00.00")
             
             rstProductos.Update
         
            
         End If
          linea = linea + 1
      Loop
      MSFlexGrid1.Clear
End Sub

Private Sub BotonMostrarProductos_Click(Index As Integer)

     Call muestrodatos
     
End Sub





Private Sub BotonSalir_Click()

    Unload Me

End Sub


Private Sub cmbLista_Change()

    Select Case cmbLista.ListIndex
    
        Case 0
            Call MuestroDatosCuotas(2)
        
        Case 1
            Call MuestroDatosCuotas(3)
            
        Case 2
            Call MuestroDatosCuotas(6)
        
        Case 3
            Call MuestroDatosCuotas(12)
    
    End Select

End Sub


Private Sub cmbLista_Click()

    Select Case cmbLista.ListIndex
        Case 0
            Call MuestroDatosCuotas(2)
        
        Case 1
            Call MuestroDatosCuotas(3)
            
        Case 2
            Call MuestroDatosCuotas(6)
        
        Case 3
            Call MuestroDatosCuotas(12)
        End Select

End Sub

Private Sub cmdInteres_Click()
    Dim tIntereses
    
    If ModifIntereses = 0 Then
        txt2Cuotas.Enabled = True
        txt3Cuotas.Enabled = True
        txt6Cuotas.Enabled = True
        txt12Cuotas.Enabled = True
        
        cmdInteres.Caption = "Guardar Cambios"
        ModifIntereses = 1
        txt2Cuotas.SetFocus
        Exit Sub
    End If

    If ModifIntereses = 1 Then
        ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set tIntereses = db.OpenRecordset("Intereses", dbOpenTable)
        
        tIntereses.MoveFirst
       
        While Not tIntereses.EOF
            Select Case tIntereses!Cuotas
                Case 2
                    tIntereses.Edit
                         tIntereses!Interes = txt2Cuotas.Text
                     tIntereses.Update
                Case 3
                    tIntereses.Edit
                         tIntereses!Interes = txt3Cuotas.Text
                     tIntereses.Update
                Case 6
                    tIntereses.Edit
                         tIntereses!Interes = txt6Cuotas.Text
                     tIntereses.Update
                Case 12
                    tIntereses.Edit
                         tIntereses!Interes = txt12Cuotas.Text
                     tIntereses.Update
            End Select
            tIntereses.MoveNext
        Wend
        
        cmdInteres.Caption = "Intereses"
        ModifIntereses = 0
        tIntereses.Close
        
        txt2Cuotas.Enabled = False
        txt3Cuotas.Enabled = False
        txt6Cuotas.Enabled = False
        txt12Cuotas.Enabled = False
        
        MsgBox ("Intereses Actualizados !!!")
        
    End If
    
End Sub

Private Sub cmdPrintList_Click()

    Dim Nombre As String
    Dim direccion As String
    
    Dim objPrinterFlex As PrinterFlex
    Set objPrinterFlex = New PrinterFlex
    
    Nombre = "    QUILPLAC S.A."
    direccion = "     Av. Andres Baranda Nº520 Quilmes"
    With objPrinterFlex
      
      'Asignamos los valores de los encabezados, el pie de página, el color_
      'del texto y el tamaño de la fuente
        
        'texto de los encabezdos y el pie de pagina
        .TextEncabezado1 = Chr(9) & "Lista de Precios"
            
                    'nombre = Chr(9) & direccion
'                    Pie = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Liquidación Total: " & FormatCurrency(txtImporteTotal.Text, 2)
                    
                     '& Chr(10) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Desarrollado por SPC Consulting"
                    'Pie = "Desarrollado por SPC Software Integral"
        
        If cmbLista.Text = "" Then Pie = "Lista de Precios Oficial"
        If cmbLista.Text = "Lista 2 Cuotas" Then Pie = "Lista de Precios con Intereses en 2 PAGOS con TARJETA"
        If cmbLista.Text = "Lista 3 Cuotas" Then Pie = "Lista de Precios con Intereses en 3 PAGOS con TARJETA"
        If cmbLista.Text = "Lista 6 Cuotas" Then Pie = "Lista de Precios con Intereses en 6 PAGOS con TARJETA"
        If cmbLista.Text = "Lista 12 Cuotas" Then Pie = "Lista de Precios con Intereses en 12 PAGOS con TARJETA"
        
        .TextEncabezado2 = Chr(9) & Nombre & Chr(10) & Chr(9) & direccion & Chr(9) & " ->  " & Format(Date, "DD/MM/YYYY")
                
        'CGrid.Row = 1
        'CGrid.Col = 10
        'Anio = CGrid.Text
        'CGrid.Col = 11
        'Periodo = CGrid.Text
        
        .TextPiePagina = Chr(9) & Pie
               
        'Colores de la fuentes
        .ColorPiePagina = QBColor(4)
        'txtPiePagina.ForeColor
        .ColorEncabezado1 = QBColor(1)
        'txtEncabezado1.ForeColor
        .ColorEncabezado2 = QBColor(0)
        'txtEncabezado2.ForeColor
        
        'Tamaños de las fuentes
        .SizeEncabezado1 = 12
        .SizeEncabezado2 = 10
        .SizePiePagina = 11
        .AjustarColumnas = True
      
        .Orientacion = Horizontal
        'Imprimimos pasando el nombre del FlexGrid a imprimir
        .ImprimirFlexGrid MSFlexGrid1
    End With
    
    'Call objPrinterFlex.ImprimirFlexGrid(CGrid)

End Sub


Private Sub Form_Load()

    FormVerProductos.Height = 8985
    FormVerProductos.Width = 12240
    FormVerProductos.Top = 1000
    FormVerProductos.Left = 1000
    
    Call muestrodatos
    
    ModifIntereses = 0
        
End Sub

Private Sub muestrodatos()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("select all * from Productos order by CodProd")
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    cmbLista.AddItem "Lista 2 Cuotas"
    cmbLista.AddItem "Lista 3 Cuotas"
    cmbLista.AddItem "Lista 6 Cuotas"
    cmbLista.AddItem "Lista 12 Cuotas"
    
    Set tIntereses = db.OpenRecordset("Intereses", dbOpenTable)
        tIntereses.Index = "PrimaryKey"
        
        tIntereses.Seek "=", 2
        If Not tIntereses.NoMatch Then txt2Cuotas.Text = tIntereses!Interes
        tIntereses.Seek "=", 3
        If Not tIntereses.NoMatch Then txt3Cuotas.Text = tIntereses!Interes
        tIntereses.Seek "=", 6
        If Not tIntereses.NoMatch Then txt6Cuotas.Text = tIntereses!Interes
        tIntereses.Seek "=", 12
        If Not tIntereses.NoMatch Then txt12Cuotas.Text = tIntereses!Interes
    tIntereses.Close
    
    Call titulos
      
    If rstProductos.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Productos", vbCritical, "Final de la busqueda")
     
    End If
     
    linea2 = 1
    rstProductos.MoveFirst
    Do While Not rstProductos.EOF
        MSFlexGrid1.AddItem " "
        MSFlexGrid1.Row = linea2
       
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = rstProductos.Fields!CodProd
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = rstProductos.Fields!Descripcion
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = rstProductos.Fields!UnidadMedida
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = Format(rstProductos.Fields!PrecioUnitarioPresupuesto, "#00.00")
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#00.00")
            linea2 = linea2 + 1
      
       rstProductos.MoveNext
       
    Loop
End Sub


Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 46 Then KeyAscii = 44
    
    If KeyAscii >= 32 And KeyAscii <= 127 Then
            MSFlexGrid1.Text = MSFlexGrid1.Text & Chr(KeyAscii)
       End If
       
    Select Case KeyAscii
       
       Case 13
      
            MSFlexGrid1.Col = 0
            codigoprodMA = UCase(MSFlexGrid1.Text)
       
       Case vbKeyBack
            
            If Len(MSFlexGrid1) >= 1 Then
               MSFlexGrid1 = Left$(MSFlexGrid1, Len(MSFlexGrid1) - 1)
            Else
                KeyAscii = 0
            End If
           
       End Select
       
        
       codigoprod = ""

End Sub

Private Sub txt12Cuotas_GotFocus()

    txt12Cuotas.SelLength = Len(txt12Cuotas.Text)

End Sub

Private Sub txt12Cuotas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub txt2Cuotas_GotFocus()

    txt2Cuotas.SelLength = (Len(txt2Cuotas.Text))
    
End Sub


Private Sub txt2Cuotas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub txt3Cuotas_GotFocus()

    txt3Cuotas.SelLength = Len(txt3Cuotas.Text)

End Sub

Private Sub txt3Cuotas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub txt6Cuotas_GotFocus()

    txt6Cuotas.SelLength = Len(txt6Cuotas.Text)

End Sub

Private Sub txt6Cuotas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


