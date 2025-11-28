VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FormListadoVentas 
   Caption         =   "INFORME DE VENTAS"
   ClientHeight    =   7215
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15015
   LinkTopic       =   "Form1"
   ScaleHeight     =   7215
   ScaleWidth      =   15015
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cmbProductos 
      Height          =   315
      Index           =   1
      Left            =   7800
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   720
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   14895
      Begin MSComCtl2.MonthView Mv1 
         Height          =   2370
         Left            =   2880
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   117506049
         CurrentDate     =   41921
      End
      Begin VB.Frame Frame4 
         Caption         =   "Opción de Liquidación"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         TabIndex        =   24
         Top             =   1320
         Width           =   12495
         Begin VB.OptionButton OptionAll 
            Caption         =   "Todos"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   9360
            TabIndex        =   6
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OptionL2 
            Caption         =   "Linea 2"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   6120
            TabIndex        =   5
            Top             =   240
            Width           =   975
         End
         Begin VB.OptionButton OptionL1 
            Caption         =   "Linea 1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2880
            TabIndex        =   4
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbProductos 
         Height          =   315
         Index           =   0
         Left            =   7800
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   3975
      End
      Begin VB.TextBox txtFechaDesde 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtFechaHasta 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2880
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cmbVendedores 
         Height          =   315
         Index           =   0
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "Acciones"
         Height          =   1095
         Left            =   840
         TabIndex        =   18
         Top             =   5640
         Width           =   5655
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
            Height          =   495
            Left            =   480
            TabIndex        =   9
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdSalir 
            Caption         =   "&Salir"
            Height          =   495
            Left            =   4080
            TabIndex        =   11
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir"
            Height          =   495
            Left            =   2280
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Liquidación"
         Height          =   1095
         Left            =   10320
         TabIndex        =   15
         Top             =   5640
         Width           =   3855
         Begin VB.TextBox txtImporteTotal 
            Height          =   375
            Left            =   1560
            TabIndex        =   16
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Importe Total:"
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
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Width           =   1200
         End
      End
      Begin VB.CommandButton cmdLiquidar 
         Caption         =   "&Liquidar"
         Height          =   495
         Left            =   12240
         TabIndex        =   7
         Top             =   600
         Width           =   1935
      End
      Begin VB.ComboBox cmbVendedores 
         Height          =   315
         Index           =   1
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.Timer Timer1 
         Left            =   8040
         Top             =   6120
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3135
         Left            =   600
         TabIndex        =   8
         Top             =   2160
         Width           =   13935
         _ExtentX        =   24580
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Producto:"
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
         Left            =   7680
         TabIndex        =   23
         Top             =   480
         Width           =   1290
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
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
         TabIndex        =   22
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
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
         Left            =   2640
         TabIndex        =   21
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Vendedor:"
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
         Left            =   4680
         TabIndex        =   20
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblImprimiendo 
         AutoSize        =   -1  'True
         Caption         =   "Imprimiendo..."
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
         Left            =   7800
         TabIndex        =   19
         Top             =   5760
         Visible         =   0   'False
         Width           =   1200
      End
   End
End
Attribute VB_Name = "FormListadoVentas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub LlenarGrilla(CodProd, cantidad, total)
    
       'On Error GoTo CapturaErrores
        
        FG1.Row = FG1.Row + 1
        FG1.Rows = FG1.Rows + 1
        
        FG1.Col = 0
        FG1.CellAlignment = 7
        FG1.Text = CodProd
        
        FG1.Col = 1
        FG1.CellAlignment = 1
        FG1.Text = BuscarDescProd(CodProd)
        
        FG1.Col = 2
        FG1.CellAlignment = 7
        FG1.Text = Format(cantidad, "Standard")
        
        FG1.Col = 3
        FG1.CellAlignment = 7
        FG1.Text = FormatCurrency(total, 2)
        
CapturaErrores:
    Select Case Err
        Case 3021
            Resume Next
        
    End Select

End Sub

Private Sub SeteoGrilla()

    FG1.Rows = 2
    FG1.Cols = 4
    FG1.Row = 0

    FG1.Col = 0
    FG1.ColWidth(0) = 1000
    FG1.CellAlignment = 4
    FG1.Text = "Producto"
    
    FG1.Col = 1
    FG1.CellAlignment = 4
    FG1.ColWidth(1) = 4000
    FG1.Text = "Descripcion"
    
    FG1.Col = 2
    FG1.CellAlignment = 4
    FG1.ColWidth(2) = 1500
    FG1.Text = "Cantidad"
    
    FG1.Col = 3
    FG1.CellAlignment = 4
    FG1.ColWidth(3) = 1500
    FG1.Text = "Importe $"

    
End Sub

Private Sub cmbProductos_Click(Index As Integer)

    cmbProductos(1).ListIndex = cmbProductos(0).ListIndex

End Sub


Private Sub cmbProductos_KeyPress(Index As Integer, KeyAscii As Integer)

    cmbProductos(1).ListIndex = cmbProductos(0).ListIndex
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub cmbProductos_LostFocus(Index As Integer)

    cmbProductos(1).ListIndex = cmbProductos(0).ListIndex

End Sub

Private Sub cmbVendedores_Click(Index As Integer)

    cmbVendedores(0).ListIndex = cmbVendedores(1).ListIndex

End Sub

Private Sub cmbVendedores_KeyPress(Index As Integer, KeyAscii As Integer)

    cmbVendedores(0).ListIndex = cmbVendedores(1).ListIndex
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub

Private Sub cmbVendedores_LostFocus(Index As Integer)

    cmbVendedores(0).ListIndex = cmbVendedores(1).ListIndex

End Sub

Private Sub cmdExcel_Click()

    Dim NombreArchivo As String
    
    NombreArchivo = "\Liq_Ventas_" + cmbVendedores(0).Text + "_" + Format(TxtFechaHasta.Text, "MMM-YYYY") + ".xlsx"
    
    'If Exportar_Excel(App.Path & "\Comisiones.xls", MSHFlexGrid1) Then
    
    Call Exportar_Excel(App.Path & NombreArchivo, FG1)
    'If Exportar_Excel(App.Path & NombreArchivo, FG1) Then
     '   MsgBox " Datos exportados en " & App.Path & NombreArchivo, vbInformation
    'End If


End Sub

Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
  
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
      
    ' -- Bucle para Exportar los datos
    With FlexGrid
        For Fila = 1 To .Rows - 1
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            Next
        Next
    End With
    
    o_Hoja.Cells(Fila, (Columna - 1)).Value = "Liquidación Total:"
    o_Hoja.Cells(Fila, Columna).Value = FormatCurrency(txtImporteTotal.Text, 2)
    
    
    o_Libro.Close True, sOutputPath
    Set o_Libro = o_Excel.Workbooks.Open(sOutputPath)
    o_Excel.Visible = True
    
    ' -- Cerrar Excel
    'o_Excel.Quit
    
    ' -- Terminar instancias
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    Exportar_Excel = True
Exit Function
  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
End Function

' -------------------------------------------------------------------
' \\ -- Eliminar objetos para liberar recursos
' -------------------------------------------------------------------
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub

Private Sub cmdImprimir_Click()

    Dim objPrinterFlex As PrinterFlex
    Set objPrinterFlex = New PrinterFlex
    
    With objPrinterFlex
      
      'Asignamos los valores de los encabezados, el pie de página, el color_
      'del texto y el tamaño de la fuente
        
        'texto de los encabezdos y el pie de pagina
        
        
        .TextEncabezado1 = Chr(9) & "LIQUIDACION DE VENTAS POR PRODUCTO"
            
                    nVendedor = Chr(9) & cmbVendedores(1).Text
                    Pie = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Liquidación Total: " & FormatCurrency(txtImporteTotal.Text, 2)
                    'Pie = "Desarrollado por SPC Software Integral"
        
        If OptionL1.Value = True Then li = "Ventas de Línea 1"
        If OptionL2.Value = True Then li = "Ventas de Línea 2"
        If OptionAll.Value = True Then li = "Todas las Ventas"
        
        .TextEncabezado2 = Chr(9) & nVendedor & Chr(10) & Chr(9) & Chr(9) & "Desde el " & TxtFechaDesde.Text & " al " & TxtFechaHasta.Text & Chr(10) & Chr(9) & Chr(9) & li
                
        'CGrid.Row = 1
        'CGrid.Col = 10
        'Anio = CGrid.Text
        'CGrid.Col = 11
        'Periodo = CGrid.Text
        
        .TextPiePagina = Pie
               
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
      
        .Orientacion = Vertical
        'Imprimimos pasando el nombre del FlexGrid a imprimir
        .ImprimirFlexGrid FG1
    End With
    
    'Call objPrinterFlex.ImprimirFlexGrid(CGrid)


End Sub

Private Sub cmdLiquidar_Click()

    Dim cantidad As Double
    Dim total As Double
    Dim L3 As Boolean
    Dim GrandTotal As Double
    
    cantidad = 0
    total = 0
    L3 = False
    GrandTotal = 0
        
    Call SeteoGrilla
    
    'On Error GoTo CapturaErrores
    
    FechaDesde = Format(TxtFechaDesde.Text, "m/d/yyyy")
    FechaHasta = Format(TxtFechaHasta.Text, "m/d/yyyy")
    
  'Se calcula la liquidación si se elige linea 1
    If OptionL1.Value = True Then
        vSQL1 = "SELECT * FROM qListadoVentasFac WHERE FechaFactura>=#" & FechaDesde & "# AND FechaFactura <=#" & FechaHasta & "# AND CodVendedor='" & cmbVendedores(0).Text & "' AND IDCodProd Like'" & cmbProductos(1).Text & "' ORDER BY IDCodProd, FechaFactura"
        
        'MsgBox (vSQL1)
        
        Set qVentasL1 = BaseSPC.OpenRecordset(vSQL1, dbOpenDynaset)
        
        qVentasL1.MoveFirst
       
        IdCodProd = qVentasL1!IdCodProd
        While qVentasL1.EOF = False
            If qVentasL1!IdCodProd = IdCodProd Then
                Cant = Cant + qVentasL1!cantidad
                total = total + qVentasL1!totalLinea
                qVentasL1.MoveNext
             Else
                Call LlenarGrilla(IdCodProd, Cant, total)
                IdCodProd = qVentasL1!IdCodProd
                GrandTotal = GrandTotal + total
                Cant = 0
                total = 0
            End If
        Wend
        Call LlenarGrilla(IdCodProd, Cant, total)
        GrandTotal = GrandTotal + total
    End If
    
    
  'Se calcula la liquidación si se elige linea 2
    If OptionL2.Value = True Then
        vSQL2 = "SELECT * FROM qListadoVentasPres WHERE FechaPresu>=#" & FechaDesde & "# AND FechaPresu <=#" & FechaHasta & "# AND CodVendedor='" & cmbVendedores(0).Text & "' AND CodProd Like'" & cmbProductos(1).Text & "' ORDER BY CodProd, FechaPresu"
        
        Set qVentasL2 = BaseSPC.OpenRecordset(vSQL2, dbOpenDynaset)
        
        qVentasL2.MoveFirst
        
        IdCodProd = qVentasL2!CodProd
        
        While Not qVentasL2.EOF
            If qVentasL2!CodProd = IdCodProd Then
                Cant = Cant + qVentasL2!cantidad
                total = total + qVentasL2!totalLinea
                qVentasL2.MoveNext
             Else
                Call LlenarGrilla(IdCodProd, Cant, total)
                GrandTotal = GrandTotal + total
                IdCodProd = qVentasL2!CodProd
                Cant = 0
                total = 0
            End If
        Wend
        Call LlenarGrilla(IdCodProd, Cant, total)
        GrandTotal = GrandTotal + total
    End If

  'Se calcula la liquidación si se eligen ambas lineas
    If OptionAll.Value = True Then
        vSQL1 = "SELECT * FROM qListadoVentasFac WHERE FechaFactura>=#" & FechaDesde & "# AND FechaFactura <=#" & FechaHasta & "# AND CodVendedor Like'" & cmbVendedores(0).Text & "' AND IDCodProd Like'" & cmbProductos(1).Text & "' ORDER BY IDCodProd, FechaFactura"
        vSQL2 = "SELECT * FROM qListadoVentasPres WHERE FechaPresu>=#" & FechaDesde & "# AND FechaPresu <=#" & FechaHasta & "# AND CodVendedor Like'" & cmbVendedores(0).Text & "' AND CodProd Like'" & cmbProductos(1).Text & "' ORDER BY CodProd, FechaPresu"
        
        'MsgBox (vSQL1)
        'MsgBox (vSQL2)
        
        Set qVentasL1 = BaseSPC.OpenRecordset(vSQL1, dbOpenDynaset)
        Set qVentasL2 = BaseSPC.OpenRecordset(vSQL2, dbOpenDynaset)
        
       'Tabla auxiliar de ventas
        Set tVA = BaseSPC.OpenRecordset("tAuxiliarVentas", dbOpenTable)
               
            'Vacío Tabla
             If Not tVA.EOF Then
                tVA.MoveFirst
                While Not tVA.EOF
                    tVA.Delete
                    tVA.MoveNext
                Wend
             End If
        tVA.Index = "PrimaryKey"
                
      'Preparo las consultas de L1 y L2
        qVentasL1.MoveFirst
        qVentasL2.MoveFirst
        
        IdCodProd = qVentasL1!IdCodProd
        IdCodProd2 = qVentasL2!CodProd
        
      'Cargo las ventas L1 en la Tabla Auxiliar
        While Not qVentasL1.EOF
            If qVentasL1!IdCodProd = IdCodProd Then
                Cant = Cant + qVentasL1!cantidad
                total = total + qVentasL1!totalLinea
                qVentasL1.MoveNext
            Else
                tVA.AddNew
                    tVA!IdProducto = IdCodProd
                    tVA!Descripcion = BuscarDescProd(IdCodProd)
                    tVA!cantidad = Cant
                    tVA!Importe = total
                tVA.Update
                GrandTotal = GrandTotal + total
                IdCodProd = qVentasL1!IdCodProd
                Cant = 0
                total = 0
            End If
        Wend
        
        tVA.AddNew
            tVA!IdProducto = IdCodProd
            tVA!Descripcion = BuscarDescProd(IdCodProd)
            tVA!cantidad = Cant
            tVA!Importe = total
        tVA.Update
        'GrandTotal = GrandTotal + total
            
        While Not qVentasL2.EOF
            If qVentasL2!CodProd = IdCodProd2 Then
                Cant = Cant + qVentasL2!cantidad
                total = total + qVentasL2!totalLinea
                qVentasL2.MoveNext
            Else
                tVA.Seek "=", IdCodProd2
                If Not tVA.NoMatch Then
                    tVA.Edit
                        tVA!cantidad = tVA!cantidad + Cant
                        tVA!Importe = tVA!Importe + total
                    tVA.Update
                    GrandTotal = GrandTotal + total
                 Else
                    tVA.AddNew
                        tVA!IdProducto = IdCodProd2
                        tVA!Descripcion = BuscarDescProd(IdCodProd2)
                        tVA!cantidad = Cant
                        tVA!Importe = total
                    tVA.Update
                    GrandTotal = GrandTotal + total
                End If
                IdCodProd2 = qVentasL2!CodProd
                Cant = 0
                total = 0
            End If
        Wend
        
        tVA.Seek "=", IdCodProd2
        If Not tVA.NoMatch Then
            tVA.Edit
                tVA!cantidad = tVA!cantidad + Cant
                tVA!Importe = tVA!Importe + total
            tVA.Update
            GrandTotal = GrandTotal + total
         Else
            tVA.AddNew
                tVA!IdProducto = IdCodProd2
                tVA!Descripcion = BuscarDescProd(IdCodProd2)
                tVA!cantidad = Cant
                tVA!Importe = total
            tVA.Update
            GrandTotal = GrandTotal + total
        End If
        
        tVA.MoveFirst
        
        While Not tVA.EOF
            'L3 = True
            Call LlenarGrilla(tVA!IdProducto, tVA!cantidad, tVA!Importe)
            'GrandTotal = GrandTotal + tVA!Importe
            tVA.MoveNext
        Wend
        
        tVA.Close
        
    End If

    txtImporteTotal.Text = Format(GrandTotal, "Standard")
    
CapturaErrores:
    Select Case Err
        Case 3021
        '   MsgBox "No hay Pagos Para Liquidar con el Criterio Seleccionado...", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
          Resume Next
    End Select

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
        Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        
        vSQL = "SELECT * FROM Empleados Where IdPuesto=1 ORDER BY Nombre"
        vSQL2 = "SELECT * FROM Productos ORDER BY CodProd"
        
        Set tVendedores = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
        Set tProductos = BaseSPC.OpenRecordset(vSQL2, dbOpenDynaset)
        
        
        On Error GoTo CapturaErrores
        'MsgBox (vSQL)

    'Llenar Combo Vendedores
        tVendedores.MoveFirst
        
        cmbVendedores(1).Clear
        While Not tVendedores.EOF
            cmbVendedores(0).AddItem (tVendedores!Legajo)
            cmbVendedores(1).AddItem (tVendedores!nombre)
            tVendedores.MoveNext
        Wend
        
        cmbVendedores(0).AddItem ("*")
        cmbVendedores(1).AddItem ("Todos")
        
        tVendedores.Close
    
    'Llenar Combo Productos
        tProductos.MoveFirst
        cmbProductos(0).Clear
        While Not tProductos.EOF
            cmbProductos(1).AddItem (tProductos!CodProd)
            cmbProductos(0).AddItem (tProductos!Descripcion)
            tProductos.MoveNext
        Wend
        
        cmbProductos(1).AddItem ("*")
        cmbProductos(0).AddItem ("Todos")
        
        tProductos.Close
    
    'Fechas
        TxtFechaDesde.Text = Format("01/01/2014", "DD/MM/YYYY")
        TxtFechaHasta.Text = Format(Date, "DD/MM/YYYY")

CapturaErrores:

Select Case Err
    
    Case 3021

End Select

End Sub

Private Sub Mv1_DateDblClick(ByVal DateDblClicked As Date)

 If Llamado = "Desde" Then
    TxtFechaDesde.Text = Mv1.Value
    TxtFechaDesde.SetFocus
    Mv1.Visible = False
 End If
 
 If Llamado = "Hasta" Then
    TxtFechaHasta.Text = Mv1.Value
    TxtFechaHasta.SetFocus
    Mv1.Visible = False
 End If


End Sub

Private Sub txtFechaDesde_DblClick()
    
    Llamado = "Desde"
    Mv1.Visible = True
    Mv1.Top = 1080
    Mv1.Left = 840
    Mv1.SetFocus

End Sub

Private Sub txtFechaDesde_GotFocus()

    TxtFechaDesde.SelLength = Len(TxtFechaDesde.Text)

End Sub

Private Sub txtFechaDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtFechaHasta_DblClick()

    Llamado = "Hasta"
    Mv1.Visible = True
    Mv1.Top = 1080
    Mv1.Left = 2880
    Mv1.SetFocus

End Sub

Private Sub txtFechaHasta_GotFocus()

    TxtFechaHasta.SelLength = Len(TxtFechaHasta.Text)

End Sub

Private Sub txtFechaHasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


