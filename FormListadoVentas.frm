VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
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
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   7095
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   14895
      Begin VB.CommandButton cmdReporteVentas 
         Caption         =   "&Reporte Ventas"
         Height          =   615
         Left            =   8640
         TabIndex        =   28
         Top             =   6360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ComboBox cmbCliente 
         Height          =   315
         Left            =   10440
         TabIndex        =   4
         Text            =   "cmbCliente"
         Top             =   720
         Width           =   2535
      End
      Begin MSComCtl2.MonthView Mv1 
         Height          =   2370
         Left            =   4800
         TabIndex        =   26
         Top             =   2640
         Visible         =   0   'False
         Width           =   2700
         _ExtentX        =   4763
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   123928577
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
         Left            =   600
         TabIndex        =   25
         Top             =   1320
         Width           =   13935
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
            TabIndex        =   8
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
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   240
            Width           =   975
         End
      End
      Begin VB.ComboBox cmbProductos 
         Height          =   315
         Index           =   0
         Left            =   7560
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtFechaDesde 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   600
         TabIndex        =   0
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtFechaHasta 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.ComboBox cmbVendedores 
         Height          =   315
         Index           =   0
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Frame Frame2 
         Caption         =   "Acciones"
         Height          =   1095
         Left            =   840
         TabIndex        =   19
         Top             =   5760
         Width           =   5655
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
            Height          =   495
            Left            =   480
            TabIndex        =   10
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdSalir 
            Caption         =   "&Salir"
            Height          =   495
            Left            =   4080
            TabIndex        =   12
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir"
            Height          =   495
            Left            =   2280
            TabIndex        =   11
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Liquidación"
         Height          =   1095
         Left            =   10320
         TabIndex        =   16
         Top             =   5760
         Width           =   3855
         Begin VB.TextBox txtImporteTotal 
            Height          =   375
            Left            =   1560
            TabIndex        =   17
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
            TabIndex        =   18
            Top             =   360
            Width           =   1200
         End
      End
      Begin VB.CommandButton cmdLiquidar 
         Caption         =   "&Liquidar"
         Height          =   495
         Left            =   13320
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox cmbVendedores 
         Height          =   315
         Index           =   1
         Left            =   4560
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   2655
      End
      Begin VB.Timer Timer1 
         Left            =   8040
         Top             =   6240
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3135
         Left            =   600
         TabIndex        =   9
         Top             =   2160
         Width           =   14055
         _ExtentX        =   24791
         _ExtentY        =   5530
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin VB.Label lblTotal 
         AutoSize        =   -1  'True
         Caption         =   "Total..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C000&
         Height          =   195
         Left            =   840
         TabIndex        =   29
         Top             =   5400
         Visible         =   0   'False
         Width           =   630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   10320
         TabIndex        =   27
         Top             =   480
         Width           =   660
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
         Left            =   7440
         TabIndex        =   24
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
         Left            =   360
         TabIndex        =   23
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
         Left            =   2400
         TabIndex        =   22
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
         Left            =   4440
         TabIndex        =   21
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
         TabIndex        =   20
         Top             =   5880
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
        FG1.text = CodProd
        
        FG1.Col = 1
        FG1.CellAlignment = 1
        FG1.text = BuscarDescProd(CodProd)
        
        FG1.Col = 2
        FG1.CellAlignment = 7
        FG1.text = Format(cantidad, "Standard")
        
        FG1.Col = 3
        FG1.CellAlignment = 7
        FG1.text = FormatCurrency(total, 2)
        
CapturaErrores:
    Select Case Err
        Case 3021
            Resume Next
        
    End Select

End Sub

Private Sub SeteoGrilla()

    FG1.Clear

    FG1.Rows = 2
    FG1.Cols = 4
    FG1.Row = 0

    FG1.Col = 0
    FG1.ColWidth(0) = 1000
    FG1.CellAlignment = 4
    FG1.text = "Producto"
    
    FG1.Col = 1
    FG1.CellAlignment = 4
    FG1.ColWidth(1) = 4000
    FG1.text = "Descripcion"
    
    FG1.Col = 2
    FG1.CellAlignment = 4
    FG1.ColWidth(2) = 1500
    FG1.text = "Cantidad"
    
    FG1.Col = 3
    FG1.CellAlignment = 4
    FG1.ColWidth(3) = 1500
    FG1.text = "Importe $"
    
End Sub

Private Sub cmbCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub cmbCliente_LostFocus()

    lblTotal.Visible = True

End Sub

Private Sub cmbProductos_Click(Index As Integer)

    cmbProductos(1).ListIndex = cmbProductos(0).ListIndex

End Sub


Private Sub cmbProductos_KeyPress(Index As Integer, KeyAscii As Integer)

    cmbProductos(1).ListIndex = cmbProductos(0).ListIndex
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
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
            Sendkeys "{TAB}"
    End If

End Sub

Private Sub cmbVendedores_LostFocus(Index As Integer)

    cmbVendedores(0).ListIndex = cmbVendedores(1).ListIndex

End Sub

Private Sub cmdExcel_Click()

    Dim NombreArchivo As String
    
    Vendor = cmbVendedores(0).text
    
    If Vendor = "*" Then Vendor = "Todos"
    
    NombreArchivo = "\Liq_Ventas_" + Vendor + "_" + Format(txtFechaHasta.text, "MMM-YYYY") + ".xlsx"
    
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
    o_Hoja.Cells(Fila, Columna).Value = FormatCurrency(txtImporteTotal.text, 2)
    
    
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
            
                    nVendedor = Chr(9) & cmbVendedores(1).text
                    Pie = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Liquidación Total: " & FormatCurrency(txtImporteTotal.text, 2)
                    'Pie = "Desarrollado por SPC Software Integral"
        
        If OptionL1.Value = True Then li = "Ventas de Línea 1"
        If OptionL2.Value = True Then li = "Ventas de Línea 2"
        If OptionAll.Value = True Then li = "Todas las Ventas"
        
        .TextEncabezado2 = Chr(9) & nVendedor & Chr(10) & Chr(9) & Chr(9) & "Desde el " & txtFechaDesde.text & " al " & txtFechaHasta.text & Chr(10) & Chr(9) & Chr(9) & li
                
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

    Dim cantidad, totalProductos As Double
    Dim total As Double
    Dim L3 As Boolean
    Dim GrandTotal As Double
    
    cantidad = 0
    totalProductos = 0
    total = 0
    L3 = False
    GrandTotal = 0
        
    Call SeteoGrilla
    
    On Error GoTo CapturaErrores
    
    FechaDesde = Format(txtFechaDesde.text, "m/d/yyyy")
    FechaHasta = Format(txtFechaHasta.text, "m/d/yyyy")
    
   
  'Se calcula la liquidación si se elige linea 1
    If OptionL1.Value = True Then
        
        If cmbCliente.text = "Todos" Then
            vSQL1 = "SELECT * FROM qListadoVentasFac WHERE FechaFactura>=#" & FechaDesde & "# AND FechaFactura <=#" & FechaHasta & "# AND CodVendedor Like '" & cmbVendedores(0).text & "' AND IDCodProd Like'" & cmbProductos(1).text & "' ORDER BY IDCodProd, FechaFactura"
         Else
            vSQL1 = "SELECT * FROM qListadoVentasFac WHERE FechaFactura>=#" & FechaDesde & "# AND FechaFactura <=#" & FechaHasta & "# AND CodVendedor Like '" & cmbVendedores(0).text & "' AND IDCodProd Like'" & cmbProductos(1).text & "' AND CodCliente =" & Val(cmbCliente.text) & " ORDER BY IDCodProd, FechaFactura"
        End If
    
'         MsgBox (vSQL1)
        
        Set qventasL1 = BaseSPC.OpenRecordset(vSQL1, dbOpenDynaset)
        
        totalProductos = 0
        qventasL1.MoveFirst
        IdCodProd = qventasL1!IdCodProd
        
        While qventasL1.EOF = False
            If qventasL1!IdCodProd = IdCodProd Then
                Cant = Cant + qventasL1!cantidad
                total = total + qventasL1!totalLinea + (qventasL1!totalLinea * qventasL1!PorcentajeIVA / 100)
                qventasL1.MoveNext
             Else
                Call LlenarGrilla(IdCodProd, Cant, total)
                totalProductos = totalProductos + Cant
                IdCodProd = qventasL1!IdCodProd
                GrandTotal = GrandTotal + total
                Cant = 0
                total = 0
            End If
        Wend
        
        totalProductos = totalProductos + Cant
        Call LlenarGrilla(IdCodProd, Cant, total)
        GrandTotal = GrandTotal + total
        
        lblTotal.Visible = True
        lblTotal.Caption = "Total de Productos Vendidos por: " & cmbVendedores(1).text & " y del cliente -> " & cmbCliente.text & " en L1 = " & Format(totalProductos, "Standard")
        
    End If
    
    
  'Se calcula la liquidación si se elige linea 2
    If OptionL2.Value = True Then
        
        If cmbCliente.text = "Todos" Then
            vsql2 = "SELECT * FROM qListadoVentasPres WHERE FechaPresu>=#" & FechaDesde & "# AND FechaPresu <=#" & FechaHasta & "# AND CodVendedor Like '" & cmbVendedores(0).text & "' AND CodProd Like'" & cmbProductos(1).text & "' ORDER BY CodProd, FechaPresu"
         Else
            vsql2 = "SELECT * FROM qListadoVentasPres WHERE FechaPresu>=#" & FechaDesde & "# AND FechaPresu <=#" & FechaHasta & "# AND CodVendedor Like '" & cmbVendedores(0).text & "' AND CodProd Like'" & cmbProductos(1).text & "' AND CodCliente =" & Val(cmbCliente.text) & " ORDER BY CodProd, FechaPresu"
        End If
        
        'MsgBox (vsql2)
        
        Set qVentasL2 = BaseSPC.OpenRecordset(vsql2, dbOpenDynaset)
        
        qVentasL2.MoveFirst
        
        IdCodProd = qVentasL2!CodProd
        totalProductos = 0
        While Not qVentasL2.EOF
            If qVentasL2!CodProd = IdCodProd Then
                Cant = Cant + qVentasL2!cantidad
                totalProductos = totalProductos + Cant
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
        
        lblTotal.Visible = True
        lblTotal.Caption = "Total de Productos Vendidos por: " & cmbVendedores(1).text & " y del cliente -> " & cmbCliente.text & " en L2 = " & Format(totalProductos, "Standard")
        
    End If

  'Se calcula la liquidación si se eligen ambas lineas
    If OptionAll.Value = True Then
        If cmbCliente.text = "Todos" Then
            vSQL1 = "SELECT * FROM qListadoVentasFac WHERE FechaFactura>=#" & FechaDesde & "# AND FechaFactura <=#" & FechaHasta & "# AND CodVendedor Like'" & cmbVendedores(0).text & "' AND IDCodProd Like'" & cmbProductos(1).text & "' ORDER BY IDCodProd, FechaFactura"
            vsql2 = "SELECT * FROM qListadoVentasPres WHERE FechaPresu>=#" & FechaDesde & "# AND FechaPresu <=#" & FechaHasta & "# AND CodVendedor Like'" & cmbVendedores(0).text & "' AND CodProd Like'" & cmbProductos(1).text & "' ORDER BY CodProd, FechaPresu"
         Else
            vSQL1 = "SELECT * FROM qListadoVentasFac WHERE FechaFactura>=#" & FechaDesde & "# AND FechaFactura <=#" & FechaHasta & "# AND CodVendedor Like '" & cmbVendedores(0).text & "' AND IDCodProd Like'" & cmbProductos(1).text & "' AND CodCliente =" & Val(cmbCliente.text) & " ORDER BY IDCodProd, FechaFactura"
            vsql2 = "SELECT * FROM qListadoVentasPres WHERE FechaPresu>=#" & FechaDesde & "# AND FechaPresu <=#" & FechaHasta & "# AND CodVendedor Like '" & cmbVendedores(0).text & "' AND CodProd Like'" & cmbProductos(1).text & "' AND CodCliente =" & Val(cmbCliente.text) & " ORDER BY CodProd, FechaPresu"
        End If
        
        'MsgBox (vSQL1)
        'MsgBox (vSQL2)
        
        Set qventasL1 = BaseSPC.OpenRecordset(vSQL1, dbOpenDynaset)
        Set qVentasL2 = BaseSPC.OpenRecordset(vsql2, dbOpenDynaset)
        
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
        qventasL1.MoveFirst
        qVentasL2.MoveFirst
        
        IdCodProd = qventasL1!IdCodProd
        IdCodProd2 = qVentasL2!CodProd
        
      'Cargo las ventas L1 en la Tabla Auxiliar
        While Not qventasL1.EOF
            If qventasL1!IdCodProd = IdCodProd Then
                Cant = Cant + qventasL1!cantidad
                'total = total + qventasL1!totalLinea
                total = total + qventasL1!totalLinea + (qventasL1!totalLinea * qventasL1!PorcentajeIVA / 100)
                qventasL1.MoveNext
            Else
                tVA.AddNew
                    tVA!IdProducto = IdCodProd
                    tVA!Descripcion = BuscarDescProd(IdCodProd)
                    tVA!cantidad = Cant
                    tVA!Importe = total
                tVA.Update
                GrandTotal = GrandTotal + total
                IdCodProd = qventasL1!IdCodProd
                Cant = 0
                total = 0
            End If
        Wend
        
        If Not IdCodProd = "" Then
            tVA.AddNew
                tVA!IdProducto = IdCodProd
                tVA!Descripcion = BuscarDescProd(IdCodProd)
                tVA!cantidad = Cant
                tVA!Importe = total
            tVA.Update
        End If
        'esto estaba comentado
        GrandTotal = GrandTotal + total
            
        Cant = 0
        total = 0
        
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
                    If Not IdCodProd2 = "" Then
                        tVA.AddNew
                            tVA!IdProducto = IdCodProd2
                            tVA!Descripcion = BuscarDescProd(IdCodProd2)
                            tVA!cantidad = Cant
                            tVA!Importe = total
                        tVA.Update
                    End If
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
           If Not IdCodProd2 = "" Then
            tVA.AddNew
                tVA!IdProducto = IdCodProd2
                tVA!Descripcion = BuscarDescProd(IdCodProd2)
                tVA!cantidad = Cant
                tVA!Importe = total
            tVA.Update
           End If
            GrandTotal = GrandTotal + total
        End If
        
        totalProductos = 0
        tVA.MoveFirst
        
        While Not tVA.EOF
            'L3 = True
            Call LlenarGrilla(tVA!IdProducto, tVA!cantidad, tVA!Importe)
            'GrandTotal = GrandTotal + tVA!Importe
            totalProductos = totalProductos + tVA!cantidad
            tVA.MoveNext
        Wend
        
        tVA.Close
        
        lblTotal.Visible = True
        lblTotal.Caption = "Total de Productos Vendidos por: " & cmbVendedores(1).text & " y del cliente -> " & cmbCliente.text & " en L1 + L2 = " & Format(totalProductos, "Standard")
    
    End If

    txtImporteTotal.text = Format(GrandTotal, "Standard")
    
CapturaErrores:
    Select Case Err
        Case 3021
        '   MsgBox "No hay Pagos Para Liquidar con el Criterio Seleccionado...", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
          Resume Next
    End Select

End Sub

Private Sub cmdReporteVentas_Click()
    Dim db As DAO.Database
    Dim ws As DAO.Workspace
    Dim rs As DAO.Recordset
    Dim sql As String
    Dim sqlFacturas As String
    Dim sqlPresupuestos As String
    Dim sqlFinal As String
    Dim rutaBase As String
    Dim clave As String
    Dim TotalBaterias As Long
    
    ' Variables para filtros
    Dim fDesde As String, fHasta As String
    Dim filtroCliente As String
    Dim filtroVendedor As String
    Dim filtroArticulo As String
    
    rutaBase = "DB_SPC_SI.mdb"
    'clave = "theidol-1995"
    
'    On Error GoTo ErrorHandler
    
    Set ws = DBEngine.Workspaces(0)
    'Set db = ws.OpenDatabase(rutaBase, False, False, ";PWD=" & clave)
    Set db = ws.OpenDatabase(rutaBase)
    
    ' --- Preparar Grilla ---
    With FG1
        .Clear
        .Rows = 2
        .Cols = 3
        .FixedRows = 1
        .TextMatrix(0, 0) = "Tipo de Batería"
        .TextMatrix(0, 1) = "Cantidad"
        .TextMatrix(0, 2) = "Cliente"
        .ColWidth(0) = 3000
        .ColWidth(1) = 1500
        .ColWidth(2) = 4000
    End With
    
    ' --- 1. Preparar Filtros Comunes ---
    ' Formateamos las fechas para SQL
    fDesde = "#" & Format(txtFechaDesde.text, "mm/dd/yyyy") & "#"
    fHasta = "#" & Format(txtFechaHasta.text, "mm/dd/yyyy") & "#"
    
    ' Filtro Cliente
    filtroCliente = ""
    If cmbCliente.text <> "Todos" And cmbCliente.text <> "" Then
        ' Usamos Val() para extraer solo el ID del string "123 - Nombre"
        filtroCliente = " AND C.IDCliente = " & Val(cmbCliente.text)
    End If
    
    ' Filtro Vendedor
    filtroVendedor = ""
    If cmbVendedores(1).text <> "Todos" And cmbVendedores(1).text <> "" Then
        ' Nota: Asumo que en ambas tablas el campo se llama CodVendedor
        ' Se aplica sobre el alias de la tabla cabecera (FC o PC)
        filtroVendedor = " AND Cab.CodVendedor = '" & cmbVendedores(0).text & "'"
    End If
    
    ' Filtro Artículo
    filtroArticulo = ""
    If cmbProductos(0).text <> "Todos" And cmbProductos(0).text <> "" Then
        filtroArticulo = " AND P.Descripcion = '" & cmbProductos(0).text & "'"
    End If

    ' --- 2. Armar Sub-Consultas ---
    
    ' Consulta A: FACTURAS (L1)
    ' Nota: En FacturaD el campo es IDCodProd
    sqlFacturas = "SELECT P.Descripcion, FD.Cantidad, C.RazonSocial " & _
                  "FROM ((FacturaD FD INNER JOIN Productos P ON FD.IDCodProd = P.CodProd) " & _
                  "INNER JOIN FacturaC Cab ON FD.NroFactura = Cab.NroFactura) " & _
                  "INNER JOIN Clientes C ON Cab.CodCliente = C.IDCliente " & _
                  "WHERE Cab.FechaFactura >= " & fDesde & " AND Cab.FechaFactura <= " & fHasta & _
                  filtroCliente & filtroVendedor & filtroArticulo
    'MsgBox (sqlFacturas)
                  
    ' Consulta B: PRESUPUESTOS (L2)
    ' Nota: En PresupuestoD el campo es CodProd y la fecha tiene espacio [Fecha Presu]
    sqlPresupuestos = "SELECT P.Descripcion, PD.Cantidad, C.RazonSocial " & _
                      "FROM ((PresupuestoD PD INNER JOIN Productos P ON PD.CodProd = P.CodProd) " & _
                      "INNER JOIN PresupuestoC Cab ON PD.NroPresu = Cab.NroPresu) " & _
                      "INNER JOIN Clientes C ON Cab.CodCliente = C.IDCliente " & _
                      "WHERE Cab.[Fecha Presu] >= " & fDesde & " AND Cab.[Fecha Presu] <= " & fHasta & _
                      filtroCliente & filtroVendedor & filtroArticulo

    ' --- 3. Seleccionar la Consulta Final según OptionButton ---
    
    If OptionL1.Value = True Then
        ' Solo Facturas
        sqlFinal = "SELECT Descripcion, SUM(Cantidad) as Total, RazonSocial " & _
                   "FROM (" & sqlFacturas & ") " & _
                   "GROUP BY Descripcion, RazonSocial ORDER BY Descripcion"
                   
    ElseIf OptionL2.Value = True Then
        ' Solo Presupuestos
        sqlFinal = "SELECT Descripcion, SUM(Cantidad) as Total, RazonSocial " & _
                   "FROM (" & sqlPresupuestos & ") " & _
                   "GROUP BY Descripcion, RazonSocial ORDER BY Descripcion"
                   
    ElseIf OptionAll.Value = True Then
        ' Ambos (UNION ALL)
        sqlFinal = "SELECT Descripcion, SUM(Cantidad) as Total, RazonSocial " & _
                   "FROM (" & sqlFacturas & " UNION ALL " & sqlPresupuestos & ") " & _
                   "GROUP BY Descripcion, RazonSocial ORDER BY Descripcion"
    End If
    
    ' --- 4. Ejecutar y Llenar ---
    Set rs = db.OpenRecordset(sqlFinal, dbOpenSnapshot)
    
    TotalBaterias = 0
    If Not rs.EOF Then
        While Not rs.EOF
            FG1.AddItem rs!Descripcion & vbTab & _
                        rs!total & vbTab & _
                        rs!RazonSocial
                        
            TotalBaterias = TotalBaterias + Val(rs!total)
            rs.MoveNext
        Wend
    Else
        MsgBox "No se encontraron movimientos para los filtros seleccionados.", vbInformation
    End If
    
    lblTotal.Caption = "Total Baterías: " & Format(TotalBaterias, "#,##0")
    
Cierre:
    If Not rs Is Nothing Then rs.Close
    If Not db Is Nothing Then db.Close
    Set rs = Nothing
    Set db = Nothing
    Set ws = Nothing
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Cierre
End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()
    
        Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        
        vSQL = "SELECT * FROM Empleados Where IdPuesto=1 ORDER BY Nombre"
        vsql2 = "SELECT * FROM Productos ORDER BY CodProd"
        
        Set tVendedores = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
        Set tProductos = BaseSPC.OpenRecordset(vsql2, dbOpenDynaset)
        
        
        On Error GoTo CapturaErrores
        'MsgBox (vSQL)

    'Llenar Combo Vendedores
        tVendedores.MoveFirst
        
        cmbVendedores(1).Clear
        While Not tVendedores.EOF
            cmbVendedores(0).AddItem (tVendedores!Legajo)
            cmbVendedores(1).AddItem (tVendedores!Nombre)
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
    
    'Llenar Combo Vendedores
        CargarComboClientes
    
    'Fechas
        txtFechaDesde.text = Format("01/01/2014", "DD/MM/YYYY")
        txtFechaHasta.text = Format(Date, "DD/MM/YYYY")
    
    'Option por defecto
        OptionAll.Value = True

CapturaErrores:

Select Case Err
    
    Case 3021

End Select

End Sub

Private Sub Mv1_DateDblClick(ByVal DateDblClicked As Date)

 If Llamado = "Desde" Then
    txtFechaDesde.text = Mv1.Value
    txtFechaDesde.SetFocus
    Mv1.Visible = False
 End If
 
 If Llamado = "Hasta" Then
    txtFechaHasta.text = Mv1.Value
    txtFechaHasta.SetFocus
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

    txtFechaDesde.SelLength = Len(txtFechaDesde.text)

End Sub

Private Sub txtFechaDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
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

    txtFechaHasta.SelLength = Len(txtFechaHasta.text)

End Sub

Private Sub txtFechaHasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub CargarComboClientes()
   ' Dim rs As DAO.Recordset
    Dim sql As String
    
    ' Buscamos ID y RazonSocial, ordenados por ID para facilitar la búsqueda numérica
    sql = "SELECT IDCliente, RazonSocial FROM Clientes ORDER BY IDCliente"
    Set rs = BaseSPC.OpenRecordset(sql, dbOpenSnapshot)
    
    cmbCliente.Clear
    cmbCliente.AddItem "Todos"
    
    While Not rs.EOF
        If Not IsNull(rs!IDCliente) Then
            ' Formato: "1234 - Nombre del Cliente"
            ' Esto permite escribir el número para buscar
            cmbCliente.AddItem rs!IDCliente & " - " & rs!RazonSocial
        End If
        rs.MoveNext
    Wend
    
    cmbCliente.ListIndex = 0 ' Selecciona "Todos" por defecto
    
    rs.Close
    Set rs = Nothing
End Sub
