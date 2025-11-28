VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormListadoClientesPorVendedor 
   BackColor       =   &H80000005&
   Caption         =   "Listado de Saldos de Clientes - Por Vendedor"
   ClientHeight    =   7110
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15225
   LinkTopic       =   "Form1"
   ScaleHeight     =   7110
   ScaleWidth      =   15225
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6855
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15135
      Begin VB.Frame Frame4 
         Height          =   1095
         Left            =   1800
         TabIndex        =   10
         Top             =   5520
         Width           =   11775
         Begin VB.CommandButton cmdSaldoAFecha 
            Caption         =   "&Saldo a Fecha"
            Height          =   615
            Left            =   3600
            TabIndex        =   15
            Top             =   240
            Width           =   1695
         End
         Begin VB.CommandButton cmdSalir 
            Caption         =   "&Salir"
            Height          =   615
            Left            =   9720
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdExportar 
            Caption         =   "&Exportar a Excel"
            Height          =   615
            Left            =   6960
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir"
            Height          =   615
            Left            =   960
            TabIndex        =   2
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Listado"
         Height          =   3735
         Left            =   600
         TabIndex        =   9
         Top             =   1560
         Width           =   14055
         Begin MSFlexGridLib.MSFlexGrid FG1 
            Height          =   3135
            Left            =   480
            TabIndex        =   11
            Top             =   360
            Width           =   13095
            _ExtentX        =   23098
            _ExtentY        =   5530
            _Version        =   393216
            FixedCols       =   0
         End
      End
      Begin VB.Frame Frame2 
         Height          =   975
         Left            =   600
         TabIndex        =   6
         Top             =   360
         Width           =   14055
         Begin VB.CommandButton cmdVerListado 
            Caption         =   "&Ver Listado"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   11640
            TabIndex        =   1
            Top             =   240
            Width           =   1815
         End
         Begin VB.ComboBox cmbVendedor 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   6720
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   360
            Width           =   2775
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   5520
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Hoy es: Martes, 25 de Mayo de 1810"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   240
            TabIndex        =   7
            Top             =   360
            Width           =   3765
         End
      End
      Begin VB.Label Label5 
         Caption         =   "Label5"
         Height          =   255
         Left            =   240
         TabIndex        =   14
         Top             =   6360
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Label4"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   6000
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Label3"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   5640
         Visible         =   0   'False
         Width           =   615
      End
   End
End
Attribute VB_Name = "FormListadoClientesPorVendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub FormatearFila()
        
        FG1.RowHeight(FG1.Row) = 300
        
        FG1.Col = 0
        FG1.CellBackColor = QBColor(1)
        FG1.CellForeColor = QBColor(7)
        FG1.CellFontSize = 12
        FG1.CellFontBold = True
        
        FG1.Col = 1
        FG1.CellBackColor = QBColor(1)
        FG1.CellForeColor = QBColor(7)
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 2
        FG1.CellBackColor = QBColor(1)
        FG1.CellForeColor = QBColor(7)
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 3
        FG1.CellBackColor = QBColor(1)
        FG1.CellForeColor = QBColor(7)
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 4
        FG1.CellBackColor = QBColor(1)
        FG1.CellForeColor = QBColor(7)
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 5
        FG1.CellBackColor = QBColor(1)
        FG1.CellForeColor = QBColor(7)
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 6
        FG1.CellBackColor = QBColor(1)
        FG1.CellForeColor = QBColor(7)
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

'        FG1.Col = 7
'        FG1.CellBackColor = QBColor(1)
'        FG1.CellForeColor = QBColor(7)
'        FG1.CellFontSize = 12
'        FG1.CellFontBold = True
        
        FG1.CellFontSize = 10

End Sub

Private Sub FormatearFilaTotales()
        FG1.RowHeight(FG1.Row) = 300
        
        FG1.Col = 0
        FG1.CellBackColor = QBColor(8)
        FG1.CellForeColor = vbWhite
        FG1.CellFontSize = 12
        FG1.CellFontBold = True
        
        FG1.Col = 1
        FG1.CellBackColor = QBColor(8)
        FG1.CellForeColor = vbWhite
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 2
        FG1.CellBackColor = QBColor(8)
        FG1.CellForeColor = vbWhite
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 3
        FG1.CellBackColor = QBColor(8)
        FG1.CellForeColor = vbWhite
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 4
        FG1.CellBackColor = QBColor(8)
        FG1.CellForeColor = vbWhite
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 5
        FG1.CellBackColor = QBColor(8)
        FG1.CellForeColor = vbWhite
        FG1.CellFontSize = 12
        FG1.CellFontBold = True

        FG1.Col = 6
        FG1.CellBackColor = QBColor(8)
        FG1.CellFontSize = 12
        FG1.CellForeColor = vbWhite
        FG1.CellFontBold = True

'        FG1.Col = 7
'        FG1.CellBackColor = QBColor(8)
'        FG1.CellFontSize = 12
'        FG1.CellForeColor = vbWhite
'        FG1.CellFontBold = True
        
        FG1.CellFontSize = 10

End Sub


Private Sub LlenarGrilla()

    FG1.Text = qLCV!idcliente

End Sub

Private Sub SeteoGrilla()

    FG1.Rows = 2
    FG1.Cols = 7
           
    FG1.Row = 0
    FG1.Col = 0
   ' FG1.ColWidth(0) = 1500
   ' FG1.CellAlignment = 4
   ' FG1.Text = "VENDEDOR"
    
    FG1.ColWidth(0) = 600
    FG1.Col = 0
    FG1.Text = "ZONA"
           
    FG1.Col = 1
    FG1.ColWidth(1) = 800
    FG1.CellAlignment = 4
    FG1.Text = "CLIENTE"
               
    FG1.Col = 2
    FG1.ColWidth(2) = 4000
    FG1.CellAlignment = 4
    FG1.Text = "NOMBRE / RS"
                                  
    FG1.Col = 3
    FG1.CellAlignment = 4
    FG1.ColWidth(3) = 1800
    FG1.Text = "SALDO L1"

    FG1.Col = 4
    FG1.CellAlignment = 4
    FG1.ColWidth(4) = 1800
    FG1.Text = "SALDO L2"

    FG1.Col = 5
    FG1.CellAlignment = 4
    FG1.ColWidth(5) = 1800
    FG1.Text = "SALDO TOTAL"
    
    FG1.Col = 6
    FG1.ColWidth(6) = 1300
    FG1.CellAlignment = 4
    FG1.Text = "ULT. ACT."

End Sub

Private Sub cmbVendedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If


End Sub


Private Sub cmbVendedor_LostFocus()

    'global vLeg As Integer
    
    Set tEmpleados = BaseSPC.OpenRecordset("Empleados", dbOpenTable)
    
    tEmpleados.Index = "IndiceNombre"
    
    tEmpleados.Seek "=", cmbVendedor.Text
  
    If Not tEmpleados.NoMatch Then
        vLeg = tEmpleados!Legajo
    End If
    
    tEmpleados.Close
    
End Sub

Private Sub cmdExportar_Click()

    Dim NombreArchivo As String
    
    NombreArchivo = "\SALDOS_" + cmbVendedor.Text + "_" + Format(Date, "yyyy-MM-dd") + ".xlsx"
    
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
    With FG1
        For Fila = 1 To .Rows - 1
            'If linea11 = 1 Then
            '    For Columna = 0 To .Cols - 3
            '        o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            '    Next
            'End If
            
            'If linea22 = 1 Then
                
            '    For Columna = 0 To .Cols - 2
            '
            '        .ColWidth(3) = 0
            '        .Col = 3
            '        .Visible = False
            '        o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            '    Next
            'End If
            For Columna = 0 To .Cols - 2
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            Next
        
        Next
    End With
    'o_Libro.Close True, sOutputPath
    
    o_Libro.Close True, sOutputPath
    Set o_Libro = o_Excel.Workbooks.Open(sOutputPath)
    o_Excel.Visible = True
    
    'Call blanqueototal
    ' -- Cerrar Excel
    'o_Excel.Quit
    ' -- Terminar instancias
    'Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    'Exportar_Excel = True
Exit Function

  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
    
End Function
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub

Private Sub cmdExportar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdImprimir_Click()
    
    Dim objPrinterFlex As PrinterFlex
    Set objPrinterFlex = New PrinterFlex
    
    With objPrinterFlex
      
      'Asignamos los valores de los encabezados, el pie de página, el color_
      'del texto y el tamaño de la fuente
        
        'texto de los encabezdos y el pie de pagina
        
        
        .TextEncabezado1 = Chr(9) & "SALDOS DE CLIENTES POR VENDEDOR"
            
                    nVendedor = Chr(9) & cmbVendedor.Text
                    Pie = Chr(9) & Chr(9) & Chr(9) & "Totales Vendedor -> L1" & Chr(9) & FormatCurrency(Label3.Caption, 2) & Chr(9) & "-> L2" & Chr(9) & FormatCurrency(Label4.Caption, 2) & Chr(9) & "-> Total" & Chr(9) & FormatCurrency(Label5.Caption, 2)
                    'Pie = "Desarrollado por SPC Software Integral"
        
'        If OptionL1.Value = True Then li = "Ventas de Línea 1"
'        If OptionL2.Value = True Then li = "Ventas de Línea 2"
'        If OptionAll.Value = True Then li = "Todas las Ventas"
        
        .TextEncabezado2 = Chr(9) & Chr(9) & "VENDEDOR: " & cmbVendedor.Text & Chr(10) & Chr(9) & Chr(9) & Format(Date, "dd - mmmm - yyyy")
                
        'CGrid.Row = 1
        'CGrid.Col = 10
        'Anio = CGrid.Text
        'CGrid.Col = 11
        'Periodo = CGrid.Text
        
        .TextPiePagina = Pie
               
        'Colores de la fuentes
        .ColorPiePagina = QBColor(2)
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
        '.AjustarColumnas = False
      
        .Orientacion = Vertical
        'Imprimimos pasando el nombre del FlexGrid a imprimir
        .ImprimirFlexGrid FG1
        
    End With
    
    'Call objPrinterFlex.ImprimirFlexGrid(CGrid)

End Sub

Private Sub cmdImprimir_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdSaldoAFecha_Click()

    FormSaldosAFecha.Show

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub cmdSalir_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmdVerListado_Click()

    Dim qLCV
    Dim Zona
    Dim TotalZonaL1 As Double
    Dim TotalZonaL2 As Double
    Dim TotalZonaTotal As Double
    Dim TotalVendedorL1 As Double
    Dim TotalVendedorL2 As Double
    Dim TotalVendedorTotal As Double

    Call SeteoGrilla

  '  On Error GoTo CapturaErr
    
    vSQL = "SELECT * FROM qListaClientesVendedor WHERE Legajo='" & vLeg & "' Order By ZonaVenta, IdCliente"

   ' MsgBox (vSQL)
    
    Set qLCV = BaseSPC.OpenRecordset(vSQL)
    
    qLCV.MoveFirst
    qLCV.MoveLast
    
    'FG1.Rows = qLCV.RecordCount + 1
    FG1.Rows = FG1.Rows + 1
    
    qLCV.MoveFirst
    
    FG1.Row = 1
'    FG1.Col = 0
    
 '   FG1.Text = cmbVendedor.Text
    
    Zona = qLCV!ZonaVenta
    
    FG1.Col = 0
    
    If Zona <> "" Then
        FG1.Text = Zona
    End If
    
    TotalZonaL1 = 0
    TotalZonaL2 = 0
    TotalZonaTotal = 0
    TotalVendedorL1 = 0
    TotalVendedorL2 = 0
    TotalVendedorTotal = 0
    
    While Not qLCV.EOF
        If Zona = qLCV!ZonaVenta Then
            FG1.Col = 1
            FG1.Text = qLCV!idcliente
            
            FG1.Col = 2
            FG1.Text = qLCV!RazonSocial
            
            FG1.Col = 3
            FG1.CellAlignment = 7
            FG1.Text = Format(qLCV!SaldoL1, "Currency")
            If FG1.Text < 0 Then FG1.CellForeColor = vbRed
            TotalZonaL1 = TotalZonaL1 + qLCV!SaldoL1
            TotalVendedorL1 = TotalVendedorL1 + qLCV!SaldoL1
            
            FG1.Col = 4
            FG1.CellAlignment = 7
            FG1.Text = Format(qLCV!SaldoL2, "Currency")
            If FG1.Text < 0 Then FG1.CellForeColor = vbRed
            TotalZonaL2 = TotalZonaL2 + qLCV!SaldoL2
            TotalVendedorL2 = TotalVendedorL2 + qLCV!SaldoL2
            
            FG1.Col = 5
            FG1.CellAlignment = 7
            FG1.Text = Format(qLCV!SaldoTotal, "Currency")
            If FG1.Text < 0 Then FG1.CellForeColor = vbRed
            TotalZonaTotal = TotalZonaTotal + qLCV!SaldoTotal
            TotalVendedorTotal = TotalVendedorTotal + qLCV!SaldoTotal
            
            FG1.Col = 6
            FG1.CellAlignment = 7
            FG1.Text = qLCV!FechaActSaldo
         Else
            'Imprimo SubTotales por Zona
                Call FormatearFila
                FG1.Col = 2
                FG1.CellAlignment = 7
                FG1.Text = "Total Zona: " & Zona
                
                FG1.Col = 3
                FG1.CellAlignment = 7
                FG1.Text = Format(TotalZonaL1, "Currency")
                If FG1.Text < 0 Then FG1.CellForeColor = vbRed
                
                FG1.Col = 4
                FG1.CellAlignment = 7
                FG1.Text = Format(TotalZonaL2, "Currency")
                If FG1.Text < 0 Then FG1.CellForeColor = vbRed
                
                FG1.Col = 5
                FG1.CellAlignment = 7
                FG1.Text = Format(TotalZonaTotal, "Currency")
                If FG1.Text < 0 Then FG1.CellForeColor = vbRed
            
            'Agrego una fila
                FG1.Rows = FG1.Rows + 1
                FG1.Row = FG1.Row + 1
                TotalZonaL1 = 0
                TotalZonaL2 = 0
                TotalZonaTotal = 0
            'Continúo con el detalle
            Zona = qLCV!ZonaVenta
            FG1.Col = 0
            
            If Zona <> "" Then
                FG1.Text = Zona
            End If
            
            FG1.Col = 1
            FG1.Text = qLCV!idcliente
            
            FG1.Col = 2
            FG1.Text = qLCV!RazonSocial
            
            FG1.Col = 3
            FG1.CellAlignment = 7
            FG1.Text = Format(qLCV!SaldoL1, "Currency")
            If FG1.Text < 0 Then FG1.CellForeColor = vbRed
            TotalZonaL1 = TotalZonaL1 + qLCV!SaldoL1
            TotalVendedorL1 = TotalVendedorL1 + qLCV!SaldoL1
            
            FG1.Col = 4
            FG1.CellAlignment = 7
            FG1.Text = Format(qLCV!SaldoL2, "Currency")
            If FG1.Text < 0 Then FG1.CellForeColor = vbRed
            TotalZonaL2 = TotalZonaL2 + qLCV!SaldoL2
            TotalVendedorL2 = TotalVendedorL2 + qLCV!SaldoL2
            
            FG1.Col = 5
            FG1.CellAlignment = 7
            FG1.Text = Format(qLCV!SaldoTotal, "Currency")
            If FG1.Text < 0 Then FG1.CellForeColor = vbRed
            TotalZonaTotal = TotalZonaTotal + qLCV!SaldoTotal
            TotalVendedorTotal = TotalVendedorTotal + qLCV!SaldoTotal
            
            FG1.Col = 6
            FG1.CellAlignment = 7
            FG1.Text = qLCV!FechaActSaldo
        End If
        qLCV.MoveNext
        FG1.Rows = FG1.Rows + 1
        FG1.Row = FG1.Row + 1
    Wend
    
    'Imprimo Totales por Vendedor
        FG1.Rows = FG1.Rows + 1
        FG1.Row = FG1.Row + 1
        
        Call FormatearFilaTotales
        FG1.Col = 2
        FG1.CellAlignment = 7
        FG1.Text = "Total Vendedor:"
        
        FG1.Col = 3
        FG1.CellAlignment = 7
        FG1.Text = Format(TotalVendedorL1, "Currency")
        If FG1.Text < 0 Then FG1.CellForeColor = vbRed
        
        FG1.Col = 4
        FG1.CellAlignment = 7
        FG1.Text = Format(TotalVendedorL2, "Currency")
        If FG1.Text < 0 Then FG1.CellForeColor = vbRed
        FG1.Col = 5
        FG1.CellAlignment = 7
        FG1.Text = Format(TotalVendedorTotal, "Currency")
        If FG1.Text < 0 Then FG1.CellForeColor = vbRed
        
        Label3.Caption = TotalVendedorL1
        Label4.Caption = TotalVendedorL2
        Label5.Caption = TotalVendedorTotal
    
CapturaErr:

    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub

Private Sub cmdVerListado_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

   'Seteo Label de fecha
        Label1.Caption = "Hoy es " & Format(Date, "dddd") & "," & Format(Date, "dd") & " de " & Format(Date, "mmmm") & " de " & Format(Date, "yyyy")

   'Seteo tamaño y ubicacion del form
        FormListadoClientesPorVendedor.Height = 7445
        FormListadoClientesPorVendedor.Width = 15000
        FormListadoClientesPorVendedor.Top = 1000
        FormListadoClientesPorVendedor.Left = 1000
        
        
    'Abro Base de Datos
        'Seteo la captura de errores de no hay registros en el archivo
'         On Error GoTo CapturaErrores
        
        Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        
        'Tabla Empleados
            Set tVendedores = BaseSPC.OpenRecordset("Empleados", dbOpenTable)
        
        'Lleno Combo de vendedores
            While Not tVendedores.EOF
                If tVendedores!IDPuesto = 1 Then cmbVendedor.AddItem tVendedores!Nombre
                tVendedores.MoveNext
            Wend

            tVendedores.Close

End Sub

