VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormLiqComisiones 
   Caption         =   "LIQUIDACION DE COMISIONES"
   ClientHeight    =   6495
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14010
   LinkTopic       =   "Form1"
   ScaleHeight     =   6495
   ScaleWidth      =   14010
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   6255
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   13815
      Begin VB.Timer Timer1 
         Left            =   6240
         Top             =   5640
      End
      Begin VB.ComboBox cmbVendedores 
         Height          =   315
         Index           =   1
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   720
         Width           =   3855
      End
      Begin VB.CommandButton cmdLiquidar 
         Caption         =   "&Liquidar"
         Height          =   495
         Left            =   10920
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.Frame Frame3 
         Caption         =   "Liquidación"
         Height          =   1095
         Left            =   9240
         TabIndex        =   15
         Top             =   4800
         Width           =   3855
         Begin VB.TextBox txtImporteTotal 
            Height          =   375
            Left            =   1560
            TabIndex        =   9
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
            TabIndex        =   16
            Top             =   360
            Width           =   1200
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Acciones"
         Height          =   1095
         Left            =   840
         TabIndex        =   14
         Top             =   4800
         Width           =   5055
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Imprimir"
            Height          =   495
            Left            =   1920
            TabIndex        =   7
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdSalir 
            Caption         =   "&Salir"
            Height          =   495
            Left            =   3360
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
            Height          =   495
            Left            =   480
            TabIndex        =   6
            Top             =   360
            Width           =   1095
         End
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3255
         Left            =   600
         TabIndex        =   5
         Top             =   1320
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
      End
      Begin VB.ComboBox cmbVendedores 
         Height          =   315
         Index           =   0
         Left            =   5760
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   3855
      End
      Begin VB.TextBox txtFechaHasta 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   3240
         TabIndex        =   1
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox txtFechaDesde 
         Alignment       =   2  'Center
         Height          =   315
         Left            =   840
         TabIndex        =   0
         Top             =   720
         Width           =   1455
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
         Left            =   6120
         TabIndex        =   17
         Top             =   5280
         Visible         =   0   'False
         Width           =   1200
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
         Left            =   5400
         TabIndex        =   13
         Top             =   480
         Width           =   1335
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
         Left            =   3000
         TabIndex        =   12
         Top             =   480
         Width           =   1155
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
         TabIndex        =   11
         Top             =   480
         Width           =   1200
      End
   End
End
Attribute VB_Name = "FormLiqComisiones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Legajo

Private Sub SeteoGrilla()
    
    'Tamaño de Grilla
        FG1.Cols = 7
        FG1.Rows = 2
    
    'Alineación de columnas
        FG1.ColAlignment(0) = 4
        FG1.ColAlignment(1) = 4
        FG1.ColAlignment(2) = 4
        FG1.ColAlignment(3) = 4
        FG1.ColAlignment(4) = 4
        FG1.ColAlignment(5) = 4
        FG1.ColAlignment(6) = 4
    
    'Ancho de columnas
        FG1.ColWidth(0) = 1000
        FG1.ColWidth(1) = 1500
        FG1.ColWidth(2) = 1700
        FG1.ColWidth(3) = 1500
        FG1.ColWidth(4) = 3800
        FG1.ColWidth(5) = 1500
        FG1.ColWidth(6) = 1500
            
    'Títulos
        FG1.Row = 0
        FG1.Col = 0
        FG1.Text = "Nro Pago"
        FG1.Col = 1
        FG1.Text = "Fecha"
        FG1.Col = 2
        'FG1.Text = "Total" + Chr(10) + "Abonado"
        FG1.Text = "Importe" + Chr(10) + "Pago"
        FG1.Col = 3
        FG1.Text = "Forma" + Chr(10) + "de Pago"
        FG1.Col = 4
        FG1.Text = "Cliente"
        FG1.Col = 5
        FG1.Text = "Comisión %"
        FG1.Col = 6
        FG1.Text = "Importe" + Chr(10) + "Comisión"
        
    
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
    
    NombreArchivo = "\LiqCom_Vend_" + cmbVendedores(0).Text + "_" + Format(TxtFechaHasta.Text, "MMM-YYYY") + ".xlsx"
    
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
        .TextEncabezado1 = Chr(9) & "LIQUIDACION DE COMISIONES"
            
                    nVendedor = Chr(9) & cmbVendedores(1).Text
                    Pie = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Liquidación Total: " & FormatCurrency(txtImporteTotal.Text, 2)
                    'Pie = "Desarrollado por SPC Software Integral"
        
        .TextEncabezado2 = Chr(9) & nVendedor & Chr(10) & Chr(9) & Chr(9) & "Desde el " & TxtFechaDesde.Text & " al " & TxtFechaHasta.Text
                
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
      
        .Orientacion = Horizontal
        'Imprimimos pasando el nombre del FlexGrid a imprimir
        .ImprimirFlexGrid FG1
    End With
    
    'Call objPrinterFlex.ImprimirFlexGrid(CGrid)
    
    

End Sub

Private Sub cmdLiquidar_Click()

  'Timer1.Enabled = True
  Timer1.Interval = 10000
  'Timer1.Enabled = False
  lblImprimiendo.Visible = False
  
    'lblImprimiendo.Visible = False
    'MsgBox (fin)

 'Selecciono los registros de la consulta que corresponden a la fecha y al vendedor
        
    Dim LiqLinea As Double
    Dim LiqTotal As Double
    Dim FormaPago As String
    Dim TotalFormaPago As Double
        
    LiqLinea = 0
    LiqTotal = 0
    TotalFormaPago = 0
    
        
    FG1.Rows = 2
    
    'On Error GoTo CapturaErrores
    
    FechaDesde = Format(TxtFechaDesde.Text, "m/d/yyyy")
    FechaHasta = Format(TxtFechaHasta.Text, "m/d/yyyy")
    
    vSQL = "SELECT * FROM qLiqComisiones WHERE FechaPago>=#" & FechaDesde & "# AND FechaPago <=#" & FechaHasta & "# AND Legajo='" & cmbVendedores(0).Text & "' ORDER BY FormaPago, FechaPago"
    'MsgBox (vSQL)
    
    Set qComisiones = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    qComisiones.MoveFirst
    
    FormaPago = qComisiones!FormaPago
    
    FG1.Row = 1
    
    While Not qComisiones.EOF
      FG1.CellFontBold = False
      If FormaPago = qComisiones!FormaPago Then
        FG1.Col = 0
        FG1.Text = Format(qComisiones.[PagoC.NroPago], "General Number")
        FG1.Col = 1
        FG1.Text = Format(qComisiones!FechaPago, "DD-MMM-YY")
        FG1.Col = 2
        FG1.CellAlignment = 7
        'FG1.Text = Format$(qComisiones!TotalAbonado, "Standard")
        'FG1.Text = FormatCurrency(qComisiones!TotalAbonado, 2)
        FG1.Text = FormatCurrency(qComisiones!ImportePago, 2)
        TotalFormaPago = TotalFormaPago + qComisiones!ImportePago
        FG1.Col = 3
        FG1.Text = qComisiones!FormaPago
        FG1.Col = 4
        FG1.Text = qComisiones!RazonSocial
        FG1.Col = 5
        FG1.CellAlignment = 7
        FG1.Text = Format$(qComisiones!Comision, "Standard")
        FG1.Col = 6
        'LiqLinea = (qComisiones!TotalAbonado * qComisiones!Comision) / 100
        FG1.CellAlignment = 7
        'LiqLinea = (qComisiones!TotalAbonado * qComisiones!Comision) / 100
        LiqLinea = (qComisiones!ImportePago * qComisiones!Comision) / 100
        
        'FG1.Text = Format$(LiqLinea, "Standard")
        FG1.Text = FormatCurrency(LiqLinea, 2)
        
        LiqTotal = LiqTotal + LiqLinea
        
        qComisiones.MoveNext
        
        FG1.Rows = FG1.Rows + 1
        FG1.Row = FG1.Row + 1
        
       Else
            'Subtotales
                FG1.Col = 2
                FG1.CellFontBold = True
                FG1.Text = "Total en " & FormaPago + ": "
                FG1.Col = 3
                FG1.CellFontBold = True
                FG1.Text = FormatCurrency(TotalFormaPago, 2)
                
                FormaPago = qComisiones!FormaPago
                LiqLinea = 0
                TotalFormaPago = 0
                FG1.Rows = FG1.Rows + 1
                FG1.Row = FG1.Row + 1
       End If
    Wend
        'Ultima Linea
            FG1.Col = 2
            FG1.CellFontBold = True
            FG1.Text = "Total en " & FormaPago + ": "
            FG1.Col = 3
            FG1.CellFontBold = True
            FG1.Text = FormatCurrency(TotalFormaPago, 2)
            FG1.Rows = FG1.Rows + 1
            FG1.Row = FG1.Row + 1
       
        'Totales
            txtImporteTotal.Alignment = 1
            txtImporteTotal.FontBold = True
            txtImporteTotal.FontSize = 10
            'txtImporteTotal.Text = Format$(LiqTotal, "Standard")
            txtImporteTotal.Text = FormatCurrency(LiqTotal, 2)
    
CapturaErrores:
    Select Case Err
        Case 3021
          MsgBox "No hay Pagos Para Liquidar con el Criterio Seleccionado...", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
          Resume Next
    End Select

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub Form_Load()

    'Llenar Combo Vendedores
        Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        vSQL = "SELECT * FROM Empleados Where IdPuesto=1 ORDER BY Nombre"
        Set tVendedores = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
        
        On Error GoTo CapturaErrores
        'MsgBox (vSQL)

        tVendedores.MoveFirst
        
        cmbVendedores(1).Clear
        While Not tVendedores.EOF
            cmbVendedores(0).AddItem (tVendedores!Legajo)
            cmbVendedores(1).AddItem (tVendedores!nombre)
            tVendedores.MoveNext
        Wend
        
        tVendedores.Close
    'Fechas
        TxtFechaDesde.Text = Format(Date, "DD/MM/YYYY")
        TxtFechaHasta.Text = Format(Date, "DD/MM/YYYY")
        
    'Setear Grilla
        Call SeteoGrilla

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "ERRROR !!! No hay Vendedores Ingresados en La Base de Datos...", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select

End Sub


Private Sub Timer1_Timer()
' Cambiar Visible Label BackColor entre rojo y azul.
   
   While Timer1.Interval <> 0
    If lblImprimiendo.ForeColor = QBColor(1) Then
       lblImprimiendo.Visible = True
       lblImprimiendo.ForeColor = QBColor(4)
       Timer1.Interval = Timer1.Interval - 1
    Else
       lblImprimiendo.Visible = True
       lblImprimiendo.ForeColor = QBColor(1)
       Timer1.Interval = Timer1.Interval - 1
    End If
   Wend
    lblImprimiendo.Visible = False
    Timer1.Enabled = False
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


Private Sub txtFechaDesde_LostFocus()

    If Not IsDate(TxtFechaDesde.Text) Then
        MsgBox "ERRROR !!! Formato de Fecha Incorrecto...", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
        'Resume Next
    End If

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


