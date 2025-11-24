VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FormSaldosAFecha 
   Caption         =   "Saldos de Todos los Clientes a una Fecha Solicitada"
   ClientHeight    =   8730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15345
   LinkTopic       =   "Form2"
   ScaleHeight     =   8730
   ScaleWidth      =   15345
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   14655
      Begin VB.Frame Frame2 
         Height          =   855
         Left            =   960
         TabIndex        =   6
         Top             =   7080
         Width           =   12615
         Begin VB.CommandButton cmdExcel 
            Caption         =   "&Excel"
            Height          =   495
            Left            =   840
            TabIndex        =   8
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdSalir 
            Caption         =   "&Salir"
            Height          =   495
            Left            =   9000
            TabIndex        =   7
            Top             =   240
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdSaldos 
         Caption         =   "&Ver Saldos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   5
         Top             =   360
         Width           =   2175
      End
      Begin MSComCtl2.MonthView dateSelect 
         Height          =   2310
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         StartOfWeek     =   121438209
         CurrentDate     =   43950
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   5775
         Left            =   600
         TabIndex        =   4
         Top             =   1080
         Width           =   13455
         _ExtentX        =   23733
         _ExtentY        =   10186
         _Version        =   393216
         FixedCols       =   0
         AllowUserResizing=   1
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Text            =   "29/02/2020"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   300
         Left            =   9000
         TabIndex        =   10
         Top             =   480
         Width           =   165
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Clientes Procesados:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Left            =   6360
         TabIndex        =   9
         Top             =   480
         Width           =   2550
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         TabIndex        =   1
         Top             =   480
         Width           =   600
      End
   End
End
Attribute VB_Name = "FormSaldosAFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub SeteoGrilla()

    FG1.Rows = 2
    FG1.FixedRows = 1
    FG1.Cols = 7

    FG1.Col = 0
    FG1.Row = 0
    FG1.CellAlignment = 3
    FG1.ColWidth(0) = 1000
    FG1.Text = "Cliente"
    
    FG1.Col = 1
    FG1.CellAlignment = 3
    FG1.ColWidth(1) = 3500
    FG1.Text = "Nombre"
    
    FG1.Col = 2
    FG1.CellAlignment = 3
    FG1.ColWidth(2) = 1500
    FG1.Text = "Saldo L1"

    FG1.Col = 3
    FG1.CellAlignment = 3
    FG1.ColWidth(3) = 1500
    FG1.Text = "Saldo L2"
    
    FG1.Col = 4
    FG1.CellAlignment = 3
    FG1.ColWidth(4) = 1500
    FG1.Text = "Saldo Total"
    
    FG1.Col = 5
    FG1.CellAlignment = 3
    FG1.ColWidth(5) = 1500
    FG1.Text = "Fecha Consulta"
    
    FG1.Col = 6
    FG1.CellAlignment = 3
    FG1.ColWidth(6) = 2500
    FG1.Text = "Vendedor"
    
End Sub

Private Sub cmdExcel_Click()
    
    Dim Nombre As String
    
    Nombre = "\Saldos_al_" & Format(txtFecha.Text, "DD-MM-YYYY") & ".xlsx"
    
    'MsgBox (Nombre)
    
    If Exportar_Excel(App.Path & Nombre, FG1) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If

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
            For Columna = 0 To .Cols - 1
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
Private Sub cmdSaldos_Click()

    dateSelect.Visible = False
    
    If Not IsDate(txtFecha.Text) Then
        MsgBox ("ERROR EN FECHA")
        Exit Sub
    End If
    
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    'Set db = DBEngine.OpenDatabase(ruta)
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenDynaset)
    
    tClientes.MoveFirst
    Call SeteoGrilla
        
    Label1.FontSize = 12
    Label1.ForeColor = vbBlue
    
    While Not tClientes.EOF
        Call BuscarSaldos(tClientes!IdCliente)
        tClientes.MoveNext
        cont = cont + 1
        Label2.Caption = cont
        FormSaldosAFecha.Refresh
    Wend
       
    'Label2.Caption = cont
    'FormSaldosAFecha.Refresh
    
    MsgBox ("TOTAL DE CLIENTES PROCESADOS: " & CStr(cont))

End Sub
Private Sub BuscarSaldos(idC As Long)

On Error GoTo Error_Handler

Dim SaldoL1, SaldoL2, SaldoTotal As Double
Dim cont As Integer
    
    SaldoL1 = 0
    SaldoL2 = 0
    SaldoTotal = 0
    cont = 0

    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")

    vSQL = "SELECT * FROM MovimientosCtaCte WHERE IDCliente =" & idC & " ORDER BY Fecha ASC"
    
    Set tMovCC = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    tMovCC.MoveFirst
    
    While (Not tMovCC.EOF) And (CVDate(tMovCC!Fecha) <= CVDate(txtFecha.Text))
        SaldoL1 = SaldoL1 + tMovCC!ImporteLinea1
        SaldoL2 = SaldoL2 + tMovCC!ImporteLinea2
        tMovCC.MoveNext
    Wend
    
        SaldoTotal = SaldoL1 + SaldoL2
        FG1.Rows = FG1.Rows + 1
        FG1.Col = 0
        FG1.ColAlignment(0) = flexAlignRightCenter
        FG1.Text = idC
        FG1.Col = 1
        FG1.ColAlignment(1) = flexAlignLeftCenter
        FG1.Text = tClientes!RazonSocial
        FG1.Col = 2
        FG1.ColAlignment(2) = flexAlignRightCenter
        FG1.Text = Format$(SaldoL1, "Currency")
        FG1.Col = 3
        FG1.ColAlignment(3) = flexAlignRightCenter
        FG1.Text = Format$(SaldoL2, "Currency")
        FG1.ColAlignment(4) = flexAlignRightCenter
        FG1.Col = 4
        FG1.Text = Format$(SaldoTotal, "Currency")
        FG1.Col = 5
        FG1.ColAlignment(5) = flexAlignCenterCenter
        FG1.Text = txtFecha.Text
        
        Set tVendedores = BaseSPC.OpenRecordset("Empleados", dbOpenTable)
        tVendedores.Index = "PrimaryKey"
        tVendedores.Seek "=", tClientes!Vendedor
        If Not tVendedores.NoMatch Then
           FG1.Col = 6
           FG1.ColAlignment(6) = flexAlignLeftCenter
           FG1.Text = tVendedores!Legajo & " - " & tVendedores!Nombre
        End If
        tVendedores.Close
        
        FG1.Row = FG1.Row + 1
         
Error_Handler:
    
'flexAlignLeftCenter
'MSFlexGrid1.ColAlignment(1) = flexAlignCenterCenter
'MSFlexGrid1.ColAlignment(2) = flexAlignRightCenter
    
    
    
    If Err = 3021 Or Err = 440 Then
        SaldoTotal = SaldoL1 + SaldoL2
        FG1.Rows = FG1.Rows + 1
        FG1.Col = 0
        FG1.ColAlignment(0) = flexAlignRightCenter
        FG1.Text = idC
        FG1.Col = 1
        FG1.ColAlignment(1) = flexAlignLeftCenter
        FG1.Text = tClientes!RazonSocial
        FG1.Col = 2
        FG1.ColAlignment(2) = flexAlignRightCenter
        FG1.Text = Format$(SaldoL1, "Currency")
        FG1.Col = 3
        FG1.ColAlignment(3) = flexAlignRightCenter
        FG1.Text = Format$(SaldoL2, "Currency")
        FG1.ColAlignment(4) = flexAlignRightCenter
        FG1.Col = 4
        FG1.Text = Format$(SaldoTotal, "Currency")
        FG1.Col = 5
        FG1.ColAlignment(5) = flexAlignCenterCenter
        FG1.Text = txtFecha.Text
        
        Set tVendedores = BaseSPC.OpenRecordset("Empleados", dbOpenTable)
        tVendedores.Index = "PrimaryKey"
        tVendedores.Seek "=", tClientes!Vendedor
        If Not tVendedores.NoMatch Then
           FG1.Col = 6
           FG1.ColAlignment(6) = flexAlignLeftCenter
           FG1.Text = tVendedores!Legajo & " - " & tVendedores!Nombre
        End If
        tVendedores.Close
        
        FG1.Row = FG1.Row + 1
        
    End If
    
    If Err = 30015 Then
        'Nada solo para capturar el error.
    End If
   
End Sub

Private Sub cmdSalir_Click()
    
    Unload Me

End Sub


Private Sub dateSelect_DateClick(ByVal DateClicked As Date)
    
    txtFecha.Text = ""
    txtFecha.Text = dateSelect.Value
    
End Sub

Private Sub dateSelect_LostFocus()

    dateSelect.Visible = False

End Sub

Private Sub Form_Load()
            
    txtFecha.ToolTipText = "Doble Click Para Calendario"
    txtFecha.Text = Format(Now, "DD/MM/YYYY")
    
    FormSaldosAFecha.Width = 15105
    FormSaldosAFecha.Height = 8970
    
    FormSaldosAFecha.Left = 2500
    FormSaldosAFecha.Top = 500
                
End Sub

Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub
Private Sub txtFecha_Change()

    txtFecha.SelLength = Len(txtFecha.Text)
    
End Sub

Private Sub txtFecha_DblClick()

    dateSelect.Visible = True

End Sub

