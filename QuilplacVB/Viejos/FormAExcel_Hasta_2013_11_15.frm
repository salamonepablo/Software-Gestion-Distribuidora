VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormAExcel 
   Caption         =   "Exportacion de Retenciones de Proveedores a Excel"
   ClientHeight    =   4800
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10740
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4800
   ScaleWidth      =   10740
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   240
      TabIndex        =   6
      Top             =   3600
      Width           =   10335
      Begin VB.CommandButton CmdExit 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   9240
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExportar 
         Caption         =   "&Exportar"
         Height          =   735
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10455
      Begin VB.TextBox TxtFechaDesde 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3120
         TabIndex        =   2
         Text            =   "01/01/2013"
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtFechaHasta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   6120
         TabIndex        =   1
         Text            =   "31/12/2013"
         Top             =   480
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2295
         Left            =   0
         TabIndex        =   3
         Top             =   960
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         Enabled         =   0   'False
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde"
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
         Left            =   3120
         TabIndex        =   5
         Top             =   240
         Width           =   1140
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta"
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
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FormAExcel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdBuscar_Click()
    Call busco
End Sub
Private Sub busco()

    
'***************Busco en PagoProvret
    
On Error GoTo Error_Handler

    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
      
    Desde = "#" & Format$(TxtFechaDesde.Text, "mm/dd/yyyy") & "#"
    Hasta = "#" & Format$(TxtFechaHasta.Text, "mm/dd/yyyy") & "#"
    
    eseqele = "SELECT * FROM PagoProvRet WHERE FechaPago >=" & Desde & " AND FechaPago <=" & Hasta & " Order By NroPago, FechaPago "
    
    Set rst = db.OpenRecordset(eseqele, dbOpenDynaset)
   
      FG1.Rows = 2
      FG1.Clear
      FG1.Visible = True
       
       Call SeteoGrilla
       
       rst.MoveFirst

       linea2 = 1
       Do While Not rst.NoMatch
          FG1.AddItem " "
          FG1.Row = linea2
          FG1.Col = 0
          FG1.Text = rst.Fields!Cuit
          FG1.Col = 1
          FG1.Text = rst.Fields!NombreProv
          FG1.Col = 2
          FG1.Text = rst.Fields!FechaPago
          FG1.Col = 3
          FG1.Text = Format(rst.Fields!TotalReten, "#0.00")
          linea2 = linea2 + 1
         ' rst.FindNext
          rst.MoveNext
          
       Loop
       
Error_Handler:
    
    If Err = 3021 Or Err = 440 Then
        'Nada solo para capturar el error.
    End If
    
    Exit Sub
End Sub

Private Sub CmdExit_Click()
    Unload FormAExcel
End Sub

Private Sub CmdExportar_Click()
       If Exportar_Excel(App.Path & "\Retencion.xls", FG1) Then
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
            For Columna = 0 To .Cols - 1
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix(Fila, Columna)
            Next
        Next
    End With
    o_Libro.Close True, sOutputPath
    ' -- Cerrar Excel
    o_Excel.Quit
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
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub

Private Sub Form_Load()

    FormAExcel.Height = 5370
    FormAExcel.Width = 10980
    
    Call SeteoGrilla
    
End Sub

Sub SeteoGrilla()
    
    'FG1.AutoSizeMode = klexAutoSizeColWidth
    FG1.Row = 0
    FG1.Col = 0
    
    FG1.ColWidth(0) = 1800
    FG1.ColAlignment(0) = flexAlignCenterCenter
    FG1.Text = "Cuit"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 4500
    FG1.Text = "Nombre Proveedor"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 1800
    FG1.Text = "Fecha Pago"
    FG1.ColAlignment(2) = flexAlignCenterCenter
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1800
    FG1.Text = "Total Retenido"
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    
End Sub

