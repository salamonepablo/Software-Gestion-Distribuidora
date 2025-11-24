VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormAExcel 
   Caption         =   "Exportacion de Retenciones de Proveedores a Excel"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4860
   ScaleWidth      =   8535
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   8
      Top             =   3600
      Width           =   8295
      Begin VB.CommandButton CmdExportarTXT 
         Caption         =   "&Exportar TXT"
         Height          =   735
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   6600
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExportar 
         Caption         =   "&Exportar Excel"
         Height          =   735
         Left            =   3720
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         Height          =   735
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   8295
      Begin VB.TextBox TxtFechaDesde 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtFechaHasta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   4920
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2295
         Left            =   240
         TabIndex        =   5
         Top             =   960
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   5
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
         Left            =   1920
         TabIndex        =   7
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
         Left            =   4920
         TabIndex        =   6
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
    
'On Error GoTo Error_Handler

    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
      
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
       
       While Not rst.EOF
          FG1.AddItem " "
          FG1.Row = linea2
          FG1.Col = 0
          FG1.Text = Format(rst.Fields!Cuit, "#00-00000000-0")
          FG1.Col = 1
          FG1.Text = rst.Fields!FechaPago
          FG1.Col = 2
          FG1.Text = Format(rst.Fields!NroPago, "#000100000000")
          FG1.Col = 3
          FG1.Text = Format(rst.Fields!TotalReten, "#00000000.00")
          FG1.Col = 4
          FG1.Text = "A"
          linea2 = linea2 + 1
         ' rst.FindNext
          rst.MoveNext
       Wend
       
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
    'Esto lo marco PVS 2014-08-24 o_Excel.Quit
    ' -- Terminar instancias
    
    Set o_Libro = o_Excel.Workbooks.Open(sOutputPath)
    o_Excel.Visible = True

    
    
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

Private Sub CmdExportarTXT_Click()

    Dim ret As Boolean
  
    ' Le envia el control MsFlexgrid, el path del archivo _
     txt y el delimitador
    CAMINO = App.Path & "\AR-30708432543-"
    CAMINO = CAMINO + Format(Date, "yyyy")
    CAMINO = CAMINO + Format(Date, "mm") + "-LOTE"
    
    Set Padron = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
    Set UN = Padron.OpenRecordset("UltNums", dbOpenDynaset)
         UN.FindFirst "TipoDoc=" + "'RET'"
        
         If Not UN.NoMatch Then
        
            CAMINO = CAMINO + Trim(Str(UN!UltNum + 1)) + ".txt"
            
            UN.Edit
            UN!UltNum = UN!UltNum + 1
            UN.Update
            
         End If
    
    MsgBox (CAMINO)
    UN.Close
    Padron.Close
    
    ret = Exportar_FlexGrid_txt(FG1, CAMINO, vbTab)
      
    If ret Then
        MsgBox "Archivo ARBA " & CAMINO & " Generado con éxito", vbInformation
    End If
End Sub
Public Function Exportar_FlexGrid_txt(FlexGrid As Object, ByVal Path_Txt As String, Delimitador As Variant) As Boolean
  
    On Error GoTo Err_Funcion
    Dim Fila As Integer
    Dim Columna As Integer
    Dim Free_File As Integer
   
    ' Número de  archivo libre para crear el archivo de texto
    Free_File = FreeFile
    ' Abre y crea el  archivo
    Open Path_Txt For Output As #Free_File
      
    ' Recorre las filas del Flexgrid
    For Fila = 1 To _
        FG1.Rows - 2
        FG1.Row = Fila
          
        ' Recorre las columnas
        For Columna = 0 To _
            FG1.Cols - 1
            FG1.Col = Columna
            ' escribe el Delimitador
          '  If Columna > 0 Then
          '      Print #Free_File, Delimitador;
          '  End If
            ' Escribe el dato
            Print #Free_File, vbNullString & FG1.Text & vbNullString;
        Next
          
        Print #Free_File, ""
      
    Next
    Close
    Exportar_FG1_txt = True
      
    ' Fin
    mensaje = MsgBox("El Archivo se genero sin problemas", vbInformation)
    Exit Function
  
' error
Err_Funcion:
    Close #Free_File
    MsgBox Err.Description, vbCritical
End Function

Private Sub Form_Load()

    FormAExcel.Height = 5430
    FormAExcel.Width = 8775
    
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
    FG1.ColWidth(1) = 1800
    FG1.Text = "Fecha Pago"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 1800
    FG1.Text = "Nro Pago"
    FG1.ColAlignment(2) = flexAlignCenterCenter
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1800
    FG1.Text = "Total Retenido"
    FG1.ColAlignment(3) = flexAlignCenterCenter
        
    FG1.Col = 4
    FG1.ColWidth(4) = 300
    FG1.Text = "A"
    FG1.ColAlignment(4) = flexAlignCenterCenter
    
    
End Sub

Private Sub TxtFechaDesde_KeyPress(KeyAscii As Integer)
 
 If KeyAscii = 13 Then
    KeyAscii = 0
    SendKeys "{TAB}"
 End If
 
End Sub

Private Sub TxtFechaHasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
End Sub
