VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormTxtPercepciones 
   Caption         =   "Exportar Percepciones TXT"
   ClientHeight    =   4740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11085
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   3480
      Width           =   10335
      Begin VB.CommandButton CmdExit 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   8520
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExportarTXT 
         Caption         =   "&Exportar TXT"
         Height          =   735
         Left            =   960
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Label1"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   480
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10455
      Begin VB.TextBox TxtFechaHasta 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   8280
         TabIndex        =   2
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TxtFechaDesde 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2295
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   8
         FixedCols       =   0
         Enabled         =   0   'False
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
         Left            =   8280
         TabIndex        =   5
         Top             =   240
         Width           =   1095
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
         Left            =   1080
         TabIndex        =   4
         Top             =   240
         Width           =   1140
      End
   End
End
Attribute VB_Name = "FormTxtPercepciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Function BuscarCuit(CodCliente As String) As String

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db1 = DBEngine.OpenDatabase(ruta)
    
    Set tClientes = db1.OpenRecordset("Clientes", dbOpenTable)
    
    tClientes.Index = "PrimaryKey"
    
    tClientes.Seek "=", CodCliente
    
    If Not tClientes.NoMatch Then
        BuscarCuit = tClientes!CUIT
     Else
        BuscarCuit = "99-99999999-9"
    End If
    
    db1.Close

End Function


Private Sub CmdExit_Click()

    Unload FormTxtPercepciones

End Sub

Private Sub CmdExit_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload FormTxtPercepciones
        End
    End If


End Sub


Private Sub CmdExportarTXT_Click()

    Dim per As Boolean
  
    'Variable para usar WSH
        Dim Wscript As Object
      
    'Creamos la referencia para usar Windows Scripting Host
        Set Wscript = CreateObject("WScript.Shell")
        
'        NombreArchivo = Wscript.SpecialFolders("Desktop") & "\VentasCITI_" & Format(TextFechaDesde.Text, "yyyy") & Format(TextFechaDesde.Text, "mm") & Format(TextFechaDesde.Text, "dd")
 '       NombreArchivo = NombreArchivo & "_" & Format(TextFechaHasta.Text, "yyyy") & Format(TextFechaHasta.Text, "mm") & Format(TextFechaHasta.Text, "dd") & ".txt"

  
    ' Le envia el control MsFlexgrid, el path del archivo _
     txt y el delimitador
    'CAMINO = App.Path & "\AR-30708432543-"
    CAMINO = Wscript.SpecialFolders("Desktop") & "\AR-30708432543-"
    CAMINO = CAMINO + Format(Date, "yyyy")
    CAMINO = CAMINO + Format(Date, "mm") + "-LOTE"
    
    Set bd = DBEngine.OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    Set UN = bd.OpenRecordset("UltimosNumeros", dbOpenDynaset)
         UN.FindFirst "IDTabla=" + "'tLotePercepciones'"
        
         If Not UN.NoMatch Then
        
            CAMINO = CAMINO + Trim(Str(UN!UltimoNumero + 1)) + ".txt"
            
            UN.Edit
            UN!UltimoNumero = UN!UltimoNumero + 1
            UN.Update
            
         End If
    
    MsgBox (CAMINO)
    UN.Close
    bd.Close
    
    per = Exportar_FlexGrid_txt(FG1, CAMINO, vbTab)
      
    If per Then
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
            Print #Free_File, vbNullString & FG1.text & vbNullString;
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

Private Sub CmdExportarTXT_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload FormTxtPercepciones
        End
    End If

End Sub

Private Sub Form_Load()

    TxtFechaDesde.text = FormLibroIvaVentas.TextFechaDesde.text
    TxtFechaHasta.text = FormLibroIvaVentas.TextFechaHasta.text
    
    Label1.Caption = "©® SPC Software Integral 2015 / 2016 All Rights Reserved"
    Label1.FontBold = True
    
    Call busco

   
    
End Sub
Private Sub busco()

    
'*************** Busco en Facturas y Nota de Credito **********************
    
On Error GoTo Error_Handler

    Set db = DBEngine.OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
      
    Desde = "#" & Format$(TxtFechaDesde.text, "mm/dd/yyyy") & "#"
    Hasta = "#" & Format$(TxtFechaHasta.text, "mm/dd/yyyy") & "#"
    
    eseqele = "SELECT * FROM FacturaC WHERE FechaFactura >=" & Desde & " AND FechaFactura <=" & Hasta & " AND ImportePercepIIBB > 0 Order By TipoFactura, NroFactura, FechaFactura"
    eseqele2 = "SELECT * FROM NotaCreditoC WHERE FechaNotaCredito >=" & Desde & " AND FechaNotaCredito <=" & Hasta & " AND ImportePercepIIBB > 0 Order By TipoNotaCredito, NroNotaCredito, FechaNotaCredito"
    eseqele3 = "SELECT * FROM NotaDebitoC WHERE FechaDebito >=" & Desde & " AND FechaDebito <=" & Hasta & " AND ImportePercepIIBB > 0 Order By TipoDebito, NroDebito, FechaDebito"
    
    'MsgBox (eseqele)
    'MsgBox (eseqele2)
    
    Set rst = db.OpenRecordset(eseqele, dbOpenDynaset)
    Set rst2 = db.OpenRecordset(eseqele2, dbOpenDynaset)
    Set rst3 = db.OpenRecordset(eseqele3, dbOpenDynaset)
   
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
          FG1.text = Format(BuscarCuit(rst.Fields!CodCliente), "#00-00000000-0")
          FG1.Col = 1
          FG1.text = Format(rst.Fields!FechaFactura, "dd/mm/yyyy")
          FG1.Col = 2
          FG1.text = "F"
          FG1.Col = 3
          FG1.text = rst.Fields!TipoFactura
          FG1.Col = 4
          FG1.text = "0004"
          FG1.Col = 5
          FG1.text = Format(rst.Fields!NroFactura, "#00000000")
          FG1.Col = 6
          FG1.text = Format(rst.Fields!SubTotalFactura, "#000000000.00")
          FG1.text = Replace(FG1.text, ",", ".")
          FG1.Col = 7
          FG1.text = Format(rst.Fields!ImportePercepIIBB, "#00000000.00")
          FG1.text = Replace(FG1.text, ",", ".")
          FG1.Col = 8
          FG1.text = "A"
          
          linea2 = linea2 + 1
          
          rst.MoveNext
       Wend
       
       While Not rst2.EOF
          FG1.AddItem " "
          FG1.Row = linea2
          FG1.Col = 0
          FG1.text = Format(BuscarCuit(rst2.Fields!CodCliente), "#00-00000000-0")
          FG1.Col = 1
          FG1.text = Format(rst2.Fields!FechaNotaCredito, "dd/mm/yyyy")
          FG1.Col = 2
          FG1.text = "C"
          FG1.Col = 3
          FG1.text = rst2.Fields!TipoNotaCredito
          FG1.Col = 4
          FG1.text = "0004"
          FG1.Col = 5
          FG1.text = Format(rst2.Fields!NroNotaCredito, "#00000000")
          FG1.Col = 6
          FG1.text = (Format(rst2.Fields!SubTotalNotaCredito, "-#00000000.00"))
          FG1.text = Replace(FG1.text, ",", ".")
          FG1.Col = 7
          FG1.text = (Format(rst2.Fields!ImportePercepIIBB, "-#0000000.00"))
          FG1.text = Replace(FG1.text, ",", ".")
          FG1.Col = 8
          FG1.text = "A"
          
          linea2 = linea2 + 1
          
          rst2.MoveNext
       Wend
               
       While Not rst3.EOF
          FG1.AddItem " "
          FG1.Row = linea2
          FG1.Col = 0
          FG1.text = Format(BuscarCuit(rst3.Fields!CodCliente), "#00-00000000-0")
          FG1.Col = 1
          FG1.text = Format(rst3.Fields!FechaDebito, "dd/mm/yyyy")
          FG1.Col = 2
          FG1.text = "D"
          FG1.Col = 3
          FG1.text = rst3.Fields!TipoDebito
          FG1.Col = 4
          FG1.text = "0004"
          FG1.Col = 5
          FG1.text = Format(rst3.Fields!NroDebito, "#00000000")
          FG1.Col = 6
          FG1.text = (Format(rst3.Fields!SubTotalDebito, "#00000000.00"))
          FG1.text = Replace(FG1.text, ",", ".")
          FG1.Col = 7
          FG1.text = (Format(rst3.Fields!ImportePercepIIBB, "#0000000.00"))
          FG1.text = Replace(FG1.text, ",", ".")
          FG1.Col = 8
          FG1.text = "A"
          
          linea2 = linea2 + 1
          
          rst3.MoveNext
       Wend
              
       FG1.Enabled = True
       
Error_Handler:
    If Err = 3021 Or Err = 440 Then
        'Nada solo para capturar el error.
    End If
    
    Exit Sub

End Sub

Sub SeteoGrilla()
    
    'FG1.AutoSizeMode = klexAutoSizeColWidth
    'FG1.AutoSizeMode = True
    FG1.Cols = 9
    FG1.Row = 0
    FG1.Col = 0
    
    FG1.ColWidth(0) = 1200
    FG1.ColAlignment(0) = flexAlignCenterCenter
    FG1.text = "Cuit"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 1500
    FG1.text = "Fecha Percepcion"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 1500
    FG1.text = "Tipo Comprobante"
    FG1.ColAlignment(2) = flexAlignCenterCenter
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1500
    FG1.text = "Letra Comprobante"
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 900
    FG1.text = "Sucursal"
    FG1.ColAlignment(4) = flexAlignCenterCenter
        
    FG1.Col = 5
    FG1.ColWidth(5) = 1300
    FG1.text = "Numero Emision"
    FG1.ColAlignment(5) = flexAlignCenterCenter
    
    FG1.Col = 6
    FG1.ColWidth(6) = 1500
    FG1.text = "Monto Imponible"
    FG1.ColAlignment(6) = flexAlignCenterCenter
    
    FG1.Col = 7
    FG1.ColWidth(7) = 1500
    FG1.text = "Importe Percepcion"
    FG1.ColAlignment(7) = flexAlignCenterCenter
    
    FG1.Col = 8
    FG1.ColWidth(8) = 1300
    FG1.text = "Tipo Operacion"
    FG1.ColAlignment(8) = flexAlignCenterCenter
    
End Sub

Private Sub txtFechaDesde_GotFocus()

   TxtFechaDesde.SelLength = Len(TxtFechaDesde.text)

End Sub


Private Sub txtFechaDesde_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload FormTxtPercepciones
        End
    End If

End Sub


Private Sub txtFechaHasta_GotFocus()

    TxtFechaHasta.SelLength = Len(TxtFechaHasta.text)

End Sub


Private Sub txtFechaHasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 0
       Sendkeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload FormTxtPercepciones
        End
    End If

End Sub


