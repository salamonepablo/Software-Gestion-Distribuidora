VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormImportTxt 
   Caption         =   "IMPORTAR PADRON"
   ClientHeight    =   3195
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7905
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   7905
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   960
      TabIndex        =   8
      Top             =   1560
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid DG1 
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   2880
      Visible         =   0   'False
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   873
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCompactarBBDD 
      Caption         =   "&Compactar BBDD"
      Height          =   615
      Left            =   5280
      TabIndex        =   6
      Top             =   120
      Width           =   1575
   End
   Begin VB.CommandButton cmdIniciar 
      Caption         =   "&Bajar Padron"
      Height          =   660
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1620
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   495
      Left            =   960
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "&Importar Padron"
      Height          =   615
      Left            =   3000
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
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
      Left            =   3480
      TabIndex        =   7
      Top             =   2160
      Width           =   3615
   End
   Begin VB.Label lblCronometro 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      Height          =   300
      Left            =   960
      TabIndex        =   4
      Top             =   960
      Width           =   1080
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "Label1"
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
      TabIndex        =   3
      Top             =   960
      Width           =   735
   End
End
Attribute VB_Name = "FormImportTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim I As Long   'Contador.
Dim Tiempo As String  'Tiempo total transcurrido.
Private Sub Command1_Click()
    
    
    'Le asignamos  el Datasource del datagrid a la función que devuelve el recordset
        'Set DataGrid1.DataSource = Leer_Txt_con_Ado
    vRta = MsgBox("Va a Borrar los Datos en la tabla PADRON" & Chr(10) & "¿Resguarda la BBDD Original?", vbYesNo, "RESGUARDAD BASE DE DATOS")
    
    If vRta = 6 Then
        SourceDB = App.Path & "\Padron.mdb"
        DestDB = App.Path & "\Padron_Hasta_" & Format(Date, "DD-MM-YYYY") & ".mdb"
        
        FileCopy SourceDB, DestDB  ' Use copy of database; preserve original.
        
    End If
    
    Set bbdd = OpenDatabase(App.Path & "\Padron.mdb")
    bbdd.Execute "DELETE * FROM Padron"
    
    Set TPadron = bbdd.OpenRecordset("Padron", dbOpenTable)
    
    'db.Execute "Delete From BadTable"
    
    Dim cn As New ADODB.Connection
    Dim rs As New ADODB.Recordset
  
  ' cursor del lado del cliente
    cn.CursorLocation = adUseClient
      
    ' abre el archivo de texto datos.txt ubicado en el app.path
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
                     & "Data Source=" & App.Path & ";" _
                    & "Extended Properties='text;HDR=NO;FMT=DELIMITED'"
      
    ' Ejecuta la consulta sql para llenar el recordset
      
    rs.Open "Select * from Padron.txt ", cn, adOpenStatic
    
    rs.MoveFirst
    rs.MoveLast
    
    PgB1.Max = rs.RecordCount
      
    ' enlaza el recordset con el DataGrid
    'DG1.ClearFields
    Set DG1.DataSource = rs
    
    rs.MoveFirst
    cont = 0
    PgB1.Value = 0
    
   
    While Not rs.EOF
    ' Visualizamos el porcentaje en el Label
        With TPadron
            .AddNew
                FG1.Col = 0
                !Regimen = rs!F1
                FG1.Text = !Regimen
                
                FG1.Col = 1
                !FPub = (rs!F2)
                FG1.Text = !FPub
                
                FG1.Col = 2
                !FVigDde = (rs!F3)
                FG1.Text = !FVigDde
                
                FG1.Col = 3
                !FVigHta = (rs!F4)
                FG1.Text = !FVigHta
                
                FG1.Col = 4
                !CUIT = (rs!f5)
                FG1.Text = !CUIT
                
                FG1.Col = 5
                !TipoContrInsc = rs!F6
                FG1.Text = !TipoContrInsc
                
                FG1.Col = 6
                !MarcaAltaBajaSujeto = rs!F7
                FG1.Text = !MarcaAltaBajaSujeto
                
                FG1.Col = 7
                !MCbioAlicuota = rs!f8
                FG1.Text = !MCbioAlicuota
                
                FG1.Col = 8
                !AlicuotaRetencion = Format(rs!f9, "Standard")
                FG1.Text = !AlicuotaRetencion
                
                FG1.Col = 9
                !NroGrupoRetencion = rs!F10
                FG1.Text = !NroGrupoRetencion
                
            .Update
        End With
      
      PgB1.Value = cont + 1
      Label1 = "Registro Nro: " & cont & "   " & CLng((PgB1.Value * 100) / PgB1.Max) & " %"
      rs.MoveNext
      FG1.Rows = FG1.Rows + 1
      FG1.Row = FG1.Row + 1
    
    Wend
        
End Sub

Private Sub cmdCompactarBBDD_Click()

    DatabaseName = "DB_SPC_SI.mdb"
    CommonDatabaseName = "DB_SPC_SI.mdb"
    
    Set db = OpenDatabase(App.Path & "\DB_SPC_SI.MDB")

'Close DB
db.Close

Screen.MousePointer = vbHourglass

'Compact DB
DBEngine.CompactDatabase App.Path & "\" & CommonDatabaseName, App.Path & "\" & CommonDatabaseName & "_"

'Delete old file and rename new one
Kill App.Path & "\" & DatabaseName
Name App.Path & "\" & DatabaseName & "_" As App.Path & "\" & DatabaseName

Screen.MousePointer = vbDefault
End Sub

Sub Compactar()
'Compactar una base de datos con ADO
    Dim sDBTmp As String
   'Dim je As Database
    '
    On Error GoTo ErrCompactar
    '
    Set je = OpenDatabase("DB_SPC_SI.MDB")
'
' Crear un nombre "medio" aleatorio
    sDBTmp = "DB_SPC_SI" & Format$(Minute(Now), "00") & Format$(Second(Now), "00") & ".mdb"
' Asegurarnos de que no existe una base con el nombre temporal
    If Len(Dir$(sDBTmp)) Then
        Kill sDBTmp
    End If
'
    MsgBox " Compactando la base de datos..."

' Compactar la base de datos
    
    je.CompactDatabase "Data Source=" & "DB_SPC_SI.mdb" & ";", "Data Source=" & sDBTmp & ";"
    
'
' Eliminar la base de datos original
    Kill "DB_SPC_SI.mdb"
'
' Renombrar la base temporal con el original
    Name sDBTmp As "DB_SPC_SI.mdb"
'
    MsgBox " Base de datos compactada."
'
    Exit Sub
'
ErrCompactar:
' Mostrar el mensaje de error
    MsgBox "Error al compactar la base de datos:" & vbCrLf & _
    Err.Number & " " & Err.Description, _
    vbExclamation, "Error al compactar la base de datos"
    Err.Clear
    MsgBox " *** Error al compactar la base de datos ***"

End Sub

Private Sub cmdIniciar_Click()
    
 '   I = 0 'Inicializar el contador.
 '   Timer1.Interval = 0    'Detener el cronometro
 '   lblCronometro.Caption = ""  'Limpiar la etiqueta
 '   Timer1.Interval = 1    'Iniciar el cronometro
    
'    Command2.Visible = True
    
    Dim vVal As Double
    
    vVal = Shell(App.Path & "\PadronArba.exe", 1)


End Sub

Private Sub cmdIniciar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub

Private Sub cmdSalir_Click()
    
    End
    
End Sub

Private Sub Command2_Click()
  
  
  'para leer:
  'declaras una variable en donde pones la ruta del archivo, por ejemplo:

    Dim strRuta As String
    Dim vNombrePadron As String
    
   ' On Error GoTo CapturaErrores

  'y declaras otra variable en donde pones la linea que estas leyendo

    vRta = MsgBox("Va a Borrar los Datos en la tabla PADRON" & Chr(10) & "¿Resguarda la BBDD Original?", vbYesNo, "RESGUARDAD BASE DE DATOS")
    
    'FormCronometro.Show
    
    If vRta = 6 Then
        SourceDB = App.Path & "\DB_SPC_SI.mdb"
        DestDB = App.Path & "\DB_SPC_SI_Hasta_" & Format(Date, "DD-MM-YYYY") & ".mdb"
        
        FileCopy SourceDB, DestDB  ' Use copy of database; preserve original.
        
    End If
    
    Label1.Visible = True
    Label1.Caption = "Importando Registros - Esto Puede Tardar unos Minutos..."

    I = 0 'Inicializar el contador.
'    Timer1.Interval = 0    'Detener el cronometro
    lblCronometro.Caption = ""  'Limpiar la etiqueta
 '   Timer1.Interval = 1    'Iniciar el cronometro
    
    
    Set bbdd = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    bbdd.Execute "DELETE * FROM Padron"

    Set TPadron = bbdd.OpenRecordset("Padron", dbOpenTable)
    
    
    Dim strLinea As String

    vNombrePadron = "PadronRGSPer" & Format(Date, "mm") & Format(Date, "yyyy") & ".txt"
    
        
    strRuta = "c:\TEMP\" & vNombrePadron
    
    'MsgBox (strRuta)
    Label2.Visible = True
    Label2.Caption = "Importando Registros - Esto Puede Tardar unos Minutos..."
    'Label2.BackColor = vbblink
        
    PgB1.Max = Lineas(strRuta)
    
    'MsgBox (PgB1.Max)

 '"con esto abres el archovo"
    Open strRuta For Input As #1


     '"y con esto lees linea por linea"
 '   Line Input #1, strLinea
 
 '"y finalmente con esto cierras el archivo"
  '  Close #1

  '"y esto es para que leas el archivo de principio a fin, linea por linea"

    PgB1.Value = 0
    cont = 0
    
    DG1.Rows = 2
    DG1.Cols = 10
    
    DG1.Row = 0
    DG1.Col = 0
    DG1.Text = "F1"
    DG1.Col = 1
    DG1.Text = "F2"
    DG1.Col = 2
    DG1.Text = "F3"
    DG1.Col = 3
    DG1.Text = "F4"
    DG1.Col = 4
    DG1.Text = "F5"
    DG1.Col = 5
    DG1.Text = "F6"
    DG1.Col = 6
    DG1.Text = "F7"
    DG1.Col = 7
    DG1.Text = "F8"
    DG1.Col = 8
    DG1.Text = "F9"
    DG1.Col = 9
    DG1.Text = "F10"
    
    DG1.Row = 1
    
    Do While Not EOF(1)
        Line Input #1, strLinea
      
        ' enlaza el recordset con el DataGrid
        With TPadron
            .AddNew
             '   DG1.Col = 0
                !Regimen = Mid(strLinea, 1, 1)
             '   DG1.Text = !Regimen
             '   DG1.Col = 1
                !FPub = Mid(strLinea, 3, 8)
             '   DG1.Text = !FPub
             '   DG1.Col = 2
                !FVigDde = Mid(strLinea, 12, 8)
              '  DG1.Text = !FVigDde
              '  DG1.Col = 3
                !FVigHta = Mid(strLinea, 21, 8)
              '  DG1.Text = !FVigHta
              '  DG1.Col = 4
                !CUIT = Mid(strLinea, 30, 11)
              '  DG1.Text = !CUIT
              '  DG1.Col = 5
                !TipoContrInsc = Mid(strLinea, 42, 1)
              '  DG1.Text = !TipoContrInsc
              '  DG1.Col = 6
                !MarcaAltaBajaSujeto = Mid(strLinea, 44, 1)
              '  DG1.Text = !MarcaAltaBajaSujeto
              '  DG1.Col = 7
                !MCbioAlicuota = Mid(strLinea, 46, 1)
              '  DG1.Text = !MCbioAlicuota
              '  DG1.Col = 8
                !AlicuotaPercepcion = Format(Mid(strLinea, 48, 4), "Standard")
              '  DG1.Text = !AlicuotaRetencion
              '  DG1.Col = 9
                !NroGrupoPercepcion = Mid(strLinea, 53, 2)
              '  DG1.Text = !NroGrupoRetencion
            .Update
        End With
      cont = cont + 1
      PgB1.Value = cont
      Label1 = "Registro Nro: " & cont & Chr(10) & CLng((PgB1.Value * 100) / PgB1.Max) & " %"
      
   '   FormImportTxt.Refresh
      
      'Label2.Caption = Format(Now, "HH:MM:SS")
      'DG1.Rows = DG1.Rows + 1
      'DG1.Row = DG1.Row + 1
    Loop
    
    Close #1
    
'    DG1.ClearFields
'    Set DG1.DataSource = TPadron
        
    A = MsgBox("PROCESO DE ACTUALIZACIÓN REALIZADO CON ÉXITO" & Chr(10) & cont & " REGISTROS CARGADOS", vbOKOnly, "INFO DEL SISTEMA")
        
  '  FormImportTxt.Refresh
    
    'Timer1.Interval = 0
    
CapturaErrores:
    Select Case Err
        Case 53
            A = MsgBox("ERROR, No se encuentra el Archivo " & App.Path & "\" & vNombrePadron, vbCritical, "INFO DEL SISTEMA")
            Exit Sub
        'Case 94
        '    a = MsgBox("ERROR, No se encuentra el Archivo " & App.Path & "\Padron.txt", , "INFO DEL SISTEMA")
    End Select
    
End Sub


Private Sub Form_Load()
    Label2.Visible = False
    Label2.Caption = ""
  
'    Dim cn As New ADODB.Connection
'    Dim rs As New ADODB.Recordset
  
  ' cursor del lado del cliente
 '   cn.CursorLocation = adUseClient
      
    ' abre el archivo de texto datos.txt ubicado en el app.path
 '   cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;" _
  '                   & "Data Source=" & App.Path & ";" _
  '                  & "Extended Properties='text;HDR=NO;FMT=DELIMITED'"
      
    ' Ejecuta la consulta sql para llenar el recordset
      
  '  rs.Open "Select * from Padron.txt ", cn, adOpenStatic
      
    ' enlaza el recordset con el DataGrid
  '  Set DG1.DataSource = rs
    
End Sub

Private Function Lineas(ByVal strRuta As String) As Long
    Dim arrLineas() As String
    arrLineas = Split(LeerArchivo(strRuta), vbNewLine)
    Lineas = UBound(arrLineas) - LBound(arrLineas) + 1
    Erase arrLineas
End Function

Private Function LeerArchivo(ByVal strRuta As String) As String
    Dim f As Integer
    f = FreeFile
    Open strRuta For Input As #f
    LeerArchivo = Input(LOF(f), #f)
    Close #f
End Function

Private Sub Form_Resize()
    
    Move (Screen.Width - Width) \ 29, (Screen.Height - Height) \ 29 'Centra el formulario completamente
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    End
    
End Sub

Private Sub Timer1_Timer()
    
    I = I + 1
    Tiempo = Format(Int(I / 36000) Mod 24, "00") & ":" & _
             Format(Int(I / 600) Mod 60, "00") & ":" & _
             Format(Int(I / 10) Mod 60, "00") & ":" & _
             Format(I Mod 10, "00")
    lblCronometro.Caption = Tiempo

    If Label2.Visible = True Then
        Label2.Visible = False
    Else
        Label2.Visible = True
    End If
    

End Sub
