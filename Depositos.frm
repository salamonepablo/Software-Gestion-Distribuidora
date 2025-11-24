VERSION 5.00
Begin VB.Form FormDepositos 
   Caption         =   "Form1"
   ClientHeight    =   3855
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   ScaleHeight     =   3855
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Depositos_frame 
      Caption         =   " Depositos "
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   12735
      Begin VB.TextBox txt_dire 
         Height          =   495
         Left            =   1200
         TabIndex        =   3
         Top             =   1680
         Width           =   2535
      End
      Begin VB.TextBox txt_nombre 
         Height          =   495
         Left            =   1200
         TabIndex        =   2
         Top             =   1080
         Width           =   2535
      End
      Begin VB.TextBox txt_depo 
         Height          =   495
         Left            =   1200
         TabIndex        =   1
         Top             =   480
         Width           =   2535
      End
      Begin VB.Frame Frame5 
         Caption         =   "Acciones"
         Height          =   1215
         Left            =   240
         TabIndex        =   12
         Top             =   2280
         Width           =   12255
         Begin VB.CommandButton btnPrimero 
            Caption         =   "|<"
            Height          =   615
            Left            =   240
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAtras 
            Caption         =   "<<"
            Height          =   615
            Left            =   1440
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAdelante 
            Caption         =   ">>"
            Height          =   615
            Left            =   2640
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnUltimo 
            Caption         =   ">|"
            Height          =   615
            Left            =   3840
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnGrabar 
            Caption         =   "&Grabar"
            Height          =   615
            Left            =   6240
            TabIndex        =   5
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnModificar 
            Caption         =   "&Modificar"
            Height          =   615
            Left            =   7440
            TabIndex        =   6
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnBuscar 
            Caption         =   "&Buscar"
            Height          =   615
            Left            =   8640
            TabIndex        =   7
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnEliminar 
            Caption         =   "&Eliminar"
            Height          =   615
            Left            =   9840
            TabIndex        =   8
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnSalir 
            Caption         =   "&Salir"
            Height          =   615
            Left            =   11040
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnLimpiar 
            Caption         =   "&Limpiar"
            Height          =   615
            Left            =   5040
            TabIndex        =   4
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Label Direccion_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Direccion:"
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
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   1800
         Width           =   885
      End
      Begin VB.Label lbl_nombre 
         AutoSize        =   -1  'True
         Caption         =   "Nombre:"
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
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Deposito_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Depósito:"
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
         TabIndex        =   9
         Top             =   600
         Width           =   810
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nombre:"
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
      Index           =   1
      Left            =   6360
      TabIndex        =   18
      Top             =   1440
      Width           =   720
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Direccion"
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
      Index           =   1
      Left            =   360
      TabIndex        =   17
      Top             =   1800
      Width           =   825
   End
End
Attribute VB_Name = "FormDepositos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private KeyRetroceso As Boolean

Private Sub btnAdelante_Click()

On Error GoTo CapturaErrores
    
    If Not tDepositos.EOF Then
        tDepositos.MoveNext
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
                
    Else
        MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
        tDepositos.MoveLast
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            'MsgBox "Ultimo Registro", vbDefaultButton1, "INFO DEL SISTEMA"
            MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
            tDepositos.MoveLast
            Call Mostrar
            Resume Next
    End Select


End Sub

Private Sub btnAtras_Click()

On Error GoTo CapturaErrores
    
    If Not tDepositos.BOF Then
        tDepositos.MovePrevious
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
        
        
    Else
        MsgBox "Primer Registro", vbInformation, "INFO DEL SISTEMA"
        tDepositos.MoveFirst
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "Primer Registro", vbInformation + vbOKOnly, "INFO DEL SISTEMA"
            'MsgBox "No hay registros !!!", vbDefaultButton1, "INFO DEL SISTEMA"
            tDepositos.MoveFirst
            Call Mostrar
            Resume Next
    End Select

End Sub

Private Sub btnBuscar_Click()

    If vFlagBuscar = 0 Then
        vFlagBuscar = 1
        txt_depo.Enabled = True
        txt_dire.Text = ""
        txt_depo.SetFocus
     Else
        
        If txt_depo.Text <> "" Then
            Campo = "IDDeposito= "
      '       Valor = "'" + txt_depo.Text + "*'"
            Valor = txt_depo.Text


         Else
            If txt_nombre.Text <> "" Then
                Campo = "Nombre Like "
                Valor = "'" + txt_nombre.Text + "*'"
             Else
                If txt_dire.Text <> "" Then
                    Campo = "Direccion Like "
                    Valor = "'" + txt_dire.Text + "*'"
                End If
            End If
        End If
        
        'vSQL = "SELECT IDProv, Descripcion FROM Provincias Where IDPais=" & tPaises!IDPais & " ORDER BY Descripcion"
        vSQL = "SELECT * FROM Depositos WHERE " & Campo & Valor & " ORDER BY IDDeposito"
        
        'MsgBox (vsql)
        
       ' Set tDepositos = BaseSPC.OpenRecordset(vSQL)
    
        If Not tDepositos.NoMatch Then
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
        Else
            MsgBox "No se encuentran registros", vbCritical, "ERROR"
        End If
        
        vFlagBuscar = 0
        
    End If


End Sub

Private Sub btnEliminar_Click()

b = MsgBox("¿ Seguro Desea Eliminar Deposito ?", vbQuestion + vbOKCancel, "Eliminar Deposito")
    
    With tDepositos
        .Delete
    End With
    
    Call EnabledTextBox(FormDepositos)
    Call LimpiarPantalla

End Sub

Private Sub btnGrabar_Click()

    A = MsgBox("¿ Seguro Genera Nuevo Deposito ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
    tDepositos.Index = "PrimaryKey"
    
     With tDepositos
        .AddNew
        'Datos Depositos -------------------------------------------------------------------
            !IDDEPOSITO = txt_depo.Text
            !Descripcion = txt_nombre.Text
            !direccion = txt_dire.Text
            
           ' Call LimpiarPantalla
                   
        .Update
     End With
     
   ' tDepositos.Index = "PrimaryKey"
   ' tDepositos.Seek "=", "tEmpleados"
   
   '  tDepositos.Edit
   '  tUltimosNumeros!UltimoNumero = IDDeposito.Text
   ' .Update
    
'    txtLegEmple.Text = tUltimosNumeros!UltimoNumero + 1
    Call LimpiarPantalla
    txt_depo.SetFocus

CapturaErrores:

    Select Case Err
        Case 3021
            Resume Next
    End Select


End Sub

Private Sub btnLimpiar_Click()
    Call LimpiarPantalla
End Sub

Private Sub btnSalir_Click()
    Unload Me
End Sub

Private Sub btnUltimo_Click()

On Error GoTo CapturaErrores
    
    If Not tDepositos.EOF Then
        tDepositos.MoveLast
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
     Else
        vSQL = "SELECT * FROM Depositos ORDER BY IDDepositos"
        'MsgBox (vsql)
        Set tDepositos = BaseSPC.OpenRecordset(vSQL)
        
        If Not tDepositos.EOF Then
            tDepositos.Last
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay más registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select



End Sub

Private Sub Form_Load()
Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
'Tabla Depositos
            Set tDepositos = BaseSPC.OpenRecordset("Depositos", dbOpenTable)
        
End Sub


Private Sub LimpiarPantalla()

    txt_depo.Text = ""
    txt_nombre.Text = ""
    txt_dire.Text = ""

    btnGrabar.Enabled = False
    btnEliminar.Enabled = False
    btnModificar.Caption = "&Modificar"
    
    'txt_depo.Text = tUltimosNumeros!UltimoNumero + 1
    Call EnabledTextBox(FormDepositos)

    txt_depo.SetFocus
    

End Sub

Private Sub Mostrar()

     With tDepositos
            txt_depo.Text = !IDDEPOSITO
            txt_nombre.Text = !Descripcion
            txt_dire.Text = !direccion
            
           ' txt_depo.Text = "PrimaryKey"
                          
     End With
     
     Call DisabledTextBox(FormDepositos)

End Sub



Private Sub txt_depo_GotFocus()
  txt_depo.SelLength = Len(txt_depo.Text)
End Sub





Private Sub txt_dire_GotFocus()
    txt_dire.SelLength = Len(txt_dire.Text)
End Sub

Private Sub txt_nombre_GotFocus()
    txt_nombre.SelLength = Len(txt_nombre.Text)
End Sub
