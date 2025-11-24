VERSION 5.00
Begin VB.Form FormUltimosNumerosFactura 
   Caption         =   "Numeros de Factura"
   ClientHeight    =   4140
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   5775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      Top             =   2880
      Width           =   5535
      Begin VB.CommandButton BotonGuardar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   750
         Left            =   1800
         TabIndex        =   9
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   2880
         TabIndex        =   8
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5535
      Begin VB.TextBox TextUltimoNumero 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   6
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TextNumeroActual 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   5
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox ComboTipoFactura 
         Height          =   315
         Left            =   2640
         TabIndex        =   1
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Utimo Numero:"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   1320
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero Actual:"
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
         Left            =   1200
         TabIndex        =   3
         Top             =   1920
         Width           =   1320
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Left            =   1200
         TabIndex        =   2
         Top             =   720
         Width           =   450
      End
   End
End
Attribute VB_Name = "FormUltimosNumerosFactura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstUltimosNumeros As DAO.Recordset

Private Sub BotonGuardar_Click()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)

    Dim busco As String
       
    If ComboTipoFactura.Text = "A" Then
        busco = "tFacturaA"
    End If
    
    If ComboTipoFactura = "B" Then
        busco = "tFacturaB"
    End If
    
    rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
    
    rstUltimosNumeros.Edit
    rstUltimosNumeros.Fields!UltimoNumero = TextNumeroActual.Text
    rstUltimosNumeros.Update
    
    Call blanco
    
    BotonGuardar.Enabled = False

End Sub
Private Sub blanco()

    TextUltimoNumero.Text = ""
    TextNumeroActual.Text = ""

End Sub

Private Sub BotonSalir_Click()

    Unload FormUltimosNumerosFactura

End Sub

Private Sub ComboTipoFactura_Click()

    Dim NumeroFactura As Integer
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
    
    Dim busco As String
       
    If ComboTipoFactura.Text = "A" Then
        busco = "tFacturaA"
    End If
    
    If ComboTipoFactura.Text = "B" Then
        busco = "tFacturaB"
    End If
    
    rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
    NumeroFactura = rstUltimosNumeros.Fields!UltimoNumero
    
    If rstUltimosNumeros.NoMatch Then
       FG1.Visible = False
       mensaje = MsgBox("No existen Numeros de Factura", vbCritical, "Final de la busqueda")
    End If
    
    TextUltimoNumero.Text = NumeroFactura

  
End Sub

Private Sub ComboTipoFactura_GotFocus()
    ComboTipoFactura.SelLength = Len(ComboTipoFactura.Text)
End Sub

Private Sub Form_Load()

    FormUltimosNumerosFactura.Height = 4710
    FormUltimosNumerosFactura.Width = 6015
    
End Sub

Private Sub TextNumeroActual_Change()

    If TextNumeroActual.Text <> "" Then
        BotonGuardar.Enabled = True
    End If

End Sub

Private Sub TextNumeroActual_GotFocus()
    TextNumeroActual.SelLength = Len(TextNumeroActual.Text)
End Sub

Private Sub TextUltimoNumero_GotFocus()
    TextUltimoNumero.SelLength = Len(TextUltimoNumero.Text)
End Sub
