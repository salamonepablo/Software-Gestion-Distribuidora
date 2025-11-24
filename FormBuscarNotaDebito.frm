VERSION 5.00
Begin VB.Form FormBuscarNotaDebito 
   Caption         =   "Buscar Nota de Debito"
   ClientHeight    =   2250
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2250
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox TextNumeroFactura 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextA 
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TextTipo 
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox ComboTipo 
         Height          =   315
         ItemData        =   "FormBuscarNotaDebito.frx":0000
         Left            =   2400
         List            =   "FormBuscarNotaDebito.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nota Debito:"
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
         Index           =   0
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   1320
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   555
      End
   End
End
Attribute VB_Name = "FormBuscarNotaDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rstFacturaC As DAO.Recordset



Private Sub ComboTipo_Click()

    TextTipo.Text = ComboTipo.Text
    TextNumeroFactura.Enabled = True
    TextNumeroFactura.SetFocus
    
End Sub

Private Sub Form_Load()

    FormBuscarNotaDebito.Height = 2310
    FormBuscarNotaDebito.Width = 4800
    FormBuscarNotaDebito.Top = 1500
    FormBuscarNotaDebito.Left = 1500
    
    ComboTipo.AddItem ("A")
    ComboTipo.AddItem ("B")

End Sub



Private Sub TextNumeroFactura_KeyPress(KeyAscii As Integer)

    
        TextA.Text = 1
        
        ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db1 = DBEngine.OpenDatabase(ruta)
        
        Set rstDebC = db1.OpenRecordset("NotaDebitoC", dbOpenTable)
'
        If KeyAscii = 13 Then

                rstDebC.Index = "PrimaryKey"
        
                rstDebC.Seek "=", TextTipo, Str(TextNumeroFactura.Text)

             If rstDebC.NoMatch Then
                mensaje = MsgBox("Nota Debito Inexistente", vbCritical, "Final de la busqueda")
                TextNumeroFactura.Text = ""
                TextNumeroFactura.SetFocus
             Else
      
                TextCodigoCliente.Text = rstDebC.Fields!CodCliente
                TextTipo.Text = rstDebC.Fields!TipoDebito
                TextNumeroFactura.Enabled = False
                FormVerNotaDebito.Show
            
        End If
    End If
     
End Sub

