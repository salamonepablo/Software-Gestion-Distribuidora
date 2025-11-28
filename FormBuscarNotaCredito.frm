VERSION 5.00
Begin VB.Form FormBuscarNotaCredito 
   Caption         =   "Buscar Nota Credito"
   ClientHeight    =   1740
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4560
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.TextBox TextTipo 
         Height          =   285
         Left            =   0
         TabIndex        =   7
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.ComboBox ComboTipo 
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.TextBox TextA 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   600
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   480
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextNumeroFactura 
         Alignment       =   1  'Right Justify
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
         TabIndex        =   1
         Top             =   840
         Width           =   1935
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
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nota de Credito:"
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
         Left            =   240
         TabIndex        =   2
         Top             =   960
         Width           =   1710
      End
   End
End
Attribute VB_Name = "FormBuscarNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rstNotaCreditoC As DAO.Recordset

Private Sub Form_Load()

    FormBuscarNotaCredito.Height = 2310
    FormBuscarNotaCredito.Width = 4800
    FormBuscarNotaCredito.Top = 1500
    FormBuscarNotaCredito.Left = 1500
    
    ComboTipo.AddItem ("A")
    ComboTipo.AddItem ("B")

End Sub


Private Sub ComboTipo_Click()

    ComboTipo.SetFocus
    TextNumeroFactura.Enabled = True
    TextTipo.Text = ComboTipo.Text
    'TextNumeroFactura.SetFocus
    
End Sub

Private Sub TextNumeroFactura_KeyPress(KeyAscii As Integer)

    
        TextA.Text = 1
        
        
        numDoc = Val(TextNumeroFactura.Text)
    
         ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db1 = DBEngine.OpenDatabase(ruta)
        
        Set rstNotaCreditoC = db1.OpenRecordset("NotaCreditoC", dbOpenTable)
    
        If KeyAscii = 13 Then
        
               rstNotaCreditoC.Index = "PrimaryKey"
        
               rstNotaCreditoC.Seek "=", TextTipo, Str(TextNumeroFactura.Text)

             
             If rstNotaCreditoC.NoMatch Then
                mensaje = MsgBox("Nota Credito Inexistente", vbCritical, "Final de la busqueda")
                TextNumeroFactura.Text = ""
                TextNumeroFactura.SetFocus
             Else
                
                TextCodigoCliente.Text = rstNotaCreditoC.Fields!CodCliente
                TextTipo.Text = rstNotaCreditoC.Fields!TipoNotaCredito
                'TextNumeroFactura.Enabled = False
                
'                rstNotaCreditoC.FindFirst "NroNotaCredito= " + Str(numDoc)
'                TextTipo.Text = rstFacC.Fields!TipoNotaCredito
'                TextCodigoCliente.Text = rstNotaCreditoC.Fields!CodCliente
            
                FormVerNotacredito.Show
            
        End If
    End If
End Sub
