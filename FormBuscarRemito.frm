VERSION 5.00
Begin VB.Form FormBuscarRemito 
   Caption         =   "Buscar Remito"
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
      Begin VB.TextBox TextA 
         Height          =   285
         Left            =   2040
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   240
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
         Left            =   1920
         TabIndex        =   1
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Numero Remito:"
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
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1680
      End
   End
End
Attribute VB_Name = "FormBuscarRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim rstRemitoC As DAO.Recordset

Private Sub Form_Load()

    FormBuscarFactura.Height = 2310
    FormBuscarFactura.Width = 4800
    FormBuscarFactura.Top = 1500
    FormBuscarFactura.Left = 1500

End Sub



Private Sub TextNumeroFactura_KeyPress(KeyAscii As Integer)

    
        TextA.Text = 1
        
        numDoc = Val(TextNumeroFactura.Text)
    
        ruta = App.Path & "\DB_SPC_SI.mdb"
        
        Set db1 = DBEngine.OpenDatabase(ruta)
            
        Set rstRemitoC = db1.OpenRecordset("RemitoC", dbOpenDynaset)
    
        If KeyAscii = 13 Then
        
             Dim busca1 As String, busca2 As String
             busca1 = RTrim(LTrim(TextNumeroFactura.Text))
             busca2 = busca1 + "z"
            
             rstRemitoC.FindFirst "NroRemito >= '" & busca1 & "' and NroRemito <= '" & busca2 & "'"
'             rstRemitoC.FindFirst "NroRemito= " + Str(numDoc)
             If rstRemitoC.Fields!NroRemito <> Val(TextNumeroFactura.Text) Then
                mensaje = MsgBox("Remito Inexistente", vbCritical, "Final de la busqueda")
                TextNumeroFactura.Text = ""
                TextNumeroFactura.SetFocus
             Else

'                rstRemitoC.FindFirst "NroRemito= " + Str(numDoc)

'                 Dim busca1 As String, busca2 As String
                 busca1 = RTrim(LTrim(TextNumeroFactura.Text))
                 busca2 = busca1 + "z"
            
                rstRemitoC.FindFirst "NroRemito >= '" & busca1 & "' and NroRemito <= '" & busca2 & "'"
                
                TextCodigoCliente.Text = rstRemitoC.Fields!CodCliente
            
                FormVerRemito.Show
            
        End If
    End If
End Sub
