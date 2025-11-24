VERSION 5.00
Begin VB.Form FormBuscarRemito 
   Caption         =   "Buscar Remito"
   ClientHeight    =   1950
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1950
   ScaleWidth      =   4830
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
      Begin VB.ComboBox cmbSucursal 
         Height          =   315
         Left            =   1920
         TabIndex        =   5
         Text            =   "Combo1"
         Top             =   360
         Width           =   1935
      End
      Begin VB.TextBox TextA 
         Height          =   285
         Left            =   1320
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   0
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   0
         TabIndex        =   3
         Top             =   0
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
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Sucursal:"
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
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   975
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
         Top             =   960
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

    Dim tSucursales
    
    FormBuscarFactura.Height = 2310
    FormBuscarFactura.Width = 4800
    FormBuscarFactura.Top = 1500
    FormBuscarFactura.Left = 1500
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db1 = DBEngine.OpenDatabase(ruta)
        
    Set tSucursales = db1.OpenRecordset("Sucursales", dbOpenTable)
    
    tSucursales.MoveFirst
    
    Do
        cmbSucursal.AddItem tSucursales!IdSucursal & " - " & tSucursales!NombreSucursal
        tSucursales.MoveNext
    
    Loop Until tSucursales.EOF
    
    cmbSucursal.ListIndex = 1

End Sub



Private Sub TextNumeroFactura_KeyPress(KeyAscii As Integer)

        Dim rstRemitoC, rstRemitoD
        'Dim IdSucursal As Long
        
    
        TextA.text = 1
        
        numDoc = Val(TextNumeroFactura.text)
    
        ruta = App.Path & "\DB_SPC_SI.mdb"
        
        Set db1 = DBEngine.OpenDatabase(ruta)
            
        'Set rstRemitoC = db1.OpenRecordset("RemitoC", dbOpenDynaset)
        Set rstRemitoC = db1.OpenRecordset("RemitoC", dbOpenTable)
        
        rstRemitoC.Index = "PrimaryKey"
        
        IdSucursal = CLng(Left(cmbSucursal.text, 1))
    
        If KeyAscii = 13 Then
        
             'Dim busca1 As String, busca2 As String
             Dim busca2 As String
             Dim busca1 As Long
             'busca1 = RTrim(LTrim(TextNumeroFactura.text))
             busca1 = CLng(RTrim(LTrim(TextNumeroFactura.text)))
           '  busca2 = busca1 + "z"
            
             'rstRemitoC.FindFirst "NroRemito >= '" & busca1 & "' and NroRemito <= '" & busca2 & "'"
             'rstRemitoC.FindFirst "NroRemito >= " & busca1 & " and NroRemito <= " & busca2 & ""
             'rstRemitoC.FindFirst "NroRemito= " + Str(numDoc)
             
             rstRemitoC.Seek "=", IdSucursal, busca1
             
             If rstRemitoC.NoMatch Then
             
             'If rstRemitoC.Fields!NroRemito <> Val(TextNumeroFactura.text) Then
                mensaje = MsgBox("Remito Inexistente", vbCritical, "Final de la busqueda")
                TextNumeroFactura.text = ""
                TextNumeroFactura.SetFocus
              Else

'                rstRemitoC.FindFirst "NroRemito= " + Str(numDoc)

'                 Dim busca1 As String, busca2 As String

'************** modificaciones agregadas con el numero de sucursal 2025-06 ****************************************
'                 busca1 = RTrim(LTrim(TextNumeroFactura.text))
'                 busca2 = busca1 + "z"
            
'                rstRemitoC.FindFirst "NroRemito >= '" & busca1 & "' and NroRemito <= '" & busca2 & "'"
                
                TextCodigoCliente.text = rstRemitoC.Fields!CodCliente
            
                FormVerRemito.Show
            
        End If
    End If
End Sub
