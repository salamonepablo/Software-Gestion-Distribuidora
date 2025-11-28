VERSION 5.00
Begin VB.Form FormIngresoFacturas 
   Caption         =   "Ingreso de Facturas de Proveedores"
   ClientHeight    =   4125
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4125
   ScaleWidth      =   10785
   Begin VB.Frame Frame3 
      Height          =   1215
      Left            =   120
      TabIndex        =   23
      Top             =   2760
      Width           =   10575
      Begin VB.CommandButton CmdExit 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   8760
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   1080
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos Factura"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   10575
      Begin VB.TextBox TextPercepcion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   7560
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtIVA 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6000
         TabIndex        =   8
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtSubtotalFactura 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   4440
         TabIndex        =   7
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtTipoFactura 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   3480
         TabIndex        =   6
         Top             =   720
         Width           =   495
      End
      Begin VB.TextBox TxtTotFac 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   9000
         TabIndex        =   10
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtFF 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   1800
         TabIndex        =   5
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TxtNroFac 
         Alignment       =   2  'Center
         Height          =   285
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Percepcion"
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
         Left            =   7440
         TabIndex        =   22
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "IVA"
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
         Left            =   5880
         TabIndex        =   21
         Top             =   480
         Width           =   315
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal Factura"
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
         Left            =   4320
         TabIndex        =   20
         Top             =   480
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Factura"
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
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Total Factura"
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
         Left            =   8880
         TabIndex        =   18
         Top             =   480
         Width           =   1155
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Factura"
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
         Left            =   1680
         TabIndex        =   17
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nro Factura"
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
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Clientes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   6975
      Begin VB.TextBox TxtCUIT 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   5040
         TabIndex        =   14
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox cmbCodProv 
         Height          =   315
         Index           =   0
         Left            =   2040
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   120
         Visible         =   0   'False
         Width           =   4815
      End
      Begin VB.TextBox TxtProvName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2040
         TabIndex        =   11
         Top             =   120
         Visible         =   0   'False
         Width           =   4215
      End
      Begin VB.ComboBox cmbCodProv 
         Height          =   315
         Index           =   1
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   600
         Width           =   4815
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Cod. Proveedor"
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
         Left            =   600
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Nombre / Razón Social"
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
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1995
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "CUIT:"
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
         Left            =   5040
         TabIndex        =   2
         Top             =   360
         Width           =   510
      End
   End
End
Attribute VB_Name = "FormIngresoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdExit_Click()
     Unload Me
End Sub

Private Sub CmdSave_Click()
    
    ruta = App.Path & "\Padron.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstFacturaCProv = db.OpenRecordset("FacturaCProv", dbOpenDynaset)
    
    '*** Busco Factutra Existente
        
      
        
        ruta = App.Path & "\Padron.mdb"
    
        Set db1 = DBEngine.OpenDatabase(ruta)
        
        Set rstFacC = db1.OpenRecordset("FacturaCProv", dbOpenTable)
        
        rstFacC.Index = "PrimaryKey"
        
        rstFacC.Seek "=", TxtTipoFactura, Str(TxtNroFac.text)

        If Not rstFacC.NoMatch Then
            A = MsgBox("Factura Existente", vbCritical, "INFO DEL SISTEMA")
           
            TxtNroFac.text = num
            TxtNroFac.SetFocus
        Else
        
            rstFacC.Close
            db1.Close
      
            rstFacturaCProv.AddNew
            rstFacturaCProv.Fields!NroFactura = TxtNroFac.text
            rstFacturaCProv.Fields!TipoFactura = UCase(TxtTipoFactura.text)
            rstFacturaCProv.Fields!FechaFactura = TxtFF.text
            rstFacturaCProv.Fields!SubTotalFactura = Format(TxtSubtotalFactura.text, "#,###,###,#0.00")
            rstFacturaCProv.Fields!totalIva = Format(TxtIVA.text, "#,###,###,#0.00")
            rstFacturaCProv.Fields!percepcion = Format(TextPercepcion.text, "#,###,###,#0.00")
            rstFacturaCProv.Fields!TotalFactura = Format(TxtTotFac.text, "#,###,###,#0.00")
            rstFacturaCProv.Fields!CodProv = cmbCodProv(0).text
            rstFacturaCProv.Fields!Paga = "no"
            rstFacturaCProv.Update
        End If
        CmdSave.Enabled = False
        Call blanco
End Sub
Private Sub blanco()

    TxtNroFac.text = ""
    TxtTipoFactura = ""
    TxtFF.text = ""
    TxtSubtotalFactura.text = ""
    TxtIVA.text = ""
    TextPercepcion.text = ""
    TxtTotFac.text = ""
'    cmbCodProv(0).Text = ""
End Sub

Private Sub Form_Load()

    FormIngresoFacturas.Height = 4695
    FormIngresoFacturas.Width = 11025
    FormIngresoFacturas.Top = 300
    FormIngresoFacturas.Left = 300
    
    Set Padron = OpenDatabase(App.Path & "\Padron.mdb")
    Set Provs = Padron.OpenRecordset("Proveedores")



    With Provs
        .MoveFirst
        While Not .EOF
           cmbCodProv(0).AddItem (!CodProv)
           cmbCodProv(1).AddItem (!NombreProv)
           .MoveNext
        Wend
    End With
    

End Sub

Private Sub cmbCodProv_Click(Index As Integer)

    cmbCodProv(0).ListIndex = cmbCodProv(1).ListIndex

End Sub

Private Sub CmbCodProv_KeyPress(Index As Integer, KeyAscii As Integer)

    cmbCodProv(0).ListIndex = cmbCodProv(1).ListIndex
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub

Private Sub CmbCodProv_LostFocus(Index As Integer)

    cmbCodProv(0).ListIndex = cmbCodProv(1).ListIndex
    
    
 If cmbCodProv(0).text <> "" Then
        
     With Provs
        .Index = "Primary"
        Provs.Seek "=", cmbCodProv(0).text
       
        If .NoMatch = False Then
            ' TxtProvName.Text = rst.Fields!NombreProv
             TxtCUIT.text = !CUIT
        End If
     End With
     
  
  Else
        A = MsgBox("Debe Ingresar un Proveedor", vbOKOnly, "ERROR")
 End If

End Sub






Private Sub TextPercepcion_KeyPress(KeyAscii As Integer)

     If KeyAscii = 13 Then
       KeyAscii = 9
       Sendkeys "{TAB}"
    End If

End Sub

Private Sub TxtFF_KeyPress(KeyAscii As Integer)

    
    If KeyAscii = 13 Then
       KeyAscii = 5
       Sendkeys "{TAB}"
    End If

End Sub



Private Sub TxtIva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
       KeyAscii = 9
       Sendkeys "{TAB}"
    End If

End Sub

Private Sub TxtNroFac_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
       KeyAscii = 4
       Sendkeys "{TAB}"
    End If

End Sub




Private Sub TxtSubtotalFactura_KeyPress(KeyAscii As Integer)

    
    If KeyAscii = 13 Then
       KeyAscii = 7
       Sendkeys "{TAB}"
    End If

End Sub

Private Sub TxtTipoFactura_KeyPress(KeyAscii As Integer)

     KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
     
    If KeyAscii = 13 Then
       KeyAscii = 6
       Sendkeys "{TAB}"
    End If

End Sub

Private Sub TxtTotFac_Change()
    
    If TxtTotFac.text <> "" Then
        CmdSave.Enabled = True
    End If
End Sub
