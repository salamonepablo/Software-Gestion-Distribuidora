VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FormDomiciliosRemito 
   Caption         =   "Domicilos Entrega"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11955
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   11955
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   3120
      Width           =   11655
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   5280
         TabIndex        =   2
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Direccion"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   2295
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   11175
         _ExtentX        =   19711
         _ExtentY        =   4048
         _Version        =   393216
         Cols            =   6
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   8040
         TabIndex        =   4
         Top             =   240
         Visible         =   0   'False
         Width           =   2055
      End
   End
End
Attribute VB_Name = "FormDomiciliosRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstDomiciliosClientes As DAO.Recordset

Private Sub BotonSalir_Click()

    Unload FormDomiciliosRemito

End Sub

Private Sub Form_Load()

    FormDomiciliosRemito.Height = 4905
    FormDomiciliosRemito.Width = 11955
    FormDomiciliosRemito.Top = 3000
    FormDomiciliosRemito.Left = 3000

       
    TextCodigoCliente.Text = FormRemito.TextCodigoCliente
    
End Sub

Private Sub titulos()

    MSHFlexGrid1.Row = 0
    
    MSHFlexGrid1.Col = 0
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.ColAlignment(0) = flexAlignCenterCenter
    MSHFlexGrid1.Text = "Item"
    MSHFlexGrid1.ColWidth(0) = 0
    
        
    MSHFlexGrid1.Col = 1
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.ColAlignment(1) = flexAlignCenterCenter
    MSHFlexGrid1.Text = "Domicilio"
    MSHFlexGrid1.ColWidth(1) = 4000
    
    MSHFlexGrid1.Col = 2
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.ColAlignment(2) = flexAlignCenterCenter
    MSHFlexGrid1.Text = "Localidad"
    MSHFlexGrid1.ColWidth(2) = 2000
    
    MSHFlexGrid1.Col = 3
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.ColAlignment(3) = flexAlignCenterCenter
    MSHFlexGrid1.Text = "Cod.Pos"
    MSHFlexGrid1.ColWidth(3) = 800
    
    MSHFlexGrid1.Col = 4
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.ColAlignment(4) = flexAlignCenterCenter
    MSHFlexGrid1.Text = "Provincia"
    MSHFlexGrid1.ColWidth(4) = 1900
    
    MSHFlexGrid1.Col = 5
    MSHFlexGrid1.CellFontBold = True
    MSHFlexGrid1.ColAlignment(5) = flexAlignCenterCenter
    MSHFlexGrid1.Text = "Pais"
    MSHFlexGrid1.ColWidth(5) = 1900
    
   
 
 End Sub
 
 Private Sub buscodirecciones()
 
 MSHFlexGrid1.Clear
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstDomiciliosClientes = db.OpenRecordset("DomiciliosClientes", dbOpenDynaset)
    
    
    MSHFlexGrid1.Rows = 2
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Visible = True
    
    Call titulos
    
    
    CodigoClie = Val(TextCodigoCliente.Text)
    
    rstDomiciliosClientes.FindFirst "IDCliente= " + Str(CodigoClie)
    'facturacancelada = rstDomiciliosClientes.Fields!Cancelada
    codigoclientedetalle = rstDomiciliosClientes.Fields!IDCliente
    
    If rstDomiciliosClientes.Fields!IDCliente <> Val(TextCodigoCliente.Text) Then
        MSHFlexGrid1.Visible = False
        mensaje = MsgBox("No Existen Domicilios", vbCritical, "Final de la busqueda")
        TextCodigoCliente.Text = ""
        'Call blanco
        'TextCodigoCliente.SetFocus
    End If
    
    If codigoclientedetalle = CodigoClie Then
        MSHFlexGrid1.Rows = 2
        MSHFlexGrid1.Clear
        MSHFlexGrid1.Visible = True
    
       
    Else
        MSHFlexGrid1.Visible = False
    End If
    Call titulos
    linea2 = 1
    Do While Not rstDomiciliosClientes.NoMatch
            MSHFlexGrid1.AddItem " "
            MSHFlexGrid1.Row = linea2
       
            MSHFlexGrid1.Col = 0
            MSHFlexGrid1.Text = rstDomiciliosClientes.Fields!item
            MSHFlexGrid1.Col = 1
            MSHFlexGrid1.Text = rstDomiciliosClientes.Fields!Domicilio
            MSHFlexGrid1.Col = 2
            MSHFlexGrid1.Text = rstDomiciliosClientes.Fields!Localidad
            MSHFlexGrid1.Col = 3
            MSHFlexGrid1.Text = rstDomiciliosClientes.Fields!CP
            MSHFlexGrid1.Col = 4
            MSHFlexGrid1.Text = rstDomiciliosClientes.Fields!Prov
            MSHFlexGrid1.Col = 5
            MSHFlexGrid1.Text = rstDomiciliosClientes.Fields!Pais
            'facturacancelada = rstDomiciliosClientes.Fields!Cancelada
            'If facturacancelada = True Then
            '    MSHFlexGrid1.Col = 4
            '    MSHFlexGrid1.Text = "SI"
            'Else
            '    MSHFlexGrid1.Col = 4
            '    MSHFlexGrid1.Text = "NO"
            'End If
            linea2 = linea2 + 1
                
            rstDomiciliosClientes.FindNext "IDCliente= " + Str(CodigoClie)
    Loop
    
    

 
 End Sub

Private Sub MSHFlexGrid1_Click()

    MSHFlexGrid1.Col = 0
    FormRemito.TextItemDomicilio.Text = MSHFlexGrid1.Text

    MSHFlexGrid1.Col = 1
    FormRemito.TextDireccion.Text = MSHFlexGrid1.Text
    
    MSHFlexGrid1.Col = 2
    FormRemito.TextLocalidad.Text = MSHFlexGrid1.Text
    
    MSHFlexGrid1.Col = 3
    FormRemito.TextCodigoPostal = MSHFlexGrid1.Text
    
    MSHFlexGrid1.Col = 4
    FormRemito.TextProvincia.Text = MSHFlexGrid1.Text
    
    'MSHFlexGrid1.Col = 5
    'FormRemito.TextProvincia.Text = MSHFlexGrid1.Text
    
    Unload FormDomiciliosRemito
    
    FormRemito.BotonNueva.SetFocus
    

End Sub

Private Sub TextCodigoCliente_Change()

    Call buscodirecciones
    
End Sub
