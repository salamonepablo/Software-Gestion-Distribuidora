VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FormProveedores 
   Caption         =   "Proveedores"
   ClientHeight    =   4035
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   7380
   Begin VB.Frame Frame1 
      Caption         =   "Datos Proveedores"
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   1935
         Left            =   120
         TabIndex        =   6
         Top             =   1200
         Visible         =   0   'False
         Width           =   6855
         _ExtentX        =   12091
         _ExtentY        =   3413
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   3
      End
      Begin VB.TextBox TextCodigoProveedor 
         Height          =   285
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextCUIT 
         Height          =   285
         Left            =   1920
         TabIndex        =   8
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TextNombreProveedor 
         Height          =   285
         Left            =   1920
         TabIndex        =   7
         Top             =   1080
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Codigo Proveedor:"
         Height          =   255
         Left            =   360
         TabIndex        =   12
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre Proveedor:"
         Height          =   255
         Left            =   360
         TabIndex        =   11
         Top             =   1080
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "CUIT:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1560
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   2520
      Width           =   7095
      Begin VB.CommandButton ComEliminar 
         Caption         =   "&Eliminar"
         Height          =   735
         Left            =   2640
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton ComGuardar 
         Caption         =   "&Guardar"
         Height          =   735
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   5760
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Agregar"
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FormProveedores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
    Dim db As DAO.Database
    Dim rst As DAO.Recordset
    
   

Private Sub CmdAgregar_Click()

    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
    Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)
    
    Call blanco
    
    rst.MoveLast
    CodigoCliente = rst.Fields!CodProv + 1
    TextCodigoProveedor = CodigoCliente
   
    TextNombreProveedor.SetFocus

End Sub

Private Sub CmdExit_Click()
     Unload FormProveedores
End Sub

Private Sub ComAgregar_Click()

    
    
End Sub

Private Sub ComEliminar_Click()

    Set db = DBEngine.OpenDatabase("C:\QuilplacVB 20-10-2013\Padron.mdb")
    Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)

    If TextNombreProveedor.Text = "" Then
        respuesta = MsgBox("Ingrese los Datos a Borrar", vbCritical, "")
        TextNombreProveedor.SetFocus
    Else
        respuesta = MsgBox("Esta Seguro que Desea Eliminar el Proveedor?", vbYesNo, "Borrar el Proveedor")
        If respuesta = vbYes Then
            CodigoProveedor = Val(TextCodigoProveedor.Text)
            rst.FindFirst "CodProv= " + Str(CodigoProveedor)
            rst.Delete
            Call blanco
         End If
         TextCodigoProveedor.SetFocus
    End If


End Sub

Private Sub ComGuardar_Click()

    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
    Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)
 
    If TextCodigoProveedor.Text = "" Or TextNombreProveedor.Text = "" Then
        respuesta = MsgBox("Complete los Datos Faltantes", vbCritical, " ")
        TextNombreProveedor.SetFocus
    Else
        CodigoProveedor = Val(TextCodigoProveedor.Text)
        rst.FindFirst "CodProv= " + Str(CodigoProveedor)
        If rst.Fields!CodProv <> Val(TextCodigoProveedor.Text) Then
           rst.AddNew
           Call muevo
           Call blanco
           TextCodigoProveedor.Enabled = True
        Else
            CodigoProveedor = Val(TextCodigoProveedor.Text)
            rst.FindFirst "CodProv= " + Str(CodigoProveedor)
            rst.Edit
            Call muevo
            Call blanco
            TextCodigoProveedor.Enabled = True
            TextCodigoProveedor.SetFocus
        End If
    End If

End Sub
Private Sub muevo()

    rst.Fields!CodProv = Val(TextCodigoProveedor.Text)
    rst.Fields!NombreProv = TextNombreProveedor.Text
    rst.Fields!Cuit = TextCUIT.Text
    rst.Update

End Sub



Private Sub ComSalir_Click()
   
End Sub



Private Sub TextCodigoProveedor_KeyPress(KeyAscii As Integer)

'    Dim db As DAO.Database
'    Dim rst As DAO.Recordset
    
    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
    Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)
 
    
' Base Datos Padron
' Tabla Proveedores
' Campos
    ' CodProv Numerico
    ' NombreProv Texto 50
    ' Cuit Texto 11
    ' Estado Texto 2
    

    If KeyAscii = 13 Then
    If TextCodigoProveedor.Text = "" Then
       TextNombreProveedor.SetFocus
    Else
        CodigoProveedor = Val(TextCodigoProveedor.Text)
      
        rst.FindFirst "CodProv= " + Str(CodigoProveedor)
        If rst.Fields!CodProv <> Val(TextCodigoProveedor.Text) Then
           mensaje = MsgBox("Proveedor Inexistente", vbCritical, "Final de la busqueda")
           TextCodigoProveedor.Text = ""
           Call blanco
           TextCodigoProveedor.SetFocus
        Else
           TextCodigoProveedor.Text = rst.Fields!CodProv
           TextNombreProveedor.Text = rst.Fields!NombreProv
           TextCUIT.Text = rst.Fields!Cuit
        End If
          
    End If
    End If

End Sub

Private Sub blanco()

    TextCodigoProveedor.Text = ""
    TextNombreProveedor.Text = ""
    TextCUIT.Text = ""

End Sub

    

Private Sub Form_Load()
    FormProveedores.Height = 4605
    FormProveedores.Width = 7620

    

   ' Set Padron = OpenDatabase("C:\QuilplacVB 20-10-2013\Padron.mdb")
   ' Set Provs = Padron.OpenRecordset("Proveedores")
    
    Call blanco
   

End Sub


Private Sub TextNombreProveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    Call busco
End If

End Sub

Private Sub busco()

    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
    Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)
    
    MSHFlexGrid1.Rows = 2
    MSHFlexGrid1.Clear
    MSHFlexGrid1.Visible = True
    
    Call titulos
    
    Dim busca1 As String, busca2 As String
    busca1 = RTrim(LTrim(TextNombreProveedor.Text))
    busca2 = busca1 + "z"
    
    rst.FindFirst "NombreProv >= '" & busca1 & "' and NombreProv <= '" & busca2 & "'"
    
    If rst.NoMatch Then
       MSHFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Proveedores", vbCritical, "Final de la busqueda")
       TextNombreProveedor.Text = ""
       Call blanco
       TextNombreProveedor.SetFocus
    End If
     
    linea2 = 1
    Do While Not rst.NoMatch
        MSHFlexGrid1.AddItem " "
        MSHFlexGrid1.Row = linea2
        If rst.Fields!Estado = "si" Then
            MSHFlexGrid1.Col = 0
            MSHFlexGrid1.Text = rst.Fields!CodProv
            MSHFlexGrid1.Col = 1
            MSHFlexGrid1.Text = rst.Fields!NombreProv
            MSHFlexGrid1.Col = 2
            MSHFlexGrid1.Text = rst.Fields!Cuit
            linea2 = linea2 + 1
       End If
       rst.FindNext "NombreProv >= '" & busca1 & "' and NombreProv <= '" & busca2 & "'"
       
    Loop

End Sub

Private Sub titulos()

    MSHFlexGrid1.Row = 0
    
    MSHFlexGrid1.Col = 0
    MSHFlexGrid1.Text = "Codigo"
    MSHFlexGrid1.ColWidth(0) = 900
    
    MSHFlexGrid1.Col = 1
    MSHFlexGrid1.Text = "Apellido y Nombre"
    MSHFlexGrid1.ColWidth(1) = 4700
    
    MSHFlexGrid1.Col = 2
    MSHFlexGrid1.Text = "CUIT"
    MSHFlexGrid1.ColWidth(2) = 1200

End Sub

Private Sub MSHFlexGrid1_DblClick()

    MSHFlexGrid1.Col = 0
    TextCodigoProveedor.Text = MSHFlexGrid1.Text
    
    MSHFlexGrid1.Col = 1
    TextNombreProveedor.Text = MSHFlexGrid1.Text
    
    MSHFlexGrid1.Col = 2
    TextCUIT.Text = MSHFlexGrid1.Text
    
    MSHFlexGrid1.Visible = False

End Sub

