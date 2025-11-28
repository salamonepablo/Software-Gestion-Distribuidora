VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FormProveedores 
   Caption         =   "Proveedores"
   ClientHeight    =   7380
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10965
   Icon            =   "FormProveedores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7380
   ScaleWidth      =   10965
   Begin VB.Frame Frame1 
      Caption         =   "Datos Proveedores"
      Height          =   5535
      Left            =   360
      TabIndex        =   13
      Top             =   120
      Width           =   10695
      Begin VB.TextBox TextProvincia 
         Height          =   285
         Left            =   600
         TabIndex        =   4
         Top             =   2040
         Width           =   1935
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid FG1 
         Height          =   2655
         Left            =   720
         TabIndex        =   11
         Top             =   2640
         Visible         =   0   'False
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   4683
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.TextBox TextNombreProveedor 
         Height          =   285
         Left            =   2040
         TabIndex        =   1
         Top             =   600
         Width           =   8055
      End
      Begin VB.TextBox TextCUIT 
         Height          =   285
         Left            =   600
         TabIndex        =   2
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox TextCodigoProveedor 
         Height          =   285
         Left            =   600
         TabIndex        =   0
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TextLocalidad 
         Height          =   285
         Left            =   3000
         TabIndex        =   5
         Top             =   2040
         Width           =   4935
      End
      Begin VB.TextBox TextCodigoPostal 
         Height          =   285
         Left            =   8520
         TabIndex        =   6
         Top             =   2040
         Width           =   1575
      End
      Begin VB.TextBox TextDireccion 
         Height          =   285
         Left            =   2280
         TabIndex        =   3
         Top             =   1320
         Width           =   7815
      End
      Begin VB.Image Image1 
         Height          =   2610
         Left            =   2400
         Picture         =   "FormProveedores.frx":000C
         Stretch         =   -1  'True
         Top             =   2640
         Width           =   5775
      End
      Begin VB.Label Label7 
         Caption         =   "Provincia:"
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Postal:"
         Height          =   195
         Left            =   8280
         TabIndex        =   19
         Top             =   1800
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Direccion"
         Height          =   195
         Left            =   1920
         TabIndex        =   18
         Top             =   1080
         Width           =   675
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "CUIT"
         Height          =   195
         Left            =   360
         TabIndex        =   17
         Top             =   1080
         Width           =   375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nombre Proveedor"
         Height          =   195
         Left            =   1920
         TabIndex        =   16
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Codigo Proveedor"
         Height          =   195
         Left            =   360
         TabIndex        =   15
         Top             =   360
         Width           =   1275
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Localidad:"
         Height          =   195
         Left            =   2760
         TabIndex        =   14
         Top             =   1800
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   120
      TabIndex        =   12
      Top             =   5880
      Width           =   10695
      Begin VB.CommandButton cmdVerProv 
         Caption         =   "&Ver Proveedores"
         Height          =   735
         Left            =   5880
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton cmdModifica 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Height          =   735
         Left            =   4200
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton ComEliminar 
         Caption         =   "&Eliminar"
         Height          =   735
         Left            =   7920
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton ComGuardar 
         Caption         =   "&Guardar"
         Height          =   735
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdExit 
         Caption         =   "&Salir"
         Height          =   735
         Left            =   9480
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
      Begin VB.CommandButton CmdAgregar 
         Caption         =   "&Nuevo"
         Height          =   735
         Left            =   480
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   360
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
    
   

Private Sub ActualizarUN()

    Set baseQP = OpenDatabase(App.Path & "\Padron.mdb")
    Set tUltNums = baseQP.OpenRecordset("UltNums", dbOpenTable)
    
    tUltNums.Index = "PrimaryKey"
    
    tUltNums.Seek "=", "PROV"
    
    If Not tUltNums.NoMatch Then
        If Not tUltNums.EOF Then
            tUltNums.Edit
                tUltNums!UltNum = TextCodigoProveedor.Text
            tUltNums.Update
        End If
    End If
    
    tUltNums.Close
    baseQP.Close

End Sub

Private Sub SeteoGrilla()

    FG1.Visible = True
    
    FG1.Rows = 2
    FG1.Cols = 7

    FG1.Row = 0
    
    FG1.Col = 0
    FG1.CellAlignment = 7
    FG1.ColWidth(0) = 800
    FG1.Text = "Proveedor"
    
    FG1.Col = 1
    FG1.CellAlignment = 1
    FG1.ColWidth(1) = 2000
    FG1.Text = "Nombre"
    
    FG1.Col = 2
    FG1.CellAlignment = 4
    FG1.ColWidth(2) = 1200
    FG1.Text = "CUIT"
    
    FG1.Col = 3
    FG1.CellAlignment = 1
    FG1.ColWidth(3) = 2000
    FG1.Text = "Dirección"
    
    FG1.Col = 4
    FG1.CellAlignment = 4
    FG1.ColWidth(4) = 2000
    FG1.Text = "Localidad"
    
    FG1.Col = 5
    FG1.CellAlignment = 4
    FG1.ColWidth(5) = 1000
    FG1.Text = "Provincia"
    
    FG1.Col = 6
    FG1.CellAlignment = 4
    FG1.ColWidth(6) = 1300
    FG1.Text = "CP"

End Sub

Private Sub CmdAgregar_Click()

    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
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


Private Sub cmdModifica_Click()

    tProveedores.Index = "Primary"
    tProveedores.Seek "=", TextCodigoProveedor.Text
    
    If Not tProveedores.NoMatch Then
        tProveedores.Edit
            
            tProveedores!CodProv = Val(TextCodigoProveedor.Text)
            tProveedores!NombreProv = TextNombreProveedor.Text
            tProveedores!Cuit = TextCUIT.Text
            tProveedores!Direccion = TextDireccion.Text
            tProveedores!Localidad = TextLocalidad.Text
            tProveedores!Cp = TextCodigoPostal.Text
            tProveedores!Provincia = TextProvincia.Text
        
        tProveedores.Update
        
        a = MsgBox("Registro Modificado con Exito !!!", vbOKOnly, "INFO DEL SISTEMA")
    
    End If

End Sub

Private Sub cmdVerProv_Click()

    'On Error GoTo CapturaErrores
    
    Set baseQP = OpenDatabase(App.Path & "\Padron.mdb")
    Set tProveedores = baseQP.OpenRecordset("Proveedores", dbOpenTable)
    
    Call SeteoGrilla
    
    FG1.Visible = True
    
    FG1.Row = 1
    
    tProveedores.MoveFirst
    
    While Not tProveedores.EOF
        FG1.Col = 0
        FG1.Text = tProveedores!CodProv
        FG1.Col = 1
        If tProveedores!NombreProv <> "" Then FG1.Text = tProveedores!NombreProv
        FG1.Col = 2
        If tProveedores!Cuit <> "" Then FG1.Text = tProveedores!Cuit
        FG1.Col = 3
        If tProveedores!Direccion <> "" Then FG1.Text = tProveedores!Direccion
        FG1.Col = 4
        If tProveedores!Localidad <> "" Then FG1.Text = tProveedores!Localidad
        FG1.Col = 5
        If tProveedores!Provincia <> "" Then FG1.Text = tProveedores!Provincia
        FG1.Col = 6
        If tProveedores!Cp <> "" Then FG1.Text = tProveedores!Cp
            
        tProveedores.MoveNext
        FG1.Rows = FG1.Rows + 1
        FG1.Row = FG1.Row + 1
    Wend
    
    cmdModifica.Enabled = True

CapturaErrores:
    Select Case Err
        Case 3021
            a = MsgBox("No Hay Registros en el Archivo", vbOKOnly, "ERROR")
            Resume Next
        Case 94
            Resume Next
        Case Else
       
    End Select

End Sub

Private Sub ComEliminar_Click()

    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
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

    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
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
           Call ActualizarUN
           Call blanco
           TextCodigoProveedor.Enabled = True
           TextCodigoProveedor.SetFocus
        Else
            a = MsgBox("Código de Proveedor Existente !!!", vbOKOnly, "ERROR !!!")
            TextCodigoProveedor.SetFocus
           ' CodigoProveedor = Val(TextCodigoProveedor.Text)
           ' rst.FindFirst "CodProv= " + Str(CodigoProveedor)
           ' rst.Edit
           ' Call muevo
           ' Call blanco
           ' TextCodigoProveedor.Enabled = True
           ' TextCodigoProveedor.SetFocus
        End If
    End If

End Sub
Private Sub muevo()

    rst.Fields!CodProv = Val(TextCodigoProveedor.Text)
    rst.Fields!NombreProv = TextNombreProveedor.Text
    rst.Fields!Cuit = TextCUIT.Text
    rst.Fields!Direccion = TextDireccion.Text
    rst.Fields!Localidad = TextLocalidad.Text
    rst.Fields!Cp = TextCodigoPostal.Text
    rst.Fields!Provincia = TextProvincia.Text
    rst.Update

End Sub



Private Sub FG1_DblClick()

    FG1.Col = 0
    TextCodigoProveedor.Text = FG1.Text
    
    FG1.Col = 1
    TextNombreProveedor.Text = FG1.Text
    
    FG1.Col = 2
    TextCUIT.Text = FG1.Text
    
    FG1.Col = 3
    TextDireccion.Text = FG1.Text
    
    FG1.Col = 4
    TextLocalidad.Text = FG1.Text
    
    FG1.Col = 5
    TextCodigoPostal.Text = FG1.Text
    
    FG1.Col = 6
    TextProvincia.Text = FG1.Text
    
    FG1.Visible = False

End Sub


Private Sub TextCodigoPostal_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub TextCodigoProveedor_GotFocus()

    TextCodigoProveedor.SelLength = Len(TextCodigoProveedor.Text)
    

End Sub

Private Sub TextCodigoProveedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If


'    Set db = DBEngine.OpenDatabase("C:\QuilplacVB\Padron.mdb")
'    Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)
 
    
' Base Datos Padron
' Tabla Proveedores
' Campos
    ' CodProv Numerico
    ' NombreProv Texto 50
    ' Cuit Texto 11
    ' Estado Texto 2
    

'    If KeyAscii = 13 Then
'    If TextCodigoProveedor.Text = "" Then
'       TextNombreProveedor.SetFocus
'    Else
'        CodigoProveedor = Val(TextCodigoProveedor.Text)
      
'        rst.FindFirst "CodProv= " + Str(CodigoProveedor)
'        If rst.Fields!CodProv <> Val(TextCodigoProveedor.Text) Then
'           mensaje = MsgBox("Proveedor Inexistente", vbCritical, "Final de la busqueda")
'           TextCodigoProveedor.Text = ""
'           Call blanco
'           TextCodigoProveedor.SetFocus
'        Else
'           TextCodigoProveedor.Text = rst.Fields!CodProv
'           TextNombreProveedor.Text = rst.Fields!NombreProv
'           TextCUIT.Text = rst.Fields!CUIT
'           TextDireccion.Text = rst.Fields!Direccion
'           TextLocalidad.Text = rst.Fields!Localidad
'           TextCodigoPostal.Text = rst.Fields!CP
'           TextProvincia.Text = rst.Fields!Provincia
'        End If
          
'    End If
'    End If

End Sub

Private Sub blanco()

    TextCodigoProveedor.Text = ""
    TextNombreProveedor.Text = ""
    TextCUIT.Text = ""
    TextDireccion.Text = ""
    TextLocalidad.Text = ""
    TextCodigoPostal.Text = ""
    TextProvincia.Text = ""
    
End Sub

    

Private Sub Form_Load()
    
    FormProveedores.Height = 7950
    FormProveedores.Width = 11205
    
    

    Set baseQP = OpenDatabase(App.Path & "\Padron.mdb")
    Set tUltNums = baseQP.OpenRecordset("UltNums", dbOpenTable)
    
    tUltNums.Index = "PrimaryKey"
    
    tUltNums.Seek "=", "PROV"
    
    If Not tUltNums.NoMatch Then
        If Not tUltNums.EOF Then
            TextCodigoProveedor.Text = tUltNums!UltNum + 1
        End If
    End If
    
    tUltNums.Close
    baseQP.Close
    
    'Call blanco
   

End Sub


Private Sub TextCUIT_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub TextDireccion_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
    
End Sub


Private Sub TextLocalidad_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If


End Sub


Private Sub TextNombreProveedor_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

'    If KeyAscii = 13 Then
'        Call busco
'    End If
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub

Private Sub busco()

    Set db = DBEngine.OpenDatabase(App.Path & "\Padron.mdb")
    Set rst = db.OpenRecordset("Proveedores", dbOpenDynaset)
    
    FG1.Rows = 2
    FG1.Clear
    FG1.Visible = True
    
    Call titulos
    
    Dim busca1 As String, busca2 As String
    busca1 = RTrim(LTrim(TextNombreProveedor.Text))
    busca2 = busca1 + "z"
    
    rst.FindFirst "NombreProv >= '" & busca1 & "' and NombreProv <= '" & busca2 & "'"
    
    If rst.NoMatch Then
       FG1.Visible = False
       mensaje = MsgBox("No existen Proveedores", vbCritical, "Final de la busqueda")
       TextNombreProveedor.Text = ""
       Call blanco
       TextNombreProveedor.SetFocus
    End If
     
    linea2 = 1
    Do While Not rst.NoMatch
        FG1.AddItem " "
        FG1.Row = linea2
       
            FG1.Col = 0
            FG1.Text = rst.Fields!CodProv
            FG1.Col = 1
            FG1.Text = rst.Fields!NombreProv
            FG1.Col = 2
            FG1.Text = rst.Fields!Cuit
            FG1.Col = 3
            FG1.Text = rst.Fields!Direccion
            FG1.Col = 4
            FG1.Text = rst.Fields!Localidad
            FG1.Col = 5
            FG1.Text = rst.Fields!Cp
            FG1.Col = 6
            FG1.Text = rst.Fields!Provincia
            linea2 = linea2 + 1
      
       rst.FindNext "NombreProv >= '" & busca1 & "' and NombreProv <= '" & busca2 & "'"
       
    Loop
    
    
End Sub

Private Sub titulos()

    FG1.Row = 0
    
    FG1.Col = 0
    FG1.Text = "Codigo"
    FG1.ColWidth(0) = 900
    
    FG1.Col = 1
    FG1.Text = "Apellido y Nombre"
    FG1.ColWidth(1) = 4700
    
    FG1.Col = 2
    FG1.Text = "CUIT"
    FG1.ColWidth(2) = 1200
    
    FG1.Col = 3
    FG1.Text = "Direccion"
    FG1.ColWidth(3) = 0
    
    FG1.Col = 4
    FG1.Text = "Localidad"
    FG1.ColWidth(4) = 0
    
    FG1.Col = 5
    FG1.Text = "CP"
    FG1.ColWidth(5) = 0
    
    FG1.Col = 6
    FG1.Text = "Provincia"
    FG1.ColWidth(6) = 0

End Sub

Private Sub TextProvincia_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


