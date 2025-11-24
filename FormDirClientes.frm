VERSION 5.00
Begin VB.Form FormDireccionesEntrega 
   Caption         =   "CLIENTES - DIRECCIONES DE ENTREGA"
   ClientHeight    =   5085
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5085
   ScaleWidth      =   13980
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   0
      Width           =   13695
      Begin VB.TextBox txtRazonSocial 
         Height          =   375
         Left            =   5160
         TabIndex        =   30
         Top             =   360
         Width           =   6495
      End
      Begin VB.TextBox txtIDCliente 
         Height          =   375
         Left            =   3600
         TabIndex        =   29
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblItem 
         AutoSize        =   -1  'True
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   12600
         TabIndex        =   32
         Top             =   240
         Width           =   225
      End
      Begin VB.Image Image1 
         Height          =   570
         Left            =   240
         Picture         =   "FormDirClientes.frx":0000
         Stretch         =   -1  'True
         Top             =   240
         Width           =   2295
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         Left            =   2760
         TabIndex        =   31
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Acciones"
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   3360
      Width           =   13695
      Begin VB.CommandButton btnPrimero 
         Caption         =   "|<"
         Height          =   615
         Left            =   240
         TabIndex        =   27
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton btnAtras 
         Caption         =   "<<"
         Height          =   615
         Left            =   1440
         TabIndex        =   26
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton btnAdelante 
         Caption         =   ">>"
         Height          =   615
         Left            =   2640
         TabIndex        =   25
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton btnUltimo 
         Caption         =   ">|"
         Height          =   615
         Left            =   3840
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton btnGrabar 
         Caption         =   "&Grabar"
         Height          =   615
         Left            =   6480
         TabIndex        =   8
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton btnModificar 
         Caption         =   "&Modificar"
         Height          =   615
         Left            =   7920
         TabIndex        =   23
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton btnBuscar 
         Caption         =   "&Buscar"
         Height          =   615
         Left            =   9360
         TabIndex        =   22
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton btnEliminar 
         Caption         =   "&Eliminar"
         Height          =   615
         Left            =   10800
         TabIndex        =   21
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton btnSalir 
         Caption         =   "&Salir"
         Height          =   615
         Left            =   12240
         TabIndex        =   20
         Top             =   360
         Width           =   1200
      End
      Begin VB.CommandButton btnLimpiar 
         Caption         =   "&Limpiar"
         Height          =   615
         Left            =   5040
         TabIndex        =   19
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Dirección de Entrega Alternativa"
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   13695
      Begin VB.TextBox txtDomicilio 
         Height          =   375
         Left            =   6840
         TabIndex        =   3
         Top             =   600
         Width           =   6495
      End
      Begin VB.TextBox txtCP 
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cmbProv 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cmbPais 
         Height          =   315
         Left            =   480
         TabIndex        =   0
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox txtTel 
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox txtCel 
         Height          =   375
         Left            =   4320
         TabIndex        =   6
         Top             =   1440
         Width           =   2175
      End
      Begin VB.TextBox txtEmail 
         Height          =   375
         Left            =   6840
         TabIndex        =   7
         Top             =   1440
         Width           =   6495
      End
      Begin VB.ComboBox cmbLocalidad 
         Height          =   315
         Left            =   3840
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Domicilio"
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
         Left            =   6720
         TabIndex        =   17
         Top             =   360
         Width           =   780
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "CPA"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1200
         Width           =   375
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Provincia"
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
         Left            =   1560
         TabIndex        =   15
         Top             =   360
         Width           =   810
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "País"
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
         Left            =   360
         TabIndex        =   14
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Teléfono"
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
         Left            =   1920
         TabIndex        =   13
         Top             =   1200
         Width           =   750
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Celular"
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
         Left            =   4200
         TabIndex        =   12
         Top             =   1200
         Width           =   720
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "e-mail"
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
         Left            =   6720
         TabIndex        =   11
         Top             =   1200
         Width           =   510
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Localidad"
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
         Left            =   3720
         TabIndex        =   10
         Top             =   360
         Width           =   840
      End
   End
End
Attribute VB_Name = "FormDireccionesEntrega"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private KeyRetroceso As Boolean
Private vContadorDirecciones As Integer

Private Sub btnAdelante_Click()

    On Error GoTo CapturaErrores
    
    If Not tDomiciliosClientes.EOF Then
        tDomiciliosClientes.MoveNext
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
    Else
        MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
        tDomiciliosClientes.MoveLast
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            'MsgBox "Ultimo Registro", vbDefaultButton1, "INFO DEL SISTEMA"
            MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
            tDomiciliosClientes.MoveLast
            Call Mostrar
            Resume Next
    End Select


End Sub

Private Sub btnAtras_Click()

    On Error GoTo CapturaErrores
    
    If Not tDomiciliosClientes.BOF Then
        tDomiciliosClientes.MovePrevious
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
    Else
        MsgBox "Primer Registro", vbInformation, "INFO DEL SISTEMA"
        tDomiciliosClientes.MoveFirst
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "Primer Registro", vbInformation + vbOKOnly, "INFO DEL SISTEMA"
            'MsgBox "No hay registros !!!", vbDefaultButton1, "INFO DEL SISTEMA"
            tDomiciliosClientes.MoveFirst
            Call Mostrar
            Resume Next
    End Select


End Sub

Private Sub btnBuscar_Click()

        vSQL = "SELECT * FROM DomiciliosClientes WHERE IDCliente =" & txtIDCliente.Text & " ORDER BY IDCliente"
        'MsgBox (vsql)
        Set tDomiciliosClientes = BaseSPC.OpenRecordset(vSQL)
        
        If Not tDomiciliosClientes.EOF Then
            tDomiciliosClientes.MoveFirst
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If
 '   End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select


End Sub

Private Sub btnEliminar_Click()
    
    'Cuando esten listas las facturas y los presupuestos
    'hay que controlar que no existan movimientos con el cliente.
    ' de existir los clientes hay que marcarlo como baja en el campo dado de baja.
    
    A = MsgBox("¿ Confirma Eliminar el Registro ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
    If A = 1 Then
        With tDomiciliosClientes
            .Delete
        End With
    End If
    
    Call EnabledTextBox(FormDireccionesEntrega)
    txtIDCliente.Enabled = False
    txtRazonSocial.Enabled = False
    
    Call LimpiarPantalla

End Sub

Private Sub btnGrabar_Click()
    
    A = MsgBox("¿ Seguro Genera Nuevo Registro ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
'    tDomiciliosClientes.Index = "PrimaryKey"
    
     With tDomiciliosClientes
        .AddNew
            !IdCliente = txtIDCliente.Text
            !item = CInt(lblItem.Caption)
            'lblItem.Caption
            !Domicilio = Format(txtDomicilio.Text, ">")
            !Localidad = Format(cmbLocalidad.Text, ">")
            !Pais = Format(cmbPais.Text, ">")
            !Prov = Format(cmbProv.Text, ">")
            !Localidad = Format(cmbLocalidad.Text, ">")
            !CP = txtCP.Text
            !Tel = txtTel.Text
            !Cel = txtCel.Text
            !email = Format(txtEmail.Text, ">")
     
            If A = 1 Then
               Call LimpiarPantalla
               cmbPais.SetFocus
               lblItem.Caption = CInt(lblItem.Caption) + 1
               CantidadDirecciones = lblItem.Caption
            End If
        .Update
     End With
    

CapturaErrores:
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub

Private Sub btnLimpiar_Click()

    Call LimpiarPantalla

End Sub

Private Sub btnModificar_Click()

    If vFlagBuscar = 0 Then
        Call EnabledTextBox(FormDireccionesEntrega)
        btnModificar.Caption = "Guardar Cambios"
        cmbPais.SetFocus
        txtIDCliente.Enabled = False
        txtRazonSocial.Enabled = False
        vFlagBuscar = 1
        
    Else
        A = MsgBox("¿Seguro desea Guardar las Modificaciones Realizadas?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
        
        '1 es Ok 2 es Cancel
            If A = 1 Then
                With tDomiciliosClientes
                    .Edit
                        !Domicilio = Format(txtDomicilio.Text, ">")
                        !Localidad = Format(cmbLocalidad.Text, ">")
                        !Pais = Format(cmbPais.Text, ">")
                        !Prov = Format(cmbProv.Text, ">")
                        !Localidad = Format(cmbLocalidad.Text, ">")
                        !CP = txtCP.Text
                        !Tel = txtTel.Text
                        !Cel = txtCel.Text
                        !email = Format(txtEmail.Text, ">")
                        
                       ' tEmpleados.Index = "IndiceNombre"
                       ' tEmpleados.Seek "=", cmbVendedor.Text
                        
                       ' If Not tEmpleados.NoMatch Then !Vendedor = tEmpleados!Legajo
                    .Update
                End With
                
                Call EnabledTextBox(FormClientes)
                
                txtIDCliente.Enabled = False
                txtRazonSocial.Enabled = False
                
                Call LimpiarPantalla
                
             Else
                If A = 2 Then
                    b = MsgBox("¿Limpia Pantalla?", vbQuestion + vbYesNo, "INFO DEL SISTEMA")
                    If b = 6 Then
                        Call LimpiarPantalla
                     Else
                          Call EnabledTextBox(FormDireccionesEntrega)
                          txtIDCliente.Enabled = False
                          txtRazonSocial.Enabled = False
                          cmbPais.SetFocus
                    End If
                End If
            End If
        End If

End Sub

Private Sub LimpiarPantalla()

    txtDomicilio.Text = ""
    cmbLocalidad.Text = ""
    txtCP.Text = ""
    txtTel.Text = ""
    txtCel.Text = ""
    txtEmail.Text = ""
    cmbPais.Text = ""
    cmbProv.Text = ""
    cmbLocalidad.Text = ""
    btnGrabar.Enabled = True
    btnEliminar.Enabled = False
    btnModificar.Caption = "&Modificar"
    btnModificar.Enabled = False
    
    Call EnabledTextBox(FormDireccionesEntrega)
    lblItem.Caption = CantidadDirecciones
    cmbPais.SetFocus

End Sub

Private Sub btnPrimero_Click()
    
    On Error GoTo CapturaErrores
    
   ' If Not tDomiciliosClientes.EOF Then
   '     tDomiciliosClientes.MoveFirst
   '     Call Mostrar
   '     btnGrabar.Enabled = False
   '     btnModificar.Enabled = True
   '     btnEliminar.Enabled = True
   '  Else
        vSQL = "SELECT * FROM DomiciliosClientes WHERE IDCliente =" & txtIDCliente.Text & " ORDER BY IDCliente"
        'MsgBox (vSQL)
        
        Set tDomiciliosClientes = BaseSPC.OpenRecordset(vSQL)
        
        If Not tDomiciliosClientes.EOF Then
            tDomiciliosClientes.MoveFirst
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If
   ' End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select

End Sub

Private Sub Mostrar()
     
     With tDomiciliosClientes
            txtIDCliente.Text = !IdCliente
            txtDomicilio.Text = !Domicilio
            cmbLocalidad.Text = !Localidad
            cmbPais.Text = !Pais
            cmbProv.Text = !Prov
            txtCP.Text = !CP
            
            If !Tel <> "" Then txtTel.Text = !Tel
            If !Cel <> "" Then txtCel.Text = !Cel
            If !email <> "" Then txtEmail.Text = !email
            
            lblItem.Caption = !item
            
            vFlagBuscar = 0
            
     End With
     
     Call DisabledTextBox(FormDireccionesEntrega)
     
End Sub

Private Sub btnSalir_Click()
    
    Unload Me
    
End Sub

Private Sub btnUltimo_Click()

    On Error GoTo CapturaErrores
    
'    If Not tDomiciliosClientes.EOF Then
'        tDomiciliosClientes.MoveLast
'        Call Mostrar
'        btnGrabar.Enabled = False
'        btnModificar.Enabled = True
'        btnEliminar.Enabled = True
'     Else
        'vSQL = "SELECT * FROM Clientes ORDER BY IDCliente"
        vSQL = "SELECT * FROM DomiciliosClientes WHERE IDCliente =" & txtIDCliente.Text & " ORDER BY IDCliente"
        'MsgBox (vsql)
        Set tDomiciliosClientes = BaseSPC.OpenRecordset(vSQL)
        
        If Not tDomiciliosClientes.EOF Then
            tDomiciliosClientes.Last
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If
 '   End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay más registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select


End Sub

Private Sub cmbLocalidad_Change()
    
    Autocompletar_Combo cmbLocalidad

End Sub

Private Sub cmbLocalidad_GotFocus()
    cmbLocalidad.SelLength = Len(cmbLocalidad.Text)
End Sub

Private Sub cmbLocalidad_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete, vbKeyReturn
            Select Case Len(cmbLocalidad.Text)
                Case Is <> 0
                    KeyRetroceso = True
              End Select
    End Select

End Sub


Private Sub cmbLocalidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    'Si le pasamos a SendMessageLong el valor False lo cierra
        resp = SendMessageLong(cmbLocalidad.hwnd, &H14F, False, 0)
        SendKeys "{TAB}"
     Else
    'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
    'presionamos una tecla diferente al Enter
        resp = SendMessageLong(cmbLocalidad.hwnd, &H14F, True, 0)
    End If

End Sub


Private Sub cmbLocalidad_LostFocus()
    
    tLocalidades.MoveLast
    tLocalidades.FindFirst tLocalidades.IDLocalidad = cmbLocalidad.Text
   
    If Not tLocalidades.NoMatch Then
        vSQL = "SELECT IDLocalidad, Descripcion FROM Localidades Where IDProv=" & tProvincias!IDProv & " ORDER BY Descripcion"
        Set tLocalidades = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
        
        cmbLocalidad.Clear
        While Not tLocalidades.EOF
            cmbLocalidad.AddItem tLocalidades!Descripcion
            tLocalidades.MoveNext
        Wend
        
    Else
        'MsgBox ("No se encuentra la Provincia Seleccionada")
    End If

End Sub


Private Sub cmbPais_Change()
    
    Autocompletar_Combo cmbPais

End Sub

Private Sub cmbPais_GotFocus()
    cmbPais.SelLength = Len(cmbPais.Text)
End Sub

Private Sub cmbPais_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete, vbKeyReturn
            Select Case Len(cmbPais.Text)
                Case Is <> 0
                    KeyRetroceso = True
  
            End Select
    End Select
    
End Sub


Private Sub cmbPais_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
    'Si le pasamos a SendMessageLong el valor False lo cierra
        resp = SendMessageLong(cmbPais.hwnd, &H14F, False, 0)
        SendKeys "{TAB}"
     Else
    'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
    'presionamos una tecla diferente al Enter
        resp = SendMessageLong(cmbPais.hwnd, &H14F, True, 0)
    End If

End Sub


Private Sub cmbPais_LostFocus()

    tPaises.Index = "IndiceDescripcion"
    tPaises.Seek "=", cmbPais.Text
    
    If tPaises.NoMatch Then
        'MsgBox ("No se encuentra el País Seleccionado")
    Else
        
        vSQL = "SELECT IDProv, Descripcion FROM Provincias Where IDPais=" & tPaises!IDPais & " ORDER BY Descripcion"
        Set tProvincias = BaseSPC.OpenRecordset(vSQL)
        
        cmbProv.Clear
        While Not tProvincias.EOF
            cmbProv.AddItem tProvincias!Descripcion
            tProvincias.MoveNext
        Wend
    End If


End Sub


Private Sub cmbProv_Change()
    
    Autocompletar_Combo cmbProv
    
End Sub


Private Sub cmbProv_GotFocus()
    cmbProv.SelLength = Len(cmbProv.Text)
End Sub

Private Sub cmbProv_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete, vbKeyReturn
            Select Case Len(cmbProv.Text)
                Case Is <> 0
                    KeyRetroceso = True
  
            End Select
    End Select

End Sub


Private Sub cmbProv_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    'Si le pasamos a SendMessageLong el valor False lo cierra
        resp = SendMessageLong(cmbProv.hwnd, &H14F, False, 0)
        SendKeys "{TAB}"
     Else
    'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
    'presionamos una tecla diferente al Enter
        resp = SendMessageLong(cmbProv.hwnd, &H14F, True, 0)
    End If

End Sub


Private Sub cmbProv_LostFocus()
    tProvincias.MoveLast
    
    strProvincia = "Descripcion = '" & cmbProv.Text & "'"

    tProvincias.FindFirst strProvincia
    'tProvincias.Descripcion = cmbProv.Text
    
    If tProvincias.NoMatch Then
         
       MsgBox Provincias!Descripcion & " No Existe", vbCritical + vbOKOnly, "ERROR"
    
    End If
    
   
    If Not tProvincias.EOF Then
        vSQL = "SELECT IDLocalidad, Descripcion FROM Localidades Where IDProv=" & tProvincias!IDProv & " ORDER BY Descripcion"
        Set tLocalidades = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
        
        cmbLocalidad.Clear
        While Not tLocalidades.EOF
            cmbLocalidad.AddItem tLocalidades!Descripcion
            tLocalidades.MoveNext
        Wend
        
    Else
        'MsgBox ("No se encuentra la Provincia Seleccionada")
    End If

End Sub


Private Sub Form_Load()

    'Declaro variable contador de direcciones
        'Dim vContadorDirecciones As Integer
        vContadorDirecciones = 0
    
    'Paso los datos del cliente del form anterior
        txtIDCliente.Text = FormClientes.txtIDCliente.Text
        txtIDCliente.Enabled = False
        txtRazonSocial.Text = FormClientes.txtRazonSocial.Text
        txtRazonSocial.Enabled = False
        
    'Controlo que exista relación
      '  Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
      '      tClientes.Index = "PrimaryKey"
      '      tClientes.Seek "=", FormClientes.txtIDCliente.Text
      '      If tClientes.NoMatch Then
                
      '          If MsgBox("Debe Grabar el Cliente para Luego Ingresar las Direcciones de Entrega" & Chr(10) & "¿Graba el Registro?", vbQuestion + vbYesNo, "ERROR") Then
      '              FormClientes.btnGrabar_Click
      '           Else
      '              End
      '              GoTo CapturaErrores
      '          End If
      '      End If
    
    'Inicializo el Label del contador de direcciones
        lblItem.Caption = vContadorDirecciones
        
    'Etiqueta controladora de errores de bbdd
        'On Error GoTo CapturaErrores
    
    'Lleno combo de países
        tPaises.MoveFirst
        While Not tPaises.EOF
            cmbPais.AddItem tPaises!Descripcion
            tPaises.MoveNext
        Wend
    
    'Tabla Direcciones de Clientes
        Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
        
        tDomiciliosClientes.Index = "PrimaryKey"
        tDomiciliosClientes.Seek "=", txtIDCliente.Text
        
        If Not tDomiciliosClientes.NoMatch Then
            While Not tDomiciliosClientes.EOF
                item = tDomiciliosClientes!item
                tDomiciliosClientes.MoveNext
            Wend
            lblItem.Caption = CInt(item + 1)
            CantidadDirecciones = lblItem.Caption
        Else
            lblItem.Caption = 1
            CantidadDirecciones = lblItem.Caption
        End If

CapturaErrores:
    
  'Captura el error de no hay registros en la tabla
    Select Case Err
        Case 3021
            Resume Next
    End Select

End Sub

Private Sub txtCel_GotFocus()
    txtCel.SelLength = Len(txtCel.Text)
End Sub

Private Sub txtCel_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
    
     KeyAscii = Verificar_Tecla(KeyAscii)


End Sub


Private Sub txtCP_GotFocus()
    txtCP.SelLength = Len(txtCP.Text)
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub txtDomicilio_GotFocus()
    txtDomicilio.SelLength = Len(txtDomicilio.Text)
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If


End Sub


Private Sub txtEmail_GotFocus()
    txtEmail.SelLength = Len(txtEmail.Text)
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbLowerCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub



Private Sub txtIDCliente_KeyPress(KeyAscii As Integer)
    txtIDCliente.SelLength = Len(txtIDCliente.Text)
End Sub



Private Sub txtRazonSocial_GotFocus()
    txtRazonSocial.SelLength = Len(txtRazonSocial.Text)
End Sub

Private Sub txtTel_GotFocus()
    txtTel.SelLength = Len(txtTel.Text)
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
    
     KeyAscii = Verificar_Tecla(KeyAscii)

End Sub


