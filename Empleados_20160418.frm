VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FormEmpleados 
   Caption         =   "ABM Empleados"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13770
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   13770
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Empleado"
      Height          =   7695
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   13695
      Begin VB.Frame Frame5 
         Caption         =   "Acciones"
         Height          =   1215
         Left            =   360
         TabIndex        =   20
         Top             =   6000
         Width           =   12855
         Begin VB.CommandButton btnLimpiar 
            Caption         =   "&Limpiar"
            Height          =   615
            Left            =   5640
            TabIndex        =   42
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnSalir 
            Caption         =   "&Salir"
            Height          =   615
            Left            =   11760
            TabIndex        =   43
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnEliminar 
            Caption         =   "&Eliminar"
            Height          =   615
            Left            =   9960
            TabIndex        =   41
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnBuscar 
            Caption         =   "&Buscar"
            Height          =   615
            Left            =   8880
            TabIndex        =   40
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnModificar 
            Caption         =   "&Modificar"
            Height          =   615
            Left            =   7800
            TabIndex        =   39
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnGrabar 
            Caption         =   "&Grabar"
            Height          =   615
            Left            =   6720
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnUltimo 
            Caption         =   ">|"
            Height          =   615
            Left            =   3480
            TabIndex        =   38
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAdelante 
            Caption         =   ">>"
            Height          =   615
            Left            =   2400
            TabIndex        =   37
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAtras 
            Caption         =   "<<"
            Height          =   615
            Left            =   1320
            TabIndex        =   36
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnPrimero 
            Caption         =   "|<"
            Height          =   615
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos Personales"
         Height          =   2655
         Left            =   360
         TabIndex        =   19
         Top             =   3240
         Width           =   12975
         Begin VB.TextBox txtCP 
            CausesValidation=   0   'False
            Height          =   285
            Left            =   8880
            TabIndex        =   50
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox txtDniEmple 
            Height          =   405
            Left            =   360
            TabIndex        =   9
            Top             =   1320
            Width           =   1935
         End
         Begin VB.TextBox txtFechaNacEmple 
            Height          =   375
            Left            =   8040
            TabIndex        =   8
            Top             =   480
            Width           =   2055
         End
         Begin MSComCtl2.MonthView MonthView2 
            Height          =   2370
            Left            =   10200
            TabIndex        =   48
            Top             =   120
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            StartOfWeek     =   74383361
            CurrentDate     =   41765
         End
         Begin VB.ComboBox cmbLocalidad 
            Height          =   315
            Left            =   6000
            TabIndex        =   12
            Top             =   1320
            Width           =   2775
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   4920
            TabIndex        =   15
            Top             =   2040
            Width           =   5175
         End
         Begin VB.TextBox txtCel 
            Height          =   375
            Left            =   2640
            TabIndex        =   14
            Top             =   2040
            Width           =   2055
         End
         Begin VB.TextBox txtTel 
            Height          =   375
            Left            =   360
            TabIndex        =   13
            Top             =   2040
            Width           =   2055
         End
         Begin VB.ComboBox cmbPais 
            Height          =   315
            Left            =   2640
            TabIndex        =   10
            Top             =   1320
            Width           =   1095
         End
         Begin VB.ComboBox cmbProv 
            Height          =   315
            Left            =   3960
            TabIndex        =   11
            Top             =   1320
            Width           =   1815
         End
         Begin VB.TextBox txtDomicilio 
            Height          =   375
            Left            =   360
            TabIndex        =   7
            Top             =   480
            Width           =   7455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "DNI"
            Height          =   195
            Left            =   480
            TabIndex        =   49
            Top             =   1080
            Width           =   285
         End
         Begin VB.Label Label18 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Nacimiento"
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
            Index           =   1
            Left            =   8040
            TabIndex        =   46
            Top             =   240
            Width           =   1815
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Height          =   195
            Left            =   360
            TabIndex        =   45
            Top             =   2040
            Width           =   45
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
            Left            =   6000
            TabIndex        =   44
            Top             =   1080
            Width           =   840
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
            Index           =   0
            Left            =   4920
            TabIndex        =   34
            Top             =   1800
            Width           =   510
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
            Left            =   2640
            TabIndex        =   33
            Top             =   1800
            Width           =   720
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
            Left            =   360
            TabIndex        =   32
            Top             =   1800
            Width           =   750
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
            Left            =   2640
            TabIndex        =   31
            Top             =   1080
            Width           =   390
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
            Left            =   4080
            TabIndex        =   30
            Top             =   1080
            Width           =   810
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "Cod. Postal"
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
            Left            =   9000
            TabIndex        =   29
            Top             =   1080
            Width           =   990
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
            Left            =   360
            TabIndex        =   28
            Top             =   240
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Información Principal"
         Height          =   2655
         Left            =   360
         TabIndex        =   18
         Top             =   360
         Width           =   12975
         Begin VB.TextBox txtFechaIngEmple 
            Height          =   375
            Left            =   8040
            TabIndex        =   3
            Top             =   600
            Width           =   2055
         End
         Begin MSComCtl2.MonthView MonthView1 
            Height          =   2370
            Left            =   10200
            TabIndex        =   47
            Top             =   120
            Width           =   2595
            _ExtentX        =   4577
            _ExtentY        =   4180
            _Version        =   393216
            ForeColor       =   -2147483630
            BackColor       =   -2147483633
            Appearance      =   1
            StartOfWeek     =   74383361
            CurrentDate     =   41765
         End
         Begin VB.TextBox txtPuestoEmple 
            Alignment       =   2  'Center
            CausesValidation=   0   'False
            Height          =   375
            Left            =   2280
            TabIndex        =   5
            Top             =   1800
            Width           =   1095
         End
         Begin VB.TextBox txtObservaciones 
            CausesValidation=   0   'False
            Height          =   735
            Left            =   3840
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   6
            Top             =   1800
            Width           =   6255
         End
         Begin VB.TextBox txtComsionEmple 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   1800
            Width           =   1455
         End
         Begin VB.TextBox txtCUILEmple 
            Height          =   375
            Left            =   6480
            TabIndex        =   2
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtNombreEmple 
            Height          =   375
            Left            =   1440
            TabIndex        =   1
            Top             =   600
            Width           =   4815
         End
         Begin VB.TextBox txtLegEmple 
            Height          =   375
            Left            =   360
            TabIndex        =   0
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
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
            Left            =   4080
            TabIndex        =   27
            Top             =   1560
            Width           =   1275
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Comision "
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
            Left            =   480
            TabIndex        =   26
            Top             =   1560
            Width           =   825
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Ingreso"
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
            Left            =   8160
            TabIndex        =   25
            Top             =   360
            Width           =   1230
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Puesto "
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
            Left            =   2400
            TabIndex        =   24
            Top             =   1560
            Width           =   660
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "CUIL"
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
            Left            =   6480
            TabIndex        =   23
            Top             =   360
            Width           =   435
         End
         Begin VB.Label txtNombreEmpleado 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
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
            TabIndex        =   22
            Top             =   360
            Width           =   660
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Legajo"
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
            Left            =   480
            TabIndex        =   21
            Top             =   360
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "FormEmpleados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private KeyRetroceso As Boolean

Private Sub LimpiarPantalla()

    txtLegEmple.Text = ""
    txtNombreEmple.Text = ""
    txtCUILEmple.Text = ""
    txtFechaIngEmple.Text = ""
    txtDniEmple.Text = ""
    txtDomicilio.Text = ""
    cmbLocalidad.Text = ""
    txtCP.Text = ""
    txtTel.Text = ""
    txtCel.Text = ""
    txtEmail.Text = ""
    cmbPais.Text = ""
    cmbProv.Text = ""
    cmbLocalidad.Text = ""
    btnGrabar.Enabled = False
    btnEliminar.Enabled = False
    btnModificar.Caption = "&Modificar"
    txtObservaciones = ""
    'txtLegEmple.Text = tUltimosNumeros!UltimoNumero + 1
    Call EnabledTextBox(FormEmpleados)
    txtNombreEmple.SetFocus

End Sub

Private Sub Mostrar()

     With tEmpleados
    'Info Principal ---------------------------------------
            txtLegEmple.Text = !Legajo
            txtNombreEmple.Text = !nombre
            txtCUILEmple.Text = !Cuil
            txtFechaIngEmple.Text = !FechaIngreso
            txtComsionEmple.Text = !Comision
            txtPuestoEmple.Text = !IDPuesto
            txtObservaciones.Text = !Observaciones
    'Datos Personales---------------------------------------
            txtDomicilio.Text = !Domicilio
            txtFechaNacEmple.Text = !FechaNacimiento
            txtDniEmple.Text = !DNI
            cmbPais.Text = !Pais
            cmbProv.Text = !Prov
            cmbLocalidad.Text = !Localidad
            txtCP.Text = !CP
            txtTel.Text = !Tel
            txtCel.Text = !Cel
            txtEmail.Text = !emaill
            
     End With
     
     Call DisabledTextBox(FormEmpleados)

End Sub

Private Sub btnAdelante_Click()

    On Error GoTo CapturaErrores
    
    If Not tEmpleados.EOF Then
        tEmpleados.MoveNext
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
                
    Else
        MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
        tEmpleados.MoveLast
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            'MsgBox "Ultimo Registro", vbDefaultButton1, "INFO DEL SISTEMA"
            MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
            tEmpleados.MoveLast
            Call Mostrar
            Resume Next
    End Select

End Sub

Private Sub btnAtras_Click()
    
    On Error GoTo CapturaErrores
    
    If Not tEmpleados.BOF Then
        tEmpleados.MovePrevious
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
        
        
    Else
        MsgBox "Primer Registro", vbInformation, "INFO DEL SISTEMA"
        tEmpleados.MoveFirst
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "Primer Registro", vbInformation + vbOKOnly, "INFO DEL SISTEMA"
            'MsgBox "No hay registros !!!", vbDefaultButton1, "INFO DEL SISTEMA"
            tEmpleados.MoveFirst
            Call Mostrar
            Resume Next
    End Select

End Sub


Private Sub btnBuscar_Click()

    If vFlagBuscar = 0 Then
        vFlagBuscar = 1
        txtLegEmple.Enabled = True
        txtLegEmple.Text = ""
        txtLegEmple.SetFocus
     Else
        
        If txtLegEmple.Text <> "" Then
            Campo = "Legajo= "
            Valor = txtLegEmple.Text
         Else
            If txtNombreEmple.Text <> "" Then
                Campo = "Nombre Like "
                Valor = "'" + txtNombreEmple.Text + "*'"
             Else
                If txtNombreEmple.Text <> "" Then
                    Campo = "Nombre Like "
                    Valor = "'" + txtNombreEmple.Text + "*'"
                End If
            End If
        End If
        
        'vSQL = "SELECT IDProv, Descripcion FROM Provincias Where IDPais=" & tPaises!IDPais & " ORDER BY Descripcion"
        vSQL = "SELECT * FROM Empleados WHERE " & Campo & Valor & " ORDER BY Legajo"
        
        'MsgBox (vsql)
        
        Set tEmpleados = BaseSPC.OpenRecordset(vSQL)
    
        If Not tEmpleados.NoMatch Then
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
        Else
            MsgBox "No se encuentran registros", vbCritical, "ERROR"
        End If
        
        vFlagBuscar = 0
        
    End If
    
End Sub

Private Sub btnEliminar_Click()

    'Cuando esten listas las facturas y los presupuestos
    'hay que controlar que no existan movimientos con el cliente.
    ' de existir los clientes hay que marcarlo como baja en el campo dado de baja.
     
   b = MsgBox("¿ Seguro Desea Eliminar Empleado ?", vbQuestion + vbOKCancel, "Eliminar Empleado")
    
    With tEmpleados
        .Delete
    End With
    
    Call EnabledTextBox(FormEmpleados)
    Call LimpiarPantalla

End Sub

Private Sub btnGrabar_Click()

    A = MsgBox("¿ Seguro Genera Nuevo Empleado ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
    tEmpleados.Index = "PrimaryKey"
    
     With tEmpleados
        .AddNew
        'Datos Empleados -------------------------------------------------------------------
            !Legajo = txtLegEmple.Text
            !nombre = Format(txtNombreEmple.Text, ">")
            !Cuil = txtCUILEmple.Text
            !FechaIngreso = txtFechaIngEmple.Text
            !IDPuesto = txtPuestoEmple.Text
        'Datos Personales  -------------------------------------------------------------------
            !Domicilio = Format(txtDomicilio.Text, ">")
            !Localidad = Format(cmbLocalidad.Text, ">")
            !DNI = txtDniEmple.Text
            !Pais = Format(cmbPais.Text, ">")
            !Prov = Format(cmbProv.Text, ">")
            !Localidad = Format(cmbLocalidad.Text, ">")
            !CP = txtCP.Text
            !Tel = txtTel.Text
            !Cel = txtCel.Text
            !emaill = Format(txtEmail.Text, ">")
            !FechaNacimiento = txtFechaNacEmple.Text
            !Observaciones = Format(txtObservaciones.Text, ">")
            !IDPuesto = txtPuestoEmple.Text
            
            Call LimpiarPantalla
                   
        .Update
     
     End With
     
'    tUltimosNumeros.Index = "PrimaryKey"
'    tUltimosNumeros.Seek "=", "tEmpleados"
    
'    tUltimosNumeros.Edit
'    tUltimosNumeros!UltimoNumero = txtLegEmple.Text
'    tUltimosNumeros.Update
    
'    txtLegEmple.Text = tUltimosNumeros!UltimoNumero + 1
'    txtRazonSocial.SetFocus

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
        Call EnabledTextBox(FormEmpleados)
        btnModificar.Caption = "Guardar Cambios"
        txtLegEmple.SetFocus
        vFlagBuscar = 1
        txtLegEmple.SetFocus
    Else
        A = MsgBox("¿Seguro desea Guardar las Modificaciones Realizadas?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
       
        '1 es Ok 2 es Cancel
            If A = 1 Then
                With tEmpleados
                    .Edit
               'Datos Empleados -------------------------------------------------------------------
                    !Legajo = txtLegEmple.Text
                    !nombre = Format(txtNombreEmple.Text, ">")
                    !Cuil = txtCUILEmple.Text
                    !FechaIngreso = txtFechaIngEmple.Text
                    !Comision = txtComsionEmple.Text
                    !IDPuesto = txtPuestoEmple.Text
                    !Observaciones = Format(txtObservaciones.Text, ">")
                    
               'Datos Personales  -------------------------------------------------------------------
                    !Domicilio = Format(txtDomicilio.Text, ">")
                    !Localidad = Format(cmbLocalidad.Text, ">")
                    !DNI = txtDniEmple.Text
                    !Pais = Format(cmbPais.Text, ">")
                    !Prov = Format(cmbProv.Text, ">")
                    !Localidad = Format(cmbLocalidad.Text, ">")
                    !CP = txtCP.Text
                    !Tel = txtTel.Text
                    !Cel = txtCel.Text
                    !emaill = Format(txtEmail.Text, ">")
                    !FechaNacimiento = txtFechaNacEmple.Text
                    
                    .Update
                End With
                
                Call EnabledTextBox(FormEmpleados)
                Call LimpiarPantalla
                
                'tUltimosNumeros.MoveLast
                'txtIDCliente.Text = tUltimosNumeros!UltimoNumero + 1
                'txtIDCliente.BackColor = Me.BackColor
                'txtIDCliente.Enabled = False
                'txtRazonSocial.SetFocus
                'vFlagBuscar = 0
                
             Else
                If A = 2 Then
                    b = MsgBox("¿Limpia Pantalla?", vbQuestion + vbYesNo, "INFO DEL SISTEMA")
                    If b = 6 Then
                        Call LimpiarPantalla
                     Else
                        Call EnabledTextBox(FormEmpleados)
                        txtLegEmple.SetFocus
                    End If
                End If
            
            End If
        End If

End Sub


Private Sub btnPrimero_Click()

    On Error GoTo CapturaErrores
    
    If Not tEmpleados.EOF Then
        tEmpleados.MoveFirst
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
     Else
        vSQL = "SELECT * FROM Empleados ORDER BY Legajo"
        'MsgBox (vsql)
        Set tEmpleados = BaseSPC.OpenRecordset(vSQL)
        
        If Not tEmpleados.EOF Then
            tEmpleados.MoveFirst
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select


End Sub

Private Sub btnSalir_Click()

    Unload Me

End Sub

Private Sub btnUltimo_Click()

    On Error GoTo CapturaErrores
    
    If Not tEmpleados.EOF Then
        tEmpleados.MoveLast
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
     Else
        vSQL = "SELECT * FROM Empleados ORDER BY Legajos"
        MsgBox (vSQL)
        Set tEmpleados = BaseSPC.OpenRecordset(vSQL)
        
        If Not tEmpleados.EOF Then
            tEmpleados.Last
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
         Else
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
        End If
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay más registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select


End Sub

Private Sub cmbCondicionIva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

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

   ' If KeyAscii = 13 Then
   '         KeyAscii = 0
   '         SendKeys "{TAB}"
   ' End If
   
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

'    If KeyAscii = 13 Then
'            KeyAscii = 0
'            SendKeys "{TAB}"
'    End If

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


Private Sub cmbVendedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub Form_Load()
   
   'Seteo tamaño y ubicacion del form
        FormEmpleados.Height = 8460
        FormEmpleados.Width = 14205
        FormEmpleados.Top = 1000
        FormEmpleados.Left = 1000
        
        btnModificar.Enabled = False
        btnEliminar.Enabled = False
        
    'Abro Base de Datos
        'Seteo la captura de errores de no hay registros en el archivo
         On Error GoTo CapturaErrores
        
        Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        
        'Tabla Clientes
        '    Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
            
        'Tabla Empleados
            Set tEmpleados = BaseSPC.OpenRecordset("Empleados", dbOpenTable)
        
        'Tabla Países
            Set tPaises = BaseSPC.OpenRecordset("Paises", dbOpenTable)
        
            tPaises.Index = "PrimaryKey"
            tPaises.MoveFirst
        
            'Lleno combo de países
                While Not tPaises.EOF
                    cmbPais.AddItem tPaises!Descripcion
                    tPaises.MoveNext
                Wend
        
        'Tabla Condiciones de IVA
            Set tCondicionIVA = BaseSPC.OpenRecordset("CondicionIva", dbOpenTable)
        
          'Lleno Combo de condiciones de iva
            While Not tCondicionIVA.EOF
                cmbCondicionIva.AddItem tCondicionIVA!Descripcion
                tCondicionIVA.MoveNext
            Wend
        
        'Tabla Vendedores
            Set tVendedores = BaseSPC.OpenRecordset("Vendedores", dbOpenTable)
          
          'Lleno Combo de vendedores
            While Not tVendedores.EOF
                cmbVendedor.AddItem tVendedores!nombre
                tVendedores.MoveNext
            Wend
        
        'Tabla Ultimos Numeros de Clientes
            Set tUltimosNumeros = BaseSPC.OpenRecordset("UltimosNumeros", dbOpenTable)
            
            tUltimosNumeros.Index = "PrimaryKey"
            
            tUltimosNumeros.Seek "=", "tEmpleados"
            txtIDCliente.Text = tUltimosNumeros!UltimoNumero + 1
            txtIDCliente.BackColor = Me.BackColor
            txtIDCliente.Enabled = False
            
        'Seteo variable bandera de busqueda
            vFlagBuscar = 0
       
CapturaErrores:
    
  'Captura el error de no hay registros en la tabla
    Select Case Err
        Case 3021
            Resume Next
    End Select
        
End Sub


Public Function Autocompletar_Combo(Combo As ComboBox)
  
 Dim i As Integer, posSelect As Integer
  
    Select Case (KeyRetroceso Or Len(Combo.Text) = 0)
        Case True
            KeyRetroceso = False
            Exit Function
    End Select
  
    With Combo
  
    'Recorremos todos los elementos del combo
    For i = 0 To .ListCount - 1
        'Si hay coincidencia
        If InStr(1, .List(i), .Text, vbTextCompare) = 1 Then
            posSelect = .SelStart
            'Mostramos el texto en el combo
            .Text = .List(i)
            'Indicamos el comienzo de la selección
            .SelStart = posSelect
            'Acá seleccionamos el texto
            .SelLength = Len(.Text) - posSelect
  
            Exit For
        End If
    Next i
  
    End With

End Function


Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
  txtFechaIngEmple = MonthView1.Value
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
     txtFechaNacEmple = MonthView2.Value
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


Private Sub txtCP2_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub txtCUIT_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)
    
End Sub


Private Sub txtCUIT_LostFocus()

    If txtCUIT.Text <> "" Then
        If Len(txtCUIT.Text) = 11 Then
            txtCUIT.Text = Left(txtCUIT.Text, 2) + "-" + Mid(txtCUIT.Text, 3, 8) + "-" + Right(txtCUIT.Text, 1)
         Else
            MsgBox "Error en Nro de CUIT", vbCritical, "ERROR"
        End If
    Else
        MsgBox "Error en Nro de CUIT", vbCritical, "ERROR"
    End If
    

End Sub


Private Sub txtComsionEmple_GotFocus()
    txtComsionEmple.SelLength = Len(txtComsionEmple.Text)
End Sub

Private Sub txtComsionEmple_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
KeyAscii = Verificar_Tecla(KeyAscii)
End Sub



Private Sub txtCP_GotFocus()
    txtCP.SelLength = Len(txtCP.Text)
End Sub

Private Sub txtCUILEmple_GotFocus()
    txtCUILEmple.SelLength = Len(txtCUILEmple.Text)
End Sub

Private Sub txtCUILEmple_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)
End Sub

Private Sub txtCUILEmple_LostFocus()
    If txtCUILEmple.Text <> "" Then
        If Len(txtCUILEmple.Text) = 11 Then
            txtCUILEmple.Text = Left(txtCUILEmple.Text, 2) + "-" + Mid(txtCUILEmple.Text, 3, 8) + "-" + Right(txtCUILEmple.Text, 1)
         Else
            MsgBox "Nro de CUIL Mal Ingresado", vbCritical, "ERROR"
        End If
    Else
        MsgBox "Error en Nro de CUIL", vbCritical, "ERROR"
    End If
End Sub

Private Sub txtDniEmple_GotFocus()
    txtDniEmple.SelLength = Len(txtDniEmple.Text)
End Sub

Private Sub txtDniEmple_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
    
    KeyAscii = Verificar_TeclaDNI(KeyAscii)
    
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

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub

Private Sub txtLimiteCredito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)

End Sub

Private Sub txtLimiteCredito_LostFocus()

    txtLimiteCredito.Text = Format(txtLimiteCredito.Text, "Standard")

End Sub

Private Sub txtNombreFantasia_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub

Private Sub txtFechaIngEmple_GotFocus()
    txtFechaIngEmple.SelLength = Len(txtFechaIngEmple.Text)
End Sub

Private Sub txtFechaIngEmple_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
End Sub

Private Sub txtFechaNacEmple_GotFocus()
    txtFechaNacEmple.SelLength = Len(txtFechaNacEmple.Text)
End Sub

Private Sub txtFechaNacEmple_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
        End If
End Sub

Private Sub txtLegEmple_GotFocus()
    txtLegEmple.SelLength = Len(txtLegEmple.Text)
End Sub

Private Sub txtLegEmple_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
End Sub

Private Sub txtNombreEmple_GotFocus()
    txtNombreEmple.SelLength = Len(txtNombreEmple.Text)
End Sub

Private Sub txtNombreEmple_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
End Sub

Private Sub txtObservaciones_GotFocus()
    txtObservaciones.SelLength = Len(txtObservaciones.Text)
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
End Sub

Function Verificar_Tecla(Tecla_Presionada)
    
Dim Teclas As String
    
    'Acepta todos los números, la tecla Backspace, _
     la tecla Enter, la coma y el punto
    
    Teclas = "1234567890.," & Chr(vbKeyBack)
    
    If InStr(1, Teclas, Chr(Tecla_Presionada)) Then
        
        Verificar_Tecla = Tecla_Presionada
    Else
        ' Si no es ninguna de las indicadas retorna 0
        Verificar_Tecla = 0
    End If

End Function

Function Verificar_TeclaDNI(Tecla_Presionada)
    
Dim Teclas2 As String
    
    'Acepta todos los números, la tecla Backspace, _
     la tecla Enter
    
    Teclas2 = "1234567890" & Chr(vbKeyBack)
        If InStr(1, Teclas2, Chr(Tecla_Presionada)) Then
        
        Verificar_TeclaDNI = Tecla_Presionada
    Else
        ' Si no es ninguna de las indicadas retorna 0
        Verificar_TeclaDNI = 0
    End If
    
    

End Function

Private Sub txtPorcentajeDescuento_LostFocus()

    txtPorcentajeDescuento.Text = Format(txtPorcentajeDescuento.Text, "Standard")

End Sub

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub txtPuestoEmple_GotFocus()
    txtPuestoEmple.SelLength = Len(txtPuestoEmple.Text)
End Sub

Private Sub txtPuestoEmple_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
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


Private Sub txtZonaVenta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)

End Sub


