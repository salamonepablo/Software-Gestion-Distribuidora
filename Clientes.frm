VERSION 5.00
Begin VB.Form FormClientes 
   Caption         =   "ABM Clientes"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14895
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   14895
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Cliente"
      Height          =   7695
      Left            =   120
      TabIndex        =   20
      Top             =   120
      Width           =   14655
      Begin VB.Frame Frame5 
         Caption         =   "Acciones"
         Height          =   1215
         Left            =   480
         TabIndex        =   23
         Top             =   6120
         Width           =   13695
         Begin VB.CommandButton btnLimpiar 
            Caption         =   "&Nuevo"
            Height          =   615
            Left            =   5040
            TabIndex        =   0
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnSalir 
            Caption         =   "&Salir"
            Height          =   615
            Left            =   12240
            TabIndex        =   48
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnCtaCte 
            Caption         =   "&Cta Cte"
            Height          =   615
            Left            =   11040
            TabIndex        =   47
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnEliminar 
            Caption         =   "&Eliminar"
            Height          =   615
            Left            =   9840
            TabIndex        =   46
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnBuscar 
            Caption         =   "&Buscar"
            Height          =   615
            Left            =   8640
            TabIndex        =   45
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnModificar 
            Caption         =   "&Modificar"
            Height          =   615
            Left            =   7440
            TabIndex        =   44
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnGrabar 
            Caption         =   "&Grabar "
            Height          =   615
            Left            =   6240
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnUltimo 
            Caption         =   ">|"
            Height          =   615
            Left            =   3840
            TabIndex        =   43
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAdelante 
            Caption         =   ">>"
            Height          =   615
            Left            =   2640
            TabIndex        =   42
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAtras 
            Caption         =   "<<"
            Height          =   615
            Left            =   1440
            TabIndex        =   41
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnPrimero 
            Caption         =   "|<"
            Height          =   615
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Datos de Contacto / Facturación"
         Height          =   2175
         Left            =   480
         TabIndex        =   22
         Top             =   3720
         Width           =   13695
         Begin VB.CommandButton btnDireccionesEntrega 
            Caption         =   "&Direcciones de Entrega"
            Enabled         =   0   'False
            Height          =   495
            Left            =   11160
            TabIndex        =   50
            Top             =   1320
            Width           =   2175
         End
         Begin VB.ComboBox cmbLocalidad 
            Height          =   315
            Left            =   3840
            Sorted          =   -1  'True
            TabIndex        =   13
            Top             =   600
            Width           =   2775
         End
         Begin VB.TextBox txtEmail 
            Height          =   375
            Left            =   6840
            TabIndex        =   18
            Top             =   1440
            Width           =   3975
         End
         Begin VB.TextBox txtCel 
            Height          =   375
            Left            =   4320
            TabIndex        =   17
            Top             =   1440
            Width           =   2175
         End
         Begin VB.TextBox txtTel 
            Height          =   375
            Left            =   2040
            TabIndex        =   16
            Top             =   1440
            Width           =   1935
         End
         Begin VB.ComboBox cmbPais 
            Height          =   315
            Left            =   480
            TabIndex        =   11
            Top             =   600
            Width           =   1095
         End
         Begin VB.ComboBox cmbProv 
            Height          =   315
            Left            =   1800
            TabIndex        =   12
            Top             =   600
            Width           =   1815
         End
         Begin VB.TextBox txtCP 
            Height          =   375
            Left            =   480
            TabIndex        =   15
            Top             =   1440
            Width           =   1215
         End
         Begin VB.TextBox txtDomicilio 
            Height          =   375
            Left            =   6840
            TabIndex        =   14
            Top             =   600
            Width           =   6495
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
            TabIndex        =   49
            Top             =   360
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
            Left            =   6720
            TabIndex        =   39
            Top             =   1200
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
            Left            =   4200
            TabIndex        =   38
            Top             =   1200
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
            Left            =   1920
            TabIndex        =   37
            Top             =   1200
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
            Left            =   360
            TabIndex        =   36
            Top             =   360
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
            Left            =   1560
            TabIndex        =   35
            Top             =   360
            Width           =   810
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
            TabIndex        =   34
            Top             =   1200
            Width           =   375
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
            TabIndex        =   33
            Top             =   360
            Width           =   780
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Información Principal"
         Height          =   3015
         Left            =   480
         TabIndex        =   21
         Top             =   480
         Width           =   13695
         Begin VB.TextBox txtZonaVenta 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   1920
            TabIndex        =   8
            Top             =   2280
            Width           =   1095
         End
         Begin VB.ComboBox cmbVendedor 
            Height          =   315
            Left            =   3600
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   2280
            Width           =   3015
         End
         Begin VB.TextBox txtObservaciones 
            Height          =   1215
            Left            =   8880
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   10
            Top             =   1440
            Width           =   4455
         End
         Begin VB.TextBox txtLimiteCredito 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   7080
            TabIndex        =   7
            Top             =   2280
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.TextBox txtPorcentajeDescuento 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   360
            TabIndex        =   6
            Top             =   2280
            Width           =   1095
         End
         Begin VB.ComboBox cmbCondicionIva 
            Height          =   315
            Left            =   11160
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   600
            Width           =   2175
         End
         Begin VB.TextBox txtCUIT 
            Height          =   375
            Left            =   9480
            TabIndex        =   3
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtNombreFantasia 
            Height          =   375
            Left            =   360
            TabIndex        =   5
            Top             =   1440
            Width           =   8175
         End
         Begin VB.TextBox txtRazonSocial 
            Height          =   375
            Left            =   1560
            TabIndex        =   2
            Top             =   600
            Width           =   7575
         End
         Begin VB.TextBox txtIDCliente 
            CausesValidation=   0   'False
            Height          =   375
            HideSelection   =   0   'False
            Left            =   360
            TabIndex        =   1
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Límite de Crédito"
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
            Left            =   6960
            TabIndex        =   51
            Top             =   2040
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor"
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
            Left            =   3480
            TabIndex        =   32
            Top             =   2040
            Width           =   1305
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
            Left            =   8760
            TabIndex        =   31
            Top             =   1200
            Width           =   1275
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "% Descuento"
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
            Left            =   240
            TabIndex        =   30
            Top             =   2040
            Width           =   1125
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Zona de Venta"
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
            Left            =   1800
            TabIndex        =   29
            Top             =   2040
            Width           =   1275
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Condición ante IVA"
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
            Left            =   11160
            TabIndex        =   28
            Top             =   360
            Width           =   1635
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "CUIT"
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
            Left            =   9360
            TabIndex        =   27
            Top             =   360
            Width           =   450
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nombre Fantasía"
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
            Left            =   240
            TabIndex        =   26
            Top             =   1200
            Width           =   1455
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Razón Social"
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
            Left            =   1440
            TabIndex        =   25
            Top             =   360
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Código"
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
            Left            =   240
            TabIndex        =   24
            Top             =   360
            Width           =   585
         End
      End
   End
End
Attribute VB_Name = "FormClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private KeyRetroceso As Boolean


    
Private Sub LimpiarPantalla()

    txtIDCliente.text = ""
    txtRazonSocial.text = ""
    TxtCUIT.text = ""
    txtNombreFantasia.text = ""
    txtObservaciones.text = ""
    txtPorcentajeDescuento.text = ""
    txtLimiteCredito.text = ""
    txtZonaVenta.text = ""
    txtDomicilio.text = ""
    cmbLocalidad.text = ""
    txtCP.text = ""
    txtTel.text = ""
    txtCel.text = ""
    txtEmail.text = ""
    cmbPais.text = ""
    cmbProv.text = ""
    cmbLocalidad.text = ""
'    cmbCondicionIva = ""
'    cmbVendedor = ""
    btnGrabar.Enabled = True
    btnEliminar.Enabled = False
    btnModificar.Caption = "&Modificar"
    btnModificar.Enabled = False
        
    txtIDCliente.text = tUltimosNumeros!UltimoNumero + 1
    Call EnabledTextBox(FormClientes)
    txtIDCliente.SetFocus

End Sub

Private Sub Mostrar()

     On Error GoTo CapturaErrores
     
     With tClientes
            txtIDCliente.text = !IdCliente
            txtRazonSocial.text = !RazonSocial
            txtNombreFantasia.text = !NombreFantasia
            
            If !CUIT <> "" Then
                TxtCUIT.text = !CUIT
            End If
            
            If Not IsNull(!Domicilio) Then txtDomicilio.text = !Domicilio
            If Not IsNull(!localidad) Then cmbLocalidad.text = !localidad
            If Not IsNull(!Pais) Then cmbPais.text = !Pais
            If Not IsNull(!Prov) Then cmbProv.text = !Prov
            If Not IsNull(!localidad) Then cmbLocalidad.text = !localidad
            If Not IsNull(!CP) Then txtCP.text = !CP
            If Not IsNull(!Tel) Then txtTel.text = !Tel
            If Not IsNull(!Cel) Then txtCel.text = !Cel
            If Not IsNull(!email) Then txtEmail.text = !email
            If Not IsNull(!PorcentajeDescuento) Then txtPorcentajeDescuento.text = !PorcentajeDescuento
            'txtLimiteCredito.Text = !LimiteCredito
            If Not IsNull(!ZonaVenta) Then txtZonaVenta.text = !ZonaVenta
            If Not IsNull(!observaciones) Then txtObservaciones.text = !observaciones
            
            tCondicionIVA.Index = "PrimaryKey"
            tCondicionIVA.Seek "=", !condicionIva
            
            If Not tCondicionIVA.NoMatch Then
                cmbCondicionIva.text = tCondicionIVA!Descripcion
            End If
            
            tEmpleados.Index = "PrimaryKey"
            tEmpleados.Seek "=", !Vendedor
            
            If Not tEmpleados.NoMatch Then
                cmbVendedor.text = tEmpleados!Nombre
            End If
                                    
            vFlagBuscar = 0
            
     End With
     
     Call DisabledTextBox(FormClientes)
     
     If txtRazonSocial.text <> "" Then btnDireccionesEntrega.Enabled = True
     
CapturaErrores:
    Select Case Err
        Case 94
            Resume Next
    End Select

End Sub

Private Sub btnAdelante_Click()

  '  On Error GoTo CapturaErrores
    
    If Not tClientes.EOF Then
        tClientes.MoveNext
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
    Else
        MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
        tClientes.MoveLast
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            'MsgBox "Ultimo Registro", vbDefaultButton1, "INFO DEL SISTEMA"
            MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
            tClientes.MoveLast
            Call Mostrar
            Resume Next
    End Select

End Sub

Private Sub btnAtras_Click()
    
    On Error GoTo CapturaErrores
    
    If Not tClientes.BOF Then
        tClientes.MovePrevious
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
    Else
        MsgBox "Primer Registro", vbInformation, "INFO DEL SISTEMA"
        tClientes.MoveFirst
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "Primer Registro", vbInformation + vbOKOnly, "INFO DEL SISTEMA"
            'MsgBox "No hay registros !!!", vbDefaultButton1, "INFO DEL SISTEMA"
            tClientes.MoveFirst
            Call Mostrar
            Resume Next
    End Select

End Sub


Private Sub btnBuscar_Click()

    If vFlagBuscar = 0 Then
        vFlagBuscar = 1
        txtIDCliente.Enabled = True
        txtIDCliente.text = ""
        txtIDCliente.SetFocus
     Else
        
        If txtIDCliente.text <> "" Then
            Campo = "IDCliente= "
            Valor = txtIDCliente.text
         Else
            If txtRazonSocial.text <> "" Then
                Campo = "RazonSocial Like "
                Valor = "'*" + txtRazonSocial.text + "*'"
             Else
                If txtNombreFantasia.text <> "" Then
                    Campo = "NombreFantasia Like "
                    Valor = "'*" + txtNombreFantasia.text + "*'"
                 Else
                    A = MsgBox("DEBE INGRESAR UN VALOR DE BUSQUEDA", vbCritical, "ERROR !!!")
                    txtIDCliente.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        'vSQL = "SELECT IDProv, Descripcion FROM Provincias Where IDPais=" & tPaises!IDPais & " ORDER BY Descripcion"
        vSQL = "SELECT * FROM Clientes WHERE " & Campo & Valor & " ORDER BY IDCliente"
        
        'MsgBox (vSQL)
        
        Set tClientes = BaseSPC.OpenRecordset(vSQL)
    
        If Not tClientes.NoMatch Then
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

Public Sub btnDireccionesEntrega_Click()
    
    FormDireccionesEntrega.Show

End Sub

Private Sub btnEliminar_Click()

    'Cuando esten listas las facturas y los presupuestos
    'hay que controlar que no existan movimientos con el cliente.
    ' de existir los clientes hay que marcarlo como baja en el campo dado de baja.
    
    On Error GoTo CapturaErrores
    
    A = MsgBox("¿ Confirma Eliminar el Registro ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
    If A = 1 Then
        With tClientes
            .Delete
        End With
    End If
    
    Call EnabledTextBox(FormDireccionesEntrega)
    Call LimpiarPantalla

CapturaErrores:
    Select Case Err
        Case 3200
            E = MsgBox("EL CLIENTE TIENE REGISTROS RELACIONADOS, NO PUEDE ELMINARSE", vbCritical, "ALERTA DEL SISTEMA")
            txtRazonSocial.SetFocus
    End Select
End Sub

Private Sub btnGrabar_Click()

    A = MsgBox("¿ Seguro Genera Nuevo Registro ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
  If A = 1 Then
    
'    tClientes.Index = "PrimaryKey"
'     On Error GoTo CapturaErrores
    
     With tClientes
        .AddNew
            !IdCliente = txtIDCliente.text
            !RazonSocial = Format(txtRazonSocial.text, ">")
            !NombreFantasia = Format(txtNombreFantasia.text, ">")
            
            If Len(TxtCUIT.text) = 13 Then
                CUIT = Left(TxtCUIT.text, 2) + Mid(TxtCUIT.text, 4, 8) + Right(TxtCUIT.text, 1)
              Else
                CUIT = TxtCUIT.text
            End If
            'MsgBox (Cuit)
            !CUIT = CUIT
            'txtCUIT.Text
            !Domicilio = Format(txtDomicilio.text, ">")
            !localidad = Format(cmbLocalidad.text, ">")
            !Pais = Format(cmbPais.text, ">")
            !Prov = Format(cmbProv.text, ">")
            !localidad = Format(cmbLocalidad.text, ">")
            !CP = txtCP.text
            !Tel = txtTel.text
            !Cel = txtCel.text
            !email = Format(txtEmail.text, ">")
            If txtPorcentajeDescuento.text = "" Then txtPorcentajeDescuento.text = 0
            !PorcentajeDescuento = txtPorcentajeDescuento.text
            '!LimiteCredito = txtLimiteCredito.Text
            !ZonaVenta = txtZonaVenta.text
            !observaciones = Format(txtObservaciones.text, ">")
            
            tCondicionIVA.Index = "IndiceDescripcion"
            tCondicionIVA.Seek "=", cmbCondicionIva.text
            
            If Not tCondicionIVA.NoMatch Then !condicionIva = tCondicionIVA!IdCondicionIVA
            
            tEmpleados.Index = "IndiceNombre"
            tEmpleados.Seek "=", cmbVendedor.text
            
            If Not tEmpleados.NoMatch Then !Vendedor = tEmpleados!Legajo
            
        .Update
        
        'Grabo el Domicilio
            Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
            
            With tDomiciliosClientes
               .AddNew
                   !IdCliente = txtIDCliente.text
                   !item = 1
                   'lblItem.Caption
                   !Domicilio = Format(txtDomicilio.text, ">")
                   !Pais = Format(cmbPais.text, ">")
                   !Prov = Format(cmbProv.text, ">")
                   !localidad = Format(cmbLocalidad.text, ">")
                   !CP = txtCP.text
                   !Tel = txtTel.text
                   !Cel = txtCel.text
                   !email = Format(txtEmail.text, ">")
               .Update
            End With
     
     End With
        
    Set tCtaCte = BaseSPC.OpenRecordset("CtaCte", dbOpenTable)
        tCtaCte.AddNew
            tCtaCte!IdCliente = txtIDCliente.text
            tCtaCte!SaldoL1 = 0
            tCtaCte!SaldoL2 = 0
            tCtaCte!SaldoTotal = 0
            tCtaCte!FechaActSaldo = Format(Date, "dd/mm/yyyy")
        tCtaCte.Update
        
        tCtaCte.Close
    
    If A = 1 Then
       Call LimpiarPantalla
      'Lleno Combo de condiciones de iva
        tCondicionIVA.MoveFirst
        While Not tCondicionIVA.EOF
            cmbCondicionIva.AddItem tCondicionIVA!Descripcion
            tCondicionIVA.MoveNext
        Wend
      
      'Lleno Combo de vendedores
        tEmpleados.MoveFirst
        While Not tEmpleados.EOF
            cmbVendedor.AddItem tEmpleados!Nombre
            tEmpleados.MoveNext
        Wend
     Else
       txtRazonSocial.SetFocus
    End If
     
    tUltimosNumeros.Index = "PrimaryKey"
    tUltimosNumeros.Seek "=", "tClientes"
    
    tUltimosNumeros.Edit
        tUltimosNumeros!UltimoNumero = txtIDCliente.text
    tUltimosNumeros.Update
    
    txtIDCliente.text = tUltimosNumeros!UltimoNumero + 1
    txtRazonSocial.SetFocus

  Else
    txtRazonSocial.SetFocus
End If

CapturaErrores:
    Select Case Err
        Case 3021
            Resume Next
        Case 3421
            A = MsgBox("NO SE PUEDE GRABAR EL REGISTRO POR FALTA DE DATOS O DATOS INCORRECTOS", vbCritical, "ERROR !!")
            txtRazonSocial.SetFocus
            Exit Sub
    End Select

End Sub


Private Sub btnLimpiar_Click()

    Call LimpiarPantalla

End Sub


Private Sub btnLimpiar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub btnModificar_Click()

    If vFlagBuscar = 0 Then
        Call EnabledTextBox(FormClientes)
        btnModificar.Caption = "Guardar Cambios"
        btnGrabar.Enabled = False
        vFlagBuscar = 1
        txtRazonSocial.SetFocus
        txtIDCliente.Enabled = False
         
    Else
        A = MsgBox("¿Seguro desea Guardar las Modificaciones Realizadas?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
        
        '1 es Ok 2 es Cancel
            If A = 1 Then
                With tClientes
                    .Edit
                        !IdCliente = txtIDCliente.text
                        !RazonSocial = Format(txtRazonSocial.text, ">")
                        !NombreFantasia = Format(txtNombreFantasia.text, ">")
                        
                        If Len(TxtCUIT.text) = 13 Then CUIT = Left(TxtCUIT.text, 2) + Mid(TxtCUIT.text, 4, 8) + Right(TxtCUIT.text, 1)
                        If Len(TxtCUIT.text) = 11 Then CUIT = TxtCUIT.text
                        If Len(TxtCUIT.text) = 8 Then CUIT = TxtCUIT.text
                        !CUIT = CUIT
                        'txtCUIT.Text
                        '!CUIT = txtCUIT.Text
                        
                        !Domicilio = Format(txtDomicilio.text, ">")
                        !localidad = Format(cmbLocalidad.text, ">")
                        !Pais = Format(cmbPais.text, ">")
                        !Prov = Format(cmbProv.text, ">")
                        !localidad = Format(cmbLocalidad.text, ">")
                        !CP = txtCP.text
                        !Tel = txtTel.text
                        !Cel = txtCel.text
                        !email = Format(txtEmail.text, ">")
                        !PorcentajeDescuento = txtPorcentajeDescuento.text
                        '!LimiteCredito = txtLimiteCredito.Text
                        If (txtZonaVenta.text = "") Then txtZonaVenta.text = 0
                        !ZonaVenta = txtZonaVenta.text
                        !observaciones = Format(txtObservaciones.text, ">")
                        
                        tCondicionIVA.Index = "IndiceDescripcion"
                        tCondicionIVA.Seek "=", cmbCondicionIva.text
                        
                        If Not tCondicionIVA.NoMatch Then !condicionIva = tCondicionIVA!IdCondicionIVA
                        
                        tEmpleados.Index = "IndiceNombre"
                        tEmpleados.Seek "=", cmbVendedor.text
                        
                        If Not tEmpleados.NoMatch Then !Vendedor = tEmpleados!Legajo
                    .Update
                End With
                
                Call EnabledTextBox(FormClientes)
                Call LimpiarPantalla
                btnGrabar.Enabled = True
                
                'tUltimosNumeros.MoveLast
                txtIDCliente.text = tUltimosNumeros!UltimoNumero + 1
                txtIDCliente.BackColor = Me.BackColor
                txtIDCliente.Enabled = False
                txtRazonSocial.SetFocus
                vFlagBuscar = 0
                
             Else
                If A = 2 Then
                    b = MsgBox("¿Limpia Pantalla?", vbQuestion + vbYesNo, "INFO DEL SISTEMA")
                    If b = 6 Then
                        Call LimpiarPantalla
                     Else
                        Call EnabledTextBox(FormClientes)
                        txtIDCliente.SetFocus
                    End If
                End If
            
            End If
        End If

End Sub


Private Sub btnPrimero_Click()

    On Error GoTo CapturaErrores
    
    If Not tClientes.EOF Then
        tClientes.MoveFirst
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
     Else
        vSQL = "SELECT * FROM Clientes ORDER BY IDCliente"
        'MsgBox (vsql)
        Set tClientes = BaseSPC.OpenRecordset(vSQL)
        
        If Not tClientes.EOF Then
            tClientes.MoveFirst
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
    
    If Not tClientes.EOF Then
        tClientes.MoveLast
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
     Else
        vSQL = "SELECT * FROM Clientes ORDER BY IDCliente"
        'MsgBox (vsql)
        Set tClientes = BaseSPC.OpenRecordset(vSQL)
        
        If Not tClientes.EOF Then
            tClientes.Last
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

Private Sub cmbCondicionIva_GotFocus()
'    cmbCondicionIva.SelLength = Len(cmbCondicionIva.Text)
End Sub

Private Sub cmbCondicionIva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub

Private Sub cmbLocalidad_Change()
    
    Autocompletar_Combo cmbLocalidad

End Sub

Private Sub cmbLocalidad_GotFocus()
    cmbLocalidad.SelLength = Len(cmbLocalidad.text)
End Sub

Private Sub cmbLocalidad_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete
        ', vbKeyReturn
            Select Case Len(cmbLocalidad.text)
                Case Is <> 0
                    KeyRetroceso = True
              End Select
    End Select

End Sub


Private Sub cmbLocalidad_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
    'Si le pasamos a SendMessageLong el valor False lo cierra
        resp = SendMessageLong(cmbLocalidad.hwnd, &H14F, False, 0)
        Sendkeys "{TAB}"
     Else
    'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
    'presionamos una tecla diferente al Enter
        resp = SendMessageLong(cmbLocalidad.hwnd, &H14F, True, 0)
    End If

End Sub

Private Sub cmbLocalidad_LostFocus()

  If cmbLocalidad.text <> "" Then
    tLocalidades.MoveLast
    tLocalidades.FindFirst tLocalidades.IDLocalidad = cmbLocalidad.text
   
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
  End If

End Sub


Private Sub cmbPais_Change()
    
    Autocompletar_Combo cmbPais

End Sub

Private Sub cmbPais_GotFocus()
    cmbPais.SelLength = Len(cmbPais.text)
End Sub

Private Sub cmbPais_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete
        ', vbKeyReturn
            Select Case Len(cmbPais.text)
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
        Sendkeys "{TAB}"
     Else
    'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
    'presionamos una tecla diferente al Enter
        resp = SendMessageLong(cmbPais.hwnd, &H14F, True, 0)
    End If

End Sub


Private Sub cmbPais_LostFocus()

    tPaises.Index = "IndiceDescripcion"
    tPaises.Seek "=", cmbPais.text
    
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
    cmbProv.SelLength = Len(cmbProv.text)
End Sub

Private Sub cmbProv_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        'Si la tecla presionada es Backspace o la tecla Delete
        Case vbKeyBack, vbKeyDelete
            Select Case Len(cmbProv.text)
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
        Sendkeys "{TAB}"
     Else
    'si le pasamos True a SendMessageLong lo adespliega, es decir cuando
    'presionamos una tecla diferente al Enter
        resp = SendMessageLong(cmbProv.hwnd, &H14F, True, 0)
    End If

End Sub


Private Sub cmbProv_LostFocus()

    On Error GoTo CapturaErrores
    
    tProvincias.MoveLast
    
    strProvincia = "Descripcion = '" & cmbProv.text & "'"

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

CapturaErrores:

Select Case Err
    Case 3021
        Resume Next
End Select



End Sub


Private Sub cmbVendedor_GotFocus()
'    cmbVendedor.SelLength = Len(cmbVendedor.Text)
End Sub

Private Sub cmbVendedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub cmbVendedor_LostFocus()

    If cmbVendedor.text = "" Then
        A = MsgBox("NO SE PUEDE DEJAR EL CAMPO VENDEDOR SIN VALOR", vbCritical, "ERROR !!")
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()
   
   'Seteo tamaño y ubicacion del form
        FormClientes.Height = 8505
        FormClientes.Width = 15180
        FormClientes.Top = 1000
        FormClientes.Left = 1000
        
        btnModificar.Enabled = False
        btnEliminar.Enabled = False
        
    'Abro Base de Datos
        'Seteo la captura de errores de no hay registros en el archivo
         On Error GoTo CapturaErrores
        
        Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        
        'Tabla Clientes
            Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
        
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
            Set tEmpleados = BaseSPC.OpenRecordset("Empleados", dbOpenTable)
          
          'Lleno Combo de vendedores
            While Not tEmpleados.EOF
                If tEmpleados!IDPuesto = 1 Then cmbVendedor.AddItem tEmpleados!Nombre
                tEmpleados.MoveNext
            Wend
        
        'Tabla Ultimos Numeros de Clientes
            Set tUltimosNumeros = BaseSPC.OpenRecordset("UltimosNumeros", dbOpenTable)
            
            tUltimosNumeros.Index = "PrimaryKey"
            
            tUltimosNumeros.Seek "=", "tClientes"
            txtIDCliente.text = tUltimosNumeros!UltimoNumero + 1
            txtIDCliente.BackColor = Me.BackColor
            'txtIDCliente.Enabled = False
            
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
  
 Dim I As Integer, posSelect As Integer
  
    Select Case (KeyRetroceso Or Len(Combo.text) = 0)
        Case True
            KeyRetroceso = False
            Exit Function
    End Select
  
    With Combo
  
    'Recorremos todos los elementos del combo
    For I = 0 To .ListCount - 1
        'Si hay coincidencia
        If InStr(1, .List(I), .text, vbTextCompare) = 1 Then
            posSelect = .SelStart
            'Mostramos el texto en el combo
            .text = .List(I)
            'Indicamos el comienzo de la selección
            .SelStart = posSelect
            'Acá seleccionamos el texto
            .SelLength = Len(.text) - posSelect
  
            Exit For
        End If
        
    Next I
    
    End With
    
    KeyRetroceso = False

End Function

Private Sub txtCel_GotFocus()
    txtCel.SelLength = Len(txtCel.text)
End Sub

Private Sub txtCel_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If
    
     KeyAscii = Verificar_Tecla(KeyAscii)

End Sub


Private Sub txtCP_GotFocus()
    txtCP.SelLength = Len(txtCP.text)
End Sub

Private Sub txtCP_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub txtCUIT_GotFocus()
    TxtCUIT.SelLength = Len(TxtCUIT.text)
End Sub

Private Sub txtCUIT_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)
    
End Sub

Private Sub txtCUIT_LostFocus()

    If TxtCUIT.text <> "" Then
        If Len(TxtCUIT.text) = 11 Then
            TxtCUIT.text = Left(TxtCUIT.text, 2) + "-" + Mid(TxtCUIT.text, 3, 8) + "-" + Right(TxtCUIT.text, 1)
         Else
            MsgBox "Error en Nro de CUIT", vbCritical, "ERROR"
        End If
    Else
        MsgBox "Error en Nro de CUIT", vbCritical, "ERROR"
    End If
    
End Sub


Private Sub txtDomicilio_GotFocus()
    txtDomicilio.SelLength = Len(txtDomicilio.text)
End Sub

Private Sub txtDomicilio_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub txtEmail_GotFocus()
    txtEmail.SelLength = Len(txtEmail.text)
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbLowerCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub

Private Sub txtIDCliente_GotFocus()
    txtIDCliente.SelLength = Len(txtIDCliente.text)
End Sub

Private Sub txtIDCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub txtIDCliente_LostFocus()

    On Error GoTo CapturaErrores
    
    If vFlagBuscar = 1 Then
   '     vFlagBuscar = 1
   '     txtIDCliente.Enabled = True
   '     txtIDCliente.Text = ""
   '     txtIDCliente.SetFocus
   '  Else
        
        If txtIDCliente.text <> "" Then
            Campo = "IDCliente= "
            Valor = txtIDCliente.text
         Else
            If txtRazonSocial.text <> "" Then
                Campo = "RazonSocial Like "
                Valor = "'*" + txtRazonSocial.text + "*'"
             Else
                If txtNombreFantasia.text <> "" Then
                    Campo = "NombreFantasia Like "
                    Valor = "'*" + txtNombreFantasia.text + "*'"
                 Else
                    A = MsgBox("DEBE INGRESAR UN VALOR DE BUSQUEDA", vbCritical, "ERROR !!!")
                    txtIDCliente.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        'vSQL = "SELECT IDProv, Descripcion FROM Provincias Where IDPais=" & tPaises!IDPais & " ORDER BY Descripcion"
        vSQL = "SELECT * FROM Clientes WHERE " & Campo & Valor & " ORDER BY IDCliente"
        
'        MsgBox (vSQL)
        
        Set tClientes = BaseSPC.OpenRecordset(vSQL)
    
        If Not tClientes.EOF Then
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
        Else
            MsgBox "No se encuentran registros", vbCritical, "ERROR"
            txtIDCliente.SetFocus
        End If
        
        'vFlagBuscar = 0
        
    Else
                
        'Si doy de alta un cliente nuevo, controlo que no exista
        If txtIDCliente.text <> "" Then
            Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
            tClientes.Index = "PrimaryKey"
            tClientes.Seek "=", txtIDCliente.text
        
            If Not tClientes.NoMatch Then
                MsgBox "Nro de Cliente Ya Existe", vbCritical, "ERROR"
                txtIDCliente.SetFocus
            End If
            'tClientes.Close
        End If
    
    End If

CapturaErrores:

    Select Case Err
        Case 3022
            MsgBox "No se encuentran registros", vbCritical, "ERROR"
            Exit Sub
    End Select

End Sub


Private Sub txtLimiteCredito_GotFocus()
    txtLimiteCredito.SelLength = Len(txtLimiteCredito.text)
End Sub

Private Sub txtLimiteCredito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)

End Sub




Private Sub txtLimiteCredito_LostFocus()

    txtLimiteCredito.text = Format(txtLimiteCredito.text, "Standard")

End Sub

Private Sub txtNombreFantasia_GotFocus()
    txtNombreFantasia.SelLength = Len(txtNombreFantasia.text)
End Sub

Private Sub txtNombreFantasia_KeyPress(KeyAscii As Integer)
    
    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub txtObservaciones_GotFocus()
    txtObservaciones.SelLength = Len(txtObservaciones.text)
End Sub

Private Sub txtObservaciones_KeyPress(KeyAscii As Integer)

    'If KeyAscii = 13 Then
    '        KeyAscii = 0
    '        SendKeys "{TAB}"
    'End If

End Sub


Private Sub txtPorcentajeDescuento_GotFocus()
    txtPorcentajeDescuento.SelLength = Len(txtPorcentajeDescuento.text)
End Sub

Private Sub txtPorcentajeDescuento_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)

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

Private Sub txtPorcentajeDescuento_LostFocus()

    txtPorcentajeDescuento.text = Format(txtPorcentajeDescuento.text, "Standard")

End Sub



Private Sub txtRazonSocial_GotFocus()
    txtRazonSocial.SelLength = Len(txtRazonSocial.text)
End Sub

Private Sub txtRazonSocial_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub txtRazonSocial_LostFocus()

    If txtRazonSocial.text <> "" Then
        btnDireccionesEntrega.Enabled = True
        btnGrabar.Enabled = True
    End If
End Sub


Private Sub txtTel_GotFocus()
    txtTel.SelLength = Len(txtTel.text)
End Sub

Private Sub txtTel_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If
    
     KeyAscii = Verificar_Tecla(KeyAscii)

End Sub


Private Sub txtZonaVenta_GotFocus()
   txtZonaVenta.SelLength = Len(txtZonaVenta.text)
End Sub

Private Sub txtZonaVenta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)

End Sub


Private Sub txtZonaVenta_LostFocus()

    If txtZonaVenta.text = "" Then
        A = MsgBox("NO SE PUEDE DEJAR EL CAMPO ZONA DE VENTA SIN VALOR", vbCritical, "ERROR !!")
    End If
    
End Sub
