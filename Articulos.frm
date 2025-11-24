VERSION 5.00
Begin VB.Form FormProductos 
   Caption         =   "MAESTRO DE PRODUCTOS"
   ClientHeight    =   6480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14250
   LinkTopic       =   "Form1"
   ScaleHeight     =   6480
   ScaleWidth      =   14250
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FormArticulos 
      Height          =   6255
      Left            =   120
      TabIndex        =   13
      Top             =   0
      Width           =   14055
      Begin VB.Frame Frame5 
         Caption         =   "Acciones"
         Height          =   1215
         Left            =   360
         TabIndex        =   26
         Top             =   4800
         Width           =   13455
         Begin VB.CommandButton btnPrimero 
            Caption         =   "|<"
            Height          =   615
            Left            =   240
            TabIndex        =   35
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAtras 
            Caption         =   "<<"
            Height          =   615
            Left            =   1560
            TabIndex        =   34
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnAdelante 
            Caption         =   ">>"
            Height          =   615
            Left            =   2880
            TabIndex        =   33
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnUltimo 
            Caption         =   ">|"
            Height          =   615
            Left            =   4200
            TabIndex        =   32
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnGrabar 
            Caption         =   "&Grabar"
            Height          =   615
            Left            =   6840
            TabIndex        =   11
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnModificar 
            Caption         =   "&Modificar"
            Height          =   615
            Left            =   8160
            TabIndex        =   31
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnBuscar 
            Caption         =   "&Buscar"
            Height          =   615
            Left            =   9600
            TabIndex        =   30
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnEliminar 
            Caption         =   "&Eliminar"
            Height          =   615
            Left            =   10800
            TabIndex        =   29
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnSalir 
            Caption         =   "&Salir"
            Height          =   615
            Left            =   12120
            TabIndex        =   28
            Top             =   360
            Width           =   975
         End
         Begin VB.CommandButton btnLimpiar 
            Caption         =   "&Nuevo"
            Height          =   615
            Left            =   5520
            TabIndex        =   27
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Observaciones / Descripción Ampliada"
         Height          =   2055
         Left            =   9480
         TabIndex        =   25
         Top             =   2520
         Width           =   4335
         Begin VB.TextBox txtObservaciones 
            CausesValidation=   0   'False
            Height          =   1455
            Left            =   240
            MultiLine       =   -1  'True
            ScrollBars      =   3  'Both
            TabIndex        =   9
            Top             =   360
            Width           =   3855
         End
      End
      Begin VB.Frame Frame3 
         Height          =   2055
         Left            =   360
         TabIndex        =   12
         Top             =   240
         Width           =   13455
         Begin VB.TextBox txtCodBarra 
            Height          =   375
            Left            =   6480
            TabIndex        =   10
            Top             =   600
            Width           =   4350
         End
         Begin VB.TextBox txtCodProd 
            Height          =   375
            Left            =   360
            TabIndex        =   0
            Top             =   1320
            Width           =   1335
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   375
            Left            =   2040
            TabIndex        =   1
            Top             =   1320
            Width           =   11055
         End
         Begin VB.Image Image1 
            Height          =   690
            Left            =   240
            Picture         =   "Articulos.frx":0000
            Stretch         =   -1  'True
            Top             =   240
            Width           =   2535
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Codigo de Barra"
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
            Left            =   4920
            TabIndex        =   24
            Top             =   720
            Width           =   1380
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
            TabIndex        =   23
            Top             =   1080
            Width           =   585
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Descricpion"
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
            TabIndex        =   22
            Top             =   1080
            Width           =   1020
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Datos Complementarios"
         Height          =   2055
         Left            =   5280
         TabIndex        =   19
         Top             =   2520
         Width           =   3855
         Begin VB.ComboBox cmbRubro 
            Height          =   315
            Left            =   360
            TabIndex        =   8
            Text            =   "cmbRubro"
            Top             =   1440
            Width           =   3135
         End
         Begin VB.CheckBox chkPromo 
            Caption         =   "Promoción"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   7
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtPuntoPedido 
            Height          =   375
            Left            =   3240
            TabIndex        =   6
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            Caption         =   "Rubro"
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
            TabIndex        =   21
            Top             =   1200
            Width           =   525
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Punto Pedido"
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
            TabIndex        =   20
            Top             =   360
            Visible         =   0   'False
            Width           =   1155
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Caracteristicas del Producto"
         Height          =   2055
         Left            =   360
         TabIndex        =   14
         Top             =   2520
         Width           =   4695
         Begin VB.ComboBox cmbUnidadMedida 
            Height          =   315
            Left            =   1560
            TabIndex        =   3
            Text            =   "cmbUnidadMedida"
            Top             =   720
            Width           =   1455
         End
         Begin VB.TextBox txtPrecioUnitarioPresupuesto 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   3000
            TabIndex        =   5
            Top             =   1440
            Width           =   1095
         End
         Begin VB.TextBox txtTamaño 
            Height          =   375
            Left            =   3480
            TabIndex        =   2
            Top             =   720
            Visible         =   0   'False
            Width           =   2175
         End
         Begin VB.TextBox txtPrecioUnitarioFactura 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   360
            TabIndex        =   4
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            Caption         =   "<<< $$$ >>>"
            Height          =   195
            Left            =   1800
            TabIndex        =   36
            Top             =   1560
            Width           =   900
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Precio Presupuesto"
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
            TabIndex        =   18
            Top             =   1200
            Width           =   1665
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tamaño / Tipo"
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
            TabIndex        =   17
            Top             =   360
            Visible         =   0   'False
            Width           =   1260
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Unidad de Medida"
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
            TabIndex        =   16
            Top             =   480
            Width           =   1560
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Precio Factura"
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
            TabIndex        =   15
            Top             =   1200
            Width           =   1260
         End
      End
   End
End
Attribute VB_Name = "FormProductos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Buscar()
        On Error GoTo Captura_Errores
        
        If txtCodProd.text <> "" Then
            Campo = "CodProd Like "
            Valor = "'*" + txtCodProd.text + "*'"
         Else
            If txtDescripcion.text <> "" Then
                Campo = "Descripcion Like "
                Valor = "'*" + txtDescripcion.text + "*'"
             Else
                If txtCodBarra.text <> "" Then
                    Campo = "CodBarra Like "
                    Valor = "'*" + txtCodBarra.text + "*'"
                Else
                    M = MsgBox("Ingrese un Criterio de Búsqueda", vbInformation, "INFO DEL SISTEMA")
                    txtCodProd.SetFocus
                    Exit Sub
                End If
            End If
        End If
        
        'vSQL = "SELECT IDProv, Descripcion FROM Provincias Where IDPais=" & tPaises!IDPais & " ORDER BY Descripcion"
        vSQL = "SELECT * FROM Productos WHERE " & Campo & Valor & " ORDER BY CodProd"
        
        'MsgBox (vSQL)
        
        Set tProductos = BaseSPC.OpenRecordset(vSQL)
    
        If Not tProductos.NoMatch Then
            Call Mostrar
            btnGrabar.Enabled = False
            btnModificar.Enabled = True
            btnEliminar.Enabled = True
        Else
            MsgBox "No se encuentran registros", vbCritical, "ERROR"
        End If
        
        vFlagBuscar = 0
        
   ' End If

Captura_Errores:
    Select Case Err
        Case 3021
            MsgBox "No se encuentran registros", vbCritical, "ERROR"
            Resume Next
    End Select

End Sub

Private Sub LimpiarPantalla()

    txtCodBarra.text = ""
    txtCodProd.text = ""
    txtDescripcion.text = ""
    txtTamaño.text = ""
    txtPrecioUnitarioFactura.text = ""
    txtPrecioUnitarioPresupuesto.text = ""
    txtPuntoPedido.text = ""
    txtObservaciones.text = ""
    
    
    Call EnabledTextBox(FormProductos)
    
    'Call ReseteaBbdd
    
    vFlagBuscar = 0
    txtCodProd.SetFocus
    

End Sub

Private Sub ReseteaBbdd()

       On Error GoTo CapturaErrores
        
      '  Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        
        'Tabla Productos
            tProductos.Close
            Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
        
        'Tabla Unidades de Medida
            tUnidadesMedida.Close
            Set tUnidadesMedida = BaseSPC.OpenRecordset("UnidadesMedida", dbOpenTable)
        
            tUnidadesMedida.Index = "PrimaryKey"
            tUnidadesMedida.MoveFirst
        
            'Lleno combo Unidades de Medida
                cmbUnidadMedida.Clear
                While Not tUnidadesMedida.EOF
                    cmbUnidadMedida.AddItem tUnidadesMedida!Descripcion
                    tUnidadesMedida.MoveNext
                Wend
        
        'Tabla Rubros
            tRubros.Close
            Set tRubros = BaseSPC.OpenRecordset("Rubros", dbOpenTable)
        
            tRubros.Index = "PrimaryKey"
            tRubros.MoveFirst
        
            'Lleno combo Unidades de Medida
                cmbRubro.Clear
                While Not tRubros.EOF
                    cmbRubro.AddItem tRubros!Descripcion
                    tRubros.MoveNext
                Wend
        
        
            vFlagBuscar = 0
       
CapturaErrores:
    
  'Captura el error de no hay registros en la tabla
    Select Case Err
        Case 3021
            Resume Next
    End Select


End Sub

Private Sub btnAdelante_Click()

    On Error GoTo CapturaErrores
    
    If Not tProductos.EOF Then
        tProductos.MoveNext
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
    Else
        MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
        tProductos.MoveLast
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            'MsgBox "Ultimo Registro", vbDefaultButton1, "INFO DEL SISTEMA"
            MsgBox "Ultimo Registro", vbInformation, "INFO DEL SISTEMA"
            tProductos.MoveLast
            Call Mostrar
            Resume Next
    End Select

End Sub

Private Sub btnAtras_Click()
    
    On Error GoTo CapturaErrores
    
    If Not tProductos.BOF Then
        tProductos.MovePrevious
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
    Else
        MsgBox "Primer Registro", vbInformation, "INFO DEL SISTEMA"
        tProductos.MoveFirst
        Call Mostrar
    End If

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "Primer Registro", vbInformation + vbOKOnly, "INFO DEL SISTEMA"
            'MsgBox "No hay registros !!!", vbDefaultButton1, "INFO DEL SISTEMA"
            tProductos.MoveFirst
            Call Mostrar
            Resume Next
    End Select

End Sub

Private Sub btnBuscar_Click()

    Call Buscar
    
End Sub


Private Sub btnEliminar_Click()
    
    'Cuando esten listas las facturas y los presupuestos
    'hay que controlar que no existan movimientos con el item.
    'de existir los productos hay que marcarlo como baja en el campo dado de baja.
    
    A = MsgBox("¿ Confirma Eliminar el Registro ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
    If A = 1 Then
        With tProductos
            .Delete
        End With
    End If
    
    Call LimpiarPantalla

End Sub

Private Sub btnGrabar_Click()

    A = MsgBox("¿ Seguro Genera Nuevo Registro ?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
    
    tProductos.Index = "PrimaryKey"
    
     With tProductos
        .AddNew
            !CodProd = txtCodProd.text
            !Descripcion = Format(txtDescripcion.text, ">")
            !Stock = 0
            '!PuntoPedido = txtPuntoPedido.Text
            '!Tamaño = txtTamaño.Text
            !UnidadMedida = cmbUnidadMedida.text
            !PrecioUnitarioFactura = txtPrecioUnitarioFactura.text
            !PrecioUnitarioPresupuesto = txtPrecioUnitarioPresupuesto.text
            !Rubro = Format(cmbRubro.text, ">")
            !codbarra = txtCodBarra.text
            !Promo = chkPromo.Value
            !observaciones = Format(txtObservaciones.text, ">")
                        
            If A = 1 Then
              Call LimpiarPantalla
              'Lleno Combo de condiciones de iva
              
            Else
               txtCodProd.SetFocus
            End If
            
        .Update
     
     End With
     
CapturaErrores:

    Select Case Err
        Case 3021
            Resume Next
    End Select


End Sub

Private Sub Mostrar()

     With tProductos
            
            txtCodProd.text = !CodProd
            txtDescripcion.text = !Descripcion
            A = !Stock
            
            'txtPuntoPedido.Text = !PuntoPedido
            'txtTamaño.Text = !Tamaño
            
            If cmbUnidadMedida.Enabled = False Then
                cmbUnidadMedida.Enabled = True
                cmbUnidadMedida.text = !UnidadMedida
             Else
                cmbUnidadMedida.text = !UnidadMedida
            End If
            
            If txtPrecioUnitarioFactura.Enabled = False Then
                txtPrecioUnitarioFactura.Enabled = True
                
             Else
                txtPrecioUnitarioFactura.text = !PrecioUnitarioFactura
            End If
            
            If txtPrecioUnitarioPresupuesto.Enabled = False Then
               txtPrecioUnitarioPresupuesto.Enabled = True
               txtPrecioUnitarioPresupuesto.text = !PrecioUnitarioPresupuesto
             Else
               txtPrecioUnitarioPresupuesto.text = !PrecioUnitarioPresupuesto
            End If
            
            If cmbRubro.Enabled = False Then
                cmbRubro.Enabled = True
                cmbRubro.text = !Rubro
             Else
                cmbRubro.text = !Rubro
            End If
                        
            
            If txtCodBarra.Enabled = False Then
                txtCodBarra.Enabled = True
                If IsNull(!codbarra) Then
                    txtCodBarra.text = "NO RELEVADO"
                 Else
                    txtCodBarra.text = !codbarra
                End If
             Else
                If IsNull(!codbarra) Then
                    txtCodBarra.text = "NO RELEVADO"
                 Else
                    txtCodBarra.text = !codbarra
                End If
            End If
            
            If chkPromo.Enabled = False Then
                chkPromo.Enabled = True
                If IsNull(!Promo) Then
                    chkPromo.Value = False
                 Else
                    chkPromo.Value = !Promo
                End If
             Else
                If IsNull(!Promo) Then
                    chkPromo.Value = False
                 Else
                    chkPromo.Value = !Promo
                End If
            End If
            
            If txtObservaciones.Enabled = False Then
                txtObservaciones.Enabled = True
                If Not IsNull(!observaciones) Then txtObservaciones.text = !observaciones
             Else
                If Not IsNull(!observaciones) Then txtObservaciones.text = !observaciones
            End If
            
            vFlagBuscar = 0
            
     End With
     
     Call DisabledTextBox(FormProductos)
     btnModificar.Enabled = True

End Sub
Private Sub btnLimpiar_Click()

    Call LimpiarPantalla

End Sub

Private Sub btnModificar_Click()

    If vFlagBuscar = 0 Then
        Call EnabledTextBox(FormProductos)
        btnModificar.Caption = "Guardar Cambios"
        btnGrabar.Enabled = False
        txtCodProd.SetFocus
        vFlagBuscar = 1
        txtCodProd.SetFocus
    Else
        A = MsgBox("¿Seguro desea Guardar las Modificaciones Realizadas?", vbQuestion + vbOKCancel, "INFO DEL SISTEMA")
        
        '1 es Ok 2 es Cancel
            If A = 1 Then
                With tProductos
                    .Edit
                        !CodProd = txtCodProd.text
                        !Descripcion = Format(txtDescripcion.text, ">")
                        !Stock = 0
                        '!PuntoPedido = txtPuntoPedido.Text
                        '!Tamaño = txtTamaño.Text
                        !UnidadMedida = cmbUnidadMedida.text
                        !PrecioUnitarioFactura = txtPrecioUnitarioFactura.text
                        !PrecioUnitarioPresupuesto = txtPrecioUnitarioPresupuesto.text
                        !Rubro = Format(cmbRubro.text, ">")
                        !codbarra = txtCodBarra.text
                        !Promo = chkPromo.Value
                        !observaciones = Format(txtObservaciones.text, ">")
                    .Update
                End With
                
                Call LimpiarPantalla
                
             Else
                If A = 2 Then
                    b = MsgBox("¿Limpia Pantalla?", vbQuestion + vbYesNo, "INFO DEL SISTEMA")
                    If b = 6 Then
                        Call LimpiarPantalla
                     Else
                        Call EnabledTextBox(FormProductos)
                        txtCodProd.SetFocus
                    End If
                End If
            
            End If
        End If

End Sub


Private Sub btnPrimero_Click()
    
    On Error GoTo CapturaErrores
    
    If Not tProductos.EOF Then
        tProductos.MoveFirst
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
     Else
        vSQL = "SELECT * FROM Productos ORDER BY CodProd"
        'MsgBox (vsql)
        Set tProductos = BaseSPC.OpenRecordset(vSQL)
        
        If Not tProductos.EOF Then
            tProductos.MoveFirst
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
    
    If Not tProductos.EOF Then
        tProductos.MoveLast
        Call Mostrar
        btnGrabar.Enabled = False
        btnModificar.Enabled = True
        btnEliminar.Enabled = True
     Else
        vSQL = "SELECT * FROM Productos ORDER BY IDCliente"
        'MsgBox (vsql)
        Set tProductos = BaseSPC.OpenRecordset(vSQL)
        
        If Not tProductos.EOF Then
            tProductos.Last
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

Private Sub chkPromo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub cmbRubro_GotFocus()
    
    cmbRubro.SelLength = Len(cmbRubro.text)
    
End Sub

Private Sub cmbRubro_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub cmbUnidadMedida_GotFocus()

    cmbUnidadMedida.SelLength = Len(cmbUnidadMedida.text)

End Sub

Private Sub cmbUnidadMedida_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub Form_Load()


   'Seteo tamaño y ubicacion del form
        FormProductos.Height = 6915
        FormProductos.Width = 14490
        FormProductos.Top = 1000
        FormProductos.Left = 1000
        
        btnModificar.Enabled = False
        btnEliminar.Enabled = False
        
    'Abro Base de Datos
        'Seteo la captura de errores de no hay registros en el archivo
         On Error GoTo CapturaErrores
        
        Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
        
        'Tabla Productos
            Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
        
        'Tabla Unidades de Medida
            Set tUnidadesMedida = BaseSPC.OpenRecordset("UnidadesMedida", dbOpenTable)
        
            tUnidadesMedida.Index = "PrimaryKey"
            tUnidadesMedida.MoveFirst
        
            'Lleno combo Unidades de Medida
                While Not tUnidadesMedida.EOF
                    cmbUnidadMedida.AddItem tUnidadesMedida!Descripcion
                    tUnidadesMedida.MoveNext
                Wend
        
        'Tabla Rubros
            Set tRubros = BaseSPC.OpenRecordset("Rubros", dbOpenTable)
        
            tRubros.Index = "PrimaryKey"
            tRubros.MoveFirst
        
            'Lleno combo Unidades de Medida
                While Not tRubros.EOF
                    cmbRubro.AddItem tRubros!Descripcion
                    tRubros.MoveNext
                Wend
        
        
            vFlagBuscar = 0
       
CapturaErrores:
    
  'Captura el error de no hay registros en la tabla
    Select Case Err
        Case 3021
            Resume Next
    End Select
        
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

Private Sub txtCodBarra_GotFocus()

    txtCodBarra.SelLength = Len(txtCodBarra.text)

End Sub

Private Sub txtCodBarra_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub txtCodProd_GotFocus()

    txtCodProd.SelLength = Len(txtCodProd.text)

End Sub

Private Sub txtCodProd_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If
    
   If KeyAscii = 27 Then
        Unload Me
   End If

End Sub


Private Sub txtCodProd_LostFocus()

   On Error GoTo CapErr
  
  If vFlagBuscar = 0 Then
    If txtCodProd.text <> "" Then
        
        tProductos.Index = "PrimaryKey"
        tProductos.Seek "=", txtCodProd.text
        
        If Not tProductos.NoMatch() Then
        
            Call Mostrar
            'vFlagBuscar = 1
            btnModificar.SetFocus
            
        End If
    End If
  End If
  
CapErr:
    Select Case Err
        Case 3251
            txtCodProd.SetFocus
            Exit Sub
    End Select
End Sub


Private Sub txtDescripcion_GotFocus()

    txtDescripcion.SelLength = Len(txtDescripcion.text)

End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


Private Sub txtDescripcion_LostFocus()

  If txtCodProd.text <> "" Then Exit Sub
  
  If vFlagBuscar = 0 Then
        Call Buscar
        If btnModificar.Enabled = True Then
            btnModificar.SetFocus
        End If
  End If

End Sub

Private Sub txtObservaciones_GotFocus()

    txtObservaciones.SelLength = Len(txtObservaciones.text)

End Sub


Private Sub txtPrecioUnitarioFactura_GotFocus()
    
    txtPrecioUnitarioFactura.SelLength = Len(txtPrecioUnitarioFactura.text)

End Sub


Private Sub txtPrecioUnitarioFactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)

End Sub


Private Sub txtPrecioUnitarioFactura_LostFocus()

    txtPrecioUnitarioFactura.text = Format(txtPrecioUnitarioFactura.text, "#,###,###,#0.00")

End Sub


Private Sub txtPrecioUnitarioPresupuesto_GotFocus()

    txtPrecioUnitarioPresupuesto.SelLength = Len(txtPrecioUnitarioPresupuesto.text)

End Sub

Private Sub txtPrecioUnitarioPresupuesto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)

End Sub

Private Sub txtPrecioUnitarioPresupuesto_LostFocus()

    txtPrecioUnitarioPresupuesto.text = Format(txtPrecioUnitarioPresupuesto.text, "#,###,###,#0.00")

End Sub


Private Sub txtPuntoPedido_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    KeyAscii = Verificar_Tecla(KeyAscii)

End Sub

Private Sub txtPuntoPedido_LostFocus()

    txtPuntoPedido.text = Format(txtPuntoPedido.text, "#,###,###,#0.00")

End Sub


Private Sub txtTamaño_KeyPress(KeyAscii As Integer)

    KeyAscii = Asc(StrConv(Chr$(KeyAscii), vbUpperCase))

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

End Sub


