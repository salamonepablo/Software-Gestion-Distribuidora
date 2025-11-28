VERSION 5.00
Begin VB.Form MenuPrincipal 
   AutoRedraw      =   -1  'True
   Caption         =   "MenuPrincipal"
   ClientHeight    =   8325
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   15540
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   15540
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdImprimirCheques 
      Caption         =   "Print C&heques"
      Height          =   735
      Left            =   6360
      TabIndex        =   12
      Top             =   4800
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000010&
      Caption         =   "ACCESOS DIRECTOS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   1800
      TabIndex        =   11
      Top             =   720
      Width           =   12975
      Begin VB.CommandButton cmdConsignaciones 
         Caption         =   "C&onsignaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10440
         TabIndex        =   20
         Top             =   480
         Width           =   1755
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "S&alir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         TabIndex        =   19
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdVerFormularios 
         Caption         =   "Buscar Formularios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5400
         TabIndex        =   18
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscarPresupuestos 
         Caption         =   "Buscar Presupuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2040
         TabIndex        =   17
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdBuscarFactura 
         Caption         =   "Buscar Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   16
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdVentas 
         Caption         =   "&Ventas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         TabIndex        =   15
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdLibroIva 
         Caption         =   "Libro &Iva"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5400
         TabIndex        =   14
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdComisiones 
         Caption         =   "C&omisiones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   8760
         TabIndex        =   13
         Top             =   1440
         Width           =   1455
      End
      Begin VB.CommandButton cmdCallProductos 
         Caption         =   "Lista de P&recios"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2040
         TabIndex        =   9
         Top             =   2400
         Width           =   1400
      End
      Begin VB.CommandButton cmdCallClientes 
         Caption         =   "&Clientes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   480
         TabIndex        =   8
         Top             =   2400
         Width           =   1400
      End
      Begin VB.CommandButton cmdCallMovStock 
         Caption         =   "&Movimientos Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   10
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdNotaCredito 
         Caption         =   "&Nota de Credito"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton cmdBajaPadron 
         Caption         =   "&Bajar Padron"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7080
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
      Begin VB.CommandButton cmdCallCtacte 
         Caption         =   "&Cuenta Corriente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   480
         TabIndex        =   6
         Top             =   480
         Width           =   1400
      End
      Begin VB.CommandButton cmdCallPresupuesto 
         Caption         =   "&Presupuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   480
         Width           =   1400
      End
      Begin VB.CommandButton cmdCallRemitos 
         Caption         =   "&Remitos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   5400
         TabIndex        =   2
         Top             =   480
         Width           =   1400
      End
      Begin VB.CommandButton cmdCallFacturas 
         Caption         =   "&Facturas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   3720
         TabIndex        =   1
         Top             =   480
         Width           =   1400
      End
      Begin VB.CommandButton cmdCallPagos 
         Caption         =   "&Pagos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8760
         TabIndex        =   4
         Top             =   480
         Width           =   1400
      End
      Begin VB.CommandButton cmdCallVentas 
         Caption         =   "&Saldos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   480
         TabIndex        =   5
         Top             =   1440
         Width           =   1400
      End
      Begin VB.Shape Shape3 
         Height          =   855
         Left            =   360
         Top             =   2280
         Width           =   12255
      End
      Begin VB.Shape Shape2 
         Height          =   855
         Left            =   360
         Top             =   1320
         Width           =   12255
      End
      Begin VB.Shape Shape1 
         Height          =   855
         Left            =   360
         Top             =   360
         Width           =   12255
      End
   End
   Begin VB.Image Image1 
      Height          =   8295
      Left            =   0
      Top             =   0
      Width           =   15495
   End
   Begin VB.Menu Archivo 
      Caption         =   "&Archivo"
      Begin VB.Menu Sale 
         Caption         =   "&Salir"
      End
   End
   Begin VB.Menu Clientes 
      Caption         =   "&Clientes"
      Begin VB.Menu A 
         Caption         =   "&ABM"
      End
   End
   Begin VB.Menu Produsctos 
      Caption         =   "&Productos"
      Begin VB.Menu Alta 
         Caption         =   "&ABM"
         Index           =   2
      End
      Begin VB.Menu VerProductos 
         Caption         =   "&Ver Productos"
      End
   End
   Begin VB.Menu Usuarios 
      Caption         =   "&Usuarios"
      Begin VB.Menu AltaU 
         Caption         =   "&ABM"
      End
   End
   Begin VB.Menu Facturas 
      Caption         =   "&Facturas"
      Begin VB.Menu AltaF 
         Caption         =   "&Nueva"
      End
      Begin VB.Menu bus 
         Caption         =   "&Buscar por Cliente"
      End
      Begin VB.Menu Pago 
         Caption         =   "&Pago"
         Begin VB.Menu NuevoPago 
            Caption         =   "N&uevo Pago"
         End
         Begin VB.Menu AnularPago 
            Caption         =   "Anu&lar Pago"
         End
      End
      Begin VB.Menu notacredito 
         Caption         =   "&Nota de Credito"
      End
   End
   Begin VB.Menu Presupuestos 
      Caption         =   "&Presupuestos"
      Begin VB.Menu NuevoP 
         Caption         =   "&Nuevo"
         Index           =   1
      End
   End
   Begin VB.Menu Remitos 
      Caption         =   "&Remitos"
      Begin VB.Menu Nuevo 
         Caption         =   "&Nuevo"
         Index           =   2
      End
   End
   Begin VB.Menu CtaCte 
      Caption         =   "&Cuenta Corriente"
      Begin VB.Menu Movimientos 
         Caption         =   "&Movimientos"
      End
   End
   Begin VB.Menu Empleados 
      Caption         =   "&Empleados"
      Begin VB.Menu AltaV 
         Caption         =   "&ABM"
      End
   End
   Begin VB.Menu Liquidacion 
      Caption         =   "&Liquidación"
      Begin VB.Menu LiqComisiones 
         Caption         =   "&Comisiones por Vendedor"
      End
      Begin VB.Menu ListadoVentas 
         Caption         =   "&Listado de Ventas"
      End
      Begin VB.Menu LibroIvaVentas 
         Caption         =   "&Libro Iva Ventas"
      End
   End
   Begin VB.Menu Inventario 
      Caption         =   "&Inventario"
      Begin VB.Menu Stock 
         Caption         =   "&Stock"
      End
   End
   Begin VB.Menu Herramientas 
      Caption         =   "&Herramientas"
      Begin VB.Menu Reportes 
         Caption         =   "&Reportes"
         Begin VB.Menu cli 
            Caption         =   "&Clientes"
         End
         Begin VB.Menu Prod 
            Caption         =   "&Productos"
         End
      End
   End
   Begin VB.Menu AboutS 
      Caption         =   "Acerca de..."
      Begin VB.Menu AboutSPC 
         Caption         =   "SPC Software"
      End
   End
End
Attribute VB_Name = "MenuPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub A_Click()
    FormClientes.Show
End Sub

Private Sub abmdepo_Click()
   FormDepositos.Show
End Sub

Private Sub AboutSPC_Click()

    
    vMensaje = "       SPC INTEGRATED SOFWTWARE ®" + Chr(10) + "                 Copyright © 2014 BY " + Chr(10) + Chr(10) + "                      VICTOR POMPA" + Chr(10) + "                    IGNACIO CHAFFIN" + Chr(10) + "                    PABLO SALAMONE"
    
    A = MsgBox(vMensaje, vbOKOnly, "©® SPC SOFTWARE")

End Sub

Private Sub Alta_Click(Index As Integer)
    FormProductos.Show
End Sub

Private Sub AltaF_Click()
    FormFactura.Show
End Sub

Private Sub AltaU_Click()
    FormEmpleados.Show
End Sub

Private Sub AltaV_Click()
    FormEmpleados.Show
End Sub

Private Sub AnularPago_Click()

    FormAnulacionPago.Show

End Sub

Private Sub bus_Click()
    
    FormBusquedaFacturaPorCliente.Show

End Sub

Private Sub cli_Click()
    FormPreviewClientes.lblReporte.Caption = "Clientes"
    FormPreviewClientes.Show 1
End Sub

Private Sub cmdBajaPadron_Click()
    
'    Dim vVal As Double
    
'    vVal = Shell(App.Path & "\PadronArba.exe", 1)

    FormImportTxt.Show
    
End Sub

Private Sub cmdBajaPadron_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdBuscarFactura_Click()

    FormBuscarFactura.Show

End Sub

Private Sub cmdBuscarFactura_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub

Private Sub cmdBuscarPresupuestos_Click()

    FormBuscarPresupuesto.Show

End Sub

Private Sub cmdBuscarPresupuestos_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallClientes_Click()

    FormClientes.Show
    
End Sub

Private Sub cmdCallClientes_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallCtacte_Click()

    FormMovimientosCuentaCorriente.Show

End Sub

Private Sub cmdCallCtacte_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallFacturas_Click()

    FormFactura.Show

End Sub

Private Sub cmdCallFacturas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallLiquidaciones_Click()

    FormLiqComisiones.Show

End Sub

Private Sub cmdCallLiquidaciones_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallMovStock_Click()
    
    FormMovimientosStock.Show

End Sub

Private Sub cmdCallMovStock_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallPagos_Click()

    FormPagoFacturas.Show

End Sub

Private Sub cmdCallPagos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallPresupuesto_Click()

    FormPresupuesto.Show

End Sub

Private Sub cmdCallPresupuesto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallProductos_Click()
    
    FormVerProductos.Show

End Sub

Private Sub cmdCallProductos_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdCallRemitos_Click()

    FormRemito.Show

End Sub

Private Sub cmdCallRemitos_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub

Private Sub cmdCallVentas_Click()

    Call FormListadoClientesPorVendedor.Show

End Sub

Private Sub cmdCallVentas_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdComisiones_Click()
    
    FormLiqComisiones.Show
    
End Sub

Private Sub cmdComisiones_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub

Private Sub cmdConsignaciones_Click()

    FormConsignaciones.Show

End Sub

Private Sub cmdImprimirCheques_Click()

    FormChequesBapro.Show

End Sub

Private Sub cmdImprimirCheques_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
    End If

End Sub

Private Sub cmdLibroIva_Click()
    
    FormLibroIvaVentas.Show
    
End Sub

Private Sub cmdLibroIva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub cmdNotaCredito_Click()

    FormNotaCredito.Show

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdNotaCredito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub

Private Sub cmdSalir_Click()

    Unload Me

End Sub

Private Sub cmdVentas_Click()
    
    FormListadoVentas.Show
    
End Sub

Private Sub Command3_Click()

End Sub

Private Sub cmdVentas_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub

Private Sub Command1_Click()



End Sub

Private Sub cmdVerFormularios_Click()

    FormBuscarFormularios.Show

End Sub

Private Sub cmdVerFormularios_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload MenuPrincipal
        End
    End If

End Sub


Private Sub Form_Load()

    Image1.Top = 0
    Image1.Left = 0
    
    LlamaPagoPresup = False

End Sub


Private Sub LibroIvaVentas_Click()

    FormLibroIvaVentas.Show

End Sub

Private Sub LiqComisiones_Click()

    FormLiqComisiones.Show

End Sub

Private Sub ListadoVentas_Click()

    FormListadoVentas.Show

End Sub

Private Sub modifactura_Click()
    FormBusquedaFactura.Show
End Sub

Private Sub Movimientos_Click()
    FormMovimientosCuentaCorriente.Show
End Sub

Private Sub notacredito_Click()
    FormNotaCredito.Show
End Sub

Private Sub Nuevo_Click(Index As Integer)
    FormRemito.Show
End Sub

Private Sub NuevoP_Click(Index As Integer)
    FormPresupuesto.Show
End Sub

Private Sub NuevoPago_Click()

    FormPagoFacturas.Show

End Sub

Private Sub Prod_Click()
      FormPreviewProductos.lblReporte.Caption = "ListadoProductos"
      FormPreviewProductos.Show 1
End Sub

Private Sub Sale_Click()
    
    Unload MenuPrincipal
    End

End Sub

Private Sub Stock_Click()

    
    FormMovimientosStock.Show

End Sub

Private Sub VerProductos_Click()

    FormVerProductos.Show

End Sub
