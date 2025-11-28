VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormPresupuesto 
   BackColor       =   &H00808000&
   Caption         =   "Presupuestos"
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12045
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   12045
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   4200
      TabIndex        =   46
      Top             =   2280
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808000&
      Height          =   975
      Left            =   120
      TabIndex        =   60
      Top             =   2160
      Width           =   11655
      Begin VB.TextBox TextNumeroPresupuesto 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextFechaPresupuesto 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2880
         TabIndex        =   11
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox ComboVendedor 
         Height          =   315
         Left            =   4920
         TabIndex        =   12
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox TextSaldoCliente 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8040
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox CheckModificaStock 
         BackColor       =   &H00808000&
         Caption         =   "Modifica Stock"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   9840
         TabIndex        =   15
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.TextBox TextDescuentoCliente 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         TabIndex        =   13
         Top             =   480
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Nº Presupuesto"
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
         TabIndex        =   65
         Top             =   240
         Width           =   1320
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Fecha Presupuesto"
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
         TabIndex        =   64
         Top             =   240
         Width           =   1650
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
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
         Left            =   4920
         TabIndex        =   63
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Saldo"
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
         Left            =   8400
         TabIndex        =   62
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Descuento"
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
         Left            =   6840
         TabIndex        =   61
         Top             =   240
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808000&
      Caption         =   "Productos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   51
      Top             =   3120
      Width           =   11655
      Begin VB.ComboBox ComboArticulo 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TextCantidad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7440
         TabIndex        =   6
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox TextPrecioUnitario 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6120
         TabIndex        =   5
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TextPorDexcuento 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8400
         TabIndex        =   7
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TextUnidadMedida 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5520
         TabIndex        =   4
         Top             =   600
         Width           =   495
      End
      Begin VB.TextBox TextTotalLineaProd 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9000
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextDescripcion 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   600
         Width           =   3975
      End
      Begin VB.CommandButton BotonAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   10560
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.TextBox TextPorDesc 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9960
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Total Línea"
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
         Left            =   9120
         TabIndex        =   59
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Cantidad"
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
         TabIndex        =   58
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "% Desc."
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
         Left            =   8280
         TabIndex        =   57
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Precio Unit."
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
         Left            =   6240
         TabIndex        =   56
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "UM"
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
         Left            =   5640
         TabIndex        =   55
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Descripcion"
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
         Left            =   2880
         TabIndex        =   54
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Articulo"
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
         TabIndex        =   53
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00808000&
      Height          =   1935
      Left            =   120
      TabIndex        =   25
      Top             =   120
      Width           =   11655
      Begin VB.TextBox TextItemDomicilio 
         Height          =   285
         Left            =   4320
         TabIndex        =   48
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton BotonDomicilio 
         BackColor       =   &H00808000&
         Caption         =   "&Domicilio Entrega"
         Enabled         =   0   'False
         Height          =   510
         Left            =   9960
         MaskColor       =   &H00808000&
         TabIndex        =   47
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TextDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   37
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TextApellidoNombre 
         Height          =   285
         Left            =   7080
         TabIndex        =   17
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   29
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   28
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   27
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TextProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   26
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Código Cliente:"
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
         TabIndex        =   36
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Apellido Nombre:"
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
         Left            =   5160
         TabIndex        =   35
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
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
         Left            =   120
         TabIndex        =   34
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Dirección:"
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
         Left            =   5160
         TabIndex        =   33
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Código Postal:"
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
         TabIndex        =   32
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Localidad:"
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
         TabIndex        =   31
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Provincia:"
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
         Left            =   7080
         TabIndex        =   30
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808000&
      Height          =   1095
      Left            =   120
      TabIndex        =   39
      Top             =   8400
      Width           =   11655
      Begin VB.TextBox Textpre 
         Height          =   375
         Left            =   6960
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton BotonGrabar 
         BackColor       =   &H00808000&
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   750
         Left            =   1560
         MaskColor       =   &H00808000&
         TabIndex        =   16
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonPago 
         BackColor       =   &H00808000&
         Caption         =   "&Pago"
         Enabled         =   0   'False
         Height          =   750
         Left            =   9600
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton BotonSalir 
         BackColor       =   &H00808000&
         Caption         =   "&Salir"
         Height          =   750
         Left            =   4920
         TabIndex        =   43
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonCancelar 
         BackColor       =   &H00808000&
         Caption         =   "&Cancelar"
         Height          =   750
         Left            =   4080
         TabIndex        =   42
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonImprimir 
         BackColor       =   &H00808000&
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   750
         Left            =   2400
         TabIndex        =   41
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonNueva 
         BackColor       =   &H00808000&
         Caption         =   "&Nuevo"
         Enabled         =   0   'False
         Height          =   750
         Left            =   720
         MaskColor       =   &H00808000&
         TabIndex        =   40
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00808000&
      Height          =   3015
      Left            =   120
      TabIndex        =   38
      Top             =   4320
      Width           =   11655
      Begin VB.CommandButton BotonEliminarfila 
         Caption         =   "&Eliminar Fila"
         Height          =   495
         Left            =   10800
         TabIndex        =   49
         Top             =   1440
         Width           =   735
      End
      Begin VB.CommandButton BotonBuscarProducto 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   10800
         MaskColor       =   &H00808000&
         TabIndex        =   45
         Top             =   2160
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2655
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4683
         _Version        =   393216
         Rows            =   16
         Cols            =   9
         FixedCols       =   0
         Enabled         =   0   'False
         GridLines       =   2
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00808000&
      Height          =   975
      Left            =   120
      TabIndex        =   18
      Top             =   7320
      Width           =   11655
      Begin VB.TextBox TextSubtotalPresupuesto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   480
         TabIndex        =   21
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextDescuentos 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         TabIndex        =   20
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextTotalPresupuesto 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9840
         TabIndex        =   19
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Subtotal Presupuesto:"
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
         Top             =   240
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Descuentos:"
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
         Left            =   2520
         TabIndex        =   23
         Top             =   240
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00808000&
         Caption         =   "Total Presupuesto:"
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
         Left            =   9720
         TabIndex        =   22
         Top             =   240
         Width           =   1620
      End
   End
End
Attribute VB_Name = "FormPresupuesto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstPresupuestoC As DAO.Recordset
 Dim rstPresupuestoD As DAO.Recordset
 Dim rstPadron As DAO.Recordset
 Dim rstUltimosNumeros As DAO.Recordset
 Dim rstIva As DAO.Recordset
 Dim rstCtaCte As DAO.Recordset
 Dim rstMovimientosCtaCte As DAO.Recordset
 Dim cantidadProducto As Integer
 Dim descuentos As Double
 Dim vendedorCliente As String
 Dim nombreVendedor As Integer
 Dim LegajoEmpleado As String
 Dim Alicuota As Double
 Dim condicionIva As String
 Dim modificaStock As Integer
 Dim saldo1 As Double
 Dim saldo2 As Double
 Dim saldoLi1 As Double
 Dim Fila As Integer
 Dim fila2 As Integer
 Dim renglon As Integer
 Dim PresuC
 Dim PresuD
 Dim vSQLPC
 Dim vSQLPD

Private Sub BotonAgregar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If


End Sub

Private Sub BotonBuscarProducto_Click()

    FormBusquedaProductosPresupuesto.Show

End Sub

Private Sub BotonCancelar_Click()

    Call blanqueototal
    
End Sub
Private Sub BotonAgregar_Click()

    If fila2 < renglon Then
        If Fila > 1 Then
         '   FG1.AddItem (ComboArticulo.Text)
            FG1.Row = Fila
        Else
            FG1.Row = 1
            'FG1.Col = 0
            'FG1.Text = ComboArticulo.Text
        End If
        
        FG1.Col = 0
        FG1.Text = ComboArticulo.Text
        FG1.Col = 1
        FG1.Text = TextDescripcion.Text
        FG1.Col = 2
        FG1.Text = TextUnidadMedida.Text
        FG1.Col = 3
        FG1.Text = Format(TextPrecioUnitario.Text, "#0.00")
        FG1.Col = 4
        FG1.Text = Format(TextPorDexcuento.Text, "#0.00")
        FG1.Col = 5
        FG1.Text = TextCantidad.Text
        FG1.Col = 6
        FG1.Text = Format(TextTotalLineaProd.Text, "#0.00")
        FG1.Col = 7
        FG1.Text = Format(TextPorDesc.Text, "#0.00")
                
        Fila = Fila + 1
        fila2 = fila2 + 1
            
    
        Call calculototalpresupuesto
        
'       ComboArticulo.Text = ""
        TextDescripcion.Text = ""
        TextUnidadMedida.Text = ""
        TextPrecioUnitario.Text = ""
        TextPorDexcuento.Text = ""
        TextCantidad.Text = ""
        TextTotalLineaProd.Text = ""
        TextPorDesc.Text = ""
        
        ComboArticulo.SetFocus
        BotonGrabar.Enabled = True
        BotonImprimir.Enabled = False
    
  End If

End Sub

Private Sub calculototalpresupuesto()

    
    Dim total
    Dim subtotalFacturaForm
    Dim porcentajePrecioUnitario As Double
    Dim descuentoFactura As Double
    Dim totalFacturaForm As Double
    Dim iva As Double
    Dim impuesto As Double
    Dim percepcion As Double
    Dim preuni As Double
    Dim modifico As Integer
    Dim nnmodifico As Integer

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstIva = db.OpenRecordset("Iva", dbOpenDynaset)

    iva = rstIva.Fields!iva
    
   

    '**** suma total linea
             
    'total = Val(TextTotalLineaProd.Text) + Val(TextSubtotalFactura.Text)
    If (TextTotalPresupuesto.Text = "") Then TextTotalPresupuesto.Text = 0
    total = CDbl(TextTotalLineaProd.Text) + CDbl(TextTotalPresupuesto.Text)
    subtotalFacturaForm = total
                            
    TextTotalPresupuesto.Text = Format(total, "#0.00")
    
'    '**** suma descuentos
'
'    'porcentajePrecioUnitario = Val(porcentajePrecioUnitario) + Val(TextPorDesc.Text)
'    If (TextDescuentos.Text = "") Then TextDescuentos.Text = 0
'    porcentajePrecioUnitario = CDbl(TextDescuentos.Text) + CDbl(TextPorDesc.Text)
'    descuentoFactura = porcentajePrecioUnitario
'
'    TextDescuentos.Text = Format(descuentoFactura, "#0.00")
'
'    '**** calculo alicuota
'
'    TextAlicuota.Text = Format(Alicuota, "#0.00")
'
'    If TextTipoFactura.Text = "A" Then
'        percepcion = subtotalFacturaForm * Alicuota / 100
'        TextPercepcionIIBB.Text = Format(percepcion, "#0.00")
'
'    End If
'
'    '**** calculo impuesto
'
'    If TextTipoFactura.Text = "A" Then
'       impuesto = subtotalFacturaForm * iva / 100
'       TextImpuesto.Text = Format(impuesto, "#0.00")
'    End If
    
    '**** calculo total factura
    
    'totalFacturaForm = (subtotalFacturaForm - descuentoFactura + percepcion + impuesto)
    
'    totalFacturaForm = (subtotalFacturaForm + percepcion + impuesto)
'
'    TextTotalFactura.Text = Format(totalFacturaForm, "#0.00")
'
    If CDec(totalFacturaForm) <> 0 Then
         BotonGrabar.Enabled = True
         'BotonImprimir.Enabled = True
         'BotonPago.Enabled = True
    End If

End Sub

Public Sub blanqueototal()

    TextCodigoCliente.Text = ""
    TextApellidoNombre.Text = ""
    TextCuit.Text = ""
    TextDireccion.Text = ""
    TextLocalidad.Text = ""
    TextCodigoPostal.Text = ""
    TextProvincia.Text = ""
    TextNumeroPresupuesto.Text = ""
    ComboVendedor.Text = ""
    TextDescuentoCliente.Text = ""
    TextSubtotalPresupuesto.Text = ""
    TextDescuentos.Text = ""
    TextTotalPresupuesto.Text = ""
    TextSaldoCliente.Text = ""
    ComboVendedor.Text = ""
    TextDescuentoCliente.Text = ""
    'ComboArticulo.Text = ""
    TextDescripcion.Text = ""
    TextUnidadMedida.Text = ""
    TextPrecioUnitario.Text = ""
    'CheckModificaStock.Value = Unchecked
    FG1.Clear
    FG1.Enabled = False
    'fila2 = 0
    'Fila = 0
    'renglon = 0
    
 
    Call SeteoGrilla

End Sub

Private Sub BotonCancelar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonDomicilio_Click()

    FormDomiciliosPresupuesto.Show
    
End Sub

Private Sub BotonEliminarfila_Click()

    If FG1.Row <= 0 Then
        MsgBox "Debe Seleccionar una fila"
    'ElseIf MSFlexGrid1.Row = 1 Then
    ' MSFlexGrid1.Clear
    Else
        FG1.RemoveItem (FG1.Row)
        Call calculos
    End If
  
End Sub

Private Sub BotonGrabar_Click()

        Dim descuentoCantidad As Long
        Dim ultimo As Long
        Dim existeNumeroBD As Integer
        Dim existeTipoBD As String
        Dim existeNumero As Long
        Dim existeTipo As String
        
        Textpre.Text = 1
        
        ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstPresupuestoC = db.OpenRecordset("PresupuestoC", dbOpenDynaset)
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstPresupuestoD = db.OpenRecordset("PresupuestoD", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstMovimientosCtaCte = db.OpenRecordset("MovimientosCtaCte", dbOpenDynaset)
        
        existeNumero = TextNumeroPresupuesto.Text
       
        'buscamos el deposito para descontar el stock
        
        Set tDepositos = db.OpenRecordset("Depositos", dbOpenTable)
      '     On Error GoTo CapturaErrores
   
           tDepositos.Index = "IndXVendedor"
           
           tDepositos.MoveFirst
           tDepositos.Seek "=", LegajoEmpleado
           
           If Not tDepositos.NoMatch Then
            DepoOrigen = tDepositos!IdDeposito
            'MsgBox (DepoOrigen)
           Else
            A = MsgBox("ERROR !!", vbCritical, "Vendedor sin Depósito Asociado")
           End If
              
           tDepositos.Close
    
        
        '*** Busco Numero de Presupuesto Existentes
        
        
            
               
        'existeNumero = Val(TextNumeroPresupuesto.Text)
      
        'rstPresupuestoC.FindFirst "NroFactura= " + Str(existeNumero) And "TipoFactura >= '" & existeTipo & "'"
        'rstPresupuestoC.FindFirst "NroFactura= " + Str(existeNumero) And "TipoFactura = '" & existeTipo & "'"
        'rstPresupuestoC.FindFirst "NroFactura= " + Str(existeNumero)
        'rstPresupuestoC.FindFirst "TipoFactura >= '" & busca5 & "' and TipoFactura <= '" & busca6 & "'"
       
        'existeNumeroBD = rstPresupuestoC.Fields!NroFactura
        'existeTipoBD = rstPresupuestoC.Fields!TipoFactura
     
        'If existeNumero = existeNumeroBD And existeTipo = existeTipoBD Then
        '    mensaje = MsgBox("Numero y Tipo de Factura Existentes", vbCritical, "Final de la busqueda")
        '    TextNumeroPresupuesto.Text = ""
        '    TextNumeroPresupuesto.SetFocus
        'else
            rstPresupuestoC.AddNew
            rstPresupuestoC.Fields!NroPresu = TextNumeroPresupuesto.Text
            rstPresupuestoC.Fields!FechaPresu = TextFechaPresupuesto.Text
            rstPresupuestoC.Fields!TotalPresu = TextTotalPresupuesto.Text
            'rstPresupuestoC.Fields!SubTotalPresu = TextSubtotalPresupuesto.Text
            rstPresupuestoC.Fields!SubTotalPresu = TextTotalPresupuesto.Text
            rstPresupuestoC.Fields!CodCliente = TextCodigoCliente.Text
            rstPresupuestoC.Fields!PorcentajeDesc = TextDescuentoCliente.Text
'            rstPresupuestoC.Fields!ImporteDesc = TextDescuentos.Text
            
            rstPresupuestoC.Fields!codVendedor = LegajoEmpleado
            rstPresupuestoC.Update
            
            FG1.Col = 0
            FG1.Row = 1
            Filas = FG1.Rows
            linea = 1
            Do While linea < Filas
                  
                  FG1.Row = linea
                  FG1.Col = 0
                  If FG1.Text <> "" Then
                        rstPresupuestoD.AddNew
                    
                        rstPresupuestoD.Fields!NroPresu = TextNumeroPresupuesto.Text
                        'rstPresupuestoD.Fields!TipoFactura = TextTipoFactura.Text
                    
                        FG1.Col = 0
                        rstPresupuestoD.Fields!CodProd = FG1.Text
                    
                        FG1.Col = 2
                        rstPresupuestoD.Fields!UnidadMedida = FG1.Text
                        
                        FG1.Col = 3
                        rstPresupuestoD.Fields!precioUnitario = Format(FG1.Text, "#00.00")
                        
                        FG1.Col = 4
                        des = FG1.Text
                        If des <> "" Then
                           rstPresupuestoD.Fields!PorcentajeDescuento = Val(des)
                        Else
                           rstPresupuestoD.Fields!PorcentajeDescuento = Val(TextDescuentoCliente.Text)
                        End If
                        FG1.Col = 5
                        rstPresupuestoD.Fields!cantidad = Val(FG1.Text)
                        descuentoCantidad = Val(FG1.Text)
                        
                        '*** Modifico Stock Producto
                       

                       'Call DesHagoStock(CodProd, descuentoCantidad)
                        
                        If modificaStock = 1 Then
                            FG1.Col = 0
                            codigoprod = FG1.Text
                            
                            'Dim busca1 As String, busca2 As String
                            'busca1 = RTrim(LTrim(codigoprod))
                            'busca2 = busca1 + "z"
                       
                            'rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
                            
                            'rstProductos.Edit
                            'rstProductos.Fields!Stock = cantidadProducto - descuentoCantidad
                            'rstProductos.Update
                            
                            Call ActualizarStock(codigoprod, DepoOrigen, descuentoCantidad)
                        End If
                        
                        FG1.Col = 6
                        rstPresupuestoD.Fields!totalLinea = Format(FG1.Text, "#00.00")
                        
                        FG1.Col = 7
                        rstPresupuestoD.Fields!ImporteDescuento = Format(FG1.Text, "#00.00")
                        
                        FG1.Col = 8
                        rstPresupuestoD.Fields!ItemPresu = Val(FG1.Text)
                         
                        rstPresupuestoD.Update
                  End If
                  linea = linea + 1
            Loop
            
            '*** Grabo Linea 1 en Cuenta Corriente
            
            CodigoClie = Val(TextCodigoCliente.Text)
      
            rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCliente.Fields!IdCliente <> Val(TextCodigoCliente.Text) Then
                mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                'TextCodigoCliente.Text = ""
                'Call blanqueototal
                'TextCodigoCliente.SetFocus
            Else
                rstCtaCte.Edit
                saldo1 = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
                saldo2 = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
                saldoLi2 = Format(TextTotalPresupuesto.Text, "#0.00")
                rstCtaCte.Fields!SaldoL2 = saldoLi2 + saldo2
                saldo2 = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
                rstCtaCte.Fields!SaldoTotal = saldo1 + saldo2
                rstCtaCte.Fields!FechaActSaldo = Format(Date, "DD/MM/YYYY")
                rstCtaCte.Update
            End If
        
            
            '*** Grabo Movimientos Cuente corriente
        
            rstMovimientosCtaCte.AddNew
            'rstMovimientosCtaCte.Fields!Fecha = Format(Date, "dd/mm/yyyy")
            rstMovimientosCtaCte.Fields!Fecha = Format(TextFechaPresupuesto.Text, "DD/MM/YYYY")
            rstMovimientosCtaCte.Fields!IdCliente = TextCodigoCliente.Text
            rstMovimientosCtaCte.Fields!tipoDoc = "Presupuesto"
           
            rstMovimientosCtaCte.Fields!NroDoc = TextNumeroPresupuesto.Text
            rstMovimientosCtaCte.Fields!ImporteLinea1 = 0
            rstMovimientosCtaCte.Fields!ImporteLinea2 = TextTotalPresupuesto.Text
            rstMovimientosCtaCte.Update
            
            '*** Actualizo Ultimo Numero Presupuesto
            
            Set db = DBEngine.OpenDatabase(ruta)
            Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
        
            Dim Busco As String
       
             Busco = "tPresupuestoC"
                    
            rstUltimosNumeros.FindFirst "IDTabla >= '" & Busco & "' "
            ultimo = rstUltimosNumeros.Fields!UltimoNumero
            
            If ultimo < Val(TextNumeroPresupuesto.Text) Then
                rstUltimosNumeros.Edit
                rstUltimosNumeros.Fields!UltimoNumero = TextNumeroPresupuesto.Text
                rstUltimosNumeros.Update
            End If
        ' End If
        
        
        
        BotonGrabar.Enabled = False
        BotonNueva.Enabled = False
        BotonImprimir.Enabled = True
        BotonPago.Enabled = False
                 
        modificaStock = 0
        
         respuesta = MsgBox("Desea Realizar un Pago", vbYesNo, "Pago")
         
         If respuesta = vbYes Then
            FormPagoFacturasDesdeFactura.Show
         Else
            respuesta2 = MsgBox("¿Desea Imprimir el Presupuesto?", vbYesNo, "Pago")
         
            If respuesta2 = vbYes Then
               Call Imprimir
            Else
               Call blanqueototal
               BotonNueva.Enabled = True
               MSFlexGrid1.Visible = False
               TextCodigoCliente.SetFocus
            End If
         End If
         
         
        
    fila2 = 0
    Fila = 1
        
CapturaErrores:
        Select Case Err
            Case 3021
                Resume Next
        End Select
    
'    fila2 = 0
'    Fila = 0
    
End Sub

Public Sub Imprimir()
    Dim PU, TL, Cant, TotalPres As Variant
    'PU = 0
    'TL = 0
    X = -4
    Y = -4
    renglon = 0
     
    With Printer
        'On Error GoTo CapturaErrores
        
        'Seteo escala a mm
            .ScaleMode = 6
        
        'Cantidad de Impresiones
            .Copies = 3
            
        'Imprimir Codigo de Cliente
            .CurrentX = X + 25
            .CurrentY = Y + 25
            .Font = "Courier New"
            .FontSize = 16
            .FontBold = True
            .ForeColor = vbRed
            Printer.Print TextCodigoCliente.Text
        
        'Imprimir Fecha
            .ForeColor = vbBlack
            .CurrentX = X + 120
            .CurrentY = Y + 27
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print Format(TextFechaPresupuesto.Text, "DD/MM/YYYY")
        
        'Imprimir Nombres
            .CurrentX = X + 37
            .CurrentY = Y + 52
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = True
            Printer.Print TextApellidoNombre.Text
            
        'Imprimir Direccion
            .CurrentX = X + 37
            .CurrentY = Y + 59
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            'Printer.Print TextDireccion.Text & Chr(9) & Chr(9) & Chr(9) & Chr(9) & TextLocalidad.Text
            Printer.Print TextDireccion.Text
        
        'Imprimir Localidad
            .CurrentX = X + 120
            .CurrentY = Y + 59
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextLocalidad.Text
            
            
        'Imprimir Detalle
            Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
            vSQLPC = "SELECT * FROM PresupuestoC WHERE NroPresu=" & TextNumeroPresupuesto.Text & " ORDER By NroPresu"
            vSQLPD = "SELECT * FROM PresupuestoD WHERE NroPresu=" & TextNumeroPresupuesto.Text & " ORDER By NroPresu"
            
            Set PresuC = BaseSPC.OpenRecordset(vSQLPC, dbOpenDynaset)
            Set PresuD = BaseSPC.OpenRecordset(vSQLPD, dbOpenDynaset)
            
           
            PresuC.MoveFirst
            PresuD.MoveFirst
                
                    While Not PresuD.EOF
                        'Imprimo el detalle
                            .CurrentX = X + 13
                            .CurrentY = Y + 90 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            
                            Cant = CDbl(PresuD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print PresuD!cantidad
                            
                        'Detalle
                            .CurrentX = X + 30
                            .CurrentY = Y + 90 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            'Printer.Print PresuD!IdCodProd & Chr(9) & Descripcion(PresuD!IdCodProd)
                            Printer.Print BuscarDescProd(PresuD!CodProd)
                        
                        'Precio
                            .CurrentX = X + 115
                            .CurrentY = Y + 90 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            PU = CDbl(PresuD!precioUnitario) - (CDbl(PresuD!precioUnitario) * CDbl(PresuD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU
                        
                        'Importe
                            .CurrentX = X + 132
                            .CurrentY = Y + 90 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            TL = Format(PresuD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                        
                         renglon = renglon + 5
                            
                        PresuD.MoveNext
                    Wend
            
            'Importe Total
                .CurrentX = X + 132
                .CurrentY = Y + 197
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                
                TotalPres = Format(CDbl(PresuC!TotalPresu), "Standard")
                Hasta = CInt(14 - Len(TotalPres))
                For I = 0 To Hasta
                    TotalPres = " " & TotalPres
                Next I
                
                Printer.Print TotalPres
            
        .EndDoc
        
    End With
    
    PresuC.Close
    PresuD.Close
    BaseSPC.Close
    
    Call blanqueototal
    MSFlexGrid1.Visible = False
    TextCodigoCliente.SetFocus
    Unload FormPagoFacturasDesdeFactura
        
CapturaErrores:
    'If Err = 321 Then
    'End If


End Sub


Private Sub ActualizarStock(CodProd, IdDepoOrigen, Cant)

    'Sumo el Stock en Depósito Destino
        Set tS = db.OpenRecordset("Stock", dbOpenTable)
        
        tS.Index = "PrimaryKey"
        tS.MoveFirst
        
        'Resto el Stock en Depósito Origen
          tS.Seek "=", CodProd, IdDepoOrigen
            
        If Not tS.NoMatch Then
            tS.Edit
                tS!CodProd = CodProd
                tS!IdDeposito = IdDepoOrigen
                tS!cantidad = tS.cantidad - FormatNumber(Cant, 2)
                tS!FechaUM = Format(Date, "DD/MM/YYYY")
            tS.Update
        End If
    
    'Sumo el Stock en Depósito Destino
    '    tS.Seek "=", CodProd, IdDepoDestino
              
        'Si tiene stock de este producto
     '       If Not tS.NoMatch Then
                'CantIni = tSotck!Stock
     '           tS.Edit
     '               tS.CodProd = CodProd
     '               tS.IdDeposito = IdDepoDestino
     '              tS.cantidad = tS!cantidad + FormatNumber(Cant, 2)
     '               tS.FechaUM = Format(Date, "DD/MM/YYYY")
     '           tS.Update
        'Si no tiene stock de este producto
     '        Else
     '           tS.AddNew
     '               tS!CodProd = CodProd
     '               tS!IdDeposito = IdDepoDestino
     '               tS!cantidad = FormatNumber(Cant, 2)
     '               tS!FechaUM = Format(Date, "DD/MM/YYYY")
     '           tS.Update
     '       End If
    
End Sub


Private Sub BotonGrabar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonImprimir_Click()

    Call Imprimir

End Sub

Private Sub BotonImprimir_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub BotonNueva_Click()

    Dim NumeroPresupuesto As Long
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
 
    Dim Busco As String
       
    Busco = "tPresupuestoC"
        
    rstUltimosNumeros.FindFirst "IDTabla >= '" & Busco & "' "
    NumeroPresupuesto = rstUltimosNumeros.Fields!UltimoNumero
    
   
    TextNumeroPresupuesto.Text = NumeroPresupuesto + 1

    If TextCuit.Text <> "" Then
       FG1.Enabled = True
    End If
    
    BotonNueva.Enabled = False
    
    FG1.Enabled = True
    FG1.Row = 1
    FG1.Col = 0
    'FG1.SetFocus
    TextNumeroPresupuesto.SetFocus
    'TextCodigoCliente.SetFocus
    
End Sub

Private Sub BotonNueva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonPago_Click()

    'FormPagoFacturas.Show
    
End Sub

Private Sub BotonPago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub BotonSalir_Click()

    Unload FormPresupuesto

End Sub

Private Sub BotonSalir_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub CheckModificaStock_Click()

    If CheckModificaStock.Value = Unchecked Then
        modificaStock = 0
    End If
    
End Sub

Private Sub CheckModificaStock_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub ComboArticulo_Click()

    TextPrecioUnitario.Text = ""
    TextPorDexcuento.Text = ""
    TextCantidad.Text = ""
    TextTotalLineaProd.Text = ""
    TextPorDesc.Text = ""
    
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
   
      
    Dim busca1 As String
    busca1 = RTrim(LTrim(ComboArticulo.Text))
   
    
    rstProductos.FindFirst "CodProd >= '" & busca1 & "' "
    
   
    TextDescripcion.Text = rstProductos.Fields!Descripcion
    TextUnidadMedida.Text = rstProductos.Fields!UnidadMedida
    TextPrecioUnitario.Text = Format(rstProductos.Fields!PrecioUnitarioPresupuesto, "#0.00")
      
   ' TextCantidad.SetFocus
    
    
    
End Sub
Private Sub ComboArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'Call calculoprecios
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

    
End Sub
Private Sub ComboArticulo_LostFocus()
    
    TextPrecioUnitario.Text = ""
    TextPorDexcuento.Text = ""
    TextCantidad.Text = ""
    TextTotalLineaProd.Text = ""
    TextPorDesc.Text = ""
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    'If KeyAscii = 13 Then
      
           Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
        
   
           tProductos.Index = "PrimaryKey"
           
           tProductos.MoveFirst
           tProductos.Seek "=", ComboArticulo.Text
           
           If Not tProductos.NoMatch Then
                produ = tProductos!CodProd
                'MsgBox (DepoOrigen)
                 Dim busca1 As String, busca2 As String
                 busca1 = RTrim(LTrim(ComboArticulo.Text))
                 busca2 = busca1 + "z"
                                     
                 rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
                             
                 TextDescripcion.Text = rstProductos.Fields!Descripcion
                 TextUnidadMedida.Text = rstProductos.Fields!UnidadMedida
                 TextPrecioUnitario.Text = Format(rstProductos.Fields!PrecioUnitarioPresupuesto, "#0.00")
                
           Else
                mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                'ComboArticulo.Text = ""
                TextDescripcion.Text = ""
                TextUnidadMedida.Text = ""
                TextPrecioUnitario.Text = ""
                ComboArticulo.SetFocus
           End If
              
           tProductos.Close
          'TextCantidad.SetFocus
    'End If

End Sub



Private Sub ComboVendedor_GotFocus()
    ComboVendedor.SelLength = Len(ComboVendedor.Text)
'    ComboVendedor.SelLength = Len(ComboVendedor.Text)

End Sub


Private Sub ComboVendedor_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextApellidoNombre_Change()
    Columna = 1
    Call FiltrarGrilla(MSFlexGrid1, TextApellidoNombre, CLng(Columna))
End Sub

Private Sub TextApellidoNombre_GotFocus()
'    TextApellidoNombre.SelLength = Len(TextApellidoNombre.Text)
End Sub

Private Sub TextCantidad_GotFocus()

    TextCantidad.SelLength = Len(TextCantidad.Text)

End Sub
Private Sub FiltrarGrilla(MSFlexGrid1 As Object, TBox As TextBox, Columna As Long)
    
    Dim A As Integer
    
    
    If (KeyRetroceso Or Len(TBox.Text) = 0) Then
        'KeyRetroceso = False
        'Exit Sub
        TBox.Text = ""
    End If
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")

    
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    
    Call titulos
   
    A = Len(TBox.Text)

    If A >= 4 Then
    
        vSQL = "SELECT * FROM Clientes WHERE RazonSocial Like '*" & TBox.Text & "*' ORDER BY RazonSocial"
        
        Set tClientes = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
        
        
        linea2 = 1
        
        Do While Not tClientes.EOF()
                MSFlexGrid1.AddItem " "
                MSFlexGrid1.Row = linea2
            
            
                MSFlexGrid1.Col = 0
                MSFlexGrid1.Text = tClientes.Fields!IdCliente
                
                With Me.MSFlexGrid1

                    MSFlexGrid1.ColAlignment(1) = flexAlignLeftTop
                    MSFlexGrid1.Col = 0
                    MSFlexGrid1.Text = tClientes.Fields!IdCliente
                    MSFlexGrid1.Col = 1
                    MSFlexGrid1.Text = tClientes.Fields!RazonSocial
                    MSFlexGrid1.Col = 2
                    If tClientes.Fields!CUIT <> "" Then
                       MSFlexGrid1.Text = tClientes.Fields!CUIT
                    End If
                    MSFlexGrid1.Col = 3
                    MSFlexGrid1.Text = tClientes.Fields!Domicilio
                    MSFlexGrid1.Col = 4
                    MSFlexGrid1.Text = tClientes.Fields!Localidad
                    MSFlexGrid1.Col = 5
                    If tClientes.Fields!CP <> "" Then
                        MSFlexGrid1.Text = tClientes.Fields!CP
                    End If
                    MSFlexGrid1.Col = 6
                    MSFlexGrid1.Text = tClientes.Fields!Prov
                    MSFlexGrid1.Col = 7
                    MSFlexGrid1.Text = tClientes.Fields!PorcentajeDescuento
                    
                End With
                linea2 = linea2 + 1
                tClientes.MoveNext
        Loop
    End If
MSFlexGrid1.Col = 4
'MSFlexGrid1.Sort = flexSortGenericAscending


End Sub

Private Sub TextCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'Call calculoprecios
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If

    
End Sub

Private Sub TextCodigoCliente_GotFocus()
    TextCodigoCliente.SelLength = Len(TextCodigoCliente.Text)
End Sub

Private Sub TextDescripcion_GotFocus()

    TextDescripcion.SelLength = Len(TextDescripcion.Text)

End Sub

Private Sub TextDescripcion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If


End Sub

Private Sub TextDescuentoCliente_GotFocus()

    TextDescuentoCliente.SelLength = Len(TextDescuentoCliente.Text)
    
End Sub

Private Sub TextDescuentoCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextFechaPresupuesto_GotFocus()

    TextFechaPresupuesto.SelLength = Len(TextFechaPresupuesto.Text)

End Sub

Private Sub TextFechaPresupuesto_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextFechaPresupuesto_LostFocus()

    TextFechaPresupuesto.SelLength = Len(TextFechaPresupuesto.Text)

End Sub


Private Sub TextNumeroPresupuesto_GotFocus()
    TextNumeroPresupuesto.SelLength = Len(TextNumeroPresupuesto.Text)
End Sub

Private Sub TextNumeroPresupuesto_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextNumeroPresupuesto_LostFocus()

    If TextNumeroPresupuesto.Text <> "" Then
        Set tPresupuestoC = db.OpenRecordset("PresupuestoC", dbOpenTable)
        
        tPresupuestoC.Index = "PrimaryKey"
        
        tPresupuestoC.Seek "=", TextNumeroPresupuesto.Text
        
        If Not tPresupuestoC.NoMatch Then
            MsgBox ("NUMERO DE PRESUPUESTO EXISTENTE !!")
        End If
        
        tPresupuestoC.Close
    End If

End Sub

Private Sub TextPorDexcuento_GotFocus()

    TextPorDexcuento.SelLength = Len(TextPorDexcuento.Text)

End Sub

Private Sub TextPorDexcuento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub
Private Sub TextPorDexcuento_LostFocus()
     
    'If KeyAscii = 13 Then
        Call calculoprecios
    'End If

End Sub
Private Sub TextCantidad_LostFocus()
    
    If TextCantidad.Text = "" Then
        A = MsgBox("NO PUEDE DEJAR LA CANTIDAD EN BLANCO", vbCritical, "E R R O R ! ! !")
        TextCantidad.SetFocus
    End If
    
    Call calculoprecios
    
End Sub

Private Sub TextPrecioUnitario_GotFocus()

    TextPrecioUnitario.SelLength = Len(TextPrecioUnitario.Text)
    
End Sub

Private Sub TextSaldoCliente_GotFocus()

    TextSaldoCliente.SelLength = Len(TextSaldoCliente.Text)

End Sub

Private Sub TextSaldoCliente_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextTotalLineaProd_GotFocus()

    TextTotalLineaProd.SelLength = Len(TextTotalLineaProd.Text)

End Sub

Private Sub TextTotalLineaProd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextUnidadMedida_GotFocus()

    TextUnidadMedida.SelLength = Len(TextUnidadMedida.Text)

End Sub

Private Sub TextUnidadMedida_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If


End Sub
Private Sub TextPrecioUnitario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
   
    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
   
End Sub
Private Sub TextPrecioUnitario_LostFocus()
     
    'If KeyAscii = 13 Then
        Call calculoprecios
    'End If

End Sub
Private Sub calculoprecios()

    Dim precioUnitario As Double
    Dim porcentaje As Double
    Dim cantidad As Long
    Dim total As Double
    Dim porcentajePrecioUnitario As Double
    Dim totalLinea As Double
    
    'If KeyAscii = 13 Then
    If TextCantidad.Text <> "" Then
         
        precioUnitario = Format(TextPrecioUnitario.Text, "#00.00")
        If TextPorDexcuento.Text <> "" Then
            porcentaje = Val(TextPorDexcuento.Text)
        Else
            porcentaje = Val(TextDescuentoCliente.Text)
        End If
        cantidad = Val(TextCantidad.Text)
        total = (precioUnitario * cantidad)
        porcentajePrecioUnitario = ((precioUnitario * porcentaje) / 100) * cantidad
        totalLinea = total - ((total * porcentaje) / 100)
        TextTotalLineaProd.Text = Format(totalLinea, "#00.00")
        TextPorDesc.Text = Format(porcentajePrecioUnitario, "#00.00")
             
    End If
         
    ' End If
    

End Sub

Private Sub ComboVendedor_Click()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    Dim busca1 As String, busca2 As String
    busca1 = RTrim(LTrim(ComboVendedor.Text))
    busca2 = busca1 + "z"
    
    rstEmpleado.FindFirst "Nombre >= '" & busca1 & "' and Nombre <= '" & busca2 & "'"
    
    If rstEmpleado.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Clientes", vbCritical, "Final de la busqueda")
       ComboVendedor.Text = ""
       Call blanco
       ComboVendedor.SetFocus
    End If
     
   LegajoEmpleado = rstEmpleado.Fields!Legajo
  
End Sub


'Private Sub FG1_KeyPress(KeyAscii As Integer)

'    Dim precioUnitario As Double
'    Dim cantidad As Integer
'    Dim porcentaje As Double
'    Dim total
'    Dim totalLinea As Double
'    Dim totalGrilla
'    Dim subtotalPresuForm
'    Dim porcentajePrecioUnitario As Double
'    Dim descuentoPresup As Double
'    Dim totalPresuForm As Double
'    Dim iva As Double
'    Dim impuesto As Double
'    Dim percepcion As Double
'    Dim columnaSeis As Integer
'    Dim columnaSiete As Integer
'    Dim bandera As Integer

'    ruta = App.Path & "\DB_SPC_SI.mdb"
    
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
     
    
'    If KeyAscii >= 32 And KeyAscii <= 127 Then
'        FG1.Text = FG1.Text & Chr(KeyAscii)
'    End If

'    Select Case KeyAscii
'       Case 13
      
     
'            FG1.Col = 0
'            codigoprodMA = UCase(FG1.Text)
                   
             '******* Busco Producto
            
'           Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
        
   
'           tProductos.Index = "PrimaryKey"
           
'           tProductos.MoveFirst
'           tProductos.Seek "=", codigoprodMA
           
'           If Not tProductos.NoMatch Then
'                produ = tProductos!CodProd
                'MsgBox (DepoOrigen)
'                 Dim busca1 As String, busca2 As String
'                 busca1 = RTrim(LTrim(codigoprodMA))
'                 busca2 = busca1 + "z"
                                     
'                 rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
'                 codigoProdTabla = rstProductos.Fields!CodProd
            
'                cantidadProducto = rstProductos.Fields!Stock
'                FG1.Col = 1
'                FG1.Text = rstProductos.Fields!Descripcion
'                FG1.Col = 2
'                FG1.Text = rstProductos.Fields!UnidadMedida
'                FG1.Col = 3
'                FG1.Text = Format(rstProductos.Fields!PrecioUnitarioPresupuesto, "#00.00")
'                FG1.Col = FG1.Col + 2
    
'           Else
'                  mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
'                 codigoprodMA = ""
'                 FG1.Col = 1
'                 FG1.Text = ""
'                 FG1.Col = 2
'                 FG1.Text = ""
'                 FG1.Col = 3
'                 FG1.Text = ""
'                 FG1.Col = 4
'                 FG1.Text = ""
'                 FG1.Col = 5
'                 FG1.Text = ""
'                 FG1.Col = 6
'                 FG1.Text = ""
'                 FG1.Col = 7
'                 FG1.Text = ""
'                 FG1.Col = 0
'                 FG1.Text = ""
'                 FG1.SetFocus
'                 bandera = 1
'           End If
              
'           tProductos.Close
            
'            If bandera <> 1 Then
'                Call muestrodatosproductos
'                FG1.Col = FG1.Col + 2
'                 codigoprodMA = ""
'            End If
            
            '***********************
          
           '*** descuento
'           If FG1.Col = 4 And FG1.Text <> "" Then
'                If KeyAscii = 13 Then
'                   'FG1.Col = FG1.Col + 1
'                   FG1.Col = 3
'                   precioUnitario = Val(FG1.Text)
'                   FG1.Col = 4
'                   porcentaje = Val(FG1.Text)
'                   FG1.Col = 5
'                   cantidad = Val(FG1.Text)
'                   total = (precioUnitario * cantidad)
'                   porcentajePrecioUnitario = ((precioUnitario * porcentaje) / 100) * cantidad
'                   totalLinea = total - ((total * porcentaje) / 100)
'                   FG1.Col = 6
'                   FG1.Text = Format(totalLinea, "#00.00")
'                   FG1.Col = 7
'                   FG1.Text = Format(porcentajePrecioUnitario, "#00.00")
'                End If
'           End If
'
'           '**** cantidad
'           If FG1.Col = 5 And FG1.Text <> "" Then
'                If KeyAscii = 13 Then
'                    FG1.Col = FG1.Col + 1
'                    FG1.Col = 3
'                    precioUnitario = Format(FG1.Text, "#00.00")
'                    FG1.Col = 4
'                    If FG1.Text <> "" Then
'                        porcentaje = Val(FG1.Text)
'                    Else
'                        porcentaje = TextDescuentoCliente.Text
'                    End If
'                    FG1.Col = 5
'                    cantidad = Val(FG1.Text)
'                    '*** verfico stock de producto
'                    'If cantidad > cantidadProducto Then
'                    '    MsgBox "La cantidad ingresada supera al Stock Actual: " & cantidadProducto & ""
'                    '    FG1.Col = 5
'                    '    FG1.Text = ""
'                    '    FG1.SetFocus
'                    'Else
'                        total = (precioUnitario * cantidad)
'                        porcentajePrecioUnitario = ((precioUnitario * porcentaje) / 100) * cantidad
'                        totalLinea = total - ((total * porcentaje) / 100)
'                        FG1.Col = 6
'                        FG1.Text = Format(totalLinea, "#00.00")
'                        FG1.Col = 7
'                        FG1.Text = Format(porcentajePrecioUnitario, "#00.00")
'                    'End If
'                End If
'            End If
'
'            '**** suma total linea
'
'            columnaSeis = 6
'
'            total = SumarTotalGrilla(FG1, columnaSeis)
'            subtotalPresuForm = total
'
'            TextSubtotalPresupuesto.Text = Format(total, "#00.00")
'
'            '**** suma descuentos
'
'            columnaSiete = 7
'
'            porcentajePrecioUnitario = SumarTotalDescuentos(FG1, columnaSiete)
'            descuentoPresup = porcentajePrecioUnitario
'
'            TextDescuentos.Text = Format(descuentoPresup, "#0.00")
'
'
'            '**** calculo total factura
'
'            totalPresuForm = subtotalPresuForm
'            'totalPresuForm = (subtotalPresuForm - descuentoPresup)
'
'            TextTotalPresupuesto.Text = Format(totalPresuForm, "#00.00")
'
'            If CDec(totalPresuForm) <> 0 Then
'                 BotonGrabar.Enabled = True
'                 BotonImprimir.Enabled = True
'                 BotonPago.Enabled = True
'             End If
'
'            If FG1.Col = 7 And FG1.Text <> "" Then
'                FG1.Col = 0
'                'If FG1.Row < 2 Then
'                    FG1.Row = FG1.Row + 1
'                    FG1.SetFocus
'                    BotonGrabar.Enabled = True
'                    BotonImprimir.Enabled = True
'                'End If
'            End If
'
'
'       Case vbKeyBack
'
'            If Len(FG1) >= 1 Then
'               FG1 = Left$(FG1, Len(FG1) - 1)
'            Else
'                KeyAscii = 0
'            End If
'
'       End Select
'
'
'       codigoprod = ""
'
'End Sub
Private Sub calculos()

Dim precioUnitario As Double
    Dim cantidad As Integer
    Dim porcentaje As Double
    Dim total
    Dim totalLinea As Double
    Dim totalGrilla
    Dim subtotalPresuForm
    Dim porcentajePrecioUnitario As Double
    Dim descuentoPresup As Double
    Dim totalPresuForm As Double
    Dim iva As Double
    Dim impuesto As Double
    Dim percepcion As Double
    Dim columnaSeis As Integer
    Dim columnaSiete As Integer
    
    

       
            '**** suma total linea
            
            columnaSeis = 6
             
            total = SumarTotalGrilla(FG1, columnaSeis)
            subtotalPresuForm = total
                                    
            TextSubtotalPresupuesto.Text = Format(total, "#00.00")
            
            '**** suma descuentos
            
            columnaSiete = 7
             
            porcentajePrecioUnitario = SumarTotalDescuentos(FG1, columnaSiete)
            descuentoPresup = porcentajePrecioUnitario
                                    
            TextDescuentos.Text = Format(descuentoPresup, "#0.00")
            
           
            '**** calculo total factura
            
            totalPresuForm = (subtotalPresuForm - descuentoPresup)
            
            TextTotalPresupuesto.Text = Format(totalPresuForm, "#00.00")
            
    
End Sub


Private Sub muestrodatosproductos()

    cantidadProducto = rstProductos.Fields!Stock
    FG1.Col = 1
    FG1.Text = rstProductos.Fields!Descripcion
    FG1.Col = 2
    FG1.Text = rstProductos.Fields!UnidadMedida
    FG1.Col = 3
    FG1.Text = Format(rstProductos.Fields!PrecioUnitarioPresupuesto, "#00.00")
    'FG1.Col = 4
    'FG1.Text = TextDescuentocliente.Text
           
End Sub

Function SumarTotalGrilla(MSFlexGrid3 As Object, columnaSeis As Integer) As Currency

 On Error GoTo error_function
  
    With MSFlexGrid3
        Dim totalLinea As Currency
        Dim I As Long
           
        If columnaSeis > MSFlexGrid3.Cols Then
           MsgBox "Columna no válida", vbExclamation
           Exit Function
        End If
          
        ' recorrer  las filas de la grilla
        For I = 1 To MSFlexGrid3.Rows - 1
            ' comprobar que el dato es de tipo numérico con la función IsNumeric de vb
            If IsNumeric(MSFlexGrid3.TextMatrix(I, columnaSeis)) Then
                ' Sumar, obteniendo el valor de la celda con TextMatrix
                totalLinea = totalLinea + MSFlexGrid3.TextMatrix(I, columnaSeis)
            End If
        Next
           
        ' retornar el total de la suma a la función
       SumarTotalGrilla = totalLinea

    End With
    
Exit Function
error_function:
  
MsgBox Err.Description, vbCritical, "error al sumar"
                        
       
End Function

Function SumarTotalDescuentos(MSFlexGrid3 As Object, columnaSiete As Integer) As Currency

 On Error GoTo error_function
  
    With MSFlexGrid3
        Dim totalDescuento As Currency
        Dim I As Long
           
        If columnaSiete > MSFlexGrid3.Cols Then
           MsgBox "Columna no válida", vbExclamation
           Exit Function
        End If
          
        ' recorrer  las filas de la grilla
        For I = 1 To MSFlexGrid3.Rows - 1
            ' comprobar que el dato es de tipo numérico con la función IsNumeric de vb
            If IsNumeric(MSFlexGrid3.TextMatrix(I, columnaSiete)) Then
                ' Sumar, obteniendo el valor de la celda con TextMatrix
                totalDescuento = totalDescuento + MSFlexGrid3.TextMatrix(I, columnaSiete)
            End If
        Next
           
        ' retornar el total de la suma a la función
       SumarTotalDescuentos = totalDescuento

    End With
    
Exit Function
error_function:
  
MsgBox Err.Description, vbCritical, "error al sumar"
                        
       
End Function

Private Sub Form_Load()

    FormPresupuesto.Height = 10110
    FormPresupuesto.Width = 12135
    FormPresupuesto.Top = 1000
    FormPresupuesto.Left = 50
        
    Call SeteoGrilla
    
    Fila = 1
    renglon = 16
      
    Call Cargo
    Call buscoarticulo
    
    TextFechaPresupuesto.Text = Format(Date, "dd/mm/yyyy")
    
    'bansera = 0
    modificaStock = 1
End Sub
Private Sub buscoarticulo()

    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    
   
    Do While Not rstProductos.EOF
        
       ComboArticulo.AddItem rstProductos!CodProd
       rstProductos.MoveNext
       
    Loop
    
    
End Sub

Sub SeteoGrilla()
    
    FG1.Row = 0
    FG1.Col = 0
    
    FG1.ColWidth(0) = 1000
    FG1.CellFontBold = True
    FG1.ColAlignment(0) = flexAlignCenterCenter
    FG1.Text = "Articulo"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 4400
    FG1.CellFontBold = True
    FG1.Text = "Descripción"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 1000
    FG1.CellFontBold = True
    FG1.Text = "UM"
    FG1.ColAlignment(2) = flexAlignCenterCenter
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1100
    FG1.CellFontBold = True
    FG1.Text = "Precio Unit."
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 720
    FG1.CellFontBold = True
    FG1.Text = "% Desc."
    FG1.ColAlignment(4) = flexAlignCenterCenter
    
    FG1.Col = 5
    FG1.ColWidth(5) = 900
    FG1.CellFontBold = True
    FG1.Text = "Cantidad"
    FG1.ColAlignment(5) = flexAlignCenterCenter
        
    FG1.Col = 6
    FG1.ColWidth(6) = 1100
    FG1.CellFontBold = True
    FG1.Text = "Total Línea"
    FG1.ColAlignment(6) = flexAlignCenterCenter
    
    FG1.Col = 7
    FG1.ColWidth(7) = 1100
    FG1.CellFontBold = True
    FG1.Text = "Importe Descuento"
    
    FG1.Col = 8
    FG1.ColWidth(8) = 0
    FG1.CellFontBold = True
    FG1.Text = "Item"
    
    FG1.Row = 1
    item = 1
    linea = 1
    Do While FG1.Row <= 14
        FG1.Col = 8
        FG1.Text = item
        item = (item + 1)
        FG1.Row = (Val(FG1.Row) + 1)
    Loop
    
    
End Sub

Private Sub Cargo()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    rstEmpleado.MoveFirst
    Do While Not rstEmpleado.EOF
        ComboVendedor.AddItem rstEmpleado!nombre
        rstEmpleado.MoveNext
    Loop

End Sub

Private Sub Busco()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    Call titulos
    
    Dim busca1 As String, busca2 As String
    busca1 = RTrim(LTrim(TextApellidoNombre.Text))
    busca2 = busca1 + "z"
    
    rstCliente.FindFirst "Razonsocial >= '" & busca1 & "' and RazonSocial <= '" & busca2 & "'"
    
    If rstCliente.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Clientes", vbCritical, "Final de la busqueda")
       TextApellidoNombre.Text = ""
       Call blanco
       TextApellidoNombre.SetFocus
    End If
     
    linea2 = 1
    Do While Not rstCliente.NoMatch
        MSFlexGrid1.AddItem " "
        MSFlexGrid1.Row = linea2
       
            MSFlexGrid1.Col = 0
            MSFlexGrid1.Text = rstCliente.Fields!IdCliente
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = rstCliente.Fields!RazonSocial
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = rstCliente.Fields!CUIT
            MSFlexGrid1.Col = 3
            MSFlexGrid1.Text = rstCliente.Fields!Domicilio
            MSFlexGrid1.Col = 4
            MSFlexGrid1.Text = rstCliente.Fields!Localidad
            MSFlexGrid1.Col = 5
            MSFlexGrid1.Text = rstCliente.Fields!CP
            MSFlexGrid1.Col = 6
            MSFlexGrid1.Text = rstCliente.Fields!Prov
            MSFlexGrid1.Col = 7
            MSFlexGrid1.Text = rstCliente.Fields!PorcentajeDescuento
            linea2 = linea2 + 1
      
       rstCliente.FindNext "RazonSocial >= '" & busca1 & "' and RazonSocial <= '" & busca2 & "'"
       
    Loop
    
    FG1.Enabled = True
    
End Sub

Private Sub titulos()

    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.Text = "Codigo"
    MSFlexGrid1.ColWidth(0) = 900
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.Text = "Apellido y Nombre"
    MSFlexGrid1.ColWidth(1) = 4700
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.Text = "CUIT"
    MSFlexGrid1.ColWidth(2) = 1200
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.Text = "Direccion"
    MSFlexGrid1.ColWidth(3) = 0
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.Text = "Localidad"
    MSFlexGrid1.ColWidth(4) = 0
    
    MSFlexGrid1.Col = 5
    MSFlexGrid1.Text = "CP"
    MSFlexGrid1.ColWidth(5) = 0
    
    MSFlexGrid1.Col = 6
    MSFlexGrid1.Text = "Provincia"
    MSFlexGrid1.ColWidth(6) = 0
    
    MSFlexGrid1.Col = 7
    MSFlexGrid1.Text = "Porcentaje Descuento"
    MSFlexGrid1.ColWidth(7) = 0

    
 End Sub

Private Sub MSFlexGrid1_Click()
   
    
    MSFlexGrid1.Col = 0
    TextCodigoCliente.Text = MSFlexGrid1.Text
    
    MSFlexGrid1.Col = 1
    TextApellidoNombre.Text = MSFlexGrid1.Text
    
    MSFlexGrid1.Col = 2
    TextCuit.Text = MSFlexGrid1.Text
    
    MSFlexGrid1.Col = 3
    TextDireccion.Text = MSFlexGrid1.Text
    
    MSFlexGrid1.Col = 4
    TextLocalidad.Text = MSFlexGrid1.Text
    
    MSFlexGrid1.Col = 5
    TextCodigoPostal.Text = MSFlexGrid1.Text
    
    MSFlexGrid1.Col = 6
    TextProvincia.Text = MSFlexGrid1.Text
        
    'MSFlexGrid1.Col = 7
    'descuentos = MSFlexGrid1.Text
    
    Call buscocuilyvendedor
    
    MSFlexGrid1.Visible = False
    
    FG1.Enabled = True

End Sub

Private Sub TextApellidoNombre_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call Busco
    End If
    
End Sub

Private Sub blanco()

    TextCodigoCliente.Text = ""
    TextApellidoNombre.Text = ""
    TextCuit.Text = ""
    TextDireccion.Text = ""
    TextLocalidad.Text = ""
    TextCodigoPostal.Text = ""
    TextProvincia.Text = ""
    
End Sub

Private Sub TextCodigoCliente_KeyPress(KeyAscii As Integer)
   
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
   

    If KeyAscii = 13 Then
        If TextCodigoCliente.Text = "" Then
            TextApellidoNombre.SetFocus
        Else
            CodigoClie = Val(TextCodigoCliente.Text)
      
            rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCliente.Fields!IdCliente <> Val(TextCodigoCliente.Text) Then
                mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                TextCodigoCliente.Text = ""
                Call blanqueototal
                TextCodigoCliente.SetFocus
            Else
                TextCodigoCliente.Text = rstCliente.Fields!IdCliente
                TextApellidoNombre.Text = rstCliente.Fields!RazonSocial
                MSFlexGrid1.Visible = False
                'If Not rstCliente.Fields!Cuit Then
                If rstCliente.Fields!CUIT <> "" Then
                    TextCuit.Text = rstCliente.Fields!CUIT
                End If
                TextDireccion.Text = rstCliente.Fields!Domicilio
                TextLocalidad.Text = rstCliente.Fields!Localidad
                If rstCliente.Fields!CP <> Null Then
                    TextCodigoPostal.Text = rstCliente.Fields!CP
                End If
                TextProvincia.Text = rstCliente.Fields!Prov
                TextDescuentoCliente.Text = rstCliente.Fields!PorcentajeDescuento
                vendedorCliente = rstCliente.Fields!Vendedor
                Call buscocuilyvendedor
            End If
        End If
        TextNumeroPresupuesto.Text = ""
    End If
    
    If TextNumeroPresupuesto <> "" Then
        FG1.Enabled = True
    Else
        FG1.Enabled = False
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub buscocuilyvendedor()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
    '*** Busco CUIT
    
    CodigoClie = Val(TextCodigoCliente.Text)
      
    rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
    
    'If Not rstCliente.Fields!Cuit Then
    If rstCliente.Fields!CUIT <> "" Then
        TextCuit.Text = rstCliente.Fields!CUIT
    End If
    codigovendedor = rstCliente!Vendedor
      
    TextDescuentoCliente.Text = rstCliente.Fields!PorcentajeDescuento
    
    '*** Busco Vendedor
    
    CodigoVend = codigovendedor
      
    rstEmpleado.FindFirst "Legajo >= '" & CodigoVend & "'"
    
    LegajoEmpleado = rstEmpleado.Fields!Legajo
    ComboVendedor.Text = rstEmpleado.Fields!nombre
    
    '*** Busco Saldo
    
   rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
    
   TextSaldoCliente.Text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
   
   BotonDomicilio.Enabled = True
   BotonNueva.Enabled = True
   BotonNueva.SetFocus
    
End Sub

Private Sub TextCuit_Change()

    If TextCuit.Text <> "" Then
        BotonNueva.Enabled = True
    End If
        
End Sub



Private Sub TextProvincia_Change()

    If TextProvincia.Text <> "" Then
        ComboVendedor.SetFocus
    End If
End Sub


