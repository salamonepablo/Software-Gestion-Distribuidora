VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormNotaCredito 
   BackColor       =   &H00008080&
   Caption         =   "Nota de Credito"
   ClientHeight    =   9630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1335
      Left            =   4320
      TabIndex        =   50
      Top             =   1800
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   2355
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00008080&
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
      TabIndex        =   64
      Top             =   2520
      Width           =   11655
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
      Begin VB.TextBox TextPorDescuento 
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
         Left            =   10440
         TabIndex        =   9
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox ComboArticulo 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   600
         Width           =   1215
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
         TabIndex        =   65
         Top             =   120
         Visible         =   0   'False
         Width           =   1455
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   72
         Top             =   360
         Width           =   990
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   71
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   70
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   69
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   68
         Top             =   360
         Width           =   300
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   67
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   66
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008080&
      Height          =   975
      Left            =   120
      TabIndex        =   51
      Top             =   1560
      Width           =   11655
      Begin VB.OptionButton opContado 
         BackColor       =   &H00008080&
         Caption         =   "Contado"
         Height          =   255
         Left            =   9600
         TabIndex        =   75
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opCtaCte 
         BackColor       =   &H00008080&
         Caption         =   "Cta Cte"
         Height          =   255
         Left            =   10560
         TabIndex        =   74
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox TextTipoNotaCredito 
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
         Left            =   1800
         TabIndex        =   57
         Top             =   480
         Width           =   375
      End
      Begin VB.TextBox TextNumeroNotaCredito 
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
         Left            =   360
         TabIndex        =   1
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextFechaNotaCredito 
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
         Left            =   2640
         TabIndex        =   56
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox ComboVendedor 
         Height          =   315
         Left            =   4800
         TabIndex        =   55
         Top             =   480
         Width           =   1455
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
         Left            =   6840
         TabIndex        =   54
         Top             =   480
         Width           =   735
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
         Left            =   7920
         TabIndex        =   53
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox CheckModificaStock 
         BackColor       =   &H00008080&
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
         Left            =   9600
         TabIndex        =   52
         Top             =   600
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Tipo"
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
         TabIndex        =   63
         Top             =   240
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Nº Nota Credito"
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
         TabIndex        =   62
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Fecha Nota Credito"
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
         TabIndex        =   61
         Top             =   240
         Width           =   1665
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         Left            =   4800
         TabIndex        =   60
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         Left            =   6720
         TabIndex        =   59
         Top             =   240
         Width           =   930
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         Left            =   8280
         TabIndex        =   58
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00008080&
      Height          =   1455
      Left            =   120
      TabIndex        =   26
      Top             =   120
      Width           =   11655
      Begin VB.TextBox TextProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   27
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   38
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
         TabIndex        =   11
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   30
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   29
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   28
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   37
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   36
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   35
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   34
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   33
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   32
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         TabIndex        =   31
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Height          =   1095
      Left            =   120
      TabIndex        =   40
      Top             =   8520
      Width           =   11655
      Begin VB.CommandButton BotonGrabar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   750
         Left            =   1560
         TabIndex        =   48
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonPago 
         Caption         =   "&Pago"
         Enabled         =   0   'False
         Height          =   750
         Left            =   3240
         TabIndex        =   46
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   5280
         TabIndex        =   44
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonCancelar 
         Caption         =   "&Cancelar"
         Height          =   750
         Left            =   4080
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton BotonImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   750
         Left            =   2400
         TabIndex        =   42
         Top             =   240
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton BotonNueva 
         Caption         =   "&Nueva"
         Enabled         =   0   'False
         Height          =   750
         Left            =   720
         TabIndex        =   41
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00008080&
      Height          =   3615
      Left            =   120
      TabIndex        =   39
      Top             =   3600
      Width           =   11655
      Begin VB.CommandButton BotonEliminarfila 
         Caption         =   "&Eliminar Fila"
         Height          =   495
         Left            =   10800
         TabIndex        =   49
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton BotonBuscarProducto 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   10800
         TabIndex        =   47
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   5741
         _Version        =   393216
         Rows            =   16
         Cols            =   9
         FixedCols       =   0
         Enabled         =   0   'False
         GridLines       =   2
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00008080&
      Height          =   1215
      Left            =   120
      TabIndex        =   12
      Top             =   7320
      Width           =   11655
      Begin VB.CommandButton BotonCalcular 
         Caption         =   "&Calcular"
         Height          =   375
         Left            =   9960
         TabIndex        =   73
         Top             =   840
         Width           =   1215
      End
      Begin VB.TextBox TextSubTotalNotaCredito 
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
         TabIndex        =   19
         Top             =   480
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
         Left            =   1560
         TabIndex        =   18
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextPercepcionIIBB 
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
         Left            =   2400
         TabIndex        =   17
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextAlicuota 
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
         Left            =   4200
         TabIndex        =   16
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextImpuesto 
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
         Left            =   6000
         TabIndex        =   15
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Textiva 
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
         Left            =   7920
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextTotalNotaCredito 
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
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Alicuota:"
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
         Left            =   4440
         TabIndex        =   45
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Subtotal Nota Credito:"
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
         TabIndex        =   25
         Top             =   240
         Width           =   1905
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
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
         Left            =   1680
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Percepción IIBB:"
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
         TabIndex        =   23
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Impuesto:"
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
         TabIndex        =   22
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "iva:"
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
         TabIndex        =   21
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Total Nota Credito:"
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
         TabIndex        =   20
         Top             =   240
         Width           =   1635
      End
   End
End
Attribute VB_Name = "FormNotaCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstNotaCreditoC As DAO.Recordset
 Dim rstNotaCreditoD As DAO.Recordset
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
 Dim num As Integer
 Dim Fila As Integer
 Dim fila2 As Integer
 Dim renglon As Integer

 

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
        FG1.Text = Format(TextPorDescuento.Text, "#0.00")
        FG1.Col = 5
        FG1.Text = TextCantidad.Text
        FG1.Col = 6
        FG1.Text = Format(TextTotalLineaProd.Text, "#0.00")
        FG1.Col = 7
        FG1.Text = Format(TextPorDesc.Text, "#0.00")
                
        Fila = Fila + 1
        fila2 = fila2 + 1
            
    
        Call calculototalnotacredito
        
'        ComboArticulo.Text = ""
        TextDescripcion.Text = ""
        TextUnidadMedida.Text = ""
        TextPrecioUnitario.Text = ""
        TextPorDescuento.Text = ""
        TextCantidad.Text = ""
        TextTotalLineaProd.Text = ""
        TextPorDesc.Text = ""
        
        ComboArticulo.SetFocus
    
  End If

End Sub
Private Sub calculototalnotacredito()

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
    
    If TextTipoNotaCredito.Text = "A" Then
       Textiva.Text = Format(iva, "#00.00")
    End If

    '**** suma total linea
             
    'total = Val(TextTotalLineaProd.Text) + Val(TextSubTotalNotaCredito.Text)
    If (TextSubTotalNotaCredito.Text = "") Then TextSubTotalNotaCredito.Text = 0
    total = CDbl(TextTotalLineaProd.Text) + CDbl(TextSubTotalNotaCredito.Text)
    subtotalFacturaForm = total
                            
    TextSubTotalNotaCredito.Text = Format(total, "#0.00")
    
    '**** suma descuentos
     
    'porcentajePrecioUnitario = Val(porcentajePrecioUnitario) + Val(TextPorDesc.Text)
    If (TextDescuentos.Text = "") Then TextDescuentos.Text = 0
    porcentajePrecioUnitario = CDbl(TextDescuentos.Text) + CDbl(TextPorDesc.Text)
    descuentoFactura = porcentajePrecioUnitario
                            
    TextDescuentos.Text = Format(descuentoFactura, "#0.00")
    
    '**** calculo alicuota

    TextAlicuota.Text = Format(Alicuota, "#0.00")
                    
   If TextTipoNotaCredito.Text = "A" Then
        percepcion = subtotalFacturaForm * Alicuota / 100
        TextPercepcionIIBB.Text = Format(percepcion, "#0.00")
        
    End If
    
    '**** calculo impuesto
    
    If TextTipoNotaCredito.Text = "A" Then
       impuesto = subtotalFacturaForm * iva / 100
       TextImpuesto.Text = Format(impuesto, "#0.00")
    End If
    
    '**** calculo total factura
    
    'totalFacturaForm = (subtotalFacturaForm - descuentoFactura + percepcion + impuesto)
    
    totalFacturaForm = (subtotalFacturaForm + percepcion + impuesto)
    
    TextTotalNotaCredito.Text = Format(totalFacturaForm, "#0.00")
    
    If CDec(totalFacturaForm) <> 0 Then
         BotonGrabar.Enabled = True
         'BotonImprimir.Enabled = True
         'BotonPago.Enabled = True
    End If

End Sub
Private Sub calculomanual()

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
                      
    total = Format(TextSubTotalNotaCredito.Text, "#0.00")
    
    TextAlicuota.Text = Format(Alicuota, "#0.00")
                    
    percepcion = TextPercepcionIIBB.Text
       
    impuesto = TextImpuesto.Text
    
    totalFacturaForm = (total + subtotalFacturaForm + percepcion + impuesto)
    
    TextTotalNotaCredito.Text = Format(totalFacturaForm, "#0.00")
    
    If CDec(totalFacturaForm) <> 0 Then
         BotonGrabar.Enabled = True
         'BotonImprimir.Enabled = True
         'BotonPago.Enabled = True
    End If


End Sub
Private Sub BotonAgregar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonBuscarProducto_Click()

    FormBusquedaProducto.Show

End Sub

Private Sub BotonCalcular_Click()
     Call calculomanual
End Sub

Private Sub BotonCancelar_Click()

    Call blanqueototal
    
End Sub

Public Sub blanqueototal()

    TextCodigoCliente.Text = ""
    TextApellidoNombre.Text = ""
    TextCuit.Text = ""
    TextDireccion.Text = ""
    TextLocalidad.Text = ""
    TextCodigoPostal.Text = ""
    TextProvincia.Text = ""
    TextNumeroNotaCredito.Text = ""
    TextTipoNotaCredito.Text = ""
    ComboVendedor.Text = ""
    TextDescuentoCliente.Text = ""
    TextSubTotalNotaCredito.Text = ""
    TextDescuentos.Text = ""
    TextPercepcionIIBB.Text = ""
    TextAlicuota.Text = ""
    TextImpuesto.Text = ""
    Textiva.Text = ""
    TextTotalNotaCredito.Text = ""
    TextSaldoCliente.Text = ""
    ComboVendedor.Text = ""
    TextDescuentoCliente.Text = ""
    'CheckModificaStock.Value = Unchecked
    FG1.Clear
    FG1.Enabled = False
   
    Call SeteoGrilla

End Sub

Private Sub BotonCancelar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

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

Private Sub BotonGrabar_Click()

        Dim descuentoCantidad As Long
        Dim ultimo As Long
        Dim existeNumeroBD As Long
        Dim existeTipoBD As String
        Dim existeNumero As Long
        Dim existeTipo As String
        
'///Declaraciones para FE SPC
        Dim DocNro As Double
        Dim CbteDesde As Double
        Dim CbteHasta As Double
        Dim ImpTotal As Double
        Dim ImpNeto As Double
        Dim MonCotiz As Double
        Dim BaseImpIVA As Double
        Dim ImpIva As Double
        Dim BaseImpTributo As Double
        Dim Alicuota As Double
        Dim ImpAlicuota As Double
        Dim DocTipo As Long
        Dim AlicIVA As Long
        Dim IdTributo As Long
        Dim PtoVta As Long
        Dim TipoComp As Long
        Dim CbteFch As String
        Dim MonId As String
        Dim DescTributo As String
        Dim ImporteExento As Double
'///Declaraciones para FE SPC
                
        ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstNotaCreditoC = db.OpenRecordset("NotaCreditoC", dbOpenDynaset)
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstNotaCreditoD = db.OpenRecordset("NotaCreditoD", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstMovimientosCtaCte = db.OpenRecordset("MovimientosCtaCte", dbOpenDynaset)
        
        
        'buscamos el deposito para descontar el stock
        
        Set tDepositos = db.OpenRecordset("Depositos", dbOpenTable)
           On Error GoTo CapturaErrores
   
           tDepositos.Index = "IndXVendedor"
           
           tDepositos.MoveFirst
           tDepositos.Seek "=", LegajoEmpleado
           
           If Not tDepositos.NoMatch Then
            DepoOrigen = tDepositos!IDDEPOSITO
            'MsgBox (DepoOrigen)
           Else
            A = MsgBox("ERROR !!", vbCritical, "Vendedor sin Depósito Asociado")
           End If
              
           tDepositos.Close
        
    '**************************************************
    
        
        '*** Busco Nota Credito Existente
        
       
        ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db1 = DBEngine.OpenDatabase(ruta)
        
        Set rstFacC = db1.OpenRecordset("NotaCreditoC", dbOpenTable)
        
        rstFacC.Index = "PrimaryKey"
        
        rstFacC.Seek "=", TextTipoNotaCredito, Str(TextNumeroNotaCredito.Text)

        If Not rstFacC.NoMatch Then
            A = MsgBox("Factura Existente", vbCritical, "INFO DEL SISTEMA")
           
            TextNumeroNotaCredito.Text = num
            TextNumeroNotaCredito.SetFocus
        Else
        
        rstFacC.Close
        db1.Close
        
        
     
            rstNotaCreditoC.AddNew
            rstNotaCreditoC.Fields!NroNotaCredito = TextNumeroNotaCredito.Text
            rstNotaCreditoC.Fields!TipoNotaCredito = UCase(TextTipoNotaCredito.Text)
            rstNotaCreditoC.Fields!FechaNotaCredito = TextFechaNotaCredito.Text
            rstNotaCreditoC.Fields!TotalNotaCredito = TextTotalNotaCredito.Text
            If Textiva.Text <> "" Then
                rstNotaCreditoC.Fields!PorcentajeIVA = Textiva.Text
            Else
                rstNotaCreditoC.Fields!PorcentajeIVA = "0,00"
            End If
            
            If TextSubTotalNotaCredito.Text = TextTotalNotaCredito.Text Then TextSubTotalNotaCredito.Text = (TextTotalNotaCredito.Text / 1.21)
            rstNotaCreditoC.Fields!SubTotalNotaCredito = Format(TextSubTotalNotaCredito.Text, "#0.00")
            
            If TextImpuesto.Text <> "" Then
                rstNotaCreditoC.Fields!totalIva = TextImpuesto.Text
            Else
                'rstNotaCreditoC.Fields!totalIva = "0,00"
                rstNotaCreditoC.Fields!totalIva = Format((TextSubTotalNotaCredito.Text * 21) / 100, "#0.00")
            End If
                        
           ' rstNotaCreditoC.Fields!SubTotalNotaCredito = TextSubTotalNotaCredito.Text
           ' If TextImpuesto.Text <> "" Then
           '     rstNotaCreditoC.Fields!totalIva = TextImpuesto.Text
           ' Else
           '     rstNotaCreditoC.Fields!totalIva = "0,00"
           ' End If
            
            If TextAlicuota.Text = "" Then TextAlicuota.Text = 0
            rstNotaCreditoC.Fields!AlicuotaIIBB = TextAlicuota.Text
            If TextPercepcionIIBB.Text <> "" Then
                rstNotaCreditoC.Fields!ImportePercepIIBB = TextPercepcionIIBB.Text
            End If
            rstNotaCreditoC.Fields!CodCliente = TextCodigoCliente.Text
            rstNotaCreditoC.Fields!PorcentajeDesc = TextDescuentoCliente.Text
            rstNotaCreditoC.Fields!ImporteDesc = TextDescuentos.Text
            rstNotaCreditoC.Fields!codVendedor = LegajoEmpleado
            
            If opContado.Value = True Then rstNotaCreditoC.Fields!CondicionVenta = "Contado"
            If opCtaCte.Value = True Then rstNotaCreditoC.Fields!CondicionVenta = "Cuenta Corriente"

            rstNotaCreditoC.Update
            
            FG1.Col = 0
            FG1.Row = 1
            Filas = FG1.Rows
            linea = 1
            Do While linea < Filas
                  
                  FG1.Row = linea
                  FG1.Col = 0
                  If FG1.Text <> "" Then
                        rstNotaCreditoD.AddNew
                    
                        rstNotaCreditoD.Fields!NroNotaCredito = TextNumeroNotaCredito.Text
                        rstNotaCreditoD.Fields!TipoNotaCredito = TextTipoNotaCredito.Text
                    
                        FG1.Col = 0
                        rstNotaCreditoD.Fields!IdCodProd = FG1.Text
                    
                        FG1.Col = 2
                        rstNotaCreditoD.Fields!UnidadMedida = FG1.Text
                        
                        FG1.Col = 3
                        rstNotaCreditoD.Fields!precioUnitario = Format(FG1.Text, "#00.00")
                        
                        FG1.Col = 4
                        des = FG1.Text
                        If des <> "" Then
                           rstNotaCreditoD.Fields!PorcentajeDescuento = Val(des)
                        Else
                           rstNotaCreditoD.Fields!PorcentajeDescuento = Val(TextDescuentoCliente.Text)
                        End If
                        FG1.Col = 5
                        rstNotaCreditoD.Fields!cantidad = Val(FG1.Text)
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
                        rstNotaCreditoD.Fields!totalLinea = Format(FG1.Text, "#00.00")
                        
                        FG1.Col = 7
                        rstNotaCreditoD.Fields!ImporteDescuento = Format(FG1.Text, "#00.00")
                        
                        FG1.Col = 8
                        rstNotaCreditoD.Fields!ItemNotaCredito = Val(FG1.Text)
                         
                        rstNotaCreditoD.Update
                  End If
                  linea = linea + 1
            Loop
            
     '/////////  GENERAMOS LA FACTURA ELECTRONICA DESDE SPC /////////////////////////////
                If rstNotaCreditoD.Fields!IdCodProd = "CHE" Then
                    ImporteExento = rstNotaCreditoD.Fields!precioUnitario * rstNotaCreditoD.Fields!cantidad
                Else
                    ImporteExento = 0
                End If
     
     
                PtoVta = 3
                If TextTipoNotaCredito.Text = "A" Then
                    TipoComp = 3
                 Else
                    If TextTipoNotaCredito.Text = "B" Then
                        TipoComp = 8
                    End If
                End If
                
                If TextCuit.Text <> "" Then
                    DocTipo = 80
                    DocNro = CDbl(TextCuit.Text)
                 Else
                    DocTipo = 96
                    DocNro = 11111111
                End If
                
                CbteDesde = CDbl(TextNumeroNotaCredito.Text)
                CbteHasta = CDbl(TextNumeroNotaCredito.Text)
                CbteFch = CStr(Format(TextFechaNotaCredito.Text, "YYYYMMDD"))
                ImpTotal = Format(CDbl(TextTotalNotaCredito.Text), "Standard")
                
                If TipoComp = 3 Then ImpNeto = Format(CDbl(TextSubTotalNotaCredito.Text), "Standard")
                If TipoComp = 8 Then ImpNeto = Format(CDbl(TextTotalNotaCredito.Text), "Standard")
                
                MonId = "PES"
                MonCotiz = 1
                
                If Textiva.Text = "" Then AlicIVA = 3
                
                If Textiva.Text = "21,00" Then
                    AlicIVA = 5
                Else
                    If Textiva.Text = "10,5" Then
                        AlicIVA = 4
                    End If
                End If
                
                If TipoComp = 3 Then BaseImpIVA = Format(CDbl(TextSubTotalNotaCredito.Text), "Standard")
                If TipoComp = 8 Then BaseImpIVA = ImpNeto
                
                If TextImpuesto.Text <> "" Then
                    ImpIva = Format(CDbl(TextImpuesto.Text), "Standard")
                 Else
                    ImpIva = 0
                End If
                
                IdTributo = 2
                DescTributo = "IIBB"
                
                If TipoComp = 3 Then BaseImpTributo = Format(CDbl(TextSubTotalNotaCredito.Text), "Standard")
                
                Alicuota = Format(CDbl(TextAlicuota.Text), "Standard")
                
                If TextPercepcionIIBB.Text <> "" Then
                    ImpAlicuota = Format(CDbl(TextPercepcionIIBB.Text), "Standard")
                 Else
                    ImpAlicuota = 0
                End If
                Call FacturaElectronicaSPC(PtoVta, DocTipo, DocNro, TipoComp, CbteDesde, CbteHasta, CbteFch, ImpTotal, ImpNeto, MonId, MonCotiz, AlicIVA, BaseImpIVA, ImpIva, IdTributo, DescTributo, BaseImpTributo, Alicuota, ImpAlicuota, ImporteExento)
     '//////////////////////////////////////////////////////////////////////////////////
            
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
                saldoLi1 = Format(-TextTotalNotaCredito.Text, "#0.00")
                rstCtaCte.Fields!SaldoL1 = saldoLi1 + saldo1
                saldo1 = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
                rstCtaCte.Fields!SaldoTotal = saldo1 + saldo2
                rstCtaCte.Update
            End If
        
            
            '*** Grabo Movimientos Cuente corriente
        
            rstMovimientosCtaCte.AddNew
            'rstMovimientosCtaCte.Fields!Fecha = Format(Date, "dd/mm/yyyy")
            rstMovimientosCtaCte.Fields!Fecha = Format(TextFechaNotaCredito.Text, "dd/mm/yyyy")
            rstMovimientosCtaCte.Fields!IdCliente = TextCodigoCliente.Text
            If TextTipoNotaCredito.Text = "A" Then
                rstMovimientosCtaCte.Fields!tipoDoc = "Nota Credito A"
            End If
            If TextTipoNotaCredito.Text = "B" Then
                rstMovimientosCtaCte.Fields!tipoDoc = "Nota Credito B"
            End If
            rstMovimientosCtaCte.Fields!NroDoc = TextNumeroNotaCredito.Text
            rstMovimientosCtaCte.Fields!ImporteLinea1 = -TextTotalNotaCredito.Text
            rstMovimientosCtaCte.Fields!ImporteLinea2 = 0
            rstMovimientosCtaCte.Update
            
            '*** Actualizo Ultimo Numero Cuenta Corriente
            
            Set db = DBEngine.OpenDatabase(ruta)
            Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
        
            Dim busco As String
       
            If TextTipoNotaCredito.Text = "A" Then
                busco = "tNotaCreditoA"
            End If
            
            If TextTipoNotaCredito.Text = "B" Then
                busco = "tNotaCreditoB"
            End If
    
            '*****************
            'Guardo el nro de NC en la variable global para luego poder imprimir
                vNroNCImp = TextNumeroNotaCredito.Text
                vTipoNCImp = TextTipoNotaCredito.Text
            '*****************
    
    
            'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
            rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
            ultimo = rstUltimosNumeros.Fields!UltimoNumero
            
            If ultimo < Val(TextNumeroNotaCredito.Text) Then
                rstUltimosNumeros.Edit
                'If ultimo < rstUltimosNumeros.Fields!UltimoNumero Then
                     rstUltimosNumeros.Fields!UltimoNumero = TextNumeroNotaCredito.Text
                'End If
                rstUltimosNumeros.Update
            End If
            
             BotonGrabar.Enabled = False
        BotonNueva.Enabled = False
        
                 
        modificaStock = 0
        
        respuesta = MsgBox("Desea Imprimir la Nota de Crédito", vbYesNo, "Nota de Crédito")
    
        If respuesta = vbYes Then
          'Call ImprimirNotaCredito
          FormImprimirNC.Show
        End If
        
        'Call blanqueototal
        'Call SeteoGrilla
        
        MSFlexGrid1.Visible = False
        
        'TextCodigoCliente.SetFocus
        End If
        
        
        BotonPago.Enabled = True
        BotonImprimir.Enabled = True
       ' BotonImprimir.SetFocus
       
CapturaErrores:
        
        Select Case Err
            Case 3021
                Resume Next
        End Select
     
    
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
                tS!IDDEPOSITO = IdDepoOrigen
                tS!cantidad = tS.cantidad + FormatNumber(Cant, 2)
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

    Call ImprimirNotaCredito
        
  '  respuesta = MsgBox("Desea Imprimir el Remito", vbYesNo, "Imprimir Remito")
  '  If respuesta = vbYes Then
  '      FormImprimeRemito.Show
  '  End If
    
  '  BotonImprimir.Enabled = False
  '  BotonPago.SetFocus
    
    
End Sub

Private Sub ImprimirNotaCredito()

    X = -4
    Y = -4
    renglon = 0
    vNroRemito = "0003- "
    '& TextNumeroRemito.Text
    
    With Printer
        
        'On Error GoTo CapturaErrores
            .Copies = 2
        'Seteo escala a mm
            .ScaleMode = 6
        
        'Imprimir Fecha
            .CurrentX = X + 120
            .CurrentY = Y + 27
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print Format(TextFechaNotaCredito.Text, "DD/MM/YYYY")
        
        'Imprimir Nombres
            .CurrentX = X + 37
            .CurrentY = Y + 54
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = True
            Printer.Print TextApellidoNombre.Text
            
        'Imprimir Direccion
            .CurrentX = X + 37
            .CurrentY = Y + 60
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextDireccion.Text
            
        'Imprimir Localidad
            .CurrentX = X + 37
            .CurrentY = Y + 65
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextLocalidad.Text
            
        'Imprimir CUIT
            .CurrentX = X + 125
            .CurrentY = Y + 67
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextCuit.Text
            
        'Imprimir Marca Responsable Inscripto
            .CurrentX = X + 57
            .CurrentY = Y + 70
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca Contado
            .CurrentX = X + 70
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca CtaCte
            .CurrentX = X + 100
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Nro Remito
            .CurrentX = X + 138
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 9
            .FontBold = False
            Printer.Print vNroRemito
            
        'Imprimir Detalle
            
            sqlfc = "SELECT * FROM NotaCreditoC WHERE TipoNotaCredito='" & TextTipoNotaCredito.Text & "' AND NroNotaCredito=" & TextNumeroNotaCredito.Text & " ORDER By NroNotaCredito"
            vsqlFD = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & TextTipoNotaCredito.Text & "' AND NroNotaCredito=" & TextNumeroNotaCredito.Text & " ORDER By NroNotaCredito"
            
            Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
                        
            Set NotaCC = BaseSPC.OpenRecordset(sqlfc, dbOpenDynaset)
            Set NotaCD = BaseSPC.OpenRecordset(vsqlFD, dbOpenDynaset)
            
           
            NotaCC.MoveFirst
            NotaCD.MoveFirst
                
                    While Not NotaCD.EOF
                        'Imprimo el detalle
                            .CurrentX = X + 20
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Printer.Print NotaCD!cantidad
                            
                        'Detalle
                            .CurrentX = X + 40
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Printer.Print NotaCD!IdCodProd & Chr(9) & Descripcion(NotaCD!IdCodProd)
                        
                        'Precio
                            .CurrentX = X + 123
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            PU = CDbl(NotaCD!precioUnitario) - (CDbl(NotaCD!precioUnitario) * CDbl(NotaCD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            
                            Printer.Print PU
                        
                        'Importe
                            .CurrentX = X + 143
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Printer.Print NotaCD!totalLinea
                        
                         renglon = renglon + 5
                            
                        NotaCD.MoveNext
                    Wend
           
            'Importe SubTotal
                .CurrentX = X + 143
                .CurrentY = Y + 176
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print NotaCC!SubTotalNotaCredito
                
            'Alicuota IVA
                .CurrentX = X + 131
                .CurrentY = Y + 182
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print "21"
            
            'Importe IVA
                .CurrentX = X + 143
                .CurrentY = Y + 182
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print NotaCC!totalIva
            
            If NotaCC!ImportePercepIIBB > 0 Then
                'Alicuota IIBB
                    .CurrentX = X + 123
                    .CurrentY = Y + 187
                    .Font = "Courier New"
                    .FontSize = 8
                    .FontBold = False
                    Printer.Print "Per.IIBB"
                
                'Importe IIBB
                    .CurrentX = X + 143
                    .CurrentY = Y + 187
                    .Font = "Courier New"
                    .FontSize = 8
                    .FontBold = False
                    Printer.Print NotaCC!ImportePercepIIBB
            End If
            
            'Importe Total
                .CurrentX = X + 143
                .CurrentY = Y + 194
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print NotaCC!TotalNotaCredito
        .EndDoc
        
    End With
    
    NotaCC.Close
    NotaCD.Close
        
CapturaErrores:
    'If Err = 321 Then
    '    Resume Next
    'End If

End Sub

Public Function Descripcion(IdCodProd As Variant) As String

    Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    
    tProductos.Index = "PrimaryKey"
    
    tProductos.Seek "=", IdCodProd

    If Not tProductos.NoMatch Then Descripcion = tProductos!Descripcion

End Function

Private Sub BotonImprimir_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonNueva_Click()

    Dim NumeroNotaCredito As Long
    Dim num As Long
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
    
    
    
    Dim busco As String
       
   If TextTipoNotaCredito.Text = "A" Then
        busco = "tNotaCreditoA"
    End If
    
    If TextTipoNotaCredito.Text = "B" Then
        busco = "tNotaCreditoB"
    End If
    
    
    'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
    rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
    NumeroNotaCredito = rstUltimosNumeros.Fields!UltimoNumero
    
    'If rstUltimosNumeros.NoMatch Then
    '   FG1.Visible = False
    '   mensaje = MsgBox("No existen Numeros de Factura", vbCritical, "Final de la busqueda")
    'End If
    
    TextNumeroNotaCredito.Text = NumeroNotaCredito + 1
    
    num = Val(TextNumeroNotaCredito.Text)
    
    If TextCuit.Text <> "" Then
       FG1.Enabled = True
    End If
    
    BotonNueva.Enabled = False
    
    'ComboVendedor.SetFocus
   
    FG1.Row = 1
    FG1.Col = 0
    FG1.Enabled = True
    'FG1.SetFocus
        
    TextNumeroNotaCredito.SetFocus
  
End Sub

Private Sub BotonNueva_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonPago_Click()

    FormPagoFacturas.Show
    
End Sub

Private Sub BotonPago_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonSalir_Click()

    Unload FormNotaCredito

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

Private Sub ComboArticulo_Click()

    Dim precioUnitario As Double

    TextPrecioUnitario.Text = ""
    TextPorDescuento.Text = ""
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
    
    If TextTipoNotaCredito.Text = "B" Then
        precioUnitario = rstProductos.Fields!PrecioUnitarioFactura * 1.21
        TextPrecioUnitario.Text = Format(precioUnitario, "#0.00")
    Else
        TextPrecioUnitario.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#0.00")
    End If
    
    'TextPrecioUnitario.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#0.00")
      
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
    TextPorDescuento.Text = ""
    TextCantidad.Text = ""
    TextTotalLineaProd.Text = ""
    TextPorDesc.Text = ""
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    'If KeyAscii = 13 Then
    If ComboArticulo.Text <> "" Then
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
                 
                 If TextTipoNotaCredito.Text = "B" Then
                    precioUnitario = rstProductos.Fields!PrecioUnitarioFactura * 1.21
                    TextPrecioUnitario.Text = Format(precioUnitario, "#0.00")
                  Else
                    TextPrecioUnitario.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#0.00")
                End If
                 
                ' TextPrecioUnitario.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#0.00")
                
           Else
                mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                ComboArticulo.Text = ""
                TextDescripcion.Text = ""
                TextUnidadMedida.Text = ""
                TextPrecioUnitario.Text = ""
           End If
              
           tProductos.Close
          'TextCantidad.SetFocus
    'End If
    End If
    
End Sub


Private Sub ComboVendedor_GotFocus()
    ComboVendedor.SelLength = Len(ComboVendedor.Text)
End Sub

Private Sub TextApellidoNombre_Change()
    Columna = 1
    Call FiltrarGrilla(MSFlexGrid1, TextApellidoNombre, CLng(Columna))
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
                    If Not IsNull(tClientes.Fields!Domicilio) Then MSFlexGrid1.Text = tClientes.Fields!Domicilio
                    MSFlexGrid1.Col = 4
                    If Not IsNull(tClientes.Fields!localidad) Then MSFlexGrid1.Text = tClientes.Fields!localidad
                    MSFlexGrid1.Col = 5
                    If tClientes.Fields!CP <> "" Then
                        MSFlexGrid1.Text = tClientes.Fields!CP
                    End If
                    MSFlexGrid1.Col = 6
                    If Not IsNull(tClientes.Fields!Prov) Then MSFlexGrid1.Text = tClientes.Fields!Prov
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

Private Sub TextApellidoNombre_GotFocus()
'    TextApellidoNombre.SelLength = Len(TextApellidoNombre.Text)
End Sub

Private Sub TextCantidad_GotFocus()
    TextCantidad.SelLength = Len(TextCantidad.Text)
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

End Sub




Private Sub TextNumeroNotaCredito_GotFocus()
    TextNumeroNotaCredito.SelLength = Len(TextNumeroNotaCredito.Text)
End Sub


Private Sub TextNumeroNotaCredito_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextPorDescuento_GotFocus()
    TextPorDescuento.SelLength = Len(TextPorDescuento.Text)
End Sub

Private Sub TextPrecioUnitario_GotFocus()
    TextPrecioUnitario.SelLength = Len(TextPrecioUnitario.Text)
End Sub

Private Sub TextUnidadMedida_GotFocus()
    TextUnidadMedida.SelLength = Len(TextUnidadMedida.Text)
End Sub

Private Sub TextUnidadMedida_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub
Private Sub TextPrecioUnitario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If
    
    If KeyAscii = 46 Then KeyAscii = 44
   
End Sub

Private Sub TextPrecioUnitario_LostFocus()
     
    'If KeyAscii = 13 Then
        Call calculoprecios
    'End If

End Sub
Private Sub TextPorDescuento_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub TextPorDescuento_LostFocus()
     
    'If KeyAscii = 13 Then
        Call calculoprecios
    'End If

End Sub
Private Sub TextCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'Call calculoprecios
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
    
End Sub

Private Sub TextTotalLineaProd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub
Private Sub calculoprecios()

    Dim precioUnitario As Double
    Dim porcentaje As Double
    Dim cantidad As Integer
    Dim total As Double
    Dim porcentajePrecioUnitario As Double
    Dim totalLinea As Double
    
    'If KeyAscii = 13 Then
    If TextCantidad.Text <> "" Then
         
        precioUnitario = Format(TextPrecioUnitario.Text, "#00.00")
        If TextPorDescuento.Text <> "" Then
            porcentaje = Val(TextPorDescuento.Text)
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

Private Sub TextCantidad_LostFocus()
    
    If TextCantidad.Text = "" Then
        A = MsgBox("NO PUEDE DEJAR LA CANTIDAD EN BLANCO", vbCritical, "E R R O R ! ! !")
        TextCantidad.SetFocus
    End If
    
    Call calculoprecios
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

Private Sub CommandSalir_Click()

    Unload FormNotaCreditoCliente

End Sub

'Private Sub FG1_KeyPress(KeyAscii As Integer)
'
'    Dim precioUnitario As Double
'    Dim cantidad As Integer
'    Dim porcentaje As Double
'    Dim total
'    Dim totalLinea As Double
'    Dim totalGrilla
'    Dim subTotalNotaCreditoForm
'    Dim porcentajePrecioUnitario As Double
'    Dim descuentoFactura As Double
'    Dim TotalNotaCreditoForm As Double
'    Dim iva As Double
'    Dim impuesto As Double
'    Dim percepcion As Double
'    Dim columnaSeis As Integer
'    Dim columnaSiete As Integer
'    Dim bandera As Integer
'
'
'
'    ruta = App.Path & "\DB_SPC_SI.mdb"
'
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
'
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstIva = db.OpenRecordset("Iva", dbOpenDynaset)
'
'    iva = rstIva.Fields!iva
'
'    If TextTipoNotaCredito.Text = "A" Then
'        Textiva.Text = Format(iva, "#00.00")
'    End If
'
'    If KeyAscii >= 32 And KeyAscii <= 127 Then
'        FG1.Text = FG1.Text & Chr(KeyAscii)
'    End If
'
'    Select Case KeyAscii
'       Case 13
'
'
'            FG1.Col = 0
'            codigoprodMA = UCase(FG1.Text)
'
'              '******* Busco Producto
'
'           Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
'
'
'           tProductos.Index = "PrimaryKey"
'
'           tProductos.MoveFirst
'           tProductos.Seek "=", codigoprodMA
'
'           If Not tProductos.NoMatch Then
'                produ = tProductos!CodProd
'                'MsgBox (DepoOrigen)
'                 Dim busca1 As String, busca2 As String
'                 busca1 = RTrim(LTrim(codigoprodMA))
'                 busca2 = busca1 + "z"
'
'                 rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
'
'                 codigoProdTabla = rstProductos.Fields!CodProd
'
'                cantidadProducto = rstProductos.Fields!Stock
'                FG1.Col = 1
'                FG1.Text = rstProductos.Fields!Descripcion
'                FG1.Col = 2
'                FG1.Text = rstProductos.Fields!UnidadMedida
'                FG1.Col = 3
'                FG1.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#00.00")
'                FG1.Col = FG1.Col + 2
'
'           Else
'                 mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
'                 codigoprodMA = ""
'                 Textiva.Text = "0,00"
'                 TextPercepcionIIBB.Text = "0,00"
'                 TextTotalNotaCredito.Text = "0,00"
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
'
'           tProductos.Close
'
'            If bandera <> 1 Then
'                Call muestrodatosproductos
'                FG1.Col = FG1.Col + 2
'            End If
'
'            '***********************
'
'           '*** descuento
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
'
'            End If
'
'            '**** suma total linea
'
'            columnaSeis = 6
'
'            total = SumarTotalGrilla(FG1, columnaSeis)
'            subTotalNotaCreditoForm = total
'
'            TextSubTotalNotaCredito.Text = Format(total, "#00.00")
'
'            '**** suma descuentos
'
'            columnaSiete = 7
'
'            porcentajePrecioUnitario = SumarTotalDescuentos(FG1, columnaSiete)
'            descuentoFactura = porcentajePrecioUnitario
'
'            TextDescuentos.Text = Format(descuentoFactura, "#0.00")
'
'            '**** calculo alicuota
'
'            TextAlicuota.Text = Format(Alicuota, "#0.00")
'
'            If TextTipoNotaCredito.Text = "A" Then
'                percepcion = (subTotalNotaCreditoForm - descuentoFactura) * Alicuota / 100
'                TextPercepcionIIBB.Text = Format(percepcion, "#0.00")
'
'            End If
'
'            '**** calculo impuesto
'
'            If TextTipoNotaCredito.Text = "A" Then
'               impuesto = (subTotalNotaCreditoForm - descuentoFactura) * iva / 100
'               TextImpuesto.Text = Format(impuesto, "#0.00")
'            End If
'
'            '**** calculo total factura
'
'            TotalNotaCreditoForm = (subTotalNotaCreditoForm - descuentoFactura + percepcion + impuesto)
'
'            TextTotalNotaCredito.Text = Format(TotalNotaCreditoForm, "#00.00")
'
'            If CDec(TotalNotaCreditoForm) <> 0 Then
'                 BotonGrabar.Enabled = True
'                 'BotonImprimir.Enabled = True
'                 'BotonPago.Enabled = True
'            End If
'
'
'            If FG1.Col = 7 And FG1.Text <> "" Then
'                FG1.Col = 0
'                'If FG1.Row < 2 Then
'                    FG1.Row = FG1.Row + 1
'                    FG1.SetFocus
'                    BotonGrabar.Enabled = True
'                    'BotonImprimir.Enabled = True
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

Private Sub muestrodatosproductos()

    cantidadProducto = rstProductos.Fields!Stock
    FG1.Col = 1
    FG1.Text = rstProductos.Fields!Descripcion
    FG1.Col = 2
    FG1.Text = rstProductos.Fields!UnidadMedida
    FG1.Col = 3
    FG1.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#00.00")
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

    FormNotaCredito.Height = 10200
    FormNotaCredito.Width = 12135
    FormNotaCredito.Top = 600
    FormNotaCredito.Left = 50
        
    Fila = 1
    renglon = 16
    
    Call SeteoGrilla
      
    Call Cargo
    Call buscoarticulo
    
    TextFechaNotaCredito.Text = Format(Date, "dd/mm/yyyy")
    
    opContado.Value = True
    opCtaCte.Value = False
    
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
    
    Dim item As Integer
    Dim linea As Integer
    
    FG1.Row = 0
    FG1.Col = 0
    
    FG1.ColWidth(0) = 1000
    FG1.CellFontBold = True
    FG1.ColAlignment(0) = flexAlignCenterCenter
    FG1.Text = "Articulo"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 4700
    FG1.CellFontBold = True
    FG1.Text = "Descripción"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 700
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
    FG1.ColWidth(7) = 0
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
        ComboVendedor.AddItem rstEmpleado!Nombre
        rstEmpleado.MoveNext
    Loop

End Sub

Private Sub busco()

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
            MSFlexGrid1.Text = rstCliente.Fields!localidad
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
        Call busco
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

    TextAlicuota.Text = ""
   
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
                TextCuit.Text = rstCliente.Fields!CUIT
                If Not IsNull(rstCliente.Fields!Domicilio) Then TextDireccion.Text = rstCliente.Fields!Domicilio
                If Not IsNull(rstCliente.Fields!localidad) Then TextLocalidad.Text = rstCliente.Fields!localidad
                If Not IsNull(rstCliente.Fields!CP) Then TextCodigoPostal.Text = rstCliente.Fields!CP
                If Not IsNull(rstCliente.Fields!Prov) Then TextProvincia.Text = rstCliente.Fields!Prov
                TextDescuentoCliente.Text = rstCliente.Fields!PorcentajeDescuento
                vendedorCliente = rstCliente.Fields!Vendedor
                Call buscocuilyvendedor
            End If
        End If
        TextNumeroNotaCredito.Text = ""
    End If
    
    If TextNumeroNotaCredito <> "" Then
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
    
    TextCuit.Text = rstCliente.Fields!CUIT
    codigovendedor = rstCliente!Vendedor
      
    Set rstPadron = db.OpenRecordset("Padron", dbOpenTable)
    
    rstPadron.Index = "CUIT"
    
    With rstPadron
        rstPadron.Seek "=", TextCuit.Text
        If .NoMatch = False Then
            Alicuota = !AlicuotaPercepcion
        End If
    End With
    
    
    TextDescuentoCliente.Text = rstCliente.Fields!PorcentajeDescuento
    
    '*** Busco Vendedor
    
    CodigoVend = codigovendedor
      
    rstEmpleado.FindFirst "Legajo >= '" & CodigoVend & "'"
    
    LegajoEmpleado = rstEmpleado.Fields!Legajo
    ComboVendedor.Text = rstEmpleado.Fields!Nombre
    
    '*** Busco Saldo
    
   rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
    
   TextSaldoCliente.Text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
   
    '*** Busco Condicion IVA
    
    condicionIva = rstCliente.Fields!condicionIva
    If condicionIva = "RI" Then
        TextTipoNotaCredito.Text = "A"
    End If
    If condicionIva = "CF" Then
        TextTipoNotaCredito.Text = "B"
    End If
    
    'Mas Condiciones para factura B agregado por PVS 25/04/2016
    
    If condicionIva = "EX" Then
        TextTipoNotaCredito.Text = "B"
    End If
    
    If condicionIva = "NR" Then
        TextTipoNotaCredito.Text = "B"
    End If
    
    If condicionIva = "MO" Then
        TextTipoNotaCredito.Text = "B"
    End If

    If condicionIva = "RN" Then
        TextTipoNotaCredito.Text = "B"
    End If
'************************************************************

    If TextTipoNotaCredito.Text = "A" Then
        TextAlicuota.Text = Format(Alicuota, "#0.00")
    End If
    
    BotonNueva.Enabled = True
    BotonNueva.SetFocus
    
    'Agregado por PVS 22/04/2016
        MSFlexGrid1.Visible = False
    
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



