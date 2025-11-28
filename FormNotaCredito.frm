VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormNotaCredito 
   BackColor       =   &H00008080&
   Caption         =   "Nota de Credito"
   ClientHeight    =   9630
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   13560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9630
   ScaleWidth      =   13560
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1335
      Left            =   2880
      TabIndex        =   51
      Top             =   3480
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
      TabIndex        =   65
      Top             =   2520
      Width           =   13095
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
         Left            =   8400
         TabIndex        =   7
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
         Left            =   7080
         TabIndex        =   6
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
         Left            =   9360
         TabIndex        =   8
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
         Left            =   6480
         TabIndex        =   5
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
         Left            =   9960
         TabIndex        =   9
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
         TabIndex        =   4
         Top             =   600
         Width           =   4935
      End
      Begin VB.CommandButton BotonAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   11760
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox ComboArticulo 
         Height          =   315
         Left            =   120
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
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
         TabIndex        =   66
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
         Left            =   10080
         TabIndex        =   73
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
         Left            =   8400
         TabIndex        =   72
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
         Left            =   9240
         TabIndex        =   71
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
         Left            =   7200
         TabIndex        =   70
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
         Left            =   6600
         TabIndex        =   69
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
         TabIndex        =   68
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
         TabIndex        =   67
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00008080&
      Height          =   975
      Left            =   120
      TabIndex        =   52
      Top             =   1560
      Width           =   13095
      Begin VB.ComboBox cmbFacturaReferencia 
         Height          =   315
         Left            =   6240
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox txtTipoFacturaReferencia 
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
         Left            =   8040
         TabIndex        =   77
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton opContado 
         BackColor       =   &H00008080&
         Caption         =   "Contado"
         Height          =   255
         Left            =   11040
         TabIndex        =   76
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton opCtaCte 
         BackColor       =   &H00008080&
         Caption         =   "Cta Cte"
         Height          =   255
         Left            =   12000
         TabIndex        =   75
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
         TabIndex        =   58
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
         TabIndex        =   57
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox ComboVendedor 
         Height          =   315
         Left            =   4320
         TabIndex        =   56
         Top             =   480
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
         Left            =   8880
         TabIndex        =   55
         Top             =   480
         Width           =   975
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
         Left            =   9480
         TabIndex        =   54
         Top             =   480
         Visible         =   0   'False
         Width           =   1215
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
         Left            =   11040
         TabIndex        =   53
         Top             =   600
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Nº Fact Ref."
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
         TabIndex        =   79
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackColor       =   &H00008080&
         Caption         =   "Tipo FO"
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
         Left            =   7920
         TabIndex        =   78
         Top             =   240
         Width           =   690
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
         TabIndex        =   64
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
         TabIndex        =   63
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
         TabIndex        =   62
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
         Left            =   4680
         TabIndex        =   61
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
         Left            =   8880
         TabIndex        =   60
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
         Left            =   9720
         TabIndex        =   59
         Top             =   240
         Visible         =   0   'False
         Width           =   495
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00008080&
      Height          =   1455
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   13095
      Begin VB.TextBox TextProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   9000
         TabIndex        =   28
         Top             =   960
         Width           =   3615
      End
      Begin VB.TextBox TextDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   39
         Top             =   600
         Width           =   5535
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   1560
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TextApellidoNombre 
         Height          =   285
         Left            =   7080
         TabIndex        =   12
         Top             =   240
         Width           =   5535
      End
      Begin VB.TextBox TextCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   31
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1560
         TabIndex        =   30
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   6480
         TabIndex        =   29
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
         TabIndex        =   38
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
         TabIndex        =   37
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
         TabIndex        =   36
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
         TabIndex        =   35
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
         Left            =   5160
         TabIndex        =   34
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
         TabIndex        =   33
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
         Left            =   8040
         TabIndex        =   32
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00008080&
      Height          =   1095
      Left            =   120
      TabIndex        =   41
      Top             =   8520
      Width           =   13095
      Begin VB.CommandButton BotonGrabar 
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   750
         Left            =   2160
         TabIndex        =   49
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonPago 
         Caption         =   "&Pago"
         Enabled         =   0   'False
         Height          =   750
         Left            =   8760
         TabIndex        =   47
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   11640
         TabIndex        =   45
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonCancelar 
         Caption         =   "&Cancelar"
         Height          =   750
         Left            =   10200
         TabIndex        =   44
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton BotonImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   750
         Left            =   3600
         TabIndex        =   43
         Top             =   240
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton BotonNueva 
         Caption         =   "&Nueva"
         Enabled         =   0   'False
         Height          =   750
         Left            =   720
         TabIndex        =   42
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00008080&
      Height          =   3615
      Left            =   120
      TabIndex        =   40
      Top             =   3600
      Width           =   13095
      Begin VB.CommandButton BotonEliminarfila 
         Caption         =   "&Eliminar Fila"
         Height          =   495
         Left            =   12120
         TabIndex        =   50
         Top             =   1920
         Width           =   735
      End
      Begin VB.CommandButton BotonBuscarProducto 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   12120
         TabIndex        =   48
         Top             =   2880
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   11775
         _ExtentX        =   20770
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
      TabIndex        =   13
      Top             =   7320
      Width           =   13095
      Begin VB.CommandButton BotonCalcular 
         Caption         =   "&Calcular"
         Height          =   375
         Left            =   11160
         TabIndex        =   74
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
         TabIndex        =   20
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
         Left            =   1680
         TabIndex        =   19
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
         Left            =   2520
         TabIndex        =   18
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
         Left            =   4680
         TabIndex        =   17
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
         Left            =   6720
         TabIndex        =   16
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
         Left            =   8880
         TabIndex        =   15
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
         Left            =   11040
         TabIndex        =   14
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
         Left            =   4920
         TabIndex        =   46
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
         TabIndex        =   26
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
         Left            =   1800
         TabIndex        =   25
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
         Left            =   2520
         TabIndex        =   24
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
         Left            =   6960
         TabIndex        =   23
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
         Left            =   9360
         TabIndex        =   22
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
         Left            =   10920
         TabIndex        =   21
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

Private Sub LlenoComboFacturasDeReferencia(txtTipoFacturaReferencia As String)

    vSQL = "SELECT * FROM FacturaC WHERE CodCliente=" & Val(TextCodigoCliente.text) & " ORDER BY FechaFactura DESC"
    
    Set tP = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    tP.MoveFirst
    
    While Not tP.EOF
        cmbFacturaReferencia.AddItem tP!NroFactura
        tP.MoveNext
    Wend
    
    tP.Close

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
        FG1.text = ComboArticulo.text
        FG1.Col = 1
        FG1.text = TextDescripcion.text
        FG1.Col = 2
        FG1.text = TextUnidadMedida.text
        FG1.Col = 3
        FG1.text = Format(TextPrecioUnitario.text, "#0.00")
        FG1.Col = 4
        FG1.text = Format(TextPorDescuento.text, "#0.00")
        FG1.Col = 5
        FG1.text = TextCantidad.text
        FG1.Col = 6
        FG1.text = Format(TextTotalLineaProd.text, "#0.00")
        FG1.Col = 7
        FG1.text = Format(TextPorDesc.text, "#0.00")
                
        Fila = Fila + 1
        fila2 = fila2 + 1
            
    
        Call calculototalnotacredito
        
'        ComboArticulo.Text = ""
        TextDescripcion.text = ""
        TextUnidadMedida.text = ""
        TextPrecioUnitario.text = ""
        TextPorDescuento.text = ""
        TextCantidad.text = ""
        TextTotalLineaProd.text = ""
        TextPorDesc.text = ""
        
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
    
    If TextTipoNotaCredito.text = "A" Then
       Textiva.text = Format(iva, "#00.00")
    End If

    '**** suma total linea
             
    'total = Val(TextTotalLineaProd.Text) + Val(TextSubTotalNotaCredito.Text)
    If (TextSubTotalNotaCredito.text = "") Then TextSubTotalNotaCredito.text = 0
    total = CDbl(TextTotalLineaProd.text) + CDbl(TextSubTotalNotaCredito.text)
    subtotalFacturaForm = total
                            
    TextSubTotalNotaCredito.text = Format(total, "#0.00")
    
    '**** suma descuentos
     
    'porcentajePrecioUnitario = Val(porcentajePrecioUnitario) + Val(TextPorDesc.Text)
    If (TextDescuentos.text = "") Then TextDescuentos.text = 0
    porcentajePrecioUnitario = CDbl(TextDescuentos.text) + CDbl(TextPorDesc.text)
    descuentoFactura = porcentajePrecioUnitario
                            
    TextDescuentos.text = Format(descuentoFactura, "#0.00")
    
    '**** calculo alicuota

    TextAlicuota.text = Format(Alicuota, "#0.00")
                    
   If TextTipoNotaCredito.text = "A" Then
        percepcion = subtotalFacturaForm * Alicuota / 100
        TextPercepcionIIBB.text = Format(percepcion, "#0.00")
        
    End If
    
    '**** calculo impuesto
    
    If TextTipoNotaCredito.text = "A" Then
       impuesto = subtotalFacturaForm * iva / 100
       TextImpuesto.text = Format(impuesto, "#0.00")
    End If
    
    '**** calculo total factura
    
    'totalFacturaForm = (subtotalFacturaForm - descuentoFactura + percepcion + impuesto)
    
    totalFacturaForm = (subtotalFacturaForm + percepcion + impuesto)
    
    TextTotalNotaCredito.text = Format(totalFacturaForm, "#0.00")
    
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
                      
    total = Format(TextSubTotalNotaCredito.text, "#0.00")
    
    TextAlicuota.text = Format(Alicuota, "#0.00")
                    
    percepcion = TextPercepcionIIBB.text
       
    impuesto = TextImpuesto.text
    
    totalFacturaForm = (total + subtotalFacturaForm + percepcion + impuesto)
    
    TextTotalNotaCredito.text = Format(totalFacturaForm, "#0.00")
    
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

    TextCodigoCliente.text = ""
    TextApellidoNombre.text = ""
    TextCuit.text = ""
    TextDireccion.text = ""
    TextLocalidad.text = ""
    TextCodigoPostal.text = ""
    TextProvincia.text = ""
    TextNumeroNotaCredito.text = ""
    TextTipoNotaCredito.text = ""
    ComboVendedor.text = ""
    TextDescuentoCliente.text = ""
    TextSubTotalNotaCredito.text = ""
    TextDescuentos.text = ""
    TextPercepcionIIBB.text = ""
    TextAlicuota.text = ""
    TextImpuesto.text = ""
    Textiva.text = ""
    TextTotalNotaCredito.text = ""
    TextSaldoCliente.text = ""
    ComboVendedor.text = ""
    TextDescuentoCliente.text = ""
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
                                    
            TextSubtotalPresupuesto.text = Format(total, "#00.00")
            
            '**** suma descuentos
            
            columnaSiete = 7
             
            porcentajePrecioUnitario = SumarTotalDescuentos(FG1, columnaSiete)
            descuentoPresup = porcentajePrecioUnitario
                                    
            TextDescuentos.text = Format(descuentoPresup, "#0.00")
            
           
            '**** calculo total factura
            
            totalPresuForm = (subtotalPresuForm - descuentoPresup)
            
            TextTotalPresupuesto.text = Format(totalPresuForm, "#00.00")
            
    
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
        
        rstFacC.Seek "=", TextTipoNotaCredito, Str(TextNumeroNotaCredito.text)

        If Not rstFacC.NoMatch Then
            A = MsgBox("Factura Existente", vbCritical, "INFO DEL SISTEMA")
           
            TextNumeroNotaCredito.text = num
            TextNumeroNotaCredito.SetFocus
        Else
        
        rstFacC.Close
        db1.Close
        
        
     
            rstNotaCreditoC.AddNew
            rstNotaCreditoC.Fields!NroNotaCredito = TextNumeroNotaCredito.text
            rstNotaCreditoC.Fields!TipoNotaCredito = UCase(TextTipoNotaCredito.text)
            rstNotaCreditoC.Fields!FechaNotaCredito = TextFechaNotaCredito.text
            rstNotaCreditoC.Fields!TotalNotaCredito = TextTotalNotaCredito.text
            If Textiva.text <> "" Then
                rstNotaCreditoC.Fields!PorcentajeIVA = Textiva.text
            Else
                rstNotaCreditoC.Fields!PorcentajeIVA = "0,00"
            End If
            
            If TextSubTotalNotaCredito.text = TextTotalNotaCredito.text Then TextSubTotalNotaCredito.text = (TextTotalNotaCredito.text / 1.21)
            rstNotaCreditoC.Fields!SubTotalNotaCredito = Format(TextSubTotalNotaCredito.text, "#0.00")
            
            If TextImpuesto.text <> "" Then
                rstNotaCreditoC.Fields!totalIva = TextImpuesto.text
            Else
                'rstNotaCreditoC.Fields!totalIva = "0,00"
                rstNotaCreditoC.Fields!totalIva = Format((TextSubTotalNotaCredito.text * 21) / 100, "#0.00")
            End If
            If TextAlicuota.text = "" Then TextAlicuota.text = 0
            rstNotaCreditoC.Fields!AlicuotaIIBB = TextAlicuota.text
            If TextPercepcionIIBB.text <> "" Then
                rstNotaCreditoC.Fields!ImportePercepIIBB = TextPercepcionIIBB.text
            End If
            rstNotaCreditoC.Fields!CodCliente = TextCodigoCliente.text
            rstNotaCreditoC.Fields!PorcentajeDesc = TextDescuentoCliente.text
            rstNotaCreditoC.Fields!ImporteDesc = TextDescuentos.text
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
                  If FG1.text <> "" Then
                        rstNotaCreditoD.AddNew
                    
                        rstNotaCreditoD.Fields!NroNotaCredito = TextNumeroNotaCredito.text
                        rstNotaCreditoD.Fields!TipoNotaCredito = TextTipoNotaCredito.text
                    
                        FG1.Col = 0
                        rstNotaCreditoD.Fields!IdCodProd = FG1.text
                    
                        FG1.Col = 2
                        rstNotaCreditoD.Fields!UnidadMedida = FG1.text
                        
                        FG1.Col = 3
                        rstNotaCreditoD.Fields!precioUnitario = Format(FG1.text, "#00.00")
                        
                        FG1.Col = 4
                        des = FG1.text
                        If des <> "" Then
                           rstNotaCreditoD.Fields!PorcentajeDescuento = Val(des)
                        Else
                           rstNotaCreditoD.Fields!PorcentajeDescuento = Val(TextDescuentoCliente.text)
                        End If
                        FG1.Col = 5
                        rstNotaCreditoD.Fields!cantidad = Val(FG1.text)
                        descuentoCantidad = Val(FG1.text)
                        
                        '*** Modifico Stock Producto
                       

                       'Call DesHagoStock(CodProd, descuentoCantidad)
                        
                        If modificaStock = 1 Then
                            FG1.Col = 0
                            codigoprod = FG1.text
                            
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
                        rstNotaCreditoD.Fields!totalLinea = Format(FG1.text, "#00.00")
                        
                        FG1.Col = 7
                        rstNotaCreditoD.Fields!ImporteDescuento = Format(FG1.text, "#00.00")
                        
                        FG1.Col = 8
                        rstNotaCreditoD.Fields!ItemNotaCredito = Val(FG1.text)
                         
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
                If TextTipoNotaCredito.text = "A" Then
                    TipoComp = 3
                 Else
                    If TextTipoNotaCredito.text = "B" Then
                        TipoComp = 8
                    End If
                End If
                
                If TextCuit.text <> "" Then
                    DocTipo = 80
                    DocNro = CDbl(TextCuit.text)
                 Else
                    DocTipo = 96
                    DocNro = 11111111
                End If
                
                CbteDesde = CDbl(TextNumeroNotaCredito.text)
                CbteHasta = CDbl(TextNumeroNotaCredito.text)
                CbteFch = CStr(Format(TextFechaNotaCredito.text, "YYYYMMDD"))
                ImpTotal = Format(CDbl(TextTotalNotaCredito.text), "Standard")
                
               ' If TipoComp = 3 Then ImpNeto = Format(CDbl(TextSubTotalNotaCredito.Text), "Standard")
               ' If TipoComp = 8 Then ImpNeto = Format(CDbl(TextTotalNotaCredito.Text), "Standard")
                
                ImpNeto = Format(CDbl(TextSubTotalNotaCredito.text), "Standard")
                BaseImpIVA = ImpNeto
                BaseImpTributo = ImpNeto

                
                MonId = "PES"
                MonCotiz = 1
                
                If Textiva.text = "" Then AlicIVA = 5
                
                If Textiva.text = "21,00" Then
                    AlicIVA = 5
                Else
                    If Textiva.text = "10,5" Then
                        AlicIVA = 4
                    End If
                End If
                
                'If TipoComp = 3 Then BaseImpIVA = Format(CDbl(TextSubTotalNotaCredito.Text), "Standard")
                'If TipoComp = 8 Then BaseImpIVA = ImpNeto
                
                If TextImpuesto.text <> "" Then
                    ImpIva = Format(CDbl(TextImpuesto.text), "Standard")
                 Else
                    'ImpIva = 0
                    ImpIva = Format(ImpTotal - ImpNeto, "Standard")
                End If
                
                IdTributo = 2
                DescTributo = "IIBB"
                
                If TipoComp = 3 Then BaseImpTributo = Format(CDbl(TextSubTotalNotaCredito.text), "Standard")
                
                Alicuota = Format(CDbl(TextAlicuota.text), "Standard")
                
                If TextPercepcionIIBB.text <> "" Then
                    ImpAlicuota = Format(CDbl(TextPercepcionIIBB.text), "Standard")
                 Else
                    ImpAlicuota = 0
                End If
                
            'Buscamos el comprobante asociado a la NC
                If cmbFacturaReferencia.text = "" Then
                    z = MsgBox("Debe Elegir un Comprobante Asociado a la Nota de Crédito", vbOKOnly, "ERROR !!!")
                    cmbFacturaReferencia.SetFocus
                End If
                
                Call BuscaCbteAsociado(CLng(cmbFacturaReferencia.text), CStr(txtTipoFacturaReferencia.text))
            
'LOCAL
              'Prueba de generar la cabecera luego de poder realizar la factura electrónica
              '  If Not FacturaElectronicaSPC(PtoVta, DocTipo, DocNro, TipoComp, CbteDesde, CbteHasta, CbteFch, ImpTotal, ImpNeto, MonId, MonCotiz, AlicIVA, BaseImpIVA, ImpIva, IdTributo, DescTributo, BaseImpTributo, Alicuota, ImpAlicuota, ImporteExento, TipoCbteAsoc, NroCbteAsoc, FechaCbteAsoc) Then
                    
              '      Call RevertirNC(TipoComp, CbteDesde)
              '      Fila = Fila - 1
              '      fila2 = fila2 - 1
                    
              '      Call blanqueototal
              '      MSFlexGrid1.Visible = False
              '      TextCodigoCliente.SetFocus
              '      Exit Sub
              '  End If

'///////////////////////FIN FE SPC //////////////////////////////////////////////////////////////////////////////////////
                
                'Call FacturaElectronicaSPC(PtoVta, DocTipo, DocNro, TipoComp, CbteDesde, CbteHasta, CbteFch, ImpTotal, ImpNeto, MonId, MonCotiz, AlicIVA, BaseImpIVA, ImpIva, IdTributo, DescTributo, BaseImpTributo, Alicuota, ImpAlicuota, ImporteExento, TipoCbteAsoc, NroCbteAsoc, FechaCbteAsoc)
     '//////////////////////////////////////////////////////////////////////////////////
            
            
            '*** Grabo Linea 1 en Cuenta Corriente
            
            CodigoClie = Val(TextCodigoCliente.text)
      
            rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCliente.Fields!IdCliente <> Val(TextCodigoCliente.text) Then
                mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                'TextCodigoCliente.Text = ""
                'Call blanqueototal
                'TextCodigoCliente.SetFocus
            Else
                rstCtaCte.Edit
                saldo1 = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
                saldo2 = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
                saldoLi1 = Format(-TextTotalNotaCredito.text, "#0.00")
                rstCtaCte.Fields!SaldoL1 = saldoLi1 + saldo1
                saldo1 = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
                rstCtaCte.Fields!SaldoTotal = saldo1 + saldo2
                rstCtaCte.Update
            End If
        
            
            '*** Grabo Movimientos Cuente corriente
        
            rstMovimientosCtaCte.AddNew
            'rstMovimientosCtaCte.Fields!Fecha = Format(Date, "dd/mm/yyyy")
            rstMovimientosCtaCte.Fields!Fecha = Format(TextFechaNotaCredito.text, "dd/mm/yyyy")
            rstMovimientosCtaCte.Fields!IdCliente = TextCodigoCliente.text
            If TextTipoNotaCredito.text = "A" Then
                rstMovimientosCtaCte.Fields!tipoDoc = "Nota Credito A"
            End If
            If TextTipoNotaCredito.text = "B" Then
                rstMovimientosCtaCte.Fields!tipoDoc = "Nota Credito B"
            End If
            rstMovimientosCtaCte.Fields!NroDoc = TextNumeroNotaCredito.text
            rstMovimientosCtaCte.Fields!ImporteLinea1 = -TextTotalNotaCredito.text
            rstMovimientosCtaCte.Fields!ImporteLinea2 = 0
            rstMovimientosCtaCte.Update
            
            '*** Actualizo Ultimo Numero Cuenta Corriente
            
            Set db = DBEngine.OpenDatabase(ruta)
            Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
        
            Dim busco As String
       
            If TextTipoNotaCredito.text = "A" Then
                busco = "tNotaCreditoA"
            End If
            
            If TextTipoNotaCredito.text = "B" Then
                busco = "tNotaCreditoB"
            End If
    
            '*****************
            'Guardo el nro de NC en la variable global para luego poder imprimir
                vNroNCImp = TextNumeroNotaCredito.text
                vTipoNCImp = TextTipoNotaCredito.text
            '*****************
    
    
            'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
            rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
            ultimo = rstUltimosNumeros.Fields!UltimoNumero
            
            If ultimo < Val(TextNumeroNotaCredito.text) Then
                rstUltimosNumeros.Edit
                'If ultimo < rstUltimosNumeros.Fields!UltimoNumero Then
                     rstUltimosNumeros.Fields!UltimoNumero = TextNumeroNotaCredito.text
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
            Printer.Print Format(TextFechaNotaCredito.text, "DD/MM/YYYY")
        
        'Imprimir Nombres
            .CurrentX = X + 37
            .CurrentY = Y + 54
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = True
            Printer.Print TextApellidoNombre.text
            
        'Imprimir Direccion
            .CurrentX = X + 37
            .CurrentY = Y + 60
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextDireccion.text
            
        'Imprimir Localidad
            .CurrentX = X + 37
            .CurrentY = Y + 65
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextLocalidad.text
            
        'Imprimir CUIT
            .CurrentX = X + 125
            .CurrentY = Y + 67
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextCuit.text
            
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
            
            sqlfc = "SELECT * FROM NotaCreditoC WHERE TipoNotaCredito='" & TextTipoNotaCredito.text & "' AND NroNotaCredito=" & TextNumeroNotaCredito.text & " ORDER By NroNotaCredito"
            vsqlFD = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & TextTipoNotaCredito.text & "' AND NroNotaCredito=" & TextNumeroNotaCredito.text & " ORDER By NroNotaCredito"
            
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
       
   If TextTipoNotaCredito.text = "A" Then
        busco = "tNotaCreditoA"
    End If
    
    If TextTipoNotaCredito.text = "B" Then
        busco = "tNotaCreditoB"
    End If
    
    
    'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
    rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
    NumeroNotaCredito = rstUltimosNumeros.Fields!UltimoNumero
    
    'If rstUltimosNumeros.NoMatch Then
    '   FG1.Visible = False
    '   mensaje = MsgBox("No existen Numeros de Factura", vbCritical, "Final de la busqueda")
    'End If
    
    TextNumeroNotaCredito.text = NumeroNotaCredito + 1
    
    num = Val(TextNumeroNotaCredito.text)
    
    If TextCuit.text <> "" Then
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

    TextPrecioUnitario.text = ""
    TextPorDescuento.text = ""
    TextCantidad.text = ""
    TextTotalLineaProd.text = ""
    TextPorDesc.text = ""
    
    Dim precioUnitario As Double
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
   
      
    Dim busca1 As String
    busca1 = RTrim(LTrim(ComboArticulo.text))
   
    
    rstProductos.FindFirst "CodProd >= '" & busca1 & "' "
    
   
    TextDescripcion.text = rstProductos.Fields!Descripcion
    TextUnidadMedida.text = rstProductos.Fields!UnidadMedida
    
    If TextTipoNotaCredito.text = "B" Then
        precioUnitario = rstProductos.Fields!PrecioUnitarioFactura * 1.21
        TextPrecioUnitario.text = Format(precioUnitario, "#0.00")
     Else
        TextPrecioUnitario.text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#0.00")
    End If

    
    'TextPrecioUnitario.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#0.00")
      
   ' TextCantidad.SetFocus
    
    
    
End Sub
Private Sub ComboArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'Call calculoprecios
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub
Private Sub ComboArticulo_LostFocus()
    
    TextPrecioUnitario.text = ""
    TextPorDescuento.text = ""
    TextCantidad.text = ""
    TextTotalLineaProd.text = ""
    TextPorDesc.text = ""
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    'If KeyAscii = 13 Then
    If ComboArticulo.text <> "" Then
           Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
        
   
           tProductos.Index = "PrimaryKey"
           
           tProductos.MoveFirst
           tProductos.Seek "=", ComboArticulo.text
           
           If Not tProductos.NoMatch Then
                produ = tProductos!CodProd
                'MsgBox (DepoOrigen)
                 Dim busca1 As String, busca2 As String
                 busca1 = RTrim(LTrim(ComboArticulo.text))
                 busca2 = busca1 + "z"
                                     
                 rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
                             
                 TextDescripcion.text = rstProductos.Fields!Descripcion
                 TextUnidadMedida.text = rstProductos.Fields!UnidadMedida
                 
                 If TextTipoNotaCredito.text = "B" Then
                    precioUnitario = rstProductos.Fields!PrecioUnitarioFactura * 1.21
                    TextPrecioUnitario.text = Format(precioUnitario, "#0.00")
                  Else
                    TextPrecioUnitario.text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#0.00")
                 End If
                 
                 'TextPrecioUnitario.Text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#0.00")
                
           Else
                mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                ComboArticulo.text = ""
                TextDescripcion.text = ""
                TextUnidadMedida.text = ""
                TextPrecioUnitario.text = ""
           End If
              
           tProductos.Close
          'TextCantidad.SetFocus
    'End If
    End If
    
End Sub


Private Sub ComboVendedor_GotFocus()
    ComboVendedor.SelLength = Len(ComboVendedor.text)
End Sub

Private Sub TextApellidoNombre_Change()
    Columna = 1
    Call FiltrarGrilla(MSFlexGrid1, TextApellidoNombre, CLng(Columna))
End Sub
Private Sub FiltrarGrilla(MSFlexGrid1 As Object, TBox As TextBox, Columna As Long)
    
    Dim A As Integer
    
    
    If (KeyRetroceso Or Len(TBox.text) = 0) Then
        'KeyRetroceso = False
        'Exit Sub
        TBox.text = ""
    End If
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")

    
    MSFlexGrid1.Rows = 2
    MSFlexGrid1.Clear
    MSFlexGrid1.Visible = True
    
    
    Call titulos
   
    A = Len(TBox.text)

    If A >= 4 Then
    
        vSQL = "SELECT * FROM Clientes WHERE RazonSocial Like '*" & TBox.text & "*' ORDER BY RazonSocial"
        
        Set tClientes = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
        
        
        linea2 = 1
        
        Do While Not tClientes.EOF()
                MSFlexGrid1.AddItem " "
                MSFlexGrid1.Row = linea2
            
            
                MSFlexGrid1.Col = 0
                MSFlexGrid1.text = tClientes.Fields!IdCliente
                
                With Me.MSFlexGrid1

                    MSFlexGrid1.ColAlignment(1) = flexAlignLeftTop
                    MSFlexGrid1.Col = 0
                    MSFlexGrid1.text = tClientes.Fields!IdCliente
                    MSFlexGrid1.Col = 1
                    MSFlexGrid1.text = tClientes.Fields!RazonSocial
                    MSFlexGrid1.Col = 2
                    If tClientes.Fields!CUIT <> "" Then
                       MSFlexGrid1.text = tClientes.Fields!CUIT
                    End If
                    MSFlexGrid1.Col = 3
                    MSFlexGrid1.text = tClientes.Fields!Domicilio
                    MSFlexGrid1.Col = 4
                    MSFlexGrid1.text = tClientes.Fields!localidad
                    MSFlexGrid1.Col = 5
                    If tClientes.Fields!CP <> "" Then
                        MSFlexGrid1.text = tClientes.Fields!CP
                    End If
                    MSFlexGrid1.Col = 6
                    MSFlexGrid1.text = tClientes.Fields!Prov
                    MSFlexGrid1.Col = 7
                    MSFlexGrid1.text = tClientes.Fields!PorcentajeDescuento
                    
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
    TextCantidad.SelLength = Len(TextCantidad.text)
End Sub

Private Sub TextCodigoCliente_GotFocus()
    TextCodigoCliente.SelLength = Len(TextCodigoCliente.text)
End Sub

Private Sub TextDescripcion_GotFocus()
    TextDescripcion.SelLength = Len(TextDescripcion.text)
End Sub

Private Sub TextDescripcion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

End Sub




Private Sub TextNumeroNotaCredito_GotFocus()
    TextNumeroNotaCredito.SelLength = Len(TextNumeroNotaCredito.text)
End Sub


Private Sub TextNumeroNotaCredito_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextPorDescuento_GotFocus()
    TextPorDescuento.SelLength = Len(TextPorDescuento.text)
End Sub

Private Sub TextPrecioUnitario_GotFocus()
    TextPrecioUnitario.SelLength = Len(TextPrecioUnitario.text)
End Sub

Private Sub TextUnidadMedida_GotFocus()
    TextUnidadMedida.SelLength = Len(TextUnidadMedida.text)
End Sub

Private Sub TextUnidadMedida_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

End Sub
Private Sub TextPrecioUnitario_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
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
        Sendkeys "{TAB}"
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
            Sendkeys "{TAB}"
    End If
    
End Sub

Private Sub TextTotalLineaProd_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

End Sub
Private Sub calculoprecios()

    Dim precioUnitario As Double
    Dim porcentaje As Double
    Dim cantidad As Double
    Dim total As Double
    Dim porcentajePrecioUnitario As Double
    Dim totalLinea As Double
    
    'If KeyAscii = 13 Then
    If TextCantidad.text <> "" Then
         
        precioUnitario = Format(TextPrecioUnitario.text, "#00.00")
        If TextPorDescuento.text <> "" Then
            porcentaje = Val(TextPorDescuento.text)
        Else
            porcentaje = Val(TextDescuentoCliente.text)
        End If
        cantidad = CDbl(TextCantidad.text)
        total = (precioUnitario * cantidad)
        porcentajePrecioUnitario = ((precioUnitario * porcentaje) / 100) * cantidad
        totalLinea = total - ((total * porcentaje) / 100)
        TextTotalLineaProd.text = Format(totalLinea, "#00.00")
        TextPorDesc.text = Format(porcentajePrecioUnitario, "#00.00")
             
    End If
         
    ' End If
    

End Sub

Private Sub TextCantidad_LostFocus()
    
    If TextCantidad.text = "" Then
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
    busca1 = RTrim(LTrim(ComboVendedor.text))
    busca2 = busca1 + "z"
    
    rstEmpleado.FindFirst "Nombre >= '" & busca1 & "' and Nombre <= '" & busca2 & "'"
    
    If rstEmpleado.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Clientes", vbCritical, "Final de la busqueda")
       ComboVendedor.text = ""
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
    FG1.text = rstProductos.Fields!Descripcion
    FG1.Col = 2
    FG1.text = rstProductos.Fields!UnidadMedida
    FG1.Col = 3
    FG1.text = Format(rstProductos.Fields!PrecioUnitarioFactura, "#00.00")
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
    FormNotaCredito.Width = 13600
    FormNotaCredito.Top = 600
    FormNotaCredito.Left = 50
        
    Fila = 1
    renglon = 16
    
    Call SeteoGrilla
      
    Call Cargo
    Call buscoarticulo
    
    TextFechaNotaCredito.text = Format(Date, "dd/mm/yyyy")
    
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
    FG1.text = "Articulo"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 5800
    FG1.CellFontBold = True
    FG1.text = "Descripción"
    FG1.ColAlignment(1) = flexAlignCenterCenter
    
    FG1.Col = 2
    FG1.ColWidth(2) = 700
    FG1.CellFontBold = True
    FG1.text = "UM"
    FG1.ColAlignment(2) = flexAlignCenterCenter
    
    FG1.Col = 3
    FG1.ColWidth(3) = 1100
    FG1.CellFontBold = True
    FG1.text = "Precio Unit."
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 720
    FG1.CellFontBold = True
    FG1.text = "% Desc."
    FG1.ColAlignment(4) = flexAlignCenterCenter
    
    FG1.Col = 5
    FG1.ColWidth(5) = 900
    FG1.CellFontBold = True
    FG1.text = "Cantidad"
    FG1.ColAlignment(5) = flexAlignCenterCenter
        
    FG1.Col = 6
    FG1.ColWidth(6) = 1100
    FG1.CellFontBold = True
    FG1.text = "Total Línea"
    FG1.ColAlignment(6) = flexAlignCenterCenter
    
    FG1.Col = 7
    FG1.ColWidth(7) = 0
    FG1.CellFontBold = True
    FG1.text = "Importe Descuento"
    
    FG1.Col = 8
    FG1.ColWidth(8) = 0
    FG1.CellFontBold = True
    FG1.text = "Item"
    
    FG1.Row = 1
    item = 1
    linea = 1
    Do While FG1.Row <= 14
        FG1.Col = 8
        FG1.text = item
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
    busca1 = RTrim(LTrim(TextApellidoNombre.text))
    busca2 = busca1 + "z"
    
    rstCliente.FindFirst "Razonsocial >= '" & busca1 & "' and RazonSocial <= '" & busca2 & "'"
    
    If rstCliente.NoMatch Then
       MSFlexGrid1.Visible = False
       mensaje = MsgBox("No existen Clientes", vbCritical, "Final de la busqueda")
       TextApellidoNombre.text = ""
       Call blanco
       TextApellidoNombre.SetFocus
    End If
     
    linea2 = 1
    Do While Not rstCliente.NoMatch
        MSFlexGrid1.AddItem " "
        MSFlexGrid1.Row = linea2
       
            MSFlexGrid1.Col = 0
            MSFlexGrid1.text = rstCliente.Fields!IdCliente
            MSFlexGrid1.Col = 1
            MSFlexGrid1.text = rstCliente.Fields!RazonSocial
            MSFlexGrid1.Col = 2
            MSFlexGrid1.text = rstCliente.Fields!CUIT
            MSFlexGrid1.Col = 3
            MSFlexGrid1.text = rstCliente.Fields!Domicilio
            MSFlexGrid1.Col = 4
            MSFlexGrid1.text = rstCliente.Fields!localidad
            MSFlexGrid1.Col = 5
            MSFlexGrid1.text = rstCliente.Fields!CP
            MSFlexGrid1.Col = 6
            MSFlexGrid1.text = rstCliente.Fields!Prov
            MSFlexGrid1.Col = 7
            MSFlexGrid1.text = rstCliente.Fields!PorcentajeDescuento
            linea2 = linea2 + 1
      
       rstCliente.FindNext "RazonSocial >= '" & busca1 & "' and RazonSocial <= '" & busca2 & "'"
       
    Loop
    
    FG1.Enabled = True
    
End Sub

Private Sub titulos()

    MSFlexGrid1.Row = 0
    
    MSFlexGrid1.Col = 0
    MSFlexGrid1.text = "Codigo"
    MSFlexGrid1.ColWidth(0) = 900
    
    MSFlexGrid1.Col = 1
    MSFlexGrid1.text = "Apellido y Nombre"
    MSFlexGrid1.ColWidth(1) = 4700
    
    MSFlexGrid1.Col = 2
    MSFlexGrid1.text = "CUIT"
    MSFlexGrid1.ColWidth(2) = 1200
    
    MSFlexGrid1.Col = 3
    MSFlexGrid1.text = "Direccion"
    MSFlexGrid1.ColWidth(3) = 0
    
    MSFlexGrid1.Col = 4
    MSFlexGrid1.text = "Localidad"
    MSFlexGrid1.ColWidth(4) = 0
    
    MSFlexGrid1.Col = 5
    MSFlexGrid1.text = "CP"
    MSFlexGrid1.ColWidth(5) = 0
    
    MSFlexGrid1.Col = 6
    MSFlexGrid1.text = "Provincia"
    MSFlexGrid1.ColWidth(6) = 0
    
    MSFlexGrid1.Col = 7
    MSFlexGrid1.text = "Porcentaje Descuento"
    MSFlexGrid1.ColWidth(7) = 0

    
 End Sub

Private Sub MSFlexGrid1_Click()
   
    
    MSFlexGrid1.Col = 0
    TextCodigoCliente.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 1
    TextApellidoNombre.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 2
    TextCuit.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 3
    TextDireccion.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 4
    TextLocalidad.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 5
    TextCodigoPostal.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 6
    TextProvincia.text = MSFlexGrid1.text
        
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

    TextCodigoCliente.text = ""
    TextApellidoNombre.text = ""
    TextCuit.text = ""
    TextDireccion.text = ""
    TextLocalidad.text = ""
    TextCodigoPostal.text = ""
    TextProvincia.text = ""
    
End Sub

Private Sub TextCodigoCliente_KeyPress(KeyAscii As Integer)

    TextAlicuota.text = ""
   
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
   

    If KeyAscii = 13 Then
        If TextCodigoCliente.text = "" Then
            TextApellidoNombre.SetFocus
        Else
            CodigoClie = Val(TextCodigoCliente.text)
      
            rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCliente.Fields!IdCliente <> Val(TextCodigoCliente.text) Then
                mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                TextCodigoCliente.text = ""
                Call blanqueototal
                TextCodigoCliente.SetFocus
            Else
                TextCodigoCliente.text = rstCliente.Fields!IdCliente
                TextApellidoNombre.text = rstCliente.Fields!RazonSocial
                If rstCliente.Fields!CUIT <> "" Then TextCuit.text = rstCliente.Fields!CUIT
                TextDireccion.text = rstCliente.Fields!Domicilio
                TextLocalidad.text = rstCliente.Fields!localidad
                TextCodigoPostal.text = rstCliente.Fields!CP
                TextProvincia.text = rstCliente.Fields!Prov
                TextDescuentoCliente.text = rstCliente.Fields!PorcentajeDescuento
                vendedorCliente = rstCliente.Fields!Vendedor
                Call buscocuilyvendedor
            End If
        End If
        TextNumeroNotaCredito.text = ""
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

Dim FD As String
Dim FH As String

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
    '*** Busco CUIT
    
    
    CodigoClie = Val(TextCodigoCliente.text)
      
    rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
    
    If rstCliente.Fields!CUIT <> "" Then TextCuit.text = rstCliente.Fields!CUIT
    codigovendedor = rstCliente!Vendedor
      
    'Set rstPadron = db.OpenRecordset("Padron", dbOpenTable)
    
    'rstPadron.Index = "CUIT"
    
    'With rstPadron
    '    rstPadron.Seek "=", TextCuit.text
    '    If .NoMatch = False Then
    '        Alicuota = !AlicuotaPercepcion
    '    End If
    'End With
    
    FD = Format(Date, "YYYY") + Format(Date, "MM") + "01"
    FH = Left(FD, 6)
    
    Select Case Format(Date, "MM")
        Case 1, 3, 5, 7, 8, 10, 12
            FH = FH + "31"
        
        Case 4, 6, 9, 11
            FH = FH + "30"
        
        Case 2
            FH = FH + "28"
    End Select
            
    Alicuota = ePadron(FD, FH, TextCuit.text)
    
    TextDescuentoCliente.text = rstCliente.Fields!PorcentajeDescuento
    
    '*** Busco Vendedor
    
    CodigoVend = codigovendedor
      
    rstEmpleado.FindFirst "Legajo >= '" & CodigoVend & "'"
    
    LegajoEmpleado = rstEmpleado.Fields!Legajo
    ComboVendedor.text = rstEmpleado.Fields!Nombre
    
    '*** Busco Saldo
    
   rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
    
   TextSaldoCliente.text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
   
    '*** Busco Condicion IVA
    
    condicionIva = rstCliente.Fields!condicionIva
    If condicionIva = "RI" Then
        TextTipoNotaCredito.text = "A"
    End If
    If condicionIva = "CF" Then
        TextTipoNotaCredito.text = "B"
    End If
    
    'Mas Condiciones para factura B agregado por PVS 25/04/2016
    
    If condicionIva = "EX" Then
        TextTipoNotaCredito.text = "B"
    End If
    
    If condicionIva = "NR" Then
        TextTipoNotaCredito.text = "B"
    End If
    
    If condicionIva = "MO" Then
        TextTipoNotaCredito.text = "A"
    End If

    If condicionIva = "RN" Then
        TextTipoNotaCredito.text = "B"
    End If
'************************************************************
    
    
    If TextTipoNotaCredito.text = "A" Then
        TextAlicuota.text = Format(Alicuota, "#0.00")
    End If
    
    txtTipoFacturaReferencia.text = TextTipoNotaCredito.text
    
    Call LlenoComboFacturasDeReferencia(txtTipoFacturaReferencia.text)
    
    BotonNueva.Enabled = True
    BotonNueva.SetFocus
    
    'Agregado por PVS 22/04/2016
        MSFlexGrid1.Visible = False
    
End Sub

Private Sub TextCuit_Change()

    If TextCuit.text <> "" Then
        BotonNueva.Enabled = True
    End If
        
End Sub




Private Sub TextProvincia_Change()

    If TextProvincia.text <> "" Then
        ComboVendedor.SetFocus
    End If
End Sub



