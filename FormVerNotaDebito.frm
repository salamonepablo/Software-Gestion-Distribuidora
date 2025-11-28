VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormVerNotaDebito 
   Caption         =   "Consulta Nota Debito"
   ClientHeight    =   8445
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   4455
      Left            =   120
      TabIndex        =   41
      Top             =   1800
      Width           =   11655
      Begin VB.TextBox TextDescuentoCliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   7800
         TabIndex        =   48
         Top             =   600
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox ComboVendedor 
         Appearance      =   0  'Flat
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
         Height          =   315
         Left            =   5160
         TabIndex        =   47
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox TextFechaFactura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   3000
         TabIndex        =   46
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextNumeroFactura 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   45
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextTipoFactura 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   44
         Top             =   600
         Width           =   375
      End
      Begin VB.OptionButton opContado 
         Caption         =   "Contado"
         Height          =   255
         Left            =   9240
         TabIndex        =   43
         Top             =   360
         Width           =   1095
      End
      Begin VB.OptionButton opCtaCte 
         Caption         =   "Cta Cte"
         Height          =   255
         Left            =   9240
         TabIndex        =   42
         Top             =   720
         Width           =   855
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3135
         Left            =   480
         TabIndex        =   49
         Top             =   1080
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   5530
         _Version        =   393216
         Rows            =   15
         Cols            =   8
         FixedCols       =   0
         Enabled         =   0   'False
         GridLines       =   2
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
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
         Left            =   7560
         TabIndex        =   54
         Top             =   360
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label15 
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
         Left            =   5160
         TabIndex        =   53
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Debito"
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
         Left            =   3000
         TabIndex        =   52
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Debito"
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
         Left            =   600
         TabIndex        =   51
         Top             =   360
         Width           =   825
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
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
         Left            =   1920
         TabIndex        =   50
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      TabIndex        =   26
      Top             =   6240
      Width           =   11655
      Begin VB.TextBox TextSubtotalFactura 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   33
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextDescuentos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   32
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextPercepcionIIBB 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   31
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextAlicuota 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   4320
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextImpuesto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6120
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox Textiva 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   7800
         TabIndex        =   28
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextTotalFactura 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   27
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
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
         Left            =   4560
         TabIndex        =   40
         Top             =   240
         Width           =   765
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal Debito:"
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
         TabIndex        =   39
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
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
         Left            =   1920
         TabIndex        =   38
         Top             =   840
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
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
         TabIndex        =   37
         Top             =   240
         Width           =   1440
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
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
         Left            =   6360
         TabIndex        =   36
         Top             =   240
         Width           =   840
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         Left            =   8280
         TabIndex        =   35
         Top             =   240
         Width           =   330
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Debito:"
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
         Left            =   9840
         TabIndex        =   34
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   11655
      Begin VB.TextBox TextDireccion 
         Appearance      =   0  'Flat
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
         Left            =   7080
         TabIndex        =   18
         Top             =   600
         Width           =   4335
      End
      Begin VB.TextBox TextCodigoCliente 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   17
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TextApellidoNombre 
         Appearance      =   0  'Flat
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
         Left            =   7080
         TabIndex        =   16
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCuit 
         Appearance      =   0  'Flat
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
         Left            =   1920
         TabIndex        =   15
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextLocalidad 
         Appearance      =   0  'Flat
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
         Left            =   1080
         TabIndex        =   14
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Appearance      =   0  'Flat
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
         Left            =   5640
         TabIndex        =   13
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TextProvincia 
         Appearance      =   0  'Flat
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
         Left            =   8280
         TabIndex        =   12
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
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
         TabIndex        =   25
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
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
         TabIndex        =   24
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
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
         TabIndex        =   23
         Top             =   600
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
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
         TabIndex        =   22
         Top             =   600
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
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
         TabIndex        =   21
         Top             =   960
         Width           =   1230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
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
         TabIndex        =   20
         Top             =   960
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
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
         TabIndex        =   19
         Top             =   960
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   960
         TabIndex        =   10
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton BotonImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   750
         Left            =   3600
         TabIndex        =   9
         Top             =   240
         Width           =   1230
      End
      Begin VB.CommandButton cmdImprimirFE 
         Caption         =   "Imprimir &NDE"
         Height          =   735
         Left            =   6480
         TabIndex        =   8
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGenerarFE 
         Caption         =   "Generar &PDF"
         Height          =   735
         Left            =   9000
         TabIndex        =   7
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   11655
      Begin VB.OptionButton OptionFacturaPaga 
         Caption         =   "Pagas"
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
         Left            =   240
         TabIndex        =   5
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptionFacturaImpaga 
         Caption         =   "Impagas"
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
         Left            =   1200
         TabIndex        =   4
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.OptionButton OptionFacturaTodas 
         Caption         =   "Todas"
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
         Left            =   1320
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.PictureBox PictureQP 
      Height          =   975
      Left            =   12000
      ScaleHeight     =   915
      ScaleWidth      =   195
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.PictureBox PicBC 
      Height          =   975
      Left            =   12000
      ScaleHeight     =   915
      ScaleWidth      =   195
      TabIndex        =   0
      Top             =   3720
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "FormVerNotaDebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstNotaDebitoC As DAO.Recordset
 Dim rstDebC1 As DAO.Recordset
 Dim rstNotaDebitoD As DAO.Recordset
 Dim rstPadron As DAO.Recordset
 Dim cantidadProducto As Integer
 Dim descuentos As Double
 Dim LegajoEmpleado As Integer
 Dim numDoc As Long
 Dim tipoDoc As String
 Dim codCli As Integer
' Dim cl As New arisBarcode
 
Private Function BarCodeIL2of5(Cadena As String) As String
    
    Dim I As Long
    
    BarCodeIL2of5 = Chr(40)
    
    For I = 1 To Len(Cadena) Step 2
        If Val(Mid(Cadena, I, 2)) < 50 Then
          BarCodeIL2of5 = BarCodeIL2of5 & Chr(Val(Mid(Cadena, I, 2)) + 48)
        Else
          BarCodeIL2of5 = BarCodeIL2of5 & Chr(Val(Mid(Cadena, I, 2)) + 142)
        End If
    Next I
    
    BarCodeIL2of5 = BarCodeIL2of5 & Chr(41)


End Function

Private Sub GenerarFEB()
        
        'On Error GoTo CapturaErrores
        Dim NroFactura As String
        Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        Dim tCmp As Long
        
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaDebitoC = BaseSPC.OpenRecordset("NotaDebitoC", dbOpenTable)
'          Set tNotaDebitoD = BaseSPC.OpenRecordset("NotaDebitoD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaDebitoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaDebitoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tNotaDebitoC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tNotaDebitoC.Seek "=", "A", TextNumeroFactura.Text
           tNotaDebitoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaDebitoC.NoMatch Then
                
                If IsNull(tNotaDebitoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Nota de Debito no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    'Busco cual es la Impresora en PDF
                        For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        Next
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 80
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 07"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaDebitoC!NroRemito)
                                'LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                                'For I = 1 To LargoR
                                '    NroRemito = "0" & NroRemito
                                'Next I
                                
                               ' Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM NotaDebitoD WHERE NroDebito=" & tNotadebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaDebitoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100))
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format((PU * tNotaDebitoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                     '       .CurrentX = 135
                     '       .CurrentY = 245
                     '       .FontName = "Arial"
                     '       .FontSize = 10
                     '       '.FontBold = True
                     '       Printer.Print ("Sub Total: ")
                     '       .FontName = "Courier New"
                     '       SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalFactura), "Standard")
                     '       Hasta = CInt(14 - Len(SubTotalFac))
                     '       For I = 0 To Hasta
                     '           SubTotalFac = " " & SubTotalFac
                     '       Next I
                     '       .CurrentX = 165
                     '       .CurrentY = 245
                     '       Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                     '       .CurrentX = 135
                     '       .CurrentY = 250
                     '       .Font = "Arial"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                     '       .CurrentX = 165
                     '       .CurrentY = 250
                     '       .Font = "Courier New"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       ImpIva = Format(CDbl(tNotaDebitoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                     '           Hasta = CInt(14 - Len(ImpIIBB))
                     '           For I = 0 To Hasta
                     '               ImpIIBB = " " & ImpIIBB
                     '           Next I
                     '           Printer.Print ImpIIBB
                     '   End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
'///////////////// 'IMPRESION DE HOJA DUPLICADO ///////////////////////////////////////////////////////////////////////////
                    'Seteo de Tamaño de Papel
                        .NewPage
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                        
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 80
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 07"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaDebitoC!NroRemito)
                                'LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                                'For I = 1 To LargoR
                                '    NroRemito = "0" & NroRemito
                                'Next I
                                
                               ' Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM NotaDebitoD WHERE NroDebito=" & tNotadebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaDebitoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100))
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format((PU * tNotaDebitoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                     '       .CurrentX = 135
                     '       .CurrentY = 245
                     '       .FontName = "Arial"
                     '       .FontSize = 10
                     '       '.FontBold = True
                     '       Printer.Print ("Sub Total: ")
                     '       .FontName = "Courier New"
                     '       SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalFactura), "Standard")
                     '       Hasta = CInt(14 - Len(SubTotalFac))
                     '       For I = 0 To Hasta
                     '           SubTotalFac = " " & SubTotalFac
                     '       Next I
                     '       .CurrentX = 165
                     '       .CurrentY = 245
                     '       Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                     '       .CurrentX = 135
                     '       .CurrentY = 250
                     '       .Font = "Arial"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                     '       .CurrentX = 165
                     '       .CurrentY = 250
                     '       .Font = "Courier New"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       ImpIva = Format(CDbl(tNotaDebitoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                     '           Hasta = CInt(14 - Len(ImpIIBB))
                     '           For I = 0 To Hasta
                     '               ImpIIBB = " " & ImpIIBB
                     '           Next I
                     '           Printer.Print ImpIIBB
                     '   End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Nota de Débito Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub

Private Sub Imprimir()
    Dim PU, TL, ImpIva, ImpIIBB, SubTotalFac, TotalFac, Cant As Variant
    'PU = 0
    'TL = 0
    x = -4
    Y = -4
    renglon = 0
    vNroRemito = "0004- "
    
    With Printer
        
        'On Error GoTo CapturaErrores
        
        'Seteo escala a mm
            .ScaleMode = 6
            
        'Imprimir Fecha
            .CurrentX = x + 120
            .CurrentY = Y + 27
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print Format(TextFechaFactura.text, "DD/MM/YYYY")
        
        'Imprimir Nombres
            .CurrentX = x + 37
            .CurrentY = Y + 54
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = True
            Printer.Print TextApellidoNombre.text
            
        'Imprimir Direccion
            .CurrentX = x + 37
            .CurrentY = Y + 60
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextDireccion.text
            
        'Imprimir Localidad
            .CurrentX = x + 37
            .CurrentY = Y + 65
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextLocalidad.text
            
        'Imprimir CUIT
            .CurrentX = x + 125
            .CurrentY = Y + 67
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print TextCuit.text
            
        'Imprimir Marca Responsable Inscripto
            .CurrentX = x + 57
            .CurrentY = Y + 70
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca Contado
            .CurrentX = x + 70
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Marca CtaCte
            .CurrentX = x + 100
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print "X"
            
        'Imprimir Nro Remito
            .CurrentX = x + 138
            .CurrentY = Y + 80
            .Font = "Courier New"
            .FontSize = 9
            .FontBold = False
            Printer.Print vNroRemito
            
        'Imprimir Detalle
            
            vsqlFC = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & TextTipoFactura.text & "' AND NroDebito=" & TextNumeroFactura.text & " ORDER By NroDebito"
            vsqlFD = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & TextTipoFactura.text & "' AND NroDebito=" & TextNumeroFactura.text & " ORDER By NroDebito"
            
            Set DebC = BaseSPC.OpenRecordset(vsqlFC, dbOpenDynaset)
            Set DebD = BaseSPC.OpenRecordset(vsqlFD, dbOpenDynaset)
            
           
            DebC.MoveFirst
            DebD.MoveFirst
                
                    While Not DebD.EOF
                        'Imprimo el detalle
                            .CurrentX = x + 22
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Cant = CDbl(DebD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print DebD!cantidad
                            
                        'Detalle
                            .CurrentX = x + 40
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            'Printer.Print DebD!IdCodProd & Chr(9) & Descripcion(DebD!IdCodProd)
                            Printer.Print Descripcion(DebD!IdCodProd)
                        
                        'Precio
                            .CurrentX = x + 122
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            PU = CDbl(DebD!precioUnitario) - (CDbl(DebD!precioUnitario) * CDbl(DebD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU
                        
                        'Importe
                            .CurrentX = x + 142
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            TL = Format(DebD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                        
                         renglon = renglon + 5
                            
                        DebD.MoveNext
                    Wend
           
            'Importe SubTotal
                .CurrentX = x + 142
                .CurrentY = Y + 176
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                
                SubTotalFac = Format(CDbl(DebC!SubTotalFactura), "Standard")
                Hasta = CInt(14 - Len(SubTotalFac))
                For I = 0 To Hasta
                    SubTotalFac = " " & SubTotalFac
                Next I

                Printer.Print SubTotalFac
                
            'Alicuota IVA
                .CurrentX = x + 132
                .CurrentY = Y + 182
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print "21"
            
            'Importe IVA
                .CurrentX = x + 142
                .CurrentY = Y + 182
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                ImpIva = Format(CDbl(DebC!totalIva), "Standard")
                Hasta = CInt(14 - Len(ImpIva))
                For I = 0 To Hasta
                    ImpIva = " " & ImpIva
                Next I
                
                Printer.Print ImpIva
            
            If DebC!ImportePercepIIBB > 0 Then
                'Alicuota IIBB
                    .CurrentX = x + 122
                    .CurrentY = Y + 187
                    .Font = "Courier New"
                    .FontSize = 8
                    .FontBold = False
                    Printer.Print "Per.IIBB"
                
                'Importe IIBB
                    .CurrentX = x + 142
                    .CurrentY = Y + 187
                    .Font = "Courier New"
                    .FontSize = 8
                    .FontBold = False
                    
                    ImpIIBB = Format(CDbl(DebC!ImportePercepIIBB), "Standard")
                    Hasta = CInt(14 - Len(ImpIIBB))
                    For I = 0 To Hasta
                        ImpIIBB = " " & ImpIIBB
                    Next I
                    
                    
                    Printer.Print ImpIIBB
            End If
            
            'Importe Total
                .CurrentX = x + 142
                .CurrentY = Y + 194
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                
                TotalFac = Format(CDbl(DebC!TotalFactura), "Standard")
                Hasta = CInt(14 - Len(TotalFac))
                For I = 0 To Hasta
                    TotalFac = " " & TotalFac
                Next I
                
                Printer.Print TotalFac
            
        .EndDoc
        
    End With
    
    DebC.Close
    DebD.Close
        
CapturaErrores:
    'If Err = 321 Then
    'End If

End Sub

Public Function Descripcion(IdCodProd As Variant) As String

    Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    
    tProductos.Index = "PrimaryKey"
    
    tProductos.Seek "=", IdCodProd

    If Not tProductos.NoMatch Then Descripcion = tProductos!Descripcion

End Function
Private Sub ImprimirFEB()

        'On Error GoTo CapturaErrores
        Dim NroFactura As String
        Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaDebitoC = BaseSPC.OpenRecordset("NotaDebitoC", dbOpenTable)
'          Set tNotaDebitoD = BaseSPC.OpenRecordset("NotaDebitoD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaDebitoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaDebitoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tNotaDebitoC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tNotaDebitoC.Seek "=", "A", TextNumeroFactura.Text
           tNotaDebitoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaDebitoC.NoMatch Then
                
                If IsNull(tNotaDebitoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Nota de Debito no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    Set Printer = Printers(4)
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 80
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 07"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaDebitoC!NroRemito)
                                'LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                                'For I = 1 To LargoR
                                '    NroRemito = "0" & NroRemito
                                'Next I
                                
                               ' Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM NotaDebitoD WHERE NroDebito=" & tNotadebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaDebitoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100))
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format((PU * tNotaDebitoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                     '       .CurrentX = 135
                     '       .CurrentY = 245
                     '       .FontName = "Arial"
                     '       .FontSize = 10
                     '       '.FontBold = True
                     '       Printer.Print ("Sub Total: ")
                     '       .FontName = "Courier New"
                     '       SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalFactura), "Standard")
                     '       Hasta = CInt(14 - Len(SubTotalFac))
                     '       For I = 0 To Hasta
                     '           SubTotalFac = " " & SubTotalFac
                     '       Next I
                     '       .CurrentX = 165
                     '       .CurrentY = 245
                     '       Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                     '       .CurrentX = 135
                     '       .CurrentY = 250
                     '       .Font = "Arial"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                     '       .CurrentX = 165
                     '       .CurrentY = 250
                     '       .Font = "Courier New"
                     '       .FontSize = 10
                            '.FontBold = False
                     '       ImpIva = Format(CDbl(tNotaDebitoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                     '           Hasta = CInt(14 - Len(ImpIIBB))
                     '           For I = 0 To Hasta
                     '               ImpIIBB = " " & ImpIIBB
                     '           Next I
                     '           Printer.Print ImpIIBB
                     '   End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 15
                        .CurrentY = 245
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 15
                        .CurrentY = 252
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 12
                        .CurrentY = 263
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        .FontName = "Interleaved 2of5"
                        .FontSize = 20
                        Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Nota de Débito Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:


End Sub

Private Sub BotonImprimir_Click()

    Call Imprimir

End Sub

Private Sub BotonSalir_Click()

    Unload FormVerNotaDebito

End Sub

Private Sub ImprimirFE()

        'On Error GoTo CapturaErrores
        Dim NroFactura As String
        Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        Dim tCmp As Long
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaDebitoC = BaseSPC.OpenRecordset("NotaDebitoC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaDebitoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaDebitoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tNotaDebitoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaDebitoC.NoMatch Then
                
                If IsNull(tNotaDebitoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Nota de Débito no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    'Busco cual es la Impresora en PDF
                        'For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                        '    If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        'Next
                    Set Printer = Printers(5)
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 02"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaDebitoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                               ' For I = 1 To LargoR
                               '     NroRemito = "0" & NroRemito
                               ' Next I
                                
                               ' Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & tFacturaC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tFacturaD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format(tNotaDebitoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                            .CurrentX = 135
                            .CurrentY = 245
                            .FontName = "Arial"
                            .FontSize = 10
                            '.FontBold = True
                            Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalDebito), "Standard")
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                            Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                            .CurrentX = 165
                            .CurrentY = 250
                            .Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            ImpIva = Format(CDbl(tNotaDebitoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                                Hasta = CInt(14 - Len(ImpIIBB))
                                For I = 0 To Hasta
                                    ImpIIBB = " " & ImpIIBB
                                Next I
                                Printer.Print ImpIIBB
                        End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                    'Leyenda para Monotributo
                        If tClientes!condicionIva = "MO" Then
                            .FontBold = False
                            .FontName = "Arial"
                            .ForeColor = vbBlack
                            .FontSize = 7
                            .CurrentX = 40
                            .CurrentY = 242
                            Printer.Print "El crédito fiscal discriminado en el presente comprobante, solo podrá ser"
                            .CurrentX = 40
                            .CurrentY = 245
                            Printer.Print "computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal"
                            .CurrentX = 40
                            .CurrentY = 248
                            Printer.Print "para Pequeños Contribuyentes de la Ley Nº 27.618"
                        End If
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
'///////////////// 'IMPRESION DE HOJA DUPLICADO ///////////////////////////////////////////////////////////////////////////
                    'Seteo de Tamaño de Papel
                        .NewPage
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                        
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "DUPLICADO"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 02"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaDebitoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                               ' For I = 1 To LargoR
                               '     NroRemito = "0" & NroRemito
                               ' Next I
                                
                               ' Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & tFacturaC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tFacturaD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format(tNotaDebitoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                            .CurrentX = 135
                            .CurrentY = 245
                            .FontName = "Arial"
                            .FontSize = 10
                            '.FontBold = True
                            Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalDebito), "Standard")
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                            Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                            .CurrentX = 165
                            .CurrentY = 250
                            .Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            ImpIva = Format(CDbl(tNotaDebitoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                                Hasta = CInt(14 - Len(ImpIIBB))
                                For I = 0 To Hasta
                                    ImpIIBB = " " & ImpIIBB
                                Next I
                                Printer.Print ImpIIBB
                        End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                    'Leyenda para Monotributo
                        If tClientes!condicionIva = "MO" Then
                            .FontBold = False
                            .FontName = "Arial"
                            .ForeColor = vbBlack
                            .FontSize = 7
                            .CurrentX = 40
                            .CurrentY = 242
                            Printer.Print "El crédito fiscal discriminado en el presente comprobante, solo podrá ser"
                            .CurrentX = 40
                            .CurrentY = 245
                            Printer.Print "computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal"
                            .CurrentX = 40
                            .CurrentY = 248
                            Printer.Print "para Pequeños Contribuyentes de la Ley Nº 27.618"
                        End If
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Factura Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub

Private Sub cmdGenerarFE_Click()
    
    If TextTipoFactura.text = "A" Then
        Call GenerarFE
        'MsgBox ("Genera Duplicado")
        'Call GenerarFED
    End If

    If TextTipoFactura.text = "B" Then
        Call GenerarFEB
        'MsgBox ("Genera Duplicado")
        'Call GenerarFEBD
    End If
    
End Sub

Private Sub GenerarFE()

        'On Error GoTo CapturaErrores
        Dim NroFactura As String
        Dim NroRemito As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        Dim tCmp As Long
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaDebitoC = BaseSPC.OpenRecordset("NotaDebitoC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaDebitoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaDebitoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tNotaDebitoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaDebitoC.NoMatch Then
                
                If IsNull(tNotaDebitoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Nota de Débito no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    'Busco cual es la Impresora en PDF
                        For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        Next
                                             
                    'Seteo de Tamaño de Papel
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                    
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 02"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaDebitoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                               ' For I = 1 To LargoR
                               '     NroRemito = "0" & NroRemito
                               ' Next I
                                
                               ' Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & tFacturaC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tFacturaD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format(tNotaDebitoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                            .CurrentX = 135
                            .CurrentY = 245
                            .FontName = "Arial"
                            .FontSize = 10
                            '.FontBold = True
                            Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalDebito), "Standard")
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                            Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                            .CurrentX = 165
                            .CurrentY = 250
                            .Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            ImpIva = Format(CDbl(tNotaDebitoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                                Hasta = CInt(14 - Len(ImpIIBB))
                                For I = 0 To Hasta
                                    ImpIIBB = " " & ImpIIBB
                                Next I
                                Printer.Print ImpIIBB
                        End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                    'Leyenda para Monotributo
                        If tClientes!condicionIva = "MO" Then
                            .FontBold = False
                            .FontName = "Arial"
                            .ForeColor = vbBlack
                            .FontSize = 7
                            .CurrentX = 40
                            .CurrentY = 242
                            Printer.Print "El crédito fiscal discriminado en el presente comprobante, solo podrá ser"
                            .CurrentX = 40
                            .CurrentY = 245
                            Printer.Print "computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal"
                            .CurrentX = 40
                            .CurrentY = 248
                            Printer.Print "para Pequeños Contribuyentes de la Ley Nº 27.618"
                        End If
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
'///////////////// 'IMPRESION DE HOJA DUPLICADO ///////////////////////////////////////////////////////////////////////////
                    'Seteo de Tamaño de Papel
                        .NewPage
                        .ScaleHeight = 297
                        .ScaleWidth = 210
                        
                    'Imprimir el Logo
                        PictureQP.ScaleMode = 6
                        PictureQP.Picture = LoadPicture(App.Path & "\Quilplac.JPG")
                        Printer.PaintPicture PictureQP.Picture, 10, 9, 40, 10
                    
                    'Datos de La Empresa y Comprobante
                        .FontItalic = False
                        .DrawWidth = 10
                        Printer.Line (10, 7)-(200, 7)
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE DEBITO"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "DUPLICADO"
                        
                        .DrawWidth = 5
                        Printer.Line (93, 17)-(102, 17)
                        Printer.Line (93, 17)-(93, 25)
                        Printer.Line (102, 17)-(102, 25)
                        Printer.Line (93, 25)-(102, 25)
                        
                        .CurrentX = 95
                        .CurrentY = 16
                        .FontSize = 20
                        'Printer.Print "A"
                        Printer.Print TextTipoFactura.text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 02"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaDebitoC!NroDebito)
                        Largo = 8 - Len(tNotaDebitoC!NroDebito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaDebitoC!FechaDebito, "DD/MM/YYYY")
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 9
                        .FontBold = False
                        Printer.Print "C.U.I.T Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Ing.Brutos Nº 30-70843254-3"
                        .CurrentX = 150
                        Printer.Print "Inicio de Actividades: 11-06-2003"
                        .CurrentX = 150
                        Printer.Print "I.V.A. Responsable Inscripto"
                        
                        .DrawWidth = 10
                        Printer.Line (10, 42)-(200, 42)
                        
                    'Datos de la Empresa
                        .CurrentX = 12
                        .CurrentY = 20
                        .Font = "Arial"
                        .FontSize = 10
                        .FontBold = True
                        .FontUnderline = False
                        .FontSize = 10
                        Printer.Print "QUILPLAC S.A."
                        .CurrentX = 12
                        'Printer.Print "Andrés Baranda 520" & Chr(9) & "(1878) - Quilmes"
                        Printer.Print "Andrés Baranda 520 - CP (1878) - Quilmes"
                        .CurrentX = 12
                        Printer.Print "Pcia. Buenos Aires"
                        .CurrentX = 12
                        Printer.Print "Tel. 4257-5875"
                        
                        '.DrawWidth = 5
                        'Printer.Line (10, 27)-(50, 27)
                        '.CurrentX = 12
                        '.FontBold = True
                        '.FontSize = 8
                        '.CurrentY = 30
                        'Printer.Print "I.V.A. Responsable Inscripto"
                                
                    'Recuadro de datos del cliente
                        .DrawWidth = 10
                        Printer.Line (10, 47)-(200, 47)
                        Printer.Line (10, 47)-(10, 75)
                        Printer.Line (200, 47)-(200, 75)
                        Printer.Line (10, 75)-(200, 75)
                            
                    'Datos del Cliente
                        tClientes.MoveFirst
                        tClientes.Seek "=", tNotaDebitoC!CodCliente
                        If Not tClientes.NoMatch Then
                            
                            .CurrentX = 15
                            .CurrentY = 48
                            .FontSize = 10
                            .FontBold = True
                            Printer.Print "Señor(es): "
                            .CurrentX = 35
                            .CurrentY = 48
                            .FontBold = False
                            Printer.Print tClientes!RazonSocial
                            
                            .CurrentX = 130
                            .CurrentY = 48
                            .FontBold = True
                            Printer.Print "C.U.I.T Nº:"
                            .CurrentX = 150
                            .CurrentY = 48
                            .FontBold = False
                            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
                            Printer.Print CUIT
                             
                            tDomiciliosClientes.Seek "=", tClientes!IdCliente
                                If Not tDomiciliosClientes.NoMatch Then
                                  'Domicilio
                                    .CurrentX = 15
                                    .CurrentY = 55
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Domicilio: "
                                    .CurrentX = 35
                                    .CurrentY = 55
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!Domicilio
                                   
                                   'Localidad
                                    .CurrentX = 15
                                    .CurrentY = 62
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "Localidad: "
                                    .CurrentX = 35
                                    .CurrentY = 62
                                    .FontBold = False
                                     Printer.Print tDomiciliosClientes!localidad
                                     
                                    'Telefono
                                      .CurrentX = 130
                                      .CurrentY = 62
                                      .FontBold = True
                                      Printer.Print "Teléfono: "
                                      .CurrentX = 150
                                      .CurrentY = 62
                                      .FontBold = False
                                      Printer.Print tClientes!Tel
                                    
                                   'Condicion ante el IVA
                                    .CurrentX = 15
                                    .CurrentY = 69
                                    .FontSize = 10
                                    .FontBold = True
                                    Printer.Print "I.V.A: "
                                    .CurrentX = 35
                                    .CurrentY = 69
                                    .FontBold = False
                                     Printer.Print BuscarCondicionIva(tClientes!condicionIva)
                                End If
                         'Condiciones de venta
                            'Recuadro
                                .DrawWidth = 10
                                Printer.Line (10, 78)-(200, 78)
                                Printer.Line (10, 78)-(10, 85)
                                Printer.Line (200, 78)-(200, 85)
                                Printer.Line (10, 85)-(200, 85)
                                
                                .CurrentX = 15
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Condiciones de Venta: "
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                Printer.Print tNotaDebitoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaDebitoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaDebitoC!NroRemito)
                               ' For I = 1 To LargoR
                               '     NroRemito = "0" & NroRemito
                               ' Next I
                                
                               ' Printer.Print "0004-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                        .CurrentX = 18
                        .CurrentY = 92
                        .FontSize = 8
                        .FontBold = True
                        Printer.Print "CANTIDAD"
                        .DrawWidth = 5
                        Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 70
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "DETALLE"
                        Printer.Line (130, 91)-(130, 240)
                        
                        .CurrentX = 140
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "UNITARIO"
                        Printer.Line (165, 91)-(165, 240)
                        
                        .CurrentX = 175
                        .CurrentY = 92
                        .FontSize = 8
                        Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle de La Factura
                       
                        'vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & tFacturaC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tNotaDebitoC!TipoDebito & "' AND NroDebito=" & tNotaDebitoC!NroDebito & " ORDER BY NroDebito, ItemDebito"
                        'MsgBox (vSQL)
                        
                        Set tNotaDebitoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaDebitoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaDebitoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaDebitoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tFacturaD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaDebitoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaDebitoD!precioUnitario) - (CDbl(tNotaDebitoD!precioUnitario) * CDbl(tNotaDebitoD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            Hasta = CInt(10 - Len(PU))
                            For I = 0 To Hasta
                                PU = " " & PU
                            Next I
                            Printer.Print PU

                            Printer.Line (165, 91)-(165, 240)
                            
                            .CurrentX = 165
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            TL = Format(tNotaDebitoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaDebitoD.MoveNext
                            linea = .CurrentY + 3
                        Wend
                        
                        'Recuadro Subtotal / Total
                            Printer.Line (130, 240)-(130, 262)
                            Printer.Line (200, 240)-(200, 262)
                            Printer.Line (130, 240)-(130, 262)
                        
                        
                        'Importe SubTotal
                            .CurrentX = 135
                            .CurrentY = 245
                            .FontName = "Arial"
                            .FontSize = 10
                            '.FontBold = True
                            Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            SubTotalFac = Format(CDbl(tNotaDebitoC!SubTotalDebito), "Standard")
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            Printer.Print SubTotalFac
                            
                        'Alicuota IVA
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                            Printer.Print "I.V.A. 21%: "
                        
                        'Importe IVA
                            .CurrentX = 165
                            .CurrentY = 250
                            .Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            ImpIva = Format(CDbl(tNotaDebitoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaDebitoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaDebitoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaDebitoC!ImportePercepIIBB), "Standard")
                                Hasta = CInt(14 - Len(ImpIIBB))
                                For I = 0 To Hasta
                                    ImpIIBB = " " & ImpIIBB
                                Next I
                                Printer.Print ImpIIBB
                        End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = Format(CDbl(tNotaDebitoC!TotalDebito), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                            .Font = "Courier New"
                            .FontSize = 12
                            .CurrentX = 160
                            .CurrentY = 264
                            Printer.Print TotalFac
                        
                    'Leyenda para Monotributo
                        If tClientes!condicionIva = "MO" Then
                            .FontBold = False
                            .FontName = "Arial"
                            .ForeColor = vbBlack
                            .FontSize = 7
                            .CurrentX = 40
                            .CurrentY = 242
                            Printer.Print "El crédito fiscal discriminado en el presente comprobante, solo podrá ser"
                            .CurrentX = 40
                            .CurrentY = 245
                            Printer.Print "computado a efectos del Régimen de Sostenimiento e Inclusión Fiscal"
                            .CurrentX = 40
                            .CurrentY = 248
                            Printer.Print "para Pequeños Contribuyentes de la Ley Nº 27.618"
                        End If
                        
                        .FontBold = True
                        .FontName = "Arial"
                        .ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 45
                        .CurrentY = 255
                        Printer.Print "C.A.E: " & tNotaDebitoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaDebitoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaDebitoC!TipoDebito
                            Case "A"
                                tCmp = 2
                            Case "B"
                                tCmp = 7
                        End Select
                        
                        Call CrearQR(CStr(tNotaDebitoC!FechaDebito), 30708432543#, 4, tCmp, CDbl(tNotaDebitoC!NroDebito), CDbl(tNotaDebitoC!TotalDebito), "PES", 1, 80, CUITCliente(tNotaDebitoC!CodCliente), "E", CDbl(tNotaDebitoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_ND" & tNotaDebitoC!TipoDebito & "_" & "4_" & tNotaDebitoC!NroDebito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 24, 24
                        '.FontName = "Interleaved 2of5"
                        '.FontSize = 20
                        'Printer.Print BarCodeIL2of5(CalcularBarCode)
                        
                    .EndDoc
                End With
             Else
                A = MsgBox("Factura Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:
End Sub

Private Sub cmdImprimirFE_Click()

    If TextTipoFactura.text = "A" Then
        Call ImprimirFE
        'Call ImprimirFED
    End If
    
    If TextTipoFactura.text = "B" Then
        Call ImprimirFEB
        'Call ImprimirFEBD
    End If

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
       'Call blanco
       
    End If
     
   LegajoEmpleado = rstEmpleado.Fields!Legajo
End Sub


Sub SeteoGrilla()
    
    FG1.Row = 0
    FG1.Col = 0
    
    FG1.ColWidth(0) = 1000
    FG1.CellFontBold = True
    FG1.ColAlignment(0) = flexAlignCenterCenter
    FG1.text = "Articulo"
    
    FG1.Col = 1
    FG1.ColWidth(1) = 4700
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
    
   
    
End Sub


Private Sub Form_Load()

    FormVerNotaDebito.Height = 8970
    FormVerNotaDebito.Width = 12135
    FormVerNotaDebito.Top = 1000
    FormVerNotaDebito.Left = 1000
    
'    numDoc = Val(FormMovimientosCuentaCorriente.TextNumeroDocumento)
'    tipoDoc = FormMovimientosCuentaCorriente.TextTipodocumento
'    codCli = Val(FormMovimientosCuentaCorriente.TextCodigoCliente)
    
    If CDbl(FormBuscarNotaDebito.TextA) = 1 Then
        codCli = Val(FormBuscarNotaDebito.TextCodigoCliente)
        numDoc = Val(FormBuscarNotaDebito.TextNumeroFactura)
        tipoDoc = FormBuscarNotaDebito.TextTipo
        Call SeteoGrilla
        Call buscofactura
    End If

End Sub

Private Sub buscofactura()

     Dim tip As String

'    tip = tipoDoc
     

    ruta = App.Path & "\DB_SPC_SI.mdb"

'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstNotaDebitoC = db.OpenRecordset("NotaDebitoC", dbOpenDynaset)

'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstNotaDebitoD = db.OpenRecordset("NotaDebitoD", dbOpenDynaset)

    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)

    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)


    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)

    
    
    '**************Busco Cliente
    
    rstCliente.FindFirst "IDCliente= " + Str(codCli)
   
    TextCodigoCliente.text = rstCliente.Fields!IdCliente
    TextApellidoNombre.text = rstCliente.Fields!RazonSocial
    TextCuit.text = rstCliente.Fields!CUIT
    TextDireccion.text = rstCliente.Fields!Domicilio
    TextLocalidad.text = rstCliente.Fields!localidad
    TextCodigoPostal.text = rstCliente.Fields!CP
    TextProvincia.text = rstCliente.Fields!Prov
    TextDescuentoCliente.text = rstCliente.Fields!PorcentajeDescuento
 
       
'    Call SeteoGrilla
    
    '*************Busco Factura Cabecera
    
'    ruta = App.Path & "\DB_SPC_SI.mdb"
'
'    Set db1 = DBEngine.OpenDatabase(ruta)
'
'    Set rstNotaDebitoC = db1.OpenRecordset("NotaDebitoC", dbOpenDynaset)
'
'    rstNotaDebitoC.Index = "PrimaryKey"
'
'    rstNotaDebitoC.Seek "=", tipo, numDoc
'
'    rstNotaDebitoC.FindFirst "NroFactura= " + Str(numDoc)
    
    
    Set db1 = DBEngine.OpenDatabase(ruta)

    Set rstNotaDebitoC = db1.OpenRecordset("NotaDebitoC", dbOpenTable)

    rstNotaDebitoC.Index = "PrimaryKey"

    rstNotaDebitoC.Seek "=", tipoDoc, Str(numDoc)
   
    
    TextNumeroFactura.text = rstNotaDebitoC.Fields!NroDebito
    TextTipoFactura.text = rstNotaDebitoC.Fields!TipoDebito
    TextFechaFactura.text = rstNotaDebitoC.Fields!FechaDebito
    
    If rstNotaDebitoC.Fields!CondicionVenta = "Contado" Then opContado.Value = True
    If rstNotaDebitoC.Fields!CondicionVenta = "Cuenta Corriente" Then opCtaCte.Value = True
    
    'rstDebC1.Close
    'db1.Close
    
    '******************Busco Factura Detalle
    
'    rstNotaDebitoD.FindFirst "NroFactura= " + Str(numDoc)

    Set db2 = DBEngine.OpenDatabase(ruta)

    'Set rstNotaDebitoD = db2.OpenRecordset("NotaDebitoD", dbOpenTable)

    vSQL = "SELECT * FROM NotaDebitoD WHERE TipoDebito='" & tipoDoc & "' AND NroDebito=" & numDoc & " ORDER BY ItemDebito"

    'rstNotaDebitoD.Index = "PrimaryKey"

    'rstNotaDebitoD.Seek "=", tipoDoc, Str(numDoc)
    
    Set rstNotaDebitoD = db2.OpenRecordset(vSQL, dbOpenDynaset)
    
    rstNotaDebitoD.MoveFirst

    linea2 = 1
   While Not rstNotaDebitoD.EOF
            FG1.AddItem " "
            FG1.Row = linea2

            FG1.Col = 0
            FG1.text = rstNotaDebitoD.Fields!IdCodProd

            FG1.Col = 0
            codigoprod = FG1.text

            Dim busca1 As String, busca2 As String
            busca1 = RTrim(LTrim(codigoprod))
            busca2 = busca1 + "z"

            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"

            FG1.Col = 1
            FG1.text = rstProductos.Fields!Descripcion

            FG1.Col = 2
            FG1.text = rstNotaDebitoD.Fields!UnidadMedida
            FG1.Col = 3
            FG1.text = rstNotaDebitoD.Fields!precioUnitario
            FG1.Col = 4
            FG1.text = rstNotaDebitoD.Fields!PorcentajeDescuento
            FG1.Col = 5
            FG1.text = rstNotaDebitoD.Fields!cantidad
            FG1.Col = 6
            FG1.text = rstNotaDebitoD.Fields!totalLinea

           rstNotaDebitoD.MoveNext

           linea2 = linea2 + 1
    Wend
    
    
    '*** buscar vendedor
            
    codigovendedor = Val(rstNotaDebitoC.Fields!codVendedor)
         
    rstEmpleado.FindFirst "Legajo >= '" & codigovendedor & "'"
    ComboVendedor.text = rstEmpleado.Fields!Nombre

    '****
    
    TextSubtotalFactura.text = Format(rstNotaDebitoC.Fields!SubTotalDebito, "#,###,###,#0.00")
'    TextDescuentos.Text = rstNotaDebitoC.Fields!ImporteDesc
    TextPercepcionIIBB.text = Format(rstNotaDebitoC.Fields!ImportePercepIIBB, "#,###,###,#0.00")
    TextAlicuota.text = Format(rstNotaDebitoC.Fields!AlicuotaIIBB, "#,###,###,#0.00")
    TextImpuesto.text = Format(rstNotaDebitoC.Fields!totalIva, "#,###,###,#0.00")
    Textiva.text = Format(rstNotaDebitoC.Fields!PorcentajeIVA, "#,###,###,#0.00")
    TextTotalFactura.text = Format(rstNotaDebitoC.Fields!TotalDebito, "#,###,###,#0.00")
    
    
    
End Sub

Private Function BuscarCondicionIva(CI As String) As String
    
    Set tCondicionIVA = BaseSPC.OpenRecordset("CondicionIVA", dbOpenTable)

    tCondicionIVA.Index = "PrimaryKey"
    
    tCondicionIVA.Seek "=", CI

    If Not tCondicionIVA.NoMatch Then BuscarCondicionIva = tCondicionIVA!Descripcion
    
    tCondicionIVA.Close
    
End Function

Private Sub CrearBarCode(Texto As String)
    
    PicBC.FontName = Me.FontName
    PicBC.FontSize = Me.FontSize
    PicBC.Cls
    
    cl.Code128 PicBC, 0.5, Texto, True
    SavePicture PicBC.Picture, App.Path & "\BarCode.jpg"

End Sub

Private Function CalcularBarCode() As String
    
    Dim TipoC, FechaVC As String
    
    If tNotaDebitoC!TipoDebito = "A" Then TipoC = "02"
    If tNotaDebitoC!TipoDebito = "B" Then TipoC = "07"
    
    FechaVC = Year(tNotaDebitoC!FechaVC) & Format(Month(tNotaDebitoC!FechaVC), "00") & Format(Day(tNotaDebitoC!FechaVC), "00")
    
    'MsgBox (FechaVC)

    CalcularBarCode = "30708432543" & TipoC & "0004" & tNotaDebitoC!CAE & FechaVC & CalculoDigitoVerificador("30708432543")

End Function

Public Function CalculoDigitoVerificador(CUIT As String) As String

    Dim Texto As Variant
    Dim SumaImp, SumaPar, SumaTotal  As Long
    
    SumaImp = 0
    SumaPar = 0
    SumaTotal = 0
    
    Texto = CUIT
    
    For I = 1 To 11
        Select Case I
            Case 1
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 2
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 3
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 4
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 5
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 6
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 7
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 8
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 9
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case 10
                SumaPar = SumaPar + CInt(Mid(Texto, I, 1))
            Case 11
                SumaImp = SumaImp + CInt(Mid(Texto, I, 1))
            Case Else
        End Select
    Next I

    SumaImp = SumaImp * 3
    SumaTotal = SumaImp + SumaPar
    
    For J = 0 To 9
        
        If (SumaTotal + J) Mod (10) = 0 Then
            CalculoDigitoVerificador = CStr(J)
            Exit For
        End If
    Next J

   ' MsgBox (CalculoDigitoVerificador)
    
End Function


