VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormVerNotaCreditoCtaCte 
   Caption         =   "Consulta Cuenta Corriente"
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicBC 
      Height          =   375
      Left            =   1680
      ScaleHeight     =   315
      ScaleWidth      =   1275
      TabIndex        =   52
      Top             =   8400
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.PictureBox PictureQP 
      Height          =   375
      Left            =   120
      ScaleHeight     =   315
      ScaleWidth      =   1395
      TabIndex        =   51
      Top             =   8400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   45
      Top             =   1560
      Width           =   11655
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
         TabIndex        =   48
         Top             =   240
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
         TabIndex        =   47
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
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
         TabIndex        =   46
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   42
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton cmdImprimirNcE 
         Caption         =   "Imprimir NC E"
         Enabled         =   0   'False
         Height          =   735
         Left            =   3960
         TabIndex        =   50
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdGenerarNcE 
         Caption         =   "Generar PDF NC E"
         Height          =   735
         Left            =   7200
         TabIndex        =   49
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton BotonImprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   750
         Left            =   10200
         TabIndex        =   44
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   1080
         TabIndex        =   43
         Top             =   240
         Width           =   1590
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   120
      TabIndex        =   27
      Top             =   120
      Width           =   11655
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
         TabIndex        =   34
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
         TabIndex        =   33
         Top             =   960
         Width           =   1215
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
         TabIndex        =   32
         Top             =   960
         Width           =   3135
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
         TabIndex        =   31
         Top             =   600
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
         TabIndex        =   30
         Top             =   240
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
         TabIndex        =   29
         Top             =   240
         Width           =   1815
      End
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
         TabIndex        =   28
         Top             =   600
         Width           =   4335
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
         TabIndex        =   41
         Top             =   960
         Width           =   870
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
         TabIndex        =   40
         Top             =   960
         Width           =   900
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
         TabIndex        =   39
         Top             =   960
         Width           =   1230
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
         TabIndex        =   38
         Top             =   600
         Width           =   870
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
         TabIndex        =   37
         Top             =   600
         Width           =   510
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
         TabIndex        =   36
         Top             =   240
         Width           =   1455
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
         TabIndex        =   35
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      TabIndex        =   12
      Top             =   6240
      Width           =   11655
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
         TabIndex        =   19
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
         Left            =   8280
         TabIndex        =   18
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
         Left            =   6720
         TabIndex        =   17
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
         Left            =   5160
         TabIndex        =   16
         Top             =   480
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
         Left            =   3600
         TabIndex        =   15
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
         Left            =   2040
         TabIndex        =   14
         Top             =   480
         Width           =   1335
      End
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
         TabIndex        =   13
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
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
         TabIndex        =   26
         Top             =   240
         Width           =   1635
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
         Left            =   8760
         TabIndex        =   25
         Top             =   240
         Width           =   330
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
         Left            =   6960
         TabIndex        =   24
         Top             =   240
         Width           =   840
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
         Left            =   3600
         TabIndex        =   23
         Top             =   240
         Width           =   1440
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
         Left            =   2160
         TabIndex        =   22
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
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
         TabIndex        =   21
         Top             =   240
         Width           =   1905
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
         Left            =   5400
         TabIndex        =   20
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.Frame Frame5 
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   11655
      Begin VB.TextBox TextTipoNotaCredito 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TextNumeroNotaCredito 
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextFechaNotaCredito 
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
         Left            =   2400
         TabIndex        =   3
         Top             =   600
         Width           =   1335
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
         TabIndex        =   2
         Top             =   600
         Width           =   1455
      End
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
         TabIndex        =   1
         Top             =   600
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3135
         Left            =   480
         TabIndex        =   6
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
         Left            =   1560
         TabIndex        =   11
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
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
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
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
         Left            =   2280
         TabIndex        =   9
         Top             =   360
         Width           =   1665
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
         TabIndex        =   8
         Top             =   360
         Width           =   825
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
         TabIndex        =   7
         Top             =   360
         Width           =   930
      End
   End
End
Attribute VB_Name = "FormVerNotaCreditoCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstNotaCreditoC As DAO.Recordset
 Dim rstNotCre1 As DAO.Recordset
 Dim rstNotaCreditoD As DAO.Recordset
 Dim rstPadron As DAO.Recordset
 Dim cantidadProducto As Integer
 Dim descuentos As Double
 Dim LegajoEmpleado As Integer
 Dim numDoc As Long
 Dim tipoDoc As String
 Dim codCli As Integer


Private Sub BotonImprimir_Click()

    Call ImprimirNotaCredito

End Sub
Private Sub ImprimirNotaCredito()

    x = -4
    Y = -4
    renglon = 0
    vNroRemito = "0004- "
    '& TextNumeroRemito.Text
    
    With Printer
        
        'On Error GoTo CapturaErrores
            .Copies = 2
        'Seteo escala a mm
            .ScaleMode = 6
        
        'Imprimir Fecha
            .CurrentX = x + 120
            .CurrentY = Y + 27
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print Format(TextFechaNotaCredito.text, "DD/MM/YYYY")
        
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
            
            sqlfc = "SELECT * FROM NotaCreditoC WHERE TipoNotaCredito='" & TextTipoNotaCredito.text & "' AND NroNotaCredito=" & TextNumeroNotaCredito.text & " ORDER By NroNotaCredito"
            vsqlFD = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & TextTipoNotaCredito.text & "' AND NroNotaCredito=" & TextNumeroNotaCredito.text & " ORDER By NroNotaCredito"
            
            Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
                        
            Set NotaCC = BaseSPC.OpenRecordset(sqlfc, dbOpenDynaset)
            Set NotaCD = BaseSPC.OpenRecordset(vsqlFD, dbOpenDynaset)
            
           
            NotaCC.MoveFirst
            NotaCD.MoveFirst
                
                    While Not NotaCD.EOF
                        'Imprimo el detalle
                            .CurrentX = x + 20
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Printer.Print NotaCD!cantidad
                            
                        'Detalle
                            .CurrentX = x + 40
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Printer.Print NotaCD!IdCodProd & Chr(9) & Descripcion(NotaCD!IdCodProd)
                        
                        'Precio
                            .CurrentX = x + 123
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            PU = CDbl(NotaCD!precioUnitario) - (CDbl(NotaCD!precioUnitario) * CDbl(NotaCD!PorcentajeDescuento) / 100)
                            PU = Format(PU, "Standard")
                            
                            Printer.Print PU
                        
                        'Importe
                            .CurrentX = x + 143
                            .CurrentY = Y + 96 + renglon
                            .Font = "Courier New"
                            .FontSize = 8
                            .FontBold = False
                            Printer.Print NotaCD!totalLinea
                        
                         renglon = renglon + 5
                            
                        NotaCD.MoveNext
                    Wend
           
            'Importe SubTotal
                .CurrentX = x + 143
                .CurrentY = Y + 176
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print NotaCC!SubTotalNotaCredito
                
            'Alicuota IVA
                .CurrentX = x + 131
                .CurrentY = Y + 182
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print "21"
            
            'Importe IVA
                .CurrentX = x + 143
                .CurrentY = Y + 182
                .Font = "Courier New"
                .FontSize = 8
                .FontBold = False
                Printer.Print NotaCC!totalIva
            
            If NotaCC!ImportePercepIIBB > 0 Then
                'Alicuota IIBB
                    .CurrentX = x + 123
                    .CurrentY = Y + 187
                    .Font = "Courier New"
                    .FontSize = 8
                    .FontBold = False
                    Printer.Print "Per.IIBB"
                
                'Importe IIBB
                    .CurrentX = x + 143
                    .CurrentY = Y + 187
                    .Font = "Courier New"
                    .FontSize = 8
                    .FontBold = False
                    Printer.Print NotaCC!ImportePercepIIBB
            End If
            
            'Importe Total
                .CurrentX = x + 143
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


Private Sub BotonSalir_Click()

    Unload FormVerNotaCreditoCtaCte

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


Private Sub cmdGenerarNcE_Click()

'    If TextTipoNotaCredito.text = "A" Then
'        Call GenerarFE
        'MsgBox ("Genera Duplicado")
        'Call GenerarFED
'    End If

 '   If TextTipoNotaCredito.text = "B" Then
 '       Call GenerarFEB
        'MsgBox ("Genera Duplicado")
        'Call GenerarFEBD
 '   End If
 
 FormImprimirNC.Show

End Sub

Private Sub cmdImprimirNcE_Click()

    If TextTipoNotaCredito.text = "A" Then
        Call ImprimirFE
        'Call ImprimirFED
    End If
    
    If TextTipoNotaCredito.text = "B" Then
        Call ImprimirFEB
        'Call ImprimirFEBD
    End If

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
           
           Set tNotaCreditoC = BaseSPC.OpenRecordset("NotaCreditoC", dbOpenTable)
'          Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaCreditoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaCreditoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           tNotaCreditoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaCreditoC.NoMatch Then
                
                If IsNull(tNotaCreditoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
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
                        
                        .CurrentX = 78
                        .CurrentY = 9
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "NOTA DE CREDITO"
                        
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
                        Printer.Print "Código 03"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaCreditoC!NroNotaCredito)
                        Largo = 8 - Len(tNotaCreditoC!NroNotaCredito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaCreditoC!FechaNotaCredito, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", tNotaCreditoC!CodCliente
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
                                Printer.Print tNotaCreditoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaCreditoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaCreditoC!NroRemito)
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
                        vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tNotaCreditoC!TipoNotaCredito & "' AND NroNotaCredito=" & tNotaCreditoC!NroNotaCredito & " ORDER BY NroNotaCredito, ItemNotaCredito"
                        'MsgBox (vSQL)
                        
                        Set tNotaCreditoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaCreditoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaCreditoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaCreditoD!cantidad)
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
                            Printer.Print Descripcion(tNotaCreditoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaCreditoD!precioUnitario) - (CDbl(tNotaCreditoD!precioUnitario) * CDbl(tNotaCreditoD!PorcentajeDescuento) / 100)
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
                            TL = Format(tNotaCreditoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaCreditoD.MoveNext
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
                            SubTotalFac = Format(CDbl(tNotaCreditoC!SubTotalNotaCredito), "Standard")
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
                            ImpIva = Format(CDbl(tNotaCreditoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaCreditoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaCreditoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaCreditoC!ImportePercepIIBB), "Standard")
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
                            TotalFac = Format(CDbl(tNotaCreditoC!TotalNotaCredito), "Standard")
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
                        Printer.Print "C.A.E: " & tNotaCreditoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaCreditoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaCreditoC!TipoNotaCredito
                            Case "A"
                                tCmp = 3
                            Case "B"
                                tCmp = 8
                        End Select
                        
                        Call CrearQR(CStr(tNotaCreditoC!FechaNotaCredito), 30708432543#, 4, tCmp, CDbl(tNotaCreditoC!NroNotaCredito), CDbl(tNotaCreditoC!TotalNotaCredito), "PES", 1, 80, CUITCliente(tNotaCreditoC!CodCliente), "E", CDbl(tNotaCreditoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_NC" & tNotaCreditoC!TipoNotaCredito & "_" & "4_" & tNotaCreditoC!NroNotaCredito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 23, 23
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
                        Printer.Print "NOTA DE CREDITO"
                        
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
                        Printer.Print "Código 03"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaCreditoC!NroNotaCredito)
                        Largo = 8 - Len(tNotaCreditoC!NroNotaCredito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaCreditoC!FechaNotaCredito, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", tNotaCreditoC!CodCliente
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
                                Printer.Print tNotaCreditoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                               ' Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaCreditoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaCreditoC!NroRemito)
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
                        vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tNotaCreditoC!TipoNotaCredito & "' AND NroNotaCredito=" & tNotaCreditoC!NroNotaCredito & " ORDER BY NroNotaCredito, ItemNotaCredito"
                        'MsgBox (vSQL)
                        
                        Set tNotaCreditoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaCreditoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaCreditoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaCreditoD!cantidad)
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
                            Printer.Print Descripcion(tNotaCreditoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaCreditoD!precioUnitario) - (CDbl(tNotaCreditoD!precioUnitario) * CDbl(tNotaCreditoD!PorcentajeDescuento) / 100)
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
                            TL = Format(tNotaCreditoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaCreditoD.MoveNext
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
                            SubTotalFac = Format(CDbl(tNotaCreditoC!SubTotalNotaCredito), "Standard")
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
                            ImpIva = Format(CDbl(tNotaCreditoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaCreditoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaCreditoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaCreditoC!ImportePercepIIBB), "Standard")
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
                            TotalFac = Format(CDbl(tNotaCreditoC!TotalNotaCredito), "Standard")
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
                        Printer.Print "C.A.E: " & tNotaCreditoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaCreditoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaCreditoC!TipoNotaCredito
                            Case "A"
                                tCmp = 3
                            Case "B"
                                tCmp = 8
                        End Select
                        
                        Call CrearQR(CStr(tNotaCreditoC!FechaNotaCredito), 30708432543#, 4, tCmp, CDbl(tNotaCreditoC!NroNotaCredito), CDbl(tNotaCreditoC!TotalNotaCredito), "PES", 1, 80, CUITCliente(tNotaCreditoC!CodCliente), "E", CDbl(tNotaCreditoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_NC" & tNotaCreditoC!TipoNotaCredito & "_" & "4_" & tNotaCreditoC!NroNotaCredito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 23, 23
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


Private Function CalcularBarCode() As String
    
    Dim TipoC, FechaVC As String
    
    If tNotaCreditoC!TipoNotaCredito = "A" Then TipoC = "01"
    If tNotaCreditoC!TipoNotaCredito = "B" Then TipoC = "06"
    
    'FechaVC = Year(tFacturaC!FechaVC) & Month(tFacturaC!FechaVC) & Day(tFacturaC!FechaVC)
    FechaVC = Year(tNotaCreditoC!FechaVC) & Format(Month(tNotaCreditoC!FechaVC), "00") & Format(Day(tNotaCreditoC!FechaVC), "00")
    
  '  MsgBox (FechaVC)

    CalcularBarCode = "30708432543" & TipoC & "0004" & tNotaCreditoC!CAE & FechaVC & CalculoDigitoVerificador("30708432543")

End Function



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


Private Function BuscarCondicionIva(CI As String) As String
    
    Set tCondicionIVA = BaseSPC.OpenRecordset("CondicionIVA", dbOpenTable)

    tCondicionIVA.Index = "PrimaryKey"
    
    tCondicionIVA.Seek "=", CI

    If Not tCondicionIVA.NoMatch Then BuscarCondicionIva = tCondicionIVA!Descripcion
    
    tCondicionIVA.Close
    
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
           
           Set tNotaCreditoC = BaseSPC.OpenRecordset("NotaCreditoC", dbOpenTable)
'          Set tNotaCreditoD = BaseSPC.OpenRecordset("NotaCreditoD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaCreditoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaCreditoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
               ' TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tNotaCreditoC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           tNotaCreditoC.Seek "=", FormVerNotacredito.TextTipoNotaCredito.text, FormVerNotacredito.TextNumeroNotaCredito.text
            
           If Not tNotaCreditoC.NoMatch Then
                
                If IsNull(tNotaCreditoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    
                    'Busco cual es la Impresora en PDF
                        For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        Next
                    
                    'Set Printer = Printers(6)
                                             
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
                        Printer.Print "NOTA DE CREDITO"
                        
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
                        Printer.Print "A"
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 03"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaCreditoC!NroNotaCredito)
                        Largo = 8 - Len(tNotaCreditoC!NroNotaCredito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaCreditoC!FechaNotaCredito, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", tNotaCreditoC!CodCliente
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
                                Printer.Print tNotaCreditoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaCreditoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaCreditoC!NroRemito)
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
                       
                        'vSQL = "SELECT * FROM NotaCreditoD WHERE NroFactura=" & tNotaCreditoC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tNotaCreditoC!TipoNotaCredito & "' AND NroNotaCredito=" & tNotaCreditoC!NroNotaCredito & " ORDER BY NroNotaCredito, ItemNotaCredito"
                        'MsgBox (vSQL)
                        
                        Set tNotaCreditoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaCreditoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaCreditoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaCreditoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaCreditoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaCreditoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaCreditoD!precioUnitario) - (CDbl(tNotaCreditoD!precioUnitario) * CDbl(tNotaCreditoD!PorcentajeDescuento) / 100)
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
                            TL = Format(tNotaCreditoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaCreditoD.MoveNext
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
                            SubTotalFac = Format(CDbl(tNotaCreditoC!SubTotalNotaCredito), "Standard")
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
                            ImpIva = Format(CDbl(tNotaCreditoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaCreditoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaCreditoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaCreditoC!ImportePercepIIBB), "Standard")
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
                            TotalFac = Format(CDbl(tNotaCreditoC!TotalNotaCredito), "Standard")
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
                        Printer.Print "C.A.E: " & tNotaCreditoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaCreditoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaCreditoC!TipoNotaCredito
                            Case "A"
                                tCmp = 3
                            Case "B"
                                tCmp = 8
                        End Select
                        
                        Call CrearQR(CStr(tNotaCreditoC!FechaNotaCredito), 30708432543#, 4, tCmp, CDbl(tNotaCreditoC!NroNotaCredito), CDbl(tNotaCreditoC!TotalNotaCredito), "PES", 1, 80, CUITCliente(tNotaCreditoC!CodCliente), "E", CDbl(tNotaCreditoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_NC" & tNotaCreditoC!TipoNotaCredito & "_" & "4_" & tNotaCreditoC!NroNotaCredito & ".jpg"
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
                        Printer.Print "NOTA DE CREDITO"
                        
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
                        Printer.Print "A"
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        Printer.Print "Código 03"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaCreditoC!NroNotaCredito)
                        Largo = 8 - Len(tNotaCreditoC!NroNotaCredito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaCreditoC!FechaNotaCredito, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", tNotaCreditoC!CodCliente
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
                                Printer.Print tNotaCreditoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                               ' NroRemito = CStr(tNotaCreditoC!NroRemito)
                               ' LargoR = 8 - Len(tNotaCreditoC!NroRemito)
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
                       
                        'vSQL = "SELECT * FROM NotaCreditoD WHERE NroFactura=" & tNotaCreditoC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tNotaCreditoC!TipoNotaCredito & "' AND NroNotaCredito=" & tNotaCreditoC!NroNotaCredito & " ORDER BY NroNotaCredito, ItemNotaCredito"
                        'MsgBox (vSQL)
                        
                        Set tNotaCreditoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaCreditoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaCreditoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaCreditoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaCreditoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaCreditoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
                            PU = CDbl(tNotaCreditoD!precioUnitario) - (CDbl(tNotaCreditoD!precioUnitario) * CDbl(tNotaCreditoD!PorcentajeDescuento) / 100)
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
                            TL = Format(tNotaCreditoD!totalLinea, "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaCreditoD.MoveNext
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
                            SubTotalFac = Format(CDbl(tNotaCreditoC!SubTotalNotaCredito), "Standard")
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
                            ImpIva = Format(CDbl(tNotaCreditoC!totalIva), "Standard")
                            Hasta = CInt(14 - Len(ImpIva))
                            For I = 0 To Hasta
                                ImpIva = " " & ImpIva
                            Next I
                            
                            Printer.Print ImpIva
                        
                        If tNotaCreditoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                                .CurrentX = 135
                                .CurrentY = 255
                                .Font = "Arial"
                                .FontSize = 10
                                '.FontBold = False
                                Printer.Print "Per.IIBB: " & tNotaCreditoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                                .CurrentX = 165
                                .CurrentY = 255
                                .Font = "Courier New"
                                .FontSize = 10
                                '.FontBold = False
                                ImpIIBB = Format(CDbl(tNotaCreditoC!ImportePercepIIBB), "Standard")
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
                            TotalFac = Format(CDbl(tNotaCreditoC!TotalNotaCredito), "Standard")
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
                        Printer.Print "C.A.E: " & tNotaCreditoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaCreditoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaCreditoC!TipoNotaCredito
                            Case "A"
                                tCmp = 3
                            Case "B"
                                tCmp = 8
                        End Select
                        
                        Call CrearQR(CStr(tNotaCreditoC!FechaNotaCredito), 30708432543#, 4, tCmp, CDbl(tNotaCreditoC!NroNotaCredito), CDbl(tNotaCreditoC!TotalNotaCredito), "PES", 1, 80, CUITCliente(tNotaCreditoC!CodCliente), "E", CDbl(tNotaCreditoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_NC" & tNotaCreditoC!TipoNotaCredito & "_" & "4_" & tNotaCreditoC!NroNotaCredito & ".jpg"
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
                A = MsgBox("Nota de Crédito Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:
End Sub


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
           
           Set tNotaCreditoC = BaseSPC.OpenRecordset("NotaCreditoC", dbOpenTable)
'          Set tNotaCreditoD = BaseSPC.OpenRecordset("NotaCreditoD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaCreditoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaCreditoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tNotaCreditoC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tNotaCreditoC.Seek "=", "A", TextNumeroFactura.Text
           tNotaCreditoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaCreditoC.NoMatch Then
                
                If IsNull(tNotaCreditoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
                    Exit Sub
                End If
                
                With Printer
                    'Busco cual es la Impresora en PDF
                        For I = 0 To Printers.Count - 1
                            'List1.AddItem "Número:" & I & " - " & Printers(I).DeviceName
                            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
                        Next
                    
                    'Set Printer = Printers(6)
                                             
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
                        Printer.Print "NOTA DE CREDITO"
                        
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
                        Printer.Print "Código 08"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaCreditoC!NroNotaCredito)
                        Largo = 8 - Len(tNotaCreditoC!NroNotaCredito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaCreditoC!FechaNotaCredito, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", tNotaCreditoC!CodCliente
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
                                Printer.Print tNotaCreditoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaCreditoC!NroRemito)
                                'LargoR = 8 - Len(tNotaCreditoC!NroRemito)
                                'For I = 1 To LargoR
                                '    NroRemito = "0" & NroRemito
                                'Next I
                                
                                'Printer.Print "0004-" & NroRemito
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
                       
                        'vSQL = "SELECT * FROM NotaCreditoD WHERE NroFactura=" & tNotaCreditoC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tNotaCreditoC!TipoNotaCredito & "' AND NroNotaCredito=" & tNotaCreditoC!NroNotaCredito & " ORDER BY NroNotaCredito, ItemNotaCredito"
                        'MsgBox (vSQL)
                        
                        Set tNotaCreditoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaCreditoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaCreditoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaCreditoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaCreditoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaCreditoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaCreditoD!precioUnitario) - (CDbl(tNotaCreditoD!precioUnitario) * CDbl(tNotaCreditoD!PorcentajeDescuento) / 100))
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
                            TL = Format((PU * tNotaCreditoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaCreditoD.MoveNext
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
                     '       SubTotalFac = Format(CDbl(tNotaCreditoC!SubTotalFactura), "Standard")
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
                     '       ImpIva = Format(CDbl(tNotaCreditoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaCreditoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaCreditoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaCreditoC!ImportePercepIIBB), "Standard")
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
                            TotalFac = Format(CDbl(tNotaCreditoC!TotalNotaCredito), "Standard")
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
                        Printer.Print "C.A.E: " & tNotaCreditoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaCreditoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaCreditoC!TipoNotaCredito
                            Case "A"
                                tCmp = 3
                            Case "B"
                                tCmp = 8
                        End Select
                        
                        Call CrearQR(CStr(tNotaCreditoC!FechaNotaCredito), 30708432543#, 4, tCmp, CDbl(tNotaCreditoC!NroNotaCredito), CDbl(tNotaCreditoC!TotalNotaCredito), "PES", 1, 80, CUITCliente(tNotaCreditoC!CodCliente), "E", CDbl(tNotaCreditoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_NC" & tNotaCreditoC!TipoNotaCredito & "_" & "4_" & tNotaCreditoC!NroNotaCredito & ".jpg"
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
                        Printer.Print "NOTA DE CREDITO"
                        
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
                        Printer.Print "Código 08"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaCreditoC!NroNotaCredito)
                        Largo = 8 - Len(tNotaCreditoC!NroNotaCredito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaCreditoC!FechaNotaCredito, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", tNotaCreditoC!CodCliente
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
                                Printer.Print tNotaCreditoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaCreditoC!NroRemito)
                                'LargoR = 8 - Len(tNotaCreditoC!NroRemito)
                                'For I = 1 To LargoR
                                '    NroRemito = "0" & NroRemito
                                'Next I
                                
                                'Printer.Print "0004-" & NroRemito
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
                       
                        'vSQL = "SELECT * FROM NotaCreditoD WHERE NroFactura=" & tNotaCreditoC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tNotaCreditoC!TipoNotaCredito & "' AND NroNotaCredito=" & tNotaCreditoC!NroNotaCredito & " ORDER BY NroNotaCredito, ItemNotaCredito"
                        'MsgBox (vSQL)
                        
                        Set tNotaCreditoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaCreditoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaCreditoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaCreditoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaCreditoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaCreditoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaCreditoD!precioUnitario) - (CDbl(tNotaCreditoD!precioUnitario) * CDbl(tNotaCreditoD!PorcentajeDescuento) / 100))
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
                            TL = Format((PU * tNotaCreditoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaCreditoD.MoveNext
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
                     '       SubTotalFac = Format(CDbl(tNotaCreditoC!SubTotalFactura), "Standard")
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
                     '       ImpIva = Format(CDbl(tNotaCreditoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaCreditoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaCreditoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaCreditoC!ImportePercepIIBB), "Standard")
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
                            TotalFac = Format(CDbl(tNotaCreditoC!TotalNotaCredito), "Standard")
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
                        Printer.Print "C.A.E: " & tNotaCreditoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaCreditoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaCreditoC!TipoNotaCredito
                            Case "A"
                                tCmp = 3
                            Case "B"
                                tCmp = 8
                        End Select
                        
                        Call CrearQR(CStr(tNotaCreditoC!FechaNotaCredito), 30708432543#, 4, tCmp, CDbl(tNotaCreditoC!NroNotaCredito), CDbl(tNotaCreditoC!TotalNotaCredito), "PES", 1, 80, CUITCliente(tNotaCreditoC!CodCliente), "E", CDbl(tNotaCreditoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_NC" & tNotaCreditoC!TipoNotaCredito & "_" & "4_" & tNotaCreditoC!NroNotaCredito & ".jpg"
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
                A = MsgBox("Nota de Credito Inexistente !!!", vbCritical, "E R R O R")
        End If
    
CapturaErrores:

End Sub



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
        Dim tCmp As Long
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           Set tNotaCreditoC = BaseSPC.OpenRecordset("NotaCreditoC", dbOpenTable)
'          Set tNotaCreditoD = BaseSPC.OpenRecordset("NotaCreditoD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           
           tNotaCreditoC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           
           tNotaCreditoC.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tNotaCreditoC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tNotaCreditoC.Seek "=", "A", TextNumeroFactura.Text
           tNotaCreditoC.Seek "=", TextTipoFactura.text, TextNumeroFactura.text
            
           If Not tNotaCreditoC.NoMatch Then
                
                If IsNull(tNotaCreditoC!CAE) Then
                    b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
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
                        Printer.Print "NOTA DE CREDITO"
                        
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
                        Printer.Print "Código 08"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaCreditoC!NroNotaCredito)
                        Largo = 8 - Len(tNotaCreditoC!NroNotaCredito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaCreditoC!FechaNotaCredito, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", tNotaCreditoC!CodCliente
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
                                Printer.Print tNotaCreditoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaCreditoC!NroRemito)
                                'LargoR = 8 - Len(tNotaCreditoC!NroRemito)
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
                       
                        'vSQL = "SELECT * FROM NotaCreditoD WHERE NroFactura=" & tNotaCreditoC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tNotaCreditoC!TipoNotaCredito & "' AND NroNotaCredito=" & tNotaCreditoC!NroNotaCredito & " ORDER BY NroNotaCredito, ItemNotaCredito"
                        'MsgBox (vSQL)
                        
                        Set tNotaCreditoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaCreditoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaCreditoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaCreditoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaCreditoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaCreditoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaCreditoD!precioUnitario) - (CDbl(tNotaCreditoD!precioUnitario) * CDbl(tNotaCreditoD!PorcentajeDescuento) / 100))
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
                            TL = Format((PU * tNotaCreditoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaCreditoD.MoveNext
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
                     '       SubTotalFac = Format(CDbl(tNotaCreditoC!SubTotalFactura), "Standard")
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
                     '       ImpIva = Format(CDbl(tNotaCreditoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaCreditoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaCreditoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaCreditoC!ImportePercepIIBB), "Standard")
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
                            TotalFac = Format(CDbl(tNotaCreditoC!TotalNotaCredito), "Standard")
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
                        Printer.Print "C.A.E: " & tNotaCreditoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaCreditoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaCreditoC!TipoNotaCredito
                            Case "A"
                                tCmp = 3
                            Case "B"
                                tCmp = 8
                        End Select
                        
                        Call CrearQR(CStr(tNotaCreditoC!FechaNotaCredito), 30708432543#, 4, tCmp, CDbl(tNotaCreditoC!NroNotaCredito), CDbl(tNotaCreditoC!TotalNotaCredito), "PES", 1, 80, CUITCliente(tNotaCreditoC!CodCliente), "E", CDbl(tNotaCreditoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_NC" & tNotaCreditoC!TipoNotaCredito & "_" & "4_" & tNotaCreditoC!NroNotaCredito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 23, 23
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
                        Printer.Print "NOTA DE CREDITO"
                        
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
                        Printer.Print "Código 08"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        'En el numero de factura poner de la bbdd
                        NroFactura = CStr(tNotaCreditoC!NroNotaCredito)
                        Largo = 8 - Len(tNotaCreditoC!NroNotaCredito)
                        For I = 1 To Largo
                            NroFactura = "0" & NroFactura
                        Next I
                        Printer.Print "Nº: 0004-" & NroFactura
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(tNotaCreditoC!FechaNotaCredito, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", tNotaCreditoC!CodCliente
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
                                Printer.Print tNotaCreditoC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                                'NroRemito = CStr(tNotaCreditoC!NroRemito)
                                'LargoR = 8 - Len(tNotaCreditoC!NroRemito)
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
                       
                        'vSQL = "SELECT * FROM NotaCreditoD WHERE NroFactura=" & tNotaCreditoC!NroFactura & " ORDER BY NroFactura, ItemFactura"
                        vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tNotaCreditoC!TipoNotaCredito & "' AND NroNotaCredito=" & tNotaCreditoC!NroNotaCredito & " ORDER BY NroNotaCredito, ItemNotaCredito"
                        'MsgBox (vSQL)
                        
                        Set tNotaCreditoD = BaseSPC.OpenRecordset(vSQL)
                        
                        tNotaCreditoD.MoveFirst
                        linea = .CurrentY + 3
                        
                        While Not tNotaCreditoD.EOF
                            .FontBold = True
                            .CurrentX = 14
                            .CurrentY = linea
                            .Font = "Courier New"
                            .FontBold = True
                            .FontSize = 10
                           ' .FontBold = False
                            Cant = CDbl(tNotaCreditoD!cantidad)
                            Cant = Format(Cant, "Standard")
                            Hasta = CInt(6 - Len(Cant))
                            For I = 0 To Hasta
                                Cant = " " & Cant
                            Next I
                            Printer.Print Cant
                            'Printer.Print Format(tNotaCreditoD!cantidad, "Standard")
                            
                            
                            .CurrentX = 42
                            .CurrentY = linea
                            .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            Printer.Print Descripcion(tNotaCreditoD!IdCodProd)
                            
                            .CurrentX = 140
                            .CurrentY = linea
                            .FontSize = 10
                           ' .FontBold = False
            
                            PU = (CDbl(tNotaCreditoD!precioUnitario) - (CDbl(tNotaCreditoD!precioUnitario) * CDbl(tNotaCreditoD!PorcentajeDescuento) / 100))
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
                            TL = Format((PU * tNotaCreditoD!cantidad), "Standard")
                            Hasta = CInt(14 - Len(TL))
                            For I = 0 To Hasta
                                TL = " " & TL
                            Next I
                            Printer.Print TL
                            
                            tNotaCreditoD.MoveNext
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
                     '       SubTotalFac = Format(CDbl(tNotaCreditoC!SubTotalFactura), "Standard")
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
                     '       ImpIva = Format(CDbl(tNotaCreditoC!TotalIVA), "Standard")
                     '       Hasta = CInt(14 - Len(ImpIva))
                     '       For I = 0 To Hasta
                     '           ImpIva = " " & ImpIva
                     '       Next I
                            
                     '       Printer.Print ImpIva
                        
                     '   If tNotaCreditoC!ImportePercepIIBB > 0 Then
                            'Alicuota IIBB
                     '           .CurrentX = 135
                     '           .CurrentY = 255
                     '           .Font = "Arial"
                     '           .FontSize = 10
                                '.FontBold = False
                     '           Printer.Print "Per.IIBB: " & tNotaCreditoC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                     '           .CurrentX = 165
                     '           .CurrentY = 255
                     '           .Font = "Courier New"
                      ''          .FontSize = 10
                                '.FontBold = False
                     '           ImpIIBB = Format(CDbl(tNotaCreditoC!ImportePercepIIBB), "Standard")
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
                            TotalFac = Format(CDbl(tNotaCreditoC!TotalNotaCredito), "Standard")
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
                        Printer.Print "C.A.E: " & tNotaCreditoC!CAE
                        .CurrentX = 45
                        .CurrentY = 260
                        Printer.Print "Fecha Vencimiento C.A.E: " & Format(tNotaCreditoC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        .CurrentX = 15
                        .CurrentY = 260
                        'PicBC.ScaleMode = 6
                        'Printer.PaintPicture PicBC.Picture, 14, 257, 70, 12
                        
                        Select Case tNotaCreditoC!TipoNotaCredito
                            Case "A"
                                tCmp = 3
                            Case "B"
                                tCmp = 8
                        End Select
                        
                        Call CrearQR(CStr(tNotaCreditoC!FechaNotaCredito), 30708432543#, 4, tCmp, CDbl(tNotaCreditoC!NroNotaCredito), CDbl(tNotaCreditoC!TotalNotaCredito), "PES", 1, 80, CUITCliente(tNotaCreditoC!CodCliente), "E", CDbl(tNotaCreditoC!CAE))
                        
                        PicQR.ScaleMode = 6
                        'imgQR.Stretch = True
                        
                        DirectorioQRs = App.Path & "\QRs\" & "qr_NC" & tNotaCreditoC!TipoNotaCredito & "_" & "4_" & tNotaCreditoC!NroNotaCredito & ".jpg"
                        PicQR.Picture = LoadPicture(DirectorioQRs)
                        'imgQR.Picture = LoadPicture(DirectorioQRs)
                        
                        
                        'App.Path & "\QRs\qr.jpg"
                        'Printer.PaintPicture imgQR.Picture
                        
                        Printer.PaintPicture PicQR.Picture, 15, 245, 23, 23
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

Private Sub Form_Load()

    FormVerNotaCreditoCtaCte.Height = 8970
    FormVerNotaCreditoCtaCte.Width = 12135
    FormVerNotaCreditoCtaCte.Top = 1000
    FormVerNotaCreditoCtaCte.Left = 1000
    
    numDoc = Val(FormMovimientosCuentaCorriente.TextNumeroDocumento)
    tipoDoc = FormMovimientosCuentaCorriente.TextTipodocumento
    codCli = Val(FormMovimientosCuentaCorriente.TextCodigoCliente)

    If tipoDoc = "Nota Credito A" Then
        tipoDoc = "A"
    Else
        tipoDoc = "B"
    End If


    Call SeteoGrilla
    Call busconotacredito
 
' If Val(FormBuscarFactura.TextA) = 1 Then
'        codCli = Val(FormBuscarFactura.TextCodigoCliente)
'        numDoc = Val(FormBuscarFactura.TextNumeroFactura)
'        tipoDoc = FormBuscarFactura.TextTipo
'        Call SeteoGrilla
'        Call buscofactura
'End If
'
'
    
    
'    numDoc = Val(FormMovimientosCuentaCorriente.TextNumeroDocumento)
'    tipoDoc = FormMovimientosCuentaCorriente.TextTipodocumento
'    codCli = Val(FormMovimientosCuentaCorriente.TextCodigoCliente)
'
'    If Val(FormMovimientosCuentaCorriente.TextA) = 1 Then
'        codCli = Val(FormMovimientosCuentaCorriente.TextCodigoCliente)
'        numDoc = Val(FormMovimientosCuentaCorriente.TextNumeroFactura)
'        Call SeteoGrilla
'        Call busconotacredito
'    End If
'        Call SeteoGrilla
'        Call busconotacredito
'    End If
'
'    Call SeteoGrilla
'
'    Call busconotacredito
'

End Sub

Private Sub busconotacredito()

    Dim tip As String

    tip = tipoDoc

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstNotaCreditoC = db.OpenRecordset("NotaCreditoC", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstNotaCreditoD = db.OpenRecordset("NotaCreditoD", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
      
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
   
    rstCliente.FindFirst "IDCliente= " + Str(codCli)
   
    TextCodigoCliente.text = rstCliente.Fields!IdCliente
    TextApellidoNombre.text = rstCliente.Fields!RazonSocial
    TextCuit.text = rstCliente.Fields!CUIT
    TextDireccion.text = rstCliente.Fields!Domicilio
    TextLocalidad.text = rstCliente.Fields!localidad
    TextCodigoPostal.text = rstCliente.Fields!CP
    TextProvincia.text = rstCliente.Fields!Prov
    TextDescuentoCliente.text = rstCliente.Fields!PorcentajeDescuento
 
       
    Call SeteoGrilla
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
'    Set db1 = DBEngine.OpenDatabase(ruta)
        
'    Set rstNotaCreditoC = db1.OpenRecordset("NotaCreditoC", dbOpenDynaset)
'
'    rstNotaCreditoC.Index = "PrimaryKey"
'
'    rstNotaCreditoC.Seek "=", tipoDoc, numDoc
'
'    rstNotaCreditoC.FindFirst "NroNotaCredito= " + Str(numDoc)
    
    Set db1 = DBEngine.OpenDatabase(ruta)

    Set rstNotaCreditoC = db1.OpenRecordset("NotaCreditoC", dbOpenTable)

    rstNotaCreditoC.Index = "PrimaryKey"

    rstNotaCreditoC.Seek "=", tipoDoc, Str(numDoc)
    
    TextNumeroNotaCredito.text = rstNotaCreditoC.Fields!NroNotaCredito
    TextTipoNotaCredito.text = rstNotaCreditoC.Fields!TipoNotaCredito
    TextFechaNotaCredito.text = rstNotaCreditoC.Fields!FechaNotaCredito
    
    'rstNotCre1.Close
    'db1.Close
    
'    rstNotaCreditoD.FindFirst "NroNotacredito= " + Str(numDoc)


    Set db2 = DBEngine.OpenDatabase(ruta)
    
    vSQL = "SELECT * FROM NotaCreditoD WHERE TipoNotaCredito='" & tipoDoc & "' AND NroNotaCredito=" & numDoc & " ORDER BY ItemNotaCredito"

    Set rstNotaCreditoD = db2.OpenRecordset(vSQL, dbOpenDynaset)
    
    rstNotaCreditoD.MoveFirst
    
    linea2 = 1

    While Not rstNotaCreditoD.EOF
        
            FG1.AddItem " "
            FG1.Row = linea2
       
            FG1.Col = 0
            FG1.text = rstNotaCreditoD.Fields!IdCodProd
            
            FG1.Col = 0
            codigoprod = FG1.text

            Dim busca1 As String, busca2 As String
            busca1 = RTrim(LTrim(codigoprod))
            busca2 = busca1 + "z"
       
            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
            FG1.Col = 1
            FG1.text = rstProductos.Fields!Descripcion
        
            FG1.Col = 2
            FG1.text = rstNotaCreditoD.Fields!UnidadMedida
            FG1.Col = 3
            FG1.text = Format(rstNotaCreditoD.Fields!precioUnitario, "#,###,###,#0.00")
            FG1.Col = 4
            FG1.text = rstNotaCreditoD.Fields!PorcentajeDescuento
            FG1.Col = 5
            FG1.text = rstNotaCreditoD.Fields!cantidad
            FG1.Col = 6
            FG1.text = Format(rstNotaCreditoD.Fields!totalLinea, "#,###,###,#0.00")
           
       
           rstNotaCreditoD.MoveNext
           
           linea2 = linea2 + 1
     Wend
    
    

'    Set db2 = DBEngine.OpenDatabase(ruta)
'
'    Set rstNotaCreditoD = db2.OpenRecordset("NotaCreditoD", dbOpenTable)
'
'    rstNotaCreditoD.Index = "PrimaryKey"
'
'    rstNotaCreditoD.Seek "=", tipoDoc, Str(numDoc)
'
'
'    Set db2 = DBEngine.OpenDatabase(ruta)
'
'    Set rstFacturaD = db2.OpenRecordset("FacturaD", dbOpenTable)
'
'    rstFacturaD.Index = "PrimaryKey"
'
'    rstFacturaD.Seek "=", tipoDoc, Str(numDoc)
'
'
'
'    linea2 = 1
'     While Not rstNotaCreditoD.EOF
'
'
'            FG1.AddItem " "
'            FG1.Row = linea2
'
'            FG1.Col = 0
'            FG1.Text = rstNotaCreditoD.Fields!IdCodProd
'
'            FG1.Col = 0
'            codigoprod = FG1.Text
'
'            Dim busca1 As String, busca2 As String
'            busca1 = RTrim(LTrim(codigoprod))
'            busca2 = busca1 + "z"
'
'            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
'
'            FG1.Col = 1
'            FG1.Text = rstProductos.Fields!Descripcion
'
'            FG1.Col = 2
'            FG1.Text = rstNotaCreditoD.Fields!UnidadMedida
'            FG1.Col = 3
'            FG1.Text = rstNotaCreditoD.Fields!precioUnitario
'            FG1.Col = 4
'            FG1.Text = rstNotaCreditoD.Fields!PorcentajeDescuento
'            FG1.Col = 5
'            FG1.Text = rstNotaCreditoD.Fields!cantidad
'            FG1.Col = 6
'            FG1.Text = rstNotaCreditoD.Fields!totalLinea
'
'            rstNotaCreditoD.MoveNext
'           linea2 = linea2 + 1
'    Wend
'
    
    '*** buscar vendedor
            
    codigovendedor = Val(rstNotaCreditoC.Fields!codVendedor)
         
    rstEmpleado.FindFirst "Legajo >= '" & codigovendedor & "'"
    ComboVendedor.text = rstEmpleado.Fields!Nombre

    '****
    
    TextSubtotalFactura.text = Format(rstNotaCreditoC.Fields!SubTotalNotaCredito, "#,###,###,#0.00")
    TextDescuentos.text = Format(rstNotaCreditoC.Fields!ImporteDesc, "#,###,###,#0.00")
    TextPercepcionIIBB.text = Format(rstNotaCreditoC.Fields!ImportePercepIIBB, "#,###,###,#0.00")
    TextAlicuota.text = rstNotaCreditoC.Fields!AlicuotaIIBB
    TextImpuesto.text = Format(rstNotaCreditoC.Fields!totalIva, "#,###,###,#0.00")
    Textiva.text = rstNotaCreditoC.Fields!PorcentajeIVA
    TextTotalFactura.text = Format(rstNotaCreditoC.Fields!TotalNotaCredito, "#,###,###,#0.00")
    
    

'    ruta = App.Path & "\DB_SPC_SI.mdb"
'
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstNotaCreditoC = db.OpenRecordset("NotaCreditoC", dbOpenDynaset)
'
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstNotaCreditoD = db.OpenRecordset("NotaCreditoD", dbOpenDynaset)
'
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
'
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
'
'
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
'
'
'    rstCliente.FindFirst "IDCliente= " + Str(codCli)
'
'    TextCodigoCliente.Text = rstCliente.Fields!IdCliente
'    TextApellidoNombre.Text = rstCliente.Fields!RazonSocial
'    TextCuit.Text = rstCliente.Fields!CUIT
'    TextDireccion.Text = rstCliente.Fields!Domicilio
'    TextLocalidad.Text = rstCliente.Fields!Localidad
'    TextCodigoPostal.Text = rstCliente.Fields!CP
'    TextProvincia.Text = rstCliente.Fields!Prov
'    TextDescuentoCliente.Text = rstCliente.Fields!PorcentajeDescuento
'
'
'    Call SeteoGrilla
'
'    ruta = App.Path & "\DB_SPC_SI.mdb"
'
''    Set db1 = DBEngine.OpenDatabase(ruta)
''
''    Set rstNotaCreditoC = db1.OpenRecordset("NotaCreditoC", dbOpenDynaset)
''
''    'rstNotCre1.Index = "PrimaryKey"
''
''    'rstNotCre1.Seek "=", tipoDoc, numDoc
''
''    rstNotaCreditoC.FindFirst "NroNotaCredito= " + Str(numDoc)
'
'
'    Set db1 = DBEngine.OpenDatabase(ruta)
'
'    Set rstFacturaC = db1.OpenRecordset("FacturaC", dbOpenTable)
'
'    rstFacturaC.Index = "PrimaryKey"
'
'    rstFacturaC.Seek "=", tipoDoc, Str(numDoc)
'
'    TextNumeroNotaCredito.Text = rstNotaCreditoC.Fields!NroNotaCredito
'    TextTipoNotaCredito.Text = rstNotaCreditoC.Fields!TipoNotaCredito
'    TextFechaNotaCredito.Text = rstNotaCreditoC.Fields!FechaNotaCredito
'
'    'rstNotCre1.Close
'    'db1.Close
'
'    rstNotaCreditoD.FindFirst "NroNotacredito= " + Str(numDoc)
'    linea2 = 1
'    Do While Not rstNotaCreditoD.NoMatch
'
'            FG1.AddItem " "
'            FG1.Row = linea2
'
'            FG1.Col = 0
'            FG1.Text = rstNotaCreditoD.Fields!IdCodProd
'
'            FG1.Col = 0
'            codigoprod = FG1.Text
'
'            Dim busca1 As String, busca2 As String
'            busca1 = RTrim(LTrim(codigoprod))
'            busca2 = busca1 + "z"
'
'            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
'
'            FG1.Col = 1
'            FG1.Text = rstProductos.Fields!Descripcion
'
'            FG1.Col = 2
'            FG1.Text = rstNotaCreditoD.Fields!UnidadMedida
'            FG1.Col = 3
'            FG1.Text = rstNotaCreditoD.Fields!precioUnitario
'            FG1.Col = 4
'            FG1.Text = rstNotaCreditoD.Fields!PorcentajeDescuento
'            FG1.Col = 5
'            FG1.Text = rstNotaCreditoD.Fields!cantidad
'            FG1.Col = 6
'            FG1.Text = rstNotaCreditoD.Fields!totalLinea
'
'
'           rstNotaCreditoD.FindNext "NroNotaCredito= " + Str(numDoc)
'           linea2 = linea2 + 1
'    Loop
'
'
'    '*** buscar vendedor
'
'    codigovendedor = Val(rstNotaCreditoC.Fields!codVendedor)
'
'    rstEmpleado.FindFirst "Legajo >= '" & codigovendedor & "'"
'    ComboVendedor.Text = rstEmpleado.Fields!nombre
'
'    '****
'
'    TextSubtotalFactura.Text = rstNotaCreditoC.Fields!SubTotalNotaCredito
'    TextDescuentos.Text = rstNotaCreditoC.Fields!ImporteDesc
'    TextPercepcionIIBB.Text = rstNotaCreditoC.Fields!ImportePercepIIBB
'    TextAlicuota.Text = rstNotaCreditoC.Fields!AlicuotaIIBB
'    TextImpuesto.Text = rstNotaCreditoC.Fields!TotalIVA
'    Textiva.Text = rstNotaCreditoC.Fields!PorcentajeIVA
'    TextTotalFactura.Text = rstNotaCreditoC.Fields!TotalNotaCredito
'
    
    
End Sub

Public Function Descripcion(IdCodProd As Variant) As String

    Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    
    tProductos.Index = "PrimaryKey"
    
    tProductos.Seek "=", IdCodProd

    If Not tProductos.NoMatch Then Descripcion = tProductos!Descripcion

End Function


