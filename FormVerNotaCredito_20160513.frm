VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormVerNotacredito 
   Caption         =   "Consulta Notas de Credito"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Height          =   4455
      Left            =   120
      TabIndex        =   37
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
         TabIndex        =   42
         Top             =   600
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
         TabIndex        =   41
         Top             =   600
         Width           =   1455
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
         TabIndex        =   40
         Top             =   600
         Width           =   1335
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
         TabIndex        =   39
         Top             =   600
         Width           =   1335
      End
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
         TabIndex        =   38
         Top             =   600
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3135
         Left            =   480
         TabIndex        =   43
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
         TabIndex        =   48
         Top             =   360
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
         TabIndex        =   47
         Top             =   360
         Width           =   825
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
         TabIndex        =   46
         Top             =   360
         Width           =   1665
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
         TabIndex        =   45
         Top             =   360
         Width           =   1335
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
         TabIndex        =   44
         Top             =   360
         Width           =   390
      End
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      TabIndex        =   22
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
         TabIndex        =   29
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
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         Left            =   5400
         TabIndex        =   36
         Top             =   240
         Width           =   765
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
         TabIndex        =   35
         Top             =   240
         Width           =   1905
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
         TabIndex        =   34
         Top             =   240
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
         Left            =   3600
         TabIndex        =   33
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
         Left            =   6960
         TabIndex        =   32
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
         Left            =   8760
         TabIndex        =   31
         Top             =   240
         Width           =   330
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
         TabIndex        =   30
         Top             =   240
         Width           =   1635
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   120
      TabIndex        =   7
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         TabIndex        =   19
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
         TabIndex        =   18
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
         TabIndex        =   17
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
         TabIndex        =   16
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
         TabIndex        =   15
         Top             =   960
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   4
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   4200
         TabIndex        =   6
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonImprimir 
         Caption         =   "&Imprimir"
         Height          =   750
         Left            =   6240
         TabIndex        =   5
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   0
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
         TabIndex        =   3
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
End
Attribute VB_Name = "FormVerNotacredito"
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

    X = -4
    Y = -4
    renglon = 0
    vNroRemito = "0002- "
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
                Printer.Print NotaCC!TotalIVA
            
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


Private Sub BotonSalir_Click()

    Unload FormVerNotacredito

End Sub

Sub SeteoGrilla()
    
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
    
   
    
End Sub


Private Sub Form_Load()

    FormVerNotacredito.Height = 8970
    FormVerNotacredito.Width = 12135
    FormVerNotacredito.Top = 1000
    FormVerNotacredito.Left = 1000
    
    numDoc = Val(FormMovimientosCuentaCorriente.TextNumeroDocumento)
    tipoDoc = FormMovimientosCuentaCorriente.TextTipodocumento
    codCli = Val(FormMovimientosCuentaCorriente.TextCodigoCliente)
    
    If Val(FormBuscarNotaCredito.TextA) = 1 Then
        codCli = Val(FormBuscarNotaCredito.TextCodigoCliente)
        numDoc = Val(FormBuscarNotaCredito.TextNumeroFactura)
        Call SeteoGrilla
        Call busconotacredito
    Else
        Call SeteoGrilla
        Call busconotacredito
    End If
    
    Call SeteoGrilla
      
    Call busconotacredito
    

End Sub

Private Sub busconotacredito()

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
   
    TextCodigoCliente.Text = rstCliente.Fields!IDCliente
    TextApellidoNombre.Text = rstCliente.Fields!RazonSocial
    TextCuit.Text = rstCliente.Fields!Cuit
    TextDireccion.Text = rstCliente.Fields!Domicilio
    TextLocalidad.Text = rstCliente.Fields!Localidad
    TextCodigoPostal.Text = rstCliente.Fields!CP
    TextProvincia.Text = rstCliente.Fields!Prov
    TextDescuentoCliente.Text = rstCliente.Fields!PorcentajeDescuento
 
       
    Call SeteoGrilla
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db1 = DBEngine.OpenDatabase(ruta)
        
    Set rstNotaCreditoC = db1.OpenRecordset("NotaCreditoC", dbOpenDynaset)
        
    'rstNotCre1.Index = "PrimaryKey"
        
    'rstNotCre1.Seek "=", tipoDoc, numDoc
 
    rstNotaCreditoC.FindFirst "NroNotaCredito= " + Str(numDoc)
    
    TextNumeroNotaCredito.Text = rstNotaCreditoC.Fields!NroNotaCredito
    TextTipoNotaCredito.Text = rstNotaCreditoC.Fields!TipoNotaCredito
    TextFechaNotaCredito.Text = rstNotaCreditoC.Fields!FechaNotaCredito
    
    'rstNotCre1.Close
    'db1.Close
    
    rstNotaCreditoD.FindFirst "NroNotacredito= " + Str(numDoc)
    linea2 = 1
    Do While Not rstNotaCreditoD.NoMatch
        
            FG1.AddItem " "
            FG1.Row = linea2
       
            FG1.Col = 0
            FG1.Text = rstNotaCreditoD.Fields!IdCodProd
            
            FG1.Col = 0
            codigoprod = FG1.Text

            Dim busca1 As String, busca2 As String
            busca1 = RTrim(LTrim(codigoprod))
            busca2 = busca1 + "z"
       
            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
            FG1.Col = 1
            FG1.Text = rstProductos.Fields!Descripcion
        
            FG1.Col = 2
            FG1.Text = rstNotaCreditoD.Fields!UnidadMedida
            FG1.Col = 3
            FG1.Text = rstNotaCreditoD.Fields!precioUnitario
            FG1.Col = 4
            FG1.Text = rstNotaCreditoD.Fields!PorcentajeDescuento
            FG1.Col = 5
            FG1.Text = rstNotaCreditoD.Fields!cantidad
            FG1.Col = 6
            FG1.Text = rstNotaCreditoD.Fields!totalLinea
           
       
           rstNotaCreditoD.FindNext "NroNotaCredito= " + Str(numDoc)
           linea2 = linea2 + 1
    Loop
    
    
    '*** buscar vendedor
            
    codigovendedor = Val(rstNotaCreditoC.Fields!codVendedor)
         
    rstEmpleado.FindFirst "Legajo >= '" & codigovendedor & "'"
    ComboVendedor.Text = rstEmpleado.Fields!nombre

    '****
    
    TextSubtotalFactura.Text = rstNotaCreditoC.Fields!SubTotalNotaCredito
    TextDescuentos.Text = rstNotaCreditoC.Fields!ImporteDesc
    TextPercepcionIIBB.Text = rstNotaCreditoC.Fields!ImportePercepIIBB
    TextAlicuota.Text = rstNotaCreditoC.Fields!AlicuotaIIBB
    TextImpuesto.Text = rstNotaCreditoC.Fields!TotalIVA
    Textiva.Text = rstNotaCreditoC.Fields!PorcentajeIVA
    TextTotalFactura.Text = rstNotaCreditoC.Fields!TotalNotaCredito
    
    
    
End Sub

Public Function Descripcion(IdCodProd As Variant) As String

    Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    
    tProductos.Index = "PrimaryKey"
    
    tProductos.Seek "=", IdCodProd

    If Not tProductos.NoMatch Then Descripcion = tProductos!Descripcion

End Function

