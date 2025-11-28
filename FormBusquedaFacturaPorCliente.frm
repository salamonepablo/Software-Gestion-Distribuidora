VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormBusquedaFacturaPorCliente 
   Caption         =   "Busqueda Factura Por Cliente"
   ClientHeight    =   8400
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8400
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
      Height          =   3135
      Left            =   3240
      TabIndex        =   53
      Top             =   1680
      Visible         =   0   'False
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   5530
      _Version        =   393216
      Cols            =   4
      FixedCols       =   0
   End
   Begin VB.Frame Frame2 
      Height          =   135
      Left            =   120
      TabIndex        =   49
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
         TabIndex        =   52
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
         TabIndex        =   51
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
         TabIndex        =   50
         Top             =   360
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   4200
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   43
      Top             =   7200
      Width           =   11655
      Begin VB.CommandButton BotonImprimir 
         Caption         =   "&Imprimir"
         Height          =   750
         Left            =   2520
         TabIndex        =   47
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonCancelar 
         Caption         =   "&Cancelar"
         Height          =   750
         Left            =   3360
         TabIndex        =   46
         Top             =   240
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   4320
         TabIndex        =   45
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonGrabar 
         Caption         =   "&Guardar"
         Height          =   750
         Left            =   1560
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   750
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   120
      TabIndex        =   14
      Top             =   120
      Width           =   11655
      Begin VB.TextBox TextProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   20
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   19
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TextLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   18
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   17
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextApellidoNombre 
         Height          =   285
         Left            =   7080
         TabIndex        =   16
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   1920
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TextDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   15
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
         TabIndex        =   27
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
         TabIndex        =   26
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
         TabIndex        =   25
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
         TabIndex        =   24
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
         TabIndex        =   23
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
         TabIndex        =   22
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
         TabIndex        =   21
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      TabIndex        =   28
      Top             =   6240
      Width           =   11655
      Begin VB.TextBox TextTotalFactura 
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
         TabIndex        =   35
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
         Left            =   8280
         TabIndex        =   34
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
         TabIndex        =   33
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
         Left            =   5160
         TabIndex        =   32
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextPercepcionIIBB 
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
         Left            =   3600
         TabIndex        =   31
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
         Left            =   2040
         TabIndex        =   30
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextSubtotalFactura 
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
         TabIndex        =   29
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Factura:"
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
         TabIndex        =   42
         Top             =   240
         Width           =   1215
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
         TabIndex        =   41
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
         TabIndex        =   40
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
         TabIndex        =   39
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
         TabIndex        =   38
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Subtotal Factura:"
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
         TabIndex        =   37
         Top             =   240
         Width           =   1485
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
   End
   Begin VB.Frame Frame5 
      Height          =   3855
      Left            =   120
      TabIndex        =   2
      Top             =   2400
      Width           =   11655
      Begin VB.TextBox TextCantidadOriginal 
         Height          =   375
         Left            =   9840
         TabIndex        =   48
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox TextTipoFactura 
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
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   375
      End
      Begin VB.TextBox TextNumeroFactura 
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
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextFechaFactura 
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
         Left            =   2400
         TabIndex        =   5
         Top             =   600
         Width           =   1335
      End
      Begin VB.ComboBox ComboVendedor 
         Height          =   315
         Left            =   5160
         TabIndex        =   4
         Top             =   600
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
         Left            =   7800
         TabIndex        =   3
         Top             =   600
         Width           =   375
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2655
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   4683
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
         TabIndex        =   13
         Top             =   360
         Width           =   390
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Factura"
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
         TabIndex        =   12
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Factura"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1245
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   360
         Width           =   930
      End
   End
End
Attribute VB_Name = "FormBusquedaFacturaPorCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstFacturaC As DAO.Recordset
 Dim rstFacturaD As DAO.Recordset
 Dim rstPadron As DAO.Recordset
 Dim cantidadProducto As Integer
 Dim descuentos As Double
 Dim LegajoEmpleado As Integer
 Dim cantidadgrabada As Integer
 Dim Alicuota As Double



Private Sub BotonCancelar_Click()

    Call blanqueototal
    
End Sub
Private Sub blanqueototal()

    TextCodigoCliente.Text = ""
    TextApellidoNombre.Text = ""
    TextCuit.Text = ""
    TextDireccion.Text = ""
    TextLocalidad.Text = ""
    TextCodigoPostal.Text = ""
    TextProvincia.Text = ""
    TextNumeroFactura.Text = ""
    TextTipoFactura.Text = ""
    ComboVendedor.Text = ""
    TextDescuentoCliente.Text = ""
    TextSubtotalFactura.Text = ""
    TextDescuentos.Text = ""
    TextPercepcionIIBB.Text = ""
    TextAlicuota.Text = ""
    TextImpuesto.Text = ""
    Textiva.Text = ""
    TextTotalFactura.Text = ""
    FG1.Clear
    FG1.Enabled = False
    
    
    Call SeteoGrilla

End Sub

Private Sub BotonGrabar_Click()

        Dim descuentoCantidad As Integer

        ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstFacturaC = db.OpenRecordset("FacturaC", dbOpenDynaset)
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstFacturaD = db.OpenRecordset("FacturaD", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
        
        rstFacturaC.AddNew
        rstFacturaC.Fields!NroFactura = TextNumeroFactura.Text
        rstFacturaC.Fields!TipoFactura = TextTipoFactura.Text
        rstFacturaC.Fields!FechaFactura = TextFechaFactura.Text
        rstFacturaC.Fields!TotalFactura = TextTotalFactura.Text
        rstFacturaC.Fields!PorcentajeIVA = Textiva.Text
        rstFacturaC.Fields!SubTotalFactura = TextSubtotalFactura.Text
        rstFacturaC.Fields!TotalIVA = TextImpuesto.Text
        rstFacturaC.Fields!AlicuotaIIBB = TextAlicuota.Text
        rstFacturaC.Fields!ImportePercepIIBB = TextPercepcionIIBB.Text
        rstFacturaC.Fields!CodCliente = TextCodigoCliente.Text
        rstFacturaC.Fields!PorcentajeDesc = TextDescuentoCliente.Text
        rstFacturaC.Fields!ImporteDesc = TextDescuentos.Text
        rstFacturaC.Fields!codVendedor = LegajoEmpleado
        rstFacturaC.Update
        
        FG1.Col = 0
        FG1.Row = 1
        Filas = FG1.Rows
        linea = 1
        Do While linea < Filas
              
              FG1.Row = linea
              FG1.Col = 0
              If FG1.Text <> "" Then
                    rstFacturaD.AddNew
                
                    rstFacturaD.Fields!NroFactura = TextNumeroFactura.Text
                    rstFacturaD.Fields!TipoFactura = TextTipoFactura.Text
                
                    FG1.Col = 0
                    rstFacturaD.Fields!IdCodProd = FG1.Text
                
                    FG1.Col = 2
                    rstFacturaD.Fields!UnidadMedida = FG1.Text
                    
                    FG1.Col = 3
                    rstFacturaD.Fields!precioUnitario = Format(FG1.Text, "#00.00")
                    
                    FG1.Col = 4
                    des = FG1.Text
                    If des <> "" Then
                       rstFacturaD.Fields!PorcentajeDescuento = Val(des)
                    Else
                       rstFacturaD.Fields!PorcentajeDescuento = Val(TextDescuentoCliente.Text)
                    End If
                    FG1.Col = 5
                    rstFacturaD.Fields!cantidad = Val(FG1.Text)
                    descuentoCantidad = Val(FG1.Text)
                    
                    '*** Modifico Stock Producto
                    
                    FG1.Col = 0
                    codigoprod = FG1.Text
        
                    Dim busca1 As String, busca2 As String
                    busca1 = RTrim(LTrim(codigoprod))
                    busca2 = busca1 + "z"
               
                    rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
                    
                    rstProductos.Edit
                    rstProductos.Fields!Stock = cantidadProducto - descuentoCantidad
                    rstProductos.Update
                    
                    FG1.Col = 6
                    rstFacturaD.Fields!totalLinea = Format(FG1.Text, "#00.00")
                    
                    FG1.Col = 7
                    rstFacturaD.Fields!ImporteDescuento = Format(FG1.Text, "#00.00")
                     
                    rstFacturaD.Update
              End If
              linea = linea + 1
        Loop
        
        Call blanqueototal
        Call SeteoGrilla
   

End Sub

Private Sub BotonPago_Click()

    FormPagoFacturas.Show
    
End Sub

Private Sub BotonSalir_Click()

    Unload FormBusquedaFacturaPorCliente

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
       
    End If
     
   LegajoEmpleado = rstEmpleado.Fields!Legajo
End Sub

Private Sub CommandSalir_Click()

    Unload FormFacturaCliente

End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)

    Dim precioUnitario As Double
    Dim cantidad As Integer
    Dim porcentaje As Double
    Dim total
    Dim totalLinea As Double
    Dim totalGrilla
    Dim subtotalFacturaForm
    Dim porcentajePrecioUnitario As Double
    Dim descuentoFactura As Double
    Dim totalFacturaForm As Double
    Dim iva As Double
    Dim impuesto As Double
    Dim percepcion As Double
    Dim columnaSeis As Integer
    Dim columnaSiete As Integer

    iva = 21
    
    Textiva.Text = Format(iva, "#00.00")
    
        
    If KeyAscii >= 32 And KeyAscii <= 127 Then
        FG1.Text = FG1.Text & Chr(KeyAscii)
    End If

    Select Case KeyAscii
       Case 13
      
            ruta = App.Path & "\DB_SPC_SI.mdb"
    
            Set db = DBEngine.OpenDatabase(ruta)
            Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
                
            FG1.Col = 0
            codigoprod = FG1.Text
        
            Dim busca1 As String, busca2 As String
            busca1 = RTrim(LTrim(codigoprod))
            busca2 = busca1 + "z"
               
            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
            codigoProdTabla = rstProductos.Fields!CodProd
            
            'If codigoProdTabla <> RTrim(LTrim(codigoprod)) Then
            If rstProductos.NoMatch Then
                mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                codigoprod = ""
                FG1.Col = 1
                FG1.Text = ""
                FG1.Col = 2
                FG1.Text = ""
                FG1.Col = 3
                FG1.Text = ""
                FG1.Col = 4
                FG1.Text = ""
                FG1.Col = 0
                FG1.Text = ""
                FG1.SetFocus
            Else
                Call muestrodatosproductos
                FG1.Col = FG1.Col + 2
            End If
   
           
           
           '*** descuento
           If FG1.Col = 4 And FG1.Text <> "" Then
                If KeyAscii = 13 Then
                   'FG1.Col = FG1.Col + 1
                   FG1.Col = 3
                   precioUnitario = Val(FG1.Text)
                   FG1.Col = 4
                   porcentaje = Val(FG1.Text)
                   FG1.Col = 5
                   cantidad = Val(FG1.Text)
                   total = (precioUnitario * cantidad)
                   porcentajePrecioUnitario = ((precioUnitario * porcentaje) / 100) * cantidad
                   totalLinea = total - ((total * porcentaje) / 100)
                   FG1.Col = 6
                   FG1.Text = Format(totalLinea, "#00.00")
                   FG1.Col = 7
                   FG1.Text = Format(porcentajePrecioUnitario, "#00.00")
                End If
           End If
              
           '**** cantidad
           If FG1.Col = 5 And FG1.Text <> "" Then
                If KeyAscii = 13 Then
                    FG1.Col = FG1.Col + 1
                    FG1.Col = 3
                    precioUnitario = Format(FG1.Text, "#00.00")
                    FG1.Col = 4
                    If FG1.Text <> "" Then
                        porcentaje = Val(FG1.Text)
                    Else
                        porcentaje = TextDescuentoCliente.Text
                    End If
                    FG1.Col = 5
                    cantidad = Val(FG1.Text)
                    '*** verfico stock de producto
                    If cantidad > cantidadProducto Then
                        MsgBox "La cantidad ingresada supera al Stock Actual: " & cantidadProducto & ""
                        FG1.Col = 5
                        FG1.Text = TextCantidadOriginal.Text
                        FG1.SetFocus
                    Else
                        total = (precioUnitario * cantidad)
                        porcentajePrecioUnitario = ((precioUnitario * porcentaje) / 100) * cantidad
                        totalLinea = total - ((total * porcentaje) / 100)
                        FG1.Col = 6
                        FG1.Text = Format(totalLinea, "#00.00")
                        FG1.Col = 7
                        FG1.Text = Format(porcentajePrecioUnitario, "#00.00")
                    End If
                End If
            End If
                  
            '**** suma total linea
            
            columnaSeis = 6
             
            total = SumarTotalGrilla(FG1, columnaSeis)
            subtotalFacturaForm = total
                                    
            TextSubtotalFactura.Text = Format(total, "#00.00")
            
            '**** suma descuentos
            
            columnaSiete = 7
             
            porcentajePrecioUnitario = SumarTotalDescuentos(FG1, columnaSiete)
            descuentoFactura = porcentajePrecioUnitario
                                    
            TextDescuentos.Text = Format(descuentoFactura, "#00.00")
            
            '**** calculo alicuota
    
            TextAlicuota.Text = Format(Alicuota, "#0.00")
                            
            percepcion = (subtotalFacturaForm - descuentoFactura) * Alicuota / 100
            
            TextPercepcionIIBB.Text = Format(percepcion, "#0.00")
            
            '**** calculo impuesto
            
            impuesto = (subtotalFacturaForm - descuentoFactura) * iva / 100
            
            TextImpuesto.Text = Format(impuesto, "#00.00")
            
            '**** calculo total factura
            
            totalFacturaForm = (subtotalFacturaForm - descuentoFactura + percepcion + impuesto)
            
            TextTotalFactura.Text = Format(totalFacturaForm, "#00.00")
      
            If FG1.Col = 7 And FG1.Text <> "" Then
                FG1.Col = 0
                If FG1.Row < 2 Then
                    FG1.Row = FG1.Row + 1
                    FG1.SetFocus
                End If
            End If
     
       Case vbKeyBack
            FG1.Col = 5
            TextCantidadOriginal.Text = FG1.Text
            
            If Len(FG1) >= 1 Then
               FG1 = Left$(FG1, Len(FG1) - 1)
            Else
                KeyAscii = 0
            End If
           
       End Select
       
        
       codigoprod = ""
  
End Sub
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
        Dim i As Long
           
        If columnaSeis > MSFlexGrid3.Cols Then
           MsgBox "Columna no válida", vbExclamation
           Exit Function
        End If
          
        ' recorrer  las filas de la grilla
        For i = 1 To MSFlexGrid3.Rows - 1
            ' comprobar que el dato es de tipo numérico con la función IsNumeric de vb
            If IsNumeric(MSFlexGrid3.TextMatrix(i, columnaSeis)) Then
                ' Sumar, obteniendo el valor de la celda con TextMatrix
                totalLinea = totalLinea + MSFlexGrid3.TextMatrix(i, columnaSeis)
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
        Dim i As Long
           
        If columnaSiete > MSFlexGrid3.Cols Then
           MsgBox "Columna no válida", vbExclamation
           Exit Function
        End If
          
        ' recorrer  las filas de la grilla
        For i = 1 To MSFlexGrid3.Rows - 1
            ' comprobar que el dato es de tipo numérico con la función IsNumeric de vb
            If IsNumeric(MSFlexGrid3.TextMatrix(i, columnaSiete)) Then
                ' Sumar, obteniendo el valor de la celda con TextMatrix
                totalDescuento = totalDescuento + MSFlexGrid3.TextMatrix(i, columnaSiete)
            End If
        Next
           
        ' retornar el total de la suma a la función
       SumarTotalDescuentos = totalDescuento

    End With
    
Exit Function
error_function:
  
MsgBox Err.Description, vbCritical, "error al sumar"
                        
       
End Function

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


Private Sub busco()

    'MSFlexGrid2.Visible = False
    
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
            MSFlexGrid1.Text = rstCliente.Fields!IDCliente
            MSFlexGrid1.Col = 1
            MSFlexGrid1.Text = rstCliente.Fields!RazonSocial
            MSFlexGrid1.Col = 2
            MSFlexGrid1.Text = rstCliente.Fields!Cuit
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
Private Sub titulosfactura()

    MSFlexGrid2.Row = 0
    
    MSFlexGrid2.Col = 0
    MSFlexGrid2.Text = "Nº Factura"
    MSFlexGrid2.ColWidth(0) = 1100
    MSFlexGrid2.ColAlignment(0) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 1
    MSFlexGrid2.Text = "Tipo"
    MSFlexGrid2.ColWidth(1) = 900
    MSFlexGrid2.ColAlignment(1) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 2
    MSFlexGrid2.Text = "Fecha"
    MSFlexGrid2.ColWidth(2) = 1100
    MSFlexGrid2.ColAlignment(2) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 3
    MSFlexGrid2.Text = "Importe"
    MSFlexGrid2.ColWidth(3) = 1100
    MSFlexGrid2.ColAlignment(3) = flexAlignCenterCenter
    
       
        
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
    
    
    
    Call buscocuil
    
    
    
    MSFlexGrid1.Visible = False
    
    FG1.Enabled = True
    
    

End Sub



Private Sub buscofacturapaga()

    MSFlexGrid2.Clear
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstFacturaC = db.OpenRecordset("FacturaC", dbOpenDynaset)
        
    
    CodigoClie = Val(TextCodigoCliente.Text)
    
    rstFacturaC.FindFirst "CodCliente= " + Str(CodigoClie)
    facturacancelada = rstFacturaC.Fields!Cancelada
    codigoclientedetalle = rstFacturaC.Fields!CodCliente
    
    If rstFacturaC.Fields!CodCliente <> Val(TextCodigoCliente.Text) Then
        MSFlexGrid2.Visible = False
        mensaje = MsgBox("No Existen Facturas", vbCritical, "Final de la busqueda")
        TextCodigoCliente.Text = ""
        Call blanco
        TextCodigoCliente.SetFocus
    End If
    
   'If codigoclientedetalle = CodigoClie And facturacancelada = True Then
        MSFlexGrid2.Rows = 2
        MSFlexGrid2.Clear
        MSFlexGrid2.Visible = True
        Call titulosfactura
    'Else
    '    MSFlexGrid2.Visible = False
    'End If
     
    linea2 = 1
    Do While Not rstFacturaC.NoMatch
        If rstFacturaC.Fields!Cancelada = True Then
            MSFlexGrid2.AddItem " "
            MSFlexGrid2.Row = linea2
       
            MSFlexGrid2.Col = 0
            MSFlexGrid2.Text = rstFacturaC.Fields!NroFactura
            MSFlexGrid2.Col = 1
            MSFlexGrid2.Text = rstFacturaC.Fields!TipoFactura
            MSFlexGrid2.Col = 2
            MSFlexGrid2.Text = rstFacturaC.Fields!FechaFactura
            MSFlexGrid2.Col = 3
            MSFlexGrid2.Text = rstFacturaC.Fields!TotalFactura
             facturacancelada = rstFacturaC.Fields!Cancelada
            If facturacancelada = True Then
                MSFlexGrid2.Col = 4
                MSFlexGrid2.Text = "SI"
            End If
            linea2 = linea2 + 1
        End If
        rstFacturaC.FindNext "CodCliente= " + Str(CodigoClie)
    Loop
    
    FG1.Enabled = True
    
End Sub


Private Sub MSFlexGrid2_DblClick()
    
    linea2 = 0
    FG1.Clear
    
    Call SeteoGrilla
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstFacturaC = db.OpenRecordset("FacturaC", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstFacturaD = db.OpenRecordset("FacturaD", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    
    MSFlexGrid2.Col = 0
    NumeroFactura = Val(MSFlexGrid2.Text)
         
    rstFacturaC.FindFirst "NroFactura= " + Str(NumeroFactura)
    
    TextNumeroFactura.Text = rstFacturaC.Fields!NroFactura
    TextTipoFactura.Text = rstFacturaC.Fields!TipoFactura
    TextFechaFactura.Text = rstFacturaC.Fields!FechaFactura
    
    '*** buscar vendedor
            
    codigovendedor = Val(rstFacturaC.Fields!codVendedor)
         
    rstEmpleado.FindFirst "Legajo >= '" & codigovendedor & "'"
    ComboVendedor.Text = rstEmpleado.Fields!nombre

    '****
    
    TextSubtotalFactura.Text = rstFacturaC.Fields!SubTotalFactura
    TextDescuentos.Text = rstFacturaC.Fields!ImporteDesc
    TextPercepcionIIBB.Text = rstFacturaC.Fields!ImportePercepIIBB
    TextAlicuota.Text = rstFacturaC.Fields!AlicuotaIIBB
    TextImpuesto.Text = rstFacturaC.Fields!TotalIVA
    Textiva.Text = rstFacturaC.Fields!PorcentajeIVA
    TextTotalFactura.Text = rstFacturaC.Fields!TotalFactura
    
    rstFacturaD.FindFirst "NroFactura= " + Str(NumeroFactura)
    linea2 = 1
    Do While Not rstFacturaD.NoMatch
        
            FG1.AddItem " "
            FG1.Row = linea2
       
            FG1.Col = 0
            FG1.Text = rstFacturaD.Fields!IdCodProd
            
            FG1.Col = 0
            codigoprod = FG1.Text

            Dim busca1 As String, busca2 As String
            busca1 = RTrim(LTrim(codigoprod))
            busca2 = busca1 + "z"
       
            rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
            FG1.Col = 1
            FG1.Text = rstProductos.Fields!Descripcion
        
            FG1.Col = 2
            FG1.Text = rstFacturaD.Fields!UnidadMedida
            FG1.Col = 3
            FG1.Text = rstFacturaD.Fields!precioUnitario
            FG1.Col = 4
            FG1.Text = rstFacturaD.Fields!PorcentajeDescuento
            FG1.Col = 5
            FG1.Text = rstFacturaD.Fields!cantidad
            FG1.Col = 6
            FG1.Text = rstFacturaD.Fields!totalLinea
           
       
           rstFacturaD.FindNext "NroFactura= " + Str(NumeroFactura)
           linea2 = linea2 + 1
    Loop
       
    
    MSFlexGrid2.Visible = False

End Sub

Private Sub OptionFacturaImpaga_Click()

    Call buscofacturaimpaga

End Sub
Private Sub buscofacturaimpaga()

    MSFlexGrid2.Clear
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstFacturaC = db.OpenRecordset("FacturaC", dbOpenDynaset)
      
    CodigoClie = Val(TextCodigoCliente.Text)
    
    rstFacturaC.FindFirst "CodCliente= " + Str(CodigoClie)
    facturacancelada = rstFacturaC.Fields!Cancelada
    codigoclientedetalle = rstFacturaC.Fields!CodCliente
    
    If rstFacturaC.Fields!CodCliente <> Val(TextCodigoCliente.Text) Then
        MSFlexGrid2.Visible = False
        mensaje = MsgBox("No Existen Facturas", vbCritical, "Final de la busqueda")
        TextCodigoCliente.Text = ""
        Call blanco
        TextCodigoCliente.SetFocus
        
    End If
    
    'If codigoclientedetalle = CodigoClie And facturacancelada = True Then
        MSFlexGrid2.Rows = 2
        MSFlexGrid2.Clear
        MSFlexGrid2.Visible = True
        Call titulosfactura
    'Else
    '    MSFlexGrid2.Visible = False
    'End If
    
    linea2 = 1
    Do While Not rstFacturaC.NoMatch
        If rstFacturaC.Fields!Cancelada = False Then
            MSFlexGrid2.AddItem " "
            MSFlexGrid2.Row = linea2
       
            MSFlexGrid2.Col = 0
            MSFlexGrid2.Text = rstFacturaC.Fields!NroFactura
            MSFlexGrid2.Col = 1
            MSFlexGrid2.Text = rstFacturaC.Fields!TipoFactura
            MSFlexGrid2.Col = 2
            MSFlexGrid2.Text = rstFacturaC.Fields!FechaFactura
            MSFlexGrid2.Col = 3
            MSFlexGrid2.Text = rstFacturaC.Fields!TotalFactura
            'facturacancelada = rstFacturaC.Fields!Cancelada
            'If facturacancelada = False Then
            '    MSFlexGrid2.Col = 4
            '    MSFlexGrid2.Text = "NO"
            'End If
            linea2 = linea2 + 1
        End If
        rstFacturaC.FindNext "CodCliente= " + Str(CodigoClie)
    Loop
    
    FG1.Enabled = True
End Sub


Private Sub OptionFacturaPaga_Click()

    Call buscofacturapaga

End Sub

Private Sub OptionFacturaTodas_Click()

    buscofacturatodas

End Sub
Private Sub buscofacturatodas()

     MSFlexGrid2.Clear
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstFacturaC = db.OpenRecordset("FacturaC", dbOpenDynaset)
    
    
    MSFlexGrid2.Rows = 2
    MSFlexGrid2.Clear
    MSFlexGrid2.Visible = True
    
    Call titulosfactura
    
    
    CodigoClie = Val(TextCodigoCliente.Text)
    
    rstFacturaC.FindFirst "CodCliente= " + Str(CodigoClie)
    'facturacancelada = rstFacturaC.Fields!Cancelada
    codigoclientedetalle = rstFacturaC.Fields!CodCliente
    
    If rstFacturaC.Fields!CodCliente <> Val(TextCodigoCliente.Text) Then
        MSFlexGrid2.Visible = False
        mensaje = MsgBox("No Existen Facturas", vbCritical, "Final de la busqueda")
        TextCodigoCliente.Text = ""
        Call blanco
        TextCodigoCliente.SetFocus
    End If
    
    If codigoclientedetalle = CodigoClie Then
        MSFlexGrid2.Rows = 2
        MSFlexGrid2.Clear
        MSFlexGrid2.Visible = True
    
        Call titulosfactura
    Else
        MSFlexGrid2.Visible = False
    End If
     
    linea2 = 1
    Do While Not rstFacturaC.NoMatch
            MSFlexGrid2.AddItem " "
            MSFlexGrid2.Row = linea2
       
            MSFlexGrid2.Col = 0
            MSFlexGrid2.Text = rstFacturaC.Fields!NroFactura
            MSFlexGrid2.Col = 1
            MSFlexGrid2.Text = rstFacturaC.Fields!TipoFactura
            MSFlexGrid2.Col = 2
            MSFlexGrid2.Text = rstFacturaC.Fields!FechaFactura
            MSFlexGrid2.Col = 3
            MSFlexGrid2.Text = rstFacturaC.Fields!TotalFactura
            'facturacancelada = rstFacturaC.Fields!Cancelada
            'If facturacancelada = True Then
            '    MSFlexGrid2.Col = 4
            '    MSFlexGrid2.Text = "SI"
            'Else
            '    MSFlexGrid2.Col = 4
            '    MSFlexGrid2.Text = "NO"
            'End If
            linea2 = linea2 + 1
                
            rstFacturaC.FindNext "CodCliente= " + Str(CodigoClie)
    Loop
    
    FG1.Enabled = True

End Sub

Private Sub TextApellidoNombre_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        MSFlexGrid2.Visible = False
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

Private Sub TextCodigoCliente_GotFocus()
    TextCodigoCliente.SelLength = Len(TextCodigoCliente.Text)
End Sub

Private Sub TextCodigoCliente_KeyPress(KeyAscii As Integer)
   
    MSFlexGrid2.Visible = False
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
   

    If KeyAscii = 13 Then
        If TextCodigoCliente.Text = "" Then
            TextApellidoNombre.SetFocus
        Else
            CodigoClie = Val(TextCodigoCliente.Text)
      
            rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCliente.Fields!IDCliente <> Val(TextCodigoCliente.Text) Then
                mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                TextCodigoCliente.Text = ""
                Call blanco
                TextCodigoCliente.SetFocus
            Else
                TextCodigoCliente.Text = rstCliente.Fields!IDCliente
                TextApellidoNombre.Text = rstCliente.Fields!RazonSocial
                TextCuit.Text = rstCliente.Fields!Cuit
                TextDireccion.Text = rstCliente.Fields!Domicilio
                TextLocalidad.Text = rstCliente.Fields!Localidad
                TextCodigoPostal.Text = rstCliente.Fields!CP
                TextProvincia.Text = rstCliente.Fields!Prov
                TextDescuentoCliente.Text = rstCliente.Fields!PorcentajeDescuento
                Call buscocuil
            End If
        End If
         Call buscofacturatodas
    End If
    
    If TextNumeroFactura <> "" Then
        FG1.Enabled = True
    Else
        FG1.Enabled = False
    End If
    
   
End Sub
Private Sub buscocuil()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
    CodigoClie = Val(TextCodigoCliente.Text)
      
    rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
    
    TextCuit.Text = rstCliente.Fields!Cuit
   
    
    Set rstPadron = db.OpenRecordset("Padron", dbOpenTable)
    
    rstPadron.Index = "CUIT"
    
    With rstPadron
        rstPadron.Seek "=", TextCuit.Text
        If .NoMatch = False Then
            Alicuota = !AlicuotaPercepcion
        End If
    End With
    
    
  
End Sub

Private Sub Form_Load()

    FormBusquedaFacturaPorCliente.Height = 8970
    FormBusquedaFacturaPorCliente.Width = 12135
    FormBusquedaFacturaPorCliente.Top = 1000
    FormBusquedaFacturaPorCliente.Left = 1000
    
    Call SeteoGrilla
      
    Call Cargo
    

End Sub
