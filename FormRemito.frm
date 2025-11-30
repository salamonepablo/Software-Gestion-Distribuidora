VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormRemito 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Remitos"
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   12135
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
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
      TabIndex        =   44
      Top             =   3120
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
         Left            =   7800
         TabIndex        =   6
         Top             =   600
         Width           =   735
      End
      Begin VB.CommandButton BotonAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   10560
         TabIndex        =   7
         Top             =   600
         Width           =   975
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
         Left            =   3120
         TabIndex        =   4
         Top             =   600
         Width           =   3975
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
         Left            =   7200
         TabIndex        =   5
         Top             =   600
         Width           =   495
      End
      Begin VB.ComboBox ComboArticulo 
         Height          =   315
         Left            =   1800
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   7800
         TabIndex        =   48
         Top             =   360
         Width           =   765
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   2040
         TabIndex        =   47
         Top             =   360
         Width           =   660
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   4560
         TabIndex        =   46
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   7320
         TabIndex        =   45
         Top             =   360
         Width           =   300
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Height          =   975
      Left            =   120
      TabIndex        =   35
      Top             =   2160
      Width           =   11655
      Begin VB.ComboBox cmbSucursal 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   2295
      End
      Begin VB.TextBox TextNumeroRemito 
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
         Left            =   2880
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextFechaRemito 
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
         Left            =   4680
         TabIndex        =   39
         Top             =   480
         Width           =   1335
      End
      Begin VB.ComboBox ComboVendedor 
         Height          =   315
         Left            =   6360
         TabIndex        =   38
         Top             =   480
         Width           =   1575
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
         Left            =   8280
         TabIndex        =   37
         Top             =   480
         Width           =   1335
      End
      Begin VB.CheckBox CheckModificaStock 
         BackColor       =   &H00C0C0C0&
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
         TabIndex        =   36
         Top             =   480
         Value           =   1  'Checked
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Sucursal"
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
         TabIndex        =   49
         Top             =   240
         Width           =   750
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Nº Remito"
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
         TabIndex        =   43
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Fecha Remito"
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
         TabIndex        =   42
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   6360
         TabIndex        =   41
         Top             =   240
         Width           =   825
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Left            =   8640
         TabIndex        =   40
         Top             =   240
         Width           =   495
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   4200
      TabIndex        =   31
      Top             =   720
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton BotonDomicilio 
         BackColor       =   &H00808000&
         Caption         =   "&Domicilio Entrega"
         Enabled         =   0   'False
         Height          =   510
         Left            =   9960
         MaskColor       =   &H00808000&
         TabIndex        =   33
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox TextItemDomicilio 
         Height          =   285
         Left            =   4200
         TabIndex        =   32
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox TextDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   22
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
         TabIndex        =   9
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   14
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   12
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TextProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   11
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Top             =   720
         Width           =   510
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Top             =   720
         Width           =   870
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Top             =   1080
         Width           =   900
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
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
         Top             =   1080
         Width           =   870
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   1095
      Left            =   120
      TabIndex        =   24
      Top             =   8400
      Width           =   11655
      Begin VB.CommandButton BotonGrabar 
         BackColor       =   &H00808000&
         Caption         =   "&Guardar"
         Enabled         =   0   'False
         Height          =   750
         Left            =   1560
         MaskColor       =   &H00808000&
         TabIndex        =   30
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonSalir 
         BackColor       =   &H00808000&
         Caption         =   "&Salir"
         Height          =   750
         Left            =   4080
         TabIndex        =   28
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonCancelar 
         BackColor       =   &H00808000&
         Caption         =   "&Cancelar"
         Height          =   750
         Left            =   3240
         TabIndex        =   27
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonImprimir 
         BackColor       =   &H00808000&
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   750
         Left            =   2400
         TabIndex        =   26
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
         TabIndex        =   25
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   4095
      Left            =   120
      TabIndex        =   23
      Top             =   4320
      Width           =   11655
      Begin VB.CommandButton BotonEliminarfila 
         Caption         =   "&Eliminar Fila"
         Height          =   495
         Left            =   9600
         TabIndex        =   34
         Top             =   1680
         Width           =   735
      End
      Begin VB.CommandButton BotonBuscarProducto 
         Caption         =   "&Buscar"
         Height          =   495
         Left            =   10800
         MaskColor       =   &H00808000&
         TabIndex        =   29
         Top             =   2760
         Visible         =   0   'False
         Width           =   735
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   3615
         Left            =   1320
         TabIndex        =   8
         Top             =   360
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   6376
         _Version        =   393216
         Rows            =   16
         Cols            =   5
         FixedCols       =   0
         Enabled         =   0   'False
         GridLines       =   2
      End
   End
End
Attribute VB_Name = "FormRemito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstEmpleado As DAO.Recordset
 Dim rstCliente As DAO.Recordset
 Dim rstProductos As DAO.Recordset
 Dim rstRemitoC As DAO.Recordset
 Dim rstRemitoD As DAO.Recordset
 Dim rstPadron As DAO.Recordset
 Dim rstUltimosNumeros As DAO.Recordset
 Dim cantidadProducto As Integer
 Dim vendedorCliente As String
 Dim nombreVendedor As Integer
 Dim LegajoEmpleado As String
 Dim modificaStock As Integer
 Dim Fila As Integer
 Dim fila2 As Integer
 Dim renglon As Integer
 
Private Sub ImprimirRemito()

   Dim RemC
   Dim RemD
   Dim NumeroRemito As Long
   Dim DescripcionCondIVA As String
   Dim Aclaracion As String
        
    'On Error GoTo CapturaErrores

   Aclaracion = "Nota: "
   x = 0
   Y = 0
   renglon = 0
   vNroRemito = "0002- " & TextNumeroRemito.text
   NumeroRemito = CLng(TextNumeroRemito.text)
   DescripcionCondIVA = DescCondIVA(vCondIVA)
    
   vSQLRc = "SELECT * FROM RemitoC WHERE IdSucursal=" & IdSucursal & " AND NroRemito=" & NumeroRemito & " ORDER By NroRemito"
   vSQLRd = "SELECT * FROM RemitoD WHERE IdSucursal=" & IdSucursal & " AND NroRemito=" & NumeroRemito & " ORDER By NroRemito, ItemRemito"
   'vSQLRdir = "SELECT * FROM RemitoD WHERE NroRemito=" & CLng(TextNumeroRemito.text) & " ORDER By NroRemito"
    
  ' MsgBox (vSQLRc)
  ' MsgBox (vSQLRd)
  ' MsgBox (vSQLRdir)
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    
    Set RemC = BaseSPC.OpenRecordset(vSQLRc, dbOpenDynaset)
    Set RemD = BaseSPC.OpenRecordset(vSQLRd, dbOpenDynaset)
      
        
    'With p
        'Seteo escala a mm
            'Set Printer = Printer
            Printer.Copies = 3
            'Printer.Copies = 1
            Printer.ScaleMode = 6
        
        'Imprimir Fecha
            Printer.CurrentX = x + 163
            Printer.CurrentY = Y + 17
            Printer.Font = "Arial"
            Printer.FontSize = 12
            Printer.FontBold = True
            'Printer.Print Format(FormFactura.TextFechaFactura.text, "DD/MM/YYYY")
            Printer.Print Format(TextFechaRemito.text, "DD/MM/YYYY")
        
        'Imprimir Nombres
           Printer.CurrentX = x + 38
            Printer.CurrentY = Y + 49
            Printer.Font = "Arial"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print TextApellidoNombre.text
            
        'Imprimir Direccion
            Printer.CurrentX = x + 38
            Printer.CurrentY = Y + 56
            Printer.Font = "Arial"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print TextDireccion.text
            
        'Imprimir Localidad
            Printer.CurrentX = x + 38
            Printer.CurrentY = Y + 63
            Printer.Font = "Arial"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print TextLocalidad.text
            
        'Imprimir CUIT
            Printer.CurrentX = x + 148
            Printer.CurrentY = Y + 48
            Printer.Font = "Arial"
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.Print TextCuit.text
            
        'Imprimir Marca Responsable Inscripto
            Printer.CurrentX = x + 38
            Printer.CurrentY = Y + 70
            Printer.Font = "Arial"
            Printer.FontSize = 10
            Printer.FontBold = False
            'Printer.Print "X"
            Printer.Print vCondIVA & " - " & DescripcionCondIVA
            
        'Imprimir Marca Contado
         '   Printer.CurrentX = X + 70
         '   Printer.CurrentY = Y + 80
         '   Printer.Font = "Courier New"
         '   Printer.FontSize = 10
         '   Printer.FontBold = False
         '   Printer.Print "X"
            
        'Imprimir Marca CtaCte
         '   Printer.CurrentX = X + 100
         '   Printer.CurrentY = Y + 80
         '   Printer.Font = "Courier New"
         '   Printer.FontSize = 10
         '   Printer.FontBold = False
         '   Printer.Print "X"
            
        'Imprimir Nro Remito
          '  Printer.CurrentX = x + 138
          '  Printer.CurrentY = Y + 80
          '  Printer.Font = "Courier New"
          '  Printer.FontSize = 9
          '  Printer.FontBold = False
          '  Printer.Print vNroRemito
            
        'Imprimir Detalle
          '  vsqlfc = "SELECT * FROM FacturaC WHERE TipoFactura='" & TextTipoFactura.text & "' AND NroFactura=" & TextNumeroFactura.text & " ORDER By NroFactura"
          '  vsqlFD = "SELECT * FROM FacturaD WHERE TipoFactura='" & TextTipoFactura.text & "' AND NroFactura=" & TextNumeroFactura.text & " ORDER By NroFactura"
            
          '  Set RemC = BaseSPC.OpenRecordset(vsqlfc, dbOpenDynaset)
          '  Set RemD = BaseSPC.OpenRecordset(vsqlFD, dbOpenDynaset)
            
           
            RemC.MoveFirst
            RemD.MoveFirst
            
          'Guardo Aclaración
            If RemC!AclaracionRemito <> "" Then Aclaracion = Aclaracion & RemC!AclaracionRemito
                
                    While Not RemD.EOF
                        'Imprimo el detalle
                            Printer.CurrentX = x + 26
                            Printer.CurrentY = Y + 100 + renglon
                            Printer.Font = "Courier New"
                            Printer.FontSize = 10
                            Printer.FontBold = False
                            Printer.Print Format(RemD!cantidad, "#,###,###,#0.00")
                            
                        'Detalle
                            Printer.CurrentX = x + 38
                            Printer.CurrentY = Y + 100 + renglon
                            Printer.Font = "Courier New"
                            Printer.FontSize = 10
                            Printer.FontBold = False
                            'Printer.Print RemD!IdCodProd & Chr(9) & Descripcion(RemD!IdCodProd)
                            Printer.Print Chr(9) & Descripcion(RemD!IdCodProd)
                        
                         renglon = renglon + 5
                            
                        RemD.MoveNext
                    Wend
        
                        'Aclaración
                            Printer.CurrentX = x + 38
                            Printer.CurrentY = Y + 220 + renglon
                            Printer.Font = "Courier New"
                            Printer.FontSize = 10
                            Printer.FontBold = False
                            'Printer.Print RemD!IdCodProd & Chr(9) & Descripcion(RemD!IdCodProd)
                            Printer.Print Chr(9) & Aclaracion
        
        
        Printer.EndDoc
        
'    End With
    
    RemC.Close
    RemD.Close
    
    BaseSPC.Close
        
CapturaErrores:
    'If Err = 321 Then
    'End If
'    MsgBox (Err & " " & Err.Description)
'    End
    
End Sub

Public Function Descripcion(IdCodProd As Variant) As String

    Set tProductos = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    
    tProductos.Index = "PrimaryKey"
    
    tProductos.Seek "=", IdCodProd

    If Not tProductos.NoMatch Then Descripcion = tProductos!Descripcion

End Function
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
        FG1.text = TextCantidad.text
                
        Fila = Fila + 1
        fila2 = fila2 + 1
            
    
        
        
'        ComboArticulo.Text = ""
        TextDescripcion.text = ""
        TextUnidadMedida.text = ""
        TextCantidad.text = ""
        
        
        ComboArticulo.SetFocus
    
  End If
        BotonGrabar.Enabled = True
        BotonImprimir.Enabled = True
End Sub

Private Sub BotonBuscarProducto_Click()

    FormBusquedaProductosRemito.Show

End Sub

Private Sub BotonCancelar_Click()

    Call blanqueototal
    
End Sub

Private Sub blanqueototal()

    TextCodigoCliente.text = ""
    TextApellidoNombre.text = ""
    TextCuit.text = ""
    TextDireccion.text = ""
    TextLocalidad.text = ""
    TextCodigoPostal.text = ""
    TextProvincia.text = ""
    TextNumeroRemito.text = ""
    ComboVendedor.text = ""
    TextSaldoCliente.text = ""
    FG1.Clear
    FG1.Enabled = False
    
    MSFlexGrid1.Visible = False
 
    Call SeteoGrilla

End Sub

Private Sub BotonCancelar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonDomicilio_Click()

    FormDomiciliosRemito.Show
   
    
End Sub

Private Sub BotonEliminarfila_Click()

    If FG1.Row <= 0 Then
        MsgBox "Debe Seleccionar una fila"
    'ElseIf MSFlexGrid1.Row = 1 Then
    ' MSFlexGrid1.Clear
    Else
        FG1.RemoveItem (FG1.Row)
    End If
    
End Sub

Private Sub BotonGrabar_Click()

        Dim descuentoCantidad As Long
        Dim ultimo As Long
        Dim existeNumeroBD As Long
        Dim existeTipoBD As String
        Dim existeNumero As Long
        Dim existeTipo As String
        
        ruta = App.Path & "\DB_SPC_SI.mdb"
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstRemitoC = db.OpenRecordset("RemitoC", dbOpenDynaset)
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstRemitoD = db.OpenRecordset("RemitoD", dbOpenDynaset)
        
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
        
                
        existeNumero = TextNumeroRemito.text
        
        'buscamos el deposito para descontar el stock
        
        Set tDepositos = db.OpenRecordset("Depositos", dbOpenTable)
       
'        On Error GoTo CapturaErrores
   
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
    
       
        
        '*** Busco Numero y Tipo Factura Existentes
        
        'Dim busca5 As String, busca6 As String
        'busca5 = RTrim(LTrim(TextTipoFactura.Text))
        'busca6 = busca5 + "z"
            
               
        'existeNumero = Val(TextNumeroRemito.text)
      
        'rstRemitoC.FindFirst "NroFactura= " + Str(existeNumero) And "TipoFactura >= '" & existeTipo & "'"
        'rstRemitoC.FindFirst "NroFactura= " + Str(existeNumero) And "TipoFactura = '" & existeTipo & "'"
        'rstRemitoC.FindFirst "NroFactura= " + Str(existeNumero)
        'rstRemitoC.FindFirst "TipoFactura >= '" & busca5 & "' and TipoFactura <= '" & busca6 & "'"
       
        'existeNumeroBD = rstRemitoC.Fields!NroFactura
        'existeTipoBD = rstRemitoC.Fields!TipoFactura
     
        'If existeNumero = existeNumeroBD And existeTipo = existeTipoBD Then
        '    mensaje = MsgBox("Numero y Tipo de Factura Existentes", vbCritical, "Final de la busqueda")
        '    TextNumeroRemito.text = ""
        '    TextNumeroPresupuesto.SetFocus
        'else
            rstRemitoC.AddNew
                rstRemitoC.Fields!IdSucursal = CLng(Left(cmbSucursal.text, 1))
                IdSucursal = rstRemitoC.Fields!IdSucursal
                rstRemitoC.Fields!NroRemito = CLng(TextNumeroRemito.text)
                rstRemitoC.Fields!FechaRemito = TextFechaRemito.text
                rstRemitoC.Fields!CodCliente = TextCodigoCliente.text
                rstRemitoC.Fields!codVendedor = LegajoEmpleado
                'rstRemitoC.Fields!ItemDomicilio = Val(TextItemDomicilio.Text)
            
               '//Solicitud de aclaracion en el remito 2025-07-02
                Dim Rta As Long
                   Rta = MsgBox("¿Desea Agregar una Aclaración en el Remito?", vbYesNo, "Aclaración")
            
                    If Rta = vbYes Then
                       Dim Aclaracion As String
                       Aclaracion = InputBox("Ingrese Aclaración", "SPCSI Observaciones")
                       ' Validar que no se haya cancelado el InputBox
                       If Aclaracion <> "" Then
                           rstRemitoC.Fields!AclaracionRemito = Aclaracion
                       Else
                           rstRemitoC.Fields!AclaracionRemito = " "
                       End If
                    Else
                       'rstFacturaC.Fields!AclaracionRemito = " "
                   End If
              '//********************************************************************
            
            rstRemitoC.Update
            
            FG1.Col = 0
            FG1.Row = 1
            Filas = FG1.Rows
            linea = 1
            Do While linea < Filas
                  
                  FG1.Row = linea
                  FG1.Col = 0
                  If FG1.text <> "" Then
                        rstRemitoD.AddNew
                    
                        rstRemitoD.Fields!IdSucursal = CLng(Left(cmbSucursal.text, 1))
                        rstRemitoD.Fields!NroRemito = CLng(TextNumeroRemito.text)
                        'rstRemitoD.Fields!TipoFactura = TextTipoFactura.Text
                    
                        FG1.Col = 0
                        rstRemitoD.Fields!IdCodProd = FG1.text
                    
                        FG1.Col = 2
                        rstRemitoD.Fields!UnidadMedida = FG1.text
                        
                        FG1.Col = 3
                        rstRemitoD.Fields!cantidad = Val(FG1.text)
                        descuentoCantidad = Val(FG1.text)
                        
                        FG1.Col = 4
                        rstRemitoD.Fields!itemremito = Val(FG1.text)
                        
                        
                        
                        
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
                        
                       rstRemitoD.Update
                  End If
                  linea = linea + 1
            Loop
            
            
            '*** Actualizo Ultimo Numero Remito
            
            Set db = DBEngine.OpenDatabase(ruta)
            Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
        
            Dim busco As String
       
            'If TextTipoFactura.Text = "A" Then
                busco = "tRemitoC"
            'End If
            
            'If TextTipoFactura.Text = "B" Then
            '    busco = "tFacturaB"
            'End If
    
            'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
            rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
            ultimo = rstUltimosNumeros.Fields!UltimoNumero
            
            If ultimo < Val(TextNumeroRemito.text) Then
                rstUltimosNumeros.Edit
                'If ultimo < rstUltimosNumeros.Fields!UltimoNumero Then
                     rstUltimosNumeros.Fields!UltimoNumero = TextNumeroRemito.text
                'End If
                rstUltimosNumeros.Update
            End If
        ' End If
        
        
        
        BotonGrabar.Enabled = False
        BotonNueva.Enabled = False
        BotonImprimir.Enabled = False
                        
        Rta = MsgBox("¿Desea Imprimir el remito?", vbYesNo, "INFO DEL SISTEMA")
        
        If Rta = vbYes Then
            Call ImprimirRemito
        End If
        
        modificaStock = 0
        
        Call blanqueototal
        Call SeteoGrilla
        
        TextCodigoCliente.SetFocus
        
CapturaErrores:
        
        Select Case Err
            Case 3021
                Resume Next
        End Select
        fila2 = 0
        Fila = 0
        
    
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

    'FormImprimeRemito.Show

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
 
    Dim busco As String
       
    busco = "tRemitoC"
        
    'rstUltimosNumeros.FindFirst "IDTabla >= '" & busca1 & "' and IDTabla <= '" & busca2 & "'"
    rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
    NumeroPresupuesto = rstUltimosNumeros.Fields!UltimoNumero
    
    'If rstUltimosNumeros.NoMatch Then
    '   FG1.Visible = False
    '   mensaje = MsgBox("No existen Numeros de Factura", vbCritical, "Final de la busqueda")
    'End If
    
    TextNumeroRemito.text = NumeroPresupuesto + 1

    If TextCuit.text <> "" Then
       FG1.Enabled = True
    End If
    
    BotonNueva.Enabled = False
    
    FG1.Enabled = True
    FG1.Row = 1
    FG1.Col = 0
  
  ' *** AGREGADO POR PVS ***
    '  FG1.SetFocus
         TextNumeroRemito.SetFocus
    
End Sub



Private Sub BotonNueva_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub BotonSalir_Click()

    Unload FormRemito

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


Private Sub ComboVendedor_GotFocus()
    
    ComboVendedor.SelLength = Len(ComboVendedor.text)

End Sub

Private Sub ComboVendedor_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub FG1_KeyPress(KeyAscii As Integer)

 '**** PUSE ESTE IF PARA NO TENER QUE COMENTAR LINEA POR LINEA *** 'PVS' ****
 If PVS = 1 Then
    
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
    Dim bandera As Integer
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
     
    'Set db = DBEngine.OpenDatabase(ruta)
    'Set rstIva = db.OpenRecordset("Iva", dbOpenDynaset)

    'iva = rstIva.Fields!iva
    
    'If TextTipoFactura.Text = "A" Then
    '    Textiva.Text = Format(iva, "#,###,###,#0.00")
    'End If
        
    If KeyAscii >= 32 And KeyAscii <= 127 Then
        FG1.text = FG1.text & Chr(KeyAscii)
    End If

    Select Case KeyAscii
       Case 13
            
            FG1.Col = 0
            codigoprodMA = UCase(FG1.text)
                   
            '******* Busco Producto
            
           Set tProductos = db.OpenRecordset("Productos", dbOpenTable)
        
   
           tProductos.Index = "PrimaryKey"
           
           tProductos.MoveFirst
           tProductos.Seek "=", codigoprodMA
           
           If Not tProductos.NoMatch Then
                produ = tProductos!CodProd
                'MsgBox (DepoOrigen)
                 Dim busca1 As String, busca2 As String
                 busca1 = RTrim(LTrim(codigoprodMA))
                 busca2 = busca1 + "z"
                                     
                 rstProductos.FindFirst "CodProd >= '" & busca1 & "' and CodProd <= '" & busca2 & "'"
            
                 codigoProdTabla = rstProductos.Fields!CodProd
            
                cantidadProducto = rstProductos.Fields!Stock
                FG1.Col = 1
                FG1.text = rstProductos.Fields!Descripcion
                FG1.Col = 2
                FG1.text = rstProductos.Fields!UnidadMedida
                FG1.Col = 3
               
    
           Else
                 mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                 codigoprodMA = ""
                 FG1.Col = 1
                 FG1.text = ""
                 FG1.Col = 2
                 FG1.text = ""
                 FG1.Col = 3
                 FG1.text = ""
                 FG1.Col = 4
                 FG1.text = ""
                 FG1.Col = 0
                 FG1.text = ""
                 FG1.SetFocus
                 bandera = 1
           End If
              
           tProductos.Close
             
            
            
            '***********************
                   
            If FG1.Col = 3 And FG1.text <> "" Then
                FG1.Col = 0
                'If FG1.Row < 2 Then
                    FG1.Row = FG1.Row + 1
                    FG1.SetFocus
                    BotonGrabar.Enabled = True
                    BotonImprimir.Enabled = True
                'End If
            End If
     
             
       Case vbKeyBack
            
            If Len(FG1) >= 1 Then
               FG1 = Left$(FG1, Len(FG1) - 1)
            Else
                KeyAscii = 0
            End If
           
       End Select
       
        
       codigoprod = ""
  
  End If

End Sub
Private Sub muestrodatosproductos()

    cantidadProducto = rstProductos.Fields!Stock
    FG1.Col = 1
    FG1.text = rstProductos.Fields!Descripcion
    FG1.Col = 2
    FG1.text = rstProductos.Fields!UnidadMedida
    
           
End Sub


Private Sub Form_Load()

   FormRemito.Height = 10110
   FormRemito.Width = 12135
   FormRemito.Top = 600
   FormRemito.Left = 50
   
    Fila = 1
    renglon = 16
    
    Call SeteoGrilla
    
    
    Call Cargo
    Call buscoarticulo
    
    TextFechaRemito.text = Format(Date, "dd/mm/yyyy")
    
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

Private Sub ComboArticulo_Click()

    
    TextCantidad.text = ""
    
    
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
   
      
    Dim busca1 As String
    busca1 = RTrim(LTrim(ComboArticulo.text))
   
    
    rstProductos.FindFirst "CodProd >= '" & busca1 & "' "
    
   
    TextDescripcion.text = rstProductos.Fields!Descripcion
    TextUnidadMedida.text = rstProductos.Fields!UnidadMedida
    
      
   ' TextCantidad.SetFocus
    
    
    
End Sub
Private Sub ComboArticulo_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub
Private Sub ComboArticulo_LostFocus()
    
   
    TextCantidad.text = ""
   
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstProductos = db.OpenRecordset("Productos", dbOpenDynaset)
    
    'If KeyAscii = 13 Then
      
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
                 
                
           Else
                mensaje = MsgBox("Producto Inexistente", vbCritical, "Final de la busqueda")
                'ComboArticulo.text = ""
                TextDescripcion.text = ""
                ComboArticulo.SetFocus
           End If
              
           tProductos.Close
          'TextCantidad.SetFocus
    'End If

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
    FG1.ColWidth(3) = 900
    FG1.CellFontBold = True
    FG1.text = "Cantidad"
    FG1.ColAlignment(3) = flexAlignCenterCenter
    
    FG1.Col = 4
    FG1.ColWidth(4) = 0
    FG1.CellFontBold = True
    FG1.text = "Item"
    
    FG1.Row = 1
    item = 1
    linea = 1
    Do While FG1.Row <= 14
        FG1.Col = 4
        FG1.text = item
        item = (item + 1)
        FG1.Row = (Val(FG1.Row) + 1)
    Loop
        
    
    
End Sub

Private Sub Cargo()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstEmpleado = db.OpenRecordset("Empleados", dbOpenDynaset)
    Set tSucursales = db.OpenRecordset("Sucursales", dbOpenDynaset)
    
    rstEmpleado.MoveFirst
    Do While Not rstEmpleado.EOF
        ComboVendedor.AddItem rstEmpleado!Nombre
        rstEmpleado.MoveNext
    Loop
    
    tSucursales.MoveFirst
    While Not tSucursales.EOF
        cmbSucursal.AddItem tSucursales!IdSucursal & " - " & tSucursales!NombreSucursal
        tSucursales.MoveNext
    Wend
    cmbSucursal.ListIndex = 1

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
            If Not IsNull(rstCliente.Fields!CUIT) Then MSFlexGrid1.text = rstCliente.Fields!CUIT
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

Private Sub TextApellidoNombre_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call busco
    End If
    
    If KeyAscii = 27 Then
        Unload Me
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


Private Sub TextCantidad_GotFocus()
    TextCantidad.SelLength = Len(TextCantidad.text)
End Sub

Private Sub TextCantidad_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
End Sub

Private Sub TextCantidad_LostFocus()

    If TextCantidad.text = "" Then
        A = MsgBox("NO PUEDE DEJAR LA CANTIDAD EN BLANCO", vbCritical, "E R R O R ! ! !")
        TextCantidad.SetFocus
    End If

End Sub

Private Sub TextCodigoCliente_GotFocus()
    TextCodigoCliente.SelLength = Len(TextCodigoCliente.text)
End Sub

Private Sub TextCodigoCliente_KeyPress(KeyAscii As Integer)
   
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
                MSFlexGrid1.Visible = False
                TextCuit.text = rstCliente.Fields!CUIT
                TextDireccion.text = rstCliente.Fields!Domicilio
                TextLocalidad.text = rstCliente.Fields!localidad
                TextCodigoPostal.text = rstCliente.Fields!CP
                TextProvincia.text = rstCliente.Fields!Prov
                vendedorCliente = rstCliente.Fields!Vendedor
                Call buscocuilyvendedor
            End If
        End If
        TextNumeroRemito.text = ""
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
    
    CodigoClie = Val(TextCodigoCliente.text)
      
    rstCliente.FindFirst "IDCliente= " + Str(CodigoClie)
    
    TextCuit.text = rstCliente.Fields!CUIT
    codigovendedor = rstCliente!Vendedor
      
    
    '*** Busco Vendedor
    
    CodigoVend = codigovendedor
      
    rstEmpleado.FindFirst "Legajo >= '" & CodigoVend & "'"
    
    LegajoEmpleado = rstEmpleado.Fields!Legajo
    ComboVendedor.text = rstEmpleado.Fields!Nombre
    
    '*** Busco Saldo
    
   rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
    
   TextSaldoCliente.text = Format(rstCtaCte.Fields!SaldoTotal, "#,###,###,#0.00")
   
   BotonDomicilio.Enabled = True
   
   BotonNueva.Enabled = True
   BotonNueva.SetFocus
   
    
End Sub

Private Sub TextCodigoCliente_LostFocus()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
   

'    If KeyAscii = 13 Then
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
                TextCuit.text = rstCliente.Fields!CUIT
                TextDireccion.text = rstCliente.Fields!Domicilio
                TextLocalidad.text = rstCliente.Fields!localidad
                TextCodigoPostal.text = rstCliente.Fields!CP
                TextProvincia.text = rstCliente.Fields!Prov
                vendedorCliente = rstCliente.Fields!Vendedor
                vCondIVA = rstCliente.Fields!condicionIva
                Call buscocuilyvendedor
            End If
        End If
        TextNumeroRemito.text = ""
 '   End If
    
    If TextNumeroPresupuesto <> "" Then
        FG1.Enabled = True
    Else
        FG1.Enabled = False
    End If


End Sub

Private Sub TextCuit_Change()

    If TextCuit.text <> "" Then
        BotonNueva.Enabled = True
    End If
        
End Sub

Private Sub TextDescripcion_GotFocus()
    TextDescripcion.SelLength = Len(TextDescripcion.text)
End Sub

Private Sub TextDescripcion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextFechaRemito_GotFocus()
    TextFechaRemito.SelLength = Len(TextFechaRemito.text)
End Sub

Private Sub TextFechaRemito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextItemDomicilio_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextNumeroRemito_GotFocus()
    
    TextNumeroRemito.SelLength = Len(TextNumeroRemito.text)

End Sub

Private Sub TextNumeroRemito_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextUnidadMedida_GotFocus()
    TextUnidadMedida.SelLength = Len(TextUnidadMedida.text)
End Sub

Private Sub TextUnidadMedida_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextProvincia_Change()

    If TextProvincia.text <> "" Then
        ComboVendedor.SetFocus
    End If
End Sub


