VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormMovimientosCuentaCorriente 
   Caption         =   "Movimento Cuenta Corriente"
   ClientHeight    =   9480
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12390
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9480
   ScaleWidth      =   12390
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   4200
      TabIndex        =   20
      Top             =   720
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.Frame Frame3 
      Height          =   1335
      Left            =   120
      TabIndex        =   28
      Top             =   1560
      Width           =   11655
      Begin VB.OptionButton OptionLinea2 
         Caption         =   "Linea 2"
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
         Left            =   8280
         TabIndex        =   48
         Top             =   840
         Width           =   1455
      End
      Begin VB.OptionButton OptionLinea1 
         Caption         =   "Linea 1"
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
         Top             =   840
         Width           =   1575
      End
      Begin VB.OptionButton OptionTodo 
         Caption         =   "Mostrar Todo"
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
         TabIndex        =   46
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox TextFechaAnterior 
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
         Left            =   10320
         TabIndex        =   38
         Top             =   240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.TextBox TextDocumento 
         Height          =   285
         Left            =   120
         TabIndex        =   37
         Top             =   600
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TextNumeroDocumento 
         Height          =   285
         Left            =   120
         TabIndex        =   36
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TextTipodocumento 
         Height          =   285
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.TextBox TextFechaHasta 
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
         TabIndex        =   31
         Top             =   240
         Width           =   1215
      End
      Begin VB.TextBox TextFechaDesde 
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
         Left            =   3240
         TabIndex        =   29
         Text            =   "01/01/2010"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Hasta:"
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
         TabIndex        =   32
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Desde:"
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
         TabIndex        =   30
         Top             =   360
         Width           =   1200
      End
   End
   Begin VB.Frame Frame2 
      Height          =   5655
      Left            =   120
      TabIndex        =   19
      Top             =   2760
      Width           =   11655
      Begin VB.TextBox TextBusqueda 
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   1095
      End
      Begin VB.TextBox TextSaldoTotalAnt 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9650
         TabIndex        =   41
         Top             =   440
         Width           =   1695
      End
      Begin VB.TextBox TextSaldoLinea2Ant 
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
         Left            =   6510
         TabIndex        =   40
         Top             =   450
         Width           =   1695
      End
      Begin VB.TextBox TextSaldoLinea1Ant 
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
         Left            =   3000
         TabIndex        =   39
         Top             =   450
         Width           =   1695
      End
      Begin VB.TextBox TextSaldoLinea1 
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
         TabIndex        =   24
         Top             =   5085
         Width           =   1695
      End
      Begin VB.TextBox TextSaldoLinea2 
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
         Left            =   5640
         TabIndex        =   23
         Top             =   5085
         Width           =   1695
      End
      Begin VB.TextBox TextSaldoTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9240
         TabIndex        =   22
         Top             =   5085
         Width           =   1695
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
         Height          =   3975
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Visible         =   0   'False
         Width           =   11385
         _ExtentX        =   20082
         _ExtentY        =   7011
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Busqueda"
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
         TabIndex        =   45
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Total:"
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
         Left            =   8520
         TabIndex        =   44
         Top             =   480
         Width           =   1050
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Linea2:"
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
         Left            =   5280
         TabIndex        =   43
         Top             =   480
         Width           =   1185
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Linea 1:"
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
         TabIndex        =   42
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Linea 1:"
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
         TabIndex        =   27
         Top             =   5160
         Width           =   1245
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Linea2:"
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
         Left            =   4200
         TabIndex        =   26
         Top             =   5160
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Total:"
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
         TabIndex        =   25
         Top             =   5160
         Width           =   1050
      End
   End
   Begin VB.Frame Frame4 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   11655
      Begin VB.TextBox TextProvincia 
         Enabled         =   0   'False
         Height          =   285
         Left            =   8280
         TabIndex        =   7
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCodigoPostal 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TextLocalidad 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1080
         TabIndex        =   5
         Top             =   960
         Width           =   3135
      End
      Begin VB.TextBox TextCuit 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   600
         Width           =   1815
      End
      Begin VB.TextBox TextApellidoNombre 
         Height          =   285
         Left            =   7080
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
      Begin VB.TextBox TextCodigoCliente 
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   240
         Width           =   1815
      End
      Begin VB.TextBox TextDireccion 
         Enabled         =   0   'False
         Height          =   285
         Left            =   7080
         TabIndex        =   2
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
         TabIndex        =   14
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
         TabIndex        =   13
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
         TabIndex        =   12
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
         TabIndex        =   11
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
         TabIndex        =   10
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
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   240
         Width           =   1290
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   120
      TabIndex        =   15
      Top             =   8280
      Width           =   11655
      Begin VB.TextBox TextTot 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   11160
         TabIndex        =   50
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton BotonExportar 
         Caption         =   "&Exportar al Excel"
         Height          =   510
         Left            =   7560
         TabIndex        =   34
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton BotonBuscar 
         Caption         =   "&Buscar"
         Height          =   510
         Left            =   840
         TabIndex        =   33
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton BotonImprimir 
         Caption         =   "&Imprimir"
         Height          =   510
         Left            =   3120
         TabIndex        =   18
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton BotonCancelar 
         Caption         =   "&Cancelar"
         Height          =   510
         Left            =   5400
         TabIndex        =   17
         Top             =   360
         Width           =   990
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   510
         Left            =   10080
         TabIndex        =   16
         Top             =   360
         Width           =   990
      End
   End
End
Attribute VB_Name = "FormMovimientosCuentaCorriente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Dim db As DAO.Database
 Dim rstCliente As DAO.Recordset
 Dim rstMovimientosCtaCte As DAO.Recordset
 Dim saldoLinea1 As Double
 Dim saldoLinea2 As Double
 Dim SaldoTotal As Double
 Dim todo As Integer
 Dim linea1 As Integer
 Dim linea11 As Integer
 Dim linea2 As Integer
 Dim linea22 As Integer

Private Sub BuscoSaldos()
        
    Set tCtaCte = db.OpenRecordset("CtaCte", dbOpenTable, dbReadOnly)
    
    tCtaCte.Index = "PrimaryKey"
    
    tCtaCte.Seek "=", TextCodigoCliente.Text
    
    If Not tCtaCte.NoMatch Then
        TextSaldoLinea1.Text = tCtaCte!SaldoL1
        TextSaldoLinea2.Text = tCtaCte!SaldoL2
        TextSaldoTotal.Text = tCtaCte!SaldoTotal
    Else
        TextSaldoLinea1.Text = 0
        TextSaldoLinea2.Text = 0
        TextSaldoTotal.Text = 0
    End If
    
    tCtaCte.Close
    
End Sub

Private Sub ImprimirMovimientos()

    Dim objPrinterFlex As PrinterFlex
    Set objPrinterFlex = New PrinterFlex
    
    With objPrinterFlex
      
      'Asignamos los valores de los encabezados, el pie de página, el color_
      'del texto y el tamaño de la fuente
        
        'texto de los encabezdos y el pie de pagina
        .TextEncabezado1 = Chr(9) & "MOVIMIENTOS CTA. CTE.: " & Chr(9) & TextCodigoCliente.Text & " - " & TextApellidoNombre.Text
                    
        '            nVendedor = Chr(9) & cmbVendedores(1).Text
                    'pie = Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & Chr(9) & "Liquidación Total: " & FormatCurrency(txtImporteTotal.Text, 2)
                    
                    Pie = Chr(9) & Chr(9) & "Saldo Línea 1: " & FormatCurrency(TextSaldoLinea1.Text, 2) & Chr(9) & Chr(9) & "Saldo Línea 2: " & FormatCurrency(TextSaldoLinea2.Text, 2) & Chr(9) & Chr(9) & "Saldo Total: " & FormatCurrency(TextSaldoTotal.Text, 2)
                    
                    'Pie = "Desarrollado por SPC Software Integral"
        
        .TextEncabezado2 = Chr(9) & Chr(9) & "Desde el " & TextFechaDesde.Text & " al " & TextFechaHasta.Text
                
        'CGrid.Row = 1
        'CGrid.Col = 10
        'Anio = CGrid.Text
        'CGrid.Col = 11
        'Periodo = CGrid.Text
        
        .TextPiePagina = Pie
               
        'Colores de la fuentes
        .ColorPiePagina = QBColor(4)
        'txtPiePagina.ForeColor
        .ColorEncabezado1 = QBColor(1)
        'txtEncabezado1.ForeColor
        .ColorEncabezado2 = QBColor(0)
        'txtEncabezado2.ForeColor
        
        'Tamaños de las fuentes
        .SizeEncabezado1 = 12
        .SizeEncabezado2 = 10
        .SizePiePagina = 11
        '.AjustarColumnas = True
        .AjustarColumnas = False
      
        '.Orientacion = Horizontal
        .Orientacion = Vertical
        'Imprimimos pasando el nombre del FlexGrid a imprimir
        .ImprimirFlexGrid MSFlexGrid2
        'FG1
    End With
    
    'Call objPrinterFlex.ImprimirFlexGrid(CGrid)

End Sub

Public Sub BotonBuscar_Click()
    
    Call buscomovimientosCtaCte

End Sub

Private Sub BotonBuscar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub BotonCancelar_Click()

    'call blanqueototal
    
End Sub


Private Sub buscomovimientosCtaCte()

    Dim SaldoL1, SaldoL2, SaldoTotal, Saldo, totalLinea, totalLineatot As Double
    
    saldoLinea1 = 0
    saldoLinea2 = 0
    
    SaldoL1 = 0
    SaldoL2 = 0
    SaldoTotal = 0
    Saldo = 0
    totalLinea = 0
    totalLineatot = 0
    
'***************Busco en PagoProvret
    
On Error GoTo Error_Handler
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")

    'vSQL = "SELECT * FROM MovimientosCtaCte WHERE IDCliente =" & TextCodigoCliente.Text & " ORDER BY Fecha DESC"
    vSQL = "SELECT * FROM MovimientosCtaCte WHERE IDCliente =" & TextCodigoCliente.Text & " ORDER BY Fecha ASC"
    
'    MsgBox (vSQL)
    
    Set tMovCC = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    tMovCC.MoveFirst
    
    While CVDate(tMovCC!Fecha) < CVDate(TextFechaDesde.Text)
'    While CVDate(tMovCC!Fecha) > CVDate(TextFechaDesde.Text)
        SaldoL1 = SaldoL1 + tMovCC!ImporteLinea1
        SaldoL2 = SaldoL2 + tMovCC!ImporteLinea2
        tMovCC.MoveNext
    Wend
    SaldoTotal = SaldoL1 + SaldoL2
    
    TextSaldoLinea1Ant.Text = Format(SaldoL1, "#0.00")
    TextSaldoLinea2Ant.Text = Format(SaldoL2, "#0.00")
    
    TextSaldoTotalAnt.Text = Format(SaldoTotal, "#0.00")
    
    Saldo = TextSaldoTotalAnt.Text
    
   ' lblSaldoAnterior.Caption = " Línea 1: $ " + CStr(CCur(SaldoL1)) + "     Línea 2: $ " + CStr(CCur(SaldoL2)) + "     Saldo Anterior Consolidado: $ " + CStr(CCur(SaldoTotal))
    
    
    CodigoClie = Val(TextCodigoCliente.Text)
       
    Desde = "#" & Format$(TextFechaDesde.Text, "mm/dd/yyyy") & "#"
    Hasta = "#" & Format$(TextFechaHasta.Text, "mm/dd/yyyy") & "#"
    
    'eseqele = "SELECT * FROM MovimientosCtaCte WHERE IDCliente = " & CodigoClie & " AND Fecha >=" & Desde & " AND Fecha <=" & Hasta & " Order By Fecha DESC ,TipoDoc,NroDoc"
    eseqele = "SELECT * FROM MovimientosCtaCte WHERE IDCliente = " & CodigoClie & " AND Fecha >=" & Desde & " AND Fecha <=" & Hasta & " Order By Fecha ASC ,TipoDoc,NroDoc"
    
    'MsgBox (eseqele)
    
    Set rst = db.OpenRecordset(eseqele, dbOpenDynaset)
   
    MSFlexGrid2.Rows = 2
    MSFlexGrid2.Clear
    MSFlexGrid2.Visible = True
    
    If todo = 1 Then
        Call titulosmovimientostodos
    End If
    
    If linea1 = 1 Then
        Call titulosmovimientoslinea1
    End If
    
    If linea2 = 1 Then
        Call titulosmovimientoslinea2
    End If
    
    
    
    rst.MoveFirst

    linea2 = 1
    
    'While Not rst.EOF
    Do While Not rst.NoMatch
         
         '********** Muestro Todo
         
         'If todo = 1 Then
         If OptionTodo.Value = True Then
            MSFlexGrid2.AddItem " "
            MSFlexGrid2.Row = linea2
            MSFlexGrid2.Col = 0
            MSFlexGrid2.Text = rst.Fields!Fecha
            MSFlexGrid2.Col = 1
            MSFlexGrid2.Text = rst.Fields!tipoDoc
            If rst.Fields!tipoDoc = "Pago Linea 1" Or rst.Fields!tipoDoc = "Pago Linea 2" Or rst.Fields!tipoDoc = "Nota Credito A" Or rst.Fields!tipoDoc = "Nota Credito B" Then
                For J = 0 To MSFlexGrid2.Cols - 1
                    MSFlexGrid2.Col = J
                    MSFlexGrid2.CellForeColor = vbRed
                Next J
            End If
                  
            MSFlexGrid2.Col = 2
            'MSFlexGrid2.Text = FormatNumber(rst.Fields!NroDoc, 0)
            MSFlexGrid2.Text = rst.Fields!NroDoc
                    
            MSFlexGrid2.Col = 3
            MSFlexGrid2.Text = Format(rst.Fields!ImporteLinea1, "Standard")
            saldoLinea1 = saldoLinea1 + rst.Fields!ImporteLinea1
            MSFlexGrid2.Col = 4
            MSFlexGrid2.Text = Format(rst.Fields!ImporteLinea2, "Standard")
            saldoLinea2 = saldoLinea2 + rst.Fields!ImporteLinea2
            
            '**************** Calculo Saldos Parciales
            
            If inicio = 0 Then
               totalLinea = Saldo + saldoLinea1 + saldoLinea2
               inicio = 1
            Else
               totalLinea = totalLinea + saldoLinea1 + saldoLinea2
            End If
            MSFlexGrid2.Col = 5
            MSFlexGrid2.CellForeColor = vbBlack
            MSFlexGrid2.Text = Format(totalLinea, "Standard")
            
            S1 = S1 + saldoLinea1
            S2 = S2 + saldoLinea2
            saldoLinea1 = 0
            saldoLinea2 = 0
            
            totalLineatot = totalLinea + totalLineatot
                    
            If rst.Fields!tipoDoc = "Pago Linea 1" Or rst.Fields!tipoDoc = "Pago Linea 2" Then
               MSFlexGrid2.Col = 6
               MSFlexGrid2.Text = "PAGO"
            End If
            If rst.Fields!tipoDoc = "Factura A" Or rst.Fields!tipoDoc = "Factura B" Then
               MSFlexGrid2.Col = 6
               MSFlexGrid2.Text = "FACTURA"
            End If
            If rst.Fields!tipoDoc = "Presupuesto" Then
              MSFlexGrid2.Col = 6
              MSFlexGrid2.Text = "PRESUPUESTO"
            End If
            If rst.Fields!tipoDoc = "Nota Credito A" Or rst.Fields!tipoDoc = "Nota Credito B" Then
               MSFlexGrid2.Col = 6
               MSFlexGrid2.Text = "NOTA CREDITO"
            End If
            linea2 = linea2 + 1
            
            
            
         End If
                  
         'MSFlexGrid2.CellForeColor = vbBlack
         
         '********** Muestro Linea 1
         
         'If linea1 = 1 Then
         If OptionLinea1.Value = True Then
            If rst.Fields!tipoDoc = "Factura A" Or rst.Fields!tipoDoc = "Factura B" Or rst.Fields!tipoDoc = "Pago Linea 1" Or rst.Fields!tipoDoc = "Nota Credito A" Or rst.Fields!tipoDoc = "Nota Credito B" Then
                MSFlexGrid2.AddItem " "
                MSFlexGrid2.Row = linea2
                MSFlexGrid2.Col = 0
                MSFlexGrid2.Text = rst.Fields!Fecha
                MSFlexGrid2.Col = 1
                MSFlexGrid2.Text = rst.Fields!tipoDoc
                MSFlexGrid2.Col = 2
                'MSFlexGrid2.Text = FormatNumber(rst.Fields!NroDoc, 0)
                MSFlexGrid2.Text = rst.Fields!NroDoc
            
                If rst.Fields!tipoDoc = "Pago Linea 1" Or rst.Fields!tipoDoc = "Nota Credito A" Or rst.Fields!tipoDoc = "Nota Credito B" Then
                        For J = 0 To MSFlexGrid2.Cols - 1
                        MSFlexGrid2.Col = J
                        MSFlexGrid2.CellForeColor = vbRed
                    Next J
                End If
                   
                MSFlexGrid2.Col = 3
                MSFlexGrid2.Text = Format(rst.Fields!ImporteLinea1, "Standard")
                saldoLinea1 = saldoLinea1 + rst.Fields!ImporteLinea1
              
                'MSFlexGrid2.Col = 4
                'MSFlexGrid2.Text = rst.Fields!ImporteLinea2
                'saldoLinea2 = saldoLinea2 + rst.Fields!ImporteLinea2
                
                
                If inicio = 0 Then
                   totalLinea = SaldoL1 + saldoLinea1 + saldoLinea2
                   inicio = 1
                Else
                   totalLinea = totalLinea + saldoLinea1 + saldoLinea2
                End If
                
                MSFlexGrid2.Col = 5
                MSFlexGrid2.CellForeColor = vbBlack
                MSFlexGrid2.Text = Format(totalLinea, "Standard")
                
                S1 = S1 + saldoLinea1
                S2 = S2 + saldoLinea2
                saldoLinea1 = 0
                saldoLinea2 = 0
                
                totalLineatot = totalLinea + totalLineatot
                        
                If rst.Fields!tipoDoc = "Pago Linea 1" Then
                   MSFlexGrid2.Col = 6
                   MSFlexGrid2.Text = "PAGO"
                End If
                If rst.Fields!tipoDoc = "Factura A" Or rst.Fields!tipoDoc = "Factura B" Then
                   MSFlexGrid2.Col = 6
                   MSFlexGrid2.Text = "FACTURA"
                End If
                'If rst.Fields!tipoDoc = "Presupuesto" Then
                '  MSFlexGrid2.Col = 5
                '   MSFlexGrid2.Text = "PRESUPUESTO"
                'End If
                If rst.Fields!tipoDoc = "Nota Credito A" Or rst.Fields!tipoDoc = "Nota Credito B" Then
                   MSFlexGrid2.Col = 6
                   MSFlexGrid2.Text = "NOTA CREDITO"
                End If
            linea2 = linea2 + 1
            End If
         End If
         
         '********** Muestro Linea 2
         
         'If linea2 = 1 Then
          If OptionLinea2.Value = True Then
            'If rst.Fields!tipoDoc = "Pago Linea 2" Or rst.Fields!tipoDoc = "Presupuesto" Then
            If rst.Fields!tipoDoc = "Pago Linea 2" Or rst.Fields!tipoDoc = "Presupuesto" Or Left(rst.Fields!tipoDoc, 14) = "Anulacion Pago" Then
                MSFlexGrid2.AddItem " "
                MSFlexGrid2.Row = linea2
                MSFlexGrid2.Col = 0
                MSFlexGrid2.Text = rst.Fields!Fecha
                MSFlexGrid2.Col = 1
                MSFlexGrid2.Text = rst.Fields!tipoDoc
                MSFlexGrid2.Col = 2
                'MSFlexGrid2.Text = FormatNumber(rst.Fields!NroDoc, 0)
                MSFlexGrid2.Text = rst.Fields!NroDoc
            
                If rst.Fields!tipoDoc = "Pago Linea 2" Then
                        For J = 0 To MSFlexGrid2.Cols - 1
                        MSFlexGrid2.Col = J
                        MSFlexGrid2.CellForeColor = vbRed
                    Next J
                End If
                   
                MSFlexGrid2.Col = 4
                MSFlexGrid2.Text = Format(rst.Fields!ImporteLinea2, "Standard")
                saldoLinea2 = saldoLinea2 + rst.Fields!ImporteLinea2
                
                '**************** Calculo Saldos Parciales
            
                If inicio = 0 Then
                   totalLinea = SaldoL2 + saldoLinea1 + saldoLinea2
                   inicio = 1
                Else
                   totalLinea = totalLinea + saldoLinea1 + saldoLinea2
                End If
                MSFlexGrid2.Col = 5
                MSFlexGrid2.CellForeColor = vbBlack
                MSFlexGrid2.Text = Format(totalLinea, "Standard")
                
                S1 = S1 + saldoLinea1
                S2 = S2 + saldoLinea2
                saldoLinea1 = 0
                saldoLinea2 = 0
                
                totalLineatot = totalLinea + totalLineatot
                        
                If rst.Fields!tipoDoc = "Pago Linea 1" Then
                   MSFlexGrid2.Col = 6
                   MSFlexGrid2.Text = "PAGO"
                End If
                If rst.Fields!tipoDoc = "Presupuesto" Then
                  MSFlexGrid2.Col = 6
                  MSFlexGrid2.Text = "PRESUPUESTO"
                End If
            linea2 = linea2 + 1
            End If
         End If
         rst.MoveNext
   Loop
   'Wend
   
Error_Handler:
    
    If Err = 3021 Or Err = 440 Then
        'Nada solo para capturar el error.
    End If
    
   'TxtTotalRetencion.Text = Format(totalrete, "#0.00")
   ' TxtTOTAL.Text = Format(totalpa, "#0.00")
   
    'saldoLinea1 = saldoLinea1 + SaldoL1
    'saldoLinea2 = saldoLinea2 + SaldoL2
   
    saldoLinea1 = S1 + SaldoL1
    saldoLinea2 = S2 + SaldoL2
    TextSaldoLinea1.Text = Format(saldoLinea1, "#0.00")
    TextSaldoLinea2.Text = Format(saldoLinea2, "#0.00")
    
    SaldoTotal = saldoLinea1 + saldoLinea2
    
    TextSaldoTotal.Text = Format(SaldoTotal, "#0.00")
    TextTot.Text = Format(totalLineatot, "#0.00")
    Exit Sub
    
   
End Sub


Private Sub blanqueototal()

    TextCodigoCliente.Text = ""
    TextApellidoNombre.Text = ""
    TextCuit.Text = ""
    TextDireccion.Text = ""
    TextLocalidad.Text = ""
    TextCodigoPostal.Text = ""
    TextProvincia.Text = ""
    
    Label10.Visible = True
    TextSaldoLinea1Ant.Visible = True
    TextSaldoLinea1Ant.Text = ""
    
    Label11.Visible = True
    TextSaldoLinea2Ant.Visible = True
    TextSaldoLinea2Ant.Text = ""
    
    Label12.Visible = True
    TextSaldoTotalAnt.Visible = True
    TextSaldoTotalAnt.Text = ""
        
    Label3.Visible = True
    TextSaldoLinea2.Visible = True
    TextSaldoLinea2.Text = ""
       
    Label4.Visible = True
    TextSaldoLinea1.Visible = True
    TextSaldoLinea1.Text = ""
    
    Label2.Visible = True
    TextSaldoTotal.Visible = True
    TextSaldoTotal.Text = ""
    
    MSFlexGrid2.Visible = False
    
    
    Call titulosmovimientostodos

End Sub
Private Sub titulosmovimientostodos()

    MSFlexGrid2.Row = 0
    
    MSFlexGrid2.Col = 0
    MSFlexGrid2.Text = "Fecha"
    MSFlexGrid2.ColWidth(0) = 1900
    MSFlexGrid2.ColAlignment(0) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 1
    MSFlexGrid2.Text = "Tipo Documento"
    MSFlexGrid2.ColWidth(1) = 3100
    MSFlexGrid2.ColAlignment(1) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 2
    MSFlexGrid2.Text = "Nro Documento"
    MSFlexGrid2.ColWidth(2) = 1500
    MSFlexGrid2.ColAlignment(2) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 3
    MSFlexGrid2.Text = "Importe Linea 1"
    MSFlexGrid2.ColWidth(3) = 1500
    MSFlexGrid2.ColAlignment(3) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 4
    MSFlexGrid2.Text = "Importe Linea 2"
    MSFlexGrid2.ColWidth(4) = 1500
    MSFlexGrid2.ColAlignment(4) = flexAlignCenterCenter
     
    MSFlexGrid2.Col = 5
    MSFlexGrid2.Text = "Saldos Parci."
    MSFlexGrid2.ColWidth(5) = 1400
    MSFlexGrid2.ColAlignment(5) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 6
    MSFlexGrid2.Text = ""
    MSFlexGrid2.ColWidth(6) = 0
    MSFlexGrid2.ColAlignment(6) = flexAlignCenterCenter
    
    Label10.Visible = True
    TextSaldoLinea1Ant.Visible = True
    
    Label11.Visible = True
    TextSaldoLinea2Ant.Visible = True
    
    Label12.Visible = True
    TextSaldoTotalAnt.Visible = True
        
    Label3.Visible = True
    TextSaldoLinea2.Visible = True
       
    Label4.Visible = True
    TextSaldoLinea1.Visible = True
    
    Label2.Visible = True
    TextSaldoTotal.Visible = True
    
    'todo = 0
        
 End Sub
Private Sub titulosmovimientoslinea1()

    MSFlexGrid2.Row = 0
    
    MSFlexGrid2.Col = 0
    MSFlexGrid2.Text = "Fecha"
    MSFlexGrid2.ColWidth(0) = 1900
    MSFlexGrid2.ColAlignment(0) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 1
    MSFlexGrid2.Text = "Tipo Documento"
    MSFlexGrid2.ColWidth(1) = 3100
    MSFlexGrid2.ColAlignment(1) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 2
    MSFlexGrid2.Text = "Nro Documento"
    MSFlexGrid2.ColWidth(2) = 1600
    MSFlexGrid2.ColAlignment(2) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 3
    MSFlexGrid2.Text = "Importe Linea 1"
    MSFlexGrid2.ColWidth(3) = 1600
    MSFlexGrid2.ColAlignment(3) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 4
    MSFlexGrid2.Text = ""
    MSFlexGrid2.ColWidth(4) = 0
    MSFlexGrid2.ColAlignment(4) = flexAlignCenterCenter
 
    MSFlexGrid2.Col = 5
    MSFlexGrid2.Text = "Saldos Parci."
    MSFlexGrid2.ColWidth(5) = 1400
    MSFlexGrid2.ColAlignment(5) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 6
    MSFlexGrid2.Text = ""
    MSFlexGrid2.ColWidth(6) = 0
    MSFlexGrid2.ColAlignment(6) = flexAlignCenterCenter
    
    Label11.Visible = False
    TextSaldoLinea2Ant.Visible = False
    
    Label12.Visible = False
    TextSaldoTotalAnt.Visible = False
    
    Label3.Visible = False
    TextSaldoLinea2.Visible = False
    
    Label2.Visible = False
    TextSaldoTotal.Visible = False
    
    
    Label10.Visible = True
    TextSaldoLinea1Ant.Visible = True
    
    
    
    Label4.Visible = True
    TextSaldoLinea1.Visible = True
    
    
    
    'linea11 = 0
    'linea11 = 1
 End Sub
Private Sub titulosmovimientoslinea2()

    MSFlexGrid2.Row = 0
    
    MSFlexGrid2.Col = 0
    MSFlexGrid2.Text = "Fecha"
    MSFlexGrid2.ColWidth(0) = 1900
    MSFlexGrid2.ColAlignment(0) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 1
    MSFlexGrid2.Text = "Tipo Documento"
    MSFlexGrid2.ColWidth(1) = 3100
    MSFlexGrid2.ColAlignment(1) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 2
    MSFlexGrid2.Text = "Nro Documento"
    MSFlexGrid2.ColWidth(2) = 1600
    MSFlexGrid2.ColAlignment(2) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 3
    MSFlexGrid2.Text = "Importe Linea 1"
    MSFlexGrid2.ColWidth(3) = 0
    MSFlexGrid2.ColAlignment(3) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 4
    MSFlexGrid2.Text = "Importe Linea 2"
    MSFlexGrid2.ColWidth(4) = 1600
    MSFlexGrid2.ColAlignment(4) = flexAlignCenterCenter
 
    MSFlexGrid2.Col = 5
    MSFlexGrid2.Text = "Saldos Parci."
    MSFlexGrid2.ColWidth(5) = 1400
    MSFlexGrid2.ColAlignment(5) = flexAlignCenterCenter
    
    MSFlexGrid2.Col = 6
    MSFlexGrid2.Text = ""
    MSFlexGrid2.ColWidth(6) = 0
    MSFlexGrid2.ColAlignment(6) = flexAlignCenterCenter
    
    Label10.Visible = False
    TextSaldoLinea1Ant.Visible = False
    
    Label12.Visible = False
    TextSaldoTotalAnt.Visible = False
    
    Label4.Visible = False
    TextSaldoLinea1.Visible = False
    
    Label2.Visible = False
    TextSaldoTotal.Visible = False
    
    
    Label11.Visible = True
    TextSaldoLinea2Ant.Visible = True
    
    
    
    Label3.Visible = True
    TextSaldoLinea2.Visible = True
    
   
    
    'linea2 = 0
    'linea22 = 1
 End Sub


Private Sub BotonPago_Click()

    FormPagoFacturas.Show
    
End Sub

Private Sub BotonCancelar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonExportar_Click()
    
    If Exportar_Excel(App.Path & "\Movimiento Cuenta Corriente.xls", MSFlexGrid2) Then
        MsgBox " Datos exportados en " & App.Path, vbInformation
    End If
    
End Sub
Public Function Exportar_Excel(sOutputPath As String, FlexGrid As Object) As Boolean
  
    On Error GoTo Error_Handler
  
    Dim o_Excel     As Object
    Dim o_Libro     As Object
    Dim o_Hoja      As Object
    Dim Fila        As Long
    Dim Columna     As Long
      
    ' -- Crea el objeto Excel, el objeto workBook y el objeto sheet
    Set o_Excel = CreateObject("Excel.Application")
    Set o_Libro = o_Excel.Workbooks.Add
    Set o_Hoja = o_Libro.Worksheets.Add
    
    ' -- Bucle para Exportar los datos
    With MSFlexGrid2
        For Fila = 1 To .Rows - 1
            'If linea11 = 1 Then
            '    For Columna = 0 To .Cols - 3
            '        o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            '    Next
            'End If
            
            'If linea22 = 1 Then
                
            '    For Columna = 0 To .Cols - 2
            '
            '        .ColWidth(3) = 0
            '        .Col = 3
            '        .Visible = False
            '        o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            '    Next
            'End If
            For Columna = 0 To .Cols - 2
                o_Hoja.Cells(Fila, Columna + 1).Value = .TextMatrix((Fila - 1), Columna)
            Next
        
        Next
    End With
    'o_Libro.Close True, sOutputPath
    
    o_Libro.Close True, sOutputPath
    Set o_Libro = o_Excel.Workbooks.Open(sOutputPath)
    o_Excel.Visible = True
    
    Call blanqueototal
    ' -- Cerrar Excel
    'o_Excel.Quit
    ' -- Terminar instancias
    'Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    'Exportar_Excel = True
Exit Function

  
' -- Controlador de Errores
Error_Handler:
    ' -- Cierra la hoja y el la aplicación Excel
    If Not o_Libro Is Nothing Then: o_Libro.Close False
    If Not o_Excel Is Nothing Then: o_Excel.Quit
    Call ReleaseObjects(o_Excel, o_Libro, o_Hoja)
    If Err.Number <> 1004 Then MsgBox Err.Description, vbCritical
    
End Function
Private Sub ReleaseObjects(o_Excel As Object, o_Libro As Object, o_Hoja As Object)
    If Not o_Excel Is Nothing Then Set o_Excel = Nothing
    If Not o_Libro Is Nothing Then Set o_Libro = Nothing
    If Not o_Hoja Is Nothing Then Set o_Hoja = Nothing
End Sub



Private Sub BotonExportar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonImprimir_Click()

    Call ImprimirMovimientos

End Sub

Private Sub BotonImprimir_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub BotonSalir_Click()

    Unload FormMovimientosCuentaCorriente

End Sub


Private Sub Busco()

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


Private Sub BotonSalir_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

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
    
    MSFlexGrid1.Visible = False
    
   'Call buscomovimientosCtaCte

End Sub

Private Sub OptionFacturaTodas_Click()

    buscomovimientosCtaCte

End Sub

'Private Sub MSFlexGrid2_Click()

'   MSFlexGrid2.Col = 1
'   TextTipodocumento.Text = MSFlexGrid2.Text
'   MSFlexGrid2.Col = 2
'   TextNumeroDocumento.Text = Val(MSFlexGrid2.Text)
'   MSFlexGrid2.Col = 5
'   TextDocumento.Text = MSFlexGrid2.Text
   
'   If TextDocumento.Text = "PAGO" Then
'        FormVerPagoFacturas.Show
'   End If
   
'   If TextDocumento.Text = "FACTURA" Then
'        FormVerFactura.Show
'   End If
   
 '  If TextDocumento.Text = "PRESUPUESTO" Then
 '      FormVerPresupuestos.Show
 '  End If
   
 '  If TextDocumento.Text = "NOTA CREDITO" Then
 '      FormVerNotacredito.Show
 '  End If
   
'End Sub

Private Sub MSFlexGrid2_DblClick()

   MSFlexGrid2.Col = 1
   TextTipodocumento.Text = MSFlexGrid2.Text
   MSFlexGrid2.Col = 2
   TextNumeroDocumento.Text = Val(MSFlexGrid2.Text)
   MSFlexGrid2.Col = 6
   TextDocumento.Text = MSFlexGrid2.Text
   
   If TextDocumento.Text = "PAGO" Then
        FormVerPagoFacturas.Show
        'FormAnulacionPago.Show
   End If
   
   If TextDocumento.Text = "FACTURA" Then
        FormVerFacturaCtaCte.Show
   End If
   
   If TextDocumento.Text = "PRESUPUESTO" Then
       FormVerPresupuestos.Show
   End If
   
   If TextDocumento.Text = "NOTA CREDITO" Then
       'FormVerNotacredito.Show
       FormVerNotaCreditoCtaCte.Show
   End If

End Sub

Private Sub OptionLinea1_Click()
    If OptionLinea1.Value = True Then
        linea1 = 1
    End If
    linea2 = 0
    todo = 0
End Sub

Private Sub OptionLinea2_Click()
    If OptionLinea2.Value = True Then
        linea2 = 1
    End If
    linea1 = 0
    todo = 0
End Sub

Private Sub OptionTodo_Click()
    If OptionTodo.Value = True Then
        todo = 1
    End If
    linea1 = 0
    linea2 = 0
End Sub

Private Sub TextApellidoNombre_Change()
   
        Columna = 1
        Call FiltrarGrilla(MSFlexGrid1, TextApellidoNombre, CLng(Columna))
  
End Sub

Private Sub TextApellidoNombre_KeyPress(KeyAscii As Integer)

'    If KeyAscii = 13 Then
'        MSFlexGrid2.Visible = False
'        Call busco
'    End If
    
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

Private Sub blanco()

    TextCodigoCliente.Text = ""
    TextApellidoNombre.Text = ""
    TextCuit.Text = ""
    TextDireccion.Text = ""
    TextLocalidad.Text = ""
    TextCodigoPostal.Text = ""
    TextProvincia.Text = ""
    
End Sub



Private Sub TextBusqueda_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        'Damos al FlexiGrid el color de fondo por defecto
        MSFlexGrid2.BackColor = &H80000005
        'Si la caja de texto está vacía eliminamos el contenido del label y salimos
        If TextBusqueda.Text = "" Then
            Label1.Caption = ""
            Exit Sub
        End If
        I = 1
        J = 2
        'Recorremos todas la filas del FlexiGrid columna a columna
        For I = 1 To MSFlexGrid2.Rows - 1
        For J = 1 To MSFlexGrid2.Cols - 1
        'comprobamos si coincide el contenido del TextBusqueda con la celda
        If LCase(TextBusqueda.Text) = LCase(Mid(MSFlexGrid2.TextMatrix(I, J), 1, Len(TextBusqueda.Text))) Then
        'En caso afirmativo mostramos su contenido en un Label1
        Label1.Caption = MSFlexGrid2.TextMatrix(I, J)
        'Seleccionamos la celda para darle color de fondo
        MSFlexGrid2.Row = I
        MSFlexGrid2.Col = J - 1
        MSFlexGrid2.ColSel = J
        MSFlexGrid2.BackColorSel = QBColor(1)
        'Damos unos valores a I y J para que salga de nol dos For y no continue buscando. Si no hiciéramos esto el label mostraría la última celda que coincida con el contenido del TextBusqueda
        I = MSFlexGrid2.Rows + 1
        J = MSFlexGrid2.Cols + 1
        End If
        Next J
        Next I
        
        TextBusqueda.Text = ""
        TextBusqueda.SetFocus
    
    End If
    
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
            If rstCliente.Fields!IdCliente <> Val(TextCodigoCliente.Text) Then
                mensaje = MsgBox("Cliente Inexistente", vbCritical, "Final de la busqueda")
                TextCodigoCliente.Text = ""
                Call blanco
                TextCodigoCliente.SetFocus
            Else
                TextCodigoCliente.Text = rstCliente.Fields!IdCliente
                TextApellidoNombre.Text = rstCliente.Fields!RazonSocial
                If rstCliente.Fields!CUIT <> "" Then
                    TextCuit.Text = rstCliente.Fields!CUIT
                End If
                TextDireccion.Text = rstCliente.Fields!Domicilio
                TextLocalidad.Text = rstCliente.Fields!Localidad
                TextCodigoPostal.Text = rstCliente.Fields!CP
                TextProvincia.Text = rstCliente.Fields!Prov
            End If
        End If
        ' Call buscomovimientosCtaCte
         Call BuscoSaldos
         'BotonBuscar.SetFocus
         TextFechaDesde.SetFocus
    End If
   
    If KeyAscii = 27 Then
        Unload Me
    End If
    MSFlexGrid1.Visible = False
End Sub

Private Sub Form_Load()

    FormMovimientosCuentaCorriente.Height = 10050
    FormMovimientosCuentaCorriente.Width = 12105
    FormMovimientosCuentaCorriente.Top = 800
    FormMovimientosCuentaCorriente.Left = 800
    
    OptionTodo = True
    
    TextFechaHasta.Text = Format(Date, "dd/mm/yyyy")

End Sub



Private Sub TextFechaDesde_GotFocus()
    TextFechaDesde.SelLength = Len(TextFechaDesde.Text)
End Sub

Private Sub TextFechaDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub


Private Sub TextFechaDesde_LostFocus()

    If TextFechaDesde.Text = "" Then TextFechaDesde.Text = Format("01/01/1900", "dd/mm/yyyy")

End Sub


Private Sub TextFechaHasta_GotFocus()
    TextFechaHasta.SelLength = Len(TextFechaHasta.Text)
End Sub

Private Sub TextFechaHasta_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Call buscomovimientosCtaCte
        KeyAscii = 0
        SendKeys "{TAB}"
    End If

End Sub

Private Sub TextFechaHasta_LostFocus()

    If TextFechaHasta.Text = "" Then TextFechaHasta.Text = Format(Date, "dd/mm/yyyy")
    
End Sub


