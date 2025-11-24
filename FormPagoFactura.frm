VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form FormPagoFacturas 
   Caption         =   "Pago Facturas"
   ClientHeight    =   8550
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8550
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PictureQP 
      Height          =   615
      Left            =   8400
      ScaleHeight     =   555
      ScaleWidth      =   555
      TabIndex        =   48
      Top             =   4800
      Width           =   615
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   45
      Top             =   1440
      Visible         =   0   'False
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   3201
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
   End
   Begin VB.Frame Frame3 
      Caption         =   "Destinado a"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   37
      Top             =   1560
      Width           =   7695
      Begin VB.OptionButton OptionSaldoLinea2 
         Caption         =   "Saldo Linea 2"
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
         Left            =   4200
         TabIndex        =   4
         Top             =   480
         Width           =   1935
      End
      Begin VB.OptionButton OptionSaldoLinea1 
         Caption         =   "Saldo Linea 1"
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
         Left            =   1560
         TabIndex        =   3
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   18
      Top             =   7200
      Width           =   7695
      Begin VB.CommandButton cmdPagoCh 
         Caption         =   "Pago por Cheque Rech."
         Height          =   735
         Left            =   5760
         TabIndex        =   49
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdPrintRecibo 
         Caption         =   "Imprimir Recibo Oficial"
         Height          =   735
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonGuardar 
         Caption         =   "&Guardar"
         Height          =   750
         Left            =   2760
         TabIndex        =   14
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forma de Pago"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   120
      TabIndex        =   16
      Top             =   2640
      Width           =   7695
      Begin VB.TextBox textTransferencia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3600
         TabIndex        =   44
         Text            =   "Text1"
         Top             =   360
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.TextBox TextNumeroPago 
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
         Left            =   5880
         TabIndex        =   5
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextFechaPago 
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
         Left            =   5880
         TabIndex        =   13
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TextObservaciones 
         Height          =   525
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   28
         Top             =   2160
         Width           =   4335
      End
      Begin VB.TextBox TextTarjeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   12
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox TextRetencion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   11
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox TextRezago 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TextMercaderia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   10
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TextCheque 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TextResta 
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
         Left            =   2760
         TabIndex        =   24
         Top             =   3960
         Width           =   1335
      End
      Begin VB.TextBox TextEfectivo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Transferencia:"
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
         TabIndex        =   50
         Top             =   1080
         Width           =   1245
      End
      Begin VB.Label Label16 
         Caption         =   "Saldo Actualizado"
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
         Left            =   5280
         TabIndex        =   43
         Top             =   3600
         Width           =   1815
      End
      Begin VB.Label LabelSaldoTotal 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4920
         TabIndex        =   42
         Top             =   3960
         Width           =   2295
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Pago"
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
         Left            =   5880
         TabIndex        =   41
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label LabelTotalAbonado 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5040
         TabIndex        =   40
         Top             =   3000
         Width           =   75
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Abonado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3120
         TabIndex        =   39
         Top             =   3000
         Width           =   1620
      End
      Begin VB.Label LabelSaldo 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   38
         Top             =   4080
         Width           =   75
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Numero Pago"
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
         Left            =   5880
         TabIndex        =   30
         Top             =   240
         Width           =   1155
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Observaciones:"
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
         Left            =   3120
         TabIndex        =   29
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta:"
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
         Top             =   3480
         Width           =   675
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Retencion:"
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
         TabIndex        =   26
         Top             =   3000
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Resta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   25
         Top             =   4080
         Width           =   630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Cheque:"
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
         TabIndex        =   23
         Top             =   1560
         Width           =   720
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Mercaderia:"
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
         TabIndex        =   22
         Top             =   2520
         Width           =   1020
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Rezago:"
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
         TabIndex        =   21
         Top             =   2040
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Efectivo:"
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
         TabIndex        =   20
         Top             =   600
         Width           =   780
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Datos Cliente"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   7695
      Begin VB.ComboBox cmbSucursal 
         Height          =   315
         Left            =   480
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   940
         Width           =   1815
      End
      Begin VB.TextBox Textcod 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   6120
         TabIndex        =   46
         Top             =   600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextNombre 
         Height          =   285
         Left            =   2640
         TabIndex        =   1
         Top             =   360
         Width           =   4455
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
         Left            =   2640
         TabIndex        =   33
         Top             =   960
         Width           =   1335
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
         Left            =   4200
         TabIndex        =   32
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TextSaldoTotal 
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
         Left            =   5760
         TabIndex        =   31
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TextCodigoCliente 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1080
         TabIndex        =   0
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
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
         Left            =   840
         TabIndex        =   51
         Top             =   705
         Width           =   750
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
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
         Left            =   5760
         TabIndex        =   36
         Top             =   720
         Width           =   495
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Linea 1"
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
         TabIndex        =   35
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Linea 2"
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
         TabIndex        =   34
         Top             =   720
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cliente:"
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
         TabIndex        =   17
         Top             =   360
         Width           =   915
      End
   End
End
Attribute VB_Name = "FormPagoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPagoC As DAO.Recordset
Dim rstPagoD As DAO.Recordset
Dim rstCtaCte As DAO.Recordset
Dim rstMovimientosCtaCte As DAO.Recordset
Dim numeroPago As Long
Dim saldoLinea1 As Integer
Dim saldoLinea2 As Integer
Dim saldo1 As Double
Dim saldo2 As Double
Dim saldoLi1 As Double
Dim saldoLi2 As Double
Dim resta As Double
Dim suma As Double
Dim sldoTotalForm As Double
Dim efectivo As Double
Dim transferencia As Double
Dim rezago As Double
Dim mercaderia As Double
Dim cheque As Double
Dim retencion As Double
Dim tarjeta As Double


Private Sub ImprimeReciboPreImpreso()
    
    On Error GoTo CapturaErrores
    
    Dim PU, TL, ImpIva, ImpIIBB, SubTotalFac, TotalFac, Cant As Variant
    Dim vImporteEnLetras, Importe As String
    Dim vCantFac As Integer
    
    x = -4
    Y = -4
    renglon = 0
   ' vNroRemito = "0004- " & TextNumeroRemito.Text
    
        vCantFac = CInt(InputBox("Ingrese Cantidad de Facturas a Detallar", "DETALLE DE FACTURAS DEL REMITO"))
        
        If Val(vCantFac) > 6 Then
            A = MsgBox("Corrija la Cantidad de Facturas" + Chr(10) + "No pueden ser más de 6", vbOKOnly, "CANTIDAD DE FACTURAS")
            vCantFac = InputBox("Ingrese Cantidad de Facturas a Detallar", "DETALLE DE FACTURAS DEL REMITO")
        End If
            
        Dim vFF(6) As Date
        Dim vNFac(6) As String
        Dim vImpF(6) As Double
        
        For J = 1 To vCantFac
            vFF(J) = InputBox("Ingrese La Fecha de La Factura # " & J, "DETALLE DE FACTURAS DEL REMITO")
            vNFac(J) = InputBox("Ingrese El Nro de La Factura # " & J, "DETALLE DE FACTURAS DEL REMITO")
            vImpF(J) = InputBox("Ingrese El Importe de La Factura # " & J, "DETALLE DE FACTURAS DEL REMITO")
        Next J
        
        vConcepto = InputBox("Ingrese El Concepto", "DETALLE DE FACTURAS DEL REMITO")
            
    'FormRecibo.Show
    
    With Printer
        
        'On Error GoTo CapturaErrores
            .Copies = 2
        'Seteo escala a mm
            .ScaleMode = 6
        
        'Imprimir Fecha
            .CurrentX = x + 103
            .CurrentY = Y + 31
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print Format(TextFechaPago.text, "DD")
            .CurrentX = x + 113
            .CurrentY = Y + 31
            Printer.Print Format(TextFechaPago.text, "MM")
            .CurrentX = x + 122
            .CurrentY = Y + 31
            Printer.Print Format(TextFechaPago.text, "YYYY")
                
        'Imprimir Nombres
            .CurrentX = x + 23
            .CurrentY = Y + 52
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = True
            Printer.Print TextNombre.text
            
        'Imprimir Direccion
            .CurrentX = x + 31
            .CurrentY = Y + 60
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print vDomicilio
            
        'Imprimir Localidad
            .CurrentX = x + 115
            .CurrentY = Y + 60
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            Printer.Print vLocalidad
            
        'Imprimir CUIT
            .CurrentX = x + 105
            .CurrentY = Y + 67
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            
            vCUIT = Left(vCUIT, 2) + "-" + Mid(vCUIT, 3, 8) + "-" + Right(vCUIT, 1)
            Printer.Print vCUIT
            
        'Imprimir Marca Responsable Inscripto
            .CurrentX = x + 20
            .CurrentY = Y + 67
            .Font = "Courier New"
            .FontSize = 10
            .FontBold = False
            
            Select Case vCondIVA
            
                Case "RI"
                    vCondIVA = "RESPONSABLE INSCRIPTO"
                
                Case "NR"
                    vCondIVA = "NO RESPONSABLE"
            
                Case "EX"
                    vCondIVA = "EXENTO"
                
                Case "CF"
                    vCondIVA = "CONSUMIDOR FINAL"
                
                Case "MO"
                    vCondIVA = "RESPONSABLE MONOTRIBUTO"
                    
                Case "RN"
                    vCondIVA = "RESPONSABLE NO INSCRIPTO"
            
            End Select
            
            Printer.Print vCondIVA
            
         'Imprime Importe en Letras
            .CurrentX = x + 114
            .CurrentY = Y + 78
            
            Importe = CStr(Format(suma, Fixed))
            vImporteEnLetras = EnLetras(Importe)
            
            'MsgBox (vImporteEnLetras)
            
          'Imprimir importe en letras en renglones
            
            Largo = Len(vImporteEnLetras)
            I = 0
            
            If Len(vImporteEnLetras) >= 18 And Len(vImporteEnLetras) <= 84 Then
                For I = 18 To 1 Step -1
                    If Mid(vImporteEnLetras, I, 1) = " " Then Exit For
                Next I
                Printer.Print StrConv(Mid(vImporteEnLetras, 1, I), vbUpperCase)
                
                If (Largo - I) <= 34 Then
                    .CurrentX = x + 78
                    .CurrentY = Y + 85
                    Printer.Print StrConv(Mid(vImporteEnLetras, (I + 1), (Largo - I)), vbUpperCase)
                 Else
                    .CurrentX = x + 78
                    .CurrentY = Y + 85
                    SegundoTramo = StrConv(Mid(vImporteEnLetras, (I + 1), (Largo - I)), vbUpperCase)
                    Largo = Len(SegundoTramo)
                                        
                    For I = 34 To 1 Step -1
                        If Mid(SegundoTramo, I, 1) = " " Then
                            PosicionT1 = I
                            Exit For
                        End If
                    Next I
                    'Segundo tramo de la cadena
                    Printer.Print StrConv(Mid(SegundoTramo, 1, I), vbUpperCase)
                    .CurrentX = x + 78
                    .CurrentY = Y + 93
                    Printer.Print StrConv(Mid(SegundoTramo, (I + 1), (Largo - I)), vbUpperCase)
                End If
             Else
                If Len(vImporteEnLetras) >= 68 Then
                   'Primer tramo de la cadena
                    For I = 34 To 1 Step -1
                        If Mid(vImporteEnLetras, I, 1) = " " Then
                            PosicionT1 = I
                            Exit For
                        End If
                    Next I
                    Printer.Print StrConv(Mid(vImporteEnLetras, 1, I), vbUpperCase)
                    
                    'Segundo tramo de la cadena
                    For I = 68 To PosicionT1 Step -1
                        If Mid(vImporteEnLetras, I, 1) = " " Then
                            H = (I - 1) - (PosicionT1 - 1)
                            If H <= 34 Then
                                PosicionT2 = I
                                Exit For
                            End If
                        End If
                    Next I
                    
                    .CurrentX = x + 78
                    .CurrentY = Y + 85
                    Hasta = (PosicionT2 - 1) - (PosicionT1 - 1)
                    Printer.Print StrConv(Mid(vImporteEnLetras, (PosicionT1 + 1), Hasta), vbUpperCase)
                    
                    'Tercer tramo de la cadena
                    .CurrentX = x + 78
                    .CurrentY = Y + 93
                    Hasta = Largo - (PosicionT2)
                    If Hasta > 34 Then Hasta = 34
                    Printer.Print StrConv(Mid(vImporteEnLetras, (PosicionT2 + 1), Hasta), vbUpperCase)
                 Else
                    .CurrentX = x + 78
                    .CurrentY = Y + 93
                    Printer.Print StrConv(vImporteEnLetras, vbUpperCase)
                End If
            End If
                        
            .CurrentX = x + 124
            .CurrentY = Y + 124
            
            vEfe = Format(Val(TextEfectivo.text), "Standard")
            Hasta = CInt(10 - Len(vEfe))
            For I = 0 To Hasta
                vEfe = " " & vEfe
            Next I
            Printer.Print vEfe
            
            .CurrentX = x + 124
            .CurrentY = Y + 177
            
            vCheq = Format(Val(TextCheque.text), "Standard")
            Hasta = CInt(10 - Len(vCheq))
            For I = 0 To Hasta
                vCheq = " " & vCheq
            Next I
            Printer.Print vCheq
            'Printer.Print Format(Val(TextCheque.Text), "Standard")
            
            .CurrentX = x + 124
            .CurrentY = Y + 183
            
            vRet = Format(Val(TextRetencion.text), "Standard")
            Hasta = CInt(10 - Len(vRet))
            For I = 0 To Hasta
                vRet = " " & vRet
            Next I
            Printer.Print vRet
            'Printer.Print Format(Val(TextRetencion.Text), "Standard")
                        
            .CurrentX = x + 124
            .CurrentY = Y + 189
            
            vTotal = Format(LabelTotalAbonado.Caption, "Standard")
            Hasta = CInt(10 - Len(vTotal))
            For I = 0 To Hasta
                vTotal = " " & vTotal
            Next I
            Printer.Print vTotal
            'Printer.Print Format(LabelTotalAbonado.Caption, "Standard")
            
        'Imprimir Facturas
            .FontSize = 8
            
            For J = 1 To vCantFac
                Select Case J
                    Case 1
                        .CurrentX = x + 10
                        .CurrentY = Y + 84
                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                        
                        .CurrentX = x + 28
                        .CurrentY = Y + 84
                        Printer.Print vNFac(J)
                        
                        .CurrentX = x + 53
                        .CurrentY = Y + 84
                        Printer.Print Format(Val(vImpF(J)), "##,##")
                    Case 2
                        .CurrentX = x + 10
                        .CurrentY = Y + 89
                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                        
                        .CurrentX = x + 28
                        .CurrentY = Y + 89
                        Printer.Print vNFac(J)
                        
                        .CurrentX = x + 53
                        .CurrentY = Y + 89
                        Printer.Print Format(Val(vImpF(J)), "##,##")
                    Case 3
                        .CurrentX = x + 10
                        .CurrentY = Y + 94
                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                        
                        .CurrentX = x + 28
                        .CurrentY = Y + 94
                        Printer.Print vNFac(J)
                        
                        .CurrentX = x + 53
                        .CurrentY = Y + 94
                        Printer.Print Format(Val(vImpF(J)), "##,##")
                    Case 4
                        .CurrentX = x + 10
                        .CurrentY = Y + 99
                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                        
                        .CurrentX = x + 28
                        .CurrentY = Y + 99
                        Printer.Print vNFac(J)
                        
                        .CurrentX = x + 53
                        .CurrentY = Y + 99
                        Printer.Print Format(Val(vImpF(J)), "##,##")
                    Case 5
                        .CurrentX = x + 10
                        .CurrentY = Y + 104
                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                        
                        .CurrentX = x + 28
                        .CurrentY = Y + 104
                        Printer.Print vNFac(J)
                        
                        .CurrentX = x + 53
                        .CurrentY = Y + 104
                        Printer.Print Format(Val(vImpF(J)), "##,##")
                    
                    Case 6
                        .CurrentX = x + 10
                        .CurrentY = Y + 109
                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                        
                        .CurrentX = x + 28
                        .CurrentY = Y + 109
                        Printer.Print vNFac(J)
                        
                        .CurrentX = x + 53
                        .CurrentY = Y + 109
                        Printer.Print Format(Val(vImpF(J)), "##,##")
                End Select
            Next J
                    
         .FontSize = 10
        
        'Imprime Concepto
            
            If vConcepto <> "" Then
            
                'Imprime en 3 renglones
                   .CurrentX = x + 100
                   .CurrentY = Y + 100
                   
                 'Imprimir importe en letras en renglones
                   
                   Largo = Len(vConcepto)
                   I = 0
                   
                   If Len(vConcepto) >= 24 And Len(vConcepto) <= 92 Then
                       For I = 24 To 1 Step -1
                           If Mid(vConcepto, I, 1) = " " Then Exit For
                       Next I
                       Printer.Print StrConv(Mid(vConcepto, 1, I), vbUpperCase)
                       
                       If (Largo - I) <= 34 Then
                           .CurrentX = x + 78
                           .CurrentY = Y + 107
                           Printer.Print StrConv(Mid(vConcepto, (I + 1), (Largo - I)), vbUpperCase)
                        Else
                           .CurrentX = x + 78
                           .CurrentY = Y + 107
                           SegundoTramo = StrConv(Mid(vConcepto, (I + 1), (Largo - I)), vbUpperCase)
                           Largo = Len(SegundoTramo)
                                               
                           For I = 34 To 1 Step -1
                               If Mid(SegundoTramo, I, 1) = " " Then
                                   PosicionT1 = I
                                   Exit For
                               End If
                           Next I
                           'Segundo tramo de la cadena
                           Printer.Print StrConv(Mid(SegundoTramo, 1, I), vbUpperCase)
                           .CurrentX = x + 78
                           .CurrentY = Y + 114
                           Printer.Print StrConv(Mid(SegundoTramo, (I + 1), (Largo - I)), vbUpperCase)
                       End If
                    Else
                       If Len(vConcepto) >= 68 Then
                          'Primer tramo de la cadena
                           For I = 34 To 1 Step -1
                               If Mid(vConcepto, I, 1) = " " Then
                                   PosicionT1 = I
                                   Exit For
                               End If
                           Next I
                           Printer.Print StrConv(Mid(vConcepto, 1, I), vbUpperCase)
                           
                           'Segundo tramo de la cadena
                           For I = 68 To PosicionT1 Step -1
                               If Mid(vConcepto, I, 1) = " " Then
                                   H = (I - 1) - (PosicionT1 - 1)
                                   If H <= 34 Then
                                       PosicionT2 = I
                                       Exit For
                                   End If
                               End If
                           Next I
                           
                           .CurrentX = x + 78
                           .CurrentY = Y + 114
                           Hasta = (PosicionT2 - 1) - (PosicionT1 - 1)
                           Printer.Print StrConv(Mid(vConcepto, (PosicionT1 + 1), Hasta), vbUpperCase)
                           
                           'Tercer tramo de la cadena
                           .CurrentX = x + 78
                           .CurrentY = Y + 114
                           Hasta = Largo - (PosicionT2)
                           If Hasta > 34 Then Hasta = 34
                           Printer.Print StrConv(Mid(vConcepto, (PosicionT2 + 1), Hasta), vbUpperCase)
                        Else
                           .CurrentX = x + 78
                           .CurrentY = Y + 114
                           Printer.Print StrConv(vConcepto, vbUpperCase)
                       End If
                   End If
            
            End If
            
            .CurrentX = x + 26
            .CurrentY = Y + 192
            .FontBold = True
            .FontSize = 12
            Printer.Print Format(LabelTotalAbonado.Caption, "Standard")
            
        .EndDoc
        
    End With
           
CapturaErrores:
   Select Case Err
        Case 13
          'MsgBox "No hay Pagos Para Liquidar con el Criterio Seleccionado...", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
          Exit Sub
    End Select

End Sub


Private Sub ImprimirReciboE()

        'On Error GoTo CapturaErrores
       
        Dim NroFactura As String
        Dim NroRecibo As Long
        Dim NRec As String
        Dim NroRemito As String
        Dim IdSuc As String
        Dim vSQL As String
        Dim Largo As Integer
        Dim LargoR As Integer
        Dim LargoSuc As Integer
        Dim linea As Integer
        Dim PU, TL, Cant, SubTotalFac, ImpIva, ImpIIBB, TotalFac As Variant
        Dim Original As Integer
        Dim vEfete, vCheques, vRetenciones, vTransf
        
        'Buscar en bbdd
           Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
           
           'Set tFacturaC = BaseSPC.OpenRecordset("FacturaC", dbOpenTable)
           'Set tFacturaD = BaseSPC.OpenRecordset("FacturaD", dbOpenTable)
           Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
           Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
           Set tUltimosNumeros = BaseSPC.OpenRecordset("UltimosNumeros", dbOpenTable)
           
           'tFacturaC.Index = "PrimaryKey"
           tClientes.Index = "PrimaryKey"
           tDomiciliosClientes.Index = "PrimaryKey"
           tUltimosNumeros.Index = "PrimaryKey"
           
           'tFacturaC.MoveFirst
           tUltimosNumeros.MoveFirst
           tClientes.MoveFirst
           tDomiciliosClientes.MoveFirst
           
           'esto es provisorio
              '*******************************************
                'FormFactura.TextTipoFactura.Text = "A"
                'TextNumeroFactura.Text = "14789"
              '*******************************************
           
           'tFacturaC.Seek "=", FormFactura.TextTipoFactura.Text, FormImprimir.TextNumeroFactura.Text
           'tFacturaC.Seek "=", "A", TextNumeroFactura.Text
           'tFacturaC.Seek "=", TextTipoFactura.Text, TextNumeroFactura.Text
           
           tUltimosNumeros.Seek "=", "tRecibosC", CLng(Left(cmbSucursal.text, 1))
           NroRecibo = 0
           If Not tUltimosNumeros.NoMatch Then
             NroRecibo = (CLng(tUltimosNumeros!UltimoNumero))
             NRec = CStr(tUltimosNumeros!UltimoNumero)
            Else
                
           End If
                       
                       
           'If Not tFacturaC.NoMatch Then
                
            '    If IsNull(tFacturaC!CAE) Then
            '        b = MsgBox("No se puede imprimir la Factura no se ha generado el CAE !!!", vbCritical, "E R R O R")
            '        Exit Sub
            '    End If
                
            vCantFac = CInt(InputBox("Ingrese Cantidad de Facturas a Detallar", "DETALLE DE FACTURAS DEL RECIBO"))
            
            If Val(vCantFac) > 6 Then
                A = MsgBox("Corrija la Cantidad de Facturas" + Chr(10) + "No pueden ser más de 6", vbOKOnly, "CANTIDAD DE FACTURAS")
                vCantFac = InputBox("Ingrese Cantidad de Facturas a Detallar", "DETALLE DE FACTURAS DEL RECIBO")
            End If
                
            Dim vFF(6) As Date
            Dim vNFac(6) As String
            Dim vImpF(6) As Double
            
            For J = 1 To vCantFac
                vFF(J) = InputBox("Ingrese La Fecha de La Factura # " & J, "DETALLE DE FACTURAS DEL RECIBO")
                vNFac(J) = InputBox("Ingrese El Nro de La Factura # " & J, "DETALLE DE FACTURAS DEL RECIBO")
                vImpF(J) = InputBox("Ingrese El Importe de La Factura # " & J, "DETALLE DE FACTURAS DEL RECIBO")
            Next J
            
            vConcepto = InputBox("Ingrese El Concepto", "DETALLE DE FACTURAS DEL RECIBO")
                
                
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
                        
                        .CurrentX = 85
                        .CurrentY = 14
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = True
                        Printer.Print "RECIBO OFICIAL"
                        
                        .CurrentX = 15
                        .CurrentY = 2
                        .Font = "Arial"
                        .FontSize = 12
                        .FontBold = False
                        Printer.Print "ORIGINAL"
                        
                      '  .DrawWidth = 5
                      '  Printer.Line (93, 17)-(102, 17)
                      '  Printer.Line (93, 17)-(93, 25)
                      '  Printer.Line (102, 17)-(102, 25)
                      '  Printer.Line (93, 25)-(102, 25)
                        
                        '.CurrentX = 95
                        '.CurrentY = 17
                        '.FontSize = 20
                        'Printer.Print "A"
                        'Printer.Print TextTipoFactura.Text
                        
                        .CurrentX = 94
                        .CurrentY = 23
                        .FontSize = 4
                        .FontBold = True
                        'Printer.Print "Código 01"
                        
                        .FontSize = 12
                        .CurrentY = 9
                        .CurrentX = 150
                        
                        'En el numero de factura poner de la bbdd
                        NroRecibo = CStr(NroRecibo)
                        'Largo = 8 - Len(NroRecibo)
                        Largo = 8 - Len(NRec)
                        For I = 1 To Largo
                            'NroRecibo = "0" & NroRecibo
                            NRec = "0" & NRec
                        Next I
                        
                        IdSuc = CStr(Left(cmbSucursal.text, 1))
                        LargoSuc = 4 - Len(IdSuc)
                        For J = 1 To LargoSuc
                            IdSuc = "0" & IdSuc
                        Next J
                        
                        'Printer.Print "Nº: 000" & CInt(Left(cmbSucursal.text, 1)) & "-" & NroRecibo
                        'Printer.Print IdSuc & "-" & NroRecibo
                        Printer.Print IdSuc & "-" & NRec
                        
                        .CurrentX = 150
                        .CurrentY = .CurrentY + 2
                        .FontSize = 12
                        
                        'En la fecha poner la fecha de la bbdd
                        Printer.Print "Fecha: " & Format(TextFechaPago, "DD/MM/YYYY")
                        
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
                        tClientes.Seek "=", TextCodigoCliente.text
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
                                
                                .CurrentX = 83
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                Printer.Print "*** www.quilplac.com ***"
                                .CurrentX = 55
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                'Printer.Print tFacturaC!CondicionVenta
                            
                                .CurrentX = 130
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = True
                                'Printer.Print "Remito: "
                                
                                .CurrentX = 150
                                .CurrentY = 80
                                .FontSize = 10
                                .FontBold = False
                                
                           '     NroRemito = CStr(tFacturaC!NroRemito)
                           '     LargoR = 8 - Len(tFacturaC!NroRemito)
                           '     For I = 1 To LargoR
                           '         NroRemito = "0" & NroRemito
                           '     Next I
                                
                             '   Printer.Print "0002-" & NroRemito
                        End If
                     
                     'Recuadro Detalle
                        .DrawWidth = 10
                        Printer.Line (10, 90)-(200, 90)
                        Printer.Line (10, 240)-(200, 240)
                        Printer.Line (10, 90)-(10, 240)
                        Printer.Line (200, 90)-(200, 240)
                        Printer.Line (10, 97)-(200, 97)
                        
                       ' .CurrentX = 18
                       ' .CurrentY = 92
                       ' .FontSize = 8
                        .FontBold = True
                       ' Printer.Print "CANTIDAD"
                       ' .DrawWidth = 5
                       ' Printer.Line (40, 91)-(40, 240)
                        
                        
                        .CurrentX = 86
                        .CurrentY = 92
                        .FontSize = 10
                        Printer.Print "DETALLE DEL RECIBO"
                        'Printer.Line (130, 91)-(130, 240)
                        
                       ' .CurrentX = 140
                       ' .CurrentY = 92
                       '.FontSize = 8
                       ' Printer.Print "UNITARIO"
                       ' Printer.Line (165, 91)-(165, 240)
                        
                       ' .CurrentX = 175
                        '.CurrentY = 92
                       ' .FontSize = 8
                       ' Printer.Print "IMPORTE"
                        
                       'Imprimir Detalle del Recibo
                       
                                                    
                            .CurrentX = 32
                            .CurrentY = 110
                          '  .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            vEfete = Val(TextEfectivo.text)
                            Printer.Print "* Efectivo: " & Chr(9) & Chr(9) & Format(vEfete, "Currency")
                            
                            .CurrentX = 32
                            .CurrentY = 120
                          '  .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            vTransf = Val(textTransferencia.text)
                            Printer.Print "* Transferencia: " & Chr(9) & Format(vTransf, "Currency")
                            
                            .CurrentX = 32
                            .CurrentY = 130
                          '  .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            vCheques = Val(TextCheque.text)
                            Printer.Print "* Cheques Varios: " & Chr(9) & Format(vCheques, "Currency")
                        
                            .CurrentX = 32
                            .CurrentY = 140
                          '  .FontName = "Courier New"
                           ' .FontBold = False
                            .FontSize = 10
                            vRetenciones = Val(TextRetencion.text)
                            Printer.Print "* Retenciones: " & Chr(9) & Format(vRetenciones, "Currency")
                            
                            'X = 10
                            'Y = 102
                                                                    
                        'Imprimir Facturas
                            .FontSize = 8
                            
                            .CurrentX = 32
                            .CurrentY = 150
                            .FontUnderline = True
                            Printer.Print "Fecha"
                            
                            .CurrentX = 60
                            .CurrentY = 150
                            .FontUnderline = True
                            Printer.Print "Nro. Factura"
                            
                            .CurrentX = 90
                            .CurrentY = 150
                            .FontUnderline = True
                            Printer.Print "Importe"
                            
                            .FontUnderline = False
                            
                            For J = 1 To vCantFac
                                Select Case J
                                    Case 1
                                        .CurrentX = 32
                                        .CurrentY = 155
                                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                                        
                                        .CurrentX = 60
                                        .CurrentY = 155
                                        Printer.Print vNFac(J)
                                        
                                        .CurrentX = 90
                                        .CurrentY = 155
                                        Printer.Print Format(vImpF(J), "Standard")
                                    Case 2
                                        .CurrentX = 32
                                        .CurrentY = 160
                                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                                        
                                        .CurrentX = 60
                                        .CurrentY = 160
                                        Printer.Print vNFac(J)
                                        
                                        .CurrentX = 90
                                        .CurrentY = 160
                                        Printer.Print Format(Val(vImpF(J)), "##,##")
                                    Case 3
                                        .CurrentX = 32
                                        .CurrentY = 165
                                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                                        
                                        .CurrentX = 60
                                        .CurrentY = 165
                                        Printer.Print vNFac(J)
                                        
                                        .CurrentX = 90
                                        .CurrentY = 165
                                        Printer.Print Format(Val(vImpF(J)), "##,##")
                                    Case 4
                                        .CurrentX = 32
                                        .CurrentY = 170
                                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                                        
                                        .CurrentX = 60
                                        .CurrentY = 170
                                        Printer.Print vNFac(J)
                                        
                                        .CurrentX = 90
                                        .CurrentY = 170
                                        Printer.Print Format(Val(vImpF(J)), "##,##")
                                    Case 5
                                        .CurrentX = 32
                                        .CurrentY = 175
                                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                                        
                                        .CurrentX = 60
                                        .CurrentY = 175
                                        Printer.Print vNFac(J)
                                        
                                        .CurrentX = 90
                                        .CurrentY = 175
                                        Printer.Print Format(Val(vImpF(J)), "##,##")
                                    
                                    Case 6
                                        .CurrentX = 32
                                        .CurrentY = 180
                                        Printer.Print Format(vFF(J), "DD/MM/YYYY")
                                        
                                        .CurrentX = 60
                                        .CurrentY = 180
                                        Printer.Print vNFac(J)
                                        
                                        .CurrentX = 90
                                        .CurrentY = 180
                                        Printer.Print Format(Val(vImpF(J)), "##,##")
                                End Select
                            Next J
                        
                        
                        'Imprime Concepto
                            .CurrentX = 20
                            .CurrentY = 185
                            .Font = "Arial"
                            .FontSize = 10
                            Printer.Print StrConv(vConcepto, vbUpperCase)
                        
                        
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
                           ' Printer.Print ("Sub Total: ")
                            .FontName = "Courier New"
                            'SubTotalFac = CDbl(TextCheque.Text) + CDbl(TextEfectivo.Text)
                            SubTotalFac = Val(TextCheque.text) + Val(TextEfectivo.text)
                            'Format(CDbl(tFacturaC!SubTotalFactura), "Standard")
                            vSubTotal = SubTotalFac
                            Hasta = CInt(14 - Len(SubTotalFac))
                            For I = 0 To Hasta
                                SubTotalFac = " " & SubTotalFac
                            Next I
                            .CurrentX = 165
                            .CurrentY = 245
                            
                         '   Printer.Print Format(SubTotalFac, "Currency")
                        '   Printer.Print Format(vSubTotal, "Currency")
                            
                        'Percepciones
                            .CurrentX = 135
                            .CurrentY = 250
                            .Font = "Arial"
                            .FontSize = 10
                            '.FontBold = False
                         '   Printer.Print "Retenciones: "
                        
                        'Importe Percepciones
                            .CurrentX = 165
                            .CurrentY = 250
                            '.Font = "Courier New"
                            .FontSize = 10
                            '.FontBold = False
                            'ImpIva = Format(CDbl(tFacturaC!totalIva), "Standard")
                            'Hasta = CInt(14 - Len(ImpIva))
                            'For I = 0 To Hasta
                            '    ImpIva = " " & ImpIva
                            'Next I
                            
                            If TextRetencion.text = "" Then TextRetencion.text = 0
                            vRetenciones = Val(TextRetencion.text)
                            'Ret = Format(TextRetencion.Text, "Currency")
                            Ret = Format(vRetenciones, "Currency")
                            'Printer.Print CDbl(Format(TextRetencion.Text, "Currency"))
                          '  Printer.Print Ret
                        
                       ' If tFacturaC!ImportePercepIIBB > 0 Then
                       '     'Alicuota IIBB
                       '         .CurrentX = 135
                       '         .CurrentY = 255
                       '         .Font = "Arial"
                       '         .FontSize = 10
                                '.FontBold = False
                       '         Printer.Print "Per.IIBB: " & tFacturaC!AlicuotaIIBB & "%:"
                            
                            'Importe IIBB
                       '         .CurrentX = 165
                       '         .CurrentY = 255
                       '         .Font = "Courier New"
                       '         .FontSize = 10
                                '.FontBold = False
                       '         ImpIIBB = Format(CDbl(tFacturaC!ImportePercepIIBB), "Standard")
                       '         Hasta = CInt(14 - Len(ImpIIBB))
                       '         For I = 0 To Hasta
                       '             ImpIIBB = " " & ImpIIBB
                       '         Next I
                       '         Printer.Print ImpIIBB
                       ' End If
                        
                        'Importe Total
                            
                            Printer.Line (130, 262)-(200, 270), vbBlack, BF
                            
                            .CurrentX = 135
                            .CurrentY = 264
                            .Font = "Arial"
                            .FontSize = 12
                            '.FontBold = False
                            .ForeColor = vbWhite
                            Printer.Print "TOTAL: "
                            TotalFac = LabelTotalAbonado.Caption
                            'Format(CDbl(tFacturaC!TotalFactura), "Standard")
                            Hasta = CInt(14 - Len(TotalFac))
                            For I = 0 To Hasta
                                TotalFac = " " & TotalFac
                            Next I
                            
                           ' .Font = "Courier New"
                            .Font = "Arial"
                            .FontSize = 12
                            .CurrentX = 165
                            .CurrentY = 264
                            Printer.Print Format(TotalFac, "Currency")
                        
                        '.FontBold = True
                        '.FontName = "Arial"
                        '.ForeColor = vbBlack
                        .FontSize = 10
                        .CurrentX = 15
                        .CurrentY = 245
                        'Printer.Print "C.A.E: " & tFacturaC!CAE
                        '.CurrentX = 15
                        '.CurrentY = 252
                        'Printer.Print "Fecha Vencimiento C.A.E: " & Format(tFacturaC!FechaVC, "DD/MM/YYYY")
                        
                        'Call CrearBarCode(CalcularBarCode)
                        
                        Printer.Line (10, 245)-(48, 250), vbBlack, BF
                        
                        .CurrentX = 12
                        .CurrentY = 245

                        
                        Printer.Print "RECIBIMOS PESOS: "
                        .ForeColor = vbBlack
                        'Imprimir importe en letras en renglones
                            TotalFac = Format(TotalFac, "Fixed")
                            vImporteEnLetras = EnLetras(CStr(TotalFac))
                                            
                        .CurrentX = 12
                        .CurrentY = 253
                            
                            Largo = Len(vImporteEnLetras)
                            I = 0
                            
                            If Len(vImporteEnLetras) <= 50 Then
                                Printer.Print StrConv(vImporteEnLetras, vbUpperCase)
                              Else
                                'Primer Tramo de la cadena
                                     For I = 50 To 1 Step -1
                                         If Mid(vImporteEnLetras, I, 1) = " " Then Exit For
                                     Next I
                                     Printer.Print StrConv(Mid(vImporteEnLetras, 1, I), vbUpperCase)
                                
                                    'Segundo Tramo de la cadena
                                     If (Largo - I) <= 50 Then
                                         .CurrentX = 12
                                         .CurrentY = 258
                                         Printer.Print StrConv(Mid(vImporteEnLetras, (I + 1), (Largo - I)), vbUpperCase)
                                      Else
                                         .CurrentX = 12
                                         .CurrentY = 263
                                         SegundoTramo = StrConv(Mid(vImporteEnLetras, (I + 1), (Largo - I)), vbUpperCase)
                                         Largo = Len(SegundoTramo)
                                         For I = 50 To 1 Step -1
                                             If Mid(SegundoTramo, I, 1) = " " Then
                                                 PosicionT1 = I
                                                 Exit For
                                             End If
                                         Next I
                                         'Segundo tramo de la cadena
                                         Printer.Print StrConv(Mid(SegundoTramo, 1, I), vbUpperCase)
                                         .CurrentX = 12
                                         .CurrentY = 268
                                         'Tercer Tramo de la cadena
                                         Printer.Print StrConv(Mid(SegundoTramo, (I + 1), (Largo - I)), vbUpperCase)
                                     End If
                            End If
                            
                    .EndDoc
                End With
             'Else
                'A = MsgBox("Factura Inexistente !!!", vbCritical, "E R R O R")
        'End If
    
CapturaErrores:

End Sub

Private Function BuscarCondicionIva(CI As String) As String
    
    Set tCondicionIVA = BaseSPC.OpenRecordset("CondicionIVA", dbOpenTable)

    tCondicionIVA.Index = "PrimaryKey"
    
    tCondicionIVA.Seek "=", CI

    If Not tCondicionIVA.NoMatch Then BuscarCondicionIva = tCondicionIVA!Descripcion
    
    tCondicionIVA.Close
    
End Function

Private Sub BotonGuardar_Click()
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoC = db.OpenRecordset("Pagoc", dbOpenDynaset)
    
'    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoD = db.OpenRecordset("PagoD", dbOpenDynaset)
    
'    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
'    Set db = DBEngine.OpenDatabase(ruta)
    Set rstMovimientosCtaCte = db.OpenRecordset("MovimientosCtaCte", dbOpenDynaset)
    
'Agregamos las tablas de recibos
    Set tRecibosC = db.OpenRecordset("RecibosC", dbOpenTable)
    Set tRecibosD = db.OpenRecordset("RecibosD", dbOpenTable)
    
    
    '*** Grabo Pagos
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db1 = DBEngine.OpenDatabase(ruta)
    
    Set rstPagc = db1.OpenRecordset("PagoC", dbOpenTable)
    
    rstPagc.Index = "PrimaryKey"
    
    rstPagc.Seek "=", CDbl(TextNumeroPago.text), CLng(Left(cmbSucursal.text, 1))

    If Not rstPagc.NoMatch Then
        A = MsgBox("Pago Existente", vbCritical, "INFO DEL SISTEMA")
       
        TextNumeroPago.text = num
        TextNumeroPago.SetFocus
    Else
    
        rstPagc.Close
        db1.Close
        
        
        '*** Grabo Cuenta Corriente
        
        CodigoClie = Val(TextCodigoCliente.text)
        rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
            
        rstCtaCte.Edit
        rstCtaCte.Fields!IdCliente = TextCodigoCliente.text
        saldo1 = rstCtaCte.Fields!SaldoL1
        saldo2 = rstCtaCte.Fields!SaldoL2
        
        If saldoLinea1 = 1 Then
            saldoLi1 = LabelTotalAbonado.Caption
            saldoLi1 = saldo1 - saldoLi1
            rstCtaCte.Fields!SaldoL1 = Format(saldoLi1, "#0.00")
            saldoTotalForm = saldoLi1 + saldo2
            rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#0.00")
        End If
        
        If saldoLinea2 = 2 Then
            saldoLi2 = LabelTotalAbonado.Caption
            saldoLi2 = saldo2 - saldoLi2
            rstCtaCte.Fields!SaldoL2 = Format(saldoLi2, "#0.00")
            saldoTotalForm = saldoLi2 + saldo1
            rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#0.00")
        End If
        
        rstCtaCte.Fields!FechaActSaldo = Format(Date, "DD/MM/YYYY")
        rstCtaCte.Update
        
        '*** Grabo Movimientos Cuente corriente
            
        rstMovimientosCtaCte.AddNew
        rstMovimientosCtaCte.Fields!Fecha = TextFechaPago.text
        rstMovimientosCtaCte.Fields!IdCliente = TextCodigoCliente.text
        If saldoLinea1 = 1 Then
            rstMovimientosCtaCte.Fields!tipoDoc = "Pago Linea 1"
            rstMovimientosCtaCte.Fields!ImporteLinea1 = Format(LabelTotalAbonado.Caption, "-#0.00")
            rstMovimientosCtaCte.Fields!ImporteLinea2 = 0
        End If
        If saldoLinea2 = 2 Then
            rstMovimientosCtaCte.Fields!tipoDoc = "Pago Linea 2"
            rstMovimientosCtaCte.Fields!ImporteLinea2 = Format(LabelTotalAbonado.Caption, "-#0.00")
            rstMovimientosCtaCte.Fields!ImporteLinea1 = 0
        End If
        rstMovimientosCtaCte.Fields!NroDoc = TextNumeroPago.text
        
        rstMovimientosCtaCte.Update
                
          
        '*** Grabo Pagos C
        
        rstPagoC.AddNew
            rstPagoC.Fields!IdSucursal = CLng(Left(cmbSucursal.text, 1))
            rstPagoC.Fields!NroPago = TextNumeroPago.text
            rstPagoC.Fields!FechaPago = Format(Date, "dd/mm/yyyy")
            rstPagoC.Fields!IdCliente = TextCodigoCliente.text
            rstPagoC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "#0.00")
            If OptionSaldoLinea1.Value = True Then rstPagoC.Fields!Corresponde = "L1"
            If OptionSaldoLinea2.Value = True Then rstPagoC.Fields!Corresponde = "L2"
            rstPagoC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "#0.00")
        rstPagoC.Update
    
        'Grabo RecibosC 2025-06-29
        tRecibosC.AddNew
            'tRecibosC.Fields!NroPago = NroRecibo
            tRecibosC.Fields!IdSucursal = CLng(Left(cmbSucursal.text, 1))
            tRecibosC.Fields!NroPago = CLng(TextNumeroPago.text)
            tRecibosC.Fields!FechaPago = Format(Date, "dd/mm/yyyy")
            tRecibosC.Fields!IdCliente = TextCodigoCliente.text
            tRecibosC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "#0.00")
            If OptionSaldoLinea1.Value = True Then tRecibosC.Fields!Corresponde = "L1"
            If OptionSaldoLinea2.Value = True Then tRecibosC.Fields!Corresponde = "L2"
            tRecibosC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "#0.00")
        tRecibosC.Update
    
    
        'rstPagoD.AddNew
        NroLinea = 0
        
        If TextEfectivo.text <> "" Then
            rstPagoD.AddNew
                rstPagoD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Efectivo"
                rstPagoD.Fields!ImportePago = Format(Val(TextEfectivo.text), "#0.00")
            rstPagoD.Update
            
            'Agrego guardar lineas en RecibosD 2025-06-25
            tRecibosD.AddNew
                tRecibosD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                tRecibosD.Fields!NroPago = CLng(TextNumeroPago.text)
                tRecibosD.Fields!LineaPago = CInt(NroLinea)
                tRecibosD.Fields!FormaPago = "Efectivo"
                tRecibosD.Fields!ImportePago = Format(Val(TextEfectivo.text), "#0.00")
            tRecibosD.Update
        End If
        
        If textTransferencia.text <> "" Then
            rstPagoD.AddNew
                rstPagoD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Transferencia"
                rstPagoD.Fields!ImportePago = Format(Val(textTransferencia.text), "#0.00")
            rstPagoD.Update
        
            tRecibosD.AddNew
                tRecibosD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                tRecibosD.Fields!NroPago = CLng(TextNumeroPago.text)
                tRecibosD.Fields!LineaPago = CInt(NroLinea)
                tRecibosD.Fields!FormaPago = "Transferencia"
                tRecibosD.Fields!ImportePago = Format(Val(textTransferencia.text), "#0.00")
            tRecibosD.Update
        End If
            
        If TextRezago.text <> "" Then
            rstPagoD.AddNew
                rstPagoD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Rezago"
                rstPagoD.Fields!ImportePago = Format(Val(TextRezago.text), "#0.00")
            rstPagoD.Update
            
            tRecibosD.AddNew
                tRecibosD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                tRecibosD.Fields!NroPago = CLng(TextNumeroPago.text)
                tRecibosD.Fields!LineaPago = CInt(NroLinea)
                tRecibosD.Fields!FormaPago = "Rezago"
                tRecibosD.Fields!ImportePago = Format(Val(TextRezago.text), "#0.00")
            tRecibosD.Update
        End If
                 
        If TextMercaderia.text <> "" Then
            rstPagoD.AddNew
                rstPagoD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Mercaderia"
                rstPagoD.Fields!ImportePago = Format(Val(TextMercaderia.text), "#0.00")
            rstPagoD.Update
        
            tRecibosD.AddNew
                tRecibosD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                tRecibosD.Fields!NroPago = CLng(TextNumeroPago.text)
                tRecibosD.Fields!LineaPago = CInt(NroLinea)
                tRecibosD.Fields!FormaPago = "Mercaderia"
                tRecibosD.Fields!ImportePago = Format(Val(TextMercaderia.text), "#0.00")
            tRecibosD.Update
        End If
        
        If TextCheque.text <> "" Then
            rstPagoD.AddNew
                rstPagoD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Cheque"
                rstPagoD.Fields!ImportePago = Format(Val(TextCheque.text), "#0.00")
            rstPagoD.Update
        
            tRecibosD.AddNew
                tRecibosD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                tRecibosD.Fields!NroPago = CLng(TextNumeroPago.text)
                tRecibosD.Fields!LineaPago = CInt(NroLinea)
                tRecibosD.Fields!FormaPago = "Cheque"
                tRecibosD.Fields!ImportePago = Format(Val(TextCheque.text), "#0.00")
            tRecibosD.Update
        End If
        
        If TextRetencion.text <> "" Then
            rstPagoD.AddNew
                rstPagoD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Retencion"
                rstPagoD.Fields!ImportePago = Format(Val(TextRetencion.text), "#0.00")
            rstPagoD.Update
        
            tRecibosD.AddNew
                tRecibosD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                tRecibosD.Fields!NroPago = CLng(TextNumeroPago.text)
                tRecibosD.Fields!LineaPago = CInt(NroLinea)
                tRecibosD.Fields!FormaPago = "Retencion"
                tRecibosD.Fields!ImportePago = Format(Val(TextRetencion.text), "#0.00")
            tRecibosD.Update
        End If
        
        If TextTarjeta.text <> "" Then
            rstPagoD.AddNew
                rstPagoD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Tarjeta"
                rstPagoD.Fields!ImportePago = Format(Val(TextTarjeta.text), "#0.00")
            rstPagoD.Update
        
            tRecibosD.AddNew
                tRecibosD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                tRecibosD.Fields!NroPago = CLng(TextNumeroPago.text)
                tRecibosD.Fields!LineaPago = CInt(NroLinea)
                tRecibosD.Fields!FormaPago = "Tarjeta"
                tRecibosD.Fields!ImportePago = Format(Val(TextTarjeta.text), "#0.00")
            tRecibosD.Update
        End If
        
        
        If TextObservaciones.text <> "" Then
            rstPagoD.AddNew
                rstPagoD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!observaciones = TextObservaciones.text
            rstPagoD.Update
            
            tRecibosD.AddNew
                tRecibosD.Fields!IdSucursal = CInt(Left(cmbSucursal.text, 1))
                tRecibosD.Fields!NroPago = CLng(TextNumeroPago.text)
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                tRecibosD.Fields!LineaPago = CInt(NroLinea)
                tRecibosD.Fields!observaciones = TextObservaciones.text
            tRecibosD.Update
        End If
    
'    '****** Actualizo ultimo numero pago
'
'    Set db = DBEngine.OpenDatabase(ruta)
'    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
'
'    Dim busco As String
'
'    busco = "tPagosC"
'
'    rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
'    ultimo = rstUltimosNumeros.Fields!UltimoNumero
'
'    If ultimo < Val(TextNumeroPago.text) Then
'        rstUltimosNumeros.Edit
'        'If ultimo < rstUltimosNumeros.Fields!UltimoNumero Then
'             rstUltimosNumeros.Fields!UltimoNumero = TextNumeroPago.text
'        'End If
'        rstUltimosNumeros.Update
'    End If
    
  'Actualizo nuevo numero de Pago y Recibo con sucursal 2025-06-29
      Set rstUltNum = db.OpenRecordset("UltimosNumeros", dbOpenTable)
      
      rstUltNum.Index = "PrimaryKey"
    
      rstUltNum.Seek "=", "tPagosC", CInt(Left(cmbSucursal.text, 1))
        
    If Not rstUltNum.NoMatch Then
        rstUltNum.Edit
            rstUltNum!UltimoNumero = TextNumeroPago.text
            rstUltNum!IdSucursal = CLng(Left(cmbSucursal.text, 1))
        rstUltNum.Update
        
        'Busco Nuevo Número de Recibo 2025-06-28
            rstUltNum.Seek "=", "tRecibosC", CInt(Left(cmbSucursal.text, 1))
            
            If Not rstUltNum.NoMatch Then
                rstUltNum.Edit
                    rstUltNum!UltimoNumero = CLng(TextNumeroPago.text)
                    rstUltNum!IdSucursal = CLng(Left(cmbSucursal.text, 1))
                rstUltNum.Update
            End If
     Else
    End If
    
    saldoLinea1 = 0
    saldoLinea2 = 0
    
    Dim Rta, Rta2
    
    Rta2 = MsgBox("¿Desea Imprimir el recibo?", vbYesNo, "Módulo de Pagos")
    
        If Rta2 = vbYes Then
            
            Rta = MsgBox("Elija Sí Para Imprimir el Recibo Electrónico y NO para el Pre-Impreso", vbYesNoCancel, "Elegir Tipo Recibo")
            
            If Rta = vbYes Then
                Call ImprimirReciboE
             Else
                If Rta = vbNo Then
                    Call ImprimeReciboPreImpreso
                 Else
                    Exit Sub
                End If
            End If
         Else
        End If
    
    Call nuevonumeroPago
    Call blanco
End If
    
    TextCodigoCliente.SetFocus
    MSFlexGrid1.Visible = False

End Sub

Private Sub BotonGuardar_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub BotonSalir_Click()

    Unload FormPagoFacturas
    
End Sub

Private Sub BotonSalir_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub cmbSucursal_Click()
    
    Call nuevonumeroPago

End Sub

Private Sub cmbSucursal_KeyPress(KeyAscii As Integer)

 If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
 If KeyAscii = 44 Then KeyAscii = 46


End Sub


Private Sub cmbSucursal_LostFocus()

    Call nuevonumeroPago

End Sub


Private Sub cmdPrintRecibo_Click()
    
    Dim Rta
    
    Rta = MsgBox("Elija Sí Para Imprimir el Recibo Electrónico y NO para el Pre-Impreso", vbYesNoCancel, "Elegir Tipo Recibo")
    
    If Rta = vbYes Then
        Call ImprimirReciboE
     Else
        If Rta = vbNo Then
            Call ImprimeReciboPreImpreso
         Else
            Exit Sub
        End If
    End If
    
End Sub


Public Function EnLetras(numero As String) As String
    
    Dim b, paso As Integer
    Dim expresion, entero, deci, flag As String
       
    flag = "N"
    For paso = 1 To Len(numero)
        'If Mid(numero, paso, 1) = "." Then
        If Mid(numero, paso, 1) = "," Then
            flag = "S"
        Else
            If flag = "N" Then
                entero = entero + Mid(numero, paso, 1) 'Extae la parte entera del numero
            Else
                deci = deci + Mid(numero, paso, 1) 'Extrae la parte decimal del numero
            End If
        End If
    Next paso
   
    If Len(deci) = 1 Then
        deci = deci & "0"
    End If
   
    flag = "N"
    
    If Val(numero) >= -999999999 And Val(numero) <= 999999999 Then 'si el numero esta dentro de 0 a 999.999.999
        For paso = Len(entero) To 1 Step -1
            b = Len(entero) - (paso - 1)
            Select Case paso
            Case 3, 6, 9
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" And Mid(entero, b + 2, 1) = "0" Then
                            expresion = expresion & "cien "
                        Else
                            expresion = expresion & "ciento "
                        End If
                    Case "2"
                        expresion = expresion & "doscientos "
                    Case "3"
                        expresion = expresion & "trescientos "
                    Case "4"
                        expresion = expresion & "cuatrocientos "
                    Case "5"
                        expresion = expresion & "quinientos "
                    Case "6"
                        expresion = expresion & "seiscientos "
                    Case "7"
                        expresion = expresion & "setecientos "
                    Case "8"
                        expresion = expresion & "ochocientos "
                    Case "9"
                        expresion = expresion & "novecientos "
                End Select
               
            Case 2, 5, 8
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If Mid(entero, b + 1, 1) = "0" Then
                            flag = "S"
                            expresion = expresion & "diez "
                        End If
                        If Mid(entero, b + 1, 1) = "1" Then
                            flag = "S"
                            expresion = expresion & "once "
                        End If
                        If Mid(entero, b + 1, 1) = "2" Then
                            flag = "S"
                            expresion = expresion & "doce "
                        End If
                        If Mid(entero, b + 1, 1) = "3" Then
                            flag = "S"
                            expresion = expresion & "trece "
                        End If
                        If Mid(entero, b + 1, 1) = "4" Then
                            flag = "S"
                            expresion = expresion & "catorce "
                        End If
                        If Mid(entero, b + 1, 1) = "5" Then
                            flag = "S"
                            expresion = expresion & "quince "
                        End If
                        If Mid(entero, b + 1, 1) > "5" Then
                            flag = "N"
                            expresion = expresion & "dieci"
                        End If
               
                    Case "2"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "veinte "
                            flag = "S"
                        Else
                            expresion = expresion & "veinti"
                            flag = "N"
                        End If
                   
                    Case "3"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "treinta "
                            flag = "S"
                        Else
                            expresion = expresion & "treinta y "
                            flag = "N"
                        End If
               
                    Case "4"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cuarenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cuarenta y "
                            flag = "N"
                        End If
               
                    Case "5"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "cincuenta "
                            flag = "S"
                        Else
                            expresion = expresion & "cincuenta y "
                            flag = "N"
                        End If
               
                    Case "6"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "sesenta "
                            flag = "S"
                        Else
                            expresion = expresion & "sesenta y "
                            flag = "N"
                        End If
               
                    Case "7"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "setenta "
                            flag = "S"
                        Else
                            expresion = expresion & "setenta y "
                            flag = "N"
                        End If
               
                    Case "8"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "ochenta "
                            flag = "S"
                        Else
                            expresion = expresion & "ochenta y "
                            flag = "N"
                        End If
               
                    Case "9"
                        If Mid(entero, b + 1, 1) = "0" Then
                            expresion = expresion & "noventa "
                            flag = "S"
                        Else
                            expresion = expresion & "noventa y "
                            flag = "N"
                        End If
                End Select
               
            Case 1, 4, 7
                Select Case Mid(entero, b, 1)
                    Case "1"
                        If flag = "N" Then
                            If paso = 1 Then
                                expresion = expresion & "uno "
                            Else
                                expresion = expresion & "un "
                            End If
                        End If
                    Case "2"
                        If flag = "N" Then
                            expresion = expresion & "dos "
                        End If
                    Case "3"
                        If flag = "N" Then
                            expresion = expresion & "tres "
                        End If
                    Case "4"
                        If flag = "N" Then
                            expresion = expresion & "cuatro "
                        End If
                    Case "5"
                        If flag = "N" Then
                            expresion = expresion & "cinco "
                        End If
                    Case "6"
                        If flag = "N" Then
                            expresion = expresion & "seis "
                        End If
                    Case "7"
                        If flag = "N" Then
                            expresion = expresion & "siete "
                        End If
                    Case "8"
                        If flag = "N" Then
                            expresion = expresion & "ocho "
                        End If
                    Case "9"
                        If flag = "N" Then
                            expresion = expresion & "nueve "
                        End If
                End Select
            End Select
            If paso = 4 Then
                If Mid(entero, 6, 1) <> "0" Or Mid(entero, 5, 1) <> "0" Or Mid(entero, 4, 1) <> "0" Or _
                  (Mid(entero, 6, 1) = "0" And Mid(entero, 5, 1) = "0" And Mid(entero, 4, 1) = "0" And _
                   Len(entero) <= 6) Then
                    expresion = expresion & "mil "
                End If
            End If
            
            If paso = 7 Then
                'MsgBox (Mid(entero, 1, 1))
                If Len(entero) = 7 And Mid(entero, 1, 1) = "1" Then
                    expresion = expresion & "millón "
                Else
                    expresion = expresion & "millones "
                End If
            End If
        Next paso
       
        If deci <> "" Then
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion & "con " & deci & "/100"
            Else
                EnLetras = expresion & "con " & deci & "/100"
            End If
        Else
            If Mid(entero, 1, 1) = "-" Then 'si el numero es negativo
                EnLetras = "menos " & expresion
            Else
                EnLetras = expresion
            End If
        End If
    Else 'si el numero a convertir esta fuera del rango superior e inferior
        EnLetras = ""
    End If
       
End Function

Private Sub Form_Load()

'    FormPagoFacturas.Height = 8625
    FormPagoFacturas.Height = 9135
    'FormPagoFacturas.Width = 8055
    FormPagoFacturas.Width = 8300
    FormPagoFacturas.Top = 1000
    FormPagoFacturas.Left = 1000
    
    TextFechaPago.text = Format(Date, "dd/mm/yyyy")
            
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
        
    Set tSucursales = db.OpenRecordset("Sucursales", dbOpenTable)
    
    tSucursales.MoveFirst
    
    Do
        cmbSucursal.AddItem tSucursales!IdSucursal & " - " & tSucursales!NombreSucursal
        tSucursales.MoveNext
    
    Loop Until tSucursales.EOF
    
    cmbSucursal.ListIndex = 0
    
    tSucursales.Close
    db.Close
              
 '   Call nuevonumeroPago
    
    
    
End Sub

Private Sub buscosaldo()
 
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
   
    CodigoClie = Val(TextCodigoCliente.text)
 
         
    rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
    If rstCtaCte.Fields!IdCliente <> Val(TextCodigoCliente.text) Then
       mensaje = MsgBox("No Posee Cuentas Corientes", vbCritical, "Final de la busqueda")
       TextCodigoCliente.text = ""
       Call blanco
       Call nuevonumeroPago
       TextCodigoCliente.SetFocus
    Else
       TextSaldoLinea1.text = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
       TextSaldoLinea2.text = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
       TextSaldoTotal.text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
    End If
    
  
 End Sub



Private Sub nuevonumeroPago()
    
    'ruta = App.Path & "\DB_SPC_SI.mdb"
    
    'Set db = DBEngine.OpenDatabase(ruta)
    'Set rstPagoC = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)
    
    'rstPagoC.MoveLast
        
    'numeroPago = rstPagoC.Fields!NroPago
    'TextNumeroPago.Text = numeroPago + 1

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstUltNum = db.OpenRecordset("UltimosNumeros", dbOpenTable)
    
    rstUltNum.Index = "PrimaryKey"
    
    rstUltNum.Seek "=", "tPagosC", CInt(Left(cmbSucursal.text, 1))
        
    If Not rstUltNum.NoMatch Then
        numeroPago = rstUltNum!UltimoNumero
        TextNumeroPago.text = numeroPago + 1
        
        'Busco Nuevo Número de Recibo 2025-06-28
        rstUltNum.Seek "=", "tRecibosC", CInt(Left(cmbSucursal.text, 1))
        
        If Not rstUltNum.NoMatch Then
            NroRecibo = rstUltNum!UltimoNumero
            NroRecibo = NroRecibo + 1
        End If
     Else
        A = MsgBox("Error al buscar Último Nro. Contacte al Administrador del Sistema", vbOKOnly, "Mensaje del Sistema")
        
    End If
           
End Sub





Private Sub OptionSaldoLinea1_Click()

    If OptionSaldoLinea1.Value = True Then
        saldoLinea2 = 0
        LabelSaldo.Caption = "Saldo Linea 1"
        BotonGuardar.Enabled = True
        Frame1.Enabled = True
        TextResta.text = ""
        TextResta.text = Format(TextSaldoLinea1.text, "#,###,###,#0.00")
        If Val(TextSaldoTotal.text) > 0 Then
            LabelSaldoTotal.ForeColor = vbRed
            LabelSaldoTotal.Caption = TextSaldoTotal.text
        End If
        If Val(TextSaldoTotal.text) < 0 Then
            LabelSaldoTotal.ForeColor = vbBlue
            LabelSaldoTotal.Caption = TextSaldoTotal.text
        End If
        Call blancoCambioCheck
        saldoLinea1 = 1
        cmbSucursal.ListIndex = 1
    End If
    
End Sub

Private Sub OptionSaldoLinea1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub OptionSaldoLinea2_Click()

    If OptionSaldoLinea2.Value = True Then
        saldoLinea1 = 0
        LabelSaldo.Caption = "Saldo Linea 2"
        BotonGuardar.Enabled = True
        Frame1.Enabled = True
        TextResta.text = ""
        TextResta.text = Format(TextSaldoLinea2.text, "#,###,###,#0.00")
        If Val(TextSaldoTotal.text) > 0 Then
            LabelSaldoTotal.ForeColor = vbRed
            LabelSaldoTotal.Caption = TextSaldoTotal.text
        End If
        If Val(TextSaldoTotal.text) < 0 Then
            LabelSaldoTotal.ForeColor = vbBlue
            LabelSaldoTotal.Caption = TextSaldoTotal.text
        End If
        Call blancoCambioCheck
        saldoLinea2 = 2
        cmbSucursal.ListIndex = 0
    End If
    
End Sub

Private Sub OptionSaldoLinea2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Text2_Change()



End Sub

Private Sub TextCheque_Change()

     If TextCheque.text <> "" Then
        Call calculo
        Call calculoabonado
       Else
        Call calculoresta
        Call calculoabonadoresta
    End If
'
' If TextC.Text = "a" Then
'        Call calculo
'        Call calculoabonado
'        TextC.Text = "b"
'    Else
'        Call calculoresta
'        Call calculoabonadoresta
'        TextC.Text = "a"
'    End If

End Sub


Private Sub TextCheque_GotFocus()
    TextCheque.SelLength = Len(TextCheque.text)
End Sub

Private Sub TextCheque_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If KeyAscii = 44 Then KeyAscii = 46

End Sub



Private Sub Textcod_Change()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
   

    
    CodigoClie = Val(Textcod.text)
 
'    If KeyAscii = 13 Then
        If Textcod.text = "" Then
            Textcod.SetFocus
        Else
            
            rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCtaCte.Fields!IdCliente <> Val(Textcod.text) Then
                mensaje = MsgBox("No Posee Cuentas Corientes", vbCritical, "Final de la busqueda")
                Textcod.text = ""
                Call blanco
                Call nuevonumeroPago
                Textcod.SetFocus
            Else
                LabelSaldoTotal.Caption = ""
                TextSaldoLinea1.text = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
                If saldoLinea1 = 1 Then
                    TextResta.text = Format(TextSaldoLinea1.text, "#0.00")
                End If
                TextSaldoLinea2.text = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
                TextSaldoTotal.text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
                
                If Val(TextSaldoTotal.text) > 0 Then
                    LabelSaldoTotal.ForeColor = vbRed
                    LabelSaldoTotal.Caption = TextSaldoTotal.text
                End If
                If Val(TextSaldoTotal.text) < 0 Then
                    LabelSaldoTotal.ForeColor = vbBlue
                    LabelSaldoTotal.Caption = TextSaldoTotal.text
                End If
               
                'SendKeys "{TAB}"
                 'KeyAscii = 0
                 OptionSaldoLinea1.SetFocus
           End If
        End If
  
TextCodigoCliente.text = Textcod.text
    
  
End Sub


Private Sub TextCodigoCliente_GotFocus()

    TextCodigoCliente.SelLength = Len(TextCodigoCliente.text)

End Sub

Private Sub TextCodigoCliente_KeyPress(KeyAscii As Integer)
  
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
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
                Call blanco
                TextCodigoCliente.SetFocus
            Else
                TextNombre.text = rstCliente.Fields!RazonSocial
                
                vDomicilio = rstCliente.Fields!Domicilio
                vLocalidad = rstCliente.Fields!localidad
                vCondIVA = rstCliente.Fields!condicionIva
                If rstCliente.Fields!CUIT <> "" Then
                    vCUIT = rstCliente.Fields!CUIT
                 Else
                    vCUIT = " "
                End If
       
   
    CodigoClie = Val(TextCodigoCliente.text)
 
'    If KeyAscii = 13 Then
        If TextCodigoCliente.text = "" Then
            TextCodigoCliente.SetFocus
        Else
            
            rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCtaCte.Fields!IdCliente <> Val(TextCodigoCliente.text) Then
                mensaje = MsgBox("No Posee Cuentas Corientes", vbCritical, "Final de la busqueda")
                TextCodigoCliente.text = ""
                Call blanco
                Call nuevonumeroPago
                TextCodigoCliente.SetFocus
            Else
                LabelSaldoTotal.Caption = ""
                TextSaldoLinea1.text = Format(rstCtaCte.Fields!SaldoL1, "#,###,###,#0.00")
                If saldoLinea1 = 1 Then
                    TextResta.text = Format(TextSaldoLinea1.text, "#,###,###,#0.00")
                End If
                TextSaldoLinea2.text = Format(rstCtaCte.Fields!SaldoL2, "#,###,###,#0.00")
                TextSaldoTotal.text = Format(rstCtaCte.Fields!SaldoTotal, "#,###,###,#0.00")
                
                If Val(TextSaldoTotal.text) > 0 Then
                    LabelSaldoTotal.ForeColor = vbRed
                    LabelSaldoTotal.Caption = TextSaldoTotal.text
                End If
                If Val(TextSaldoTotal.text) < 0 Then
                    LabelSaldoTotal.ForeColor = vbBlue
                    LabelSaldoTotal.Caption = TextSaldoTotal.text
                End If
               
                'SendKeys "{TAB}"
                 'KeyAscii = 0
                 cmbSucursal.SetFocus
                 'OptionSaldoLinea1.SetFocus
                 
           End If
        End If
    End If
   End If
   End If
    If KeyAscii = 27 Then
        Unload Me
    End If
    MSFlexGrid1.Visible = False
      
End Sub





Private Sub TextEfectivo_Change()

    If TextEfectivo.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If

'    If TextE.Text = "a" Then
'        Call calculo
'        Call calculoabonado
'        TextE.Text = "b"
'    Else
'        Call calculoresta
'        Call calculoabonadoresta
'        TextE.Text = "a"
'    End If

End Sub

Private Sub TextEfectivo_GotFocus()
    TextEfectivo.SelLength = Len(TextEfectivo.text)
End Sub

Private Sub TextEfectivo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If KeyAscii = 44 Then KeyAscii = 46

End Sub



Private Sub TextFechaPago_GotFocus()
    TextFechaPago.SelLength = Len(TextFechaPago.text)
End Sub

Private Sub TextFechaPago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextMercaderia_Change()

    If TextMercaderia.text <> "" Then
        Call calculo
        Call calculoabonado
     Else
        Call calculoresta
        Call calculoabonadoresta
    End If

'     If TextM.Text = "a" Then
'        Call calculo
'        Call calculoabonado
'        TextM.Text = "b"
'    Else
'        Call calculoresta
'        Call calculoabonadoresta
'        TextM.Text = "a"
'    End If

End Sub
Private Sub blancoCambioCheck()

  
    TextEfectivo.text = ""
    textTransferencia.text = ""
    TextRezago.text = ""
    TextMercaderia.text = ""
    TextObservaciones.text = ""
    TextCheque.text = ""
    TextRetencion.text = ""
    TextTarjeta.text = ""
    
    
    
    

End Sub


Private Sub blanco()

    TextCodigoCliente.text = ""
    TextNombre.text = ""
    TextSaldoTotal.text = 0
    TextSaldoLinea1.text = 0
    TextSaldoLinea2.text = 0
    'TextNumeroPago.Text = 0
    TextEfectivo.text = ""
    textTransferencia.text = ""
    TextRezago.text = ""
    TextMercaderia.text = ""
    TextObservaciones.text = ""
    TextCheque.text = ""
    TextRetencion.text = ""
    TextTarjeta.text = ""
    TextResta.text = ""
    LabelSaldoTotal.Caption = ""
    TextResta.text = ""
    LabelTotalAbonado.Caption = ""
    
    BotonGuardar.Enabled = False
    Frame1.Enabled = False
    OptionSaldoLinea1.Value = False
    OptionSaldoLinea2.Value = False
    
'    TextCodigoCliente.SetFocus
    
End Sub

Private Sub TextMercaderia_GotFocus()
    TextMercaderia.SelLength = Len(TextMercaderia.text)
End Sub

Private Sub TextMercaderia_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If KeyAscii = 44 Then KeyAscii = 46

End Sub

Private Sub TextNombre_Change()
     Columna = 1
     Call FiltrarGrilla(MSFlexGrid1, TextNombre, CLng(Columna))
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
                    
                End With
                linea2 = linea2 + 1
                tClientes.MoveNext
        Loop
    End If
MSFlexGrid1.Col = 4
'MSFlexGrid1.Sort = flexSortGenericAscending


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
    Textcod.text = MSFlexGrid1.text
    
    MSFlexGrid1.Col = 1
    TextNombre.text = MSFlexGrid1.text
    
   
   
    
    MSFlexGrid1.Visible = False
    
   

End Sub




Private Sub TextNumeroPago_GotFocus()
    TextNumeroPago.SelLength = Len(TextNumeroPago.text)
End Sub

Private Sub TextNumeroPago_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub



Private Sub TextObservaciones_GotFocus()
    TextObservaciones.SelLength = Len(TextObservaciones.text)
End Sub

Private Sub TextObservaciones_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextRetencion_Change()

    If TextRetencion.text <> "" Then
        Call calculo
        Call calculoabonado
     Else
        Call calculoresta
        Call calculoabonadoresta
    End If

'     If TextRE.Text = "a" Then
'        Call calculo
'        Call calculoabonado
'        TextRE.Text = "b"
'    Else
'        Call calculoresta
'        Call calculoabonadoresta
'        TextRE.Text = "a"
'    End If
End Sub

Private Sub TextRetencion_GotFocus()
    TextRetencion.SelLength = Len(TextRetencion.text)
End Sub

Private Sub TextRetencion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

    If KeyAscii = 44 Then KeyAscii = 46

End Sub

Private Sub TextRezago_Change()

    If TextRezago.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If

'     If TextR.Text = "a" Then
'        Call calculo
'        Call calculoabonado
'        TextR.Text = "b"
'    Else
'        Call calculoresta
'        Call calculoabonadoresta
'        TextR.Text = "a"
'    End If
End Sub

Private Sub calculo()
       
        
        efectivo = Val(TextEfectivo.text)
        
        If efectivo < 0 Then
            efectivo = 0
        End If
        
        transferencia = Val(textTransferencia.text)
        
        If transferencia < 0 Then
            transferencia = 0
        End If
        
        rezago = Val(TextRezago.text)
        
        If rezago < 0 Then
            rezago = 0
        End If
        
        mercaderia = Val(TextMercaderia.text)
        
        If mercaderia < 0 Then
            mercaderia = 0
        End If
         
        cheque = Val(TextCheque.text)
        
        If cheque < 0 Then
            cheque = 0
        End If
        
        retencion = Val(TextRetencion.text)
        
        If retencion < 0 Then
            retencion = 0
        End If
        
        tarjeta = Val(TextTarjeta.text)
        
        If tarjeta < 0 Then
            tarjeta = 0
        End If
       
       
    'resta = cdec(TextSaldoLinea1.Text) - CDec(TextEfectivo.Text) - CDec(TextRezago.Text) - CDec(TextMercaderia.Text) - CDec(TextCheque.Text) - CDec(TextRetencion.Text) - CDec(TextTarjeta.Text)

    If saldoLinea1 = 1 Then
            resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
    End If
    
    If saldoLinea2 = 2 Then
            resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
    End If
    
    
    'TextResta.text = Format(resta, "#00.00")
    TextResta.text = Format(resta, "#,###,###,#0.00")
    
    LabelSaldoTotal.Caption = Format(resta2, "#,###,###,#0.00")
    
End Sub
Private Sub calculoresta()

        
        efectivo = Val(TextEfectivo.text)
        If efectivo = 0 Then
            efectivo = 0
        End If

        transferencia = Val(textTransferencia.text)
        If transferencia = 0 Then
            transferencia = 0
        End If
       
        rezago = Val(TextRezago.text)
        If rezago = 0 Then
            rezago = 0
        End If

      
        mercaderia = Val(TextMercaderia.text)
        If mercaderia = 0 Then
            mercaderia = 0
        End If

       
        cheque = Val(TextCheque.text)
        If cheque = 0 Then
            cheque = 0
        End If

       
        retencion = Val(TextRetencion.text)
        If retencion = 0 Then
            retencion = 0
        End If

       
        tarjeta = Val(TextTarjeta.text)
        If tarjeta = 0 Then
            tarjeta = 0
        End If
       
    'resta = CDec(TextSaldoLinea1.Text) - CDec(TextEfectivo.Text) - CDec(TextRezago.Text) - CDec(TextMercaderia.Text) - CDec(TextCheque.Text) - CDec(TextRetencion.Text) - CDec(TextTarjeta.Text)

    If saldoLinea1 = 1 Then
        If efectivo = 0 Then
            resta = CDec(TextSaldoLinea1.text) + CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) + CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If transferencia = 0 Then
            resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) + CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) + CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If rezago = 0 Then
             resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(transferencia) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If mercaderia = 0 Then
             resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If cheque = 0 Then
            resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If retencion = 0 Then
            resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
        End If
         If tarjeta = 0 Then
           resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
        End If
    End If
    
    If saldoLinea2 = 2 Then
       If efectivo = 0 Then
            resta = CDec(TextSaldoLinea2.text) + CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) + CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
       If transferencia = 0 Then
            resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) + CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) + CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        
        If rezago = 0 Then
             resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(transferencia) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If mercaderia = 0 Then
             resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If cheque = 0 Then
            resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If retencion = 0 Then
            resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
        End If
         If tarjeta = 0 Then
           resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(transferencia) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
        End If
     
    End If
    
    TextResta.text = Format(resta, "#00.00")
    LabelSaldoTotal.Caption = Format(resta2, "#00.00")

End Sub


Private Sub calculoabonado()

    suma = CDec(efectivo) + CDec(transferencia) + CDec(rezago) + CDec(mercaderia) + CDec(cheque) + CDec(retencion) + CDec(tarjeta)
    'suma = CDec(Format(efectivo, "#00.00")) + CDec(Format(rezago, "#00.00")) + CDec(Format(mercaderia, "#00.00")) + CDec(Format(cheque, "#00.00")) + CDec(Format(retencion, "#00.00")) + CDec(Format(tarjeta, "#00.00"))
   
    LabelTotalAbonado.Caption = Format(suma, "#,###,###,#0.00")
    
End Sub
Private Sub calculoabonadoresta()

    suma = CDec(efectivo) + CDec(transferencia) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
   
    LabelTotalAbonado.Caption = Format(suma, "#00.00")

End Sub

Private Sub TextRezago_GotFocus()
    TextRezago.SelLength = Len(TextRezago.text)
End Sub

Private Sub TextRezago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If
    
    If KeyAscii = 44 Then KeyAscii = 46

End Sub

Private Sub TextTarjeta_Change()

    If TextTarjeta.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If

'     If TextT.Text = "a" Then
'        Call calculo
'        Call calculoabonado
'        TextT.Text = "b"
'    Else
'        Call calculoresta
'        Call calculoabonadoresta
'        TextT.Text = "a"
'    End If
    
End Sub


Private Sub TextTarjeta_GotFocus()
    TextTarjeta.SelLength = Len(TextTarjeta.text)
End Sub

Private Sub TextTarjeta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

    If KeyAscii = 44 Then KeyAscii = 46

End Sub

Private Sub textTransferencia_Change()

    If textTransferencia.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If

End Sub


Private Sub TextTransferencia_GotFocus()

    textTransferencia.SelLength = Len(textTransferencia.text)

End Sub


Private Sub TextTransferencia_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        Sendkeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

    If KeyAscii = 44 Then KeyAscii = 46

End Sub


