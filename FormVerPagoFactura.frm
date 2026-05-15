VERSION 5.00
Begin VB.Form FormVerPagoFacturas 
   Caption         =   "Consulta Pago"
   ClientHeight    =   8625
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8625
   ScaleWidth      =   7950
   StartUpPosition =   3  'Windows Default
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
      TabIndex        =   31
      Top             =   1560
      Width           =   7575
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
         Left            =   4320
         TabIndex        =   38
         Top             =   480
         Visible         =   0   'False
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
         Left            =   1080
         TabIndex        =   32
         Top             =   480
         Visible         =   0   'False
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   9
      Top             =   7320
      Width           =   7575
      Begin VB.CommandButton cmdPrintRecibo 
         Caption         =   "&Imprimir"
         Height          =   750
         Left            =   2040
         TabIndex        =   45
         Top             =   240
         Width           =   870
      End
      Begin VB.CommandButton BotonModificar 
         Caption         =   "&Modificar"
         Height          =   750
         Left            =   2040
         TabIndex        =   40
         Top             =   240
         Visible         =   0   'False
         Width           =   870
      End
      Begin VB.CommandButton BotonGuardar 
         Caption         =   "&Anular"
         Height          =   750
         Left            =   3120
         TabIndex        =   0
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   4080
         TabIndex        =   10
         Top             =   240
         Width           =   750
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forma de Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   7575
      Begin VB.TextBox TextTransferencia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   42
         Text            =   "0"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TextFechaPago 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         TabIndex        =   36
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TextNumeroPago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   24
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextObservaciones 
         Appearance      =   0  'Flat
         Height          =   525
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   21
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox TextTarjeta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   19
         Text            =   "0"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.TextBox TextRetencion 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   17
         Text            =   "0"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox TextRezago 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   3
         Text            =   "0"
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TextMercaderia 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   4
         Text            =   "0"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TextCheque 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Text            =   "0"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TextResta 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         Left            =   6000
         TabIndex        =   15
         Top             =   3480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextEfectivo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         TabIndex        =   2
         Text            =   "0"
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label Label16 
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
         TabIndex        =   41
         Top             =   960
         Width           =   1245
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
         TabIndex        =   37
         Top             =   1320
         Width           =   1035
      End
      Begin VB.Label LabelTotalAbonado 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2280
         TabIndex        =   35
         Top             =   3840
         Width           =   75
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Total Abonado:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   360
         TabIndex        =   34
         Top             =   3840
         Width           =   1620
      End
      Begin VB.Label LabelSaldo 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1080
         TabIndex        =   33
         Top             =   3960
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
         TabIndex        =   23
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
         Left            =   3240
         TabIndex        =   22
         Top             =   1800
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
         TabIndex        =   20
         Top             =   3360
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
         TabIndex        =   18
         Top             =   2880
         Width           =   945
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Resta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3600
         TabIndex        =   16
         Top             =   3600
         Visible         =   0   'False
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
         TabIndex        =   14
         Top             =   2400
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
         TabIndex        =   13
         Top             =   1920
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
         TabIndex        =   12
         Top             =   1440
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
         TabIndex        =   11
         Top             =   480
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
      TabIndex        =   6
      Top             =   120
      Width           =   7575
      Begin VB.ComboBox cmbSucursal 
         Height          =   315
         Left            =   240
         TabIndex        =   43
         Text            =   "Combo1"
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
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
         Left            =   3120
         TabIndex        =   27
         Top             =   720
         Visible         =   0   'False
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
         Left            =   4680
         TabIndex        =   26
         Top             =   720
         Visible         =   0   'False
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
         Left            =   6120
         TabIndex        =   25
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextCodigoCliente 
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
         Left            =   1680
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label17 
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
         Left            =   360
         TabIndex        =   44
         Top             =   480
         Visible         =   0   'False
         Width           =   750
      End
      Begin VB.Label Anulado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   3720
         TabIndex        =   39
         Top             =   600
         Width           =   2175
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
         Left            =   6120
         TabIndex        =   30
         Top             =   480
         Visible         =   0   'False
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
         Left            =   3120
         TabIndex        =   29
         Top             =   480
         Visible         =   0   'False
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
         Left            =   4680
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "N� Cliente"
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
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "FormVerPagoFacturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rstPagoC As DAO.Recordset
Dim rstPagoD As DAO.Recordset
Dim suma As Double
Dim numDoc As Long
Dim tipoDoc As String
Dim vCorresponde As String
Dim idClienteMov As Long

Private Function BuscarCondicionIva(CI As String) As String
    Dim tCondicionIVA As DAO.Recordset
    Set tCondicionIVA = BaseSPC.OpenRecordset("CondicionIVA", dbOpenTable)
    tCondicionIVA.Index = "PrimaryKey"
    tCondicionIVA.Seek "=", CI
    If Not tCondicionIVA.NoMatch Then BuscarCondicionIva = tCondicionIVA!Descripcion
    tCondicionIVA.Close
End Function

Private Sub cmdPrintRecibo_Click()
    If MsgBox("�Desea imprimir este recibo?", vbYesNo, "M�dulo de Pagos") = vbYes Then
        If vCorresponde = "L1" Then
            Call ImprimirReciboE
        Else
            Call ImprimirReciboX
        End If
    End If
End Sub

Private Sub BotonGuardar_Click()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoC = db.OpenRecordset("Pagoc", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoD = db.OpenRecordset("PagoD", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstMovimientosCtaCte = db.OpenRecordset("MovimientosCtaCte", dbOpenDynaset)
    
      
    respuesta = MsgBox("Esta Seguro de Anular el Pago?", vbYesNo, "Pago")
    If respuesta = vbYes Then
    
            '*** Grabo Cuenta Corriente
            
            If OptionSaldoLinea1.Value = True Then saldoLinea1 = 1
            If OptionSaldoLinea2.Value = True Then saldoLinea2 = 1
            
               
            CodigoClie = Val(TextCodigoCliente.text)
            rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
            
            rstCtaCte.Edit
                rstCtaCte.Fields!IdCliente = TextCodigoCliente.text
                saldo1 = rstCtaCte.Fields!SaldoL1
                saldo2 = rstCtaCte.Fields!SaldoL2
                
                If saldoLinea1 = 1 Then
                    saldoLi1 = (LabelTotalAbonado.Caption)
                    saldoLi1 = saldo1 + saldoLi1
                    rstCtaCte.Fields!SaldoL1 = Format(saldoLi1, "#,###,###,#0.00")
                    If saldoLi1 <> 0 Then
                        saldoTotalForm = saldoLi1 + Abs(saldo2)
                    Else
                        saldoTotalForm = saldo2
                    End If
                    rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#,###,###,#0.00")
                    rstCtaCte.Fields!FechaActSaldo = Format(Date, "DD/MM/YYYY")
                End If
                
                If saldoLinea2 = 1 Then
                    saldoLi2 = (LabelTotalAbonado.Caption)
                    saldoLi2 = saldo2 + saldoLi2
                    rstCtaCte.Fields!SaldoL2 = Format(saldoLi2, "#,###,###,#0.00")
                    If saldoLi2 <> 0 Then
                        saldoTotalForm = saldoLi2 + Abs(saldo1)
                    Else
                         saldoTotalForm = saldo1
                    End If
                    rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#,###,###,#0.00")
                    rstCtaCte.Fields!FechaActSaldo = Format(Date, "DD/MM/YYYY")
                End If
            rstCtaCte.Update
        
            
            '*** Grabo Movimientos Cuente corriente
            
            
           
            rstMovimientosCtaCte.AddNew
            'rstMovimientosCtaCte.Fields!Fecha = Format(Date, "dd/mm/yyyy")
            rstMovimientosCtaCte.Fields!Fecha = Format(TextFechaPago.text, "dd/mm/yyyy")
            rstMovimientosCtaCte.Fields!IdCliente = TextCodigoCliente.text
            If OptionSaldoLinea1.Value = True Then
                rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago N� " & TextNumeroPago.text & " Linea 1"
                rstMovimientosCtaCte.Fields!ImporteLinea1 = Format(LabelTotalAbonado.Caption, "#,###,###,#0.00")
                rstMovimientosCtaCte.Fields!ImporteLinea2 = 0
            End If
            If OptionSaldoLinea2.Value = True Then
                rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago N� " & TextNumeroPago.text & " Linea 2"
                rstMovimientosCtaCte.Fields!ImporteLinea2 = Format(LabelTotalAbonado.Caption, "#,###,###,#0.00")
                rstMovimientosCtaCte.Fields!ImporteLinea1 = 0
            End If
            
            rstMovimientosCtaCte.Fields!NroDoc = 99 & TextNumeroPago.text
            
            rstMovimientosCtaCte.Update
                    
              
            '*** Grabo Pagos
            
            rstPagoC.AddNew
                rstPagoC.Fields!NroPago = 99 & TextNumeroPago.text
                rstPagoC.Fields!FechaPago = Format(Date, "dd/mm/yyyy")
                rstPagoC.Fields!IdCliente = TextCodigoCliente.text
                rstPagoC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "-#0.00")
                If OptionSaldoLinea1.Value = True Then rstPagoC.Fields!Corresponde = "L1"
                If OptionSaldoLinea2.Value = True Then rstPagoC.Fields!Corresponde = "L2"
            rstPagoC.Update
            
            CodigoPa = Val(TextNumeroPago.text)
            rstPagoC.FindFirst "NroPago= " + Str(CodigoPa)
            rstPagoC.Edit
                rstPagoC.Fields!Anulado = "si"
            rstPagoC.Update
        
            'rstPagoD.AddNew
            NroLinea = 0
            
            If TextEfectivo.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Efectivo"
                rstPagoD.Fields!ImportePago = Format(TextEfectivo.text, "-#0.00")
                rstPagoD.Update
            End If
                
            If TextRezago.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Rezago"
                rstPagoD.Fields!ImportePago = Format(TextRezago.text, "-#0.00")
                rstPagoD.Update
            End If
                     
            If TextMercaderia.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Mercaderia"
                rstPagoD.Fields!ImportePago = Format(TextMercaderia.text, "-#0.00")
                rstPagoD.Update
            End If
            
            If TextCheque.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Cheque"
                rstPagoD.Fields!ImportePago = Format(TextCheque.text, "-#0.00")
                rstPagoD.Update
            End If
            
            If TextRetencion.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Retencion"
                rstPagoD.Fields!ImportePago = Format(TextRetencion.text, "-#0.00")
                rstPagoD.Update
            End If
            
            If TextTarjeta.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Tarjeta"
                rstPagoD.Fields!ImportePago = Format(TextTarjeta.text, "-#0.00")
                rstPagoD.Update
            End If
            
           
            
            saldoLinea1 = 0
            saldoLinea2 = 0
            
           
            Call blanco
            Call FormMovimientosCuentaCorriente.BotonBuscar_Click
            Unload FormVerPagoFacturas
       Else
            Call FormMovimientosCuentaCorriente.BotonBuscar_Click
            Unload FormVerPagoFacturas
       End If

End Sub

Private Sub blanco()

    TextCodigoCliente.text = ""
    TextSaldoTotal.text = 0
    TextSaldoLinea1.text = 0
    TextSaldoLinea2.text = 0
    'TextNumeroPago.Text = 0
    TextEfectivo.text = ""
    TextTransferencia.text = ""
    TextRezago.text = ""
    TextMercaderia.text = ""
    TextCheque.text = ""
    TextRetencion.text = ""
    TextTarjeta.text = ""
    TextNumeroPago.text = ""
    TextFechaPago.text = ""
    LabelTotalAbonado.Caption = ""
    
    BotonGuardar.Enabled = False
    Frame1.Enabled = False
    OptionSaldoLinea1.Value = False
    OptionSaldoLinea2.Value = False
    
'    TextCodigoCliente.SetFocus
    
End Sub


Private Sub BotonModificar_Click()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoC = db.OpenRecordset("Pagoc", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoD = db.OpenRecordset("PagoD", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstMovimientosCtaCte = db.OpenRecordset("MovimientosCtaCte", dbOpenDynaset)
    
      
    respuesta = MsgBox("Esta Seguro de Anular el Pago?", vbYesNo, "Pago")
    If respuesta = vbYes Then
    
            '*****************************************
            '***            ANULO PAGO             ***
            '*****************************************
            
            '*** Grabo Cuenta Corriente
               
            CodigoClie = Val(TextCodigoCliente.text)
            rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
            
            rstCtaCte.Edit
            rstCtaCte.Fields!IdCliente = TextCodigoCliente.text
            saldo1 = rstCtaCte.Fields!SaldoL1
            saldo2 = rstCtaCte.Fields!SaldoL2
            
            If saldoLinea1 = 1 Then
                saldoLi1 = LabelTotalAbonado.Caption
                saldoLi1 = saldo1 + saldoLi1
                rstCtaCte.Fields!SaldoL1 = Format(saldoLi1, "#,###,###,#0.00")
                If saldoLi1 <> 0 Then
                    saldoTotalForm = saldoLi1 - saldo2
                Else
                    saldoTotalForm = saldo2
                End If
                rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#,###,###,#0.00")
            End If
            
            If saldoLinea2 = 2 Then
                saldoLi2 = LabelTotalAbonado.Caption
                saldoLi2 = saldo2 + saldoLi2
                rstCtaCte.Fields!SaldoL2 = Format(saldoLi2, "#,###,###,#0.00")
                If saldoLi2 <> 0 Then
                    saldoTotalForm = saldoLi2 - saldo1
                Else
                     saldoTotalForm = saldo1
                End If
                rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#,###,###,#0.00")
            End If
            
            rstCtaCte.Update
            
            
            '*****************************************
            '***        ACTUALIZO PAGO             ***
            '*****************************************
            
            
            
            
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
                rstCtaCte.Fields!SaldoL1 = Format(saldoLi1, "#,###,###,#0.00")
                saldoTotalForm = saldoLi1 + saldo2
                rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#,###,###,#0.00")
            End If
            
            If saldoLinea2 = 2 Then
                saldoLi2 = LabelTotalAbonado.Caption
                saldoLi2 = saldo2 - saldoLi2
                rstCtaCte.Fields!SaldoL2 = Format(saldoLi2, "#,###,###,#0.00")
                saldoTotalForm = saldoLi2 + saldo1
                rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#,###,###,#0.00")
            End If
            
            rstCtaCte.Update
            
            '*** Grabo Movimientos Cuente corriente
                
                
            '****************EDIT
            
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
                    
              
            '*** Grabo Pagos
            
            rstPagoC.AddNew
            rstPagoC.Fields!NroPago = TextNumeroPago.text
            rstPagoC.Fields!FechaPago = Format(Date, "dd/mm/yyyy")
            rstPagoC.Fields!IdCliente = TextCodigoCliente.text
            rstPagoC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "#,###,###,#0.00")
            rstPagoC.Update
        
            'rstPagoD.AddNew
            NroLinea = 0
            
            If TextEfectivo.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Efectivo"
                rstPagoD.Fields!ImportePago = TextEfectivo.text
                rstPagoD.Update
            End If
            
            If TextTransferencia.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Transferencia"
                rstPagoD.Fields!ImportePago = TextTransferencia.text
                rstPagoD.Update
            End If
                
            If TextRezago.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Rezago"
                rstPagoD.Fields!ImportePago = TextRezago.text
                rstPagoD.Update
            End If
                     
            If TextMercaderia.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Mercaderia"
                rstPagoD.Fields!ImportePago = TextMercaderia.text
                rstPagoD.Update
            End If
            
            If TextCheque.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Cheque"
                rstPagoD.Fields!ImportePago = TextCheque.text
                rstPagoD.Update
            End If
            
            If TextRetencion.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Retencion"
                rstPagoD.Fields!ImportePago = TextRetencion.text
                rstPagoD.Update
            End If
            
            If TextTarjeta.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!FormaPago = "Tarjeta"
                rstPagoD.Fields!ImportePago = TextTarjeta.text
                rstPagoD.Update
            End If
            
            
            If TextObservaciones.text <> "" Then
                rstPagoD.AddNew
                rstPagoD.Fields!NroPago = TextNumeroPago.text
                If NroLinea >= 0 Then NroLinea = NroLinea + 1
                rstPagoD.Fields!LineaPago = CInt(NroLinea)
                rstPagoD.Fields!observaciones = TextObservaciones.text
                rstPagoD.Update
            End If
        
            
            '*** Grabo Movimientos Cuente corriente
            
            
           
'            rstMovimientosCtaCte.AddNew
'            rstMovimientosCtaCte.Fields!Fecha = Format(Date, "dd/mm/yyyy")
'            rstMovimientosCtaCte.Fields!idcliente = TextCodigoCliente.Text
'            If OptionSaldoLinea1.Value = True Then
'                rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago N� " & TextNumeroPago.Text & " Linea 1"
'                rstMovimientosCtaCte.Fields!ImporteLinea1 = Format(LabelTotalAbonado.Caption, "#,###,###,#0.00")
'                rstMovimientosCtaCte.Fields!ImporteLinea2 = 0
'            End If
'            If OptionSaldoLinea2.Value = True Then
'                rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago N� " & TextNumeroPago.Text & " Linea 2"
'                rstMovimientosCtaCte.Fields!ImporteLinea2 = Format(LabelTotalAbonado.Caption, "#,###,###,#0.00")
'                rstMovimientosCtaCte.Fields!ImporteLinea1 = 0
'            End If
'
'            rstMovimientosCtaCte.Fields!NroDoc = 99 & TextNumeroPago.Text
'
'            rstMovimientosCtaCte.Update
                    
                    
                    
                    
              
'            '*** Grabo Pagos
'
'            rstPagoC.AddNew
'            rstPagoC.Fields!NroPago = 99 & TextNumeroPago.Text
'            rstPagoC.Fields!FechaPago = Format(Date, "dd/mm/yyyy")
'            rstPagoC.Fields!idcliente = TextCodigoCliente.Text
'            rstPagoC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "-#0.00")
'            rstPagoC.Update
'
'            CodigoPa = Val(TextNumeroPago.Text)
'            rstPagoC.FindFirst "NroPago= " + Str(CodigoPa)
'            rstPagoC.Edit
'            rstPagoC.Fields!Anulado = "si"
'            rstPagoC.Update
'
'            'rstPagoD.AddNew
'            NroLinea = 0
'
'            If TextEfectivo.Text <> "" Then
'                rstPagoD.AddNew
'                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
'                If NroLinea >= 0 Then NroLinea = NroLinea + 1
'                rstPagoD.Fields!LineaPago = CInt(NroLinea)
'                rstPagoD.Fields!FormaPago = "Efectivo"
'                rstPagoD.Fields!ImportePago = Format(TextEfectivo.Text, "-#0.00")
'                rstPagoD.Update
'            End If
'
'            If TextRezago.Text <> "" Then
'                rstPagoD.AddNew
'                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
'                If NroLinea >= 0 Then NroLinea = NroLinea + 1
'                rstPagoD.Fields!LineaPago = CInt(NroLinea)
'                rstPagoD.Fields!FormaPago = "Rezago"
'                rstPagoD.Fields!ImportePago = Format(TextRezago.Text, "-#0.00")
'                rstPagoD.Update
'            End If
'
'            If TextMercaderia.Text <> "" Then
'                rstPagoD.AddNew
'                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
'                If NroLinea >= 0 Then NroLinea = NroLinea + 1
'                rstPagoD.Fields!LineaPago = CInt(NroLinea)
'                rstPagoD.Fields!FormaPago = "Mercaderia"
'                rstPagoD.Fields!ImportePago = Format(TextMercaderia.Text, "-#0.00")
'                rstPagoD.Update
'            End If
'
'            If TextCheque.Text <> "" Then
'                rstPagoD.AddNew
'                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
'                If NroLinea >= 0 Then NroLinea = NroLinea + 1
'                rstPagoD.Fields!LineaPago = CInt(NroLinea)
'                rstPagoD.Fields!FormaPago = "Cheque"
'                rstPagoD.Fields!ImportePago = Format(TextCheque.Text, "-#0.00")
'                rstPagoD.Update
'            End If
'
'            If TextRetencion.Text <> "" Then
'                rstPagoD.AddNew
'                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
'                If NroLinea >= 0 Then NroLinea = NroLinea + 1
'                rstPagoD.Fields!LineaPago = CInt(NroLinea)
'                rstPagoD.Fields!FormaPago = "Retencion"
'                rstPagoD.Fields!ImportePago = Format(TextRetencion.Text, "-#0.00")
'                rstPagoD.Update
'            End If
'
'            If TextTarjeta.Text <> "" Then
'                rstPagoD.AddNew
'                rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
'                If NroLinea >= 0 Then NroLinea = NroLinea + 1
'                rstPagoD.Fields!LineaPago = CInt(NroLinea)
'                rstPagoD.Fields!FormaPago = "Tarjeta"
'                rstPagoD.Fields!ImportePago = Format(TextTarjeta.Text, "-#0.00")
'                rstPagoD.Update
'            End If
'
'
'
'            saldoLinea1 = 0
'            saldoLinea2 = 0
            
           
            Call blanco
            Unload FormVerPagoFacturas
       Else
            Unload FormVerPagoFacturas
       End If

    
End Sub

Private Sub BotonSalir_Click()

    Unload FormVerPagoFacturas
    
End Sub


Private Sub Form_Load()

    'FormVerPagoFacturas.Height = 8625
    FormVerPagoFacturas.Height = 9210
    FormVerPagoFacturas.Width = 8055
    FormVerPagoFacturas.Top = 1000
    FormVerPagoFacturas.Left = 1000
    
    numDoc = FormMovimientosCuentaCorriente.TextNumeroDocumento
    tipoDoc = FormMovimientosCuentaCorriente.TextTipodocumento
    idClienteMov = Val(FormMovimientosCuentaCorriente.TextCodigoCliente.text)

    ruta = App.Path & "\DB_SPC_SI.mdb"

    Set db = DBEngine.OpenDatabase(ruta)

    Set tSucursales = db.OpenRecordset("Sucursales", dbOpenTable)

    tSucursales.MoveFirst

    Do Until tSucursales.EOF
        cmbSucursal.AddItem (tSucursales!IdSucursal & " - " & tSucursales!NombreSucursal)
        tSucursales.MoveNext
    Loop

    cmbSucursal.ListIndex = 1

    tSucursales.Close
    db.Close
   
    Call buscodatos
    
End Sub
Private Sub buscodatos_old()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    
        Set db = DBEngine.OpenDatabase(ruta)
        Set rstPagoC = db.OpenRecordset("Pagoc", dbOpenDynaset)
        'Set rstPagoC = db.OpenRecordset("Pagoc", dbOpenTable)
        
        'Set db = DBEngine.OpenDatabase(ruta)
        Set rstPagoD = db.OpenRecordset("PagoD", dbOpenDynaset)
        'Set rstPagoD = db.OpenRecordset("PagoD", dbOpenTable)
        
        'rstPagoC.Index = "PrimaryKey"
        'rstPagoD.Index = "PrimaryKey"
        
        rstPagoC.FindFirst "NroPago= " + Str(numDoc)
        
        'rstPagoC.Seek "=", cmbSucursal.text, CLng(numDoc)
        
         
                TextCodigoCliente.text = rstPagoC.Fields!IdCliente
                LabelTotalAbonado.Caption = rstPagoC.Fields!TotalAbonado
                
                TextNumeroPago.text = rstPagoC.Fields!NroPago
                TextFechaPago.text = rstPagoC.Fields!FechaPago
                
                rstPagoD.FindFirst "NroPago= " + Str(numDoc)
            
                
                Do While Not rstPagoD.NoMatch
                  
                    If rstPagoD.Fields!FormaPago = "Efectivo" Then
                        TextEfectivo.text = rstPagoD.Fields!ImportePago
                    End If
                    If rstPagoD.Fields!FormaPago = "Transferencia" Then
                        TextTransferencia.text = rstPagoD.Fields!ImportePago
                    End If
                    If rstPagoD.Fields!FormaPago = "Rezago" Then
                        TextRezago.text = rstPagoD.Fields!ImportePago
                    End If
                    If rstPagoD.Fields!FormaPago = "Mercaderia" Then
                        TextMercaderia.text = rstPagoD.Fields!ImportePago
                    End If
                    If rstPagoD.Fields!FormaPago = "Cheque" Then
                        TextCheque.text = rstPagoD.Fields!ImportePago
                    End If
                    If rstPagoD.Fields!FormaPago = "Retencion" Then
                        TextRetencion.text = rstPagoD.Fields!ImportePago
                    End If
                    If rstPagoD.Fields!FormaPago = "Tarjeta" Then
                        TextTarjeta.text = rstPagoD.Fields!ImportePago
                    End If
                    If rstPagoD.Fields!observaciones <> "" Then
                        TextObservaciones.text = rstPagoD.Fields!observaciones
                    End If
                    
                    rstPagoD.FindNext "NroPago= " + Str(numDoc)
                    
                Loop
       
    
                If tipoDoc = "Pago Linea 1" Then
                    OptionSaldoLinea1.Visible = True
                    OptionSaldoLinea2.Visible = True
                    OptionSaldoLinea1.Value = True
                    
                End If
                
                If tipoDoc = "Pago Linea 2" Then
                    OptionSaldoLinea1.Visible = True
                    OptionSaldoLinea2.Visible = True
                    OptionSaldoLinea2.Value = True
                End If
                
   If rstPagoC.Fields!Anulado = "si" Then
      'A = MsgBox("Pago Anulado", vbOKOnly, "INFO DEL SISTEMA")
       Anulado.Caption = "PAGO ANULADO"
       BotonModificar.Visible = False
       BotonGuardar.Visible = False
       OptionSaldoLinea1.Enabled = False
       OptionSaldoLinea1.Enabled = False
       TextEfectivo.Enabled = False
       TextTransferencia.Enabled = False
       TextRezago.Enabled = False
       TextMercaderia.Enabled = False
       TextCheque.Enabled = False
       TextRetencion.Enabled = False
       TextTarjeta.Enabled = False
       TextNumeroPago.Enabled = False
       TextFechaPago.Enabled = False
       TextObservaciones.Enabled = False
   End If
End Sub

Private Sub buscodatos()
    Dim ruta As String
    ruta = App.Path & "\DB_SPC_SI.mdb"

    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoC = db.OpenRecordset("Pagoc", dbOpenDynaset)
    Set rstPagoD = db.OpenRecordset("PagoD", dbOpenDynaset)

    ' Filtrar por NroPago Y IdCliente para evitar traer pagos de otros clientes con mismo número
    rstPagoC.FindFirst "NroPago= " + Str(numDoc) + " AND IdCliente= " + Str(idClienteMov)
    If rstPagoC.NoMatch Then Exit Sub

    TextCodigoCliente.text = rstPagoC.Fields!IdCliente
    LabelTotalAbonado.Caption = rstPagoC.Fields!TotalAbonado
    TextNumeroPago.text = rstPagoC.Fields!NroPago
    TextFechaPago.text = rstPagoC.Fields!FechaPago
    '*** NUEVO: Cargar campo Corresponde para tipo de recibo
    If IsNull(rstPagoC.Fields!Corresponde) Then
        vCorresponde = "L1"
    Else
        vCorresponde = rstPagoC.Fields!Corresponde
    End If
    '*** NUEVO: sucursal persistida
    Dim idSucPago As Long, I As Integer
    If IsNull(rstPagoC.Fields!IdSucursal) Then
        idSucPago = 1
    Else
        idSucPago = CLng(rstPagoC.Fields!IdSucursal)
    End If
    For I = 0 To cmbSucursal.ListCount - 1
        If CLng(Left(cmbSucursal.List(I), 1)) = idSucPago Then
            cmbSucursal.ListIndex = I
            Exit For
        End If
    Next I

    ' Filtrar por NroPago Y IdSucursal para evitar traer detalles de otros pagos con mismo número
    rstPagoD.FindFirst "NroPago= " + Str(numDoc) + " AND IdSucursal= " + Str(idSucPago)
    Do While Not rstPagoD.NoMatch
        Select Case rstPagoD.Fields!FormaPago
            Case "Efectivo":      TextEfectivo.text = rstPagoD.Fields!ImportePago
            Case "Transferencia": TextTransferencia.text = rstPagoD.Fields!ImportePago
            Case "Rezago":        TextRezago.text = rstPagoD.Fields!ImportePago
            Case "Mercaderia":    TextMercaderia.text = rstPagoD.Fields!ImportePago
            Case "Cheque":        TextCheque.text = rstPagoD.Fields!ImportePago
            Case "Retencion":     TextRetencion.text = rstPagoD.Fields!ImportePago
            Case "Tarjeta":       TextTarjeta.text = rstPagoD.Fields!ImportePago
        End Select
        If Not IsNull(rstPagoD.Fields!observaciones) Then
            If rstPagoD.Fields!observaciones <> "" Then
                TextObservaciones.text = rstPagoD.Fields!observaciones
            End If
        End If
        rstPagoD.FindNext "NroPago= " + Str(numDoc) + " AND IdSucursal= " + Str(idSucPago)
    Loop

    If tipoDoc = "Pago Linea 1" Then
        OptionSaldoLinea1.Visible = True
        OptionSaldoLinea2.Visible = True
        OptionSaldoLinea1.Value = True
    End If
    If tipoDoc = "Pago Linea 2" Then
        OptionSaldoLinea1.Visible = True
        OptionSaldoLinea2.Visible = True
        OptionSaldoLinea2.Value = True
    End If

    If rstPagoC.Fields!Anulado = "si" Then
        Anulado.Caption = "PAGO ANULADO"
        BotonModificar.Visible = False
        BotonGuardar.Visible = False
        OptionSaldoLinea1.Enabled = False
        OptionSaldoLinea2.Enabled = False
        TextEfectivo.Enabled = False
        TextTransferencia.Enabled = False
        TextRezago.Enabled = False
        TextMercaderia.Enabled = False
        TextCheque.Enabled = False
        TextRetencion.Enabled = False
        TextTarjeta.Enabled = False
        TextNumeroPago.Enabled = False
        TextFechaPago.Enabled = False
        TextObservaciones.Enabled = False
    End If
End Sub
Private Sub TextCheque_Change()

     If TextCheque.text <> "" Then
        Call calculoabonado
    End If

End Sub

Private Sub ImprimirReciboE()

    On Error GoTo CapturaErrores

    Dim BaseSPC As DAO.Database
    Dim tClientes As DAO.Recordset
    Dim tDomiciliosClientes As DAO.Recordset

    Dim Nrec As String, IdSuc As String
    Dim Largo As Integer, LargoSuc As Integer
    Dim I As Integer, Hasta As Integer
    Dim TotalFac As Variant
    Dim vEfete As Variant, vCheques As Variant
    Dim vRetenciones As Variant, vTransf As Variant
    Dim vImporteEnLetras As String, SegundoTramo As String
    Dim CUIT As String

    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
    Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
    tClientes.Index = "PrimaryKey"
    tDomiciliosClientes.Index = "PrimaryKey"

    Nrec = CStr(TextNumeroPago.text)

    With Printer
        For I = 0 To Printers.Count - 1
            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
        Next I

        .ScaleHeight = 297
        .ScaleWidth = 210

        'Logo sin PictureQP
        Printer.PaintPicture LoadPicture(App.Path & "\Quilplac.JPG"), 10, 9, 40, 10

        .FontItalic = False
        .DrawWidth = 10
        Printer.Line (10, 7)-(200, 7)

        .CurrentX = 85: .CurrentY = 14
        .Font = "Arial": .FontSize = 12: .FontBold = True
        Printer.Print "RECIBO OFICIAL"

        .CurrentX = 15: .CurrentY = 2
        .FontSize = 12: .FontBold = False
        Printer.Print "ORIGINAL"

        'N?mero de recibo: NroPago + IdSucursal del combo
        .FontSize = 12: .CurrentY = 9: .CurrentX = 150

        Largo = 8 - Len(Nrec)
        For I = 1 To Largo
            Nrec = "0" & Nrec
        Next I

        IdSuc = CStr(Left(cmbSucursal.text, 1))
        LargoSuc = 4 - Len(IdSuc)
        For I = 1 To LargoSuc
            IdSuc = "0" & IdSuc
        Next I

        Printer.Print IdSuc & "-" & Nrec

        .CurrentX = 150: .CurrentY = .CurrentY + 2
        .FontSize = 12
        Printer.Print "Fecha: " & Format(TextFechaPago, "DD/MM/YYYY")

        .CurrentX = 150: .CurrentY = .CurrentY + 2
        .FontSize = 9: .FontBold = False
        Printer.Print "C.U.I.T N? 30-70843254-3"
        .CurrentX = 150: Printer.Print "Ing.Brutos N� 30-70843254-3"
        .CurrentX = 150: Printer.Print "Inicio de Actividades: 11-06-2003"
        .CurrentX = 150: Printer.Print "I.V.A. Responsable Inscripto"

        .DrawWidth = 10
        Printer.Line (10, 42)-(200, 42)

        'Datos empresa
        .CurrentX = 12: .CurrentY = 20
        .Font = "Arial": .FontSize = 10: .FontBold = True: .FontUnderline = False
        Printer.Print "QUILPLAC S.A."
        .CurrentX = 12: Printer.Print "Andr�s Baranda 520 - CP (1878) - Quilmes"
        .CurrentX = 12: Printer.Print "Pcia. Buenos Aires"
        .CurrentX = 12: Printer.Print "Tel. 4257-5875"

        'Recuadro cliente
        .DrawWidth = 10
        Printer.Line (10, 47)-(200, 47)
        Printer.Line (10, 47)-(10, 75)
        Printer.Line (200, 47)-(200, 75)
        Printer.Line (10, 75)-(200, 75)

        tClientes.MoveFirst
        tClientes.Seek "=", TextCodigoCliente.text
        If Not tClientes.NoMatch Then
            .CurrentX = 15: .CurrentY = 48: .FontSize = 10: .FontBold = True
            Printer.Print "Se�or(es): "
            .CurrentX = 35: .CurrentY = 48: .FontBold = False
            Printer.Print tClientes!RazonSocial

            .CurrentX = 130: .CurrentY = 48: .FontBold = True
            Printer.Print "C.U.I.T N�:"
            .CurrentX = 150: .CurrentY = 48: .FontBold = False
            CUIT = Left(tClientes!CUIT, 2) & "-" & Mid(tClientes!CUIT, 3, 8) & "-" & Right(tClientes!CUIT, 1)
            Printer.Print CUIT

            tDomiciliosClientes.Seek "=", tClientes!IdCliente
            If Not tDomiciliosClientes.NoMatch Then
                .CurrentX = 15: .CurrentY = 55: .FontBold = True
                Printer.Print "Domicilio: "
                .CurrentX = 35: .CurrentY = 55: .FontBold = False
                Printer.Print tDomiciliosClientes!Domicilio

                .CurrentX = 15: .CurrentY = 62: .FontBold = True
                Printer.Print "Localidad: "
                .CurrentX = 35: .CurrentY = 62: .FontBold = False
                Printer.Print tDomiciliosClientes!localidad

                .CurrentX = 130: .CurrentY = 62: .FontBold = True
                Printer.Print "Tel�fono: "
                .CurrentX = 150: .CurrentY = 62: .FontBold = False
                Printer.Print tClientes!Tel

                .CurrentX = 15: .CurrentY = 69: .FontBold = True
                Printer.Print "I.V.A: "
                .CurrentX = 35: .CurrentY = 69: .FontBold = False
                Printer.Print BuscarCondicionIva(tClientes!condicionIva)
            End If

            .DrawWidth = 10
            Printer.Line (10, 78)-(200, 78)
            Printer.Line (10, 78)-(10, 85)
            Printer.Line (200, 78)-(200, 85)
            Printer.Line (10, 85)-(200, 85)

            .CurrentX = 83: .CurrentY = 80: .FontSize = 10: .FontBold = True
            Printer.Print "*** www.quilplac.com ***"
        End If

        'Recuadro detalle
        .DrawWidth = 10
        Printer.Line (10, 90)-(200, 90)
        Printer.Line (10, 240)-(200, 240)
        Printer.Line (10, 90)-(10, 240)
        Printer.Line (200, 90)-(200, 240)
        Printer.Line (10, 97)-(200, 97)

        .FontBold = True
        .CurrentX = 86: .CurrentY = 92: .FontSize = 10
        Printer.Print "DETALLE DEL RECIBO"

        'Formas de pago (ya cargadas en el form desde PagoD)
        .FontBold = False
        .CurrentX = 32: .CurrentY = 110: .FontSize = 10
        vEfete = Val(TextEfectivo.text)
        Printer.Print "* Efectivo: " & Chr(9) & Chr(9) & Format(vEfete, "Currency")

        .CurrentX = 32: .CurrentY = 120
        vTransf = Val(TextTransferencia.text)
        Printer.Print "* Transferencia: " & Chr(9) & Format(vTransf, "Currency")

        .CurrentX = 32: .CurrentY = 130
        vCheques = Val(TextCheque.text)
        Printer.Print "* Cheques Varios: " & Chr(9) & Format(vCheques, "Currency")

        .CurrentX = 32: .CurrentY = 140
        vRetenciones = Val(TextRetencion.text)
        Printer.Print "* Retenciones: " & Chr(9) & Format(vRetenciones, "Currency")

        'Observaciones si hay
        If TextObservaciones.text <> "" Then
            .CurrentX = 20: .CurrentY = 185
            .Font = "Arial": .FontSize = 10
            Printer.Print "Obs.: " & StrConv(TextObservaciones.text, vbUpperCase)
        End If

        'Recuadro total
        Printer.Line (130, 240)-(130, 262)
        Printer.Line (200, 240)-(200, 262)

        Printer.Line (130, 262)-(200, 270), vbBlack, BF

        .CurrentX = 135: .CurrentY = 264
        .Font = "Arial": .FontSize = 12
        .ForeColor = vbWhite
        Printer.Print "TOTAL: "

        TotalFac = LabelTotalAbonado.Caption
        Hasta = CInt(14 - Len(CStr(TotalFac)))
        For I = 0 To Hasta
            TotalFac = " " & TotalFac
        Next I

        .Font = "Arial": .FontSize = 12
        .CurrentX = 165: .CurrentY = 264
        Printer.Print Format(TotalFac, "Currency")

        Printer.Line (10, 245)-(55, 250), vbBlack, BF
        .CurrentX = 12: .CurrentY = 245
        Printer.Print "RECIBIMOS PESOS:"
        .ForeColor = vbBlack

        'Importe en letras
        TotalFac = Format(TotalFac, "Fixed")
        vImporteEnLetras = EnLetras(CStr(TotalFac))

        .CurrentX = 12: .CurrentY = 253
        Largo = Len(vImporteEnLetras)
        I = 0

        If Len(vImporteEnLetras) <= 50 Then
            Printer.Print StrConv(vImporteEnLetras, vbUpperCase)
        Else
            For I = 50 To 1 Step -1
                If Mid(vImporteEnLetras, I, 1) = " " Then Exit For
            Next I
            Printer.Print StrConv(Mid(vImporteEnLetras, 1, I), vbUpperCase)

            If (Largo - I) <= 50 Then
                .CurrentX = 12: .CurrentY = 258
                Printer.Print StrConv(Mid(vImporteEnLetras, (I + 1), (Largo - I)), vbUpperCase)
            Else
                .CurrentX = 12: .CurrentY = 263
                SegundoTramo = StrConv(Mid(vImporteEnLetras, (I + 1), (Largo - I)), vbUpperCase)
                Largo = Len(SegundoTramo)
                For I = 50 To 1 Step -1
                    If Mid(SegundoTramo, I, 1) = " " Then Exit For
                Next I
                Printer.Print StrConv(Mid(SegundoTramo, 1, I), vbUpperCase)
                 .Font = "Arial": .FontSize = 10
                .CurrentX = 12: .CurrentY = 268
                Printer.Print StrConv(Mid(SegundoTramo, (I + 1), (Largo - I)), vbUpperCase)
            End If
        End If

        .EndDoc
    End With

    tClientes.Close
    tDomiciliosClientes.Close
    BaseSPC.Close
    Exit Sub

CapturaErrores:
    MsgBox "Error imprimiendo recibo: " & Err.Description, vbCritical, "Impresi?n"
End Sub


Private Sub ImprimirReciboX()
    '*** RECIBO X - NO OFICIAL (Linea 2) ***
    On Error GoTo CapturaErroresX

    Dim BaseSPC As DAO.Database
    Dim tClientes As DAO.Recordset
    Dim tDomiciliosClientes As DAO.Recordset

    Dim Nrec As String, IdSuc As String
    Dim Largo As Integer, LargoSuc As Integer
    Dim I As Integer, Hasta As Integer
    Dim TotalFac As Variant
    Dim vEfete As Variant, vCheques As Variant
    Dim vRetenciones As Variant, vTransf As Variant
    Dim vImporteEnLetras As String, SegundoTramo As String

    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
    Set tDomiciliosClientes = BaseSPC.OpenRecordset("DomiciliosClientes", dbOpenTable)
    tClientes.Index = "PrimaryKey"
    tDomiciliosClientes.Index = "PrimaryKey"

    Nrec = CStr(TextNumeroPago.text)

    With Printer
        For I = 0 To Printers.Count - 1
            If Printers(I).DeviceName = "CutePDF Writer" Then Set Printer = Printers(I)
        Next I

        .ScaleHeight = 297
        .ScaleWidth = 210

        'Logo
    '    Printer.PaintPicture LoadPicture(App.Path & "\Quilplac.JPG"), 10, 9, 40, 10

        .FontItalic = False
        .DrawWidth = 10
        Printer.Line (10, 7)-(200, 7)

        '*** CAMBIO: Titulo RECIBO X en lugar de RECIBO OFICIAL ***
        .CurrentX = 90: .CurrentY = 14
        .Font = "Arial": .FontSize = 14: .FontBold = True
        Printer.Print "RECIBO X"

        '*** NUEVO: Leyenda documento no valido ***
        .CurrentX = 80: .CurrentY = 20
        .FontSize = 8: .FontBold = False: .FontItalic = True
        'Printer.Print "Documento no valido como Factura"
        .FontItalic = False

        .CurrentX = 15: .CurrentY = 2
        .FontSize = 12: .FontBold = False
        Printer.Print "ORIGINAL"

        'Numero de recibo
        .FontSize = 12: .CurrentY = 9: .CurrentX = 150

        Largo = 8 - Len(Nrec)
        For I = 1 To Largo
            Nrec = "0" & Nrec
        Next I

        IdSuc = CStr(Left(cmbSucursal.text, 1))
        LargoSuc = 4 - Len(IdSuc)
        For I = 1 To LargoSuc
            IdSuc = "0" & IdSuc
        Next I

        Printer.Print IdSuc & "-" & Nrec

        .CurrentX = 150: .CurrentY = .CurrentY + 2
        .FontSize = 12
        Printer.Print "Fecha: " & Format(TextFechaPago, "DD/MM/YYYY")

        '*** CAMBIO: Sin datos fiscales para recibo no oficial ***
        .DrawWidth = 10
        Printer.Line (10, 42)-(200, 42)

        'Datos empresa (simplificados)
        .CurrentX = 12: .CurrentY = 25
        .Font = "Arial": .FontSize = 10: .FontBold = True: .FontUnderline = False
     '   Printer.Print "QUILPLAC S.A."
     '   .CurrentX = 12: Printer.Print "Andres Baranda 520 - Quilmes"
     '   .CurrentX = 12: Printer.Print "Tel. 4257-5875"

        'Recuadro cliente
        .DrawWidth = 10
        Printer.Line (10, 47)-(200, 47)
        Printer.Line (10, 47)-(10, 75)
        Printer.Line (200, 47)-(200, 75)
        Printer.Line (10, 75)-(200, 75)

        tClientes.MoveFirst
        tClientes.Seek "=", TextCodigoCliente.text
        If Not tClientes.NoMatch Then
            .CurrentX = 15: .CurrentY = 48: .FontSize = 10: .FontBold = True
            Printer.Print "Senor(es): "
            .CurrentX = 35: .CurrentY = 48: .FontBold = False
            Printer.Print tClientes!RazonSocial

            tDomiciliosClientes.Seek "=", tClientes!IdCliente
            If Not tDomiciliosClientes.NoMatch Then
                .CurrentX = 15: .CurrentY = 55: .FontBold = True
                Printer.Print "Domicilio: "
                .CurrentX = 35: .CurrentY = 55: .FontBold = False
                Printer.Print tDomiciliosClientes!Domicilio

                .CurrentX = 15: .CurrentY = 62: .FontBold = True
                Printer.Print "Localidad: "
                .CurrentX = 35: .CurrentY = 62: .FontBold = False
                Printer.Print tDomiciliosClientes!localidad

                .CurrentX = 130: .CurrentY = 62: .FontBold = True
                Printer.Print "Telefono: "
                .CurrentX = 150: .CurrentY = 62: .FontBold = False
                Printer.Print tClientes!Tel
            End If

            .DrawWidth = 10
            Printer.Line (10, 78)-(200, 78)
            Printer.Line (10, 78)-(10, 85)
            Printer.Line (200, 78)-(200, 85)
            Printer.Line (10, 85)-(200, 85)

            .CurrentX = 83: .CurrentY = 80: .FontSize = 10: .FontBold = True
      '      Printer.Print "*** www.quilplac.com ***"
           '  Printer.Print "Energ�a del Futuro... Hoy"
        End If

        'Recuadro detalle
        .DrawWidth = 10
        Printer.Line (10, 90)-(200, 90)
        Printer.Line (10, 240)-(200, 240)
        Printer.Line (10, 90)-(10, 240)
        Printer.Line (200, 90)-(200, 240)
        Printer.Line (10, 97)-(200, 97)

        .FontBold = True
        .CurrentX = 86: .CurrentY = 92: .FontSize = 10
        Printer.Print "DETALLE DEL RECIBO"

        'Formas de pago
        .FontBold = False
        .CurrentX = 32: .CurrentY = 110: .FontSize = 10
        vEfete = Val(TextEfectivo.text)
        Printer.Print "* Efectivo: " & Chr(9) & Chr(9) & Format(vEfete, "Currency")

        .CurrentX = 32: .CurrentY = 120
        vTransf = Val(TextTransferencia.text)
        Printer.Print "* Transferencia: " & Chr(9) & Format(vTransf, "Currency")

        .CurrentX = 32: .CurrentY = 130
        vCheques = Val(TextCheque.text)
        Printer.Print "* Cheques Varios: " & Chr(9) & Format(vCheques, "Currency")

        .CurrentX = 32: .CurrentY = 140
        vRetenciones = Val(TextRetencion.text)
        Printer.Print "* Retenciones: " & Chr(9) & Format(vRetenciones, "Currency")

        'Observaciones
        If TextObservaciones.text <> "" Then
            .CurrentX = 20: .CurrentY = 185
            .Font = "Arial": .FontSize = 10
            Printer.Print "Obs.: " & StrConv(TextObservaciones.text, vbUpperCase)
        End If

        'Recuadro total
        Printer.Line (130, 240)-(130, 262)
        Printer.Line (200, 240)-(200, 262)

        Printer.Line (130, 262)-(200, 270), vbBlack, BF

        .CurrentX = 135: .CurrentY = 264
        .Font = "Arial": .FontSize = 12
        .ForeColor = vbWhite
        Printer.Print "TOTAL: "

        TotalFac = LabelTotalAbonado.Caption
        Hasta = CInt(14 - Len(CStr(TotalFac)))
        For I = 0 To Hasta
            TotalFac = " " & TotalFac
        Next I

        .Font = "Arial": .FontSize = 12
        .CurrentX = 165: .CurrentY = 264
        Printer.Print Format(TotalFac, "Currency")

        Printer.Line (10, 245)-(55, 250), vbBlack, BF
        .CurrentX = 12: .CurrentY = 245
        Printer.Print "RECIBIMOS PESOS:"
        .ForeColor = vbBlack

        'Importe en letras
        TotalFac = Format(TotalFac, "Fixed")
        vImporteEnLetras = EnLetras(CStr(TotalFac))

        .CurrentX = 12: .CurrentY = 253
        Largo = Len(vImporteEnLetras)
        I = 0

        If Len(vImporteEnLetras) <= 50 Then
            Printer.Print StrConv(vImporteEnLetras, vbUpperCase)
        Else
            For I = 50 To 1 Step -1
                If Mid(vImporteEnLetras, I, 1) = " " Then Exit For
            Next I
            Printer.Print StrConv(Mid(vImporteEnLetras, 1, I), vbUpperCase)

            If (Largo - I) <= 50 Then
                .CurrentX = 12: .CurrentY = 258
                Printer.Print StrConv(Mid(vImporteEnLetras, (I + 1), (Largo - I)), vbUpperCase)
            Else
                .CurrentX = 12: .CurrentY = 263
                SegundoTramo = StrConv(Mid(vImporteEnLetras, (I + 1), (Largo - I)), vbUpperCase)
                Largo = Len(SegundoTramo)
                For I = 50 To 1 Step -1
                    If Mid(SegundoTramo, I, 1) = " " Then Exit For
                Next I
                Printer.Print StrConv(Mid(SegundoTramo, 1, I), vbUpperCase)
                .CurrentX = 12: .CurrentY = 268
                Printer.Print StrConv(Mid(SegundoTramo, (I + 1), (Largo - I)), vbUpperCase)
            End If
        End If

        .EndDoc
    End With

    tClientes.Close
    tDomiciliosClientes.Close
    BaseSPC.Close
    Exit Sub

CapturaErroresX:
    MsgBox "Error imprimiendo recibo X: " & Err.Description, vbCritical, "Impresion"
End Sub

Private Function EnLetras(numero As String) As String
    
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
                    expresion = expresion & "mill�n "
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
Private Sub TextEfectivo_Change()

    If TextEfectivo.text <> "" Then
        Call calculoabonado
    End If

End Sub

Private Sub TextMercaderia_Change()

    If TextMercaderia.text <> "" Then
        Call calculoabonado
    End If

End Sub

Private Sub TextRetencion_Change()

    If TextRetencion.text <> "" Then
        Call calculoabonado
    End If

End Sub

Private Sub TextRezago_Change()

    If TextRezago.text <> "" Then
        Call calculoabonado
    End If

End Sub

Private Sub calculoabonado()

    suma = CDec(TextEfectivo.text) + CDec(TextTransferencia.text) + CDec(TextRezago.text) + CDec(TextMercaderia.text) + CDec(TextCheque.text) + CDec(TextRetencion.text) + CDec(TextTarjeta.text)
   
    LabelTotalAbonado.Caption = Format(suma, "#,###,###,#0.00")
    
End Sub

Private Sub TextTarjeta_Change()

    If TextTarjeta.text <> "" Then
        Call calculoabonado
    End If
    
End Sub


