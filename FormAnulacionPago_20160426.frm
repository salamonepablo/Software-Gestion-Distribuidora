VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FormAnulacionPago 
   Caption         =   "Anulacion Pagos"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
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
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   1335
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2055
      Left            =   4560
      TabIndex        =   22
      Top             =   3240
      Visible         =   0   'False
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   3625
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame5 
      Caption         =   "Datos Pago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   7575
      Begin VB.TextBox TextFechaPago 
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
         Left            =   3360
         TabIndex        =   20
         Top             =   480
         Width           =   1335
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
         Left            =   3360
         TabIndex        =   21
         Top             =   240
         Width           =   1035
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
         Left            =   1920
         TabIndex        =   19
         Top             =   240
         Width           =   1155
      End
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
      Height          =   615
      Left            =   120
      TabIndex        =   16
      Top             =   1920
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
         TabIndex        =   3
         Top             =   240
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
         Left            =   1680
         TabIndex        =   2
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   6720
      Width           =   7575
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   3840
         TabIndex        =   9
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonGuardar 
         Caption         =   "&Anular"
         Height          =   750
         Left            =   2760
         TabIndex        =   4
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
      Height          =   3975
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   7575
      Begin VB.TextBox TextEfectivo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   29
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TextCheque 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   28
         Top             =   2040
         Width           =   1335
      End
      Begin VB.TextBox TextMercaderia 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   27
         Top             =   1560
         Width           =   1335
      End
      Begin VB.TextBox TextRezago 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TextRetencion 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   25
         Top             =   2520
         Width           =   1335
      End
      Begin VB.TextBox TextTarjeta 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   24
         Top             =   3000
         Width           =   1335
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
         TabIndex        =   35
         Top             =   720
         Width           =   780
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
         TabIndex        =   34
         Top             =   1200
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
         Left            =   600
         TabIndex        =   33
         Top             =   1680
         Width           =   1020
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
         TabIndex        =   32
         Top             =   2160
         Width           =   720
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
         Left            =   600
         TabIndex        =   31
         Top             =   2640
         Width           =   945
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
         TabIndex        =   30
         Top             =   3120
         Width           =   675
      End
      Begin VB.Label LabelTotalAbonado 
         AutoSize        =   -1  'True
         Caption         =   "dfdfd"
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
         Left            =   5880
         TabIndex        =   23
         Top             =   3480
         Width           =   540
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
         Left            =   4080
         TabIndex        =   17
         Top             =   3480
         Width           =   1620
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
      Height          =   855
      Left            =   120
      TabIndex        =   5
      Top             =   960
      Width           =   7575
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
         Left            =   3240
         TabIndex        =   12
         Top             =   480
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
         TabIndex        =   11
         Top             =   480
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
         TabIndex        =   10
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TextCodigoCliente 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   1800
         TabIndex        =   1
         Top             =   480
         Width           =   1335
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
         TabIndex        =   15
         Top             =   240
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
         Left            =   3240
         TabIndex        =   14
         Top             =   240
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
         TabIndex        =   13
         Top             =   240
         Width           =   1185
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Nº Cliente"
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
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
   End
End
Attribute VB_Name = "FormAnulacionPago"
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
Dim rezago As Double
Dim mercaderia As Double
Dim cheque As Double
Dim retencion As Double
Dim tarjeta As Double

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
    
    '*** Grabo Cuenta Corriente
     
    
    CodigoClie = Val(TextCodigoCliente.Text)
    rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
        
    rstCtaCte.Edit
    rstCtaCte.Fields!IdCliente = TextCodigoCliente.Text
    saldo1 = rstCtaCte.Fields!SaldoL1
    saldo2 = rstCtaCte.Fields!SaldoL2
    
    If saldoLinea1 = 1 Then
        saldoLi1 = LabelTotalAbonado.Caption
        saldoLi1 = saldo1 + saldoLi1
        rstCtaCte.Fields!SaldoL1 = Format(saldoLi1, "#0.00")
        If saldoLi1 <> 0 Then
            saldoTotalForm = saldoLi1 - saldo2
        Else
            saldoTotalForm = saldo2
        End If
        rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#0.00")
    End If
    
    If saldoLinea2 = 2 Then
        saldoLi2 = LabelTotalAbonado.Caption
        saldoLi2 = saldo2 + saldoLi2
        rstCtaCte.Fields!SaldoL2 = Format(saldoLi2, "#0.00")
        If saldoLi2 <> 0 Then
            saldoTotalForm = saldoLi2 - saldo1
        Else
             saldoTotalForm = saldo1
        End If
        rstCtaCte.Fields!SaldoTotal = Format(saldoTotalForm, "#0.00")
    End If
    
    rstCtaCte.Update

    
    '*** Grabo Movimientos Cuente corriente
   
    rstMovimientosCtaCte.AddNew
    'rstMovimientosCtaCte.Fields!Fecha = Format(Date, "dd/mm/yyyy")
    rstMovimientosCtaCte.Fields!Fecha = Format(TextFechaPago.Text, "dd/mm/yyyy")
    rstMovimientosCtaCte.Fields!IdCliente = TextCodigoCliente.Text
    If OptionSaldoLinea1.Value = True Then
        rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago Nº " & TextNumeroPago.Text & " Linea 1"
        rstMovimientosCtaCte.Fields!ImporteLinea1 = Format(LabelTotalAbonado.Caption, "#0.00")
        rstMovimientosCtaCte.Fields!ImporteLinea2 = 0
    End If
    If OptionSaldoLinea2.Value = True Then
        rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago Nº " & TextNumeroPago.Text & " Linea 2"
        rstMovimientosCtaCte.Fields!ImporteLinea2 = Format(LabelTotalAbonado.Caption, "#0.00")
        rstMovimientosCtaCte.Fields!ImporteLinea1 = 0
    End If
    
    rstMovimientosCtaCte.Fields!NroDoc = 99 & TextNumeroPago.Text
    
    rstMovimientosCtaCte.Update
            
      
    '*** Grabo Pagos
    
    rstPagoC.AddNew
    rstPagoC.Fields!NroPago = 99 & TextNumeroPago.Text
    rstPagoC.Fields!FechaPago = Format(Date, "dd/mm/yyyy")
    rstPagoC.Fields!IdCliente = TextCodigoCliente.Text
    rstPagoC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "-#0.00")
    rstPagoC.Update

    'rstPagoD.AddNew
    NroLinea = 0
    
    If TextEfectivo.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Efectivo"
        rstPagoD.Fields!ImportePago = Format(TextEfectivo.Text, "-#0.00")
        rstPagoD.Update
    End If
        
    If TextRezago.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Rezago"
        rstPagoD.Fields!ImportePago = Format(TextRezago.Text, "-#0.00")
        rstPagoD.Update
    End If
             
    If TextMercaderia.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Mercaderia"
        rstPagoD.Fields!ImportePago = Format(TextMercaderia.Text, "-#0.00")
        rstPagoD.Update
    End If
    
    If TextCheque.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Cheque"
        rstPagoD.Fields!ImportePago = Format(TextCheque.Text, "-#0.00")
        rstPagoD.Update
    End If
    
    If TextRetencion.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Retencion"
        rstPagoD.Fields!ImportePago = Format(TextRetencion.Text, "-#0.00")
        rstPagoD.Update
    End If
    
    If TextTarjeta.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = 99 & TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Tarjeta"
        rstPagoD.Fields!ImportePago = Format(TextTarjeta.Text, "-#0.00")
        rstPagoD.Update
    End If
    
   
    
    saldoLinea1 = 0
    saldoLinea2 = 0
    
   
    Call blanco

End Sub

Private Sub BotonSalir_Click()

    Unload FormAnulacionPago
    
End Sub

Private Sub Form_Load()

    FormAnulacionPago.Height = 8625
    FormAnulacionPago.Width = 8055
    FormAnulacionPago.Top = 1000
    FormAnulacionPago.Left = 1000
    
    
    
End Sub


Private Sub blanco()

    TextCodigoCliente.Text = ""
    TextSaldoTotal.Text = 0
    TextSaldoLinea1.Text = 0
    TextSaldoLinea2.Text = 0
    'TextNumeroPago.Text = 0
    TextEfectivo.Text = ""
    TextRezago.Text = ""
    TextMercaderia.Text = ""
    TextCheque.Text = ""
    TextRetencion.Text = ""
    TextTarjeta.Text = ""
    TextNumeroPago.Text = ""
    TextFechaPago.Text = ""
    LabelTotalAbonado.Caption = ""
    
    BotonGuardar.Enabled = False
    Frame1.Enabled = False
    OptionSaldoLinea1.Value = False
    OptionSaldoLinea2.Value = False
    
'    TextCodigoCliente.SetFocus
    
End Sub








Private Sub TextFechaPago_GotFocus()
    TextFechaPago.SelLength = Len(TextFechaPago.Text)
End Sub

Private Sub TextNumeroPago_GotFocus()
     TextNumeroPago.SelLength = Len(TextNumeroPago.Text)
End Sub

Private Sub TextNumeroPago_KeyPress(KeyAscii As Integer)

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoC = db.OpenRecordset("PagoC", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstPagoD = db.OpenRecordset("PagoD", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstMovimientosCtaCte = db.OpenRecordset("MovimientosCtaCte", dbOpenDynaset)
    
    If KeyAscii = 13 Then
        
        '************Busco Pago
        
        NumPago = Val(TextNumeroPago.Text)
        
        rstPagoC.FindFirst "NroPago= " + Str(NumPago)
        
        TextCodigoCliente.Text = rstPagoC.Fields!IdCliente
        TextFechaPago.Text = rstPagoC.Fields!FechaPago
        
        '********Busco Saldo en Cuenta Corriente
        
        CodigoClie = Val(TextCodigoCliente.Text)
            
        rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
        If rstCtaCte.Fields!IdCliente <> Val(TextCodigoCliente.Text) Then
            mensaje = MsgBox("No Posee Cuentas Corientes", vbCritical, "Final de la busqueda")
            TextCodigoCliente.Text = ""
            Call blanco
        Else
            TextSaldoLinea1.Text = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
            TextSaldoLinea2.Text = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
            TextSaldoTotal.Text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
            
        
        
            rstMovimientosCtaCte.FindFirst "NroDoc= " + Str(NumPago)
            pagolinea = rstMovimientosCtaCte.Fields!tipoDoc
            If pagolinea = "Pago Linea 1" Then
               OptionSaldoLinea1.Value = True
               saldoLinea1 = 1
            End If
            If pagolinea = "Pago Linea 2" Then
               OptionSaldoLinea2.Value = True
               saldoLinea2 = 2
            End If
            KeyAscii = 0
            SendKeys "{TAB}"
        End If
        
        '******Busco Detalle de Pago
        
        MSHFlexGrid1.Rows = 1
        MSHFlexGrid1.Clear
        
        
        rstPagoD.FindFirst "NroPago= " + Str(NumPago)
        
        linea2 = 1
        Do While Not rstPagoD.NoMatch
            MSHFlexGrid1.AddItem " "
            MSHFlexGrid1.Row = linea2
        
            MSHFlexGrid1.Col = 0
            MSHFlexGrid1.Text = rstPagoD.Fields!ImportePago
            If rstPagoD.Fields!FormaPago = "Efectivo" Then
               TextEfectivo.Text = MSHFlexGrid1.Text
            End If
            If rstPagoD.Fields!FormaPago = "Rezago" Then
               TextRezago.Text = MSHFlexGrid1.Text
            End If
            If rstPagoD.Fields!FormaPago = "Mercaderia" Then
               TextMercaderia.Text = MSHFlexGrid1.Text
            End If
            If rstPagoD.Fields!FormaPago = "Cheque" Then
               TextCheque.Text = MSHFlexGrid1.Text
            End If
            If rstPagoD.Fields!FormaPago = "Retencion" Then
               TextRetencion.Text = MSHFlexGrid1.Text
            End If
            If rstPagoD.Fields!FormaPago = "Tarjeta" Then
               TextTarjeta.Text = MSHFlexGrid1.Text
            End If
            LabelTotalAbonado.Caption = Val(LabelTotalAbonado.Caption) + rstPagoD.Fields!ImportePago
            linea2 = linea2 + 1
           
            rstPagoD.FindNext "NroPago= " + Str(NumPago)
            
         Loop
       
         
    End If
    

End Sub


Private Sub calculo()
        
        
        efectivo = Val(TextEfectivo.Text)
        If efectivo < 0 Then
            efectivo = 0
        End If
        
        rezago = Val(TextRezago.Text)
        If rezago < 0 Then
            rezago = 0
        End If
        
        mercaderia = Val(TextMercaderia.Text)
        If mercaderia < 0 Then
            mercaderia = 0
        End If
         
        cheque = Val(TextCheque.Text)
        If cheque < 0 Then
            cheque = 0
        End If
        
        retencion = Val(TextRetencion.Text)
        If retencion < 0 Then
            retencion = 0
        End If
        
        tarjeta = Val(TextTarjeta.Text)
        If tarjeta < 0 Then
            tarjeta = 0
        End If
       
       
    'resta = CDec(TextSaldoLinea1.Text) - CDec(TextEfectivo.Text) - CDec(TextRezago.Text) - CDec(TextMercaderia.Text) - CDec(TextCheque.Text) - CDec(TextRetencion.Text) - CDec(TextTarjeta.Text)

    If saldoLinea1 = 1 Then
        resta = CDec(TextSaldoLinea1.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
    End If
    
    If saldoLinea2 = 2 Then
        resta = CDec(TextSaldoLinea2.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
    End If
    
    TextResta.Text = Format(resta, "#00.00")
    
End Sub

