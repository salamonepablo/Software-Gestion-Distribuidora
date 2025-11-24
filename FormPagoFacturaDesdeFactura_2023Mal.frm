VERSION 5.00
Begin VB.Form FormPagoFacturasDesdeFactura 
   Caption         =   "Pago Facturas"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7935
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
      TabIndex        =   34
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
         Left            =   4200
         TabIndex        =   11
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
         TabIndex        =   10
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   16
      Top             =   6720
      Width           =   7575
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   3840
         TabIndex        =   8
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonGuardar 
         Caption         =   "&Guardar"
         Height          =   750
         Left            =   2760
         TabIndex        =   7
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
      Height          =   3975
      Left            =   120
      TabIndex        =   14
      Top             =   2640
      Width           =   7575
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
         TabIndex        =   0
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
         TabIndex        =   12
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox TextObservaciones 
         Height          =   525
         Left            =   3120
         MultiLine       =   -1  'True
         TabIndex        =   25
         Top             =   1680
         Width           =   4335
      End
      Begin VB.TextBox TextTarjeta 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox TextRetencion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TextRezago 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   3
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TextMercaderia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   4
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TextCheque 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   2
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TextResta 
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
         Left            =   2760
         TabIndex        =   21
         Top             =   3480
         Width           =   1335
      End
      Begin VB.TextBox TextEfectivo 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   1
         Top             =   480
         Width           =   1335
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
         Left            =   5280
         TabIndex        =   40
         Top             =   3480
         Width           =   1695
      End
      Begin VB.Label Label16 
         Caption         =   "Saldo Actualizado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5160
         TabIndex        =   39
         Top             =   3120
         Width           =   1935
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
         TabIndex        =   38
         Top             =   960
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
         Left            =   5040
         TabIndex        =   37
         Top             =   2520
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
         Left            =   3120
         TabIndex        =   36
         Top             =   2520
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
         TabIndex        =   35
         Top             =   3600
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
         TabIndex        =   27
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
         TabIndex        =   26
         Top             =   1440
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
         Left            =   360
         TabIndex        =   24
         Top             =   3000
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
         TabIndex        =   23
         Top             =   2520
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
         Left            =   360
         TabIndex        =   22
         Top             =   3600
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
         Left            =   360
         TabIndex        =   20
         Top             =   1080
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
         TabIndex        =   19
         Top             =   2040
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
         Left            =   360
         TabIndex        =   18
         Top             =   1560
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
         Left            =   360
         TabIndex        =   17
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
      TabIndex        =   13
      Top             =   120
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
         Left            =   2400
         TabIndex        =   30
         Top             =   720
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
         Left            =   3960
         TabIndex        =   29
         Top             =   720
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
         Left            =   5520
         TabIndex        =   28
         Top             =   720
         Width           =   1335
      End
      Begin VB.TextBox TextCodigoCliente 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         TabIndex        =   9
         Top             =   720
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
         Left            =   5520
         TabIndex        =   33
         Top             =   480
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
         Left            =   2400
         TabIndex        =   32
         Top             =   480
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
         Left            =   3960
         TabIndex        =   31
         Top             =   480
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
         Left            =   480
         TabIndex        =   15
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "FormPagoFacturasDesdeFactura"
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
Dim ImprimirPresupuesto As Integer
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
    
    '*** Grabo Pagos
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db1 = DBEngine.OpenDatabase(ruta)
    
    Set rstPagc = db1.OpenRecordset("PagoC", dbOpenTable)
    
    rstPagc.Index = "PrimaryKey"
    
    rstPagc.Seek "=", Str(TextNumeroPago.text)

    If Not rstPagc.NoMatch Then
        A = MsgBox("Pago Existente", vbCritical, "INFO DEL SISTEMA")
       
        TextNumeroPago.text = num
        TextNumeroPago.SetFocus
    Else
    
    rstPagc.Close
    db1.Close
    
    
    
    rstPagoC.AddNew
    rstPagoC.Fields!NroPago = TextNumeroPago.text
    rstPagoC.Fields!FechaPago = Format(Date, "dd/mm/yyyy")
    rstPagoC.Fields!IdCliente = TextCodigoCliente.text
    If OptionSaldoLinea1.Value = True Then rstPagoC.Fields!Corresponde = "L1"
    If OptionSaldoLinea2.Value = True Then rstPagoC.Fields!Corresponde = "L2"
    rstPagoC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "#,###,###,#0.00")
    rstPagoC.Update

    
    
    
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
            
      
    
    'rstPagoD.AddNew
    NroLinea = 0
    
    If TextEfectivo.text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Efectivo"
        rstPagoD.Fields!ImportePago = Format(Val(TextEfectivo.text), "#,###,###,#0.00")
        rstPagoD.Update
    End If
        
    If TextRezago.text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Rezago"
        rstPagoD.Fields!ImportePago = Format(Val(TextRezago.text), "#,###,###,#0.00")
        rstPagoD.Update
    End If
             
    If TextMercaderia.text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Mercaderia"
        rstPagoD.Fields!ImportePago = Format(Val(TextMercaderia.text), "#,###,###,#0.00")
        rstPagoD.Update
    End If
    
    If TextCheque.text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Cheque"
        rstPagoD.Fields!ImportePago = Format(Val(TextCheque.text), "#,###,###,#0.00")
        rstPagoD.Update
    End If
    
    If TextRetencion.text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Retencion"
        rstPagoD.Fields!ImportePago = Format(Val(TextRetencion.text), "#,###,###,#0.00")
        rstPagoD.Update
    End If
    
    If TextTarjeta.text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Tarjeta"
        rstPagoD.Fields!ImportePago = Format(Val(TextTarjeta.text), "#,###,###,#0.00")
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
    
    '****** Actualizo ultimo numero pago
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)

    Dim busco As String

    busco = "tPagosC"
    
    rstUltimosNumeros.FindFirst "IDTabla >= '" & busco & "' "
    ultimo = rstUltimosNumeros.Fields!UltimoNumero
    
    If ultimo < Val(TextNumeroPago.text) Then
        rstUltimosNumeros.Edit
        'If ultimo < rstUltimosNumeros.Fields!UltimoNumero Then
             rstUltimosNumeros.Fields!UltimoNumero = TextNumeroPago.text
        'End If
        rstUltimosNumeros.Update
    End If
    
    
    
    saldoLinea1 = 0
    saldoLinea2 = 0
    
    'Unload FormPagoFacturasDesdeFactura
    
    Call nuevonumeroPago
    
    respuesta = MsgBox("Desea Imprimir?", vbYesNo, "Facturas")
    
    If respuesta = vbYes Then
        If ImprimirPresupuesto = 1 Then
            Call blanco
            Call FormPresupuesto.Imprimir
            FormPresupuesto.MSFlexGrid1.Visible = False
            Unload Me
         Else
            Call blanco
            FormImprimir.Show
            FormPresupuesto.MSFlexGrid1.Visible = False
            Unload Me
        End If
      Else
       
       If ImprimirPresupuesto = 1 Then
            Call blanco
            Call FormPresupuesto.blanqueototal
            FormPresupuesto.TextCodigoCliente.SetFocus
            FormPresupuesto.MSFlexGrid1.Visible = False
            'FormPresupuesto.BotonNueva.Enabled = True
            'FormPresupuesto.BotonNueva.SetFocus
         Else
            Call blanco
            Call FormFactura.blanqueototal
            FormFactura.TextCodigoCliente.SetFocus
            FormFactura.MSFlexGrid1.Visible = False
            'FormFactura.BotonNueva.Enabled = True
            'FormFactura.BotonNueva.SetFocus
    '       FormFactura.TextCodigoCliente.SetFocus
       End If
       Unload FormPagoFacturasDesdeFactura
    End If
'    Call blanco
End If
End Sub

Private Sub BotonGuardar_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub BotonSalir_Click()

'Prueba de error en local 04/04/2016
    If FormFactura.TextCodigoCliente <> "" Then
 '       FormFactura.blanqueototal
 '       FormFactura.TextCodigoCliente.SetFocus
        FormImprimir.Show
    End If
                
    If FormPresupuesto.TextCodigoCliente <> "" Then
        FormPresupuesto.blanqueototal
        FormPresupuesto.BotonNueva.Enabled = True
        FormPresupuesto.TextCodigoCliente.SetFocus
        FormPresupuesto.MSFlexGrid1.Visible = False
    End If
    
    
    Unload FormPagoFacturasDesdeFactura
    
End Sub

Private Sub BotonSalir_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    FormPagoFacturasDesdeFactura.Height = 8625
    FormPagoFacturasDesdeFactura.Width = 8055
    FormPagoFacturasDesdeFactura.Top = 1000
    FormPagoFacturasDesdeFactura.Left = 1000
    
    ImprimirPresupuesto = 0
    
    'TextCodigoCliente.Enabled = True
    'TextCodigoCliente.SetFocus
       
    
    TextFechaPago.text = Format(Date, "dd/mm/yyyy")
    
    If FormFactura.Textfac.text <> "" Then
       FormPagoFacturasDesdeFactura.Top = 1000
       FormPagoFacturasDesdeFactura.Left = 12300
       TextCodigoCliente.text = FormFactura.TextCodigoCliente
    End If
    
    If FormPresupuesto.Textpre <> "" Then
        FormPagoFacturasDesdeFactura.Top = 1000
        FormPagoFacturasDesdeFactura.Left = 12300
        TextCodigoCliente.text = FormPresupuesto.TextCodigoCliente
        ImprimirPresupuesto = 1
    End If
    
    'Call nuevonumeroPago
    
    
    
    If FormPresupuesto.TextNumeroPresupuesto.text <> "" Then
        OptionSaldoLinea2.Value = True
        TextNumeroPago.text = FormPresupuesto.TextNumeroPresupuesto.text
        TextFechaPago.text = FormPresupuesto.TextFechaPresupuesto
'        FormPagoFacturas.TextEfectivo.SetFocus
        LlamaPagoPresup = False
    End If
    
   
    If FormFactura.TextNumeroFactura.text <> "" Then
        OptionSaldoLinea1.Value = True
        TextNumeroPago.text = FormFactura.TextNumeroFactura.text
        TextFechaPago.text = FormFactura.TextFechaFactura.text
'        FormPagoFacturas.TextEfectivo.SetFocus
        LlamaPagoFactura = False
    End If
    
    If TextCodigoCliente.text <> "" Then
        Call buscosaldo
    End If
    
    
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

       'TextCodigoCliente.SetFocus
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
       'TextSaldoLinea1.Text = Format(rstCtaCte.Fields!SaldoL1, "#,###,###,#0.00")
       'TextSaldoLinea2.Text = Format(rstCtaCte.Fields!SaldoL2, "#,###,###,#0.00")
       'TextSaldoTotal.Text = Format(rstCtaCte.Fields!SaldoTotal, "#,###,###,#0.00")
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
    rstUltNum.Seek "=", "tPagosC"
        
    If Not rstUltNum.NoMatch Then
        numeroPago = rstUltNum!UltimoNumero
        TextNumeroPago.text = numeroPago + 1
    End If


       
End Sub

Private Sub OptionSaldoLinea1_Click()

    If OptionSaldoLinea1.Value = True Then
        saldoLinea2 = 0
        LabelSaldo.Caption = "Saldo Linea 1"
        BotonGuardar.Enabled = True
        Frame1.Enabled = True
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

Private Sub TextCheque_Change()

     If TextCheque.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If

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
    
    If KeyAscii = 46 Then KeyAscii = 44

End Sub

Private Sub TextCodigoCliente_Change()


'    If TextCodigoCliente.Text <> "" Then
'        Call buscosaldo
'    End If
End Sub

Private Sub TextCodigoCliente_GotFocus()
    TextCodigoCliente.SelLength = Len(TextCodigoCliente.text)
End Sub

Private Sub TextCodigoCliente_KeyPress(KeyAscii As Integer)
  
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
   
    CodigoClie = Val(TextCodigoCliente.text)
 
    If KeyAscii = 13 Then
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
                KeyAscii = 0
                Sendkeys "{TAB}"
            End If
        End If
    End If
    Call buscosaldo
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextCodigoCliente_LostFocus()

    If TextCodigoCliente.text <> "" Then
        Call buscosaldo
    End If

End Sub

Private Sub TextEfectivo_Change()

    If TextEfectivo.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If

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

    If KeyAscii = 46 Then KeyAscii = 44

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

End Sub
Private Sub blancoCambioCheck()

  
    TextEfectivo.text = ""
    TextRezago.text = ""
    TextMercaderia.text = ""
    TextObservaciones.text = ""
    TextCheque.text = ""
    TextRetencion.text = ""
    TextTarjeta.text = ""
    
    
    
    

End Sub



Private Sub blanco()

    TextCodigoCliente.text = ""
    TextSaldoTotal.text = 0
    TextSaldoLinea1.text = 0
    TextSaldoLinea2.text = 0
    'TextNumeroPago.Text = 0
    TextEfectivo.text = ""
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
    
    If KeyAscii = 46 Then KeyAscii = 44

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

Private Sub TextRetencion_Change()

    If TextRetencion.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If

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
    
    If KeyAscii = 46 Then KeyAscii = 44

End Sub

Private Sub TextRezago_Change()

    If TextRezago.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If

End Sub

Private Sub calculo()
        
        
        'efectivo = Val(TextEfectivo.text)
        If TextEfectivo.text <> "" Then efectivo = CDbl(TextEfectivo.text)
        
        If efectivo < 0 Then
            efectivo = 0
        End If
        
        'rezago = Val(TextRezago.text)
        If TextRezago.text <> "" Then rezago = CDbl(TextRezago.text)
        
        If rezago < 0 Then
            rezago = 0
        End If
        
        'mercaderia = Val(TextMercaderia.text)
        If TextMercaderia.text <> "" Then mercaderia = CDbl(TextMercaderia.text)
        
        If mercaderia < 0 Then
            mercaderia = 0
        End If
         
        'cheque = Val(TextCheque.text)
        If TextCheque.text <> "" Then cheque = CDbl(TextCheque.text)
        
        If cheque < 0 Then
            cheque = 0
        End If
        
        'retencion = Val(TextRetencion.text)
        If TextRetencion.text <> "" Then retencion = CDbl(TextRetencion.text)
        
        If retencion < 0 Then
            retencion = 0
        End If
        
        'tarjeta = Val(TextTarjeta.text)
        If TextTarjeta.text <> "" Then tarjeta = CDbl(TextTarjeta.text)
        
        If tarjeta < 0 Then
            tarjeta = 0
        End If
       
       
    'resta = CDec(TextSaldoLinea1.Text) - CDec(TextEfectivo.Text) - CDec(TextRezago.Text) - CDec(TextMercaderia.Text) - CDec(TextCheque.Text) - CDec(TextRetencion.Text) - CDec(TextTarjeta.Text)

    If saldoLinea1 = 1 Then
            resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
    End If
    
    If saldoLinea2 = 2 Then
            resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
    End If
    
    TextResta.text = Format(resta, "#,###,###,#0.00")
    LabelSaldoTotal.Caption = Format(resta2, "#,###,###,#0.00")
    
End Sub

Private Sub calculoabonado()

    suma = CDec(Format(efectivo, "#,###,###,#0.00")) + CDec(Format(rezago, "#,###,###,#0.00")) + CDec(Format(mercaderia, "#,###,###,#0.00")) + CDec(Format(cheque, "#,###,###,#0.00")) + CDec(Format(retencion, "#,###,###,#0.00")) + CDec(Format(tarjeta, "#,###,###,#0.00"))
   
    LabelTotalAbonado.Caption = Format(suma, "#,###,###,#0.00")
    
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
    
    If KeyAscii = 46 Then KeyAscii = 44

End Sub

Private Sub TextSaldoLinea1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextSaldoLinea2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextSaldoTotal_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextTarjeta_Change()

    If TextTarjeta.text <> "" Then
        Call calculo
        Call calculoabonado
    Else
        Call calculoresta
        Call calculoabonadoresta
    End If
    
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
    
    If KeyAscii = 46 Then KeyAscii = 44
    
End Sub
Private Sub calculoresta()

        'efectivo = Val(TextEfectivo.text)
        If TextEfectivo.text <> "" Then efectivo = CDbl(TextEfectivo.text)
        
        If efectivo < 0 Then
            efectivo = 0
        End If
        
        'rezago = Val(TextRezago.text)
        If TextRezago.text <> "" Then rezago = CDbl(TextRezago.text)
        
        If rezago < 0 Then
            rezago = 0
        End If
        
        'mercaderia = Val(TextMercaderia.text)
        If TextMercaderia.text <> "" Then mercaderia = CDbl(TextMercaderia.text)
        
        If mercaderia < 0 Then
            mercaderia = 0
        End If
         
        'cheque = Val(TextCheque.text)
        If TextCheque.text <> "" Then cheque = CDbl(TextCheque.text)
        
        If cheque < 0 Then
            cheque = 0
        End If
        
        'retencion = Val(TextRetencion.text)
        If TextRetencion.text <> "" Then retencion = CDbl(TextRetencion.text)
        
        If retencion < 0 Then
            retencion = 0
        End If
        
        'tarjeta = Val(TextTarjeta.text)
        If TextTarjeta.text <> "" Then tarjeta = CDbl(TextTarjeta.text)
        
        If tarjeta < 0 Then
            tarjeta = 0
        End If
       
    'resta = CDec(TextSaldoLinea1.Text) - CDec(TextEfectivo.Text) - CDec(TextRezago.Text) - CDec(TextMercaderia.Text) - CDec(TextCheque.Text) - CDec(TextRetencion.Text) - CDec(TextTarjeta.Text)

    If saldoLinea1 = 1 Then
        If efectivo = 0 Then
            resta = CDec(TextSaldoLinea1.text) + CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) + CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If rezago = 0 Then
             resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If mercaderia = 0 Then
             resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If cheque = 0 Then
            resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If retencion = 0 Then
            resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
        End If
         If tarjeta = 0 Then
           resta = CDec(TextSaldoLinea1.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
        End If
    End If
    
    If saldoLinea2 = 2 Then
       If efectivo = 0 Then
            resta = CDec(TextSaldoLinea2.text) + CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) + CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If rezago = 0 Then
             resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If mercaderia = 0 Then
             resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If cheque = 0 Then
            resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If retencion = 0 Then
            resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
        End If
         If tarjeta = 0 Then
           resta = CDec(TextSaldoLinea2.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
        End If
     
    End If
    
    TextResta.text = Format(resta, "#,###,###,#0.00")
    LabelSaldoTotal.Caption = Format(resta2, "#,###,###,#0.00")


End Sub



Private Sub calculoabonadoresta()

    suma = CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
   
    LabelTotalAbonado.Caption = Format(suma, "#,###,###,#0.00")

End Sub
