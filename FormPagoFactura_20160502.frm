VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form FormPagoFacturas 
   Caption         =   "Pago Facturas"
   ClientHeight    =   8055
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8055
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1815
      Left            =   240
      TabIndex        =   43
      Top             =   1200
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
         TabIndex        =   2
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
         TabIndex        =   1
         Top             =   480
         Width           =   1935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   15
      Top             =   6720
      Width           =   7575
      Begin VB.CommandButton BotonSalir 
         Caption         =   "&Salir"
         Height          =   750
         Left            =   3840
         TabIndex        =   16
         Top             =   240
         Width           =   750
      End
      Begin VB.CommandButton BotonGuardar 
         Caption         =   "&Guardar"
         Height          =   750
         Left            =   2760
         TabIndex        =   11
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
      TabIndex        =   13
      Top             =   2640
      Width           =   7575
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3600
         TabIndex        =   41
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
         TabIndex        =   3
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
         TabIndex        =   10
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
         TabIndex        =   9
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox TextRetencion 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   2400
         Width           =   1335
      End
      Begin VB.TextBox TextRezago 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   6
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TextMercaderia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   1920
         Width           =   1335
      End
      Begin VB.TextBox TextCheque 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   1560
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   480
         Width           =   1335
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
         TabIndex        =   40
         Top             =   3120
         Width           =   1935
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
         TabIndex        =   39
         Top             =   3480
         Width           =   1695
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
      TabIndex        =   12
      Top             =   120
      Width           =   7575
      Begin VB.TextBox Textcod 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   840
         TabIndex        =   44
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox TextNombre 
         Height          =   285
         Left            =   2400
         TabIndex        =   42
         Top             =   240
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
         Left            =   2400
         TabIndex        =   30
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
         Left            =   3960
         TabIndex        =   29
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
         Left            =   5520
         TabIndex        =   28
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TextCodigoCliente 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   480
         TabIndex        =   0
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
         Left            =   2400
         TabIndex        =   32
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
         Left            =   3960
         TabIndex        =   31
         Top             =   720
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
         TabIndex        =   14
         Top             =   480
         Width           =   855
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
    
    
    '*** Grabo Pagos
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db1 = DBEngine.OpenDatabase(ruta)
    
    Set rstPagc = db1.OpenRecordset("PagoC", dbOpenTable)
    
    rstPagc.Index = "PrimaryKey"
    
    rstPagc.Seek "=", Str(TextNumeroPago.Text)

    If Not rstPagc.NoMatch Then
        A = MsgBox("Pago Existente", vbCritical, "INFO DEL SISTEMA")
       
        TextNumeroPago.Text = num
        TextNumeroPago.SetFocus
    Else
    
    rstPagc.Close
    db1.Close
    
    
    '*** Grabo Cuenta Corriente
    
    CodigoClie = Val(TextCodigoCliente.Text)
    rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
        
    rstCtaCte.Edit
    rstCtaCte.Fields!IdCliente = TextCodigoCliente.Text
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
    rstMovimientosCtaCte.Fields!Fecha = TextFechaPago.Text
    rstMovimientosCtaCte.Fields!IdCliente = TextCodigoCliente.Text
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
    rstMovimientosCtaCte.Fields!NroDoc = TextNumeroPago.Text
    
    rstMovimientosCtaCte.Update
            
      
    '*** Grabo Pagos
    
    rstPagoC.AddNew
    rstPagoC.Fields!NroPago = TextNumeroPago.Text
    rstPagoC.Fields!FechaPago = Format(Date, "dd/mm/yyyy")
    rstPagoC.Fields!IdCliente = TextCodigoCliente.Text
    rstPagoC.Fields!TotalAbonado = Format(LabelTotalAbonado.Caption, "#0.00")
    rstPagoC.Update

    'rstPagoD.AddNew
    NroLinea = 0
    
    If TextEfectivo.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Efectivo"
        rstPagoD.Fields!ImportePago = TextEfectivo.Text
        rstPagoD.Update
    End If
        
    If TextRezago.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Rezago"
        rstPagoD.Fields!ImportePago = TextRezago.Text
        rstPagoD.Update
    End If
             
    If TextMercaderia.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Mercaderia"
        rstPagoD.Fields!ImportePago = TextMercaderia.Text
        rstPagoD.Update
    End If
    
    If TextCheque.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Cheque"
        rstPagoD.Fields!ImportePago = TextCheque.Text
        rstPagoD.Update
    End If
    
    If TextRetencion.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Retencion"
        rstPagoD.Fields!ImportePago = TextRetencion.Text
        rstPagoD.Update
    End If
    
    If TextTarjeta.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!FormaPago = "Tarjeta"
        rstPagoD.Fields!ImportePago = TextTarjeta.Text
        rstPagoD.Update
    End If
    
    
    If TextObservaciones.Text <> "" Then
        rstPagoD.AddNew
        rstPagoD.Fields!NroPago = TextNumeroPago.Text
        If NroLinea >= 0 Then NroLinea = NroLinea + 1
        rstPagoD.Fields!LineaPago = CInt(NroLinea)
        rstPagoD.Fields!Observaciones = TextObservaciones.Text
        rstPagoD.Update
    End If
    
    '****** Actualizo ultimo numero pago
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstUltimosNumeros = db.OpenRecordset("UltimosNumeros", dbOpenDynaset)

    Dim Busco As String

    Busco = "tPagosC"
    
    rstUltimosNumeros.FindFirst "IDTabla >= '" & Busco & "' "
    ultimo = rstUltimosNumeros.Fields!UltimoNumero
    
    If ultimo < Val(TextNumeroPago.Text) Then
        rstUltimosNumeros.Edit
        'If ultimo < rstUltimosNumeros.Fields!UltimoNumero Then
             rstUltimosNumeros.Fields!UltimoNumero = TextNumeroPago.Text
        'End If
        rstUltimosNumeros.Update
    End If
    
    
    
    saldoLinea1 = 0
    saldoLinea2 = 0
    
    Call nuevonumeroPago
    Call blanco

    TextCodigoCliente.SetFocus
    MSFlexGrid1.Visible = False
End If
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


Private Sub Form_Load()

    FormPagoFacturas.Height = 8625
    FormPagoFacturas.Width = 8055
    FormPagoFacturas.Top = 1000
    FormPagoFacturas.Left = 1000
    
    TextFechaPago.Text = Format(Date, "dd/mm/yyyy")
    
   
    
    Call nuevonumeroPago
    
End Sub
Private Sub buscosaldo()
 
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
   
    CodigoClie = Val(TextCodigoCliente.Text)
 
         
    rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
    If rstCtaCte.Fields!IdCliente <> Val(TextCodigoCliente.Text) Then
       mensaje = MsgBox("No Posee Cuentas Corientes", vbCritical, "Final de la busqueda")
       TextCodigoCliente.Text = ""
       Call blanco
       Call nuevonumeroPago
       TextCodigoCliente.SetFocus
    Else
       TextSaldoLinea1.Text = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
       TextSaldoLinea2.Text = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
       TextSaldoTotal.Text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
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
        TextNumeroPago.Text = numeroPago + 1
    End If


       
End Sub





Private Sub OptionSaldoLinea1_Click()

    If OptionSaldoLinea1.Value = True Then
        saldoLinea2 = 0
        LabelSaldo.Caption = "Saldo Linea 1"
        BotonGuardar.Enabled = True
        Frame1.Enabled = True
        TextResta.Text = ""
        TextResta.Text = Format(TextSaldoLinea1.Text, "#0.00")
        If Val(TextSaldoTotal.Text) > 0 Then
            LabelSaldoTotal.ForeColor = vbRed
            LabelSaldoTotal.Caption = TextSaldoTotal.Text
        End If
        If Val(TextSaldoTotal.Text) < 0 Then
            LabelSaldoTotal.ForeColor = vbBlue
            LabelSaldoTotal.Caption = TextSaldoTotal.Text
        End If
        Call blancoCambioCheck
        saldoLinea1 = 1
    End If
    
End Sub

Private Sub OptionSaldoLinea1_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        KeyAscii = 0
        SendKeys "{TAB}"
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
        TextResta.Text = ""
        TextResta.Text = Format(TextSaldoLinea2.Text, "#0.00")
        If Val(TextSaldoTotal.Text) > 0 Then
            LabelSaldoTotal.ForeColor = vbRed
            LabelSaldoTotal.Caption = TextSaldoTotal.Text
        End If
        If Val(TextSaldoTotal.Text) < 0 Then
            LabelSaldoTotal.ForeColor = vbBlue
            LabelSaldoTotal.Caption = TextSaldoTotal.Text
        End If
        Call blancoCambioCheck
        saldoLinea2 = 2
    End If
    
End Sub

Private Sub OptionSaldoLinea2_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If
    
    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextCheque_Change()

     If TextCheque.Text <> "" Then
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
    TextCheque.SelLength = Len(TextCheque.Text)
End Sub

Private Sub TextCheque_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub



Private Sub Textcod_Change()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCliente = db.OpenRecordset("Clientes", dbOpenDynaset)
   

    
    CodigoClie = Val(Textcod.Text)
 
'    If KeyAscii = 13 Then
        If Textcod.Text = "" Then
            Textcod.SetFocus
        Else
            
            rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCtaCte.Fields!IdCliente <> Val(Textcod.Text) Then
                mensaje = MsgBox("No Posee Cuentas Corientes", vbCritical, "Final de la busqueda")
                Textcod.Text = ""
                Call blanco
                Call nuevonumeroPago
                Textcod.SetFocus
            Else
                LabelSaldoTotal.Caption = ""
                TextSaldoLinea1.Text = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
                If saldoLinea1 = 1 Then
                    TextResta.Text = Format(TextSaldoLinea1.Text, "#0.00")
                End If
                TextSaldoLinea2.Text = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
                TextSaldoTotal.Text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
                
                If Val(TextSaldoTotal.Text) > 0 Then
                    LabelSaldoTotal.ForeColor = vbRed
                    LabelSaldoTotal.Caption = TextSaldoTotal.Text
                End If
                If Val(TextSaldoTotal.Text) < 0 Then
                    LabelSaldoTotal.ForeColor = vbBlue
                    LabelSaldoTotal.Caption = TextSaldoTotal.Text
                End If
               
                'SendKeys "{TAB}"
                 'KeyAscii = 0
                 OptionSaldoLinea1.SetFocus
           End If
        End If
  
TextCodigoCliente.Text = Textcod.Text
    
  
End Sub


Private Sub TextCodigoCliente_GotFocus()

    TextCodigoCliente.SelLength = Len(TextCodigoCliente.Text)

End Sub

Private Sub TextCodigoCliente_KeyPress(KeyAscii As Integer)
  
    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set db = DBEngine.OpenDatabase(ruta)
    Set rstCtaCte = db.OpenRecordset("CtaCte", dbOpenDynaset)
    
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
                TextNombre.Text = rstCliente.Fields!RazonSocial
       
    
    
   
    CodigoClie = Val(TextCodigoCliente.Text)
 
'    If KeyAscii = 13 Then
        If TextCodigoCliente.Text = "" Then
            TextCodigoCliente.SetFocus
        Else
            
            rstCtaCte.FindFirst "IDCliente= " + Str(CodigoClie)
            If rstCtaCte.Fields!IdCliente <> Val(TextCodigoCliente.Text) Then
                mensaje = MsgBox("No Posee Cuentas Corientes", vbCritical, "Final de la busqueda")
                TextCodigoCliente.Text = ""
                Call blanco
                Call nuevonumeroPago
                TextCodigoCliente.SetFocus
            Else
                LabelSaldoTotal.Caption = ""
                TextSaldoLinea1.Text = Format(rstCtaCte.Fields!SaldoL1, "#0.00")
                If saldoLinea1 = 1 Then
                    TextResta.Text = Format(TextSaldoLinea1.Text, "#0.00")
                End If
                TextSaldoLinea2.Text = Format(rstCtaCte.Fields!SaldoL2, "#0.00")
                TextSaldoTotal.Text = Format(rstCtaCte.Fields!SaldoTotal, "#0.00")
                
                If Val(TextSaldoTotal.Text) > 0 Then
                    LabelSaldoTotal.ForeColor = vbRed
                    LabelSaldoTotal.Caption = TextSaldoTotal.Text
                End If
                If Val(TextSaldoTotal.Text) < 0 Then
                    LabelSaldoTotal.ForeColor = vbBlue
                    LabelSaldoTotal.Caption = TextSaldoTotal.Text
                End If
               
                'SendKeys "{TAB}"
                 'KeyAscii = 0
                 OptionSaldoLinea1.SetFocus
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

    If TextEfectivo.Text <> "" Then
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
    TextEfectivo.SelLength = Len(TextEfectivo.Text)
End Sub

Private Sub TextEfectivo_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub



Private Sub TextFechaPago_GotFocus()
    TextFechaPago.SelLength = Len(TextFechaPago.Text)
End Sub

Private Sub TextFechaPago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextMercaderia_Change()

    If TextMercaderia.Text <> "" Then
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

  
    TextEfectivo.Text = ""
    TextRezago.Text = ""
    TextMercaderia.Text = ""
    TextObservaciones.Text = ""
    TextCheque.Text = ""
    TextRetencion.Text = ""
    TextTarjeta.Text = ""
    
    
    
    

End Sub


Private Sub blanco()

    TextCodigoCliente.Text = ""
    TextNombre.Text = ""
    TextSaldoTotal.Text = 0
    TextSaldoLinea1.Text = 0
    TextSaldoLinea2.Text = 0
    'TextNumeroPago.Text = 0
    TextEfectivo.Text = ""
    TextRezago.Text = ""
    TextMercaderia.Text = ""
    TextObservaciones.Text = ""
    TextCheque.Text = ""
    TextRetencion.Text = ""
    TextTarjeta.Text = ""
    TextResta.Text = ""
    LabelSaldoTotal.Caption = ""
    TextResta.Text = ""
    LabelTotalAbonado.Caption = ""
    
    BotonGuardar.Enabled = False
    Frame1.Enabled = False
    OptionSaldoLinea1.Value = False
    OptionSaldoLinea2.Value = False
    
'    TextCodigoCliente.SetFocus
    
End Sub

Private Sub TextMercaderia_GotFocus()
    TextMercaderia.SelLength = Len(TextMercaderia.Text)
End Sub

Private Sub TextMercaderia_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextNombre_Change()
     Columna = 1
     Call FiltrarGrilla(MSFlexGrid1, TextNombre, CLng(Columna))
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
 Private Sub MSFlexGrid1_Click()
   
    
    MSFlexGrid1.Col = 0
    Textcod.Text = MSFlexGrid1.Text
    
    MSFlexGrid1.Col = 1
    TextNombre.Text = MSFlexGrid1.Text
    
   
   
    
    MSFlexGrid1.Visible = False
    
   

End Sub




Private Sub TextNumeroPago_GotFocus()
    TextNumeroPago.SelLength = Len(TextNumeroPago.Text)
End Sub

Private Sub TextNumeroPago_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub



Private Sub TextObservaciones_GotFocus()
    TextObservaciones.SelLength = Len(TextObservaciones.Text)
End Sub

Private Sub TextObservaciones_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub


Private Sub TextRetencion_Change()

    If TextRetencion.Text <> "" Then
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
    TextRetencion.SelLength = Len(TextRetencion.Text)
End Sub

Private Sub TextRetencion_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextRezago_Change()

    If TextRezago.Text <> "" Then
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
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
    End If
    
    If saldoLinea2 = 2 Then
            resta = CDec(TextSaldoLinea2.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
    End If
    
    
    TextResta.Text = Format(resta, "#00.00")
    LabelSaldoTotal.Caption = Format(resta2, "#00.00")
    
End Sub
Private Sub calculoresta()

        
        efectivo = Val(TextEfectivo.Text)
        If efectivo = 0 Then
            efectivo = 0
        End If

       
        rezago = Val(TextRezago.Text)
        If rezago = 0 Then
            rezago = 0
        End If

      
        mercaderia = Val(TextMercaderia.Text)
        If mercaderia = 0 Then
            mercaderia = 0
        End If

       
        cheque = Val(TextCheque.Text)
        If cheque = 0 Then
            cheque = 0
        End If

       
        retencion = Val(TextRetencion.Text)
        If retencion = 0 Then
            retencion = 0
        End If

       
        tarjeta = Val(TextTarjeta.Text)
        If tarjeta = 0 Then
            tarjeta = 0
        End If
       
    'resta = CDec(TextSaldoLinea1.Text) - CDec(TextEfectivo.Text) - CDec(TextRezago.Text) - CDec(TextMercaderia.Text) - CDec(TextCheque.Text) - CDec(TextRetencion.Text) - CDec(TextTarjeta.Text)

    If saldoLinea1 = 1 Then
        If efectivo = 0 Then
            resta = CDec(TextSaldoLinea1.Text) + CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) + CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If rezago = 0 Then
             resta = CDec(TextSaldoLinea1.Text) - CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If mercaderia = 0 Then
             resta = CDec(TextSaldoLinea1.Text) - CDec(efectivo) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If cheque = 0 Then
            resta = CDec(TextSaldoLinea1.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If retencion = 0 Then
            resta = CDec(TextSaldoLinea1.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
        End If
         If tarjeta = 0 Then
           resta = CDec(TextSaldoLinea1.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
        End If
    End If
    
    If saldoLinea2 = 2 Then
       If efectivo = 0 Then
            resta = CDec(TextSaldoLinea2.Text) + CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) + CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If rezago = 0 Then
             resta = CDec(TextSaldoLinea2.Text) - CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If mercaderia = 0 Then
             resta = CDec(TextSaldoLinea2.Text) - CDec(efectivo) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) + CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If cheque = 0 Then
            resta = CDec(TextSaldoLinea2.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) + CDec(cheque) - CDec(retencion) - CDec(tarjeta)
        End If
        If retencion = 0 Then
            resta = CDec(TextSaldoLinea2.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) + CDec(retencion) - CDec(tarjeta)
        End If
         If tarjeta = 0 Then
           resta = CDec(TextSaldoLinea2.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
            resta2 = CDec(TextSaldoTotal.Text) - CDec(efectivo) - CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) + CDec(tarjeta)
        End If
     
    End If
    
    TextResta.Text = Format(resta, "#00.00")
    LabelSaldoTotal.Caption = Format(resta2, "#00.00")

End Sub


Private Sub calculoabonado()

    suma = CDec(efectivo) + CDec(rezago) + CDec(mercaderia) + CDec(cheque) + CDec(retencion) + CDec(tarjeta)
   
    LabelTotalAbonado.Caption = Format(suma, "#00.00")
    
End Sub
Private Sub calculoabonadoresta()

    suma = CDec(efectivo) + CDec(rezago) - CDec(mercaderia) - CDec(cheque) - CDec(retencion) - CDec(tarjeta)
   
    LabelTotalAbonado.Caption = Format(suma, "#00.00")

End Sub

Private Sub TextRezago_GotFocus()
    TextRezago.SelLength = Len(TextRezago.Text)
End Sub

Private Sub TextRezago_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub

Private Sub TextTarjeta_Change()

    If TextTarjeta.Text <> "" Then
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
    TextTarjeta.SelLength = Len(TextTarjeta.Text)
End Sub

Private Sub TextTarjeta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

    If KeyAscii = 27 Then
        Unload Me
    End If

End Sub
