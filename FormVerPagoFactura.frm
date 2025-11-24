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
                rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago Nº " & TextNumeroPago.text & " Linea 1"
                rstMovimientosCtaCte.Fields!ImporteLinea1 = Format(LabelTotalAbonado.Caption, "#,###,###,#0.00")
                rstMovimientosCtaCte.Fields!ImporteLinea2 = 0
            End If
            If OptionSaldoLinea2.Value = True Then
                rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago Nº " & TextNumeroPago.text & " Linea 2"
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
'                rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago Nº " & TextNumeroPago.Text & " Linea 1"
'                rstMovimientosCtaCte.Fields!ImporteLinea1 = Format(LabelTotalAbonado.Caption, "#,###,###,#0.00")
'                rstMovimientosCtaCte.Fields!ImporteLinea2 = 0
'            End If
'            If OptionSaldoLinea2.Value = True Then
'                rstMovimientosCtaCte.Fields!tipoDoc = "Anulacion Pago Nº " & TextNumeroPago.Text & " Linea 2"
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
   
    ruta = App.Path & "\DB_SPC_SI.mdb"

    Set db = DBEngine.OpenDatabase(ruta)

    Set tSucursales = db.OpenRecordset("Sucursales", dbOpenTable)

    tSucursales.MoveFirst

    Do Until tSucursales.EOF
        cmbSucursal.AddItem (tSucursales!IdSucursal & " - " & tSucursales!nombreSucursal)
        tSucursales.MoveNext
    Loop

    cmbSucursal.ListIndex = 1

    tSucursales.Close
    db.Close
   
    Call buscodatos
    
End Sub
Private Sub buscodatos()

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

Private Sub TextCheque_Change()

     If TextCheque.text <> "" Then
        Call calculoabonado
    End If

End Sub

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


