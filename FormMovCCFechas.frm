VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form MovimientoCCFecha 
   Caption         =   "Movimientos de Cta Cte entre Fechas"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5850
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Calculo de Saldos Entre Fechas"
      Height          =   5535
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   9015
      Begin VB.CommandButton cmdMovimientos 
         Caption         =   "&Ver Movimientos"
         Height          =   375
         Left            =   5760
         TabIndex        =   4
         Top             =   1200
         Width           =   1935
      End
      Begin VB.ComboBox cmbIdCliente 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   480
         Width           =   1215
      End
      Begin VB.Frame Frame2 
         Caption         =   "Saldo Final"
         Height          =   855
         Left            =   360
         TabIndex        =   14
         Top             =   4440
         Width           =   7935
         Begin VB.TextBox txtSaldoTotal 
            Height          =   315
            Left            =   6240
            TabIndex        =   8
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtL2 
            Height          =   315
            Left            =   3600
            TabIndex        =   7
            Top             =   360
            Width           =   1215
         End
         Begin VB.TextBox txtL1 
            Height          =   315
            Left            =   1200
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Total:"
            Height          =   195
            Left            =   5280
            TabIndex        =   17
            Top             =   360
            Width           =   855
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Linea 2:"
            Height          =   195
            Left            =   2880
            TabIndex        =   16
            Top             =   360
            Width           =   570
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Linea 1:"
            Height          =   195
            Left            =   480
            TabIndex        =   15
            Top             =   360
            Width           =   570
         End
      End
      Begin VB.TextBox txtFechaHasta 
         Height          =   315
         Left            =   3720
         TabIndex        =   3
         Top             =   1200
         Width           =   1455
      End
      Begin VB.TextBox txtFechaDesde 
         Height          =   315
         Left            =   1200
         TabIndex        =   2
         Top             =   1200
         Width           =   1335
      End
      Begin VB.TextBox txtRSCliente 
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Top             =   480
         Width           =   5775
      End
      Begin MSFlexGridLib.MSFlexGrid FG1 
         Height          =   2175
         Left            =   360
         TabIndex        =   5
         Top             =   2160
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         FocusRect       =   0
      End
      Begin VB.Label lblSaldoAnterior 
         AutoSize        =   -1  'True
         Caption         =   "Saldo Anterior:"
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
         Width           =   1275
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
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
         TabIndex        =   12
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
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
         TabIndex        =   11
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cliente:"
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
         TabIndex        =   10
         Top             =   480
         Width           =   660
      End
   End
End
Attribute VB_Name = "MovimientoCCFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub SeteoGrilla()
   
    FG1.FixedRows = 1
    
    FG1.Cols = 8
    FG1.Rows = 2
    
    FG1.Row = 0
    
    FG1.Col = 0
    FG1.Text = "Fecha Mov."
    FG1.Col = 1
    FG1.Text = "Tipo Doc."
    FG1.Col = 2
    FG1.Text = "Nº Doc."
    FG1.Col = 3
    FG1.Text = "Importe L1"
    FG1.Col = 4
    FG1.Text = "Importe L2"
    FG1.Col = 5
    FG1.Text = "Saldo L1"
    FG1.Col = 6
    FG1.Text = "Saldo L2"
    FG1.Col = 7
    FG1.Text = "Saldo Total"
    
End Sub

Private Sub cmbIdCliente_GotFocus()
    cmbIdCliente.SelLength = Len(cmbIdCliente.Text)
End Sub

Private Sub cmbIdCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub cmdMovimientos_Click()
    
    FG1.Clear
    
    Dim SaldoL1, SaldoL2, SaldoTotal As Double
    
    SaldoL1 = 0
    SaldoL2 = 0
    SaldoTotal = 0
    
   On Error GoTo CapturaErrores:
    
    vSQL = "SELECT * FROM MovimientosCtaCte WHERE IDCliente =" & cmbIdCliente.Text & " ORDER BY Fecha"
    
    'MsgBox (vSQL)
    
    Set tMovCC = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    tMovCC.MoveFirst
    
    
    While CVDate(tMovCC!Fecha) < CVDate(txtFechaDesde.Text)
        SaldoL1 = SaldoL1 + tMovCC!ImporteLinea1
        SaldoL2 = SaldoL2 + tMovCC!ImporteLinea2
        tMovCC.MoveNext
    Wend
    SaldoTotal = SaldoL1 + SaldoL2
    
    lblSaldoAnterior.Caption = " Línea 1: $ " + CStr(CCur(SaldoL1)) + "     Línea 2: $ " + CStr(CCur(SaldoL2)) + "     Saldo Anterior Consolidado: $ " + CStr(CCur(SaldoTotal))
    
    Call SeteoGrilla
    
    While Not tMovCC.EOF
        If (CVDate(tMovCC!Fecha) <= CVDate(txtFechaHasta.Text)) Then
            'Lleno Grilla
            FG1.Row = FG1.Row + 1
            FG1.Col = 0
            FG1.Text = tMovCC!Fecha
            FG1.Col = 1
            FG1.Text = tMovCC!tipoDoc
            FG1.Col = 2
            FG1.Text = tMovCC!NroDoc
            FG1.Col = 3
            FG1.Text = tMovCC!ImporteLinea1
            FG1.Col = 4
            FG1.Text = tMovCC!ImporteLinea2
            FG1.Col = 5
            SaldoL1 = SaldoL1 + tMovCC!ImporteLinea1
            FG1.Text = SaldoL1
            FG1.Col = 6
            SaldoL2 = SaldoL2 + tMovCC!ImporteLinea2
            FG1.Text = SaldoL2
            FG1.Col = 7
            SaldoTotal = SaldoL1 + SaldoL2
            FG1.Text = SaldoTotal
            
            FG1.Rows = FG1.Rows + 1
        End If
        tMovCC.MoveNext
    Wend
    
    txtL1.Text = SaldoL1
    txtL2.Text = SaldoL2
    txtSaldoTotal.Text = SaldoTotal

CapturaErrores:
    Select Case Err
        Case 3021
            MsgBox "No hay registros !!!", vbCritical + vbDefaultButton1, "INFO DEL SISTEMA"
            Resume Next
    End Select
End Sub

Private Sub Form_Load()
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    
    Set tClientes = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
    tClientes.MoveFirst
    
    While Not tClientes.EOF
        cmbIdCliente.AddItem (tClientes!IdCliente)
        tClientes.MoveNext
    Wend
        
    txtFechaDesde.Text = Format(Date, "DD/MM/YYYY")
    txtFechaHasta.Text = Format(Date, "DD/MM/YYYY")
    
        
End Sub

Private Sub Text1_Change()

End Sub

Private Sub txtFechaDesde_GotFocus()

    txtFechaDesde.SelLength = Len(txtFechaDesde.Text)

End Sub

Private Sub txtFechaDesde_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub txtFechaHasta_GotFocus()

    txtFechaHasta.SelLength = Len(txtFechaHasta.Text)

End Sub

Private Sub txtFechaHasta_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


Private Sub txtRSCliente_GotFocus()

    tClientes.Index = "PrimaryKey"
    tClientes.Seek "=", cmbIdCliente.Text

    If Not tClientes.NoMatch Then txtRSCliente.Text = tClientes!RazonSocial

End Sub

Private Sub txtRSCliente_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = 13 Then
            KeyAscii = 0
            SendKeys "{TAB}"
    End If

End Sub


