VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FormActualizarCtaCte 
   Caption         =   "Actualizar Cuentas Corrientes"
   ClientHeight    =   3945
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   3945
   ScaleWidth      =   8550
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8295
      Begin MSComctlLib.ProgressBar PBar1 
         Height          =   255
         Left            =   1440
         TabIndex        =   3
         Top             =   1800
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
      End
      Begin VB.CommandButton cmdProcesarRegistros 
         Caption         =   "&Procesar Registros"
         Height          =   615
         Left            =   3120
         TabIndex        =   1
         Top             =   360
         Width           =   2055
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "GRACIAS POR UTILIZAR NUESTROS SERVICIOS SPC CONSULTING"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         TabIndex        =   5
         Top             =   2520
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registros Procesados: "
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
         TabIndex        =   4
         Top             =   1440
         Width           =   1980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Avance: "
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
         TabIndex        =   2
         Top             =   1440
         Width           =   780
      End
   End
End
Attribute VB_Name = "FormActualizarCtaCte"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdProcesarRegistros_Click()

    Dim SaldoL1, SaldoL2, SaldoTotal, CantReg As Double
    Dim tCtaCte, BaseSPC, Cliente, tMovCC
    
    SaldoLinea1 = 0
    SaldoLinea2 = 0
    
    SaldoL1 = 0
    SaldoL2 = 0
    SaldoTotal = 0
    
   
    'On Error GoTo Error_Handler
    
    Set BaseSPC = OpenDatabase(App.Path & "\DB_SPC_SI.mdb")
    Set tCtaCte = BaseSPC.OpenRecordset("CtaCte", dbOpenTable)
    
    tCtaCte.Index = "PrimaryKey"
    
    vSQL = "SELECT * FROM MovimientosCtaCte ORDER BY IdCliente, Fecha ASC"
    
    MsgBox (vSQL)
    
    Set tMovCC = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    tMovCC.MoveFirst
    tMovCC.MoveLast
    
    CantReg = tMovCC.RecordCount
    
    PBar1.Value = 0
    PBar1.Max = CantReg
                
                
    tMovCC.MoveFirst
    
    
    
    Cliente = tMovCC!IdCliente
        
        While Not tMovCC.EOF
          
          If Cliente = tMovCC!IdCliente Then
            
            SaldoL1 = SaldoL1 + tMovCC!ImporteLinea1
            SaldoL2 = SaldoL2 + tMovCC!ImporteLinea2
            tMovCC.MoveNext
            PBar1.Value = PBar1.Value + 1
            Label1.Caption = "Avance: % " + Format(((PBar1.Value * 100) / PBar1.Max), "Standard")
            Label2.Caption = "Registros Procesados: " + CStr(PBar1.Value)
           
           Else
             
             SaldoTotal = SaldoL1 + SaldoL2
             '** Actualizo CtaCte **
                tCtaCte.Seek "=", Cliente
                    If Not tCtaCte.NoMatch Then
                        tCtaCte.Edit
                            tCtaCte!IdCliente = Cliente
                            tCtaCte!SaldoL1 = SaldoL1
                            tCtaCte!SaldoL2 = SaldoL2
                            tCtaCte!SaldoTotal = SaldoTotal
                            tCtaCte!FechaActSaldo = Format(Date, "dd/mm/yyyy")
                        tCtaCte.Update
                    End If
             '***********************
             SaldoL1 = 0
             SaldoL2 = 0
             SaldoTotal = 0
             Cliente = tMovCC!IdCliente
          End If
        Wend
        
        Label3.Visible = True

End Sub

