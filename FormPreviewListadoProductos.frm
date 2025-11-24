VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FormReporteFacxFech 
   Caption         =   "Reporte Facturas por Fecha"
   ClientHeight    =   4065
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5880
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   5880
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtHasta 
      Height          =   375
      Left            =   4200
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.TextBox txtDesde 
      Height          =   375
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton BtnGeneraReporte 
      Caption         =   "Generar Reporte"
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   3480
      Width           =   3255
   End
   Begin MSComCtl2.MonthView MonthView2 
      Height          =   2370
      Left            =   3120
      TabIndex        =   3
      Top             =   840
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   16908289
      CurrentDate     =   41841
   End
   Begin MSComCtl2.MonthView MonthView1 
      Height          =   2370
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      Appearance      =   1
      StartOfWeek     =   16908289
      CurrentDate     =   41765
   End
   Begin VB.Line Line1 
      X1              =   2880
      X2              =   2880
      Y1              =   240
      Y2              =   3360
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Hasta"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha Desde"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "FormReporteFacxFech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbtemp As DAO.Database
Dim rstFacturaC As DAO.Recordset

Private Sub BtnGeneraReporte_Click()

    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set dbtemp = DBEngine.OpenDatabase(ruta)
    
    FechaDesde = "#" + Format(txtDesde.Text, "DD/MM/YYYY") + "#"
    FechaHasta = "#" + Format(txtHasta.Text, "DD/MM/YYYY") + "#"
    vSQL = "SELECT * FROM FacturaC WHERE FechaFactura >=" & FechaDesde & " AND FechaFactura <=" & FechaHasta & " ORDER BY NroFactura"
    
    MsgBox (vSQL)
    Set tfacturaC = dbtemp.OpenRecordset(vSQL, dbOpenDynaset)
    
    tfacturaC.MoveFirst
    MsgBox (tfacturaC!NroFactura)
        
    
End Sub

Private Sub MonthView1_DateClick(ByVal DateClicked As Date)
    txtDesde = MonthView1.Value
End Sub

Private Sub MonthView2_DateClick(ByVal DateClicked As Date)
    txtHasta = MonthView2.Value
End Sub
