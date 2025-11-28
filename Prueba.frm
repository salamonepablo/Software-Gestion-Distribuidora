VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form Prueba 
   Caption         =   "Prueba fecha"
   ClientHeight    =   4320
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7140
   LinkTopic       =   "Form2"
   ScaleHeight     =   4320
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   1440
         TabIndex        =   3
         Top             =   1200
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   116654081
         CurrentDate     =   41765
      End
      Begin VB.TextBox txtFecha 
         Height          =   375
         Left            =   1320
         TabIndex        =   2
         Top             =   600
         Width           =   2775
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   2415
      End
   End
End
Attribute VB_Name = "Prueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MonthView1_DblClick()
txtFecha = MonthView1.Value
End Sub

Private Sub txtFecha_Change()
'txtFecha.WhatsThisHelpID

End Sub
