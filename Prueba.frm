VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form FormSaldosAFecha 
   Caption         =   "Saldos de Todos los Clientes a una Fecha Solicitada"
   ClientHeight    =   7995
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15255
   LinkTopic       =   "Form2"
   ScaleHeight     =   7995
   ScaleWidth      =   15255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   7455
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   14655
      Begin MSComCtl2.MonthView dateSelect 
         Height          =   2310
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   2670
         _ExtentX        =   4710
         _ExtentY        =   4075
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         StartOfWeek     =   121241601
         CurrentDate     =   43950
      End
      Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
         Height          =   6135
         Left            =   480
         TabIndex        =   4
         Top             =   840
         Width           =   13695
         _ExtentX        =   24156
         _ExtentY        =   10821
         _Version        =   393216
      End
      Begin VB.TextBox txtFecha 
         Alignment       =   2  'Center
         BackColor       =   &H8000000B&
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   480
         Width           =   1335
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
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
         TabIndex        =   1
         Top             =   480
         Width           =   600
      End
   End
End
Attribute VB_Name = "FormSaldosAFecha"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub dateSelect_DateClick(ByVal DateClicked As Date)
    
    txtFecha.Text = ""
    txtFecha.Text = dateSelect.Value
    
End Sub

Private Sub dateSelect_LostFocus()

    dateSelect.Visible = False

End Sub

Private Sub Form_Load()
            
    txtFecha.ToolTipText = "Doble Click Para Calendario"
            
End Sub

Private Sub txtFecha_Change()
'txtFecha.WhatsThisHelpID
    txtFecha.SelLength = Len(txtFecha.Text)
End Sub

Private Sub txtFecha_DblClick()

    dateSelect.Visible = True

End Sub

