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
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   3480
         Width           =   255
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   16777217
         CurrentDate     =   41765
      End
      Begin VB.TextBox txtFecha 
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "HOY"
         Height          =   255
         Left            =   720
         TabIndex        =   6
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   3000
      TabIndex        =   5
      Top             =   1920
      Width           =   1215
   End
End
Attribute VB_Name = "Prueba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
txtFecha.Text = Date
End Sub

Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)
Private Sub Form_KeyPress(KeyAscii As Integer)

If KeyAscii = 27 Then
Unload Me
End If
End Sub
End Sub

Private Sub MonthView1_DblClick()
txtFecha = MonthView1.Value
End Sub

Private Sub txtFecha_Change()

If txtFecha <> "" Then
   If IsNumeric(txtFecha.Text) Then
        If Len(txtFecha.Text) = 8 Then
            txtFecha.Text = Left(txtFecha.Text, 2) + "/" + Mid(txtFecha.Text, 3, 2) + "/" + Right(txtFecha.Text, 4)
        End If
     Else
    'MsgBox "Error No es un Numero", vbCritical, "ERROR"
    End If
    
 Else
    'MsgBox "Mal Formato de Fecha", vbCritical, "ERROR"
 End If

End Sub
