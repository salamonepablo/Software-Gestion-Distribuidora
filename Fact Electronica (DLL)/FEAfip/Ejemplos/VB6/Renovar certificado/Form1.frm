VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo FEAFIP"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Renovar Certificado (Grafico)"
      Height          =   735
      Left            =   1080
      TabIndex        =   1
      Top             =   480
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Renovar Certificado"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   1920
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim Certificado As New FEAFIPLib.Certificado
  If Certificado.CargarInformacionCertificado("certificado.crt", "clave.key") Then
    If Certificado.RenovarCertificado("pedido.csr") Then
      MsgBox ("Certificado renovado exitosamente")
    Else
      MsgBox (Certificado.ErrorDesc)
    End If
  Else
    MsgBox (certificadoMgr.ErrorDesc)
  End If
End Sub

Private Sub Command2_Click()
  Dim Certificado As New FEAFIPLib.Certificado
  If Certificado.CargarInformacionCertificado("certificado.crt", "clave.key") Then
    Certificado.MostrarInformacionCertificado
  Else
    MsgBox (certificadoMgr.ErrorDesc)
  End If
End Sub
