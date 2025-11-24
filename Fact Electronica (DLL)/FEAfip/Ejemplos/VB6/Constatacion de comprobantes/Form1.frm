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
   Begin VB.CommandButton Command1 
      Caption         =   "Constatar Comprobante"
      Height          =   735
      Left            =   1080
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Const URL_WSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
Const URL_WSCDC = "https://wswhomo.afip.gov.ar/WSCDC/service.asmx"

Dim lwscdc As New FEAFIPLib.wscdc
lwscdc.Depurar = True
lwscdc.CUIT = 20939802593#
lwscdc.URL = URL_WSCDC
If lwscdc.login("certificado.crt", "clave.key", URL_WSAA) Then
  If lwscdc.ComprobanteConstatar("CAE", 20939802593#, 140, 1, 1588, "20170517", 1452.73, "67203477090542", "80", "27929007862") Then
    MsgBox ("Comprobante constatado con éxito.")
  Else
    MsgBox (lwscdc.ErrorDesc)
  End If
Else
  MsgBox (lwscdc.ErrorDesc)
End If
End Sub

Private Sub Command2_Click()
End Sub
