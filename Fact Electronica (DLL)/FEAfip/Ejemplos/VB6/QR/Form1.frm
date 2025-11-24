VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Generar QR"
      Height          =   615
      Left            =   1320
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim Qr As FEAFIPLib.Qr
  Set Qr = New FEAFIPLib.Qr
  Qr.ArchivoQR = Qr.RutaLibreria + "qr.bmp" ' Admite formatos BMP, PNG y JPG con solo cambiar la extension

  ver = 1
  fecha = ""
  CUIT = 20939802593#
  PtoVta = 2
  tipoComp = 1
  nroCmp = 1
  Importe = 100.2
  moneda = "PES"
  ctz = 1#
  tipoDocRec = 80
  nroDocRec = 27929007862#
  tipoCodAut = "E"  ' A = CAEA E = CAE
  codAut = 12345678901234#
  If Qr.Generar(ver, fecha, CUIT, PtoVta, tipoComp, nroCmp, Importe, moneda, ctz, tipoDocRec, nroDocRec, tipoCodAut, codAut) Then
    MsgBox ("QR generado con éxito en " + Qr.ArchivoPNG)
  Else
    MsgBox (Qr.ErrorDesc)
  End If

End Sub
