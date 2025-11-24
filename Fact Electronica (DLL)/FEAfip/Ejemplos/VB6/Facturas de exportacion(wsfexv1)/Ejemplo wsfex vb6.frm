VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
        Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
        Const URLWSW = "https://wswhomo.afip.gov.ar/wsfexv1/service.asmx"

Private Sub Form_Load()
        Dim wsfex As FEAFIPLib.wsfexv1
        Set wsfex = New FEAFIPLib.wsfexv1
        wsfex.CUIT = 20939802593#
        wsfex.URL = URLWSW
        currdate = Format(Now(), "yyyymmdd")
        
        If wsfex.login("certificado.crt", "clave.key", URLWSAA) Then
 
            Dim nro As Double, fecha As String, IdTrans As Double
            PtoVta = 100
            tipocmp = 19
            
            
            If Not wsfex.RecuperaLastCMP(PtoVta, tipocmp, nro, fecha) Then
                MsgBox wsfex.ErrorDesc
            End If
            nro = nro + 1
            If Not wsfex.UltimoIdTrans(IdTrans) Then
                MsgBox wsfex.ErrorDesc
            End If
            
            IdTrans = IdTrans + 1
            wsfex.AgregaFactura IdTrans, currdate, tipocmp, PtoVta, nro, 1, "N", 208, "chile sa", 50000000032#, "Domicilio", "", "DOL", 18, "", 100, "", "contado", "FOB", "", 1, ""
            wsfex.AgregaItem "11111", "remera ", 1, 1, 100, 100, 0
            If wsfex.Autorizar Then
 
                Dim CAE As String, Vencimiento As String, resultado As String, Reproceso As String
                wsfex.AutorizarRespuesta CAE, Vencimiento, resultado, Reproceso
                ' La variable reproceso indica que no se autorizo el comprobante sino que se obtuvo un CAE ya asignado previamente. TransId debe ser unico para cada nuevo comprobante para evitar esto
                MsgBox "Felicitaciones! si ve este mensaje es porque acaba de obtener el CAE: " + CAE + " " + Vencimiento + " " + Reproceso
            Else
                MsgBox wsfex.ErrorDesc
 
            End If
        Else
            MsgBox wsfex.ErrorDesc
        End If
 

End Sub
