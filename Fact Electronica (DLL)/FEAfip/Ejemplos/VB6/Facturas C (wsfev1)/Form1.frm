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
      Caption         =   "Iniciar Demo"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
        ' Los nombres de los parametros de las funciones se obtienen en FEAFIP.pdf
        
        'URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          ' Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
        Const URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
          ' Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx
        Dim wsfev1 As FEAFIPLib.wsfev1 ' Si esta linea falla es porqu eno agrego la referencia en a FEAFIPLib desde el menu de proyecto
        Dim nro As Double
        CAE$ = ""
        Vencimiento$ = ""
        Resultado$ = ""
        Reproceso$ = ""
        nro = 0
        PtoVta = 10  ' ATENCION! SI RECIBE UN ERROR DE FECHA O NUMERO DE COMPROBANTE EN ESTA DEMO CAMBIE ESTE VALOR POR OTRO DE 1 A 9999
        TipoComp = 11 ' Factura C(Ver codigos AFIP en excel de códigos AFIP)
        FechaComp = Format(Now(), "yyyymmdd")
         
        Set wsfev1 = New FEAFIPLib.wsfev1
        wsfev1.CUIT = 20162604032#  ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
                MsgBox (wsfev1.ErrorDesc)
            Else
                nro = nro + 1
                wsfev1.Reset
                wsfev1.AgregaFactura 1, 80, 30707219072#, nro, nro, FechaComp, 150, 0, 150, 0, "", "", "", "PES", 1
                If wsfev1.Autorizar(PtoVta, TipoComp) Then
                    wsfev1.AutorizarRespuesta 0, CAE, Vencimiento, Resultado, Reproceso
                    If Resultado = "A" Then
                        MsgBox "Felicitaciones! Si ve este mensaje instalo correctamente FEAFIP. CAE y Vencimiento: " + CAE + " " + Vencimiento
                    Else
                        MsgBox wsfev1.AutorizarRespuestaObs(0)
                    End If
    
                Else
                    MsgBox wsfev1.ErrorDesc
                End If
            End If
        Else
            MsgBox wsfev1.ErrorDesc
        End If

End Sub

