Public Class Form1


    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        ' Los nombres de los parametros de las funciones se obtienen descomprimiendo FEAFIP DOC
        ' y luego abriendo el archivo index.html de la carpeta "Doc Interfaces".
        ' la interfaz correspondiente a este ejemplo es Iwsfev1 para facturas A y B.

        'URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
        'Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms 
        Const URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
        ' Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx

        ' Si esta linea falla es porqu eno agrego la referencia en a FEAFIPLib desde el menu de proyecto
        Dim wsfev1 As FEAFIPLib.wsfev1
        Dim nro As Double
        Dim Resultado As String
        Dim Reproceso As String
        Dim PtoVta As Integer
        Dim TipoComp As Integer
        Dim FechaComp As String
        Dim CAE As String
        Dim Vencimiento As String
        CAE = ""
        Vencimiento = ""
        Resultado = ""
        Reproceso = ""
        nro = 0
        PtoVta = 180
        TipoComp = 1 ' Factura A(ir a http://www.bitingenieria.com.ar/codigos.html)
        FechaComp = Date.Today.ToString("yyyyMMdd")

        wsfev1 = New FEAFIPLib.wsfev1
        wsfev1.CUIT = 20939802593 ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
                MsgBox(wsfev1.ErrorDesc)
            End If
            nro = nro + 1
            wsfev1.Reset()
            wsfev1.AgregaFactura(1, 80, 30707219072, nro, nro, FechaComp, 121, 0, 100, 0, "", "", "", "PES", 1)
            wsfev1.AgregaIVA(5, 100, 21) 'ir a http://www.bitingenieria.com.ar/codigos.html
            If wsfev1.Autorizar(PtoVta, TipoComp) Then
                wsfev1.AutorizarRespuesta(0, CAE, Vencimiento, Resultado, Reproceso)
                If Resultado = "A" Then
                    MsgBox("Felicitaciones! Si ve este mensaje instalo correctamente FEAFIP. CAE y Vencimiento: " & _
                         CAE + " " + Vencimiento)
                Else
                    MsgBox(wsfev1.AutorizarRespuestaObs(0))
                End If
            Else
                MsgBox(wsfev1.ErrorDesc)
            End If
        Else
            MsgBox(wsfev1.ErrorDesc)
        End If

    End Sub
End Class
