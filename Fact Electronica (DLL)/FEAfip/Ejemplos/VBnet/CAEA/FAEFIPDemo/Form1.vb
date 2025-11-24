Public Class Form1

    'URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
    Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
    ' Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
    Const URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
    ' Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx
    Dim wsfev1 As FEAFIPLib.wsfev1 ' Si esta linea falla es porqu eno agrego la referencia en a FEAFIPLib desde el menu de proyecto

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Dim CAE As String = ""
        Dim Periodo As Int32 = 201404
        Dim Orden As Byte = 1
        Dim FechavigDesde As String = ""
        Dim FechaVigHasta As String = ""
        Dim FechaTope As String = ""
        Dim FechaProceso As String = ""

        wsfev1 = New FEAFIPLib.wsfev1
        wsfev1.CUIT = 20939802593.0#  ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then

            wsfev1.Reset()
            If Not wsfev1.CAEASolicitar(Periodo, Orden, CAE$, FechavigDesde, FechaVigHasta, FechaTope, FechaProceso) Then
                MsgBox(wsfev1.ErrorDesc)
            Else
                Button2.Enabled = True
                Label1.Text = CAE
                Button2.Enabled = True
            End If
        Else
            MsgBox(wsfev1.ErrorDesc)
        End If

    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        wsfev1 = New FEAFIPLib.wsfev1

        wsfev1.CUIT = 20939802593.0#  ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
            Dim CAE As String = ""
            Dim Periodo As Int32 = 201403
            Dim Orden As Byte = 2
            Dim FechavigDesde As String = ""
            Dim FechaVigHasta As String = ""
            Dim FechaTope As String = ""
            Dim FechaProceso As String = ""

            wsfev1.Reset()
            If Not wsfev1.CAEAConsultar(Periodo, Orden, CAE, FechavigDesde, FechaVigHasta, FechaTope, FechaProceso) Then
                MsgBox(wsfev1.ErrorDesc)
            Else
                Button2.Enabled = True
                Label1.Text = CAE
                Button2.Enabled = True
            End If
        Else
            MsgBox(wsfev1.ErrorDesc)
        End If
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        Dim CAE As String = ""
        Dim Vencimiento As String = ""
        Dim Resultado As String = ""
        Dim Reproceso As String = ""
        Dim nro As Int32 = 0
        Dim PtoVta As Int32 = 10
        Dim FechaComp As String = Format(Now(), "yyyymmdd")
        Dim TipoComp As Int32 = 1 ' Factura A(Ver excel referencias codigos AFIP documentacion.rar o ir a http://www.bitingenieria.com.ar/codigos.html)

        wsfev1 = New FEAFIPLib.wsfev1

        wsfev1.CUIT = 20939802593.0#  ' Cuit del vendedor
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
                MsgBox(wsfev1.ErrorDesc)
            Else
                nro = nro + 1
                wsfev1.AgregaFactura(1, 80, 30707219072.0#, nro, nro, FechaComp, 121, 0, 100, 0, "", "", "", "PES", 1)
                wsfev1.AgregaIVA(5, 100, 21) ' Ver Excel de referencias de codigos AFIP
                If wsfev1.CAEAInformar(PtoVta, TipoComp, Label1.Text) Then
                    MsgBox("Felicitaciones! Si ve este mensaje instalo correctamente FEAFIP. CAE: " + Label1.Text)
                Else
                    MsgBox(wsfev1.ErrorDesc)
                End If
            End If
        Else
            MsgBox(wsfev1.ErrorDesc)
        End If

    End Sub
End Class
