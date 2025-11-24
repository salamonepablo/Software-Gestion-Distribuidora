        Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
        Const URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
        Dim nro
        Dim POS
        nro = 0
        POS = 50

        Set wsfev1 = CreateObject("FEAFIPLib.wsfev1")
        wsfev1.CUIT = 20939802593
        wsfev1.URL = URLWSW
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsfev1.SFRecuperaLastCMP(POS, 1) Then
               MsgBox wsfev1.ErrorDesc
            else
               nro = wsfev1.SFLastCmp
            End If
            nro = nro + 1
            wsfev1.Reset
            wsfev1.AgregaFactura 1, 80, 30702637895, nro, nro, "20110420", 121, 0, 100, 0, "", "", "", "PES", 1
            wsfev1.AgregaIVA 5, 100, 21
            If wsfev1.Autorizar(POS, 1) Then
                MsgBox wsfev1.SFCAE(0) + " " + wsfev1.SFVencimiento(0) 
                If wsfev1.SFResultado(0) <> "A" Then
                    MsgBox wsfev1.AutorizarRespuestaObs(0)
                Else
                    If Not wsfev1.SFCmpConsultar(1, POS, nro) Then
                        MsgBox wsfev1.ErrorDesc
                    End If
                End If

            Else
                MsgBox wsfev1.ErrorDesc
            End If
        Else
            MsgBox wsfev1.ErrorDesc
        End If

