      PUBLIC nro as long
      PUBLIC POS as long
      PUBLIC tipo as long
 
        && Los nombres de los parametros de las funciones se obtienen en FEAFIP.pdf
        

          && URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          && Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
        URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
          && Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx

        Ptovta = 115 && ATENCION! SI RECIBE UN ERROR DE FECHA O NUMERO DE COMPROBANTE EN ESTA DEMO CAMBIE ESTE VALOR POR OTRO DE 1 A 9999
        Tipocomp = 1 &&Factura A(Ver excel referencias codigos AFIP)
        fechacmp = transform(year(date()),"@L 9999") + transform(month(date()),"@L 99") + transform(day(date()),"@L 99") && Tomo la fecha actual como ejemplo
        
        wsfev1 = CREATEOBJECT("FEAFIPLib.wsfev1")
        wsfev1.CUIT = 20939802593
        wsfev1.URL = URLWSW
        
        If wsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsfev1.SFRecuperaLastCMP(Ptovta, Tipocomp) Then
               MessageBox(wsfev1.ErrorDesc)
             else
               nro = wsfev1.SFLastCmp &&Devolucion el ultimo comprobante
           ENDIF
            nro = nro + 1
            wsfev1.Reset()
            wsfev1.AgregaFactura(1, 80, 30702637895, nro, nro, fechacmp, 121, 0, 100, 0, "", "", "", "PES", 1)
            wsfev1.AgregaIVA(5, 100, 21)  &&Ver excel referencias codigos AFIP
            If wsfev1.Autorizar(Ptovta, Tipocomp) Then
              If wsfev1.SFresultado(0)="A" Then
                MessageBox("Felicitaciones! Si ve este cartel es porque obtuvo CAE y Vencimiento. CAE:" + ;
                  wsfev1.SFCAE(0) + " Vencimiento: " + wsfev1.SFVencimiento(0))
              Else
                * observaciones
                MessageBox(wsfev1.AutorizarRespuestaObs(0))
              Endif
            Else
                MessageBox(wsfev1.ErrorDesc)
            EndIf
        Else
            MessageBox(wsfev1.ErrorDesc)
        EndIf
