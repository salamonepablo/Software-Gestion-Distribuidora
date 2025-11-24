        // Los nombres de los parametros de las funciones se obtienen descomprimiendo FEAFIP DOC
        // y luego abriendo el archivo index.html de la carpeta "Doc Interfaces".
        // la interfaz correspondiente a este ejemplo es Iwsfev1 para facturas A y B.

          // URLs de autenticacion y negocio. Cambiarlas por las de producci=n al implementarlas en el cliente(abajo)
        URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          // Producci=n: https://wsaa.afip.gov.ar/ws/services/LoginCms
        URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
          // Producci=n: https://servicios1.afip.gov.ar/wsfev1/service.asmx

        Ptovta = 110
        Tipocomp = 1 //Factura A(Ver excel referencias codigos AFIP documentacion.rar)
        fechacmp = transform(year(date()),"@L 9999") + transform(month(date()),"@L 99") + transform(day(date()),"@L 99") // Tomo la fecha actual como ejemplo
        
        wsfev1 = CreateObject("FEAFIPLib.wsfev1")
        wsfev1:CUIT = 20939802593
        wsfev1:URL = URLWSW
        
        If wsfev1:login("certificado.crt", "clave.key", URLWSAA) 
            If wsfev1:SFRecuperaLastCMP(Ptovta, Tipocomp) 
               nro = wsfev1:SFLastCmp //Devolucion el ultimo comprobante
             else
               MessageBox(0, wsfev1:ErrorDesc, "FEAFIP", 0)
           ENDIF
            nro = nro + 1
            wsfev1:Reset()
            wsfev1:AgregaFactura(1, 80, 30702637895, nro, nro, fechacmp, 121, 0, 100, 0, "", "", "", "PES", 1)
            wsfev1:AgregaIVA(5, 100, 21)  //Ver excel referencias codigos AFIP documentacion.rar
            If wsfev1:Autorizar(Ptovta, Tipocomp) 
              If wsfev1:SFresultado(0)="A" 
                MessageBox(0, "Felicitaciones! Si ve este cartel es porque obtuvo CAE y Vencimiento. CAE:" + ;
                  wsfev1:SFCAE(0) + " Vencimiento: " + wsfev1:SFVencimiento(0),"FEAFIP", 0)
              Else
                * observaciones
               MessageBox(0, wsfev1:ErrorDesc, "FEAFIP", 0)
              Endif
            Else
               MessageBox(0, wsfev1:ErrorDesc, "FEAFIP", 0)
            EndIf
        Else
               MessageBox(0, wsfev1:ErrorDesc, "FEAFIP", 0)
        EndIf