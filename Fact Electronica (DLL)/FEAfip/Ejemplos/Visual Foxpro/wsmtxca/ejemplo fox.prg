      PUBLIC nro as long
      PUBLIC POS as long
      PUBLIC tipo as long
      PUBLIC  wsmtxca
      

        && Los nombres de los parametros de las funciones se obtienen descomprimiendo FEAFIP DOC
        && y luego abriendo el archivo index.html de la carpeta "Doc Interfaces".
        && la interfaz correspondiente a este ejemplo es Iwsfev1 para facturas A y B.

          && URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          && Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
        URLWSW = "https://fwshomo.afip.gov.ar/wsmtxca/services/MTXCAService"
          && Producción: https://serviciosjava.afip.gob.ar/wsmtxca/services/MTXCAService

        Ptovta = 110
        Tipocomp = 1 &&Factura A(Ver excel referencias codigos AFIP documentacion.rar)
        fechacmp = transform(year(date()),"@L 9999") + transform(month(date()),"@L 99") + transform(day(date()),"@L 99") && Tomo la fecha actual como ejemplo
        
        wsmtxca = CREATEOBJECT("FEAFIPLib.wsmtxca")
        wsmtxca.CUIT = 20939802593
        wsmtxca.URL = URLWSW
        
        If wsmtxca.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsmtxca.SFRecuperaLastCMP(Ptovta, Tipocomp) Then
               MessageBox(wsmtxca.ErrorDesc)
             else
               nro = wsmtxca.SFLastCmp &&Devolucion el ultimo comprobante
           ENDIF
            nro = nro + 1
            && codigoTipoComprobante, numeroPuntoVenta, numeroComprobante,fechaEmision, codigoTipoDocumento, numeroDocumento, importeGravado, importeNoGravado, importeExento, importeSubtotal, importeOtrosTributos, importeTotal, codigoMoneda, cotizacionMoneda, observaciones, codigoConcepto, fechaServicioDesde, fechaServicioHasta, fechaVencimientoPago
            wsmtxca.AgregaFactura(Tipocomp, PtoVta, nro, fechacmp, 80, 30702637895, 100, 0, 0, 100, 0, 121, "PES", 1, "", 1, "", "", "")
            wsmtxca.AgregaIVA(5, 21)  &&Ver excel referencias codigos AFIP documentacion.rar
            && unidadesMtx, codigoMtx, codigo, descripcion, cantidad, codigoUnidadMedida, precioUnitario, importeBonificacion, codigoCondicionIVA, importeIVA, importeItem
            wsmtxca.AgregaItem(1, "articulo1", "articulo1", "descripcion arti 1", 1, 1, 100, 0, 5, 21, 121)
            If wsmtxca.Autorizar() Then
              If wsmtxca.SFresultado="A" Then
                MessageBox("Felicitaciones! Si ve este cartel es porque obtuvo CAE y Vencimiento. CAE:" + ;
                  wsmtxca.SFCAE + " Vencimiento: " + wsmtxca.SFVencimiento)
              Else
                * observaciones
                MessageBox(wsmtxca.AutorizarRespuestaObs)
              Endif
            Else
                MessageBox(wsmtxca.ErrorDesc)
            EndIf
        Else
            MessageBox(wsmtxca.ErrorDesc)
        EndIf
