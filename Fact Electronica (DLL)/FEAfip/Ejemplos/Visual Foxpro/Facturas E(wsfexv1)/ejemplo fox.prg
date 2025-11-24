      PUBLIC nro as long

        && Los nombres de los parametros de las funciones se obtienen descomprimiendo FEAFIP DOC
        && y luego abriendo el archivo index.html de la carpeta "Doc Interfaces".
        && la interfaz correspondiente a este ejemplo es Iwsfev1 para facturas A y B.

          && URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          && Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
        URLWSW = "https://wswhomo.afip.gov.ar/wsfexv1/service.asmx"
          && Producción: https://servicios1.afip.gov.ar/wsfexv1/service.asmx

        Ptovta = 110
        Tipocomp = 19 &&Factura A(Ver excel referencias codigos AFIP)
        fechacmp = transform(year(date()),"@L 9999") + transform(month(date()),"@L 99") + transform(day(date()),"@L 99") && Tomo la fecha actual como ejemplo
        
        wsfex = createobject("FEAFIPLib.wsfexv1") 
        wsfex.CUIT = 20939802593
        wsfex.URL = URLWSW
        
        If wsfex.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsfex.SFRecuperaLastCMP(PtoVta, Tipocomp) Then
                MsgBox wsfex.ErrorDesc
            EndIf
            nro = wsfex.SFLastCMP + 1
            If Not wsfex.SFUltimoIdTrans Then
                MsgBox wsfex.ErrorDesc
            EndIf
            IdTrans = wsfex.SFLastId + 1
            wsfex.AgregaFactura(IdTrans, fechacmp, 19, Ptovta, nro, 1, "N", 208, "chile sa", 50000000032, "Direccion", "", "DOL", 84, "", 100, "", "contado", "FOB", "", 1, "")
            wsfex.AgregaItem("11111", "remera ", 1, 1, 100, 100, 0)
            If wsfex.Autorizar Then
                && La variable reproceso indica que no se autorizo el comprobante sino que se obtuvo un CAE ya asignado previamente. TransId debe ser unico para cada nuevo comprobante para evitar esto
                MessageBox("Felicitaciones, si ve este mensaje es porque pudo obtener el CAE: " + wsfex.SFCAE + " " + wsfex.SFVencimiento + " " + wsfex.SFReproceso)
            Else
                MessageBox(wsfex.ErrorDesc)
 
            EndIf
        Else
            MessageBox(wsfex.ErrorDesc)
        EndIf
