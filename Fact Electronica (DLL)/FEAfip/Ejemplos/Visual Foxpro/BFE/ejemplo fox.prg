
        Punto_vta = 115
        tipo_cbte = 1 &&Factura A(Ver excel referencias codigos AFIP documentacion.rar)
        zona = 1
        fechacmp = transform(year(date()),"@L 9999") + transform(month(date()),"@L 99") + transform(day(date()),"@L 99") && Tomo la fecha actual como ejemplo
  
        
        wsbfev1 = CREATEOBJECT("FEAFIPLib.wsbfev1")
        wsbfev1.CUIT = 20939802593
        wsbfev1.URL = "https://wswhomo.afip.gov.ar/wsbfev1/service.asmx"
        
        If wsbfev1.login("certificado.crt", "clave.key", "https://wsaahomo.afip.gov.ar/ws/services/LoginCms")
            If Not wsbfev1.SFRecuperaLastCMP(Punto_vta, tipo_cbte) Then
               MessageBox(wsbfev1.ErrorDesc) &&
            Else
              Cbt_nro = wsbfev1.SFLastCmp &&Devolucion el ultimo comprobante autorizado
              Cbt_nro = Cbt_nro + 1
              If Not wsbfev1.SFRecuperaLastID() Then
                MessageBox(wsbfev1.ErrorDesc) 
              Else
                idtrans = wsbfev1.SFLastID
                idtrans = idtrans + 1
                wsbfev1.Reset()
                wsbfev1.AgregaFactura(idtrans, 80, 30702637895, zona, tipo_cbte, Punto_vta, Cbt_nro, 121, 0, 100, 21, 0, 0, 0, 0, 0, 0, "PES", 1, fechacmp, "")
                wsbfev1.AgregaItem("0209.00.1", "0209.00.1", "Producto 1", 1, 1,  100, 0, 121, 5)  &&Ver excel documentacion de interfaces para conocer los parametros
                If wsbfev1.Autorizar() Then
                  If wsbfev1.SFResultado =  "A" Then
                    MessageBox("Felicitaciones! Si puede ver este mensaje es porque pudo obtener CAE y Vencimiento. CAE:" + wsbfev1.SFCAE + " Vencimiento:" + wsbfev1.SFVencimiento)
                  Else
                    * observaciones
                    MessageBox(wsbfev1.AutorizarRespuestaObs)
                  Endif
                Else
                  MessageBox(wsbfev1.ErrorDesc)
                EndIf
              EndIf
            EndIf
        Else
            MessageBox(wsbfev1.ErrorDesc)
        EndIf
