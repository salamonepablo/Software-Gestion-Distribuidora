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
        fechacmp = transform(year(date()),"@L 9999") + transform(month(date()),"@L 99") + transform(day(date()),"@L 99") && Tomo la fecha actual como ejemplo
        
        wsfev1 = CreateObject("FEAFIPLib.wsfev1")
        wsfev1:CUIT = 20939802593
        wsfev1:URL = URLWSW
        
        If wsfev1:login("certificado.crt", "clave.key", URLWSAA) 
          If wsfev1:SFRecuperaLastCMP(Ptovta, Tipocomp) 
             nro = wsfev1:SFLastCmp &&Devolucion el ultimo comprobante
             If wsfev1:SFCmpConsultar(TipoComp, Ptovta, nro) // Consulto ultimo comprobante
                cbte = wsfev1:CmpConsultarCbte;
              // Ver propiedades de cbte en https://www.bitingenieria.com.ar/doc/feafip/FEAFIPLib_TLB.IComprobante.html
                CodAutorizacion = cbte:CodAutorizacion
                MessageBox(0, "Comprobante encontrado. CAE: " + CodAutorizacion, "FEAFIP", 0)
             Else
                MessageBox(0, wsfev1:ErrorDesc, "FEAFIP", 0)
             EndIf
          else
             MessageBox(0, wsfev1:ErrorDesc, "FEAFIP", 0)
          EndIf
        Else
           MessageBox(0, wsfev1:ErrorDesc, "FEAFIP", 0)
        EndIf