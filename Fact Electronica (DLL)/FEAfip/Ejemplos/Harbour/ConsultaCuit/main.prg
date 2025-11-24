    // Documentacion en http://bitingenieria.com.ar/doc/feafip/FEAFIPLib_TLB/IwsPadron.html 

    lwsPadron = CreateObject("FEAFIPLib.Padron")

    Contribuyente = CreateObject("FEAFIPLib.Contribuyente")
    
    cuitAConsultar = 20939802593
    
    IF lwsPadron:consultar(cuitAConsultar, Contribuyente)
        lbNombre = Contribuyente:nombre
        lbTipo = Contribuyente:tipoPersona
        lbEstado = Contribuyente:estadoClave
        lDomicilio = Contribuyente:domicilioFiscal
        IF lDomicilio:direccion != ""
            lbDomicilio = lDomicilio:direccion + ", " + lDomicilio:localidad + ", " + lDomicilio:provincia
            MessageBox(0, lbDomicilio)
        ENDIF
        && Solicito al cliente constancia porque no esta inscripto en ganancias
        IF lbConstancia
          lbConstancia = Contribuyente:SolicitarConstanciaInscripcion
        ENDIF
    ELSE
        MessageBox(0, lwsPadron:ErrorDesc)
    ENDIF