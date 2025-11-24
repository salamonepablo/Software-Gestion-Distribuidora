VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo FEAFIP"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Iniciar Demo"
      Height          =   735
      Left            =   1320
      TabIndex        =   0
      Top             =   1560
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
      Dim nro As Long
      Dim POS As Long
      Dim tipo As Long
      Dim wsmtxca As FEAFIPLib.wsmtxca
      

        'Este servicio solo debe implementarse si recibe una carta documento de AFIP

          ' URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          ' Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
        URLWSW = "https://fwshomo.afip.gov.ar/wsmtxca/services/MTXCAService"
          ' Producción: https://serviciosjava.afip.gob.ar/wsmtxca/services/MTXCAService

        PtoVta = 110
        Tipocomp = 1 'Factura A(Ver excel referencias codigos AFIP documentacion.rar)
        fechacmp = Format(Now(), "yyyymmdd") ' Tomo la fecha actual como ejemplo
        
        Set wsmtxca = New FEAFIPLib.wsmtxca
        wsmtxca.CUIT = 20939802593#
        wsmtxca.URL = URLWSW
        
        If wsmtxca.login("certificado.crt", "clave.key", URLWSAA) Then
            If Not wsmtxca.SFRecuperaLastCMP(PtoVta, Tipocomp) Then
               MsgBox wsmtxca.ErrorDesc
             Else
               nro = wsmtxca.SFLastCMP 'Devolucion el ultimo comprobante
           End If
            nro = nro + 1
            ' codigoTipoComprobante, numeroPuntoVenta, numeroComprobante,fechaEmision, codigoTipoDocumento, numeroDocumento, importeGravado, importeNoGravado, importeExento, importeSubtotal, importeOtrosTributos, importeTotal, codigoMoneda, cotizacionMoneda, observaciones, codigoConcepto, fechaServicioDesde, fechaServicioHasta, fechaVencimientoPago
            wsmtxca.AgregaFactura Tipocomp, PtoVta, nro, fechacmp, 80, 30702637895#, 100, 0, 0, 100, 0, 121, "PES", 1, "", 1, "", "", ""
            wsmtxca.AgregaIVA 5, 21  'Ver excel referencias codigos AFIP documentacion.rar
            ' unidadesMtx, codigoMtx, codigo, descripcion, cantidad, codigoUnidadMedida, precioUnitario, importeBonificacion, codigoCondicionIVA, importeIVA, importeItem
            wsmtxca.AgregaItem 1, "articulo1", "articulo1", "descripcion arti 1", 1, 1, 100, 0, 5, 21, 121
            If wsmtxca.Autorizar() Then
              If wsmtxca.SFResultado = "A" Then
                MsgBox "Felicitaciones! Si ve este cartel es porque obtuvo CAE y Vencimiento. CAE:" + wsmtxca.SFCAE + " Vencimiento: " + wsmtxca.SFVencimiento
              Else
                ' observaciones
                MsgBox wsmtxca.AutorizarRespuestaObs
              End If
            Else
                MsgBox wsmtxca.ErrorDesc
            End If
        Else
            MsgBox wsmtxca.ErrorDesc
        End If

End Sub
