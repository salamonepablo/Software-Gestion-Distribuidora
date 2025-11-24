VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo FEAFIP"
   ClientHeight    =   5745
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7980
   LinkTopic       =   "Form1"
   ScaleHeight     =   5745
   ScaleWidth      =   7980
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "Consulta de monto minimo"
      Height          =   615
      Left            =   480
      TabIndex        =   6
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Consultar Factura"
      Height          =   615
      Left            =   4680
      TabIndex        =   5
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Rechazar Factura"
      Height          =   615
      Left            =   4680
      TabIndex        =   4
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Aceptar Factura"
      Height          =   615
      Left            =   4680
      TabIndex        =   3
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Informar Agente Dto. Col."
      Height          =   615
      Left            =   480
      TabIndex        =   2
      Top             =   2280
      Width           =   2775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Consultar Factura"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Generar Factura"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   840
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
Const URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"

Private Sub Command1_Click()
    PtoVta = 2000 ' Un Punto de Venta dado de alta en la AFIP
    TipoComp = 206 ' Factura A(Ver excel referencias codigos AFIP documentacion.rar)
    Dim nro As Double
    fecha = Format(Now, "yyyymmdd")
    fechaVenc = Format(Now + 30, "yyyymmdd")

    Dim lwsfev1 As wsfev1
    Set lwsfev1 = New FEAFIPLib.wsfev1
    lwsfev1.CUIT = 20939802593#
    lwsfev1.URL = URLWSW

    lwsfev1.Depurar = True
    If lwsfev1.login("certificado.crt", "clave.key", URLWSAA) Then
        If Not lwsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
            MsgBox (lwsfev1.ErrorDesc)
            Exit Sub
        End If
        nro = nro + 1

        lwsfev1.AgregaFactura 1, 80, 30504507323#, nro, nro, fecha, 7260000, 0, 6000000, 0, "", "", fechaVenc, "PES", 1
        lwsfev1.AgregaIVA 5, 6000000, 1260000 ' IVA 21
        lwsfev1.AgregaOpcional "2101", "0150507801000124703453"
        If lwsfev1.Autorizar(PtoVta, TipoComp) Then
          If lwsfev1.SFResultado(0) = "A" Then
            MsgBox ("Felicitaciones! CAE y Vencimiento: " + lwsfev1.SFCAE(0) + "/" + lwsfev1.SFVencimiento(0))
          Else
            MsgBox (lwsfev1.AutorizarRespuestaObs(0))
          End If
        Else
          MsgBox (lwsfev1.ErrorDesc)
        End If
    Else
        MsgBox (lwsfev1.ErrorDesc)
    End If
End Sub

Private Sub Command2_Click()
    rolCUITRepresentada = "EMISOR"
    CUITContraparte = 0
    codTipoCmp = 0
    estadoCmp = ""
    fecha_tipo = ""
    fechaDesde = ""
    fecha_hasta = ""
    codCtaCte = 0
    estadoCtaCte = ""
    
    Dim fceObj As New wsfecred
    fceObj.CUIT = 20939802593#
    fceObj.ModoProduccion = False
    fceObj.Depurar = True
    If fceObj.login("certificado.crt", "clave.key") Then
      If fceObj.consultarComprobantes(rolCUITRepresentada, _
        CUITContraparte, codTipoCmp, estadoCmp, fecha_tipo, fechaDesde, fecha_hasta, _
        codCtaCte, estadoCtaCte) Then
        Dim consultarComprobantesReturn As ConsultarCmpReturnTy
        Set consultarComprobantesReturn = fceObj.consultarCmpReturn
        MsgBox ("Consulta realizada con éxito. Cantidad de comprobantes: " + Str(consultarComprobantesReturn.arrayComprobantesCount))
      Else
        MsgBox (fceObj.ErrorDesc)
      End If
    Else
      MsgBox (fceObj.ErrorDesc)
    End If
End Sub

Private Sub Command3_Click()
    codCtaCte = 1
    CuitEmisor = 27929007862#
    codTipoCmp = 201
    PtoVta = 10
    nroCmp = 1

    Dim fceObj As New wsfecred
    fceObj.CUIT = 20939802593#
    fceObj.ModoProduccion = False
    fceObj.Depurar = True
    If fceObj.login("certificado.crt", "clave.key") Then
        Dim informarFECred As InformarFacturaAgtDptoCltvRequestTy
        Set informarFECred = fceObj.nuevoInformarFacturaAgtDptoCltvRequestTy
        'Usar una de las dos opciones de abajo para identificar. idCtaCte o idfactura
        'informarFECred.idCtaCte(codCtaCte)
        informarFECred.idFactura CuitEmisor, codTipoCmp, PtoVta, nroCmp
        cuentaDepositante = 250
        subcuentaComitente = 120310
        denominacion = "denominacion"
        informarFECred.ctaComitente cuentaDepositante, subcuentaComitente, denominacion
        If fceObj.informarFacturaAgtDptoCltv(informarFECred) Then
            MsgBox ("Operación realizada con éxito")
        Else
            MsgBox (fceObj.ErrorDesc)
        End If
    Else
        MsgBox (fceObj.ErrorDesc)
    End If
End Sub

Private Sub Command4_Click()
    CuitEmisor = 27929007862#
    codTipoCmp = 201
    PtoVta = 10
    nroCmp = 1

    Dim fceObj As New wsfecred
    fceObj.CUIT = 20939802593#
    fceObj.ModoProduccion = False
    fceObj.Depurar = True
    If fceObj.login("certificado.crt", "clave.key") Then
        Dim aceptarFECred As AceptarFECredRequestTy
        Set aceptarFECred = fceObj.nuevoAceptarFECredRequestTy

        'Usar una de las dos opciones de abajo para identificar. idCtaCte o idfactura
        'aceptarFECred.idCtaCte codCtaCte
        aceptarFECred.idFactura CuitEmisor, codTipoCmp, PtoVta, nroCmp
        aceptarFECred.tipoCancelacion = "TOT"
        aceptarFECred.importeCancelado = 10.1
        aceptarFECred.importeTotalRetPesos = 11.2
        aceptarFECred.importeEmbargoPesos = 12.4
        aceptarFECred.saldoAceptado = 13.6
        aceptarFECred.codMoneda = "PES"
        aceptarFECred.cotizacionMonedaUlt = 1.2

        'Confirmar Notas 1..N
        cacepta = True
        cCUITEmisor = 27929007862#
        ccodTipoCmp = 201
        cptoVta = 10
        cnroCmp = 1

        aceptarFECred.arrayConfirmarNotasDC cacepta, cCUITEmisor, ccodTipoCmp, cptoVta, cnroCmp

        'Formas de cancelacion 1..N
        codigoCancelacion = 1
        descripcionCancelacion = "descripcionCancelacion"
        aceptarFECred.arrayFormasCancelacion codigoCancelacion, descripcionCancelacion

        'Retenciones 1..N
        rcodTipo = 1
        rimporte = 10.3
        rporcentaje = 5.5
        rdescMotivo = "Motivo 2"
        aceptarFECred.arrayRetenciones rcodTipo, rimporte, rporcentaje, rdescMotivo
        rcodTipo = 1
        rimporte = 0.9
        rporcentaje = 0#
        rdescMotivo = "Motivo 2"
        aceptarFECred.arrayRetenciones rcodTipo, rimporte, rporcentaje, rdescMotivo

        'Ajustes 1..N
        'acodigo = 1
        'aimporte = 10.3
        'aceptarFECred.arrayAjustesOperacion acodigo, aimporte

        If fceObj.aceptarFECred(aceptarFECred) Then
            MsgBox ("Operación realizada con éxito")
        Else
            MsgBox (fceObj.ErrorDesc)
        End If
    Else
        MsgBox (fceObj.ErrorDesc)
    End If
End Sub

Private Sub Command5_Click()
    CuitEmisor = 27929007862#
    codTipoCmp = 201
    PtoVta = 10
    nroCmp = 1

    Dim fceObj As New wsfecred
    fceObj.CUIT = 20939802593#
    fceObj.ModoProduccion = False
    fceObj.Depurar = True
    If fceObj.login("certificado.crt", "clave.key") Then
       Dim rechazarFECred As RechazarFECredRequestTy
       Set rechazarFECred = fceObj.nuevoRechazarFECredRequestTy

      'Usar una de las dos opciones de abajo para identificar. idCtaCte o idfactura
      'aceptarFECred.idCtaCte codCtaCte
      rechazarFECred.idFactura CuitEmisor, codTipoCmp, PtoVta, nroCmp

      codMotivo = 1
      descMotivo = "Mercaderia dañada"
      justificacion = "Accidente vial"

      rechazarFECred.arrayMotivosRechazo codMotivo, descMotivo, justificacion
      If fceObj.rechazarFECred(rechazarFECred) Then
        MsgBox ("Operación realizada con éxito")
      Else
        MsgBox (fceObj.ErrorDesc)
      End If
    Else
      MsgBox (fceObj.ErrorDesc)
    End If
End Sub

Private Sub Command6_Click()
    rolCUITRepresentada = "RECEPTOR"
    CUITContraparte = 0
    codTipoCmp = 0
    estadoCmp = ""
    fecha_tipo = ""
    fechaDesde = ""
    fecha_hasta = ""
    codCtaCte = 0
    estadoCtaCte = ""

    Dim fceObj As New wsfecred
    fceObj.CUIT = 30504507323#
    fceObj.ModoProduccion = False
    fceObj.Depurar = True
    If fceObj.login("certificado.crt", "clave.key") Then
      If fceObj.consultarComprobantes(rolCUITRepresentada, _
        CUITContraparte, codTipoCmp, estadoCmp, fecha_tipo, fechaDesde, fecha_hasta, _
        codCtaCte, estadoCtaCte) Then
         Dim consultarComprobantesReturn As ConsultarCmpReturnTy
         Set consultarComprobantesReturn = fceObj.consultarCmpReturn
        MsgBox ("Consulta realizada con éxito. Cantidad de comprobantes: " + Str(consultarComprobantesReturn.arrayComprobantesCount))
      Else
        MsgBox (fceObj.ErrorDesc)
      End If
    Else
      MsgBox (fceObj.ErrorDesc)
    End If
End Sub

Private Sub Command7_Click()
  Dim fceObj As FEAFIPLib.wsfecred
  Set fceObj = New FEAFIPLib.wsfecred
  fceObj.CUIT = 20939802593#
  fceObj.ModoProduccion = False
  fceObj.Depurar = True
  If fceObj.login("certificado.crt", "clave.key") Then
    cuitConsultada = 30504507323#
    fechaEmision = Format(Now, "yyyy-mm-dd")
    If fceObj.consultarMontoObligadoRecepcion(cuitConsultada, fechaEmision) Then
      Dim respuesta As FEAFIPLib.ConsultarMontoObligadoRecepcionReturnTy
      Set respuesta = fceObj.consultarMontoObligadoRecepcionReturn
      If respuesta.obligado Then
        MsgBox ("Debe emitir una factura de crédito si el total de la factura es igual o mayor a $" + Str(respuesta.montoDesde))
      Else
        MsgBox ("No necesita emitir una factura de crédito")
      End If
    Else
      MsgBox (fceObj.ErrorDesc)
    End If
  End
  Else
    MsgBox (fceObj.ErrorDesc)
  End If
End Sub
