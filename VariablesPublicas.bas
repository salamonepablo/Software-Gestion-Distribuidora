Attribute VB_Name = "Vars"
Public BaseSPC As Database
Public tClientes
Public tPaises
Public tProvincias
Public tLocalidades
Public tCondicionIVA
Public tUltimosNumeros
Public tEmpleados
Public vFlagBuscar
Public vLeg
Public tDepositos
Public tDomiciliosClientes
Public tUnidadesMedida
Public tRubros
Public tProductos
Public tMovIntStockC
Public tMovIntStockD
Public tStock
Public tVendedores
Public tCtaCte
Public tFacturaC
Public tFacturaD
Public tNotaCreditoC
Public tNotaCreditoD
Public tNotaDebitoC
Public tNotaDebitoD
Public tNotaDebitoIC
Public tNotaDebitoID
Public tConsignasC
Public tConsignasD
Public tConsC
Public tConsD

'Agrego las tablas de recibos 2025-06-29
Public tRecibosC
Public tRecibosD

Public qMovIntStock
Public codVendedor As String
Public codDeposito As String
Public codVendedorDest As String
Public codDepositoDest As String
Public linea As Integer
Public Llamado As String
Public LlamaPagoPresup As Boolean
Public LlamaPagoFactura As Boolean
Public CantidadDirecciones As Integer
Public FacC
Public FacD
Public vNroFacImp
Public vNroRemImp
Public vTipoFacImp
Public vNroNCImp
Public vTipoNCImp
Public vNroRemImpNC
Public vNroNDImp
Public vTipoNDImp
Public vNroRemImpND
Public vDomicilio As String
Public vLocalidad As String
Public vCondIVA As String
Public vCUIT As String
Public IdSucursal As Long
Public NroRecibo As Long

'Variables para el comprobante asociado a la NC
 Public TipoCbteAsoc As Long
 Public NroCbteAsoc As Double
 Public FechaCbteAsoc As String

Public Declare Function WriteProfileString Lib "kernel32" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
 
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Const HWND_BROADCAST = &HFFFF&
Public Const WM_WININICHANGE = &H1A

Public Function BuscoSucursal(Sucursal As Long) As String
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set BaseSPC = DBEngine.OpenDatabase(ruta)

    Set tS = BaseSPC.OpenRecordset("Sucursales", dbOpenTable)
    tS.Index = "PrimaryKey"
    tS.Seek "=", Sucursal
    
    If Not tS.NoMatch Then
        BuscoSucursal = tS!NombreSucursal
    End If
    
    tS.Close
    BaseSPC.Close

End Function

Public Function certPadron(CUIT) As Double

'DISTRIBUIDORA
    
    Dim tPadron
    
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set BaseSPC = DBEngine.OpenDatabase(ruta)
    
    Set tPadron = BaseSPC.OpenRecordset("Padron", dbOpenTable)
   
    tPadron.Index = "CUIT"
    
    tPadron.Seek "=", CUIT
    
    If Not tPadron.NoMatch Then
        certPadron = tPadron!AlicuotaPercepcion
    Else
        certPadron = 0
    End If
    
    tPadron.Close

End Function

Public Function DescCondIVA(condicionIva As String) As String

    ruta = App.Path & "\DB_SPC_SI.mdb"
    
    Set BaseSPC = DBEngine.OpenDatabase(ruta)
    Set tCndIVA = BaseSPC.OpenRecordset("CondicionIva", dbOpenTable)

    tCndIVA.Index = "PrimaryKey"
    
    tCndIVA.Seek "=", condicionIva
    
    If Not tCndIVA.NoMatch Then tCndIVA.MoveFirst
    
    DescCondIVA = tCndIVA!Descripcion
    
    tCndIVA.Close
    BaseSPC.Close

End Function

Public Function RevertirFactura(TipoComp As Long, NroCbte As Double)
'DISTRIBUIDORA
    Dim FacturaC
    Dim FacturaD
    Dim tipoDoc
    
    Select Case TipoComp
        Case 1
            tipoDoc = "A"
        Case 6
            tipoDoc = "B"
    End Select
        
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set BaseSPC = DBEngine.OpenDatabase(ruta)
    
    vSQL = "SELECT * FROM FacturaD WHERE NroFactura=" & NroCbte & " AND TipoFactura='" & tipoDoc & "' ORDER BY ItemFactura"
    
    Set FacturaC = BaseSPC.OpenRecordset("FacturaC", dbOpenTable)
    Set FacturaD = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    FacturaC.Index = "PrimaryKey"
    
    FacturaC.Seek "=", tipoDoc, NroCbte
    
    If Not FacturaC.NoMatch Then FacturaC.Delete
    
    FacturaD.MoveFirst
    
    While Not FacturaD.EOF
        FacturaD.Delete
        FacturaD.MoveNext
    Wend

    FacturaD.Close
    FacturaC.Close
    BaseSPC.Close

End Function
Public Function ePadron(FDesde As String, FHasta As String, CUIT) As Double
    
    Dim lwsPadron As wsPadronARBA
    Dim clienteConCertificado As Double
    
    Set lwsPadron = New wsPadronARBA
    lwsPadron.User = "30708432543"
    lwsPadron.Password = "654321"
    lwsPadron.ModoProduccion = True ' Debe dar de alta el cuit en el entorno de test de ARBA http://www.test.arba.gov.ar/
    'If lwsPadron.ConsultaAlicuota("20220301", "20220331", CDbl(Replace(Text1.text, "-", ""))) Then
    
    If lwsPadron.ConsultaAlicuota(FDesde, FHasta, CDbl(Replace(CUIT, "-", ""))) Then
       ' lbPercepcion.Caption = CStr(lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaPercepcion)
       ' lbRetencion.Caption = CStr(lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaRetencion)
       ' lbGrupoPercepcion.Caption = CStr(lwsPadron.ConsultaAlicuotaRespuesta.GrupoPercepcion)
       ' lbGrupoRetencion.Caption = CStr(lwsPadron.ConsultaAlicuotaRespuesta.GrupoRetencion)
       ePadron = CDbl(lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaPercepcion)
       
       'If CUIT = 30546250241# Then
       ' ePadron = 0.1
       'End If
       clienteConCertificado = certPadron(CUIT)
       If clienteConCertificado <> 0 Then ePadron = clienteConCertificado
       
     Else
        
        Select Case lwsPadron.ErrorCode
            Case 2, 11
                ePadron = 0
            Case Else
                MsgBox (lwsPadron.ErrorDesc)
        End Select
    
    End If

End Function


Public Function RevertirNC(TipoComp As Long, NroCbte As Double) As Boolean
'DISTRIBUIDORA
    
    Dim NotaCreditoC
    Dim NotaCreditoD
    Dim tipoDoc
    
    Select Case TipoComp
        Case 3
            tipoDoc = "A"
        Case 8
            tipoDoc = "B"
    End Select
        
    ruta = App.Path & "\DB_SPC_SI.mdb"
    Set BaseSPC = DBEngine.OpenDatabase(ruta)
    
    vSQL = "SELECT * FROM NotaCreditoD WHERE NroNotaCredito=" & NroCbte & " AND TipoNotaCredito='" & tipoDoc & "' ORDER BY ItemNotaCredito"
    
    Set NotaCreditoC = BaseSPC.OpenRecordset("NotaCreditoC", dbOpenTable)
    Set NotaCreditoD = BaseSPC.OpenRecordset(vSQL, dbOpenDynaset)
    
    NotaCreditoC.Index = "PrimaryKey"
    
    NotaCreditoC.Seek "=", tipoDoc, NroCbte
    
    If Not NotaCreditoC.NoMatch Then NotaCreditoC.Delete
    
    NotaCreditoD.MoveFirst
    
    While Not NotaCreditoD.EOF
        NotaCreditoD.Delete
        NotaCreditoD.MoveNext
    Wend

    NotaCreditoD.Close
    NotaCreditoC.Close
    BaseSPC.Close

End Function

Public Function SetDefaultPrinter(objPrn As Printer) As Boolean
    
    Dim x As Long, sztemp As String
    sztemp = objPrn.DeviceName & "," & objPrn.DriverName & "," & objPrn.Port
    x = WriteProfileString("windows", "device", sztemp)
    x = SendMessage(HWND_BROADCAST, WM_WININICHANGE, 0&, "windows")

End Function
Public Sub Sendkeys(text$, Optional wait As Boolean = False)
    Dim WshShell As Object
    Set WshShell = CreateObject("wscript.shell")
    WshShell.Sendkeys text, wait
    Set WshShell = Nothing
End Sub
Public Function ActualizarCAE(Tabla As String, Tipo As String, NumeroDoc As Double, CAE As Double, Vencimiento As String)

    Dim TablaC

    Set TablaC = BaseSPC.OpenRecordset(Tabla, dbOpenTable)
    TablaC.Index = "PrimaryKey"
    TablaC.Seek "=", Tipo, NumeroDoc
    FechaVCAE = Right(Vencimiento, 2) & "/" & Mid(Vencimiento, 5, 2) & "/" & Left(Vencimiento, 4)
    
    If Not TablaC.NoMatch Then
            TablaC.Edit
                TablaC!CAE = CAE
                TablaC!FechaVC = Format(FechaVCAE, "DD/MM/YYYY")
            TablaC.Update
    End If
    
    TablaC.Close

End Function

Public Function CUITCliente(CodCliente As Long) As Double

    Set tCli = BaseSPC.OpenRecordset("Clientes", dbOpenTable)
    tCli.Index = "PrimaryKey"
    tCli.Seek "=", CodCliente
    
    If Not tCli.NoMatch Then
        CUITCliente = CDbl(tCli!CUIT)
    End If
    
    tCli.Close

End Function

Public Sub LimpiarTextBox(frm As Form)
    ' recorre todos los controles que hay en el formulario
    For Each Control In frm.Controls
        ' verifica que el control es de tipo TextBox
        If TypeOf Control Is TextBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.text = ""
        End If
    Next
    
End Sub
Public Sub DisabledTextBox(frm As Form)
    ' recorre todos los controles que hay en el formulario
    For Each Control In frm.Controls
        ' verifica que el control es de tipo TextBox
        If TypeOf Control Is TextBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Enabled = False
        End If
        
        If TypeOf Control Is ComboBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Enabled = False
        End If
    Next
End Sub
Public Sub EnabledTextBox(frm As Form)
    ' recorre todos los controles que hay en el formulario
    For Each Control In frm.Controls
        ' verifica que el control es de tipo TextBox
        If TypeOf Control Is TextBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Enabled = True
        End If
        
        If TypeOf Control Is ComboBox Then
            '... Si es un Textbox, entonces lo limpia
            Control.Enabled = True
        End If
    Next
End Sub
Public Function Verificar_Tecla(Tecla_Presionada)
    
    
 Dim Teclas As String
    
    'Acepta todos los números, la tecla Backspace, _
     la tecla Enter, la coma y el punto
    
    Teclas = "1234567890.," & Chr(vbKeyBack)
    
    If InStr(1, Teclas, Chr(Tecla_Presionada)) Then
        
        Verificar_Tecla = Tecla_Presionada
    Else
        ' Si no es ninguna de las indicadas retorna 0
        Verificar_Tecla = 0
    End If

End Function
Public Function Autocompletar_Combo(Combo As ComboBox)
  
 Dim I As Integer, posSelect As Integer
  
    Select Case (KeyRetroceso Or Len(Combo.text) = 0)
        Case True
            KeyRetroceso = False
            Exit Function
    End Select
  
    With Combo
  
    'Recorremos todos los elementos del combo
    For I = 0 To .ListCount - 1
        'Si hay coincidencia
        If InStr(1, .List(I), .text, vbTextCompare) = 1 Then
            posSelect = .SelStart
            'Mostramos el texto en el combo
            .text = .List(I)
            'Indicamos el comienzo de la selección
            .SelStart = posSelect
            'Acá seleccionamos el texto
            .SelLength = Len(.text) - posSelect
  
            Exit For
        End If
    Next I
  
    End With

End Function

Public Function BuscarDescProd(IdCodProd) As String

    Set tP = BaseSPC.OpenRecordset("Productos", dbOpenTable)
    tP.Index = "PrimaryKey"
    tP.Seek "=", IdCodProd
    
    If Not tP.NoMatch Then
        BuscarDescProd = tP!Descripcion
    End If
    
    tP.Close

End Function

Public Sub CrearQR(Fecha As String, CUIT As Double, PtoVta As Long, TipoComp As Long, nroCmp As Double, Importe As Double, Moneda As String, ctz As Double, tipoDocRec As Long, nroDocRec As Double, TipoCodAut As String, codAut As Double)

  Dim Qr As FEAFIPLib.Qr
  Dim tipoDoc, carpetaQR As String
  
  Set Qr = New FEAFIPLib.Qr
  
  Select Case TipoComp
    Case 1
        tipoDoc = "FA"
    Case 2
        tipoDoc = "NDA"
    Case 3
        tipoDoc = "NCA"
    Case 6
        tipoDoc = "FB"
    Case 1
        tipoDoc = "NDB"
    Case 8
        tipoDoc = "NCB"
  End Select
    
  carpetaQR = "\QRs\" + "qr" + "_" & tipoDoc & "_" & PtoVta & "_" & nroCmp & ".jpg"
  Qr.ArchivoQR = Qr.RutaLibreria & carpetaQR ' Admite formatos BMP, PNG y JPG con solo cambiar la extension

  'Long
  Ver = 1
  
  'String
  'Fecha = ""
  
  'Double
 ' CUIT = 30708432543#
  
  'Long
  'PtoVta = PV
  
  'Long
 ' Select Case TextTipoFactura.Text
 '   Case "A"
 '       TipoComp = 1
 '   Case "B"
 '       TipoComp = 6
 ' End Select
  
  'Double
 ' nroCmp = TextNumeroFactura.Text
  
  'Double
 ' Importe = 100.2
  
  'String
 ' Moneda = "PES"
  
  'Double
 ' ctz = 1#
  
  'Long
 ' tipoDocRec = 80
  
  'Double
 ' nroDocRec = 27929007862#
  
  'String
  TipoCodAut = "E"  ' A = CAEA E = CAE
  
  'Double
  'codAut = 12345678901234#
  
  If Qr.Generar(Ver, Fecha, CUIT, PtoVta, TipoComp, nroCmp, Importe, Moneda, ctz, tipoDocRec, nroDocRec, TipoCodAut, codAut) Then
    '  MsgBox ("QR generado con éxito en " + Qr.ArchivoPNG)
  Else
    MsgBox (Qr.ErrorDesc)
  End If

End Sub

Public Function FacturaElectronicaSPC(PtoVta As Long, DocTipo As Long, DocNro As Double, TipoComp As Long, CbteDesde As Double, CbteHasta As Double, CbteFch As String, ImpTotal As Double, ImpNeto As Double, MonId As String, MonCotiz As Double, AlicIVA As Long, BaseImpIVA As Double, ImpIva As Double, IdTributo As Long, DescTributo As String, BaseImpTributo As Double, Alicuota As Double, ImpAlicuota As Double, ImporteExento As Double, Optional TipoCbteAsoc As Long, Optional NroCbteAsoc As Double, Optional FechaCbteAsoc As String) As Boolean

'DISTRUIBUIDORA

        ' Los nombres de los parametros de las funciones se obtienen en FEAFIP.pdf
        
        On Error GoTo CapturaErrores
        
        'URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
        Const URLWSAA = "https://wsaa.afip.gov.ar/ws/services/LoginCms"
        'Const URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
          ' Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
        Const URLWSW = "https://servicios1.afip.gov.ar/wsfev1/service.asmx"
        'Const URLWSW = "https://wswhomo.afip.gov.ar/wsfev1/service.asmx"
          ' Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx
        Dim wsfev1 As FEAFIPLib.wsfev1 ' Si esta linea falla es porqu eno agrego la referencia en a FEAFIPLib desde el menu de proyecto
        Dim Nro As Double
        CAE$ = ""
        Vencimiento$ = ""
        Resultado$ = ""
        Reproceso$ = ""
        Nro = 1
        'PtoVta = 4  ' ATENCION! SI RECIBE UN ERROR DE FECHA O NUMERO DE COMPROBANTE EN ESTA DEMO CAMBIE ESTE VALOR POR OTRO DE 1 A 9999
        'TipoComp = 1 ' Factura A(Ver excel referencias codigos AFIP)
        'FechaComp = Format(Now(), "yyyymmdd")
        
        Dim tipoDoc As String
         
        Select Case TipoComp
            Case 1
                tipoDoc = "FA"
            Case 2
                tipoDoc = "NDA"
            Case 3
                tipoDoc = "NCA"
            Case 6
                tipoDoc = "FB"
            Case 7
                tipoDoc = "NDB"
            Case 8
                tipoDoc = "NCB"
          End Select
          
        Set wsfev1 = New FEAFIPLib.wsfev1
           
        wsfev1.CUIT = 30708432543# ' Cuit del vendedor
        wsfev1.URL = URLWSW
        
        If wsfev1.login("quilplac.crt", "quilplac.key", URLWSAA) Then
            If Not wsfev1.RecuperaLastCMP(PtoVta, TipoComp, Nro) Then
                MsgBox (wsfev1.ErrorDesc)
                FacturaElectronicaSPC = False
            Else
                Nro = Nro + 1
                wsfev1.Reset
                wsfev1.AgregaFactura 1, DocTipo, DocNro, CbteDesde, CbteHasta, CbteFch, ImpTotal, 0, ImpNeto, ImporteExento, "", "", "", MonId, MonCotiz
                wsfev1.AgregaIVA AlicIVA, BaseImpIVA, ImpIva  ' Ver Excel de referencias de codigos AFIP
                
           'Acá Agregar el comprobante asociado si es NC o ND
            Select Case TipoComp
                Case 2
                    wsfev1.AgregaCompAsoc TipoCbteAsoc, PtoVta, NroCbteAsoc, 30708432543#, FechaCbteAsoc
                Case 3
                    wsfev1.AgregaCompAsoc TipoCbteAsoc, PtoVta, NroCbteAsoc, 30708432543#, FechaCbteAsoc
                Case 7
                    wsfev1.AgregaCompAsoc TipoCbteAsoc, PtoVta, NroCbteAsoc, 30708432543#, FechaCbteAsoc
                Case 8
                    wsfev1.AgregaCompAsoc TipoCbteAsoc, PtoVta, NroCbteAsoc, 30708432543#, FechaCbteAsoc
            End Select
           
           'Acá agregar la percepción de IIBB
                wsfev1.AgregaTributo IdTributo, DescTributo, BaseImpTributo, Alicuota, ImpAlicuota
                
                If wsfev1.Autorizar(PtoVta, TipoComp) Then
                    wsfev1.AutorizarRespuesta 0, CAE, Vencimiento, Resultado, Reproceso
                    If Resultado = "A" Then
                        'MsgBox "Felicitaciones! Si ve este mensaje instalo correctamente FEAFIP. CAE y Vencimiento: " + CAE + " " + Vencimiento
                        MsgBox ("Resultado: " & Resultado & Chr(10) & "Comprobante: " & tipoDoc & "-" & PtoVta & "-" & Nro & Chr(10) & "CAE: " & CAE & Chr(10) & "Vencimiento CAE :" & Vencimiento)
                        FacturaElectronicaSPC = True
                    Else
                        MsgBox wsfev1.AutorizarRespuestaObs(0)
                        FacturaElectronicaSPC = False
                        Exit Function
                    End If
                Else
                    MsgBox wsfev1.ErrorDesc
                    FacturaElectronicaSPC = False
                    Exit Function
                End If
            End If
            
            Select Case TipoComp
                Case 1
                    Call ActualizarCAE("FacturaC", "A", CbteDesde, (CAE), Vencimiento)
                Case 2
                    Call ActualizarCAE("NotaDebitoC", "A", CbteDesde, (CAE), Vencimiento)
                Case 3
                    Call ActualizarCAE("NotaCreditoC", "A", CbteDesde, (CAE), Vencimiento)
                Case 6
                    Call ActualizarCAE("FacturaC", "B", CbteDesde, (CAE), Vencimiento)
                Case 7
                    Call ActualizarCAE("NotaDebitoC", "B", CbteDesde, (CAE), Vencimiento)
                Case 8
                    Call ActualizarCAE("NotaCreditoC", "B", CbteDesde, (CAE), Vencimiento)
            End Select
        Else
            MsgBox wsfev1.ErrorDesc
            FacturaElectronicaSPC = False
            Exit Function
        End If
      
CapturaErrores:
    Select Case Err.Number
        Case -2147418113 ' &H8000FFFF - Error de conexión inesperada
            MsgBox "Error de conexión con ARCA. Verifique su conexión a internet o intente nuevamente más tarde.", vbExclamation
                FacturaElectronicaSPC = False
        Case Else
            Resume Next
            ' Otros errores no esperados
         '   MsgBox "Un error inesperado ocurrió. Por favor, contacte a soporte técnico.", vbCritical
    End Select

End Function
Public Function BuscaCbteAsociado(NroCbteAsociado As Long, TipoCbteAsociado As String)
    
    Set tP = BaseSPC.OpenRecordset("FacturaC", dbOpenTable)
    tP.Index = "PrimaryKey"
    tP.Seek "=", TipoCbteAsociado, NroCbteAsociado
    
    If Not tP.NoMatch Then
        NroCbteAsoc = tP!NroFactura
        
        Select Case tP!TipoFactura
            Case "A"
                TipoCbteAsoc = 1
            Case "B"
                TipoCbteAsoc = 6
        End Select
        
        FechaCbteAsoc = CStr(Format(tP!FechaFactura, "YYYYMMDD"))
    End If
    
    tP.Close

End Function

