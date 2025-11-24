// Parseo de un xml al consultar un comprobante FEX
// XML de ejemplo debajo

//  <?xml version="1.0" encoding="utf-8"?>
//  <soap:Envelope xmlns:soap="http://www.w3.org/2003/05/soap-envelope"
//      xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
//      xmlns:xsd="http://www.w3.org/2001/XMLSchema">
//      <soap:Body>
//          <FEXGetCMPResponse xmlns="http://ar.gov.afip.dif.fexv1/">
//              <FEXGetCMPResult>
//                  <FEXResultGet>
//                      <Id>21999700000663</Id>
//                      <Fecha_cbte>20170421</Fecha_cbte>
//                      <Cbte_tipo>19</Cbte_tipo>
//                      <Punto_vta>100</Punto_vta>
//                      <Cbte_nro>169</Cbte_nro>
//                      <Tipo_expo>1</Tipo_expo>
//                      <Permiso_existente>N</Permiso_existente>
//                      <Dst_cmp>208</Dst_cmp>
//                      <Cliente>chile sa</Cliente>
//                      <Cuit_pais_cliente>50000000032</Cuit_pais_cliente>
//                      <Domicilio_cliente>Domicilio</Domicilio_cliente>
//                      <Id_impositivo/>
//                      <Moneda_Id>DOL</Moneda_Id>
//                      <Moneda_ctz>8</Moneda_ctz>
//                      <Obs_comerciales/>
//                      <Imp_total>100</Imp_total>
//                      <Obs/>
//                      <Forma_pago>contado</Forma_pago>
//                      <Incoterms>DES</Incoterms>
//                      <Incoterms_Ds>0</Incoterms_Ds>
//                      <Idioma_cbte>1</Idioma_cbte>
//                      <Items>
//                          <Item>
//                              <Pro_codigo>11111</Pro_codigo>
//                              <Pro_ds>remera </Pro_ds>
//                              <Pro_qty>1</Pro_qty>
//                              <Pro_umed>1</Pro_umed>
//                              <Pro_precio_uni>100</Pro_precio_uni>
//                              <Pro_bonificacion>0</Pro_bonificacion>
//                              <Pro_total_item>100</Pro_total_item>
//                          </Item>
//                      </Items>
//                      <Fch_venc_Cae>20170421</Fch_venc_Cae>
//                      <Cae>67163005474368</Cae>
//                      <Resultado>A</Resultado>
//                      <Motivos_Obs/>
//                  </FEXResultGet>
//                  <FEXErr>
//                      <ErrCode>0</ErrCode>
//                      <ErrMsg>OK</ErrMsg>
//                  </FEXErr>
//                  <FEXEvents>
//                      <EventCode>0</EventCode>
//                      <EventMsg>OK</EventMsg>
//                  </FEXEvents>
//              </FEXGetCMPResult>
//          </FEXGetCMPResponse>
//      </soap:Body>
//  </soap:Envelope>

  URLWSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
  URLWSW = "https://wswhomo.afip.gov.ar/wsfexv1/service.asmx"
 
  ws = CreateObject("FEAFIPLib.wsfexv1")  
  ws:CUIT = 20939802593
  ws:URL = URLWSW
  If ws:login("certificado.crt", "clave.key", URLWSAA)
    CAE = ""
    Vencimiento = ""
    If ws:CmpConsultar(19, 100, 169, CAE, Vencimiento)
      doc = CreateObject("MSXML2.DOMDocument")

      // Remuevo los namespaces de la respuesta para parsear mas facil.
      response = strtran(ws:XMLResponse, 'xmlns:soap="http://www.w3.org/2003/05/soap-envelope"', "")
      response = strtran(response, "soap:", "")
      response = strtran(response, 'xmlns="http://ar.gov.afip.dif.fexv1/"', "")


      doc:loadXML(response)
      docNode = doc:selectSingleNode("/Envelope/Body/FEXGetCMPResponse/FEXGetCMPResult/FEXResultGet")

      Id = docNode:selectSingleNode("Id"):Text
      Fecha_cbte = docNode:selectSingleNode("Fecha_cbte"):Text

        // Agregar los campos que se necesiten
        // ................

      //Selecciono los items
      Items = docNode:selectNodes("Items/Item")
      
      For I = 0 To Items:length - 1
        // Recorro los items
        Pro_codigo = Items:Item(I):selectSingleNode("Pro_codigo"):Text
        Pro_ds = Items:Item(I):selectSingleNode("Pro_ds"):Text
        // Agregar los campos que se necesiten
        // ................
        MessageBox(0, Pro_ds, "FEAFIP", 0)
      Next
    Else
      MessageBox(0, ws:ErrorDesc, "FEAFIP", 0)
    EndIf
  Else
    MessageBox(0, ws:ErrorDesc, "FEAFIP", 0)
  EndIf