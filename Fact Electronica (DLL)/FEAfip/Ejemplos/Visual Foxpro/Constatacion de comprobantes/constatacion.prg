URL_WSAA = "https://wsaahomo.afip.gov.ar/ws/services/LoginCms"
URL_WSCDC = "https://wswhomo.afip.gov.ar/WSCDC/service.asmx"

lwscdc = CreateObject("FEAFIPLib.wscdc")
lwscdc.Depurar = .T.
lwscdc.CUIT = 20939802593
lwscdc.URL = URL_WSCDC
if lwscdc.Login("certificado.crt", "clave.key", URL_WSAA) then
  if lwscdc.ComprobanteConstatar("CAE", 20939802593, 140, 1, 1588, "20170517", 1452.73, "67203477090542", "80", "27929007862") then
    MessageBox("Comprobante constatado con éxito.")
  else
    MessageBox(lwscdc.ErrorDesc)
  endif
else
  MessageBox(lwscdc.ErrorDesc)
endif
