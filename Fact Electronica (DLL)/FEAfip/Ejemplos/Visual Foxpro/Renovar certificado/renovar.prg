  certificado = CreateObject("FEAFIPLib.Certificado")
  if certificado.CargarInformacionCertificado("certificado.crt", "clave.key") then
    if certificado.RenovarCertificado("c:\FEAFIP\pedido.csr") then
      MessageBox("Certificado renovado exitosamente")
    else
      MessageBox(certificado.ErrorDesc)
    endif
  else
    MessageBox(certificadoMgr.ErrorDesc)
  endif
