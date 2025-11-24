  certificado = CreateObject("FEAFIPLib.Certificado")
  if certificado.CargarInformacionCertificado("certificado.crt", "clave.key") then
    certificado.MostrarInformacionCertificado()
  else
    MessageBox(certificadoMgr.ErrorDesc)
  endif
