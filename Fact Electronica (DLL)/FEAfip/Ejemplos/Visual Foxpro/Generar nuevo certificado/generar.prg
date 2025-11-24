  certificado = CreateObject("FEAFIPLib.Certificado")
  if certificado.GenerarNuevoCertificado("Bit Ingeniería", "FEAFIP", 20939802593, "C:\FEAFIP\nuevopedido.csr", "c:\FEAFIP\nuevaclave.key") then
    MessageBox("Archivos generados exitosamente")
  else
    MessageBox(certificado.ErrorDesc)
  endif
