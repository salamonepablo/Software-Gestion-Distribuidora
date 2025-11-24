
* Elija la ruta de destino
lwsPadron = CreateObject("FEAFIPLib.wsPadron")
if lwsPadron.descargarConstancia(20939802593, "constancia.pdf") then
  MessageBox("Constancia descargada")
else
  MessageBox(lwsPadron.ErrorDesc)
endif
