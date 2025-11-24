  * Ver documentacion en http://www.bitingenieria.com.ar/doc/feafip/FEAFIPLib_TLB/IBarcode.html
  
  barcode = CreateObject("FEAFIPLib.Barcode")
  barcode.TamanioFuente = 8
  barcode.GenerarCodigo(20939802593, 1, 3, "12345678901234", "20171101", "C:\datos\codigobarras.bmp")
  MessageBox(barcode.Texto)
