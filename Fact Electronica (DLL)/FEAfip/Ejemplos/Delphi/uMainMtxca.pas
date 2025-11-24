unit uMainMtxca;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs;

type
  TForm1 = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses
  FEAFIPLib_TLB; // Generarla Importando el TypeLib despues de Registrar la Dll

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
const
  URLWSAA = 'https://wsaahomo.afip.gov.ar/ws/services/LoginCms';
  URLWSW = 'https://fwshomo.afip.gov.ar/wsmtxca/services/MTXCAService';
var
  lwsmtxca: wsmtxca;
  nro: Double;
  CAE, Vencimiento, Resultado, Reproceso, fecha: Widestring;
  PtoVta, TipoComp: Integer;
begin
  PtoVta := 110;
  TipoComp := 1; //Factura A(Ver excel referencias codigos AFIP documentacion.rar)
  fecha := FormatDateTime('yyyymmdd', now); // Tomo la fecha actual como ejemplo

  lwsmtxca := cowsmtxca.Create;
  lwsmtxca.CUIT := 20939802593;
  lwsmtxca.URL := URLWSW;

  If lwsmtxca.login('certificado.crt', 'clave.key', URLWSAA) Then begin
      If Not lwsmtxca.RecuperaLastCMP(PtoVta, Tipocomp, nro) Then begin
         ShowMessage( lwsmtxca.ErrorDesc);
         exit;
      End;
      nro := nro + 1;
      // codigoTipoComprobante, numeroPuntoVenta, numeroComprobante,fechaEmision, codigoTipoDocumento, numeroDocumento, importeGravado, importeNoGravado, importeExento, importeSubtotal, importeOtrosTributos, importeTotal, codigoMoneda, cotizacionMoneda, observaciones, codigoConcepto, fechaServicioDesde, fechaServicioHasta, fechaVencimientoPago
      lwsmtxca.AgregaFactura(Tipocomp, PtoVta, nro, fecha, 80, 30702637895, 100, 0, 0, 100, 0, 121, 'PES', 1, '', 1, '', '', '');
      lwsmtxca.AgregaIVA(5, 21);  //Ver excel referencias codigos AFIP documentacion.rar
      // unidadesMtx, codigoMtx, codigo, descripcion, cantidad, codigoUnidadMedida, precioUnitario, importeBonificacion, codigoCondicionIVA, importeIVA, importeItem
      lwsmtxca.AgregaItem(1, 'articulo1', 'articulo1', 'descripcion arti 1', 1, 1, 100, 0, 5, 21, 121);
      If lwsmtxca.Autorizar() Then
        If lwsmtxca.SFResultado = 'A' Then
          ShowMessage( 'Felicitaciones! Si ve este cartel es porque obtuvo CAE y Vencimiento. CAE:' + lwsmtxca.SFCAE + ' Vencimiento: ' + lwsmtxca.SFVencimiento)
        Else
          // observaciones
          Showmessage( lwsmtxca.AutorizarRespuestaObs  )
      Else
          showMessage( lwsmtxca.ErrorDesc);
  end
  else
      ShowMessage( lwsmtxca.ErrorDesc  );
end;

end.
