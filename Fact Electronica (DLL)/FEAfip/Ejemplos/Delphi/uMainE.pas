unit uMainE;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs;

type
  TForm2 = class(TForm)
    procedure FormCreate(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form2: TForm2;

implementation

uses
  FEAFIPLib_TLB; // Generarla Importando el TypeLib despues de Registrar la Dll

{$R *.dfm}

procedure TForm2.FormCreate(Sender: TObject);
const
  URLWSAA = 'https://wsaahomo.afip.gov.ar/ws/services/LoginCms';
  URLWSW = 'https://wswhomo.afip.gov.ar/wsfexv1/service.asmx';
var
  lwsfexv1: wsfexv1;
  nro, IdTrans: Double;
  CAE, Vencimiento, Resultado, Reproceso, fecha: Widestring;
  PtoVta, TipoComp: Integer;
begin
  PtoVta := 10; // Un Punto de Venta dado de alta en la AFIP;
  TipoComp := 19; //

  lwsfexv1 := Cowsfexv1.Create;
  lwsfexv1.CUIT := 20939802593;
  lwsfexv1.URL := URLWSW;

  If lwsfexv1.login('certificado.crt', 'clave.key', URLWSAA) Then
  begin
      If Not lwsfexv1.RecuperaLastCMP(PtoVta, TipoComp, nro, fecha) Then
      begin
          ShowMessage (lwsfexv1.ErrorDesc);
          Exit;
      end;
      Nro := nro + 1;
      If Not lwsfexv1.UltimoIdTrans(IdTrans) Then
      begin
          ShowMessage (lwsfexv1.ErrorDesc);
          Exit;
      end;
      IdTrans := IdTrans + 1;
      fecha := FormatDateTime('yyyymmdd', now);
      lwsfexv1.AgregaFactura(IdTrans, fecha, 19, PtoVta, nro, 2, '', 208, 'chile sa', 50000000032, 'texto', '', 'DOL', 8.52, '', 100, '', 'contado', 'DES', 1);
      lwsfexv1.AgregaItem('11111', 'remera', 1, 1, 100, 100, 0);
      If lwsfexv1.Autorizar Then
      begin
          lwsfexv1.AutorizarRespuesta(CAE, Vencimiento, Resultado, Reproceso);
          ShowMessage('Felicitaciones! Si ve este mensaje es porque pudo obtener el CAE: ' + CAE + ' ' + Vencimiento);
          If Resultado <> 'A' Then
              ShowMessage(lwsfexv1.AutorizarRespuestaObs());
      end
      else
          ShowMessage(lwsfexv1.ErrorDesc);
      with TStringList.Create do
      try
        Text := lwsfexv1.XMLRequest;
        SaveToFile('xmlRequest.xml');
        Text := lwsfexv1.XMLResponse;
        SaveToFile('xmlResponse.xml');
      finally
        Free;
      end;
  end
  else
      ShowMessage(lwsfexv1.ErrorDesc);

end;

end.
