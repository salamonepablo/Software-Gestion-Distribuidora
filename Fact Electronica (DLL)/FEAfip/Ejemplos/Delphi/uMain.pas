unit uMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm2 = class(TForm)
    Button1: TButton;
    Button2: TButton;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
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

procedure TForm2.Button2Click(Sender: TObject);
const
  URLWSAA = 'https://wsaahomo.afip.gov.ar/ws/services/LoginCms';
  URLWSW = 'https://wswhomo.afip.gov.ar/wsfev1/service.asmx';
var
  lwsfev1: wsfev1;
  nro: Double;
  CAE, Vencimiento, Resultado, Reproceso, fecha: Widestring;
  PtoVta, TipoComp: Integer;
  msgText: String;
  I: Integer;
begin
  PtoVta := 20; // ATENCION! SI RECIBE UN ERROR DE FECHA O NUMERO DE COMPROBANTE EN ESTA DEMO CAMBIE ESTE VALOR POR OTRO DE 1 A 9999
  TipoComp := 1; // Factura A(Ver excel referencias codigos AFIP documentacion.rar)

  lwsfev1 := Cowsfev1.Create;
  lwsfev1.CUIT := 20939802593;
  lwsfev1.URL := URLWSW;

  If lwsfev1.login('certificado.crt', 'clave.key', URLWSAA) Then
  begin
      If Not lwsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
      begin
          ShowMessage (lwsfev1.ErrorDesc);
          Exit;
      end;
      fecha := FormatDateTime('yyyymmdd', now);
      msgText := '';
      lwsfev1.Reset;

      // Carga de un lote
      for I := 0 to 10 do
      begin
        Nro := nro + 1;
        lwsfev1.AgregaFactura(1, 80, 30702637895, nro, nro, fecha, 121, 0, 100, 0, '', '', '', 'PES', 1);
        lwsfev1.AgregaIVA(5, 100, 21); // Ver Excel de referencias de codigos AFIP
      end;

      If lwsfev1.Autorizar(PtoVta, TipoComp) Then
      begin
          // Proceso la respuesta
          for I := 0 to 10 do
          begin
            lwsfev1.AutorizarRespuesta(I, CAE, Vencimiento, Resultado, Reproceso);
            if Resultado = 'A' then
              msgText := msgText + Format('Indice: %d, CAE: %s, Vencimiento: %s', [I, CAE, Vencimiento]) + #10
            else
              msgText := msgText + Format('Indice: %d, Error: %s', [I, lwsfev1.AutorizarRespuestaObs(I)]) + #10;
          end;
          ShowMessage(msgText);
      end
      else
          ShowMessage(lwsfev1.ErrorDesc);
  end
  else
      ShowMessage(lwsfev1.ErrorDesc);

end;

procedure TForm2.Button1Click(Sender: TObject);
const
  URLWSAA = 'https://wsaahomo.afip.gov.ar/ws/services/LoginCms';
  URLWSW = 'https://wswhomo.afip.gov.ar/wsfev1/service.asmx';
var
  lwsfev1: wsfev1;
  nro: Double;
  CAE, Vencimiento, Resultado, Reproceso, fecha: Widestring;
  PtoVta, TipoComp: Integer;
begin
  PtoVta := 20; // ATENCION! SI RECIBE UN ERROR DE FECHA O NUMERO DE COMPROBANTE EN ESTA DEMO CAMBIE ESTE VALOR POR OTRO DE 1 A 9999
  TipoComp := 1; // Factura A(Ver excel referencias codigos AFIP documentacion.rar)

  lwsfev1 := Cowsfev1.Create;
  lwsfev1.CUIT := 20939802593;
  lwsfev1.URL := URLWSW;

  If lwsfev1.login('certificado.crt', 'clave.key', URLWSAA) Then
  begin
      If Not lwsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
      begin
          ShowMessage (lwsfev1.ErrorDesc);
          Exit;
      end;
      Nro := nro + 1;
      fecha := FormatDateTime('yyyymmdd', now);
      lwsfev1.Reset;
      lwsfev1.AgregaFactura(1, 80, 30702637895, nro, nro, fecha, 121, 0, 100, 0, '', '', '', 'PES', 1);
      lwsfev1.AgregaIVA(5, 100, 21); // Ver Excel de referencias de codigos AFIP
      If lwsfev1.Autorizar(PtoVta, TipoComp) Then
      begin
          lwsfev1.AutorizarRespuesta(0, CAE, Vencimiento, Resultado, Reproceso);
          ShowMessage('Felicitaciones! Si ve este mensaje es porque pudo obtener el CAE: ' + CAE + ' ' + Vencimiento);
          If Resultado <> 'A' Then
              ShowMessage(lwsfev1.AutorizarRespuestaObs(0));
      end
      else
          ShowMessage(lwsfev1.ErrorDesc);
  end
  else
      ShowMessage(lwsfev1.ErrorDesc);
end;

end.
