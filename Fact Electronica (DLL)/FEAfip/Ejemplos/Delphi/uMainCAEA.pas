unit uMainCAEA;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm3 = class(TForm)
    btSolicitar: TButton;
    btConsultar: TButton;
    btInformar: TButton;
    GroupBox1: TGroupBox;
    lbCAE: TLabel;
    edPeriodo: TEdit;
    edOrden: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    procedure btSolicitarClick(Sender: TObject);
    procedure btConsultarClick(Sender: TObject);
    procedure btInformarClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form3: TForm3;

implementation

uses FEAFIPLib_TLB;

const
  //URLs de autenticacion y negocio. Cambiarlas por las de producción al implementarlas en el cliente(abajo)
  URLWSAA = 'https://wsaahomo.afip.gov.ar/ws/services/LoginCms';
    // Producción: https://wsaa.afip.gov.ar/ws/services/LoginCms
  URLWSW = 'https://wswhomo.afip.gov.ar/wsfev1/service.asmx';
    // Producción: https://servicios1.afip.gov.ar/wsfev1/service.asmx

{$R *.dfm}

procedure TForm3.btConsultarClick(Sender: TObject);
var
  lwsfev1: wsfev1;
  CAE: widestring;
  FechavigDesde: widestring;
  FechaVigHasta: widestring;
  FechaTope: widestring;
  FechaProceso: widestring;
begin
  lwsfev1  := Cowsfev1.Create;
  lwsfev1.CUIT := 20939802593;  // Cuit del vendedor
  lwsfev1.URL := URLWSW;
  If lwsfev1.login('certificado.crt', 'clave.key', URLWSAA) Then begin
          lwsfev1.Reset;
          If Not lwsfev1.CAEAConsultar(StrToInt(edPeriodo.Text), StrToInt(edOrden.Text), CAE, FechavigDesde, FechaVigHasta, FechaTope, FechaProceso) Then
              ShowMessage( lwsfev1.ErrorDesc )
          Else begin
              btInformar.Enabled := True;
              lbCAE.Caption := CAE;
          End;
  end
  Else
      ShowMessage( lwsfev1.ErrorDesc );
end;

procedure TForm3.btInformarClick(Sender: TObject);
var
  nro: double;
  CAE: widestring;
  Vencimiento: widestring;
  Resultado: widestring;
  Reproceso: widestring;
  PtoVta: integer;
  FechaComp: widestring;
  TipoComp: integer;
  lwsfev1: wsfev1;
begin
  TipoComp := 1;  // Factura A
  PtoVta := 10;
  FechaComp := FormatDateTime('yyyymmdd', now);

  lwsfev1 := Cowsfev1.Create;

  lwsfev1.CUIT := 20939802593; // Cuit del vendedor
  lwsfev1.URL := URLWSW;
  If lwsfev1.login('certificado.crt', 'clave.key', URLWSAA) Then begin
      If Not lwsfev1.RecuperaLastCMP(PtoVta, TipoComp, nro) Then
          ShowMessage (lwsfev1.ErrorDesc)
      Else begin
          nro := nro + 1;
          lwsfev1.AgregaFactura(1, 80, 30707219072, nro, nro, FechaComp, 121, 0, 100, 0, '', '', '', 'PES', 1);
          lwsfev1.AgregaIVA(5, 100, 21); // Ver Excel de referencias de codigos AFIP
          If lwsfev1.CAEAInformar(PtoVta, TipoComp, lbCAE.Caption) Then begin
              lwsfev1.AutorizarRespuesta( 0, CAE, Vencimiento, Resultado, Reproceso);
              If Resultado = 'A' Then
                  ShowMessage( 'Felicitaciones! Si ve este mensaje instalo correctamente FEAFIP. CAE y Vencimiento: ' + CAE + ' ' + Vencimiento)
              Else
                  ShowMessage( lwsfev1.AutorizarRespuestaObs(0) );
          end
          Else
              ShowMessage( lwsfev1.ErrorDesc );
      End
  end
  Else
      ShowMessage( lwsfev1.ErrorDesc );
end;

procedure TForm3.btSolicitarClick(Sender: TObject);
var
  nro: double;
  CAE: widestring;
  FechavigDesde: widestring;
  FechaVigHasta: widestring;
  FechaTope: widestring;
  FechaProceso: widestring;
  lwsfev1: wsfev1;
begin
  // Ver documentación www.bitingenieria.com.ar/webhelp
  lwsfev1 := Cowsfev1.Create;
  lwsfev1.CUIT := 20939802593;  // Cuit del vendedor
  lwsfev1.URL := URLWSW;
  If lwsfev1.login('certificado.crt', 'clave.key', URLWSAA) Then begin

          lwsfev1.Reset;
          If Not lwsfev1.CAEASolicitar(strtoint(edPeriodo.Text), strtoint(edOrden.Text), CAE, FechavigDesde, FechaVigHasta, FechaTope, FechaProceso) Then
              ShowMessage( lwsfev1.ErrorDesc )
          Else begin
              btInformar.Enabled := True;
              lbCAE.Caption := CAE;
          End
  end
  Else
      ShowMessage( lwsfev1.ErrorDesc );
end;

end.
