unit uMainConsultaARBA;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls;

type
  TForm1 = class(TForm)
    Button1: TButton;
    Edit1: TEdit;
    lbPercepcion: TLabel;
    lbNombrelab: TLabel;
    Label1: TLabel;
    lbRetencion: TLabel;
    procedure Button1Click(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

uses
  FEAFIPLib_TLB;

{$R *.dfm}

procedure TForm1.Button1Click(Sender: TObject);
var
  lwsPadron: IwsPadronARBA;
begin
    lwsPadron := CowsPadronARBA.Create;
    lwsPadron.ModoProduccion := true;
    lwsPadron.User := 'CUIT';
    lwsPadron.Password := 'Contraseña';
    If lwsPadron.ConsultaAlicuota('20160701', '20160731', StrToFloat(Edit1.Text)) Then
    begin
        lbPercepcion.Caption := FloatToStr(lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaPercepcion);
        lbRetencion.Caption := FloatToStr(lwsPadron.ConsultaAlicuotaRespuesta.AlicuotaRetencion);
    end
    Else
        ShowMessage (lwsPadron.ErrorDesc);
end;

end.
