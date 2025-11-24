program Facturas_A_B;

uses
  Forms,
  uMain in 'uMain.pas' {Form2},
  FEAFIPLib_TLB in 'FEAFIPLib_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm2, Form2);
  Application.Run;
end.
