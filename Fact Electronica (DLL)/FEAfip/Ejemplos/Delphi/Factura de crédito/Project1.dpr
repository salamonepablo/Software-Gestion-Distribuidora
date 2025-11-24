program Project1;

uses
  Forms,
  Unit1 in 'Unit1.pas' {Form1},
  FEAFIPLib_TLB in '..\..\..\..\..\..\Program Files (x86)\Borland\Delphi7\Imports\FEAFIPLib_TLB.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
