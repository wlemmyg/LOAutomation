program OODemo;

uses
  Vcl.Forms,
  UMain in 'UMain.pas' {Form1},
  OOObject in '..\src\OOObject.pas',
  OOTools in '..\src\OOTools.pas',
  OOWriter in '..\src\OOWriter.pas',
  OOTable in '..\src\OOTable.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TForm1, Form1);
  Application.Run;
end.
