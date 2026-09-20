{
--------------------------------------------------------------------------------
Copyright (c) 2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}
program OODemo;

uses
  Vcl.Forms,
  UMain in 'UMain.pas' {MainForm},
  OOObject in '..\src\OOObject.pas',
  OOTools in '..\src\OOTools.pas',
  OOWriter in '..\src\OOWriter.pas',
  OOTable in '..\src\OOTable.pas';

{$R *.res}

begin
  // Application.Initialize initialisiert auch COM (System.Win.ComObj); das ist Sache des Aufrufers (A7)
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
