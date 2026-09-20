{
--------------------------------------------------------------------------------
Copyright (c) 2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}
program OOAutomationTests;

{$APPTYPE CONSOLE}
{$STRONGLINKTYPES ON}

uses
  System.SysUtils,
  System.Win.ComObj,
  Winapi.ActiveX,
  DUnitX.TestFramework,
  DUnitX.Loggers.Console,
  DUnitX.MemoryLeakMonitor.FastMM4,
  OOTools in '..\src\OOTools.pas',
  OOTable in '..\src\OOTable.pas',
  OOObject in '..\src\OOObject.pas',
  OOWriter in '..\src\OOWriter.pas',
  Test.Support in 'Test.Support.pas',
  Test.OOTools in 'Test.OOTools.pas',
  Test.OOWriter in 'Test.OOWriter.pas';

var
  Runner: ITestRunner;
  Results: IRunResults;

begin
  // Was kein Test-Fenster des Leck-Monitors erfasst (Fixture-Auf-/Abbau, Programmende), meldet der
  // Speichermanager beim Beenden – im Konsolenprogramm nach stderr, ohne MessageBox (Plan 7.1, T7)
  ReportMemoryLeaksOnShutdown := True;
  // COM-Initialisierung liegt beim Aufrufer, nicht in der Bibliothek (A7)
  OleCheck(CoInitializeEx(nil, COINIT_APARTMENTTHREADED));
  try
    try
      TDUnitX.CheckCommandLine;
      Runner := TDUnitX.CreateRunner;
      Runner.UseRTTI := True;
      // Ein Test ohne Assert gilt als Fehler – sonst läuft ein leerer Test still grün durch
      Runner.FailsOnNoAsserts := True;
      Runner.AddLogger(TDUnitXConsoleLogger.Create(TDUnitX.Options.ConsoleMode = TDunitXConsoleMode.Quiet));
      WarmUpLazySingletons;
      Results := Runner.Execute;
      if not Results.AllPassed then
      begin
        System.ExitCode := EXIT_ERRORS;
      end;
      // Warten nur bei Aufruf mit --exit:pause; Build-Skripte laufen ohne Tastendruck durch
      if TDUnitX.Options.ExitBehavior = TDUnitXExitBehavior.Pause then
      begin
        System.Write('Fertig – <Enter> beendet.');
        System.Readln;
      end;
    except
      on E: Exception do
      begin
        System.Writeln(E.ClassName, ': ', E.Message);
        System.ExitCode := EXIT_ERRORS;
      end;
    end;
  finally
    // Erst die Testobjekte samt ihren COM-Verweisen freigeben, dann COM beenden
    Results := nil;
    Runner := nil;
    CoUninitialize;
  end;
end.
