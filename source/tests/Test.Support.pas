{
--------------------------------------------------------------------------------
Copyright (c) 2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}
unit Test.Support;

// Gemeinsame Helfer der Tests: Fehlerprüfung (Plan 7.1, T6) und die Umgebung der Integrationstests (T8)

interface

uses
  System.SysUtils;

type
  // Testdaten, Temp-Ordner und LibreOffice-Desktop für die Integrationstests
  TTestEnvironment = class
  strict private
    class var FRunId: string;
    class var FTestDataDir: string;
    class var FTempDir: string;
    class var FSpecialDir: string;
    class var FServiceManager: OleVariant;
    class var FDesktop: OleVariant;
    class var FOfficeWasRunning: Boolean;
    class function FindTestDataDir: string; static;
    class function IsOfficeRunning: Boolean; static;
    class function Components(AOwnOnly: Boolean): TArray<OleVariant>; static;
  public
    class procedure Start; static;
    class procedure Finish; static;
    class function CopyTestFile(const ASource, ATarget: string; AInSpecialDir: Boolean = False): string; static;
    class function IsDocumentOpen(const AFileName: string): Boolean; static;
    class procedure CloseOwnDocuments; static;
    class function OdtContains(const AFileName, AText: string): Boolean; static;
    class function CreateCalcFile(const ATarget: string): string; static;
    class function CreateMarkedDocument(const ATarget: string): string; static;
    class property TempDir: string read FTempDir;
    class property SpecialDir: string read FSpecialDir;
  end;

// Erwartet EOOAutomation, deren Meldung alle Teile enthält – geprüft werden die Schlüsselangaben, nicht der
// Wortlaut. Eine andere Ausnahme läuft durch und erscheint als Fehler des Tests.
procedure AssertRaisesOO(const AProc: TProc; const AParts: array of string);

// Legt RTL- und DUnitX-Objekte an, die beim ersten Gebrauch entstehen und erst beim Programmende frei werden.
// Ohne das rechnet der Leck-Monitor sie dem ersten Test zu, der sie berührt (Plan 7.1, T7).
procedure WarmUpLazySingletons;

// Name des Windows-Standarddruckers; ohne Standarddrucker bricht der Test laut ab
function DefaultPrinterName: string;

// Suche in einer Sequenz von PropertyValues (etwa aus getPrinter oder PrintArgs)
function HasProperty(const AProperties: OleVariant; const AName: string): Boolean;
function FindProperty(const AProperties: OleVariant; const AName: string): OleVariant;

// Die ersten ACount Bytes einer Datei als Zeichen, ohne Kodierung (etwa '%PDF')
function FileHead(const AFileName: string; ACount: Integer): string;

implementation

uses
  System.Classes,
  System.IOUtils,
  System.StrUtils,
  System.Variants,
  System.Win.ComObj,
  System.Zip,
  Winapi.TlHelp32,
  Winapi.Windows,
  Winapi.WinSpool,
  DUnitX.TestFramework,
  DUnitX.Timeout,
  OOTools;

const
  // Nur zum Anlegen des Threads; er wird sofort wieder gestoppt
  CWarmUpTimeout = 60000;

{ ===== Fehlerprüfung ===== }

procedure AssertRaisesOO(const AProc: TProc; const AParts: array of string);
var
  part: string;
begin
  try
    AProc();
  except
    on E: EOOAutomation do
    begin
      // Zählt als Prüfung, auch wenn keine Meldungsteile verlangt sind (FailsOnNoAsserts)
      Assert.InheritsFrom(E.ClassType, EOOAutomation);
      for part in AParts do
      begin
        Assert.Contains(E.Message, part, 'Meldung nennt "' + part + '" nicht: ' + E.Message);
      end;
      Exit;
    end;
  end;
  Assert.Fail('EOOAutomation erwartet, es kam keine Ausnahme');
end;

{ ===== Leck-Monitor ===== }

procedure WarmUpLazySingletons;
var
  zip: TZipFile;
  timeout: ITimeout;
begin
  // TEncoding.UTF8: FileNameToUrl, OdtContains
  TEncoding.UTF8.GetByteCount('');
  // TZipFile.FCP437Encoding (System.Zip): OdtContains
  zip := TZipFile.Create;
  try
    zip.Encoding.GetByteCount('');
  finally
    zip.Free;
  end;
  // Erster Timeout-Thread von [MaxTime]
  timeout := InitialiseTimeout(CWarmUpTimeout);
  timeout := nil;
end;

{ ===== Drucker und Eigenschaften ===== }

function DefaultPrinterName: string;
var
  size: DWORD;
begin
  size := 0;
  GetDefaultPrinter(nil, @size);
  if size = 0 then
  begin
    raise Exception.Create('Kein Standarddrucker eingerichtet; die Drucktests brauchen einen.');
  end;
  // size zählt die abschließende Null mit
  SetLength(Result, size - 1);
  if not GetDefaultPrinter(PChar(Result), @size) then
  begin
    RaiseLastOSError;
  end;
end;

function HasProperty(const AProperties: OleVariant; const AName: string): Boolean;
var
  idx: Integer;
begin
  for idx := VarArrayLowBound(AProperties, 1) to VarArrayHighBound(AProperties, 1) do
  begin
    if AProperties[idx].Name = AName then
    begin
      Exit(True);
    end;
  end;
  Result := False;
end;

function FindProperty(const AProperties: OleVariant; const AName: string): OleVariant;
var
  idx: Integer;
begin
  Result := Unassigned;
  for idx := VarArrayLowBound(AProperties, 1) to VarArrayHighBound(AProperties, 1) do
  begin
    if AProperties[idx].Name = AName then
    begin
      Exit(AProperties[idx].Value);
    end;
  end;
  Assert.Fail('Eigenschaft "' + AName + '" fehlt');
end;

function FileHead(const AFileName: string; ACount: Integer): string;
var
  bytes: TBytes;
  idx: Integer;
begin
  bytes := TFile.ReadAllBytes(AFileName);
  Result := '';
  for idx := 0 to ACount - 1 do
  begin
    if idx > High(bytes) then
    begin
      Break;
    end;
    Result := Result + Char(bytes[idx]);
  end;
end;

{ ===== Lebenszyklus ===== }

class procedure TTestEnvironment.Start;
begin
  FTestDataDir := FindTestDataDir;
  FOfficeWasRunning := IsOfficeRunning;
  // Die GUID im Pfad kennzeichnet die eigenen Dokumente – nur die werden geschlossen
  FRunId := TGUID.NewGuid.ToString.Replace('{', '').Replace('}', '');
  FTempDir := IncludeTrailingPathDelimiter(TPath.Combine(TPath.GetTempPath, 'OOAutomationTests\' + FRunId));
  FSpecialDir := FTempDir + 'Prüfung #1 50% ä\';
  ForceDirectories(FSpecialDir);
  FServiceManager := CreateOleObject('com.sun.star.ServiceManager');
  FDesktop := FServiceManager.createInstance('com.sun.star.frame.Desktop');
end;

class procedure TTestEnvironment.Finish;
begin
  try
    if not VarIsEmpty(FDesktop) then
    begin
      CloseOwnDocuments;
      // Nur beenden, was der Lauf selbst gestartet hat, und nie, solange noch fremde Komponenten offen sind
      if not FOfficeWasRunning and (Length(Components(False)) = 0) then
      begin
        if not Boolean(FDesktop.terminate) then
        begin
          raise Exception.Create('LibreOffice hat das Beenden nach dem Testlauf abgelehnt; ' +
            'soffice.bin läuft weiter und muss von Hand beendet werden.');
        end;
      end;
    end;
  finally
    FDesktop := Unassigned;
    FServiceManager := Unassigned;
  end;
  if (FTempDir <> '') and TDirectory.Exists(FTempDir) then
  begin
    TDirectory.Delete(FTempDir, True);
  end;
end;

{ ===== Testdaten ===== }

class function TTestEnvironment.FindTestDataDir: string;
begin
  Result := GetEnvironmentVariable('OOAUTOMATION_TESTDATA');
  if Result = '' then
  begin
    // IDE-Ausgabe liegt unter tests\<Plattform>\<Konfiguration>
    Result := ExpandFileName(ExtractFilePath(ParamStr(0)) + '..\..\testdata');
  end;
  Result := IncludeTrailingPathDelimiter(Result);
  if not FileExists(Result + 'test.odt') then
  begin
    raise Exception.CreateFmt('Testdaten nicht gefunden in "%s". Umgebungsvariable OOAUTOMATION_TESTDATA ' +
      'auf source\tests\testdata setzen.', [Result]);
  end;
end;

class function TTestEnvironment.CopyTestFile(const ASource, ATarget: string; AInSpecialDir: Boolean): string;
begin
  if AInSpecialDir then
  begin
    Result := FSpecialDir + ATarget;
  end
  else
  begin
    Result := FTempDir + ATarget;
  end;
  TFile.Copy(FTestDataDir + ASource, Result);
end;

class function TTestEnvironment.OdtContains(const AFileName, AText: string): Boolean;
var
  stream: TFileStream;
  zip: TZipFile;
  content: TBytes;
begin
  // Ohne Schreibsperre öffnen: LibreOffice hält die Datei offen, solange das Dokument geladen ist
  stream := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyNone);
  try
    zip := TZipFile.Create;
    try
      zip.Open(stream, zmRead);
      zip.Read('content.xml', content);
      Result := TEncoding.UTF8.GetString(content).Contains(AText);
    finally
      zip.Free;
    end;
  finally
    stream.Free;
  end;
end;

class function TTestEnvironment.CreateMarkedDocument(const ATarget: string): string;
var
  argument: OleVariant;
  document: OleVariant;
  text: OleVariant;
  cursor: OleVariant;
  bookmark: OleVariant;
  frame: OleVariant;
  pageStyle: OleVariant;
begin
  // Ein Dokument, das alles traegt, was die Reichweite des Ersetzens (B32) und das Lesen von
  // Textmarken (B31) pruefbar macht: das Wort MARKE in Fliesstext, Kopf, Fuss und Rahmen, dazu
  // eine punktfoermige und eine umspannende Textmarke.
  Result := FTempDir + ATarget;
  argument := FServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
  argument.Name := 'Hidden';
  argument.Value := True;
  document := FDesktop.loadComponentFromURL('private:factory/swriter', '_blank', 0,
    MakeSequence([argument]));
  try
    text := document.getText;
    cursor := text.createTextCursor;
    text.insertString(cursor, 'MARKE im Fliesstext ', False);

    // Punktfoermige Textmarke: bAbsorb = False, sie umspannt nichts
    bookmark := document.createInstance('com.sun.star.text.Bookmark');
    bookmark.setName('Punkt');
    text.insertTextContent(cursor, bookmark, False);

    // Umspannende Textmarke: erst schreiben, dann die Zeichen nach links auswaehlen und absorbieren
    text.insertString(cursor, 'SPANNE', False);
    cursor.goLeft(VarAsType(6, varSmallint), True);
    bookmark := document.createInstance('com.sun.star.text.Bookmark');
    bookmark.setName('Spanne');
    text.insertTextContent(cursor, bookmark, True);

    frame := document.createInstance('com.sun.star.text.TextFrame');
    text.insertTextContent(text.getEnd, frame, False);
    frame.getText.setString('MARKE im Rahmen');

    pageStyle := document.getStyleFamilies.getByName('PageStyles').getByName('Standard');
    pageStyle.HeaderIsOn := True;
    pageStyle.FooterIsOn := True;
    pageStyle.HeaderText.setString('MARKE im Kopf');
    pageStyle.FooterText.setString('MARKE im Fuss');

    argument := FServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    argument.Name := 'FilterName';
    argument.Value := OOFilterOdt;
    document.storeToURL(FileNameToUrl(Result), MakeSequence([argument]));
  finally
    document.close(True);
  end;
end;

class function TTestEnvironment.CreateCalcFile(const ATarget: string): string;
var
  argument: OleVariant;
  document: OleVariant;
begin
  // Eine echte Calc-Datei, damit TOOWriter das Nicht-Textdokument ablehnen kann (I20)
  Result := FTempDir + ATarget;
  argument := FServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
  argument.Name := 'Hidden';
  argument.Value := True;
  document := FDesktop.loadComponentFromURL('private:factory/scalc', '_blank', 0, MakeSequence([argument]));
  try
    argument := FServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
    argument.Name := 'FilterName';
    argument.Value := 'calc8';
    document.storeToURL(FileNameToUrl(Result), MakeSequence([argument]));
  finally
    document.close(True);
  end;
end;

{ ===== LibreOffice ===== }

class function TTestEnvironment.IsOfficeRunning: Boolean;
var
  snapshot: THandle;
  entry: TProcessEntry32;
  exeName: string;
begin
  Result := False;
  snapshot := CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0);
  if snapshot = INVALID_HANDLE_VALUE then
  begin
    RaiseLastOSError;
  end;
  try
    entry.dwSize := SizeOf(entry);
    if Process32First(snapshot, entry) then
    begin
      repeat
        exeName := entry.szExeFile;
        if SameText(exeName, 'soffice.bin') or SameText(exeName, 'soffice.exe') then
        begin
          Exit(True);
        end;
      until not Process32Next(snapshot, entry);
    end;
  finally
    CloseHandle(snapshot);
  end;
end;

class function TTestEnvironment.Components(AOwnOnly: Boolean): TArray<OleVariant>;
var
  enumeration: OleVariant;
  component: OleVariant;
begin
  Result := nil;
  enumeration := FDesktop.getComponents.createEnumeration;
  while enumeration.hasMoreElements do
  begin
    component := enumeration.nextElement;
    if not AOwnOnly then
    begin
      Result := Result + [component];
    end
    // Nur Dokumente haben eine URL; andere Komponenten (etwa die Basic-IDE) sind nie eigene
    else if component.supportsService('com.sun.star.document.OfficeDocument') then
    begin
      if ContainsText(component.getURL, FRunId) then
      begin
        Result := Result + [component];
      end;
    end;
  end;
end;

class function TTestEnvironment.IsDocumentOpen(const AFileName: string): Boolean;
var
  document: OleVariant;
  suffix: string;
begin
  // Vergleich über den Dateinamen: Die Testdateien heißen je Test eindeutig und rein ASCII. So hängt die
  // Prüfung nicht an der URL-Kodierung, die gerade getestet wird.
  suffix := '/' + ExtractFileName(AFileName);
  for document in Components(True) do
  begin
    if EndsText(suffix, document.getURL) then
    begin
      Exit(True);
    end;
  end;
  Result := False;
end;

class procedure TTestEnvironment.CloseOwnDocuments;
var
  document: OleVariant;
begin
  // close(True) statt close(): ohne Argument scheitert der Aufruf per OLE (B7)
  for document in Components(True) do
  begin
    document.close(True);
  end;
end;

end.
