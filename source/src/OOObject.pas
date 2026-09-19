unit OOObject;

{
--------------------------------------------------------------------------------
Autor: Wolfgang Lemmermeyer
Webseite: https://delphi-tutorials.de
Kontakt: lemmy@delphi-tutorials.de
Version: 0.2
Datum: 26.02.2005, überarbeitet 2026

Elternklasse für den OLE-Zugriff auf LibreOffice

Copyright (c) 2005-2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}

interface

uses
  OOTools;

type
  // Ein per OLE geladenes LibreOffice-Dokument. Ein Objekt hält höchstens ein Dokument; wer mehrere braucht,
  // legt mehrere Objekte an (Q1). Die COM-Initialisierung liegt beim Aufrufer (A7).
  TOOObject = class
  strict private
    function GetVisible: Boolean;
    procedure SetVisible(const AValue: Boolean);
    function ContainerWindow: OleVariant;
    procedure StoreDocument(const AFileName, AAction: string; AAsCopy: Boolean; const AFilterName: string);
    procedure ReleaseDocument;
  protected
    FServiceManager: OleVariant;
    FDesktop: OleVariant;
    FDocument: OleVariant;
    FFileName: string;
    function MakePropertyValue(const AName: string; const AValue: OleVariant): OleVariant;
    procedure RequireDocument;
    procedure ApplyPrinter(const AOptions: TOOPrintOptions);
    function PrintArgs(const AOptions: TOOPrintOptions): OleVariant;
    function PdfFilterName: string; virtual; abstract;
    procedure AfterLoadFile; virtual;
    procedure BeforeCloseFile; virtual;
    procedure AfterCloseFile; virtual;
  public
    constructor Create;
    destructor Destroy; override;
    procedure LoadFile(const AFileName: string; AHidden: Boolean = False);
    procedure Save;
    procedure SaveAs(const AFileName: string);
    procedure SaveCopyAs(const AFileName: string);
    procedure ExportPdf(const AFileName: string);
    procedure CloseFile(ASave: Boolean = False);
    procedure HandOver;
    procedure Print(const AOptions: TOOPrintOptions); virtual;
    function IsLoaded: Boolean;
    property FileName: string read FFileName;
    property Visible: Boolean read GetVisible write SetVisible;
  end;

implementation

uses
  System.SysUtils,
  System.Variants,
  System.Win.ComObj;

{ ===== Lebenszyklus ===== }

constructor TOOObject.Create;
begin
  inherited Create;
  try
    FServiceManager := CreateOleObject('com.sun.star.ServiceManager');
    FDesktop := FServiceManager.createInstance('com.sun.star.frame.Desktop');
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError('LibreOffice ist nicht erreichbar. Ist es installiert und als OLE-Server ' +
        '"com.sun.star.ServiceManager" registriert?', E);
    end;
  end;
end;

destructor TOOObject.Destroy;
begin
  // Was dem Objekt noch gehört, wird ohne Speichern geschlossen; mit HandOver abgegebene Dokumente nicht (E1).
  // Scheitert das Schließen, kommt die Meldung durch, statt verschluckt zu werden (Plan 7.2).
  if IsLoaded then
  begin
    CloseFile(False);
  end;
  inherited Destroy;
end;

function TOOObject.IsLoaded: Boolean;
begin
  Result := not VarIsEmpty(FDocument);
end;

procedure TOOObject.RequireDocument;
begin
  if not IsLoaded then
  begin
    raise EOOAutomation.Create('Kein Dokument geladen. Erst LoadFile aufrufen.');
  end;
end;

procedure TOOObject.ReleaseDocument;
begin
  FDocument := Unassigned;
  FFileName := '';
  AfterCloseFile;
end;

{ ===== Laden und Schließen ===== }

procedure TOOObject.LoadFile(const AFileName: string; AHidden: Boolean);
var
  url: string;
begin
  if IsLoaded then
  begin
    raise EOOAutomation.CreateFmt('Das Objekt hält bereits "%s". Erst CloseFile aufrufen oder für das ' +
      'weitere Dokument ein eigenes Objekt anlegen.', [FFileName]);
  end;
  url := FileNameToUrl(AFileName);
  if not FileExists(AFileName) then
  begin
    raise EOOAutomation.CreateFmt('Datei "%s" nicht gefunden.', [AFileName]);
  end;
  try
    FDocument := FDesktop.loadComponentFromURL(url, '_blank', 0,
      MakeSequence([MakePropertyValue('Hidden', AHidden)]));
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Laden von "%s" fehlgeschlagen', [AFileName]), E);
    end;
  end;
  if IsNullObject(FDocument) then
  begin
    FDocument := Unassigned;
    raise EOOAutomation.CreateFmt('LibreOffice hat für "%s" kein Dokument geliefert.', [AFileName]);
  end;
  FFileName := AFileName;
  try
    AfterLoadFile;
  except
    // Lehnt die Unterklasse das Dokument ab (etwa kein Textdokument), bleibt es nicht verwaist offen
    FDocument.close(True);
    ReleaseDocument;
    raise;
  end;
end;

procedure TOOObject.CloseFile(ASave: Boolean);
begin
  RequireDocument;
  if ASave then
  begin
    Save;
  end;
  BeforeCloseFile;
  try
    // close(True): Ohne Argument scheitert der Aufruf per OLE mit DISP_E_TYPEMISMATCH (B7)
    FDocument.close(True);
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Schließen von "%s" fehlgeschlagen', [FFileName]), E);
    end;
  end;
  ReleaseDocument;
end;

procedure TOOObject.HandOver;
begin
  RequireDocument;
  SetVisible(True);
  BeforeCloseFile;
  // Ab hier gehört das Dokument dem Benutzer; Destroy schließt es nicht mehr (E1)
  ReleaseDocument;
end;

{ ===== Speichern ===== }

procedure TOOObject.Save;
begin
  RequireDocument;
  try
    // store setzt das Dokument auf „unverändert“, storeToURL nicht. Ohne Klammern: Mit leeren Klammern schickt
    // Delphi ein Argument mit, und LibreOffice antwortet „Too many parameters“.
    FDocument.store;
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Speichern von "%s" fehlgeschlagen', [FFileName]), E);
    end;
  end;
end;

procedure TOOObject.SaveAs(const AFileName: string);
begin
  RequireDocument;
  StoreDocument(AFileName, 'Speichern unter', False, '');
  FFileName := AFileName;
end;

procedure TOOObject.SaveCopyAs(const AFileName: string);
begin
  RequireDocument;
  StoreDocument(AFileName, 'Kopie speichern', True, '');
end;

procedure TOOObject.ExportPdf(const AFileName: string);
begin
  RequireDocument;
  StoreDocument(AFileName, 'PDF-Export', True, PdfFilterName);
end;

procedure TOOObject.StoreDocument(const AFileName, AAction: string; AAsCopy: Boolean;
  const AFilterName: string);
var
  url: string;
  storeArgs: OleVariant;
begin
  url := FileNameToUrl(AFileName);
  if AFilterName = '' then
  begin
    storeArgs := MakeSequence([]);
  end
  else
  begin
    storeArgs := MakeSequence([MakePropertyValue('FilterName', AFilterName)]);
  end;
  try
    // Vorhandene Dateien werden überschrieben (LibreOffice-Vorgabe, B16)
    if AAsCopy then
    begin
      FDocument.storeToURL(url, storeArgs);
    end
    else
    begin
      FDocument.storeAsURL(url, storeArgs);
    end;
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('%s von "%s" nach "%s" fehlgeschlagen', [AAction, FFileName, AFileName]), E);
    end;
  end;
  // LibreOffice meldet auch dann Erfolg, wenn es an einen falsch gelesenen Pfad geschrieben hat (B14)
  if not FileExists(AFileName) then
  begin
    raise EOOAutomation.CreateFmt('%s meldete Erfolg, aber "%s" existiert nicht.', [AAction, AFileName]);
  end;
end;

{ ===== Drucken ===== }

procedure TOOObject.Print(const AOptions: TOOPrintOptions);
var
  printOptions: OleVariant;
begin
  RequireDocument;
  // Erst alles prüfen, dann den Drucker setzen und drucken: ein Fehler hinterlässt keinen halben Auftrag
  printOptions := PrintArgs(AOptions);
  ApplyPrinter(AOptions);
  try
    FDocument.print(printOptions);
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Drucken von "%s" fehlgeschlagen', [FFileName]), E);
    end;
  end;
end;

procedure TOOObject.ApplyPrinter(const AOptions: TOOPrintOptions);
var
  settings: TArray<OleVariant>;
begin
  RequireDocument;
  if (AOptions.PrinterName <> '') and not PrinterExists(AOptions.PrinterName) then
  begin
    raise EOOAutomation.CreateFmt('Drucker "%s" ist nicht installiert; LibreOffice würde sonst still auf den ' +
      'Standarddrucker ausweichen.', [AOptions.PrinterName]);
  end;
  settings := nil;
  if AOptions.PrinterName <> '' then
  begin
    settings := settings + [MakePropertyValue('Name', AOptions.PrinterName)];
  end;
  // Ohne OverridePaper bleiben Format und Ausrichtung des Dokuments unangetastet (E2)
  if AOptions.OverridePaper then
  begin
    // Ganzzahlen kommen als UNO-Enum an (B18); die Ordinalwerte entsprechen UNO (A11)
    settings := settings + [MakePropertyValue('PaperOrientation', Ord(AOptions.Orientation))];
    settings := settings + [MakePropertyValue('PaperFormat', Ord(AOptions.PaperFormat))];
  end;
  if Length(settings) = 0 then
  begin
    Exit;
  end;
  try
    FDocument.setPrinter(MakeSequence(settings));
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Druckereinstellung für "%s" fehlgeschlagen', [FFileName]), E);
    end;
  end;
end;

function TOOObject.PrintArgs(const AOptions: TOOPrintOptions): OleVariant;
var
  args: TArray<OleVariant>;
begin
  if (AOptions.Copies < 1) or (AOptions.Copies > High(SmallInt)) then
  begin
    raise EOOAutomation.CreateFmt('Kopienzahl %d ist ungültig; erlaubt sind 1 bis %d.',
      [AOptions.Copies, High(SmallInt)]);
  end;
  // Je Eintrag ein eigenes PropertyValue (B6). CopyCount ist in UNO ein short; ein OleVariant := SmallInt(…)
  // wird trotzdem varInteger, erst VarAsType ergibt VT_I2.
  args := [MakePropertyValue('CopyCount', VarAsType(AOptions.Copies, varSmallint)),
    MakePropertyValue('Collate', AOptions.Collate),
    MakePropertyValue('Wait', True)];
  if AOptions.Pages <> '' then
  begin
    args := args + [MakePropertyValue('Pages', AOptions.Pages)];
  end;
  Result := MakeSequence(args);
end;

{ ===== Sichtbarkeit ===== }

function TOOObject.ContainerWindow: OleVariant;
begin
  RequireDocument;
  try
    // Sichtbarkeit hängt am Container-Fenster des Rahmens, nicht am Dokument (B8)
    Result := FDocument.getCurrentController.getFrame.getContainerWindow;
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Fenster von "%s" nicht erreichbar', [FFileName]), E);
    end;
  end;
end;

function TOOObject.GetVisible: Boolean;
begin
  Result := ContainerWindow.isVisible;
end;

procedure TOOObject.SetVisible(const AValue: Boolean);
begin
  ContainerWindow.setVisible(AValue);
end;

{ ===== Helfer ===== }

function TOOObject.MakePropertyValue(const AName: string; const AValue: OleVariant): OleVariant;
begin
  // Je Aufruf ein neues Struct: Ein gemeinsames Objekt käme im Array nur mit dem letzten Stand an (B6)
  Result := FServiceManager.Bridge_GetStruct('com.sun.star.beans.PropertyValue');
  Result.Name := AName;
  Result.Value := AValue;
end;

procedure TOOObject.AfterLoadFile;
begin
end;

procedure TOOObject.BeforeCloseFile;
begin
end;

procedure TOOObject.AfterCloseFile;
begin
end;

end.
