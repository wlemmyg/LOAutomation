unit OOTools;

interface

{
--------------------------------------------------------------------------------
Autor: Wolfgang Lemmermeyer
Webseite: https://delphi-tutorials.de
Kontakt: lemmy@delphi-tutorials.de
Version: 0.2
Datum: 26.02.2005, überarbeitet 2026

Übergreifende Typen und Hilfsfunktionen, ohne LibreOffice testbar

Copyright (c) 2005-2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}

uses
  System.SysUtils;

type
  // Ausnahme der Bibliothek: Die Meldung sagt, was scheiterte und welcher Weg offensteht (A5)
  EOOAutomation = class(Exception);

  // Reihenfolge = com.sun.star.view.PaperFormat; der Ordinalwert geht unverändert an LibreOffice (A11)
  TOOPaperFormat = (pfA3, pfA4, pfA5, pfB4, pfB5, pfLetter, pfLegal, pfTabloid, pfUser);
  // Reihenfolge = com.sun.star.view.PaperOrientation
  TOOOrientation = (ooPortrait, ooLandscape);

  TOOPrintOptions = record
    PrinterName: string;       // leer = Drucker, den das Dokument bereits hat
    Copies: Integer;
    Collate: Boolean;          // mehrere Kopien sortiert: 1-2-3, 1-2-3
    Pages: string;             // leer = alle, sonst LibreOffice-Syntax, z. B. '1-3;5'
    OverridePaper: Boolean;    // nur dann gelten Orientation und PaperFormat, sonst bleibt das Dokument, wie es ist
    Orientation: TOOOrientation;
    PaperFormat: TOOPaperFormat;
    class function Default: TOOPrintOptions; static;
  end;

// Laufwerks- oder UNC-Pfad als UTF-8-%-kodierte file-URL; relative Pfade lehnt sie ab (F4)
function FileNameToUrl(const AFileName: string): string;

// Ist ein Drucker dieses Namens lokal installiert oder verbunden?
function PrinterExists(const AName: string): Boolean;

// Packt einen UNO-/OLE-Fehler samt Zusammenhang in EOOAutomation ein (A5)
function WrapUnoError(const AContext: string; AError: Exception): EOOAutomation;

// Leerer UNO-Verweis: LibreOffice liefert „nicht gefunden“ teils als null statt als Ausnahme (B19)
function IsNullObject(const AValue: OleVariant): Boolean;

// UNO-Sequenz (Variant-Array) aus einzelnen Werten; leer ergibt eine leere Sequenz
function MakeSequence(const AItems: array of OleVariant): OleVariant;

implementation

uses
  System.Variants,
  Winapi.Windows,
  Winapi.WinSpool;

const
  // RFC 3986 „unreserved“ – alles andere wird kodiert (T2)
  CUnreserved: set of AnsiChar = ['A'..'Z', 'a'..'z', '0'..'9', '-', '.', '_', '~'];

{ ===== TOOPrintOptions ===== }

class function TOOPrintOptions.Default: TOOPrintOptions;
begin
  Result.PrinterName := '';
  Result.Copies := 1;
  Result.Collate := True;
  Result.Pages := '';
  Result.OverridePaper := False;
  Result.Orientation := ooPortrait;
  Result.PaperFormat := pfA4;
end;

{ ===== Pfade ===== }

function EncodeSegment(const ASegment: string): string;
var
  bytes: TBytes;
  idx: Integer;
begin
  Result := '';
  bytes := TEncoding.UTF8.GetBytes(ASegment);
  for idx := 0 to High(bytes) do
  begin
    if AnsiChar(bytes[idx]) in CUnreserved then
    begin
      Result := Result + Char(bytes[idx]);
    end
    else
    begin
      Result := Result + '%' + IntToHex(bytes[idx], 2);
    end;
  end;
end;

function FileNameToUrl(const AFileName: string): string;
var
  path: string;
  segments: TArray<string>;
  isUnc: Boolean;
  isDrive: Boolean;
  idx: Integer;
begin
  if AFileName = '' then
  begin
    raise EOOAutomation.Create('Kein Dateiname angegeben.');
  end;
  path := AFileName.Replace('/', '\');
  isUnc := path.StartsWith('\\');
  isDrive := (Length(path) >= 3) and CharInSet(path[1], ['A'..'Z', 'a'..'z']) and (path[2] = ':') and
    (path[3] = '\');
  // Auch \Ordner\x.odt (laufwerksrelativ) und C:x.odt (verzeichnisrelativ) sind relativ (T3)
  if not isUnc and not isDrive then
  begin
    raise EOOAutomation.CreateFmt('Pfad "%s" ist nicht absolut. Erwartet wird ein Laufwerks- oder UNC-Pfad; ' +
      'einen relativen Pfad löst ExpandFileName auf.', [AFileName]);
  end;
  if isUnc then
  begin
    Result := 'file://';
    segments := path.Substring(2).Split(['\']);
  end
  else
  begin
    Result := 'file:///' + path.Substring(0, 2) + '/';
    segments := path.Substring(3).Split(['\']);
  end;
  for idx := 0 to High(segments) do
  begin
    if idx > 0 then
    begin
      Result := Result + '/';
    end;
    Result := Result + EncodeSegment(segments[idx]);
  end;
end;

{ ===== Drucker ===== }

function PrinterExists(const AName: string): Boolean;
var
  needed: DWORD;
  returned: DWORD;
  buffer: TBytes;
  info: PPrinterInfo4;
  idx: Integer;
begin
  needed := 0;
  returned := 0;
  // Erster Aufruf ermittelt nur die Puffergröße; 0 heißt: kein Drucker eingerichtet
  EnumPrinters(PRINTER_ENUM_LOCAL or PRINTER_ENUM_CONNECTIONS, nil, 4, nil, 0, needed, returned);
  if needed = 0 then
  begin
    Exit(False);
  end;
  SetLength(buffer, needed);
  if not EnumPrinters(PRINTER_ENUM_LOCAL or PRINTER_ENUM_CONNECTIONS, nil, 4, @buffer[0], needed, needed,
    returned) then
  begin
    RaiseLastOSError;
  end;
  info := PPrinterInfo4(@buffer[0]);
  for idx := 1 to Integer(returned) do
  begin
    if SameText(info^.pPrinterName, AName) then
    begin
      Exit(True);
    end;
    Inc(info);
  end;
  Result := False;
end;

{ ===== UNO-Helfer ===== }

function WrapUnoError(const AContext: string; AError: Exception): EOOAutomation;
var
  detail: string;
begin
  detail := Trim(AError.Message);
  if detail = '' then
  begin
    // LibreOffice liefert manche Ausnahmen ohne Meldungstext (B5)
    detail := AError.ClassName + ' ohne Meldungstext';
  end;
  Result := EOOAutomation.CreateFmt('%s: %s', [AContext, detail]);
end;

function IsNullObject(const AValue: OleVariant): Boolean;
begin
  case TVarData(AValue).VType of
    varEmpty, varNull:
      Result := True;
    varDispatch:
      Result := TVarData(AValue).VDispatch = nil;
    varUnknown:
      Result := TVarData(AValue).VUnknown = nil;
  else
    Result := False;
  end;
end;

function MakeSequence(const AItems: array of OleVariant): OleVariant;
var
  idx: Integer;
begin
  Result := VarArrayCreate([0, High(AItems)], varVariant);
  for idx := 0 to High(AItems) do
  begin
    Result[idx] := AItems[idx];
  end;
end;

end.
