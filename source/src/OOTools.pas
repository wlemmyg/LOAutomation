unit OOTools;

interface

{
--------------------------------------------------------------------------------
Autor: Wolfgang Lemmermeyer
Webseite: https://delphi-tutorials.de
Kontakt: lemmy@delphi-tutorials.de
Version: 0.3.1
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

  // Papierformat und Ausrichtung gibt es hier bewusst nicht: Sie kommen aus der Seitenvorlage des Dokuments.
  // Per setPrinter wirkt die Ausrichtung gar nicht, und das Format formatiert das Dokument selbst um (B25, B26, E5).
  TOOPrintOptions = record
    PrinterName: string;       // leer = Drucker, den das Dokument bereits hat
    Copies: Integer;
    Collate: Boolean;          // mehrere Kopien sortiert: 1-2-3, 1-2-3
    Pages: string;             // leer = alle, sonst LibreOffice-Syntax, z. B. '1-3;5'
    class function Default: TOOPrintOptions; static;
  end;

  // Optionen fuer Suchen und Ersetzen. Die Vorgaben sind die von LibreOffice selbst (B27).
  TOOSearchOptions = record
    CaseSensitive: Boolean;
    WholeWords: Boolean;
    RegularExpression: Boolean;
    class function Default: TOOSearchOptions; static;
  end;

const
  // Gebraeuchliche Filternamen fuer SaveCopyAs; LibreOffice kennt weit mehr (B28)
  OOFilterOdt = 'writer8';
  OOFilterDocx = 'MS Word 2007 XML';
  OOFilterRtf = 'Rich Text Format';
  OOFilterText = 'Text';
  OOFilterPdf = 'writer_pdf_Export';

// Laufwerks- oder UNC-Pfad als UTF-8-%-kodierte file-URL; relative Pfade lehnt sie ab (F4)
function FileNameToUrl(const AFileName: string): string;

// Ist ein Drucker dieses Namens lokal installiert oder verbunden?
function PrinterExists(const AName: string): Boolean;

type
  // Wie ein Dokument verloren gehen kann (E6)
  TOODocumentLoss = (dlNone, dlOfficeGone, dlDocumentClosed);

  // Fehlerweg eines geliehenen Objekts zurueck zu seinem Besitzer: packt ein und laesst ein verlorenes
  // Dokument los (E6). TOOTable bekommt ihn vom TOOWriter.
  TOODocumentErrorFunc = function(const AContext: string; AError: Exception): EOOAutomation of object;

// Sagt ein Fehler, dass das Dokument verloren ist? Dann hilft kein Wiederholen (E6)
function DocumentLoss(AError: Exception): TOODocumentLoss;

// Packt einen UNO-/OLE-Fehler samt Zusammenhang in EOOAutomation ein (A5); bei verlorenem Dokument mit
// Klartext statt „RPC-Server nicht verfügbar“ (E6)
function WrapUnoError(const AContext: string; AError: Exception): EOOAutomation; overload;

// Dieselbe Meldung, aber mit einem bereits ermittelten Befund - fuer Aufrufer, die den Verlust nicht am
// Fehlertext erkennen, sondern beim Dokument nachgefragt haben (B35, B36)
function WrapUnoError(const AContext: string; AError: Exception;
  ALoss: TOODocumentLoss): EOOAutomation; overload;

// Leerer UNO-Verweis: LibreOffice liefert „nicht gefunden“ teils als null statt als Ausnahme (B19)
function IsNullObject(const AValue: OleVariant): Boolean;

// UNO-Sequenz (Variant-Array) aus einzelnen Werten; leer ergibt eine leere Sequenz
function MakeSequence(const AItems: array of OleVariant): OleVariant;

implementation

uses
  System.Variants,
  System.Win.ComObj,
  Winapi.Windows,
  Winapi.WinSpool;

const
  // RFC 3986 „unreserved“ – alles andere wird kodiert (T2)
  CUnreserved: set of AnsiChar = ['A'..'Z', 'a'..'z', '0'..'9', '-', '.', '_', '~'];

{ ===== TOOPrintOptions ===== }

class function TOOSearchOptions.Default: TOOSearchOptions;
begin
  // LibreOffice sucht per Vorgabe ohne Ruecksicht auf Gross-/Kleinschreibung, ohne Wortgrenzen
  // und ohne regulaere Ausdruecke (B27)
  Result.CaseSensitive := False;
  Result.WholeWords := False;
  Result.RegularExpression := False;
end;

class function TOOPrintOptions.Default: TOOPrintOptions;
begin
  Result.PrinterName := '';
  Result.Copies := 1;
  Result.Collate := True;
  Result.Pages := '';
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

function DocumentLoss(AError: Exception): TOODocumentLoss;
const
  // COM-Fehlercodes, bei denen der Server weg ist: RPC_S_SERVER_UNAVAILABLE, RPC_S_CALL_FAILED(_DNE),
  // RPC_E_DISCONNECTED, RPC_E_SERVERDIED(_DNE), CO_E_OBJNOTCONNECTED
  CGoneCodes: array[0..6] of HRESULT = (HRESULT($800706BA), HRESULT($800706BE), HRESULT($800706BF),
    HRESULT($80010108), HRESULT($80010007), HRESULT($80010012), HRESULT($800401FD));
var
  idx: Integer;
begin
  Result := dlNone;
  // LibreOffice lebt, aber das Dokument ist dort geschlossen worden
  if AError.Message.Contains('com.sun.star.lang.DisposedException') then
  begin
    Exit(dlDocumentClosed);
  end;
  if AError is EOleSysError then
  begin
    for idx := Low(CGoneCodes) to High(CGoneCodes) do
    begin
      if EOleSysError(AError).ErrorCode = CGoneCodes[idx] then
      begin
        Exit(dlOfficeGone);
      end;
    end;
  end;
end;

function WrapUnoError(const AContext: string; AError: Exception): EOOAutomation;
begin
  Result := WrapUnoError(AContext, AError, DocumentLoss(AError));
end;

function WrapUnoError(const AContext: string; AError: Exception; ALoss: TOODocumentLoss): EOOAutomation;
var
  detail: string;
begin
  detail := Trim(AError.Message);
  if detail = '' then
  begin
    // LibreOffice liefert manche Ausnahmen ohne Meldungstext (B5)
    detail := AError.ClassName + ' ohne Meldungstext';
  end;
  // Der technische Text bleibt am Ende stehen, er hilft beim Nachforschen
  case ALoss of
    dlOfficeGone:
      Result := EOOAutomation.CreateFmt('%s: LibreOffice ist nicht mehr erreichbar – es wurde beendet oder ist ' +
        'abgestürzt. Das Dokument ist für dieses Objekt verloren; nicht gespeicherte Änderungen bietet ' +
        'LibreOffice beim nächsten Start eventuell zur Wiederherstellung an. Weiter: LibreOffice starten und das ' +
        'Dokument neu laden. (%s)', [AContext, detail]);
    dlDocumentClosed:
      Result := EOOAutomation.CreateFmt('%s: Das Dokument ist in LibreOffice nicht mehr offen – es wurde dort ' +
        'geschlossen. Weiter: das Dokument neu laden. (%s)', [AContext, detail]);
  else
    Result := EOOAutomation.CreateFmt('%s: %s', [AContext, detail]);
  end;
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
