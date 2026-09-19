{
--------------------------------------------------------------------------------
Autor: Wolfgang Lemmermeyer
Webseite: https://delphi-tutorials.de
Kontakt: lemmy@delphi-tutorials.de
Version: 0.2
Datum: 26.02.2005, überarbeitet 2026

Klasse für den OLE-Zugriff auf LibreOffice Writer

Copyright (c) 2005-2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}
unit OOWriter;

interface

uses
  System.Classes,
  System.Generics.Collections,
  OOObject,
  OOTable;

type
  TOOWriter = class(TOOObject)
  strict private
    FTables: TObjectList<TOOTable>;  // besitzt die TOOTable-Objekte
    function Bookmarks: OleVariant;
    function BookmarkNameList: string;
    function TableNameList: string;
  protected
    function PdfFilterName: string; override;
    procedure AfterLoadFile; override;
    procedure AfterCloseFile; override;
  public
    constructor Create;
    destructor Destroy; override;
    procedure WriteToBookmark(const AName, AValue: string);
    procedure WriteToBookmarks(AValues: TStrings);
    procedure GetBookmarkNames(AList: TStrings);
    procedure GetTableNames(AList: TStrings);
    function TableByName(const AName: string): TOOTable;
    function FindTable(const AName: string): TOOTable;
  end;

implementation

uses
  System.SysUtils,
  System.Variants,
  System.Win.ComObj,
  OOTools;

{ ===== Lebenszyklus ===== }

constructor TOOWriter.Create;
begin
  FTables := TObjectList<TOOTable>.Create(True);
  inherited Create;
end;

destructor TOOWriter.Destroy;
begin
  try
    // Schließt ein noch eigenes Dokument; AfterCloseFile leert dabei FTables
    inherited Destroy;
  finally
    FTables.Free;
  end;
end;

procedure TOOWriter.AfterLoadFile;
var
  tables: OleVariant;
  tableCount: Integer;
  idx: Integer;
begin
  inherited AfterLoadFile;
  if not Boolean(FDocument.supportsService('com.sun.star.text.TextDocument')) then
  begin
    raise EOOAutomation.CreateFmt('"%s" ist kein Textdokument; TOOWriter bearbeitet nur Writer-Dokumente.',
      [FileName]);
  end;
  tables := FDocument.getTextTables;
  tableCount := tables.getCount;
  for idx := 0 to tableCount - 1 do
  begin
    FTables.Add(TOOTable.Create(tables.getByIndex(idx)));
  end;
end;

procedure TOOWriter.AfterCloseFile;
begin
  // Geliehene TOOTable-Zeiger werden hier ungültig
  FTables.Clear;
  inherited AfterCloseFile;
end;

function TOOWriter.PdfFilterName: string;
begin
  Result := 'writer_pdf_Export';
end;

{ ===== Textmarken ===== }

procedure TOOWriter.WriteToBookmark(const AName, AValue: string);
var
  bookmark: OleVariant;
  anchor: OleVariant;
  cursor: OleVariant;
begin
  RequireDocument;
  // Textmarken sind case-sensitiv; LibreOffice selbst meldet den Fehlgriff ohne Text (B5)
  if not Boolean(Bookmarks.hasByName(AName)) then
  begin
    raise EOOAutomation.CreateFmt('Textmarke "%s" gibt es in "%s" nicht (Groß-/Kleinschreibung zählt). ' +
      'Vorhanden: %s', [AName, FileName, BookmarkNameList]);
  end;
  try
    bookmark := Bookmarks.getByName(AName);
    anchor := bookmark.getAnchor;
    cursor := anchor.getText.createTextCursorByRange(anchor);
    cursor.setString(AValue);
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Textmarke "%s" in "%s" nicht beschreibbar', [AName, FileName]), E);
    end;
  end;
end;

procedure TOOWriter.WriteToBookmarks(AValues: TStrings);
var
  missing: string;
  bookmarkName: string;
  idx: Integer;
begin
  RequireDocument;
  // Erst alle Namen prüfen, dann schreiben: ein Tippfehler hinterlässt kein halb gefülltes Dokument
  missing := '';
  for idx := 0 to AValues.Count - 1 do
  begin
    bookmarkName := AValues.Names[idx];
    if bookmarkName = '' then
    begin
      raise EOOAutomation.CreateFmt('Zeile %d ("%s") hat nicht die Form Textmarke=Wert.', [idx + 1, AValues[idx]]);
    end;
    if not Boolean(Bookmarks.hasByName(bookmarkName)) then
    begin
      missing := missing + ', "' + bookmarkName + '"';
    end;
  end;
  if missing <> '' then
  begin
    raise EOOAutomation.CreateFmt('Textmarken %s gibt es in "%s" nicht (Groß-/Kleinschreibung zählt); nichts ' +
      'wurde geschrieben. Vorhanden: %s', [missing.Substring(2), FileName, BookmarkNameList]);
  end;
  for idx := 0 to AValues.Count - 1 do
  begin
    WriteToBookmark(AValues.Names[idx], AValues.ValueFromIndex[idx]);
  end;
end;

procedure TOOWriter.GetBookmarkNames(AList: TStrings);
var
  names: OleVariant;
  idx: Integer;
begin
  RequireDocument;
  names := Bookmarks.getElementNames;
  AList.BeginUpdate;
  try
    AList.Clear;
    if VarIsArray(names) then
    begin
      for idx := VarArrayLowBound(names, 1) to VarArrayHighBound(names, 1) do
      begin
        AList.Add(names[idx]);
      end;
    end;
  finally
    AList.EndUpdate;
  end;
end;

function TOOWriter.Bookmarks: OleVariant;
begin
  try
    Result := FDocument.getBookmarks;
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Textmarken von "%s" nicht lesbar', [FileName]), E);
    end;
  end;
end;

function TOOWriter.BookmarkNameList: string;
var
  names: TStringList;
begin
  names := TStringList.Create;
  try
    GetBookmarkNames(names);
    names.QuoteChar := '"';
    names.Delimiter := ',';
    Result := names.DelimitedText;
  finally
    names.Free;
  end;
  if Result = '' then
  begin
    Result := '(keine)';
  end;
end;

{ ===== Tabellen ===== }

procedure TOOWriter.GetTableNames(AList: TStrings);
var
  table: TOOTable;
begin
  RequireDocument;
  AList.BeginUpdate;
  try
    AList.Clear;
    for table in FTables do
    begin
      AList.Add(table.Name);
    end;
  finally
    AList.EndUpdate;
  end;
end;

function TOOWriter.FindTable(const AName: string): TOOTable;
var
  table: TOOTable;
begin
  RequireDocument;
  // Tabellennamen sind wie Textmarken case-sensitiv
  for table in FTables do
  begin
    if table.Name = AName then
    begin
      Exit(table);
    end;
  end;
  Result := nil;
end;

function TOOWriter.TableByName(const AName: string): TOOTable;
begin
  Result := FindTable(AName);
  if Result = nil then
  begin
    raise EOOAutomation.CreateFmt('Tabelle "%s" gibt es in "%s" nicht. Vorhanden: %s',
      [AName, FileName, TableNameList]);
  end;
end;

function TOOWriter.TableNameList: string;
var
  names: TStringList;
begin
  names := TStringList.Create;
  try
    GetTableNames(names);
    names.QuoteChar := '"';
    names.Delimiter := ',';
    Result := names.DelimitedText;
  finally
    names.Free;
  end;
  if Result = '' then
  begin
    Result := '(keine)';
  end;
end;

end.
