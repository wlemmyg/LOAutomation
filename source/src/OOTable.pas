{
--------------------------------------------------------------------------------
Autor: Wolfgang Lemmermeyer
Webseite: https://delphi-tutorials.de
Kontakt: lemmy@delphi-tutorials.de
Version: 0.3.1
Datum: 12.11.2006, überarbeitet 2026

Hilfsklasse für den Tabellenzugriff

Copyright (c) 2006-2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}
unit OOTable;

interface

uses
  System.Classes,
  System.SysUtils,
  OOTools;

type
  // Eine Texttabelle des Dokuments. Das Objekt gehört dem TOOWriter (Aufrufer bekommen es nur geliehen) und
  // ist nach CloseFile, HandOver oder dem Laden eines anderen Dokuments ungültig.
  TOOTable = class
  strict private
    FTable: OleVariant;
    FName: string;
    FOnError: TOODocumentErrorFunc;  // Fehlerweg zum Besitzer (E6); geliehen, nicht freizugeben
    FValid: Boolean;
    function Error(const AContext: string; AError: Exception): EOOAutomation;
    procedure RequireValid;
    function CellByName(const ACellName: string): OleVariant;
    function RowCount: Integer;
    function ColumnCount: Integer;
  public
    constructor Create(const ATable: OleVariant; const AOnError: TOODocumentErrorFunc);
    function GetCell(const ACellName: string): string;
    function ReadAll: TArray<TArray<string>>;
    // Sagt der Tabelle, dass ihr Dokument nicht mehr offen ist. Der Besitzer ruft das beim Schliessen;
    // freigegeben wird die Tabelle erst spaeter (siehe TOOWriter.AfterCloseFile).
    procedure Invalidate;
    procedure SetCell(const ACellName, AValue: string);
    procedure SetCells(AValues: TStrings);
    procedure InsertRows(AAfterRow, ACount: Integer);
    procedure Fill(const AData: TArray<TArray<string>>; AStartRow: Integer = 1);
    property Name: string read FName;
  end;

implementation

uses
  System.Variants,
  System.Win.ComObj;

{ ===== Lebenszyklus ===== }

constructor TOOTable.Create(const ATable: OleVariant; const AOnError: TOODocumentErrorFunc);
begin
  inherited Create;
  FTable := ATable;
  FOnError := AOnError;
  FValid := True;
  FName := ATable.getName;
end;

procedure TOOTable.Invalidate;
begin
  FValid := False;
  FTable := Unassigned;
end;

procedure TOOTable.RequireValid;
begin
  // Ohne diesen Waechter redete die Tabelle stillschweigend mit einem toten UNO-Objekt und der Aufrufer
  // bekaeme einen COM-Fehler statt der Auskunft, was wirklich los ist
  if not FValid then
  begin
    raise EOOAutomation.CreateFmt('Tabelle "%s" gehört zu einem Dokument, das nicht mehr geladen ist. ' +
      'Weiter: das Dokument neu laden und die Tabelle erneut über TableByName holen.', [FName]);
  end;
end;

function TOOTable.Error(const AContext: string; AError: Exception): EOOAutomation;
begin
  // ACHTUNG: Der Besitzer laesst bei verlorenem Dokument los und gibt dabei diese Tabelle frei. Nach dem
  // Aufruf darf nichts mehr auf Felder zugreifen - deshalb ueberall nur "raise Error(...)" als letzte Tat.
  if Assigned(FOnError) then
  begin
    Result := FOnError(AContext, AError);
  end
  else
  begin
    Result := WrapUnoError(AContext, AError);
  end;
end;

{ ===== Zellen ===== }

function TOOTable.GetCell(const ACellName: string): string;
var
  cell: OleVariant;
begin
  RequireValid;
  cell := CellByName(ACellName);
  try
    Result := cell.getString;
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Zelle "%s" in Tabelle "%s" nicht lesbar', [ACellName, FName]), E);
    end;
  end;
end;

function TOOTable.ReadAll: TArray<TArray<string>>;
var
  rows: Integer;
  columns: Integer;
  idxRow: Integer;
  idxCol: Integer;
begin
  RequireValid;
  rows := RowCount;
  columns := ColumnCount;
  SetLength(Result, rows);
  try
    for idxRow := 0 to rows - 1 do
    begin
      SetLength(Result[idxRow], columns);
      for idxCol := 0 to columns - 1 do
      begin
        Result[idxRow][idxCol] := FTable.getCellByPosition(idxCol, idxRow).getString;
      end;
    end;
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Tabelle "%s" nicht lesbar', [FName]), E);
    end;
  end;
end;

procedure TOOTable.SetCell(const ACellName, AValue: string);
var
  cell: OleVariant;
begin
  RequireValid;
  cell := CellByName(ACellName);
  try
    cell.setString(AValue);
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Zelle "%s" in Tabelle "%s" nicht beschreibbar', [ACellName, FName]), E);
    end;
  end;
end;

procedure TOOTable.SetCells(AValues: TStrings);
var
  cells: TArray<OleVariant>;
  cellName: string;
  idx: Integer;
begin
  RequireValid;
  // Erst alle Zellen auflösen, dann schreiben: eine falsche Zelle hinterlässt keine halb gefüllte Tabelle
  SetLength(cells, AValues.Count);
  for idx := 0 to AValues.Count - 1 do
  begin
    cellName := AValues.Names[idx];
    if cellName = '' then
    begin
      raise EOOAutomation.CreateFmt('Zeile %d ("%s") für Tabelle "%s" hat nicht die Form Zelle=Wert.',
        [idx + 1, AValues[idx], FName]);
    end;
    cells[idx] := CellByName(cellName);
  end;
  try
    for idx := 0 to AValues.Count - 1 do
    begin
      cells[idx].setString(AValues.ValueFromIndex[idx]);
    end;
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Zellen in Tabelle "%s" nicht beschreibbar', [FName]), E);
    end;
  end;
end;

function TOOTable.CellByName(const ACellName: string): OleVariant;
begin
  try
    Result := FTable.getCellByName(ACellName);
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Zelle "%s" in Tabelle "%s" nicht erreichbar', [ACellName, FName]), E);
    end;
  end;
  // Eine unbekannte Zelle kommt als null zurück, nicht als Ausnahme (B19)
  if IsNullObject(Result) then
  begin
    raise EOOAutomation.CreateFmt('Zelle "%s" gibt es in Tabelle "%s" nicht.', [ACellName, FName]);
  end;
end;

{ ===== Zeilen ===== }

procedure TOOTable.InsertRows(AAfterRow, ACount: Integer);
var
  rows: Integer;
begin
  RequireValid;
  rows := RowCount;
  if (ACount < 1) or (AAfterRow < 0) or (AAfterRow > rows) then
  begin
    raise EOOAutomation.CreateFmt('InsertRows(%d, %d) passt nicht zu Tabelle "%s" mit %d Zeilen: erlaubt ist ' +
      'eine Zeile von 0 bis %d und eine Anzahl ab 1.', [AAfterRow, ACount, FName, rows, rows]);
  end;
  try
    // insertByIndex zählt ab 0 und fügt vor dieser Zeile ein – dieselbe Zahl wie „nach Zeile n“ (1-basiert);
    // insertByIndex(Zeilenzahl, n) hängt am Ende an (B19)
    FTable.getRows.insertByIndex(AAfterRow, ACount);
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Zeilen in Tabelle "%s" nicht einfügbar', [FName]), E);
    end;
  end;
end;

procedure TOOTable.Fill(const AData: TArray<TArray<string>>; AStartRow: Integer);
var
  rows: Integer;
  columns: Integer;
  needed: Integer;
  idxRow: Integer;
  idxCol: Integer;
begin
  RequireValid;
  rows := RowCount;
  columns := ColumnCount;
  if (AStartRow < 1) or (AStartRow > rows + 1) then
  begin
    raise EOOAutomation.CreateFmt('Startzeile %d liegt außerhalb von Tabelle "%s" (erlaubt 1 bis %d).',
      [AStartRow, FName, rows + 1]);
  end;
  // Erst prüfen, dann schreiben: Zu breite Daten werden nicht still abgeschnitten
  for idxRow := 0 to High(AData) do
  begin
    if Length(AData[idxRow]) > columns then
    begin
      raise EOOAutomation.CreateFmt('Datenzeile %d hat %d Spalten, Tabelle "%s" nur %d. Nichts wurde geschrieben.',
        [idxRow + 1, Length(AData[idxRow]), FName, columns]);
    end;
  end;
  needed := AStartRow - 1 + Length(AData);
  if needed > rows then
  begin
    // Fehlende Zeilen kommen ans Tabellenende (E4); Platz vor einer Fußzeile schafft der Aufrufer mit InsertRows
    InsertRows(rows, needed - rows);
  end;
  try
    for idxRow := 0 to High(AData) do
    begin
      for idxCol := 0 to High(AData[idxRow]) do
      begin
        FTable.getCellByPosition(idxCol, AStartRow - 1 + idxRow).setString(AData[idxRow][idxCol]);
      end;
    end;
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Tabelle "%s" nicht befüllbar', [FName]), E);
    end;
  end;
end;

function TOOTable.RowCount: Integer;
begin
  // Ungeschuetzt liess das einen rohen EOleSysError an der Bibliothek vorbei - und zwar aus InsertRows und
  // Fill heraus, die RowCount vor ihrem eigenen try aufrufen (Review 2026-09-20)
  try
    Result := FTable.getRows.getCount;
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Zeilenzahl von Tabelle "%s" nicht lesbar', [FName]), E);
    end;
  end;
end;

function TOOTable.ColumnCount: Integer;
begin
  try
    Result := FTable.getColumns.getCount;
  except
    on E: EOleSysError do
    begin
      raise Error(Format('Spaltenzahl von Tabelle "%s" nicht lesbar', [FName]), E);
    end;
  end;
end;

end.
