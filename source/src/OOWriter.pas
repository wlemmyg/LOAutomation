{
--------------------------------------------------------------------------------
Autor: Wolfgang Lemmermeyer
Webseite: https://delphi-tutorials.de
Kontakt: lemmy@delphi-tutorials.de
Version: 0.3.1
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
  OOTable,
  OOTools;

type
  TOOWriter = class(TOOObject)
  strict private
    FTables: TObjectList<TOOTable>;       // geliehene Sicht auf das aktuelle Dokument, besitzt nichts
    FOwnedTables: TObjectList<TOOTable>;  // besitzt alle je erzeugten TOOTable-Objekte
    function Bookmarks: OleVariant;
    function BookmarkNameList: string;
    function BookmarkAnchor(const AName: string): OleVariant;
    function GraphicFromFile(const AFileName: string): OleVariant;
    procedure InsertImage(const ABookmark, AFileName: string; AWidth, AHeight: Integer;
      AUseGivenSize: Boolean);
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
    function ReplaceAll(const ASearch, AReplace: string): Integer; overload;
    function ReplaceAll(const ASearch, AReplace: string; const AOptions: TOOSearchOptions): Integer; overload;
    function ReadBookmark(const AName: string): string;
    procedure InsertImageAtBookmark(const ABookmark, AFileName: string); overload;
    procedure InsertImageAtBookmark(const ABookmark, AFileName: string; AWidth, AHeight: Integer); overload;
    function TableByName(const AName: string): TOOTable;
    function FindTable(const AName: string): TOOTable;
  end;

implementation

uses
  System.SysUtils,
  System.Variants,
  System.Win.ComObj;

{ ===== Lebenszyklus ===== }

constructor TOOWriter.Create;
begin
  // Besitz und Sicht sind getrennt: Beim Schliessen wird die Sicht geleert, zerstoert wird erst beim
  // naechsten Laden oder hier im Destruktor. Sonst gaebe sich eine Tabelle, die gerade ein verlorenes
  // Dokument meldet, mitten im eigenen Methodenaufruf selbst frei.
  FOwnedTables := TObjectList<TOOTable>.Create(True);
  FTables := TObjectList<TOOTable>.Create(False);
  inherited Create;
end;

destructor TOOWriter.Destroy;
begin
  try
    // Schließt ein noch eigenes Dokument; AfterCloseFile leert dabei FTables
    inherited Destroy;
  finally
    FTables.Free;
    FOwnedTables.Free;
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
  // Hier ist kein Tabellen-Aufruf auf dem Stapel, also koennen die alten jetzt gefahrlos weg
  FOwnedTables.Clear;
  tables := FDocument.getTextTables;
  tableCount := tables.getCount;
  for idx := 0 to tableCount - 1 do
  begin
    // Der Fehlerweg des Besitzers: nur so kann eine geliehene Tabelle ein verlorenes Dokument melden
    // und loslassen (E6, Review 2026-09-20)
    FOwnedTables.Add(TOOTable.Create(tables.getByIndex(idx), DocumentError));
    FTables.Add(FOwnedTables.Last);
  end;
end;

procedure TOOWriter.AfterCloseFile;
var
  table: TOOTable;
begin
  // Geliehene TOOTable-Zeiger werden hier ungültig – aber NICHT freigegeben: Dieser Weg läuft auch aus einer
  // Tabelle heraus, die gerade ein verlorenes Dokument meldet (E6). Sie sagt es ab jetzt selbst, zerstört
  // wird sie beim nächsten Laden oder im Destruktor.
  for table in FTables do
  begin
    table.Invalidate;
  end;
  FTables.Clear;
  inherited AfterCloseFile;
end;

function TOOWriter.ReplaceAll(const ASearch, AReplace: string): Integer;
begin
  Result := ReplaceAll(ASearch, AReplace, TOOSearchOptions.Default);
end;

function TOOWriter.ReplaceAll(const ASearch, AReplace: string; const AOptions: TOOSearchOptions): Integer;
var
  descriptor: OleVariant;
begin
  RequireDocument;
  if ASearch = '' then
  begin
    raise EOOAutomation.Create('Kein Suchtext angegeben.');
  end;
  try
    descriptor := FDocument.createReplaceDescriptor;
    descriptor.SearchString := ASearch;
    descriptor.ReplaceString := AReplace;
    descriptor.SearchCaseSensitive := AOptions.CaseSensitive;
    descriptor.SearchWords := AOptions.WholeWords;
    descriptor.SearchRegularExpression := AOptions.RegularExpression;
    // Das Ergebnis ist die Anzahl der Ersetzungen; LibreOffice erfasst dabei auch Kopf-, Fußzeile und
    // Textrahmen (B32). Kein Treffer ist kein Fehler, sondern eine 0 (Q2)
    Result := FDocument.replaceAll(descriptor);
  except
    on E: EOleSysError do
    begin
      raise DocumentError(Format('Ersetzen von "%s" in "%s" fehlgeschlagen', [ASearch, FileName]), E);
    end;
  end;
end;

function TOOWriter.BookmarkAnchor(const AName: string): OleVariant;
begin
  RequireDocument;
  if not Boolean(Bookmarks.hasByName(AName)) then
  begin
    raise EOOAutomation.CreateFmt('Textmarke "%s" gibt es in "%s" nicht (Groß-/Kleinschreibung zählt). ' +
      'Vorhanden: %s', [AName, FileName, BookmarkNameList]);
  end;
  try
    Result := Bookmarks.getByName(AName).getAnchor;
  except
    on E: EOleSysError do
    begin
      raise DocumentError(Format('Textmarke "%s" in "%s" nicht erreichbar', [AName, FileName]), E);
    end;
  end;
end;

function TOOWriter.ReadBookmark(const AName: string): string;
var
  anchor: OleVariant;
begin
  anchor := BookmarkAnchor(AName);
  try
    Result := anchor.getString;
  except
    on E: EOleSysError do
    begin
      raise DocumentError(Format('Textmarke "%s" in "%s" nicht lesbar', [AName, FileName]), E);
    end;
  end;
  // Eine punktförmige Textmarke umspannt nichts und liefert immer '' - auch direkt nach dem Schreiben
  // (B31). Ein leerer Rückgabewert wäre von einem leeren Feld nicht zu unterscheiden, also lieber laut (Q7)
  if Result = '' then
  begin
    raise EOOAutomation.CreateFmt('Textmarke "%s" in "%s" umspannt keinen Text; punktförmige Textmarken ' +
      'haben keinen lesbaren Inhalt. Weiter: im Dokument die Textmarke über den gewünschten Text legen.',
      [AName, FileName]);
  end;
end;

function TOOWriter.GraphicFromFile(const AFileName: string): OleVariant;
var
  provider: OleVariant;
begin
  if not FileExists(AFileName) then
  begin
    raise EOOAutomation.CreateFmt('Bilddatei "%s" nicht gefunden.', [AFileName]);
  end;
  try
    // Über den GraphicProvider statt ueber die seit LibreOffice 6.1 veraltete Eigenschaft GraphicURL (B29)
    provider := FServiceManager.createInstance('com.sun.star.graphic.GraphicProvider');
    Result := provider.queryGraphic(MakeSequence([MakePropertyValue('URL', FileNameToUrl(AFileName))]));
  except
    on E: EOleSysError do
    begin
      raise WrapUnoError(Format('Bild "%s" nicht ladbar', [AFileName]), E);
    end;
  end;
  // queryGraphic meldet ein unlesbares Bild still als null statt als Ausnahme (B29)
  if IsNullObject(Result) then
  begin
    raise EOOAutomation.CreateFmt('LibreOffice konnte "%s" nicht als Bild lesen.', [AFileName]);
  end;
end;

procedure TOOWriter.InsertImage(const ABookmark, AFileName: string; AWidth, AHeight: Integer;
  AUseGivenSize: Boolean);
var
  anchor: OleVariant;
  image: OleVariant;
  naturalSize: OleVariant;
begin
  // Erst alles besorgen, dann einfügen: ein Fehler hinterlässt kein halbes Bild im Dokument
  anchor := BookmarkAnchor(ABookmark);
  image := FDocument.createInstance('com.sun.star.text.TextGraphicObject');
  image.Graphic := GraphicFromFile(AFileName);
  try
    // AS_CHARACTER: das Bild sitzt im Textfluss; LibreOffice hängt es sonst an den Absatz (B33).
    // Muss VOR dem Einfügen stehen: nachträglich gesetzt hängt LibreOffice das Bild um und es
    // landet am Absatzende statt an der Textmarke (B38)
    image.AnchorType := 1;
    anchor.getText.insertTextContent(anchor, image, False);
    // Frisch eingefuegt steht es auf 566 x 566, unabhängig vom Bild (B30) – die Größe muss gesetzt werden
    if AUseGivenSize then
    begin
      image.Width := AWidth;
      image.Height := AHeight;
    end
    else
    begin
      // ActualSize trägt erst nach dem Einfügen die natürliche Größe, Seitenverhältnis inklusive (B33)
      naturalSize := image.ActualSize;
      image.Width := naturalSize.Width;
      image.Height := naturalSize.Height;
    end;
  except
    on E: EOleSysError do
    begin
      raise DocumentError(Format('Bild "%s" an Textmarke "%s" in "%s" nicht einfügbar',
        [AFileName, ABookmark, FileName]), E);
    end;
  end;
end;

procedure TOOWriter.InsertImageAtBookmark(const ABookmark, AFileName: string);
begin
  InsertImage(ABookmark, AFileName, 0, 0, False);
end;

procedure TOOWriter.InsertImageAtBookmark(const ABookmark, AFileName: string; AWidth, AHeight: Integer);
begin
  if (AWidth < 1) or (AHeight < 1) then
  begin
    raise EOOAutomation.CreateFmt('Breite %d und Höhe %d müssen mindestens 1 sein (Maße in 1/100 mm). ' +
      'Für die natürliche Größe die Überladung ohne Maße nehmen.', [AWidth, AHeight]);
  end;
  InsertImage(ABookmark, AFileName, AWidth, AHeight, True);
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
      raise DocumentError(Format('Textmarke "%s" in "%s" nicht beschreibbar', [AName, FileName]), E);
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
  // Ohne Dokument waere FDocument leer, und der Zugriff darauf ergaebe EVariantInvalidOpError - eine
  // Ausnahme, die der Handler unten gar nicht faengt
  RequireDocument;
  try
    Result := FDocument.getBookmarks;
  except
    on E: EOleSysError do
    begin
      raise DocumentError(Format('Textmarken von "%s" nicht lesbar', [FileName]), E);
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
