{
--------------------------------------------------------------------------------
Copyright (c) 2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}
unit Test.OOWriter;

// Integrationstests gegen LibreOffice: I1–I17 aus P2 (Plan 7.1), seit P3 auf der API aus Abschnitt 5, dazu
// I18–I35 für das neue Verhalten (Plan 7.2). Gedruckt wird nie – geprüft werden nur Drucker- und
// Druckargumente.

interface

uses
  DUnitX.TestFramework,
  OOWriter;

const
  // Notbremse gegen modale LibreOffice-Dialoge, die jeden Aufruf blockieren (T8)
  CMaxTime = 60000;

type
  [TestFixture]
  [Category('LibreOffice')]
  TOOWriterTests = class
  strict private
    FWriter: TOOWriter;
    function LoadCopy(const ASource, ATarget: string; AInSpecialDir: Boolean = False): string;
    function DocumentText: string;
    function CellText(const ATable, ACell: string): string;
    function RowCount(const ATable: string): Integer;
    function PrinterSetting(const AName: string): OleVariant;
    function HeaderString: string;
    function FooterString: string;
    function FrameString: string;
    function GraphicCount: Integer;
    function GraphicWidth(AIndex: Integer): Integer;
    function GraphicHeight(AIndex: Integer): Integer;
    function GraphicAnchor(AIndex: Integer): Integer;
    function ParagraphLayout: string;
  public
    [SetupFixture]
    procedure SetupFixture;
    [TearDownFixture]
    procedure TearDownFixture;
    [Setup]
    procedure Setup;
    [TearDown]
    procedure TearDown;

    // I1–I2, I18–I20: Laden
    [Test]
    [MaxTime(CMaxTime)]
    procedure LoadHidden;
    [Test]
    [MaxTime(CMaxTime)]
    procedure LoadFromSpecialPath;
    [Test]
    [MaxTime(CMaxTime)]
    procedure LoadMissingFileRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure NoDocumentRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure LoadNonTextDocumentRaises;

    // I3–I6, I25: Textmarken
    [Test]
    [MaxTime(CMaxTime)]
    procedure WriteBookmark;
    [Test]
    [MaxTime(CMaxTime)]
    procedure WriteBookmarkList;
    [Test]
    [MaxTime(CMaxTime)]
    procedure WrongCaseBookmarkRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure BookmarkNames;
    [Test]
    [MaxTime(CMaxTime)]
    procedure WriteBookmarksChecksAllFirst;

    // I7–I8, I26–I30: Tabellen
    [Test]
    [MaxTime(CMaxTime)]
    procedure TableNamesAndCells;
    [Test]
    [MaxTime(CMaxTime)]
    procedure UnknownCellRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure TableLookup;
    [Test]
    [MaxTime(CMaxTime)]
    procedure SetCellsChecksAllFirst;
    [Test]
    [MaxTime(CMaxTime)]
    procedure InsertRowsInMiddleAndAtEnd;
    [Test]
    [MaxTime(CMaxTime)]
    procedure FillAppendsMissingRows;
    [Test]
    [MaxTime(CMaxTime)]
    procedure FillTooWideRaises;

    // I9–I12, I23–I24: Speichern und Export
    [Test]
    [MaxTime(CMaxTime)]
    procedure SaveResetsModified;
    [Test]
    [MaxTime(CMaxTime)]
    procedure SaveAsRenamesDocument;
    [Test]
    [MaxTime(CMaxTime)]
    procedure SaveAsToSpecialPath;
    [Test]
    [MaxTime(CMaxTime)]
    procedure SaveCopyAsKeepsName;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ExportPdfWritesPdf;
    [Test]
    [MaxTime(CMaxTime)]
    procedure SaveCopyAsOverwrites;

    // I13–I17, I21–I22: Schließen, Freigeben, Übergabe, Sichtbarkeit, PropertyValues, zweites Laden
    [Test]
    [MaxTime(CMaxTime)]
    procedure CloseWithoutSaving;
    [Test]
    [MaxTime(CMaxTime)]
    procedure CloseWithSaving;
    [Test]
    [MaxTime(CMaxTime)]
    procedure FreeClosesDocument;
    [Test]
    [MaxTime(CMaxTime)]
    procedure PropertyValuesAreIndependent;
    [Test]
    [MaxTime(CMaxTime)]
    procedure LoadTwiceRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure HandOverLeavesDocumentOpen;
    [Test]
    [MaxTime(CMaxTime)]
    procedure VisibleFollowsContainerWindow;

    // I31, I34, I35: Drucker und Druckargumente (I32/I33 fielen mit der Papier-Vorgabe weg, E5)
    [Test]
    [MaxTime(CMaxTime)]
    procedure UnknownPrinterRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure PrintArgsAreSeparateAndTyped;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ZeroCopiesRaises;
    // I36: Druckeinstellungen verändern das Dokument nicht (B26, E5)
    [Test]
    [MaxTime(CMaxTime)]
    procedure PrintSettingsLeaveDocumentUntouched;
    // I37: in LibreOffice geschlossenes Dokument – Klartext, und das Objekt lässt es los (E6)
    [Test]
    [MaxTime(CMaxTime)]
    procedure ClosedDocumentIsReported;

    // I38-I41: Suchen und Ersetzen (B27, B32, Q2)
    [Test]
    [MaxTime(CMaxTime)]
    procedure ReplaceAllReturnsCount;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ReplaceAllWithoutHitsReturnsZero;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ReplaceAllCoversHeaderFooterAndFrame;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ReplaceAllCaseSensitiveFindsFewer;

    // I42-I43: Export ueber einen beliebigen Filter (B28, Q4, Q5)
    [Test]
    [MaxTime(CMaxTime)]
    procedure SaveCopyAsWithFilterWritesDocx;
    [Test]
    [MaxTime(CMaxTime)]
    procedure SaveCopyAsWithUnknownFilterRaises;

    // I44-I47: Bild an einer Textmarke (B29, B30, B33, Q6)
    [Test]
    [MaxTime(CMaxTime)]
    procedure InsertImageUsesNaturalSize;
    [Test]
    [MaxTime(CMaxTime)]
    procedure InsertImageWithExplicitSize;
    [Test]
    [MaxTime(CMaxTime)]
    procedure InsertImageMissingFileRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure InsertImageUnknownBookmarkRaises;

    // I56: das Bild sitzt AN der Textmarke, nicht am Absatzende (B38, Klickprobe 2026-09-22)
    [Test]
    [MaxTime(CMaxTime)]
    procedure InsertImageSitsAtBookmark;

    // I48-I52: Lesen statt nur schreiben (B19, B31, Q7)
    [Test]
    [MaxTime(CMaxTime)]
    procedure GetCellReturnsValue;
    [Test]
    [MaxTime(CMaxTime)]
    procedure GetCellUnknownRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ReadAllReturnsGrid;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ReadPointBookmarkRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ReadSpanningBookmarkReturnsText;

    // I53-I54: verlorenes Dokument auf dem Weg ueber eine geliehene Tabelle (Review 2026-09-20, B35)
    [Test]
    [MaxTime(CMaxTime)]
    procedure ClosedDocumentIsReportedFromTableCell;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ClosedDocumentIsReportedFromTableFill;
    [Test]
    [MaxTime(CMaxTime)]
    procedure StaleTableReportsItself;
  end;

implementation

uses
  System.Classes,
  System.StrUtils,
  System.SysUtils,
  System.Variants,
  OOTable,
  OOTools,
  Test.Support;

type
  // Zugriff auf geschützte Member (FDocument, MakePropertyValue, ApplyPrinter, PrintArgs) – nur für Prüfungen
  TOOWriterAccess = class(TOOWriter);

{ ===== Lebenszyklus ===== }

procedure TOOWriterTests.SetupFixture;
begin
  TTestEnvironment.Start;
end;

procedure TOOWriterTests.TearDownFixture;
begin
  TTestEnvironment.Finish;
end;

procedure TOOWriterTests.Setup;
begin
  FWriter := TOOWriter.Create;
end;

procedure TOOWriterTests.TearDown;
begin
  FreeAndNil(FWriter);
  TTestEnvironment.CloseOwnDocuments;
end;

{ ===== Helfer ===== }

function TOOWriterTests.LoadCopy(const ASource, ATarget: string; AInSpecialDir: Boolean): string;
begin
  Result := TTestEnvironment.CopyTestFile(ASource, ATarget, AInSpecialDir);
  FWriter.LoadFile(Result, True);
end;

function TOOWriterTests.DocumentText: string;
begin
  // Die Textmarken in test.odt sind Punktmarken; ihr Anker bleibt leer, der Text steht daneben
  Result := TOOWriterAccess(FWriter).FDocument.getText.getString;
end;

function TOOWriterTests.CellText(const ATable, ACell: string): string;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getTextTables.getByName(ATable).getCellByName(ACell).getString;
end;

function TOOWriterTests.RowCount(const ATable: string): Integer;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getTextTables.getByName(ATable).getRows.getCount;
end;

function TOOWriterTests.PrinterSetting(const AName: string): OleVariant;
begin
  Result := FindProperty(TOOWriterAccess(FWriter).FDocument.getPrinter, AName);
end;

{ ===== Laden ===== }

procedure TOOWriterTests.LoadHidden;
var
  fileName: string;
  window: OleVariant;
begin
  fileName := LoadCopy('test.odt', 'I01_Laden.odt');
  Assert.IsTrue(TTestEnvironment.IsDocumentOpen(fileName), 'Dokument ist nicht auf dem Desktop');
  Assert.IsTrue(FWriter.IsLoaded, 'IsLoaded');
  // Hidden ist angekommen, wenn das Container-Fenster unsichtbar ist (B8)
  window := TOOWriterAccess(FWriter).FDocument.getCurrentController.getFrame.getContainerWindow;
  Assert.IsFalse(Boolean(window.isVisible), 'Dokument ist sichtbar');
end;

procedure TOOWriterTests.LoadFromSpecialPath;
var
  fileName: string;
begin
  fileName := LoadCopy('test.odt', 'I02_Sonderzeichen.odt', True);
  Assert.IsTrue(TTestEnvironment.IsDocumentOpen(fileName), 'Dokument ist nicht auf dem Desktop');
end;

procedure TOOWriterTests.LoadMissingFileRaises;
var
  fileName: string;
begin
  fileName := TTestEnvironment.TempDir + 'I18_Fehlt.odt';
  AssertRaisesOO(
    procedure
    begin
      FWriter.LoadFile(fileName, True);
    end,
    ['I18_Fehlt.odt']);
  Assert.IsFalse(FWriter.IsLoaded, 'IsLoaded nach gescheitertem Laden');
end;

procedure TOOWriterTests.NoDocumentRaises;
begin
  AssertRaisesOO(
    procedure
    begin
      FWriter.Save;
    end,
    []);
end;

procedure TOOWriterTests.LoadNonTextDocumentRaises;
var
  fileName: string;
begin
  fileName := TTestEnvironment.CreateCalcFile('I20_Tabelle.ods');
  AssertRaisesOO(
    procedure
    begin
      FWriter.LoadFile(fileName, True);
    end,
    ['I20_Tabelle.ods']);
  Assert.IsFalse(FWriter.IsLoaded, 'IsLoaded nach Ablehnung');
  Assert.IsFalse(TTestEnvironment.IsDocumentOpen(fileName), 'abgelehntes Dokument blieb offen');
end;

{ ===== Textmarken ===== }

procedure TOOWriterTests.WriteBookmark;
begin
  LoadCopy('test.odt', 'I03_Textmarke.odt');
  FWriter.WriteToBookmark('Test1', 'Hallo Welt');
  Assert.Contains(DocumentText, 'Hallo Welt');
end;

procedure TOOWriterTests.WriteBookmarkList;
var
  values: TStringList;
begin
  LoadCopy('test.odt', 'I04_Textmarken.odt');
  values := TStringList.Create;
  try
    values.Add('Test1=Alpha');
    values.Add('Test2=Beta');
    FWriter.WriteToBookmarks(values);
  finally
    values.Free;
  end;
  Assert.Contains(DocumentText, 'Alpha');
  Assert.Contains(DocumentText, 'Beta');
end;

procedure TOOWriterTests.WrongCaseBookmarkRaises;
var
  fileName: string;
begin
  // Textmarken sind case-sensitiv; LibreOffice meldet den Fehlgriff ohne Text (B5)
  fileName := LoadCopy('test.odt', 'I05_Schreibweise.odt');
  AssertRaisesOO(
    procedure
    begin
      FWriter.WriteToBookmark('test1', 'x');
    end,
    ['test1', ExtractFileName(fileName)]);
end;

procedure TOOWriterTests.BookmarkNames;
var
  names: TStringList;
begin
  LoadCopy('test.odt', 'I06_Namen.odt');
  names := TStringList.Create;
  try
    FWriter.GetBookmarkNames(names);
    names.Sort;
    Assert.AreEqual('Test1,Test2', names.CommaText);
  finally
    names.Free;
  end;
end;

procedure TOOWriterTests.WriteBookmarksChecksAllFirst;
var
  values: TStringList;
begin
  LoadCopy('test.odt', 'I25_Liste.odt');
  values := TStringList.Create;
  try
    values.Add('Test1=Zuerst');
    values.Add('Gibtsnicht=x');
    AssertRaisesOO(
      procedure
      begin
        FWriter.WriteToBookmarks(values);
      end,
      ['Gibtsnicht']);
  finally
    values.Free;
  end;
  Assert.DoesNotContain(DocumentText, 'Zuerst', 'Test1 wurde trotz Fehler geschrieben');
end;

{ ===== Tabellen ===== }

procedure TOOWriterTests.TableNamesAndCells;
var
  names: TStringList;
  values: TStringList;
begin
  LoadCopy('Tabellentest.odt', 'I07_Tabelle.odt');
  names := TStringList.Create;
  values := TStringList.Create;
  try
    FWriter.GetTableNames(names);
    names.Sort;
    Assert.AreEqual('Tabelle1,Tabelle2', names.CommaText);
    values.Add('A1=Hallo');
    values.Add('A2=Welt');
    FWriter.TableByName('Tabelle1').SetCells(values);
  finally
    values.Free;
    names.Free;
  end;
  Assert.AreEqual('Hallo', CellText('Tabelle1', 'A1'));
  Assert.AreEqual('Welt', CellText('Tabelle1', 'A2'));
end;

procedure TOOWriterTests.UnknownCellRaises;
var
  table: TOOTable;
begin
  LoadCopy('Tabellentest.odt', 'I08_Zelle.odt');
  table := FWriter.TableByName('Tabelle1');  // geliehen (Besitz: FWriter)
  AssertRaisesOO(
    procedure
    begin
      table.SetCell('Z99', 'x');
    end,
    ['Z99', 'Tabelle1']);
end;

procedure TOOWriterTests.TableLookup;
begin
  LoadCopy('Tabellentest.odt', 'I26_Suche.odt');
  Assert.IsNull(FWriter.FindTable('Gibtsnicht'), 'FindTable liefert eine Tabelle');
  Assert.IsNotNull(FWriter.FindTable('Tabelle2'), 'FindTable findet Tabelle2 nicht');
  AssertRaisesOO(
    procedure
    begin
      FWriter.TableByName('Gibtsnicht');
    end,
    ['Gibtsnicht']);
end;

procedure TOOWriterTests.SetCellsChecksAllFirst;
var
  values: TStringList;
  table: TOOTable;
begin
  LoadCopy('Tabellentest.odt', 'I27_Zellen.odt');
  table := FWriter.TableByName('Tabelle1');  // geliehen (Besitz: FWriter)
  values := TStringList.Create;
  try
    values.Add('A1=Zuerst');
    values.Add('Z99=x');
    AssertRaisesOO(
      procedure
      begin
        table.SetCells(values);
      end,
      ['Z99', 'Tabelle1']);
  finally
    values.Free;
  end;
  Assert.AreNotEqual('Zuerst', CellText('Tabelle1', 'A1'), 'A1 wurde trotz Fehler geschrieben');
end;

procedure TOOWriterTests.InsertRowsInMiddleAndAtEnd;
var
  table: TOOTable;
begin
  LoadCopy('Tabellentest.odt', 'I28_Zeilen.odt');
  table := FWriter.TableByName('Tabelle1');  // geliehen (Besitz: FWriter); 3 Zeilen
  table.SetCell('A2', 'Zwei');
  // Wie älterer Reportcode: InsertByIndex(1, n) lässt Platz nach der Kopfzeile
  table.InsertRows(1, 2);
  Assert.AreEqual(5, RowCount('Tabelle1'), 'Zeilen nach InsertRows(1, 2)');
  Assert.AreEqual('Zwei', CellText('Tabelle1', 'A4'), 'alte Zeile 2 ist nicht nach unten gerückt');
  table.InsertRows(5, 1);
  Assert.AreEqual(6, RowCount('Tabelle1'), 'Zeilen nach Anhängen');
  AssertRaisesOO(
    procedure
    begin
      table.InsertRows(99, 1);
    end,
    ['99', 'Tabelle1']);
end;

procedure TOOWriterTests.FillAppendsMissingRows;
var
  table: TOOTable;
  data: TArray<TArray<string>>;
begin
  LoadCopy('Tabellentest.odt', 'I29_Fuellen.odt');
  table := FWriter.TableByName('Tabelle1');  // geliehen (Besitz: FWriter); 3 Zeilen, 3 Spalten
  data := [['a1', 'b1'], ['a2', 'b2'], ['a3', 'b3'], ['a4', 'b4']];
  // Ab Zeile 2 braucht es Zeilen 2 bis 5, zwei fehlen und kommen ans Ende (E4)
  table.Fill(data, 2);
  Assert.AreEqual(5, RowCount('Tabelle1'), 'Zeilenzahl');
  Assert.AreEqual('a1', CellText('Tabelle1', 'A2'));
  Assert.AreEqual('b4', CellText('Tabelle1', 'B5'));
end;

procedure TOOWriterTests.FillTooWideRaises;
var
  table: TOOTable;
  data: TArray<TArray<string>>;
begin
  LoadCopy('Tabellentest.odt', 'I30_Breit.odt');
  table := FWriter.TableByName('Tabelle1');  // geliehen (Besitz: FWriter); 3 Spalten
  data := [['a1', 'b1', 'c1'], ['a2', 'b2', 'c2', 'd2']];
  AssertRaisesOO(
    procedure
    begin
      table.Fill(data, 1);
    end,
    ['Tabelle1']);
  Assert.AreNotEqual('a1', CellText('Tabelle1', 'A1'), 'Zeile 1 wurde trotz Fehler geschrieben');
  Assert.AreEqual(3, RowCount('Tabelle1'), 'Zeilen wurden trotz Fehler eingefügt');
end;

{ ===== Speichern und Export ===== }

procedure TOOWriterTests.SaveResetsModified;
var
  fileName: string;
begin
  fileName := LoadCopy('test.odt', 'I09_Speichern.odt');
  FWriter.WriteToBookmark('Test1', 'Gespeichert');
  FWriter.Save;
  Assert.IsTrue(TTestEnvironment.OdtContains(fileName, 'Gespeichert'), 'Änderung fehlt in der Datei');
  // store() setzt das Dokument auf „unverändert“, storeToURL() nicht
  Assert.IsFalse(Boolean(TOOWriterAccess(FWriter).FDocument.isModified), 'Dokument gilt nach Save als geändert');
end;

procedure TOOWriterTests.SaveAsRenamesDocument;
var
  target: string;
begin
  LoadCopy('test.odt', 'I10_Quelle.odt');
  target := TTestEnvironment.TempDir + 'I10_Ziel.odt';
  FWriter.SaveAs(target);
  Assert.IsTrue(FileExists(target), 'Zieldatei fehlt');
  Assert.AreEqual(target, FWriter.FileName);
  Assert.IsTrue(EndsText('/I10_Ziel.odt', TOOWriterAccess(FWriter).FDocument.getURL),
    'Dokument heißt nicht wie das Ziel');
end;

procedure TOOWriterTests.SaveAsToSpecialPath;
var
  target: string;
begin
  LoadCopy('test.odt', 'I11_Quelle.odt');
  target := TTestEnvironment.SpecialDir + 'I11_Ziel.odt';
  FWriter.SaveAs(target);
  Assert.IsTrue(FileExists(target), 'Zieldatei fehlt');
end;

procedure TOOWriterTests.SaveCopyAsKeepsName;
var
  fileName: string;
  target: string;
begin
  fileName := LoadCopy('test.odt', 'I12_Quelle.odt');
  target := TTestEnvironment.TempDir + 'I12_Kopie.odt';
  FWriter.SaveCopyAs(target);
  Assert.IsTrue(FileExists(target), 'Kopie fehlt');
  Assert.AreEqual(fileName, FWriter.FileName);
  Assert.IsTrue(EndsText('/I12_Quelle.odt', TOOWriterAccess(FWriter).FDocument.getURL),
    'Dokument hat seinen Namen gewechselt');
end;

procedure TOOWriterTests.ExportPdfWritesPdf;
var
  fileName: string;
  target: string;
begin
  fileName := LoadCopy('test.odt', 'I23_Quelle.odt');
  target := TTestEnvironment.SpecialDir + 'I23_Export.pdf';
  FWriter.ExportPdf(target);
  Assert.IsTrue(FileExists(target), 'PDF fehlt');
  Assert.AreEqual('%PDF', FileHead(target, 4), 'Datei ist kein PDF');
  Assert.AreEqual(fileName, FWriter.FileName, 'Export hat das Dokument umbenannt');
end;

procedure TOOWriterTests.SaveCopyAsOverwrites;
var
  target: string;
begin
  LoadCopy('test.odt', 'I24_Quelle.odt');
  // Das Ziel gibt es schon, mit anderem Inhalt
  target := TTestEnvironment.CopyTestFile('Tabellentest.odt', 'I24_Ziel.odt');
  FWriter.WriteToBookmark('Test1', 'Ueberschrieben');
  FWriter.SaveCopyAs(target);
  Assert.IsTrue(TTestEnvironment.OdtContains(target, 'Ueberschrieben'), 'Ziel wurde nicht überschrieben');
end;

{ ===== Schließen, Freigeben, Übergabe ===== }

procedure TOOWriterTests.CloseWithoutSaving;
var
  fileName: string;
begin
  fileName := LoadCopy('test.odt', 'I13_Schliessen.odt');
  FWriter.WriteToBookmark('Test1', 'Verworfen');
  FWriter.CloseFile(False);
  Assert.IsFalse(FWriter.IsLoaded, 'IsLoaded nach CloseFile');
  Assert.IsFalse(TTestEnvironment.IsDocumentOpen(fileName), 'Dokument ist noch offen');
  Assert.IsFalse(TTestEnvironment.OdtContains(fileName, 'Verworfen'), 'Änderung wurde gespeichert');
end;

procedure TOOWriterTests.CloseWithSaving;
var
  fileName: string;
begin
  fileName := LoadCopy('test.odt', 'I14_Schliessen.odt');
  FWriter.WriteToBookmark('Test1', 'Beim Schließen');
  FWriter.CloseFile(True);
  Assert.IsFalse(TTestEnvironment.IsDocumentOpen(fileName), 'Dokument ist noch offen');
  Assert.IsTrue(TTestEnvironment.OdtContains(fileName, 'Beim Schließen'), 'Änderung fehlt in der Datei');
end;

procedure TOOWriterTests.FreeClosesDocument;
var
  fileName: string;
begin
  fileName := LoadCopy('test.odt', 'I15_Freigeben.odt');
  FreeAndNil(FWriter);
  Assert.IsFalse(TTestEnvironment.IsDocumentOpen(fileName), 'Dokument blieb nach der Freigabe offen');
end;

procedure TOOWriterTests.PropertyValuesAreIndependent;
var
  first: OleVariant;
  second: OleVariant;
begin
  // Ein gemeinsames Objekt für alle Einträge liefert nur den letzten Stand (B6, alter Druck-Fehler)
  first := TOOWriterAccess(FWriter).MakePropertyValue('Hidden', True);
  second := TOOWriterAccess(FWriter).MakePropertyValue('ReadOnly', False);
  Assert.AreEqual('Hidden', string(first.Name), 'erstes PropertyValue wurde überschrieben');
  Assert.IsTrue(Boolean(first.Value), 'Wert des ersten PropertyValue wurde überschrieben');
  Assert.AreEqual('ReadOnly', string(second.Name));
end;

procedure TOOWriterTests.LoadTwiceRaises;
var
  firstFile: string;
  secondFile: string;
begin
  // Ein Objekt hält genau ein Dokument; wer mehrere braucht, nimmt mehrere Objekte (Q1)
  firstFile := LoadCopy('test.odt', 'I17_Erstes.odt');
  secondFile := TTestEnvironment.CopyTestFile('test.odt', 'I17_Zweites.odt');
  AssertRaisesOO(
    procedure
    begin
      FWriter.LoadFile(secondFile, True);
    end,
    [ExtractFileName(firstFile)]);
  Assert.IsFalse(TTestEnvironment.IsDocumentOpen(secondFile), 'zweites Dokument wurde trotzdem geladen');
end;

procedure TOOWriterTests.HandOverLeavesDocumentOpen;
var
  fileName: string;
  window: OleVariant;
begin
  fileName := LoadCopy('test.odt', 'I21_Uebergabe.odt');
  // Eigener Verweis auf das Fenster, um es nach der Übergabe noch prüfen zu können
  window := TOOWriterAccess(FWriter).FDocument.getCurrentController.getFrame.getContainerWindow;
  FWriter.HandOver;
  Assert.IsFalse(FWriter.IsLoaded, 'Objekt hält das Dokument nach HandOver noch');
  FreeAndNil(FWriter);
  Assert.IsTrue(TTestEnvironment.IsDocumentOpen(fileName), 'Freigeben hat das übergebene Dokument geschlossen');
  Assert.IsTrue(Boolean(window.isVisible), 'übergebenes Dokument ist nicht sichtbar');
end;

procedure TOOWriterTests.VisibleFollowsContainerWindow;
begin
  LoadCopy('test.odt', 'I22_Sichtbar.odt');
  Assert.IsFalse(FWriter.Visible, 'versteckt geladen, aber Visible');
  FWriter.Visible := True;
  Assert.IsTrue(FWriter.Visible, 'nach Visible := True nicht sichtbar');
  FWriter.Visible := False;
  Assert.IsFalse(FWriter.Visible, 'nach Visible := False sichtbar');
end;

{ ===== Drucker und Druckargumente ===== }

procedure TOOWriterTests.UnknownPrinterRaises;
var
  options: TOOPrintOptions;
begin
  LoadCopy('test.odt', 'I31_Drucker.odt');
  options := TOOPrintOptions.Default;
  options.PrinterName := 'Gibt es nicht 4711';
  AssertRaisesOO(
    procedure
    begin
      TOOWriterAccess(FWriter).ApplyPrinter(options);
    end,
    ['Gibt es nicht 4711']);
end;

procedure TOOWriterTests.PrintArgsAreSeparateAndTyped;
var
  options: TOOPrintOptions;
  args: OleVariant;
begin
  options := TOOPrintOptions.Default;
  options.Copies := 3;
  options.Collate := False;
  options.Pages := '2-3';
  args := TOOWriterAccess(FWriter).PrintArgs(options);
  Assert.AreEqual(3, Integer(FindProperty(args, 'CopyCount')), 'CopyCount');
  // CopyCount ist in UNO ein short
  Assert.AreEqual(Integer(varSmallint), Integer(VarType(FindProperty(args, 'CopyCount'))), 'Typ von CopyCount');
  Assert.IsFalse(Boolean(FindProperty(args, 'Collate')), 'Collate');
  Assert.AreEqual('2-3', string(FindProperty(args, 'Pages')), 'Pages');
  Assert.IsTrue(Boolean(FindProperty(args, 'Wait')), 'Wait (A8)');
  // Ohne Seitenangabe druckt LibreOffice alles; Pages fehlt dann ganz
  args := TOOWriterAccess(FWriter).PrintArgs(TOOPrintOptions.Default);
  Assert.IsFalse(HasProperty(args, 'Pages'), 'Pages trotz leerer Angabe');
  Assert.IsTrue(Boolean(FindProperty(args, 'Collate')), 'Collate-Vorgabe (E3)');
end;

procedure TOOWriterTests.ZeroCopiesRaises;
var
  options: TOOPrintOptions;
begin
  options := TOOPrintOptions.Default;
  options.Copies := 0;
  AssertRaisesOO(
    procedure
    begin
      TOOWriterAccess(FWriter).PrintArgs(options);
    end,
    []);
end;

procedure TOOWriterTests.PrintSettingsLeaveDocumentUntouched;
var
  options: TOOPrintOptions;
  pageStyle: OleVariant;
  widthBefore: Integer;
  heightBefore: Integer;
begin
  LoadCopy('test.odt', 'I36_Seite.odt');
  pageStyle := TOOWriterAccess(FWriter).FDocument.getStyleFamilies.getByName('PageStyles').getByName(
    TOOWriterAccess(FWriter).FDocument.getText.createTextCursor.PageStyleName);
  widthBefore := pageStyle.Width;
  heightBefore := pageStyle.Height;
  options := TOOPrintOptions.Default;
  options.PrinterName := DefaultPrinterName;
  TOOWriterAccess(FWriter).ApplyPrinter(options);
  Assert.AreEqual(DefaultPrinterName, string(PrinterSetting('Name')), 'Drucker nicht gesetzt');
  // Bis E5 setzte die Bibliothek auf Wunsch auch das Papierformat – das formatierte das Dokument um und machte
  // es „geändert“ (B26; dieser Test war damit rot)
  Assert.AreEqual(widthBefore, Integer(pageStyle.Width), 'Seitenbreite verändert');
  Assert.AreEqual(heightBefore, Integer(pageStyle.Height), 'Seitenhöhe verändert');
  Assert.IsFalse(Boolean(TOOWriterAccess(FWriter).FDocument.isModified), 'Dokument gilt als geändert');
end;

procedure TOOWriterTests.ClosedDocumentIsReported;
var
  fileName: string;
begin
  fileName := LoadCopy('test.odt', 'I37_Verloren.odt');
  // Hinter dem Rücken des Objekts schließen – so wie ein Benutzer das Fenster in LibreOffice schließt
  TOOWriterAccess(FWriter).FDocument.close(True);
  AssertRaisesOO(
    procedure
    begin
      FWriter.Save;
    end,
    [ExtractFileName(fileName), 'geschlossen']);
  Assert.IsFalse(FWriter.IsLoaded, 'Objekt hält das verlorene Dokument noch');
end;

{ ===== Helfer fuer die P6-Tests ===== }

function TOOWriterTests.HeaderString: string;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getStyleFamilies.getByName('PageStyles')
    .getByName('Standard').HeaderText.getString;
end;

function TOOWriterTests.FooterString: string;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getStyleFamilies.getByName('PageStyles')
    .getByName('Standard').FooterText.getString;
end;

function TOOWriterTests.FrameString: string;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getTextFrames.getByIndex(0).getText.getString;
end;

function TOOWriterTests.GraphicCount: Integer;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getGraphicObjects.getCount;
end;

function TOOWriterTests.GraphicWidth(AIndex: Integer): Integer;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getGraphicObjects.getByIndex(AIndex).Width;
end;

function TOOWriterTests.GraphicHeight(AIndex: Integer): Integer;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getGraphicObjects.getByIndex(AIndex).Height;
end;

function TOOWriterTests.GraphicAnchor(AIndex: Integer): Integer;
begin
  Result := TOOWriterAccess(FWriter).FDocument.getGraphicObjects.getByIndex(AIndex).AnchorType;
end;

function TOOWriterTests.ParagraphLayout: string;
var
  paragraphs: OleVariant;
  portions: OleVariant;
  portion: OleVariant;
begin
  // Der erste Absatz als Zeichenkette, das Bild als [BILD]: damit wird die POSITION pruefbar und nicht
  // nur die Anzahl. Textmarken sind eigene Abschnitte ohne Text und fallen dabei heraus.
  Result := '';
  paragraphs := TOOWriterAccess(FWriter).FDocument.getText.createEnumeration;
  portions := paragraphs.nextElement.createEnumeration;
  while Boolean(portions.hasMoreElements) do
  begin
    portion := portions.nextElement;
    if VarToStr(portion.TextPortionType) = 'Frame' then
    begin
      Result := Result + '[BILD]';
    end
    else
    begin
      Result := Result + VarToStr(portion.getString);
    end;
  end;
end;

{ ===== I38-I41: Suchen und Ersetzen ===== }

procedure TOOWriterTests.ReplaceAllReturnsCount;
begin
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I38_Marken.odt'), True);
  // MARKE steht in Fliesstext, Kopf, Fuss und Rahmen (B32)
  Assert.AreEqual(4, FWriter.ReplaceAll('MARKE', 'ERSETZT'), 'Anzahl der Ersetzungen');
  Assert.IsTrue(ContainsText(DocumentText, 'ERSETZT im Fliesstext'), 'Fliesstext nicht ersetzt');
end;

procedure TOOWriterTests.ReplaceAllWithoutHitsReturnsZero;
begin
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I39_Marken.odt'), True);
  // Kein Treffer ist kein Fehler, nur eine 0 (Q2)
  Assert.AreEqual(0, FWriter.ReplaceAll('GIBTESNICHT', 'X'), 'ohne Treffer muss 0 herauskommen');
  Assert.IsTrue(ContainsText(DocumentText, 'MARKE im Fliesstext'), 'Dokument wurde doch veraendert');
end;

procedure TOOWriterTests.ReplaceAllCoversHeaderFooterAndFrame;
begin
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I40_Marken.odt'), True);
  FWriter.ReplaceAll('MARKE', 'ERSETZT');
  Assert.AreEqual('ERSETZT im Kopf', HeaderString, 'Kopfzeile');
  Assert.AreEqual('ERSETZT im Fuss', FooterString, 'Fusszeile');
  Assert.AreEqual('ERSETZT im Rahmen', FrameString, 'Textrahmen');
end;

procedure TOOWriterTests.ReplaceAllCaseSensitiveFindsFewer;
var
  options: TOOSearchOptions;
begin
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I41_Marken.odt'), True);
  options := TOOSearchOptions.Default;
  options.CaseSensitive := True;
  Assert.AreEqual(0, FWriter.ReplaceAll('marke', 'ERSETZT', options), 'klein geschrieben darf nicht treffen');
  Assert.AreEqual(4, FWriter.ReplaceAll('MARKE', 'ERSETZT', options), 'gross geschrieben muss treffen');
end;

{ ===== I42-I43: Export ueber einen beliebigen Filter ===== }

procedure TOOWriterTests.SaveCopyAsWithFilterWritesDocx;
var
  source: string;
  target: string;
begin
  source := LoadCopy('test.odt', 'I42_Quelle.odt');
  target := TTestEnvironment.TempDir + 'I42_Kopie.docx';
  FWriter.SaveCopyAs(target, OOFilterDocx);
  Assert.IsTrue(FileExists(target), 'DOCX fehlt');
  Assert.AreEqual('PK', FileHead(target, 2), 'Datei ist kein ZIP-Container');
  Assert.AreEqual(source, FWriter.FileName, 'die Kopie hat das Dokument umbenannt');
end;

procedure TOOWriterTests.SaveCopyAsWithUnknownFilterRaises;
var
  target: string;
begin
  LoadCopy('test.odt', 'I43_Quelle.odt');
  target := TTestEnvironment.TempDir + 'I43_Murks.docx';
  // B28 meldet nur 'Error Area:Io Class:Parameter Code:26' - die Meldung muss den Filter nennen (Q5)
  AssertRaisesOO(
    procedure
    begin
      FWriter.SaveCopyAs(target, 'Gibt_Es_Nicht_Filter');
    end,
    ['Gibt_Es_Nicht_Filter']);
  Assert.IsFalse(FileExists(target), 'trotz Fehler wurde eine Datei geschrieben');
end;

{ ===== I44-I47: Bild an einer Textmarke ===== }

procedure TOOWriterTests.InsertImageUsesNaturalSize;
var
  logo: string;
begin
  logo := TTestEnvironment.CopyTestFile('logo.png', 'I44_logo.png');
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I44_Marken.odt'), True);
  FWriter.InsertImageAtBookmark('Punkt', logo);
  Assert.AreEqual(1, GraphicCount, 'Bild fehlt');
  // logo.png ist 40 x 20 Pixel, das Seitenverhaeltnis 2:1 muss stehen (B33)
  Assert.IsTrue(Abs(GraphicWidth(0) - 2 * GraphicHeight(0)) <= 2,
    Format('Seitenverhaeltnis verloren: %d x %d', [GraphicWidth(0), GraphicHeight(0)]));
  // AS_CHARACTER = 1: das Bild sitzt im Textfluss, nicht am Absatz (LibreOffice-Vorgabe waere 0)
  Assert.AreEqual(1, GraphicAnchor(0), 'Verankerung');
end;

procedure TOOWriterTests.InsertImageWithExplicitSize;
var
  logo: string;
begin
  logo := TTestEnvironment.CopyTestFile('logo.png', 'I45_logo.png');
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I45_Marken.odt'), True);
  FWriter.InsertImageAtBookmark('Punkt', logo, 2000, 1000);
  Assert.AreEqual(2000, GraphicWidth(0), 'Breite in 1/100 mm');
  Assert.AreEqual(1000, GraphicHeight(0), 'Hoehe in 1/100 mm');
end;

procedure TOOWriterTests.InsertImageMissingFileRaises;
var
  fehlt: string;
begin
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I46_Marken.odt'), True);
  fehlt := TTestEnvironment.TempDir + 'I46_gibtsnicht.png';
  // B29: queryGraphic liefert dafuer still null - das darf nicht durchrutschen
  AssertRaisesOO(
    procedure
    begin
      FWriter.InsertImageAtBookmark('Punkt', fehlt);
    end,
    ['I46_gibtsnicht.png']);
  Assert.AreEqual(0, GraphicCount, 'trotz Fehler wurde ein Bild eingefuegt');
end;

procedure TOOWriterTests.InsertImageUnknownBookmarkRaises;
var
  logo: string;
begin
  logo := TTestEnvironment.CopyTestFile('logo.png', 'I47_logo.png');
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I47_Marken.odt'), True);
  AssertRaisesOO(
    procedure
    begin
      FWriter.InsertImageAtBookmark('GibtEsNicht', logo);
    end,
    ['GibtEsNicht', 'Punkt']);
  Assert.AreEqual(0, GraphicCount, 'trotz Fehler wurde ein Bild eingefuegt');
end;

{ ===== I56: Bild an der Textmarke, nicht am Absatzende ===== }

procedure TOOWriterTests.InsertImageSitsAtBookmark;
var
  logo: string;
begin
  logo := TTestEnvironment.CopyTestFile('logo.png', 'I56_logo.png');
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I56_Marken.odt'), True);
  FWriter.InsertImageAtBookmark('Punkt', logo);
  // Die Marke Punkt steht zwischen "MARKE im Fliesstext " und "SPANNE" - genau dort gehoert das Bild hin.
  // Wird AnchorType erst NACH dem Einfuegen gesetzt, haengt LibreOffice es ans Absatzende um (B38).
  Assert.AreEqual('MARKE im Fliesstext [BILD]SPANNE', ParagraphLayout);
end;

{ ===== I48-I52: Lesen statt nur schreiben ===== }

procedure TOOWriterTests.GetCellReturnsValue;
var
  table: TOOTable;
begin
  LoadCopy('Tabellentest.odt', 'I48_Tabelle.odt');
  table := FWriter.TableByName('Tabelle1');
  table.SetCell('A1', 'WERT 42');
  Assert.AreEqual('WERT 42', table.GetCell('A1'), 'Zelle zurueckgelesen');
end;

procedure TOOWriterTests.GetCellUnknownRaises;
var
  table: TOOTable;
begin
  LoadCopy('Tabellentest.odt', 'I49_Tabelle.odt');
  table := FWriter.TableByName('Tabelle1');
  // B19 gilt auch lesend: getCellByName liefert null statt einer Ausnahme
  AssertRaisesOO(
    procedure
    begin
      table.GetCell('ZZ99');
    end,
    ['ZZ99', 'Tabelle1']);
end;

procedure TOOWriterTests.ReadAllReturnsGrid;
var
  table: TOOTable;
  grid: TArray<TArray<string>>;
begin
  LoadCopy('Tabellentest.odt', 'I50_Tabelle.odt');
  table := FWriter.TableByName('Tabelle1');
  table.SetCell('A1', 'oben links');
  grid := table.ReadAll;
  // Length liefert unter Win64 ein NativeInt, deshalb der Cast
  Assert.AreEqual(RowCount('Tabelle1'), Integer(Length(grid)), 'Zeilenzahl');
  Assert.IsTrue(Length(grid) > 0, 'Tabelle kam leer zurueck');
  Assert.AreEqual('oben links', grid[0][0], 'erste Zelle');
end;

procedure TOOWriterTests.ReadPointBookmarkRaises;
begin
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I51_Marken.odt'), True);
  FWriter.WriteToBookmark('Punkt', 'HALLO');
  // B31: getAnchor.getString liefert hier '' - auch nach dem Schreiben. Ein leerer String waere
  // nicht von einem leeren Feld zu unterscheiden, deshalb bricht das Lesen laut ab (Q7).
  AssertRaisesOO(
    procedure
    begin
      FWriter.ReadBookmark('Punkt');
    end,
    ['Punkt']);
end;

procedure TOOWriterTests.ReadSpanningBookmarkReturnsText;
begin
  FWriter.LoadFile(TTestEnvironment.CreateMarkedDocument('I52_Marken.odt'), True);
  Assert.AreEqual('SPANNE', FWriter.ReadBookmark('Spanne'), 'umspannende Textmarke');
end;

{ ===== I53-I54: verlorenes Dokument ueber eine geliehene Tabelle ===== }

procedure TOOWriterTests.ClosedDocumentIsReportedFromTableCell;
var
  table: TOOTable;
begin
  LoadCopy('Tabellentest.odt', 'I53_Verloren.odt');
  table := FWriter.TableByName('Tabelle1');
  // Hinter dem Ruecken des Objekts schliessen - wie ein Benutzer, der das Fenster zumacht
  TOOWriterAccess(FWriter).FDocument.close(True);
  AssertRaisesOO(
    procedure
    begin
      table.SetCell('A1', 'x');
    end,
    ['Tabelle1', 'geschlossen']);
  Assert.IsFalse(FWriter.IsLoaded, 'Objekt haelt das verlorene Dokument noch');
end;

procedure TOOWriterTests.ClosedDocumentIsReportedFromTableFill;
var
  table: TOOTable;
begin
  LoadCopy('Tabellentest.odt', 'I54_Verloren.odt');
  table := FWriter.TableByName('Tabelle1');
  TOOWriterAccess(FWriter).FDocument.close(True);
  // Fill ruft RowCount vor seinem try-Block; frueher kam hier ein roher EOleSysError durch
  AssertRaisesOO(
    procedure
    begin
      table.Fill([['x']]);
    end,
    ['Tabelle1', 'geschlossen']);
  Assert.IsFalse(FWriter.IsLoaded, 'Objekt haelt das verlorene Dokument noch');
end;

procedure TOOWriterTests.StaleTableReportsItself;
var
  table: TOOTable;
begin
  LoadCopy('Tabellentest.odt', 'I55_Veraltet.odt');
  table := FWriter.TableByName('Tabelle1');
  FWriter.CloseFile(False);
  // Der Zeiger bleibt gueltig, das Objekt lebt bis zum naechsten Laden - es muss jetzt selbst sagen,
  // dass sein Dokument weg ist, statt mit einem toten UNO-Objekt zu reden
  AssertRaisesOO(
    procedure
    begin
      table.SetCell('A1', 'x');
    end,
    ['Tabelle1', 'nicht mehr geladen']);
  AssertRaisesOO(
    procedure
    begin
      table.Fill([['x']]);
    end,
    ['Tabelle1', 'nicht mehr geladen']);
end;

initialization
  TDUnitX.RegisterTestFixture(TOOWriterTests);

end.
