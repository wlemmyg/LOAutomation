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

    // I31–I35: Drucker und Druckargumente
    [Test]
    [MaxTime(CMaxTime)]
    procedure UnknownPrinterRaises;
    [Test]
    [MaxTime(CMaxTime)]
    procedure PaperStaysWithoutOverride;
    [Test]
    [MaxTime(CMaxTime)]
    procedure OverridePaperSetsFormat;
    [Test]
    [MaxTime(CMaxTime)]
    procedure PrintArgsAreSeparateAndTyped;
    [Test]
    [MaxTime(CMaxTime)]
    procedure ZeroCopiesRaises;
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
  // Wie CORA heute mit InsertByIndex(1, n): Platz nach der Kopfzeile
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

procedure TOOWriterTests.PaperStaysWithoutOverride;
var
  options: TOOPrintOptions;
begin
  LoadCopy('test.odt', 'I32_Papier.odt');
  options := TOOPrintOptions.Default;
  options.OverridePaper := True;
  options.Orientation := ooLandscape;
  options.PaperFormat := pfA5;
  TOOWriterAccess(FWriter).ApplyPrinter(options);
  // Jetzt nur den Drucker setzen; Hochformat/A4 im Record dürfen ohne OverridePaper nicht ankommen (E2)
  options := TOOPrintOptions.Default;
  options.PrinterName := DefaultPrinterName;
  TOOWriterAccess(FWriter).ApplyPrinter(options);
  Assert.AreEqual(DefaultPrinterName, string(PrinterSetting('Name')), 'Drucker');
  Assert.AreEqual(Ord(ooLandscape), Integer(PrinterSetting('PaperOrientation')), 'Ausrichtung verändert');
  Assert.AreEqual(Ord(pfA5), Integer(PrinterSetting('PaperFormat')), 'Format verändert');
end;

procedure TOOWriterTests.OverridePaperSetsFormat;
var
  options: TOOPrintOptions;
begin
  LoadCopy('test.odt', 'I33_Papier.odt');
  options := TOOPrintOptions.Default;
  options.OverridePaper := True;
  options.Orientation := ooLandscape;
  options.PaperFormat := pfA5;
  TOOWriterAccess(FWriter).ApplyPrinter(options);
  Assert.AreEqual(Ord(ooLandscape), Integer(PrinterSetting('PaperOrientation')), 'Ausrichtung');
  Assert.AreEqual(Ord(pfA5), Integer(PrinterSetting('PaperFormat')), 'Format');
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

initialization
  TDUnitX.RegisterTestFixture(TOOWriterTests);

end.
