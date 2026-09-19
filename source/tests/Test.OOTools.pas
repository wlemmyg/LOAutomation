{
--------------------------------------------------------------------------------
Copyright (c) 2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}
unit Test.OOTools;

interface

uses
  DUnitX.TestFramework;

type
  [TestFixture]
  TOOToolsTests = class
  public
    // Der Ordinalwert geht per OLE unverändert als com.sun.star.view.PaperFormat an LibreOffice (A11)
    [Test]
    procedure PaperFormatOrdinalsMatchUno;

    // U1: Normalform file:///C:/… (T1), ungeschützte Zeichen bleiben, / gilt wie \ (T2)
    [Test]
    procedure DrivePathKeepsColon;
    // U2
    [Test]
    procedure UncPathBecomesHost;
    // U3–U7: %-Kodierung, Nicht-ASCII als UTF-8 (B4, T2)
    [Test]
    procedure UmlautIsUtf8Encoded;
    [Test]
    procedure SpaceIsEncoded;
    [Test]
    procedure HashIsEncoded;
    [Test]
    procedure PercentIsEncoded;
    [Test]
    procedure ReservedCharsAreEncoded;

    // U8: relative Pfade laut ablehnen statt still auflösen (F4, T3)
    [TestCase('Datei', 'Brief.odt')]
    [TestCase('Ordner', 'Ordner\Brief.odt')]
    [TestCase('Laufwerksrelativ', '\Ordner\Brief.odt')]
    [TestCase('Verzeichnisrelativ', 'C:Brief.odt')]
    procedure RelativePathRaises(const APath: string);
    [Test]
    procedure EmptyPathRaises;

    // U10–U11: Druckvorgaben (E2, E3) und Druckerprüfung, ohne LibreOffice (Plan 7.2)
    [Test]
    procedure PrintOptionsDefault;
    [Test]
    procedure PrinterExistsKnowsDefaultPrinter;
  end;

implementation

uses
  OOTools,
  Test.Support;

{ ===== Papierformat ===== }

procedure TOOToolsTests.PaperFormatOrdinalsMatchUno;
begin
  // Reihenfolge laut IDL: A3, A4, A5, B4, B5, LETTER, LEGAL, TABLOID, USER
  Assert.AreEqual(0, Ord(pfA3), 'A3');
  Assert.AreEqual(1, Ord(pfA4), 'A4');
  Assert.AreEqual(2, Ord(pfA5), 'A5');
  Assert.AreEqual(3, Ord(pfB4), 'B4');
  Assert.AreEqual(4, Ord(pfB5), 'B5');
  Assert.AreEqual(5, Ord(pfLetter), 'LETTER');
  Assert.AreEqual(6, Ord(pfLegal), 'LEGAL');
  Assert.AreEqual(7, Ord(pfTabloid), 'TABLOID');
  Assert.AreEqual(8, Ord(pfUser), 'USER');
end;

{ ===== FileNameToUrl ===== }

procedure TOOToolsTests.DrivePathKeepsColon;
begin
  Assert.AreEqual('file:///C:/Temp/Brief.odt', FileNameToUrl('C:\Temp\Brief.odt'));
  Assert.AreEqual('file:///C:/Temp/Brief-1_a.b~c.odt', FileNameToUrl('C:\Temp\Brief-1_a.b~c.odt'));
  Assert.AreEqual('file:///C:/Temp/Brief.odt', FileNameToUrl('C:/Temp/Brief.odt'));
end;

procedure TOOToolsTests.UncPathBecomesHost;
begin
  Assert.AreEqual('file://server/freigabe/Brief.odt', FileNameToUrl('\\server\freigabe\Brief.odt'));
end;

procedure TOOToolsTests.UmlautIsUtf8Encoded;
begin
  // ü = C3 BC, ö = C3 B6, ß = C3 9F
  Assert.AreEqual('file:///C:/Pr%C3%BCfung/Gr%C3%B6%C3%9Fe.odt', FileNameToUrl('C:\Prüfung\Größe.odt'));
end;

procedure TOOToolsTests.SpaceIsEncoded;
begin
  Assert.AreEqual('file:///C:/Neuer%20Ordner/Brief.odt', FileNameToUrl('C:\Neuer Ordner\Brief.odt'));
end;

procedure TOOToolsTests.HashIsEncoded;
begin
  // Roh schnitte # die URL als Fragment ab (B4)
  Assert.AreEqual('file:///C:/Akte%2312/Brief.odt', FileNameToUrl('C:\Akte#12\Brief.odt'));
end;

procedure TOOToolsTests.PercentIsEncoded;
begin
  Assert.AreEqual('file:///C:/Rabatt50%25/Brief.odt', FileNameToUrl('C:\Rabatt50%\Brief.odt'));
end;

procedure TOOToolsTests.ReservedCharsAreEncoded;
begin
  Assert.AreEqual('file:///C:/x%21%24%26%27%28%29%2B%2C%3B%3D%40y.odt',
    FileNameToUrl('C:\x!$&''()+,;=@y.odt'));
end;

procedure TOOToolsTests.RelativePathRaises(const APath: string);
begin
  AssertRaisesOO(
    procedure
    begin
      FileNameToUrl(APath);
    end,
    [APath, 'ExpandFileName']);
end;

procedure TOOToolsTests.EmptyPathRaises;
begin
  Assert.WillRaise(
    procedure
    begin
      FileNameToUrl('');
    end,
    EOOAutomation);
end;

{ ===== Druck ===== }

procedure TOOToolsTests.PrintOptionsDefault;
var
  options: TOOPrintOptions;
begin
  options := TOOPrintOptions.Default;
  Assert.AreEqual('', options.PrinterName, 'PrinterName');
  Assert.AreEqual(1, options.Copies, 'Copies');
  Assert.IsTrue(options.Collate, 'Collate (E3)');
  Assert.AreEqual('', options.Pages, 'Pages');
  Assert.IsFalse(options.OverridePaper, 'OverridePaper (E2)');
end;

procedure TOOToolsTests.PrinterExistsKnowsDefaultPrinter;
begin
  // Der Standarddrucker kommt aus GetDefaultPrinter, PrinterExists zählt über EnumPrinters auf
  Assert.IsTrue(PrinterExists(DefaultPrinterName), 'Standarddrucker nicht gefunden');
  Assert.IsFalse(PrinterExists('Gibt es nicht 4711'), 'erfundener Drucker gefunden');
end;

initialization
  TDUnitX.RegisterTestFixture(TOOToolsTests);

end.
