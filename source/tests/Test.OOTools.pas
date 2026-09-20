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

    // U10–U11: Druckvorgaben (E3) und Druckerprüfung, ohne LibreOffice (Plan 7.2)
    [Test]
    procedure PrintOptionsDefault;
    [Test]
    procedure PrinterExistsKnowsDefaultPrinter;

    // U12–U13: Klartext, wenn LibreOffice wegbricht (E6)
    [Test]
    procedure LostConnectionIsExplained;
    [Test]
    procedure OtherErrorsKeepTheirText;
  end;

implementation

uses
  System.Win.ComObj,
  OOTools,
  Test.Support;

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
end;

procedure TOOToolsTests.PrinterExistsKnowsDefaultPrinter;
begin
  // Der Standarddrucker kommt aus GetDefaultPrinter, PrinterExists zählt über EnumPrinters auf
  Assert.IsTrue(PrinterExists(DefaultPrinterName), 'Standarddrucker nicht gefunden');
  Assert.IsFalse(PrinterExists('Gibt es nicht 4711'), 'erfundener Drucker gefunden');
end;

{ ===== Fehlermeldungen ===== }

procedure TOOToolsTests.LostConnectionIsExplained;
var
  error: EOleSysError;
  wrapped: EOOAutomation;
begin
  // „Der RPC-Server ist nicht verfügbar“ sagt dem Leser nichts; die Meldung muss den Grund und den Weg nennen
  error := EOleSysError.Create('Der RPC-Server ist nicht verfügbar', HRESULT($800706BA), 0);
  try
    wrapped := WrapUnoError('Speichern von "C:\Temp\Brief.odt" fehlgeschlagen', error);
    try
      Assert.Contains(wrapped.Message, 'Speichern von "C:\Temp\Brief.odt" fehlgeschlagen', 'Zusammenhang fehlt');
      Assert.Contains(wrapped.Message, 'nicht mehr erreichbar', 'Grund fehlt');
      Assert.Contains(wrapped.Message, 'neu laden', 'Weg fehlt');
      Assert.Contains(wrapped.Message, 'RPC-Server', 'technischer Text fehlt');
    finally
      wrapped.Free;
    end;
  finally
    error.Free;
  end;
end;

procedure TOOToolsTests.OtherErrorsKeepTheirText;
var
  error: EOleSysError;
  wrapped: EOOAutomation;
begin
  // Ein gewöhnlicher UNO-Fehler wird nur eingepackt, nicht umgedeutet
  error := EOleSysError.Create('Typenkonflikt', HRESULT($80020005), 0);
  try
    wrapped := WrapUnoError('Schließen von "x.odt" fehlgeschlagen', error);
    try
      Assert.Contains(wrapped.Message, 'Schließen von "x.odt" fehlgeschlagen');
      Assert.Contains(wrapped.Message, 'Typenkonflikt');
      Assert.DoesNotContain(wrapped.Message, 'nicht mehr erreichbar');
    finally
      wrapped.Free;
    end;
  finally
    error.Free;
  end;
end;

initialization
  TDUnitX.RegisterTestFixture(TOOToolsTests);

end.
