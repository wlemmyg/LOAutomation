# LibreOffice Automation for Delphi

[English](#english) · [Deutsch](#deutsch)

---

## English

A small Delphi library that remote-controls **LibreOffice Writer** via OLE/COM: open a document, fill in
bookmarks and tables, then save it, export it as PDF or print it. There is no LibreOffice SDK to install and
no type library to import.

### Features

- Open, save, "save as", save a copy, export to PDF, print (printer, copies, collate, page range, paper
  format and orientation)
- Fill **bookmarks**, one by one or from a `TStrings` list (`Name=Value`)
- Fill **tables** by cell name (`A1`, `B3`, …) or from a two-dimensional array, adding missing rows
  automatically
- Hand a finished document over to the user, visible and still open
- Errors raise `EOOAutomation` with a plain message: what failed, in which document, and what to do instead.
  Nothing is silently truncated or skipped.

### Requirements

- Windows, Delphi **11, 12 or 13**, Win32 or Win64
- LibreOffice installed (tested with 7.5)

### Example

```delphi
uses
  OOTable, OOWriter;

procedure CreateLetter;
var
  writer: TOOWriter;
  rows: TArray<TArray<string>>;
begin
  // COM must be initialised by the caller. In a VCL application Application.Initialize already does this.
  writer := TOOWriter.Create;
  try
    writer.LoadFile('C:\Templates\Letter.odt', True);  // True = open hidden
    writer.WriteToBookmark('Recipient', 'Jane Doe');
    rows := [['1', 'Consulting', '100.00'], ['2', 'Planning', '250.00']];
    writer.TableByName('Items').Fill(rows, 2);          // start below the header row
    writer.ExportPdf('C:\Output\Letter.pdf');
    writer.CloseFile(False);
  finally
    writer.Free;
  end;
end;
```

One `TOOWriter` object holds one document. For several documents, create several objects. The library never
shuts LibreOffice down, because it shares the running LibreOffice instance with the user.

### Project layout

| Folder | Content |
|---|---|
| `source\src` | the library: `OOObject`, `OOWriter`, `OOTable`, `OOTools` |
| `source\demo` | VCL demo application |
| `source\tests` | DUnitX tests. Set `DUNITX` to the `Source` folder of a DUnitX clone and `OOAUTOMATION_TESTDATA` to `source\tests\testdata` |

### Status

Version 0.2. The library is being modernised, and the API may still change before 1.0. Only Writer is
supported so far.

### Background

The library goes back to the article "OpenOffice mit Delphi fernsteuern" in the German magazine
*Der Entwickler*, issue 6/2003. More at [delphi-tutorials.de](https://delphi-tutorials.de/projekte/).

### License

[Mozilla Public License 2.0](LICENSE). You may use the library in commercial and closed-source software.
If you distribute modified versions of the library's own files, you must publish those changes under the
MPL as well. Please keep the copyright and license notices in the source files. Issues and pull requests are
welcome.

---

## Deutsch

Eine kleine Delphi-Bibliothek, die **LibreOffice Writer** per OLE/COM fernsteuert: Dokument öffnen,
Textmarken und Tabellen füllen, dann speichern, als PDF exportieren oder drucken. Kein LibreOffice-SDK und
kein Import einer Typbibliothek nötig.

### Funktionen

- Öffnen, Speichern, „Speichern unter“, Kopie speichern, PDF-Export, Drucken (Drucker, Kopien, Sortierung,
  Seitenbereich, Papierformat und Ausrichtung)
- **Textmarken** füllen, einzeln oder aus einer `TStrings`-Liste (`Name=Wert`)
- **Tabellen** füllen, über den Zellnamen (`A1`, `B3`, …) oder aus einem zweidimensionalen Array. Fehlende
  Zeilen werden automatisch angehängt.
- Ein fertiges Dokument sichtbar und geöffnet an den Benutzer übergeben
- Fehler lösen `EOOAutomation` mit Klartext aus: was scheiterte, in welchem Dokument und welcher Weg
  offensteht. Nichts wird still abgeschnitten oder übergangen.

### Voraussetzungen

- Windows, Delphi **11, 12 oder 13**, Win32 oder Win64
- installiertes LibreOffice (getestet mit 7.5)

### Beispiel

```delphi
uses
  OOTable, OOWriter;

procedure CreateLetter;
var
  writer: TOOWriter;
  rows: TArray<TArray<string>>;
begin
  // COM initialisiert der Aufrufer. In einer VCL-Anwendung erledigt das bereits Application.Initialize.
  writer := TOOWriter.Create;
  try
    writer.LoadFile('C:\Vorlagen\Brief.odt', True);    // True = unsichtbar öffnen
    writer.WriteToBookmark('Empfaenger', 'Max Mustermann');
    rows := [['1', 'Beratung', '100,00'], ['2', 'Planung', '250,00']];
    writer.TableByName('Positionen').Fill(rows, 2);     // ab Zeile 2, unter der Kopfzeile
    writer.ExportPdf('C:\Ausgabe\Brief.pdf');
    writer.CloseFile(False);
  finally
    writer.Free;
  end;
end;
```

Ein `TOOWriter`-Objekt hält ein Dokument. Wer mehrere Dokumente braucht, legt mehrere Objekte an. Die
Bibliothek beendet LibreOffice nie, weil sie sich die laufende LibreOffice-Instanz mit dem Benutzer teilt.

### Aufbau

| Ordner | Inhalt |
|---|---|
| `source\src` | die Bibliothek: `OOObject`, `OOWriter`, `OOTable`, `OOTools` |
| `source\demo` | VCL-Demoanwendung |
| `source\tests` | DUnitX-Tests. `DUNITX` auf den `Source`-Ordner eines DUnitX-Klons und `OOAUTOMATION_TESTDATA` auf `source\tests\testdata` setzen |

### Stand

Version 0.2. Die Bibliothek wird gerade modernisiert, bis 1.0 kann sich die API noch ändern. Bisher wird nur
Writer unterstützt.

### Hintergrund

Die Bibliothek geht auf den Artikel „OpenOffice mit Delphi fernsteuern“ in *Der Entwickler*, Ausgabe 6/2003,
zurück. Mehr dazu auf [delphi-tutorials.de](https://delphi-tutorials.de/projekte/).

### Lizenz

[Mozilla Public License 2.0](LICENSE). Die Bibliothek darf auch in kommerzieller und Closed-Source-Software
eingesetzt werden. Wer geänderte Fassungen der Bibliotheksdateien weitergibt, muss diese Änderungen ebenfalls
unter der MPL veröffentlichen. Bitte die Copyright- und Lizenzhinweise in den Quelldateien stehen lassen.
Issues und Pull Requests sind willkommen.
