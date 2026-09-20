# LibreOffice Automation for Delphi

[English](#english) · [Deutsch](#deutsch)

---

## English

A small Delphi library that remote-controls **LibreOffice Writer** via OLE/COM: open a document, fill in
bookmarks and tables, then save it, export it as PDF or print it. There is no LibreOffice SDK to install and
no type library to import.

### Features

- Open, save, "save as", save a copy, export to PDF, print (printer, copies, collate, page range). Paper
  format and orientation come from the document's page style.
- Fill **bookmarks**, one by one or from a `TStrings` list (`Name=Value`)
- Fill **tables** by cell name (`A1`, `B3`, …) or from a two-dimensional array, adding missing rows
  automatically
- Hand a finished document over to the user, visible and still open
- No VCL dependency, no dialogs: the library is usable from a service or a console program
- Errors raise `EOOAutomation` with a plain message: what failed, in which document, and what to do instead.
  Nothing is silently truncated or skipped.

### Requirements

- Windows, Delphi **11, 12 or 13**, Win32 or Win64
- LibreOffice installed (tested with 7.5.3.2)
- **COM is initialised by the caller.** The library calls neither `CoInitialize` nor `CoUninitialize`. In a
  VCL application `Application.Initialize` already does this.
- For the test suite only: a [DUnitX](https://github.com/VSoftTechnologies/DUnitX) clone

### Example

```delphi
uses
  OOTable, OOTools, OOWriter;

procedure CreateLetter;
var
  writer: TOOWriter;
  rows: TArray<TArray<string>>;
  options: TOOPrintOptions;
begin
  // COM must be initialised by the caller. In a VCL application Application.Initialize already does this.
  writer := TOOWriter.Create;
  try
    writer.LoadFile('C:\Templates\Letter.odt', True);  // True = open hidden
    writer.WriteToBookmark('Recipient', 'Jane Doe');
    rows := [['1', 'Consulting', '100.00'], ['2', 'Planning', '250.00']];
    writer.TableByName('Items').Fill(rows, 2);          // start below the header row
    writer.ExportPdf('C:\Output\Letter.pdf');

    options := TOOPrintOptions.Default;                 // 1 copy, collated, all pages
    options.PrinterName := 'Invoice printer';
    options.Copies := 2;
    options.Pages := '1-3;5';
    writer.Print(options);

    writer.CloseFile(False);
  finally
    writer.Free;
  end;
end;
```

One `TOOWriter` object holds one document. For several documents, create several objects. The library never
shuts LibreOffice down, because it shares the running LibreOffice instance with the user.

### API

`TOOObject` is the abstract base class for a document; use `TOOWriter`.

**`TOOObject` – one document**

| Member | What it does |
|---|---|
| `Create` | connects to LibreOffice; if it cannot be reached, `EOOAutomation` says so in plain words |
| `LoadFile(FileName, Hidden = False)` | opens a document. The object must be empty — one object, one document |
| `Save` | stores the document under its current name |
| `SaveAs(FileName)` | stores it and renames it — the document is called that from now on |
| `SaveCopyAs(FileName)` | writes a copy; the document keeps its own name |
| `ExportPdf(FileName)` | PDF export (filter `writer_pdf_Export`) |
| `Print(Options)` | prints and **waits** until LibreOffice has handed the job over |
| `CloseFile(Save = False)` | closes the document; the object is empty afterwards |
| `HandOver` | makes the document visible and lets go of it — it stays open for the user |
| `IsLoaded` | is a document loaded? |
| `FileName` | read-only; set by `LoadFile` and `SaveAs` |
| `Visible` | shows or hides the document window |
| `Destroy` | closes a document the object still owns, without saving |

**`TOOWriter` – a Writer document**, adds to the above:

| Member | What it does |
|---|---|
| `WriteToBookmark(Name, Value)` | writes into one bookmark |
| `WriteToBookmarks(Values)` | `TStrings` with lines `Name=Value` |
| `GetBookmarkNames(List)` | fills a `TStrings` with all bookmark names |
| `GetTableNames(List)` | fills a `TStrings` with all table names |
| `TableByName(Name)` | a borrowed `TOOTable`; an unknown name raises `EOOAutomation` |
| `FindTable(Name)` | a borrowed `TOOTable`, or `nil` if there is no such table |

**`TOOTable` – a text table.** The object belongs to the `TOOWriter`; you only borrow it.

| Member | What it does |
|---|---|
| `SetCell(CellName, Value)` | by cell name, e.g. `A1` |
| `SetCells(Values)` | `TStrings` with lines `A1=Value` |
| `InsertRows(AfterRow, Count)` | 1-based; `0` inserts at the very top |
| `Fill(Data, StartRow = 1)` | fills from a two-dimensional array and appends missing rows |
| `Name` | the table's name |

**`TOOPrintOptions`** (a record in `OOTools`; `TOOPrintOptions.Default` gives one printer-less copy, collated,
all pages)

| Field | Meaning |
|---|---|
| `PrinterName` | empty = the printer the document already has |
| `Copies` | number of copies |
| `Collate` | `True` = 1-2-3, 1-2-3 instead of 1-1, 2-2, 3-3 |
| `Pages` | empty = all pages, otherwise LibreOffice syntax, e.g. `1-3;5` |

**Helpers in `OOTools`**

| Function | What it does |
|---|---|
| `FileNameToUrl(FileName)` | absolute path or UNC path → correctly encoded `file:` URL |
| `PrinterExists(Name)` | is a printer of that name installed or connected? |
| `EOOAutomation` | the exception class every error of the library uses |

### Things worth knowing

- **Bookmark names are case-sensitive.** LibreOffice reports an unknown name *without any message text*, so
  the library adds the name and the document to the error itself.
- **Table pointers are borrowed.** `TableByName` and `FindTable` return objects owned by the `TOOWriter`.
  After `CloseFile`, `HandOver` or loading another document they are invalid — fetch them again, do not
  keep them.
- **Paper format and orientation are not print options.** LibreOffice ignores an orientation passed to the
  printer, and a paper format passed there re-formats the *document* (new page breaks, document marked as
  modified). Both therefore come from the page style, where they belong.
- **Use absolute paths.** A relative path is rejected with an error instead of being resolved against some
  current directory. If you ever build `file:` URLs yourself, encode them: with a raw `#` in the path
  LibreOffice silently stores the document at the truncated path and still reports the intended one.
- **Overwriting is silent.** `SaveAs`, `SaveCopyAs` and `ExportPdf` replace an existing file without asking.
- **If LibreOffice disappears** — ended by the user, crashed, or the document closed by hand — the next call
  fails with `EOOAutomation` explaining what happened and that the document has to be loaded again. Windows
  alone would only say "The RPC server is unavailable". The object lets the lost document go, so releasing it
  afterwards stays quiet.
- **A modal LibreOffice dialog blocks automation** (document recovery after a crash, macro warning). If calls
  seem to hang, look at LibreOffice itself.
- **Threads:** the library does no synchronisation of its own. Give each thread its own objects and initialise
  COM in that thread.
- **Loading the same file twice opens a second document.** A document that a user opened never becomes "ours",
  and the library never closes LibreOffice itself.

### Project layout

| Folder | Content |
|---|---|
| `source\src` | the library: `OOObject`, `OOWriter`, `OOTable`, `OOTools` |
| `source\demo` | VCL demo application, one button per function |
| `source\tests` | DUnitX tests: unit tests plus integration tests that drive a real LibreOffice |

### Building and testing

`source\LibreOfficeAutomation.groupproj` builds the library, the demo and the tests. From a command line:

```
call "C:\Program Files (x86)\Embarcadero\Studio\<version>\bin\rsvars.bat"
msbuild source\LibreOfficeAutomation.groupproj /t:Build /p:Config=Debug /p:Platform=<Win32|Win64>
set OOAUTOMATION_TESTDATA=source\tests\testdata
OOAutomationTests.exe
```

Two environment variables:

- `DUNITX` — the `Source` folder of your DUnitX clone. The test project checks it before compiling and stops
  with a clear message if it is missing. (Without that check Delphi 11 and 12 would quietly use the DUnitX
  they ship, which is not necessarily the one you meant to test against.)
- `OOAUTOMATION_TESTDATA` — the `source\tests\testdata` folder.

The integration tests are in the DUnitX category `LibreOffice`. They work on copies in the temp folder and
leave LibreOffice as they found it: they close only the documents they opened themselves, and they shut
LibreOffice down only if it was not already running when the test run started.

### Changes in 0.2.0

The first version dates back to 2005/06, written for Delphi 5 to Turbo Delphi. What changed:

- **Delphi 11, 12 and 13, Win32 and Win64**, Unicode throughout. The old `oo.inc` with its list of compiler
  versions is gone; there are no version switches left.
- **Printing actually works.** Copies, collate and page range are applied now — the old version passed one
  and the same property object several times, so only the last setting ever arrived. (The old manual
  therefore stated that the number of copies and the pages could not be set.)
- **Closing a document works.** The UNO call needs its argument; without it every `CloseFile` failed.
- **File names become correctly encoded `file:` URLs.** Umlauts, spaces and `#` used to fail — `#` even
  silently, by storing the document somewhere else.
- **Errors are plain text.** Everything raises `EOOAutomation` saying what failed, in which document, and
  what to do instead, instead of passing raw OLE errors on. A lost document is explained rather than reported
  as an RPC failure.
- **New:** `ExportPdf`, `HandOver`, `TOOTable.Fill` and `InsertRows`, `TOOPrintOptions`, `PrinterExists`,
  `IsLoaded`.
- **Renamed:** `free(SaveFile)` → `Destroy` plus `CloseFile(Save)`, `SaveFileAs` → `SaveAs`, `SaveFileTo` →
  `SaveCopyAs`, `WriteListToBookmarks` → `WriteToBookmarks`, `pGetBookmarksList` / `pGetTablesList` →
  `GetBookmarkNames` / `GetTableNames`, `GetTableByName` → `TableByName` (raises) and `FindTable` (returns
  `nil`), `pSetTableValues` → `SetCells`. `FileName` is read-only now.
- **Removed:** `MailMerge` (it never worked as documented), the `PaperFormat` property (see above), and the
  `Vcl.Printers` dependency — `Print` takes a `TOOPrintOptions` record instead of a `TPrinter`.
- **51 DUnitX tests**, green on all three Delphi versions in both platforms. Every fixed bug has a test that
  was red before.

### Notes for contributors

Late binding to UNO has a few sharp edges that cost time to find again:

- **No empty parentheses.** `doc.store` works, `doc.store()` fails with "Too many parameters" — Delphi sends
  an argument for the `()`.
- **One property object per array entry.** Reusing a single `PropertyValue` for several entries makes all of
  them arrive with the last state.
- **UNO `short` needs `VarAsType(x, varSmallint)`**; a plain `OleVariant := SmallInt(x)` becomes `varInteger`.
- **`getCellByName` returns null** for an unknown cell instead of raising.

### Status

Version 0.2. The library is being modernised, and the API may still change before 1.0. Only Writer is
supported so far.

### Background

The library goes back to the article "OO-Automation" in the German magazine
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
  Seitenbereich). Papierformat und Ausrichtung bestimmt die Seitenvorlage des Dokuments.
- **Textmarken** füllen, einzeln oder aus einer `TStrings`-Liste (`Name=Wert`)
- **Tabellen** füllen, über den Zellnamen (`A1`, `B3`, …) oder aus einem zweidimensionalen Array. Fehlende
  Zeilen werden automatisch angehängt.
- Ein fertiges Dokument sichtbar und geöffnet an den Benutzer übergeben
- Keine VCL-Abhängigkeit, keine Dialoge: auch aus einem Dienst oder einem Konsolenprogramm nutzbar
- Fehler lösen `EOOAutomation` mit Klartext aus: was scheiterte, in welchem Dokument und welcher Weg
  offensteht. Nichts wird still abgeschnitten oder übergangen.

### Voraussetzungen

- Windows, Delphi **11, 12 oder 13**, Win32 oder Win64
- installiertes LibreOffice (getestet mit 7.5.3.2)
- **COM initialisiert der Aufrufer.** Die Bibliothek ruft weder `CoInitialize` noch `CoUninitialize`. In einer
  VCL-Anwendung erledigt das bereits `Application.Initialize`.
- nur für die Tests: ein [DUnitX](https://github.com/VSoftTechnologies/DUnitX)-Klon

### Beispiel

```delphi
uses
  OOTable, OOTools, OOWriter;

procedure CreateLetter;
var
  writer: TOOWriter;
  rows: TArray<TArray<string>>;
  options: TOOPrintOptions;
begin
  // COM initialisiert der Aufrufer. In einer VCL-Anwendung erledigt das bereits Application.Initialize.
  writer := TOOWriter.Create;
  try
    writer.LoadFile('C:\Vorlagen\Brief.odt', True);    // True = unsichtbar öffnen
    writer.WriteToBookmark('Empfaenger', 'Max Mustermann');
    rows := [['1', 'Beratung', '100,00'], ['2', 'Planung', '250,00']];
    writer.TableByName('Positionen').Fill(rows, 2);     // ab Zeile 2, unter der Kopfzeile
    writer.ExportPdf('C:\Ausgabe\Brief.pdf');

    options := TOOPrintOptions.Default;                 // 1 Kopie, sortiert, alle Seiten
    options.PrinterName := 'Rechnungsdrucker';
    options.Copies := 2;
    options.Pages := '1-3;5';
    writer.Print(options);

    writer.CloseFile(False);
  finally
    writer.Free;
  end;
end;
```

Ein `TOOWriter`-Objekt hält ein Dokument. Wer mehrere Dokumente braucht, legt mehrere Objekte an. Die
Bibliothek beendet LibreOffice nie, weil sie sich die laufende LibreOffice-Instanz mit dem Benutzer teilt.

### API

`TOOObject` ist die abstrakte Basisklasse für ein Dokument; benutzt wird `TOOWriter`.

**`TOOObject` – ein Dokument**

| Element | Wirkung |
|---|---|
| `Create` | verbindet mit LibreOffice; ist es nicht erreichbar, sagt `EOOAutomation` das im Klartext |
| `LoadFile(Datei, Versteckt = False)` | öffnet ein Dokument. Das Objekt muss leer sein — ein Objekt, ein Dokument |
| `Save` | speichert unter dem bisherigen Namen |
| `SaveAs(Datei)` | speichert und benennt um — das Dokument heißt danach so |
| `SaveCopyAs(Datei)` | schreibt eine Kopie; das Dokument behält seinen Namen |
| `ExportPdf(Datei)` | PDF-Export (Filter `writer_pdf_Export`) |
| `Print(Optionen)` | druckt und **wartet**, bis LibreOffice den Auftrag abgegeben hat |
| `CloseFile(Speichern = False)` | schließt das Dokument; das Objekt ist danach leer |
| `HandOver` | macht das Dokument sichtbar und lässt es los — es bleibt beim Benutzer offen |
| `IsLoaded` | ist ein Dokument geladen? |
| `FileName` | nur lesend; gesetzt von `LoadFile` und `SaveAs` |
| `Visible` | zeigt das Dokumentfenster oder versteckt es |
| `Destroy` | schließt ein noch eigenes Dokument, ohne zu speichern |

**`TOOWriter` – ein Writer-Dokument**, zusätzlich zu den obigen:

| Element | Wirkung |
|---|---|
| `WriteToBookmark(Name, Wert)` | schreibt in eine Textmarke |
| `WriteToBookmarks(Werte)` | `TStrings` mit Zeilen `Name=Wert` |
| `GetBookmarkNames(Liste)` | füllt eine `TStrings` mit allen Textmarkennamen |
| `GetTableNames(Liste)` | füllt eine `TStrings` mit allen Tabellennamen |
| `TableByName(Name)` | eine geliehene `TOOTable`; ein unbekannter Name löst `EOOAutomation` aus |
| `FindTable(Name)` | eine geliehene `TOOTable` oder `nil`, wenn es sie nicht gibt |

**`TOOTable` – eine Texttabelle.** Das Objekt gehört dem `TOOWriter`, der Aufrufer bekommt es nur geliehen.

| Element | Wirkung |
|---|---|
| `SetCell(Zellname, Wert)` | über den Zellnamen, z. B. `A1` |
| `SetCells(Werte)` | `TStrings` mit Zeilen `A1=Wert` |
| `InsertRows(NachZeile, Anzahl)` | 1-basiert; `0` fügt ganz oben ein |
| `Fill(Daten, AbZeile = 1)` | füllt aus einem zweidimensionalen Array und hängt fehlende Zeilen an |
| `Name` | Name der Tabelle |

**`TOOPrintOptions`** (Record in `OOTools`; `TOOPrintOptions.Default` liefert eine Kopie, sortiert, alle
Seiten, ohne festen Drucker)

| Feld | Bedeutung |
|---|---|
| `PrinterName` | leer = der Drucker, den das Dokument bereits hat |
| `Copies` | Anzahl der Kopien |
| `Collate` | `True` = 1-2-3, 1-2-3 statt 1-1, 2-2, 3-3 |
| `Pages` | leer = alle Seiten, sonst LibreOffice-Syntax, z. B. `1-3;5` |

**Helfer in `OOTools`**

| Funktion | Wirkung |
|---|---|
| `FileNameToUrl(Datei)` | absoluter Pfad oder UNC-Pfad → richtig kodierte `file:`-URL |
| `PrinterExists(Name)` | ist ein Drucker dieses Namens installiert oder verbunden? |
| `EOOAutomation` | die Ausnahmeklasse, die jeder Fehler der Bibliothek benutzt |

### Was man wissen sollte

- **Textmarkennamen sind case-sensitiv.** LibreOffice meldet einen unbekannten Namen *ohne jeden
  Meldungstext*, deshalb setzt die Bibliothek Name und Dokument selbst in den Fehler.
- **Tabellenzeiger sind geliehen.** `TableByName` und `FindTable` liefern Objekte, die dem `TOOWriter`
  gehören. Nach `CloseFile`, `HandOver` oder dem Laden eines anderen Dokuments sind sie ungültig — dann neu
  holen, nicht aufbewahren.
- **Papierformat und Ausrichtung sind keine Druckoptionen.** Eine an den Drucker übergebene Ausrichtung
  ignoriert LibreOffice, und ein übergebenes Papierformat formatiert das *Dokument* um (neuer Umbruch,
  Dokument gilt als geändert). Beides kommt deshalb aus der Seitenvorlage, wo es hingehört.
- **Absolute Pfade benutzen.** Ein relativer Pfad wird mit einer Meldung abgelehnt, statt still gegen
  irgendein Arbeitsverzeichnis aufgelöst zu werden. Wer selbst `file:`-URLs baut, muss sie kodieren: Bei
  einem rohen `#` im Pfad speichert LibreOffice das Dokument **kommentarlos** am abgeschnittenen Pfad und
  meldet trotzdem das gewünschte Ziel.
- **Überschrieben wird ohne Rückfrage.** `SaveAs`, `SaveCopyAs` und `ExportPdf` ersetzen eine vorhandene
  Datei stillschweigend.
- **Fällt LibreOffice weg** — vom Benutzer beendet, abgestürzt oder das Dokument von Hand geschlossen —,
  scheitert der nächste Aufruf mit `EOOAutomation` im Klartext samt Hinweis, dass das Dokument neu geladen
  werden muss. Windows allein meldet nur „Der RPC-Server ist nicht verfügbar“. Das Objekt lässt das verlorene
  Dokument los, das spätere Freigeben bleibt deshalb stumm.
- **Ein modaler LibreOffice-Dialog blockiert die Automation** (Wiederherstellung nach einem Absturz,
  Makro-Warnung). Wenn Aufrufe scheinbar hängen, lohnt der Blick auf LibreOffice selbst.
- **Threads:** Die Bibliothek synchronisiert nichts von sich aus. Jeder Thread bekommt eigene Objekte und
  initialisiert COM selbst.
- **Dieselbe Datei zweimal laden öffnet ein zweites Dokument.** Ein vom Benutzer geöffnetes Dokument wird nie
  „unseres“, und die Bibliothek beendet LibreOffice nie.

### Aufbau

| Ordner | Inhalt |
|---|---|
| `source\src` | die Bibliothek: `OOObject`, `OOWriter`, `OOTable`, `OOTools` |
| `source\demo` | VCL-Demoanwendung, je Funktion eine Schaltfläche |
| `source\tests` | DUnitX-Tests: Unit-Tests und Integrationstests gegen ein echtes LibreOffice |

### Bauen und testen

`source\LibreOfficeAutomation.groupproj` baut Bibliothek, Demo und Tests. Von der Kommandozeile:

```
call "C:\Program Files (x86)\Embarcadero\Studio\<Version>\bin\rsvars.bat"
msbuild source\LibreOfficeAutomation.groupproj /t:Build /p:Config=Debug /p:Platform=<Win32|Win64>
set OOAUTOMATION_TESTDATA=source\tests\testdata
OOAutomationTests.exe
```

Zwei Umgebungsvariablen:

- `DUNITX` — der `Source`-Ordner des DUnitX-Klons. Das Testprojekt prüft ihn vor dem Compiler und bricht mit
  einer klaren Meldung ab, wenn dort nichts liegt. (Ohne diese Prüfung nähmen Delphi 11 und 12 still ihre
  mitgelieferte DUnitX — nicht unbedingt die, gegen die getestet werden sollte.)
- `OOAUTOMATION_TESTDATA` — der Ordner `source\tests\testdata`.

Die Integrationstests stehen in der DUnitX-Kategorie `LibreOffice`. Sie arbeiten auf Kopien im Temp-Ordner und
hinterlassen LibreOffice so, wie sie es vorgefunden haben: Sie schließen nur die Dokumente, die sie selbst
geöffnet haben, und beenden LibreOffice nur dann, wenn es zu Beginn des Laufs nicht schon lief.

### Änderungen in 0.2.0

Die erste Fassung stammt von 2005/06 und war für Delphi 5 bis Turbo Delphi geschrieben. Was sich geändert hat:

- **Delphi 11, 12 und 13, Win32 und Win64**, durchgängig Unicode. Die alte `oo.inc` mit ihrer Liste von
  Compiler-Versionen ist weg, Versionsweichen gibt es keine mehr.
- **Drucken funktioniert wirklich.** Kopien, Sortierung und Seitenbereich kommen jetzt an — die alte Fassung
  reichte ein und dasselbe Eigenschaftsobjekt mehrfach weiter, sodass immer nur der letzte Stand ankam.
  (Das alte Handbuch hielt deshalb fest, Anzahl der Ausdrucke und Seiten ließen sich nicht festlegen.)
- **Dokumente schließen funktioniert.** Der UNO-Aufruf braucht sein Argument; ohne das scheiterte jedes
  `CloseFile`.
- **Dateinamen werden zu richtig kodierten `file:`-URLs.** Umlaute, Leerzeichen und `#` scheiterten vorher —
  `#` sogar still, indem das Dokument woanders landete.
- **Fehler sind Klartext.** Alles löst `EOOAutomation` aus und sagt, was scheiterte, in welchem Dokument und
  welcher Weg offensteht, statt rohe OLE-Fehler durchzureichen. Ein verlorenes Dokument wird erklärt, statt
  als RPC-Fehler gemeldet zu werden.
- **Neu:** `ExportPdf`, `HandOver`, `TOOTable.Fill` und `InsertRows`, `TOOPrintOptions`, `PrinterExists`,
  `IsLoaded`.
- **Umbenannt:** `free(SaveFile)` → `Destroy` und `CloseFile(Speichern)`, `SaveFileAs` → `SaveAs`,
  `SaveFileTo` → `SaveCopyAs`, `WriteListToBookmarks` → `WriteToBookmarks`, `pGetBookmarksList` /
  `pGetTablesList` → `GetBookmarkNames` / `GetTableNames`, `GetTableByName` → `TableByName` (löst aus) und
  `FindTable` (liefert `nil`), `pSetTableValues` → `SetCells`. `FileName` ist jetzt nur lesbar.
- **Entfallen:** `MailMerge` (funktionierte nie wie dokumentiert), die Eigenschaft `PaperFormat` (siehe oben)
  und die Abhängigkeit von `Vcl.Printers` — `Print` nimmt einen `TOOPrintOptions`-Record statt eines
  `TPrinter`.
- **51 DUnitX-Tests**, grün in allen drei Delphi-Versionen auf beiden Plattformen. Jeder behobene Fehler ist
  durch einen Test belegt, der vorher rot war.

### Hinweise für Mitarbeit

Late Binding an UNO hat ein paar Fallen, die man sonst zweimal sucht:

- **Keine leeren Klammern.** `doc.store` geht, `doc.store()` scheitert mit „Too many parameters“ — Delphi
  schickt bei `()` ein Argument mit.
- **Je Array-Eintrag ein eigenes Eigenschaftsobjekt.** Ein mehrfach verwendetes `PropertyValue` kommt in allen
  Einträgen mit dem letzten Stand an.
- **UNO-`short` braucht `VarAsType(x, varSmallint)`**; ein einfaches `OleVariant := SmallInt(x)` wird
  `varInteger`.
- **`getCellByName` liefert für eine unbekannte Zelle null**, statt eine Ausnahme auszulösen.

### Stand

Version 0.2. Die Bibliothek wird gerade modernisiert, bis 1.0 kann sich die API noch ändern. Bisher wird nur
Writer unterstützt.

### Hintergrund

Die Bibliothek geht auf den Artikel „OO-Automation“ in *Der Entwickler*, Ausgabe 6/2003,
zurück. Mehr dazu auf [delphi-tutorials.de](https://delphi-tutorials.de/projekte/).

### Lizenz

[Mozilla Public License 2.0](LICENSE). Die Bibliothek darf auch in kommerzieller und Closed-Source-Software
eingesetzt werden. Wer geänderte Fassungen der Bibliotheksdateien weitergibt, muss diese Änderungen ebenfalls
unter der MPL veröffentlichen. Bitte die Copyright- und Lizenzhinweise in den Quelldateien stehen lassen.
Issues und Pull Requests sind willkommen.
