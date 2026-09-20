{
--------------------------------------------------------------------------------
Copyright (c) 2026 Wolfgang Lemmermeyer

This Source Code Form is subject to the terms of the Mozilla Public
License, v. 2.0. If a copy of the MPL was not distributed with this
file, You can obtain one at https://mozilla.org/MPL/2.0/.
--------------------------------------------------------------------------------
}
unit UMain;

// Demo für ooAutomation: Jede öffentliche Funktion der Bibliothek hat hier einen Knopf (P4). Fehler zeigt die
// Demo mit Standard-VCL-Dialogen und im Protokoll; die Bibliothek selbst zeigt nie Dialoge (A4).

interface

uses
  System.Classes,
  System.SysUtils,
  Vcl.Controls,
  Vcl.Dialogs,
  Vcl.Forms,
  Vcl.Grids,
  Vcl.Samples.Spin,
  Vcl.StdCtrls,
  Vcl.ValEdit,
  OOTable,
  OOWriter;

type
  TMainForm = class(TForm)
    grpDocument: TGroupBox;
    btnOpen: TButton;
    chkHidden: TCheckBox;
    btnClose: TButton;
    chkSaveOnClose: TCheckBox;
    btnSave: TButton;
    btnSaveAs: TButton;
    btnSaveCopy: TButton;
    btnExportPdf: TButton;
    btnHandOver: TButton;
    chkVisible: TCheckBox;
    grpBookmarks: TGroupBox;
    vleBookmarks: TValueListEditor;
    btnWriteBookmarks: TButton;
    lblFile: TLabel;
    grpTables: TGroupBox;
    lblTable: TLabel;
    cboTables: TComboBox;
    grdData: TStringGrid;
    lblStartRow: TLabel;
    spnStartRow: TSpinEdit;
    btnFill: TButton;
    lblAfterRow: TLabel;
    spnAfterRow: TSpinEdit;
    lblRowCount: TLabel;
    spnRowCount: TSpinEdit;
    btnInsertRows: TButton;
    grpPrint: TGroupBox;
    lblPages: TLabel;
    edtPages: TEdit;
    btnPrint: TButton;
    lblPrintHint: TLabel;
    memLog: TMemo;
    dlgOpen: TOpenDialog;
    dlgSave: TSaveDialog;
    dlgSavePdf: TSaveDialog;
    dlgPrint: TPrintDialog;
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure btnOpenClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnSaveAsClick(Sender: TObject);
    procedure btnSaveCopyClick(Sender: TObject);
    procedure btnExportPdfClick(Sender: TObject);
    procedure btnHandOverClick(Sender: TObject);
    procedure chkVisibleClick(Sender: TObject);
    procedure btnWriteBookmarksClick(Sender: TObject);
    procedure btnFillClick(Sender: TObject);
    procedure btnInsertRowsClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
  private
    FWriter: TOOWriter;
    FUpdating: Boolean;  // Häkchen per Code setzen, ohne OnClick auszulösen
    procedure RunAction(const ACaption: string; const AAction: TProc);
    procedure Log(const AText: string);
    procedure LoadDocumentLists;
    procedure ClearDocumentLists;
    procedure UpdateControls;
    procedure SetVisibleCheck(AChecked: Boolean);
    function SelectedTable: TOOTable;
    function GridData: TArray<TArray<string>>;
  end;

var
  MainForm: TMainForm;

implementation

uses
  System.UITypes,
  Vcl.Printers,
  OOTools;

{$R *.dfm}

{ ===== Lebenszyklus ===== }

procedure TMainForm.FormCreate(Sender: TObject);
begin
  grdData.Cells[0, 0] := '1';
  grdData.Cells[1, 0] := 'Planung';
  grdData.Cells[2, 0] := '100,00';
  grdData.Cells[0, 1] := '2';
  grdData.Cells[1, 1] := 'Bauleitung';
  grdData.Cells[2, 1] := '250,00';
  // Ohne LibreOffice bleibt die Demo offen, aber alles außer dem Protokoll gesperrt
  RunAction('Verbinden mit LibreOffice',
    procedure
    begin
      FWriter := TOOWriter.Create;
    end);
end;

procedure TMainForm.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  // Schließt ein noch eigenes Dokument ohne Speichern; ein übergebenes bleibt beim Benutzer (E1)
  FreeAndNil(FWriter);
end;

{ ===== Dokument ===== }

procedure TMainForm.btnOpenClick(Sender: TObject);
begin
  if not dlgOpen.Execute then
  begin
    Exit;
  end;
  // Hält das Objekt schon ein Dokument, lehnt LoadFile laut ab (Q1) – die Demo zeigt das absichtlich
  RunAction('Öffnen ' + dlgOpen.FileName,
    procedure
    begin
      FWriter.LoadFile(dlgOpen.FileName, chkHidden.Checked);
      LoadDocumentLists;
      SetVisibleCheck(not chkHidden.Checked);
    end);
end;

procedure TMainForm.btnCloseClick(Sender: TObject);
begin
  RunAction('Schließen',
    procedure
    begin
      FWriter.CloseFile(chkSaveOnClose.Checked);
      ClearDocumentLists;
    end);
end;

procedure TMainForm.btnSaveClick(Sender: TObject);
begin
  RunAction('Speichern',
    procedure
    begin
      FWriter.Save;
    end);
end;

procedure TMainForm.btnSaveAsClick(Sender: TObject);
begin
  if not dlgSave.Execute then
  begin
    Exit;
  end;
  RunAction('Speichern unter ' + dlgSave.FileName,
    procedure
    begin
      FWriter.SaveAs(dlgSave.FileName);
    end);
end;

procedure TMainForm.btnSaveCopyClick(Sender: TObject);
begin
  if not dlgSave.Execute then
  begin
    Exit;
  end;
  RunAction('Kopie speichern unter ' + dlgSave.FileName,
    procedure
    begin
      FWriter.SaveCopyAs(dlgSave.FileName);
    end);
end;

procedure TMainForm.btnExportPdfClick(Sender: TObject);
begin
  if not dlgSavePdf.Execute then
  begin
    Exit;
  end;
  RunAction('PDF-Export nach ' + dlgSavePdf.FileName,
    procedure
    begin
      FWriter.ExportPdf(dlgSavePdf.FileName);
    end);
end;

procedure TMainForm.btnHandOverClick(Sender: TObject);
begin
  RunAction('An Benutzer übergeben',
    procedure
    begin
      FWriter.HandOver;
      ClearDocumentLists;
    end);
end;

procedure TMainForm.chkVisibleClick(Sender: TObject);
begin
  if FUpdating then
  begin
    Exit;
  end;
  RunAction('Sichtbar = ' + BoolToStr(chkVisible.Checked, True),
    procedure
    begin
      FWriter.Visible := chkVisible.Checked;
    end);
end;

{ ===== Textmarken ===== }

procedure TMainForm.btnWriteBookmarksClick(Sender: TObject);
begin
  // Die Zeilen des Editors sind schon „Textmarke=Wert“; neue Zeilen darf man anlegen, um den Fehlerfall zu sehen
  RunAction('Textmarken schreiben',
    procedure
    begin
      FWriter.WriteToBookmarks(vleBookmarks.Strings);
    end);
end;

{ ===== Tabellen ===== }

procedure TMainForm.btnFillClick(Sender: TObject);
begin
  RunAction(Format('Tabelle füllen ab Zeile %d', [spnStartRow.Value]),
    procedure
    begin
      SelectedTable.Fill(GridData, spnStartRow.Value);
    end);
end;

procedure TMainForm.btnInsertRowsClick(Sender: TObject);
begin
  RunAction(Format('%d Zeile(n) einfügen nach Zeile %d', [spnRowCount.Value, spnAfterRow.Value]),
    procedure
    begin
      SelectedTable.InsertRows(spnAfterRow.Value, spnRowCount.Value);
    end);
end;

function TMainForm.SelectedTable: TOOTable;
begin
  if cboTables.ItemIndex < 0 then
  begin
    raise Exception.Create('Keine Tabelle ausgewählt.');
  end;
  // geliehen (Besitz: FWriter)
  Result := FWriter.TableByName(cboTables.Text);
end;

function TMainForm.GridData: TArray<TArray<string>>;
var
  lastRow: Integer;
  lastCol: Integer;
  idxRow: Integer;
  idxCol: Integer;
begin
  // Bis zur letzten gefüllten Zeile, je Zeile bis zur letzten gefüllten Zelle
  lastRow := -1;
  for idxRow := 0 to grdData.RowCount - 1 do
  begin
    for idxCol := 0 to grdData.ColCount - 1 do
    begin
      if grdData.Cells[idxCol, idxRow] <> '' then
      begin
        lastRow := idxRow;
      end;
    end;
  end;
  SetLength(Result, lastRow + 1);
  for idxRow := 0 to lastRow do
  begin
    lastCol := -1;
    for idxCol := 0 to grdData.ColCount - 1 do
    begin
      if grdData.Cells[idxCol, idxRow] <> '' then
      begin
        lastCol := idxCol;
      end;
    end;
    SetLength(Result[idxRow], lastCol + 1);
    for idxCol := 0 to lastCol do
    begin
      Result[idxRow][idxCol] := grdData.Cells[idxCol, idxRow];
    end;
  end;
end;

{ ===== Drucken ===== }

procedure TMainForm.btnPrintClick(Sender: TObject);
var
  options: TOOPrintOptions;
begin
  if not dlgPrint.Execute then
  begin
    Exit;
  end;
  options := TOOPrintOptions.Default;
  // Der Dialog stellt Printer.PrinterIndex um; die Namen stimmen mit PrinterExists überein (beide Ebene 4)
  options.PrinterName := Printer.Printers[Printer.PrinterIndex];
  options.Copies := dlgPrint.Copies;
  options.Collate := dlgPrint.Collate;
  if edtPages.Text <> '' then
  begin
    options.Pages := edtPages.Text;
  end
  else if dlgPrint.PrintRange = prPageNums then
  begin
    options.Pages := Format('%d-%d', [dlgPrint.FromPage, dlgPrint.ToPage]);
  end;
  // Papierformat und Ausrichtung bestimmt die Seitenvorlage des Dokuments, nicht der Druck (E5)
  RunAction(Format('Drucken auf "%s", %d Kopie(n), sortiert=%s, Seiten="%s"',
    [options.PrinterName, options.Copies, BoolToStr(options.Collate, True), options.Pages]),
    procedure
    begin
      FWriter.Print(options);
    end);
end;

{ ===== Helfer ===== }

procedure TMainForm.RunAction(const ACaption: string; const AAction: TProc);
begin
  try
    AAction();
    Log(ACaption + ': ok');
  except
    on E: Exception do
    begin
      // Angezeigt und protokolliert, so wie die Bibliothek den Fehler meldet: Klasse und Klartext
      Log(Format('%s: %s – %s', [ACaption, E.ClassName, E.Message]));
      MessageDlg(E.Message, mtError, [mbOK], 0);
    end;
  end;
  UpdateControls;
end;

procedure TMainForm.Log(const AText: string);
begin
  memLog.Lines.Add(FormatDateTime('hh:nn:ss', Now) + '  ' + AText);
end;

procedure TMainForm.LoadDocumentLists;
var
  names: TStringList;
  idx: Integer;
begin
  names := TStringList.Create;
  try
    FWriter.GetBookmarkNames(names);
    vleBookmarks.Strings.BeginUpdate;
    try
      vleBookmarks.Strings.Clear;
      for idx := 0 to names.Count - 1 do
      begin
        vleBookmarks.Strings.Add(names[idx] + '=');
      end;
    finally
      vleBookmarks.Strings.EndUpdate;
    end;
  finally
    names.Free;
  end;
  FWriter.GetTableNames(cboTables.Items);
  if cboTables.Items.Count > 0 then
  begin
    cboTables.ItemIndex := 0;
  end;
end;

procedure TMainForm.ClearDocumentLists;
begin
  vleBookmarks.Strings.Clear;
  cboTables.Items.Clear;
  SetVisibleCheck(False);
end;

procedure TMainForm.SetVisibleCheck(AChecked: Boolean);
begin
  FUpdating := True;
  try
    chkVisible.Checked := AChecked;
  finally
    FUpdating := False;
  end;
end;

procedure TMainForm.UpdateControls;
var
  connected: Boolean;
  loaded: Boolean;
  control: TControl;
begin
  connected := FWriter <> nil;
  loaded := connected and FWriter.IsLoaded;
  btnOpen.Enabled := connected;
  chkHidden.Enabled := connected;
  for control in TArray<TControl>.Create(btnClose, chkSaveOnClose, btnSave, btnSaveAs, btnSaveCopy, btnExportPdf,
    btnHandOver, chkVisible, vleBookmarks, btnWriteBookmarks, cboTables, grdData, spnStartRow, btnFill,
    spnAfterRow, spnRowCount, btnInsertRows, edtPages, btnPrint) do
  begin
    control.Enabled := loaded;
  end;
  if loaded then
  begin
    lblFile.Caption := FWriter.FileName;
  end
  else
  begin
    lblFile.Caption := '(kein Dokument geladen)';
  end;
end;

end.
