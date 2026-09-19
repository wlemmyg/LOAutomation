unit UMain;

// P3: nur auf die neue API umgestellt, damit die Build-Matrix baut; die eigentliche Demo folgt in P4

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Printers, OOWriter, OOTools;

type
  TForm1 = class(TForm)
    ButtonOpen: TButton;
    OpenDialog1: TOpenDialog;
    SaveDialog1: TSaveDialog;
    Button2: TButton;
    PrintDialog1: TPrintDialog;
    ComboBox1: TComboBox;
    ButtonClose: TButton;
    Button4: TButton;
    Button5: TButton;
    Button_TableTest: TButton;
    procedure Button5Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure ButtonCloseClick(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure ButtonOpenClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Button_TableTestClick(Sender: TObject);
  private
    { Private-Deklarationen }
  public
    { Public-Deklarationen }
    OO:TOOWriter;
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.FormCreate(Sender: TObject);
begin
  OO:=TOOWriter.Create;
end;

procedure TForm1.ButtonOpenClick(Sender: TObject);
begin
  if OpenDialog1.Execute then
  begin
    // Ein Objekt hält genau ein Dokument (Q1)
    if OO.IsLoaded then
    begin
      OO.CloseFile(False);
    end;
    OO.LoadFile(OpenDialog1.FileName, False);
  end;
end;

procedure TForm1.Button_TableTestClick(Sender: TObject);
var
  values: TStringList;
begin
  values:=TStringList.Create;
  try
    values.Add('A1=Hallo');
    values.Add('A2=Welt');
    OO.TableByName('Tabelle1').SetCells(values);
  finally
    FreeAndNil(values);
  end;
end;

procedure TForm1.Button2Click(Sender: TObject);
var
  options: TOOPrintOptions;
begin
  if PrintDialog1.Execute then
  begin
    options:=TOOPrintOptions.Default;
    options.PrinterName:=Printer.Printers[Printer.PrinterIndex];
    options.Copies:=PrintDialog1.Copies;
    options.Collate:=PrintDialog1.Collate;
    // Ohne Auswahl bleibt das Papier des Dokuments (E2); die Einträge folgen der Reihenfolge von TOOPaperFormat
    if ComboBox1.ItemIndex>=0 then
    begin
      options.OverridePaper:=True;
      options.PaperFormat:=TOOPaperFormat(ComboBox1.ItemIndex);
      if Printer.Orientation=poLandscape then
      begin
        options.Orientation:=ooLandscape;
      end;
    end;
    OO.Print(options);
  end;
end;

procedure TForm1.ButtonCloseClick(Sender: TObject);
begin
  OO.CloseFile(False);
end;

procedure TForm1.FormDestroy(Sender: TObject);
begin
  FreeAndNil(OO);
end;

procedure TForm1.Button4Click(Sender: TObject);
begin
  if SaveDialog1.Execute then
  begin
    OO.SaveAs(SaveDialog1.FileName);
  end;
end;

procedure TForm1.Button5Click(Sender: TObject);
var
  values: TStringList;
begin
  values:=TStringList.Create;
  try
    values.Add('Test1=Hallo Welt!');
    values.Add('Test2=Hello World');
    OO.WriteToBookmarks(values);
  finally
    FreeAndNil(values);
  end;
end;

end.
