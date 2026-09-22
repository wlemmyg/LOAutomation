object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'ooAutomation '#8211' Demo'
  ClientHeight = 760
  ClientWidth = 800
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = 'Segoe UI'
  Font.Style = []
  Position = poScreenCenter
  OnClose = FormClose
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 15
  object lblFile: TLabel
    Left = 8
    Top = 212
    Width = 784
    Height = 15
    AutoSize = False
    Caption = '(kein Dokument geladen)'
    EllipsisPosition = epPathEllipsis
  end
  object grpDocument: TGroupBox
    Left = 8
    Top = 8
    Width = 384
    Height = 196
    Caption = 'Dokument'
    TabOrder = 0
    object btnOpen: TButton
      Left = 16
      Top = 24
      Width = 160
      Height = 25
      Caption = #214'ffnen '#8230
      TabOrder = 0
      OnClick = btnOpenClick
    end
    object chkHidden: TCheckBox
      Left = 192
      Top = 28
      Width = 176
      Height = 17
      Caption = 'versteckt laden'
      TabOrder = 1
    end
    object btnClose: TButton
      Left = 16
      Top = 56
      Width = 160
      Height = 25
      Caption = 'Schlie'#223'en'
      TabOrder = 2
      OnClick = btnCloseClick
    end
    object chkSaveOnClose: TCheckBox
      Left = 192
      Top = 60
      Width = 176
      Height = 17
      Caption = 'dabei speichern'
      TabOrder = 3
    end
    object btnSave: TButton
      Left = 16
      Top = 88
      Width = 160
      Height = 25
      Caption = 'Speichern'
      TabOrder = 4
      OnClick = btnSaveClick
    end
    object btnSaveAs: TButton
      Left = 192
      Top = 88
      Width = 176
      Height = 25
      Caption = 'Speichern unter '#8230
      TabOrder = 5
      OnClick = btnSaveAsClick
    end
    object btnSaveCopy: TButton
      Left = 16
      Top = 120
      Width = 160
      Height = 25
      Caption = 'Kopie speichern unter '#8230
      TabOrder = 6
      OnClick = btnSaveCopyClick
    end
    object btnExportPdf: TButton
      Left = 192
      Top = 120
      Width = 176
      Height = 25
      Caption = 'Als PDF exportieren '#8230
      TabOrder = 7
      OnClick = btnExportPdfClick
    end
    object btnHandOver: TButton
      Left = 16
      Top = 152
      Width = 160
      Height = 25
      Caption = 'An Benutzer '#252'bergeben'
      TabOrder = 8
      OnClick = btnHandOverClick
    end
    object chkVisible: TCheckBox
      Left = 192
      Top = 156
      Width = 176
      Height = 17
      Caption = 'sichtbar'
      TabOrder = 9
      OnClick = chkVisibleClick
    end
  end
  object grpBookmarks: TGroupBox
    Left = 400
    Top = 8
    Width = 392
    Height = 196
    Caption = 'Textmarken'
    TabOrder = 1
    object vleBookmarks: TValueListEditor
      Left = 16
      Top = 24
      Width = 360
      Height = 128
      KeyOptions = [keyEdit, keyAdd, keyDelete]
      TabOrder = 0
      TitleCaptions.Strings = (
        'Textmarke'
        'Wert')
      ColWidths = (
        150
        204)
    end
    object btnWriteBookmarks: TButton
      Left = 16
      Top = 160
      Width = 200
      Height = 25
      Caption = 'In Textmarken schreiben'
      TabOrder = 1
      OnClick = btnWriteBookmarksClick
    end
  end
  object grpTables: TGroupBox
    Left = 8
    Top = 236
    Width = 384
    Height = 236
    Caption = 'Tabellen'
    TabOrder = 2
    object lblTable: TLabel
      Left = 16
      Top = 28
      Width = 45
      Height = 15
      Caption = 'Tabelle:'
    end
    object lblStartRow: TLabel
      Left = 16
      Top = 178
      Width = 45
      Height = 15
      Caption = 'ab Zeile'
    end
    object lblAfterRow: TLabel
      Left = 16
      Top = 206
      Width = 56
      Height = 15
      Caption = 'nach Zeile'
    end
    object lblRowCount: TLabel
      Left = 136
      Top = 206
      Width = 40
      Height = 15
      Caption = 'Anzahl'
    end
    object cboTables: TComboBox
      Left = 80
      Top = 24
      Width = 160
      Height = 23
      Style = csDropDownList
      TabOrder = 0
    end
    object grdData: TStringGrid
      Left = 16
      Top = 56
      Width = 352
      Height = 110
      ColCount = 4
      DefaultColWidth = 84
      DefaultRowHeight = 20
      FixedCols = 0
      RowCount = 5
      FixedRows = 0
      Options = [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine, goRangeSelect, goEditing]
      TabOrder = 1
    end
    object spnStartRow: TSpinEdit
      Left = 80
      Top = 174
      Width = 56
      Height = 24
      MaxValue = 999
      MinValue = 1
      TabOrder = 2
      Value = 1
    end
    object btnFill: TButton
      Left = 144
      Top = 173
      Width = 224
      Height = 25
      Caption = 'Tabelle f'#252'llen'
      TabOrder = 3
      OnClick = btnFillClick
    end
    object spnAfterRow: TSpinEdit
      Left = 80
      Top = 202
      Width = 48
      Height = 24
      MaxValue = 999
      MinValue = 0
      TabOrder = 4
      Value = 1
    end
    object spnRowCount: TSpinEdit
      Left = 184
      Top = 202
      Width = 48
      Height = 24
      MaxValue = 99
      MinValue = 1
      TabOrder = 5
      Value = 1
    end
    object btnInsertRows: TButton
      Left = 240
      Top = 201
      Width = 128
      Height = 25
      Caption = 'Zeilen einf'#252'gen'
      TabOrder = 6
      OnClick = btnInsertRowsClick
    end
  end
  object grpPrint: TGroupBox
    Left = 400
    Top = 236
    Width = 392
    Height = 236
    Caption = 'Drucken'
    TabOrder = 3
    object lblPages: TLabel
      Left = 16
      Top = 28
      Width = 38
      Height = 15
      Caption = 'Seiten:'
    end
    object lblPrintHint: TLabel
      Left = 16
      Top = 100
      Width = 360
      Height = 64
      AutoSize = False
      Caption =
        'Drucker, Kopien und Sortierung kommen aus dem Druckdialog, die S' +
        'eiten aus dem Feld oder dem Dialog. Papierformat und Ausrichtung' +
        ' bestimmt die Seitenvorlage des Dokuments.'
      WordWrap = True
    end
    object edtPages: TEdit
      Left = 72
      Top = 24
      Width = 160
      Height = 23
      TabOrder = 0
      TextHint = 'leer = aus dem Druckdialog'
    end
    object btnPrint: TButton
      Left = 16
      Top = 60
      Width = 200
      Height = 25
      Caption = 'Drucken '#8230
      TabOrder = 1
      OnClick = btnPrintClick
    end
  end
  object grpMore: TGroupBox
    Left = 8
    Top = 480
    Width = 784
    Height = 140
    Caption = 'Suchen, Bilder, Lesen, Filter'
    TabOrder = 4
    object lblSearch: TLabel
      Left = 16
      Top = 28
      Width = 42
      Height = 15
      Caption = 'Suchen:'
    end
    object lblReplace: TLabel
      Left = 220
      Top = 28
      Width = 53
      Height = 15
      Caption = 'Ersetzen:'
    end
    object lblBookmark: TLabel
      Left = 16
      Top = 68
      Width = 62
      Height = 15
      Caption = 'Textmarke:'
    end
    object lblFilter: TLabel
      Left = 16
      Top = 106
      Width = 30
      Height = 15
      Caption = 'Filter:'
    end
    object lblMoreHint: TLabel
      Left = 468
      Top = 100
      Width = 300
      Height = 32
      AutoSize = False
      Caption =
        'Ersetzen erfasst auch Kopf-, Fu'#223'zeile und Rahmen. Lesen geht nur' +
        ' bei Textmarken, die Text umspannen.'
      WordWrap = True
    end
    object edtSearch: TEdit
      Left = 66
      Top = 24
      Width = 140
      Height = 23
      TabOrder = 0
    end
    object edtReplace: TEdit
      Left = 280
      Top = 24
      Width = 140
      Height = 23
      TabOrder = 1
    end
    object chkSearchCase: TCheckBox
      Left = 436
      Top = 26
      Width = 160
      Height = 17
      Caption = 'Gro'#223'-/Kleinschreibung'
      TabOrder = 2
    end
    object btnReplaceAll: TButton
      Left = 604
      Top = 22
      Width = 164
      Height = 25
      Caption = 'Alle ersetzen'
      TabOrder = 3
      OnClick = btnReplaceAllClick
    end
    object edtBookmark: TEdit
      Left = 86
      Top = 64
      Width = 120
      Height = 23
      TabOrder = 4
      TextHint = 'Name der Textmarke'
    end
    object btnReadBookmark: TButton
      Left = 220
      Top = 62
      Width = 160
      Height = 25
      Caption = 'Textmarke lesen'
      TabOrder = 5
      OnClick = btnReadBookmarkClick
    end
    object btnInsertImage: TButton
      Left = 392
      Top = 62
      Width = 200
      Height = 25
      Caption = 'Bild an Textmarke einf'#252'gen '#8230
      TabOrder = 6
      OnClick = btnInsertImageClick
    end
    object btnReadTable: TButton
      Left = 604
      Top = 62
      Width = 164
      Height = 25
      Caption = 'Tabelle lesen'
      TabOrder = 7
      OnClick = btnReadTableClick
    end
    object cboFilter: TComboBox
      Left = 66
      Top = 102
      Width = 140
      Height = 23
      Style = csDropDownList
      TabOrder = 8
    end
    object btnSaveFiltered: TButton
      Left = 220
      Top = 100
      Width = 220
      Height = 25
      Caption = 'Kopie speichern als '#8230
      TabOrder = 9
      OnClick = btnSaveFilteredClick
    end
  end
  object memLog: TMemo
    Left = 8
    Top = 628
    Width = 784
    Height = 112
    Anchors = [akLeft, akTop, akRight, akBottom]
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 5
  end
  object dlgImage: TOpenDialog
    Filter = 'Bilder (*.png;*.jpg;*.jpeg;*.gif;*.bmp)|*.png;*.jpg;*.jpeg;*.gif;*.bmp|Alle Dateien (*.*)|*.*'
    Options = [ofHideReadOnly, ofFileMustExist, ofEnableSizing]
    Left = 616
    Top = 636
  end
  object dlgSaveFiltered: TSaveDialog
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 616
    Top = 684
  end
  object dlgOpen: TOpenDialog
    Filter =
      'Writer-Dokumente (*.odt;*.ott;*.doc;*.docx)|*.odt;*.ott;*.doc;*.d' +
      'ocx|Alle Dateien (*.*)|*.*'
    Options = [ofHideReadOnly, ofFileMustExist, ofEnableSizing]
    Left = 704
    Top = 636
  end
  object dlgSave: TSaveDialog
    DefaultExt = 'odt'
    Filter = 'Writer-Dokument (*.odt)|*.odt|Alle Dateien (*.*)|*.*'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 704
    Top = 684
  end
  object dlgSavePdf: TSaveDialog
    DefaultExt = 'pdf'
    Filter = 'PDF (*.pdf)|*.pdf'
    Options = [ofOverwritePrompt, ofHideReadOnly, ofEnableSizing]
    Left = 744
    Top = 536
  end
  object dlgPrint: TPrintDialog
    MaxPage = 9999
    MinPage = 1
    Options = [poPageNums]
    Left = 744
    Top = 488
  end
end
