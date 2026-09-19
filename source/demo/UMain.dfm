object Form1: TForm1
  Left = 305
  Top = 186
  Caption = 'Form1'
  ClientHeight = 348
  ClientWidth = 535
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poScreenCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object ButtonOpen: TButton
    Left = 40
    Top = 24
    Width = 141
    Height = 25
    Caption = 'Dokument '#246'ffnen'
    TabOrder = 0
    OnClick = ButtonOpenClick
  end
  object Button2: TButton
    Left = 40
    Top = 136
    Width = 141
    Height = 25
    Caption = 'Dokument Drucken'
    TabOrder = 1
    OnClick = Button2Click
  end
  object ComboBox1: TComboBox
    Left = 196
    Top = 140
    Width = 93
    Height = 21
    Style = csDropDownList
    ItemHeight = 13
    ItemIndex = 1
    TabOrder = 2
    Text = 'A4'
    Items.Strings = (
      'A3'
      'A4'
      'A5'
      'B4'
      'B5'
      'LETTER'
      'LEGAL'
      'TABLOID')
  end
  object ButtonClose: TButton
    Left = 40
    Top = 56
    Width = 141
    Height = 25
    Caption = 'Dokument Schlie'#223'en'
    TabOrder = 3
    OnClick = ButtonCloseClick
  end
  object Button4: TButton
    Left = 40
    Top = 96
    Width = 141
    Height = 25
    Caption = 'Dokument Speichern'
    TabOrder = 4
    OnClick = Button4Click
  end
  object Button5: TButton
    Left = 40
    Top = 172
    Width = 141
    Height = 25
    Caption = 'Daten'#252'bergabe'
    TabOrder = 5
    OnClick = Button5Click
  end
  object Button_TableTest: TButton
    Left = 40
    Top = 232
    Width = 141
    Height = 25
    Caption = 'Tabellentest'
    TabOrder = 6
    OnClick = Button_TableTestClick
  end
  object OpenDialog1: TOpenDialog
    Left = 360
    Top = 28
  end
  object SaveDialog1: TSaveDialog
    Left = 396
    Top = 32
  end
  object PrintDialog1: TPrintDialog
    Left = 440
    Top = 28
  end
end
