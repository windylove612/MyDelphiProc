object Form1: TForm1
  Left = 404
  Top = 289
  Width = 418
  Height = 191
  Caption = #33719#21462#39033#30446#30721
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object RzLabel1: TRzLabel
    Left = 24
    Top = 120
    Width = 73
    Height = 13
    AutoSize = False
    Caption = #25805#20316#21592#20195#30721
  end
  object RzEdit1: TRzEdit
    Left = 144
    Top = 77
    Width = 241
    Height = 21
    TabOrder = 0
  end
  object RzButton1: TRzButton
    Left = 48
    Top = 74
    Caption = #33719#21462#39033#30446#30721
    HotTrack = True
    TabOrder = 1
    OnClick = RzButton1Click
  end
  object rgLX: TRzRadioGroup
    Left = 56
    Top = 16
    Width = 153
    Height = 41
    Caption = #25968#25454#24211#31867#22411
    Color = clWindow
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      'sybase'
      'oracle')
    TabOrder = 2
  end
  object RzButton2: TRzButton
    Left = 224
    Top = 24
    Width = 105
    Caption = #35835#21462#37197#32622#25991#20214
    HotTrack = True
    TabOrder = 3
    OnClick = RzButton2Click
  end
  object RzEdit2: TRzEdit
    Left = 104
    Top = 117
    Width = 81
    Height = 21
    TabOrder = 4
  end
  object RzButton3: TRzButton
    Left = 192
    Top = 114
    Caption = #25805#20316#21592#23494#30721
    HotTrack = True
    TabOrder = 5
    OnClick = RzButton3Click
  end
  object RzEdit3: TRzEdit
    Left = 280
    Top = 117
    Width = 81
    Height = 21
    TabOrder = 6
  end
  object Database1: TDatabase
    DatabaseName = 'CSCRM'
    LoginPrompt = False
    Params.Strings = (
      'SERVER NAME=BENDI'
      'USER NAME=BFCRM8'
      'NET PROTOCOL=TNS'
      'OPEN MODE=READ/WRITE'
      'SCHEMA CACHE SIZE=8'
      'LANGDRIVER='
      'SQLQRYMODE='
      'SQLPASSTHRU MODE=SHARED AUTOCOMMIT'
      'SCHEMA CACHE TIME=-1'
      'MAX ROWS=-1'
      'BATCH COUNT=200'
      'ENABLE SCHEMA CACHE=FALSE'
      'SCHEMA CACHE DIR='
      'ENABLE BCD=FALSE'
      'ENABLE INTEGERS=TRUE'
      'LIST SYNONYMS=NONE'
      'ROWSET SIZE=20'
      'BLOBS TO CACHE=64'
      'BLOB SIZE=32'
      'OBJECT MODE=TRUE'
      'PASSWORD=DHHZDHHZ')
    SessionName = 'Default'
    Left = 336
    Top = 32
  end
  object Query1: TQuery
    DatabaseName = 'CSCRM'
    Left = 16
    Top = 32
  end
end
