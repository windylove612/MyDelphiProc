object Form1: TForm1
  Left = 416
  Top = 242
  Width = 584
  Height = 329
  Caption = 'Form1'
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object RzPanel1: TRzPanel
    Left = 0
    Top = 0
    Width = 568
    Height = 291
    Align = alClient
    TabOrder = 0
    object RzLabel1: TRzLabel
      Left = 80
      Top = 48
      Width = 33
      Height = 13
      AutoSize = False
      Caption = #26263#30721
    end
    object RzLabel2: TRzLabel
      Left = 81
      Top = 94
      Width = 31
      Height = 15
      Align = alCustom
      AutoSize = False
      Caption = #26126#30721
    end
    object RzLabel3: TRzLabel
      Left = 64
      Top = 136
      Width = 57
      Height = 13
      AutoSize = False
      Caption = #25552#31034#20449#24687
    end
    object edtMM: TRzEdit
      Left = 120
      Top = 45
      Width = 233
      Height = 21
      TabOrder = 0
    end
    object edtRL: TRzEdit
      Left = 120
      Top = 92
      Width = 233
      Height = 21
      Color = clInfoBk
      ReadOnly = True
      TabOrder = 1
    end
    object RzMemo1: TRzMemo
      Left = 119
      Top = 136
      Width = 273
      Height = 129
      TabOrder = 2
    end
    object RzButton1: TRzButton
      Left = 376
      Top = 56
      Caption = #35299#23494
      TabOrder = 3
      OnClick = RzButton1Click
    end
  end
end
