object Form1: TForm1
  Left = 399
  Top = 259
  Width = 538
  Height = 273
  Caption = #29983#25104#35760#36134#21306#38388
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
    Width = 522
    Height = 235
    Align = alClient
    Color = clWindow
    TabOrder = 0
    object RzLabel1: TRzLabel
      Left = 32
      Top = 131
      Width = 33
      Height = 13
      AutoSize = False
      Caption = #34920#21517
    end
    object RzLabel3: TRzLabel
      Left = 16
      Top = 165
      Width = 49
      Height = 13
      AutoSize = False
      Caption = #24320#22987#24180
    end
    object RzLabel4: TRzLabel
      Left = 208
      Top = 165
      Width = 49
      Height = 13
      AutoSize = False
      Caption = #32467#26463#24180
    end
    object lblZY: TRzLabel
      Left = 320
      Top = 32
      Width = 3
      Height = 13
    end
    object rgLX: TRzRadioGroup
      Left = 32
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
      TabOrder = 0
      OnClick = rgLXClick
    end
    object RzButton1: TRzButton
      Left = 200
      Top = 24
      Width = 105
      Caption = #35835#21462#37197#32622#25991#20214
      HotTrack = True
      TabOrder = 1
      OnClick = RzButton1Click
    end
    object edtBM: TRzEdit
      Left = 68
      Top = 128
      Width = 121
      Height = 21
      FrameVisible = True
      TabOrder = 2
    end
    object KSRQ: TRzEdit
      Left = 68
      Top = 162
      Width = 121
      Height = 21
      FrameVisible = True
      TabOrder = 3
    end
    object JSRQ: TRzEdit
      Left = 260
      Top = 162
      Width = 121
      Height = 21
      FrameVisible = True
      TabOrder = 4
    end
    object rgJZ: TRzRadioGroup
      Left = 32
      Top = 72
      Width = 153
      Height = 41
      Caption = #35760#36134#21306#38388#31867#22411
      Color = clWindow
      Columns = 2
      ItemIndex = 0
      Items.Strings = (
        #26376
        #23395)
      TabOrder = 5
      OnClick = rgJZClick
    end
    object RzButton2: TRzButton
      Left = 104
      Top = 200
      Width = 105
      Caption = #25191#34892
      HotTrack = True
      TabOrder = 6
      OnClick = RzButton2Click
    end
    object RzButton3: TRzButton
      Left = 256
      Top = 200
      Width = 105
      Caption = #36864#20986
      HotTrack = True
      TabOrder = 7
      OnClick = RzButton3Click
    end
  end
  object Database1: TDatabase
    DatabaseName = 'JZQJ'
    LoginPrompt = False
    SessionName = 'Default'
    Left = 336
    Top = 72
  end
  object qryTmp: TQuery
    DatabaseName = 'JZQJ'
    Left = 272
    Top = 88
  end
  object qryConnect: TQuery
    DatabaseName = 'JZQJ'
    SQL.Strings = (
      'select 1 from dual')
    Left = 392
    Top = 104
  end
end
