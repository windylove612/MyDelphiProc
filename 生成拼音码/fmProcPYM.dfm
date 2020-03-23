object Form1: TForm1
  Left = 388
  Top = 259
  Width = 544
  Height = 323
  Caption = #29983#25104#25340#38899#30721
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object RzPanel1: TRzPanel
    Left = 0
    Top = 0
    Width = 528
    Height = 285
    Align = alClient
    TabOrder = 0
    object RzLabel1: TRzLabel
      Left = 32
      Top = 32
      Width = 57
      Height = 13
      AutoSize = False
      Caption = #26356#26032#34920#21517
    end
    object RzLabel2: TRzLabel
      Left = 32
      Top = 67
      Width = 49
      Height = 13
      AutoSize = False
      Caption = #20027#38190#21517
    end
    object RzLabel3: TRzLabel
      Left = 264
      Top = 68
      Width = 129
      Height = 13
      AutoSize = False
      Caption = #22810#20010#20027#38190#29992#36887#21495#38548#24320
    end
    object RzLabel4: TRzLabel
      Left = 16
      Top = 132
      Width = 73
      Height = 13
      AutoSize = False
      Caption = #26356#26032#23383#27573#21517
    end
    object LabMsg: TLabel
      Left = 144
      Top = 200
      Width = 185
      Height = 12
      AutoSize = False
    end
    object RzLabel5: TRzLabel
      Left = 16
      Top = 100
      Width = 73
      Height = 13
      AutoSize = False
      Caption = #36716#25442#23383#27573#21517
    end
    object RzLabel6: TRzLabel
      Left = 4
      Top = 164
      Width = 81
      Height = 13
      AutoSize = False
      Caption = #29305#27530#38480#21046#26465#20214
    end
    object RzLabel7: TRzLabel
      Left = 328
      Top = 164
      Width = 121
      Height = 13
      AutoSize = False
      Caption = #20363#65306'SP_ID>0'
    end
    object RzLabel8: TRzLabel
      Left = 216
      Top = 132
      Width = 145
      Height = 13
      AutoSize = False
      Caption = #22810#20010#23383#27573#29992#36887#21495#38548#24320
    end
    object edtTBL: TRzEdit
      Left = 88
      Top = 31
      Width = 121
      Height = 21
      FrameColor = clSkyBlue
      FrameVisible = True
      TabOrder = 0
    end
    object edtKEY: TRzEdit
      Left = 88
      Top = 65
      Width = 169
      Height = 21
      FrameColor = clSkyBlue
      FrameVisible = True
      TabOrder = 1
    end
    object edtZD: TRzEdit
      Left = 88
      Top = 130
      Width = 121
      Height = 21
      FrameColor = clSkyBlue
      FrameVisible = True
      TabOrder = 2
    end
    object btnOK: TRzButton
      Left = 160
      Top = 248
      Caption = #25191#34892
      HotTrack = True
      TabOrder = 3
      OnClick = btnOKClick
    end
    object btnExit: TRzButton
      Left = 296
      Top = 248
      Caption = #36864#20986
      HotTrack = True
      TabOrder = 4
      OnClick = btnExitClick
    end
    object pb1: TProgressBar
      Left = 61
      Top = 224
      Width = 383
      Height = 16
      TabOrder = 5
    end
    object edtZH: TRzEdit
      Left = 88
      Top = 98
      Width = 121
      Height = 21
      FrameColor = clSkyBlue
      FrameVisible = True
      TabOrder = 6
    end
    object edtTJ: TRzEdit
      Left = 88
      Top = 162
      Width = 233
      Height = 21
      FrameColor = clSkyBlue
      FrameVisible = True
      TabOrder = 7
    end
  end
  object Database1: TDatabase
    AliasName = 'BF'
    DatabaseName = 'PYM'
    LoginPrompt = False
    SessionName = 'Default'
    Left = 440
    Top = 24
  end
  object tblPYM: TTable
    Left = 424
    Top = 112
  end
  object qryTmp: TQuery
    DatabaseName = 'PYM'
    Left = 480
    Top = 80
  end
end
