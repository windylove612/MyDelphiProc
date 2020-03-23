object fmProcDecrypt: TfmProcDecrypt
  Left = 297
  Top = 167
  Width = 842
  Height = 403
  Caption = #25209#37327#35299#23494
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
  object Label1: TLabel
    Left = 408
    Top = 344
    Width = 185
    Height = 12
    AutoSize = False
  end
  object RzLabel1: TRzLabel
    Left = 24
    Top = 208
    Width = 361
    Height = 26
    AutoSize = False
    Caption = #29983#25104#26465#20214#21487#20197#33258#24049#20889#65292#38754#20540#21345#30340#35821#21477#25105#25918#22312#19979#38754#65292#13#25353#26377#20010#26631#35760#26469#21306#20998#30340#65292#36825#26679#20320#22909#23548
  end
  object RzLabel2: TRzLabel
    Left = 104
    Top = 168
    Width = 281
    Height = 13
    AutoSize = False
    Caption = #25805#20316#27493#39588#65306#20808#28857#25191#34892#65292#22312#28857#29983#25104#25991#26412
  end
  object GroupBox1: TGroupBox
    Left = 8
    Top = 8
    Width = 380
    Height = 136
    Caption = #29983#25104#26465#20214
    TabOrder = 0
    object Memo1: TMemo
      Left = 2
      Top = 15
      Width = 376
      Height = 119
      Align = alClient
      Lines.Strings = (
        'select HYK_NO,CDNR'
        'from HYK_HYXX'
        'where HYK_NO>='#39'00000000'#39
        'order by HYK_NO')
      TabOrder = 0
    end
  end
  object Button1: TButton
    Left = 12
    Top = 158
    Width = 75
    Height = 25
    Caption = #25191#34892
    TabOrder = 1
    OnClick = Button1Click
  end
  object DBGrid1: TRzDBGrid
    Left = 400
    Top = 8
    Width = 409
    Height = 321
    DataSource = dsSFZ
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'HYK_NO'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'CDNR'
        Width = 251
        Visible = True
      end>
  end
  object RzButton1: TRzButton
    Left = 632
    Top = 336
    Caption = #29983#25104#25991#26412
    TabOrder = 3
    OnClick = RzButton1Click
  end
  object RzMemo1: TRzMemo
    Left = 16
    Top = 248
    Width = 361
    Height = 89
    Lines.Strings = (
      'select X.HYK_NO,X.CDNR'
      'from HYK_HYXX X,HYKDEF F'
      'where X.HYKTYPE=F.HYKTYPE and F.BJ_CZK=1 and '
      'X.HYK_NO>='#39'00000000'#39
      'order by X.HYK_NO')
    TabOrder = 4
  end
  object Database1: TDatabase
    DatabaseName = 'CSCRM'
    DriverName = 'ORACLE'
    LoginPrompt = False
    SessionName = 'Default'
    Left = 232
    Top = 88
  end
  object tmpQry: TQuery
    DatabaseName = 'CSCRM'
    Left = 304
    Top = 88
  end
  object qrySave: TQuery
    DatabaseName = 'CSCRM'
    Left = 440
    Top = 112
  end
  object tblSFZ: TTable
    Left = 520
    Top = 80
    object tblSFZHYK_NO: TStringField
      DisplayLabel = #20250#21592#21345#21495
      FieldName = 'HYK_NO'
    end
    object tblSFZCDNR: TStringField
      DisplayLabel = #30913#36947#20869#23481
      FieldName = 'CDNR'
    end
  end
  object dsSFZ: TDataSource
    DataSet = tblSFZ
    Left = 584
    Top = 128
  end
  object SaveDialog: TSaveDialog
    FileName = 'sfzbh_error.TXT'
    Left = 472
    Top = 182
  end
end
