object Form1: TForm1
  Left = 283
  Top = 111
  Width = 764
  Height = 511
  Caption = #35302#21457#22120#33050#26412#29983#25104#22120
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
    Width = 281
    Height = 473
    Align = alLeft
    TabOrder = 0
    object RzLabel2: TRzLabel
      Left = 16
      Top = 50
      Width = 33
      Height = 13
      AutoSize = False
      Caption = #29992#25143
    end
    object RzLabel3: TRzLabel
      Left = 136
      Top = 50
      Width = 33
      Height = 13
      AutoSize = False
      Caption = #34920
    end
    object RzButton1: TRzButton
      Left = 16
      Top = 16
      Width = 89
      Caption = #33719#21462#29983#25104#21015#34920
      HotTrack = True
      TabOrder = 0
      OnClick = RzButton1Click
    end
    object RzButton2: TRzButton
      Left = 112
      Top = 16
      Caption = #29983#25104#33050#26412
      HotTrack = True
      TabOrder = 1
      OnClick = RzButton2Click
    end
    object RzDBGrid1: TRzDBGrid
      Left = 16
      Top = 72
      Width = 217
      Height = 385
      DataSource = dsLB
      TabOrder = 2
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'MS Sans Serif'
      TitleFont.Style = []
      Columns = <
        item
          Expanded = False
          FieldName = 'TBLNAME'
          Visible = True
        end>
    end
    object edtUSER: TRzEdit
      Left = 48
      Top = 46
      Width = 81
      Height = 21
      TabOrder = 3
    end
    object edtBM: TRzEdit
      Left = 168
      Top = 46
      Width = 105
      Height = 21
      TabOrder = 4
    end
    object RzButton3: TRzButton
      Left = 192
      Top = 16
      Caption = #29983#25104#33050#26412'('#36755#20837')'
      HotTrack = True
      TabOrder = 5
      OnClick = RzButton3Click
    end
  end
  object RzMemo1: TRzMemo
    Left = 288
    Top = 48
    Width = 457
    Height = 417
    Align = alCustom
    TabOrder = 1
  end
  object RzPanel2: TRzPanel
    Left = 289
    Top = 8
    Width = 456
    Height = 33
    Align = alCustom
    TabOrder = 2
    object RzLabel1: TRzLabel
      Left = 6
      Top = 10
      Width = 105
      Height = 13
      AutoSize = False
      Caption = #29983#25104#33050#26412#39044#35272
    end
  end
  object Query1: TQuery
    Left = 72
    Top = 104
  end
  object tblLB: TTable
    Left = 88
    Top = 176
    object tblLBTBLNAME: TStringField
      DisplayLabel = #34920#21517
      FieldName = 'TBLNAME'
    end
  end
  object dsLB: TDataSource
    DataSet = tblLB
    Left = 160
    Top = 176
  end
  object OpenDialog1: TOpenDialog
    Filter = '.*'
    Left = 200
    Top = 232
  end
end
