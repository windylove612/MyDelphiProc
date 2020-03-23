object Form1: TForm1
  Left = 318
  Top = 233
  Width = 695
  Height = 363
  Caption = 'excel'#23548#20837
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
  object RzLabel1: TRzLabel
    Left = 32
    Top = 24
    Width = 33
    Height = 13
    Caption = #34920#21517
  end
  object edtTBL: TRzEdit
    Left = 70
    Top = 22
    Width = 153
    Height = 21
    Text = 'HYK_LPFFTMPJL'
    FrameController = RzFrameController
    TabOrder = 0
  end
  object BtnSJDR: TRzButton
    Left = 18
    Top = 69
    Width = 92
    Height = 24
    FrameColor = clScrollBar
    Caption = #23548#20837'excel'#25968#25454
    Color = 14737632
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    HotTrack = True
    ParentFont = False
    TabOrder = 1
    OnClick = BtnSJDRClick
  end
  object RzDBGrid1: TRzDBGrid
    Left = 328
    Top = 8
    Width = 345
    Height = 305
    DataSource = dsExcel
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
      end>
  end
  object RzButton1: TRzButton
    Left = 122
    Top = 69
    Width = 103
    Height = 24
    FrameColor = clScrollBar
    Caption = #30830#35748#20889#20837#25968#25454#24211
    Color = 14737632
    Font.Charset = GB2312_CHARSET
    Font.Color = clWindowText
    Font.Height = -12
    Font.Name = #23435#20307
    Font.Style = []
    HotTrack = True
    ParentFont = False
    TabOrder = 3
    OnClick = RzButton1Click
  end
  object RzFrameController: TRzFrameController
    FlatButtonColor = clSkyBlue
    FrameColor = clSkyBlue
    FrameVisible = True
    Left = 126
    Top = 243
  end
  object qryExcel: TQuery
    DatabaseName = 'EXCEL'
    Left = 384
    Top = 88
  end
  object tblExcel: TTable
    Left = 152
    Top = 152
    object tblExcelHYK_NO: TStringField
      DisplayLabel = #20250#21592#21345#21495
      FieldName = 'HYK_NO'
    end
  end
  object dsExcel: TDataSource
    DataSet = tblExcel
    Left = 384
    Top = 200
  end
  object Database1: TDatabase
    AliasName = 'BF'
    DatabaseName = 'EXCEL'
    LoginPrompt = False
    SessionName = 'Default'
    Left = 224
    Top = 224
  end
  object OpenDialog1: TOpenDialog
    Left = 472
    Top = 144
  end
  object qryConnect: TQuery
    DatabaseName = 'EXCEL'
    SQL.Strings = (
      'select 1 from dual')
    Left = 264
    Top = 112
  end
  object qryID: TQuery
    DatabaseName = 'EXCEL'
    Left = 368
    Top = 160
  end
end
