object ProcZK: TProcZK
  Left = 225
  Top = 90
  Width = 950
  Height = 548
  Caption = #21046#21345
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
    Width = 934
    Height = 81
    Align = alTop
    TabOrder = 0
    object RzLabel2: TRzLabel
      Left = 144
      Top = 16
      Width = 409
      Height = 13
      AutoSize = False
      Caption = #25991#26723#26684#24335#20026#65306#21345#21495#65292#30913#36947#20869#23481
    end
    object Label7: TLabel
      Left = 96
      Top = 51
      Width = 129
      Height = 13
      AutoSize = False
      Caption = #20889#21345#22120#36890#35759#31471#21475#35774#32622':'
    end
    object lblPortSetting: TLabel
      Left = 246
      Top = 51
      Width = 84
      Height = 12
      Caption = 'lblPortSetting'
      Font.Charset = GB2312_CHARSET
      Font.Color = clBlue
      Font.Height = -12
      Font.Name = #23435#20307
      Font.Style = []
      ParentFont = False
    end
    object Bevel1: TBevel
      Left = 225
      Top = 67
      Width = 288
      Height = 12
      Shape = bsTopLine
      Style = bsRaised
    end
    object Label10: TLabel
      Left = 736
      Top = 5
      Width = 72
      Height = 13
      Caption = #25104#21151#20889#21345#25968#65306
    end
    object Label11: TLabel
      Left = 724
      Top = 52
      Width = 84
      Height = 13
      Caption = #21097#20313#26410#20889#21345#25968#65306
    end
    object LabXK: TLabel
      Left = 808
      Top = 5
      Width = 42
      Height = 12
      AutoSize = False
    end
    object LabSY: TLabel
      Left = 808
      Top = 51
      Width = 42
      Height = 12
      AutoSize = False
    end
    object Label12: TLabel
      Left = 724
      Top = 28
      Width = 84
      Height = 13
      Caption = #19981#25104#21151#20889#21345#25968#65306
    end
    object LabSBSL: TLabel
      Left = 808
      Top = 28
      Width = 42
      Height = 12
      AutoSize = False
    end
    object RzButton1: TRzButton
      Left = 552
      Top = 43
      Caption = #20889#21345
      HotTrack = True
      TabOrder = 0
      OnClick = RzButton1Click
    end
    object RzButton2: TRzButton
      Left = 640
      Top = 43
      Caption = #35835#21345
      HotTrack = True
      TabOrder = 1
      OnClick = RzButton2Click
    end
    object RzButton3: TRzButton
      Left = 8
      Top = 8
      Width = 129
      Caption = #35835#21462#38656#21046#21345#25991#26723
      HotTrack = True
      TabOrder = 2
      OnClick = RzButton3Click
    end
    object btnSetup: TBmpBtn
      Left = 10
      Top = 45
      Width = 81
      Caption = #35774#32622#31471#21475
      HotTrack = True
      TabOrder = 3
      OnClick = btnSetupClick
      Kind = bkSetup
    end
  end
  object RzDBGrid1: TRzDBGrid
    Left = 0
    Top = 81
    Width = 409
    Height = 429
    Align = alLeft
    DataSource = dsMain
    ReadOnly = True
    TabOrder = 1
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
        Width = 240
        Visible = True
      end>
  end
  object RzPanel2: TRzPanel
    Left = 409
    Top = 81
    Width = 525
    Height = 429
    Align = alClient
    TabOrder = 2
    object RzLabel1: TRzLabel
      Left = 16
      Top = 40
      Width = 106
      Height = 24
      AutoSize = False
      Caption = #24403#21069#21345#21495#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object RzLabel3: TRzLabel
      Left = 16
      Top = 160
      Width = 169
      Height = 24
      AutoSize = False
      Caption = #24403#21069#30913#36947#20869#23481#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblKH: TRzLabel
      Left = 24
      Top = 80
      Width = 345
      Height = 24
      AutoSize = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblCD: TRzLabel
      Left = 24
      Top = 208
      Width = 417
      Height = 20
      AutoSize = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object RzLabel4: TRzLabel
      Left = 17
      Top = 256
      Width = 106
      Height = 24
      AutoSize = False
      Caption = #35835#20986#21345#21495#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblDCKH: TRzLabel
      Left = 25
      Top = 296
      Width = 345
      Height = 24
      AutoSize = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object RzLabel6: TRzLabel
      Left = 17
      Top = 344
      Width = 106
      Height = 24
      AutoSize = False
      Caption = #35835#20986#30913#36947#65306
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
    object lblDCCD: TRzLabel
      Left = 25
      Top = 384
      Width = 345
      Height = 24
      AutoSize = False
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -16
      Font.Name = 'MS Sans Serif'
      Font.Style = [fsBold]
      ParentFont = False
    end
  end
  object tblMain: TTable
    Left = 96
    Top = 216
    object tblMainHYK_NO: TStringField
      DisplayLabel = #20250#21592#21345#21495
      FieldName = 'HYK_NO'
    end
    object tblMainCDNR: TStringField
      DisplayLabel = #30913#36947#20869#23481
      FieldName = 'CDNR'
      Size = 60
    end
  end
  object dsMain: TDataSource
    DataSet = tblMain
    OnDataChange = dsMainDataChange
    Left = 144
    Top = 152
  end
  object OpenDialog1: TOpenDialog
    Filter = '.*'
    Left = 200
    Top = 232
  end
end
