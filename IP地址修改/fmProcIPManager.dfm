object fmIPManager: TfmIPManager
  Left = 253
  Top = 119
  Width = 905
  Height = 476
  Caption = 'IP'#22320#22336#31649#29702
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
    Width = 889
    Height = 177
    Align = alTop
    BorderColor = clSkyBlue
    BorderShadow = clSkyBlue
    Color = clWindow
    FlatColor = clSkyBlue
    GridColor = clSkyBlue
    TabOrder = 0
    object RzLabel1: TRzLabel
      Left = 16
      Top = 24
      Width = 34
      Height = 13
      AutoSize = False
      Caption = #21517#31216
    end
    object RzLabel2: TRzLabel
      Left = 16
      Top = 57
      Width = 33
      Height = 13
      AutoSize = False
      Caption = #32593#21345
    end
    object RzLabel3: TRzLabel
      Left = 16
      Top = 89
      Width = 43
      Height = 13
      AutoSize = False
      Caption = 'IP'#22320#22336
    end
    object RzLabel4: TRzLabel
      Left = 256
      Top = 89
      Width = 61
      Height = 13
      AutoSize = False
      Caption = #23376#32593#25513#30721
    end
    object RzLabel5: TRzLabel
      Left = 12
      Top = 121
      Width = 61
      Height = 13
      AutoSize = False
      Caption = #40664#35748#32593#20851
    end
    object RzLabel6: TRzLabel
      Left = 256
      Top = 121
      Width = 59
      Height = 13
      AutoSize = False
      Caption = #39318#36873'DNS'
    end
    object RzLabel7: TRzLabel
      Left = 8
      Top = 153
      Width = 59
      Height = 13
      AutoSize = False
      Caption = #22791#29992'DNS'
    end
    object edtMC: TRzEdit
      Left = 72
      Top = 22
      Width = 169
      Height = 21
      FrameColor = clSkyBlue
      FrameHotColor = clSkyBlue
      FrameVisible = True
      TabOrder = 0
    end
    object ComWK: TRzComboBox
      Left = 72
      Top = 54
      Width = 257
      Height = 21
      Ctl3D = False
      FrameColor = clSkyBlue
      FrameVisible = True
      ItemHeight = 13
      ParentCtl3D = False
      TabOrder = 1
    end
    object edtIP: TRzEdit
      Left = 72
      Top = 86
      Width = 169
      Height = 21
      FrameColor = clSkyBlue
      FrameHotColor = clSkyBlue
      FrameVisible = True
      TabOrder = 2
    end
    object edtYM: TRzEdit
      Left = 328
      Top = 86
      Width = 169
      Height = 21
      FrameColor = clSkyBlue
      FrameHotColor = clSkyBlue
      FrameVisible = True
      TabOrder = 3
    end
    object edtWG: TRzEdit
      Left = 72
      Top = 118
      Width = 169
      Height = 21
      FrameColor = clSkyBlue
      FrameHotColor = clSkyBlue
      FrameVisible = True
      TabOrder = 4
    end
    object edtMS: TRzEdit
      Left = 328
      Top = 118
      Width = 169
      Height = 21
      FrameColor = clSkyBlue
      FrameHotColor = clSkyBlue
      FrameVisible = True
      TabOrder = 5
    end
    object BmpBtn6: TBmpBtn
      Left = 512
      Top = 141
      Caption = #24212#29992
      HotTrack = True
      TabOrder = 6
      OnClick = BmpBtn6Click
      Kind = bkExec
    end
    object edtBY: TRzEdit
      Left = 72
      Top = 150
      Width = 169
      Height = 21
      FrameColor = clSkyBlue
      FrameHotColor = clSkyBlue
      FrameVisible = True
      TabOrder = 7
    end
  end
  object RzDBGrid1: TRzDBGrid
    Left = 0
    Top = 177
    Width = 889
    Height = 220
    Align = alClient
    DataSource = dsDZ
    ReadOnly = True
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    FrameColor = clSkyBlue
    FrameHotColor = clSkyBlue
    FrameVisible = True
    LineColor = clSkyBlue
    Columns = <
      item
        Expanded = False
        FieldName = 'MC'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'WKMC'
        Width = 118
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'IP'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'YM'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'WG'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'DNS'
        Visible = True
      end
      item
        Expanded = False
        FieldName = 'BYDNS'
        Visible = True
      end>
  end
  object RzPanel2: TRzPanel
    Left = 0
    Top = 397
    Width = 889
    Height = 41
    Align = alBottom
    BorderColor = clSkyBlue
    BorderShadow = clSkyBlue
    Color = clWindow
    FlatColor = clSkyBlue
    GradientColorStop = clSkyBlue
    GridColor = clSkyBlue
    TabOrder = 2
    object btnAdd: TBmpBtn
      Left = 185
      Top = 8
      HotTrack = True
      TabOrder = 0
      OnClick = btnAddClick
      Kind = bkAdd
    end
    object btnXG: TBmpBtn
      Left = 273
      Top = 8
      HotTrack = True
      TabOrder = 1
      OnClick = btnXGClick
      Kind = bkModify
    end
    object btnDel: TBmpBtn
      Left = 361
      Top = 8
      HotTrack = True
      TabOrder = 2
      OnClick = btnDelClick
      Kind = bkDelete
    end
    object btnSave: TBmpBtn
      Left = 449
      Top = 8
      HotTrack = True
      TabOrder = 3
      OnClick = btnSaveClick
      Kind = bkSave
    end
    object btnCancel: TBmpBtn
      Left = 537
      Top = 8
      HotTrack = True
      TabOrder = 4
      OnClick = btnCancelClick
      Kind = bkCancel
    end
    object btnClose: TBmpBtn
      Left = 625
      Top = 8
      HotTrack = True
      TabOrder = 5
      OnClick = btnCloseClick
      Kind = bkExit
    end
  end
  object tblDZ: TTable
    Left = 240
    Top = 264
    object tblDZMC: TStringField
      DisplayLabel = #21517#31216
      FieldName = 'MC'
    end
    object tblDZWKMC: TStringField
      DisplayLabel = #32593#21345
      FieldName = 'WKMC'
      Size = 40
    end
    object tblDZIP: TStringField
      DisplayLabel = 'IP'#22320#22336
      FieldName = 'IP'
    end
    object tblDZYM: TStringField
      DisplayLabel = #23376#32593#25513#30721
      FieldName = 'YM'
    end
    object tblDZWG: TStringField
      DisplayLabel = #32593#20851
      FieldName = 'WG'
    end
    object tblDZDNS: TStringField
      DisplayLabel = #39318#36873'DNS'
      FieldName = 'DNS'
    end
    object tblDZBYDNS: TStringField
      DisplayLabel = #22791#29992'DNS'
      FieldName = 'BYDNS'
    end
  end
  object dsDZ: TDataSource
    DataSet = tblDZ
    OnDataChange = dsDZDataChange
    Left = 320
    Top = 240
  end
end
