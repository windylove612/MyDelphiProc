object fmProcCSRQ_SFZ: TfmProcCSRQ_SFZ
  Left = 291
  Top = 251
  Width = 759
  Height = 326
  Caption = #36523#20221#35777#29983#25104#20986#29983#26085#26399
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
  object LabMsg: TLabel
    Left = 96
    Top = 168
    Width = 185
    Height = 12
    AutoSize = False
  end
  object Label1: TLabel
    Left = 400
    Top = 248
    Width = 185
    Height = 12
    AutoSize = False
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
        'select G.HYID,G.SFZBH'
        'from   HYK_GRXX G  '
        'where G.SFZBH is not null'
        'and G.CSRQ is  null'
        'and (length(trim(G.SFZBH))=18 or length(trim(G.SFZBH))=15)'
        'order by G.HYID')
      TabOrder = 0
    end
  end
  object pb1: TProgressBar
    Left = 13
    Top = 192
    Width = 383
    Height = 16
    TabOrder = 1
  end
  object Button1: TButton
    Left = 156
    Top = 230
    Width = 75
    Height = 25
    Caption = #25191#34892
    TabOrder = 2
    OnClick = Button1Click
  end
  object DBGrid1: TRzDBGrid
    Left = 400
    Top = 8
    Height = 233
    DataSource = dsSFZ
    TabOrder = 3
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'MS Sans Serif'
    TitleFont.Style = []
    Columns = <
      item
        Expanded = False
        FieldName = 'SFZBH'
        Width = 217
        Visible = True
      end>
  end
  object RzButton1: TRzButton
    Left = 624
    Top = 248
    Caption = #29983#25104#25991#26412
    TabOrder = 4
    OnClick = RzButton1Click
  end
  object Database1: TDatabase
    AliasName = 'BF'
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
    Left = 224
    Top = 112
  end
  object tmpQry: TQuery
    DatabaseName = 'CSCRM'
    Left = 304
    Top = 88
  end
  object qrySave: TQuery
    DatabaseName = 'CSCRM'
    Left = 432
    Top = 56
  end
  object tblSFZ: TTable
    Left = 520
    Top = 80
    object tblSFZSFZBH: TStringField
      FieldName = 'SFZBH'
    end
  end
  object dsSFZ: TDataSource
    DataSet = tblSFZ
    OnDataChange = dsSFZDataChange
    Left = 584
    Top = 128
  end
  object SaveDialog: TSaveDialog
    FileName = 'sfzbh_error.TXT'
    Left = 464
    Top = 158
  end
end
