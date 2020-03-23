object Form1: TForm1
  Left = 230
  Top = 158
  Width = 878
  Height = 442
  Caption = #38656#27714#20462#25913#35760#24405
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = FormCreate
  OnKeyDown = FormKeyDown
  PixelsPerInch = 96
  TextHeight = 13
  object RzPanel1: TRzPanel
    Left = 0
    Top = 0
    Width = 862
    Height = 404
    Align = alClient
    TabOrder = 0
    object RzLabel1: TRzLabel
      Left = 24
      Top = 24
      Width = 41
      Height = 13
      AutoSize = False
      Caption = #26085#26399
    end
    object RzLabel2: TRzLabel
      Left = 16
      Top = 96
      Width = 65
      Height = 13
      AutoSize = False
      Caption = #39033#30446#21517#31216
    end
    object RzLabel3: TRzLabel
      Left = 16
      Top = 62
      Width = 65
      Height = 13
      AutoSize = False
      Caption = #39033#30446#21015#34920
    end
    object RzLabel4: TRzLabel
      Left = 24
      Top = 128
      Width = 49
      Height = 13
      AutoSize = False
      Caption = #20869#23481
    end
    object RQ: TRzDateTimeEdit
      Left = 80
      Top = 20
      Width = 121
      Height = 21
      EditType = etDate
      TabOrder = 0
      OnKeyDown = RQKeyDown
    end
    object edtMC: TRzEdit
      Left = 80
      Top = 93
      Width = 265
      Height = 21
      TabOrder = 1
      OnKeyDown = edtMCKeyDown
    end
    object ComXM: TRzComboBox
      Left = 80
      Top = 59
      Width = 265
      Height = 21
      ItemHeight = 13
      TabOrder = 2
      OnChange = ComXMChange
      OnKeyDown = ComXMKeyDown
    end
    object edtNR: TRzMemo
      Left = 80
      Top = 128
      Width = 489
      Height = 201
      TabOrder = 3
      OnKeyDown = edtNRKeyDown
    end
    object btnSure: TRzButton
      Left = 184
      Top = 352
      Width = 81
      Caption = #30830#23450#28155#21152'(F1)'
      HotTrack = True
      TabOrder = 4
      OnClick = btnSureClick
    end
    object btnAdd: TRzButton
      Left = 360
      Top = 89
      Width = 113
      Caption = #28155#21152#39033#30446#21517#31216
      HotTrack = True
      TabOrder = 5
      OnClick = btnAddClick
    end
    object RzButton1: TRzButton
      Left = 480
      Top = 89
      Width = 113
      Caption = #26597#30475#20462#25913#20869#23481'(F7)'
      HotTrack = True
      TabOrder = 6
      OnClick = RzButton1Click
    end
    object RzButton2: TRzButton
      Left = 384
      Top = 352
      Width = 81
      Caption = #28165#38500#20869#23481'(F3)'
      HotTrack = True
      TabOrder = 7
      OnClick = RzButton2Click
    end
    object RzButton3: TRzButton
      Left = 576
      Top = 129
      Width = 113
      Caption = #20869#23481#23383#20307#35774#32622
      HotTrack = True
      TabOrder = 8
      OnClick = RzButton3Click
    end
    object btnShow: TRzButton
      Left = 583
      Top = 161
      Width = 65
      Height = 49
      Caption = #39044#35272
      Color = clWindow
      Font.Charset = GB2312_CHARSET
      Font.Color = clWindowText
      Font.Height = -14
      Font.Name = #23435#20307
      Font.Style = [fsBold]
      ParentFont = False
      TabOrder = 9
    end
    object RzButton4: TRzButton
      Left = 288
      Top = 352
      Caption = #37325#20889#20837'(F2)'
      HotTrack = True
      TabOrder = 10
      OnClick = RzButton4Click
    end
    object RzButton5: TRzButton
      Left = 576
      Top = 233
      Width = 113
      Caption = #21387#32553#31243#24207'(F4)'
      HotTrack = True
      TabOrder = 11
      OnClick = RzButton5Click
    end
    object RzButton6: TRzButton
      Left = 576
      Top = 265
      Width = 113
      Caption = 'sqlmonitor(F5)'
      HotTrack = True
      TabOrder = 12
      OnClick = RzButton6Click
    end
    object RzButton7: TRzButton
      Left = 576
      Top = 297
      Width = 113
      Caption = #25171#24320#25991#20214#30446#24405'(F6)'
      HotTrack = True
      TabOrder = 13
      OnClick = RzButton7Click
    end
    object RzButton8: TRzButton
      Left = 704
      Top = 233
      Width = 113
      Caption = #30913#36947#20869#23481#20114#36716'(F8)'
      HotTrack = True
      TabOrder = 14
      OnClick = RzButton8Click
    end
    object RzButton9: TRzButton
      Left = 704
      Top = 265
      Width = 113
      Caption = #19975#33021#21152#35299#23494'(F9'
      HotTrack = True
      TabOrder = 15
      OnClick = RzButton9Click
    end
    object RzButton10: TRzButton
      Left = 704
      Top = 297
      Width = 113
      Caption = #21152#23494'new(F10)'
      HotTrack = True
      TabOrder = 16
      OnClick = RzButton10Click
    end
  end
  object FontDialog1: TFontDialog
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = []
    Left = 328
    Top = 176
  end
end
