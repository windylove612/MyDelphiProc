object Form1: TForm1
  Left = 388
  Top = 292
  Width = 583
  Height = 235
  Caption = #21475#20196#21152#23494#35299#23494
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object RzLabel1: TRzLabel
    Left = 64
    Top = 40
    Width = 57
    Height = 33
    AutoSize = False
    Caption = #26126#25991
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object RzLabel2: TRzLabel
    Left = 64
    Top = 104
    Width = 57
    Height = 33
    AutoSize = False
    Caption = #26263#25991
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object RzEdit1: TRzEdit
    Left = 128
    Top = 42
    Width = 313
    Height = 21
    FrameColor = clSkyBlue
    FrameVisible = True
    TabOrder = 0
  end
  object RzEdit2: TRzEdit
    Left = 128
    Top = 106
    Width = 313
    Height = 21
    FrameColor = clSkyBlue
    FrameVisible = True
    TabOrder = 1
  end
  object RzButton1: TRzButton
    Left = 448
    Top = 40
    Caption = #36716#26263#25991
    HotTrack = True
    TabOrder = 2
    OnClick = RzButton1Click
  end
  object RzButton2: TRzButton
    Left = 448
    Top = 104
    Caption = #36716#26126#25991
    HotTrack = True
    TabOrder = 3
    OnClick = RzButton2Click
  end
end
