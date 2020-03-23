object MainForm: TMainForm
  Left = 0
  Top = 0
  Caption = 'GetPhyMac'
  ClientHeight = 293
  ClientWidth = 426
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object GetMac: TButton
    Left = 24
    Top = 248
    Width = 377
    Height = 25
    Caption = #33719#21462#32593#21345#29289#29702#22320#22336
    TabOrder = 0
    OnClick = GetMacClick
  end
  object Memo: TMemo
    Left = 0
    Top = 0
    Width = 426
    Height = 226
    Align = alTop
    ScrollBars = ssVertical
    TabOrder = 1
  end
end
