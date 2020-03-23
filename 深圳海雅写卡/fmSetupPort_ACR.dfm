object SetPort_ACR: TSetPort_ACR
  Left = 645
  Top = 248
  Width = 296
  Height = 329
  Caption = #36890#35759#31471#21475#35774#32622
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = False
  PixelsPerInch = 96
  TextHeight = 13
  object grpBaud: TRzRadioGroup
    Left = 18
    Top = 8
    Width = 123
    Height = 265
    Caption = 'USB'#31471#21475
    Color = clWindow
    ItemFrameColor = clSkyBlue
    ItemIndex = 0
    Items.Strings = (
      'USB1'
      'USB2'
      'USB3'
      'USB4'
      'USB5'
      'USB6'
      'USB7'
      'USB8'
      'USB9'
      'USB10'
      'USB11'
      'USB12')
    TabOrder = 0
  end
  object btnOk: TBmpBtn
    Left = 166
    Top = 87
    FrameColor = clScrollBar
    Color = 14737632
    HotTrack = True
    TabOrder = 1
    OnClick = btnOkClick
    Kind = bkOk
  end
  object btnCancel: TBmpBtn
    Left = 166
    Top = 145
    FrameColor = clScrollBar
    Color = 14737632
    HotTrack = True
    TabOrder = 2
    OnClick = btnCancelClick
    Kind = bkCancel
  end
end
