object fmQImport3License: TfmQImport3License
  Left = 343
  Top = 197
  BorderIcons = [biSystemMenu]
  BorderStyle = bsDialog
  Caption = 'EMS AdvancedImport License Agreement'
  ClientHeight = 304
  ClientWidth = 438
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object reLicense: TRichEdit
    Left = 8
    Top = 8
    Width = 420
    Height = 257
    ReadOnly = True
    ScrollBars = ssVertical
    TabOrder = 0
  end
  object btnPrint: TButton
    Left = 8
    Top = 272
    Width = 90
    Height = 25
    Caption = 'Print'
    TabOrder = 1
    OnClick = btnPrintClick
  end
  object btnOK: TButton
    Left = 338
    Top = 272
    Width = 90
    Height = 25
    Caption = 'OK'
    ModalResult = 1
    TabOrder = 2
  end
end
