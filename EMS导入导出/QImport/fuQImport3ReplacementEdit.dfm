object fmQImport3ReplacementEdit: TfmQImport3ReplacementEdit
  Left = 356
  Top = 279
  ActiveControl = edTextToFind
  BorderStyle = bsDialog
  Caption = 'Replacement'
  ClientHeight = 150
  ClientWidth = 314
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
  object Bevel: TBevel
    Left = 3
    Top = 3
    Width = 308
    Height = 116
  end
  object laTextToFind: TLabel
    Left = 9
    Top = 7
    Width = 53
    Height = 13
    Caption = 'Text to find'
  end
  object laReplaceWith: TLabel
    Left = 9
    Top = 48
    Width = 62
    Height = 13
    Caption = 'Replace with'
  end
  object bOk: TButton
    Left = 64
    Top = 126
    Width = 75
    Height = 22
    Caption = 'OK'
    Default = True
    TabOrder = 3
    OnClick = bOkClick
  end
  object bCancel: TButton
    Left = 176
    Top = 126
    Width = 75
    Height = 22
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 4
  end
  object chIgnoreCase: TCheckBox
    Left = 9
    Top = 94
    Width = 298
    Height = 17
    Caption = 'Ignore Case'
    TabOrder = 2
  end
  object edTextToFind: TEdit
    Left = 9
    Top = 24
    Width = 298
    Height = 21
    TabOrder = 0
  end
  object edReplaceWith: TEdit
    Left = 9
    Top = 65
    Width = 298
    Height = 21
    TabOrder = 1
  end
end
