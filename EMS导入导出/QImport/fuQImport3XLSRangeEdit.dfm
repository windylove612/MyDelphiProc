object fmQImport3XLSRangeEdit: TfmQImport3XLSRangeEdit
  Left = 335
  Top = 219
  BorderStyle = bsDialog
  Caption = 'Range'
  ClientHeight = 313
  ClientWidth = 316
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
  object gbRangeType: TGroupBox
    Left = 4
    Top = 1
    Width = 307
    Height = 48
    Caption = ' Range Type '
    TabOrder = 0
    object laColRowCell: TLabel
      Left = 229
      Top = 22
      Width = 15
      Height = 13
      Alignment = taRightJustify
      Caption = 'Col'
    end
    object edColRowCell: TEdit
      Left = 250
      Top = 18
      Width = 50
      Height = 21
      TabOrder = 1
    end
    object cbRangeType: TComboBox
      Left = 11
      Top = 18
      Width = 184
      Height = 21
      Style = csDropDownList
      ItemHeight = 13
      TabOrder = 0
      OnChange = UpdateControls
      Items.Strings = (
        'Col'
        'Row'
        'Cell')
    end
  end
  object gbFinishCondition: TGroupBox
    Left = 160
    Top = 50
    Width = 151
    Height = 68
    Caption = ' Finish '
    TabOrder = 2
    object rbFinishColRow: TRadioButton
      Left = 9
      Top = 41
      Width = 78
      Height = 17
      Caption = 'Finish Row'
      TabOrder = 1
      OnClick = UpdateControls
    end
    object rbFinishData: TRadioButton
      Left = 9
      Top = 17
      Width = 113
      Height = 17
      Caption = 'While Data Exists'
      Checked = True
      TabOrder = 0
      TabStop = True
      OnClick = UpdateControls
    end
    object edFinishColRow: TEdit
      Left = 94
      Top = 39
      Width = 50
      Height = 21
      TabOrder = 2
    end
  end
  object gbSheetType: TGroupBox
    Left = 4
    Top = 166
    Width = 308
    Height = 117
    Caption = ' Sheet '
    TabOrder = 4
    object rbDefaultSheet: TRadioButton
      Left = 11
      Top = 20
      Width = 290
      Height = 17
      Caption = 'Default Sheet'
      Checked = True
      TabOrder = 0
      TabStop = True
      OnClick = UpdateControls
    end
    object rbCustomSheet: TRadioButton
      Left = 11
      Top = 42
      Width = 290
      Height = 17
      Caption = 'Custom Sheet'
      TabOrder = 1
      OnClick = UpdateControls
    end
    object paCustomSheet: TPanel
      Left = 26
      Top = 60
      Width = 277
      Height = 51
      BevelOuter = bvNone
      TabOrder = 2
      object rbSheetNumber: TRadioButton
        Left = 2
        Top = 4
        Width = 113
        Height = 17
        Caption = 'Sheet Number'
        Checked = True
        TabOrder = 0
        TabStop = True
        OnClick = UpdateControls
      end
      object edSheetNumber: TComboBox
        Left = 122
        Top = 2
        Width = 55
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 1
      end
      object rbSheetName: TRadioButton
        Left = 2
        Top = 29
        Width = 113
        Height = 17
        Caption = 'Sheet Name'
        TabOrder = 2
        OnClick = UpdateControls
      end
      object edSheetName: TComboBox
        Left = 122
        Top = 27
        Width = 153
        Height = 21
        Style = csDropDownList
        ItemHeight = 13
        TabOrder = 3
      end
    end
  end
  object bCancel: TButton
    Left = 182
    Top = 288
    Width = 80
    Height = 22
    Cancel = True
    Caption = 'Cancel'
    ModalResult = 2
    TabOrder = 6
  end
  object bOk: TButton
    Left = 54
    Top = 288
    Width = 80
    Height = 22
    Caption = 'OK'
    Default = True
    TabOrder = 5
    OnClick = bOkClick
  end
  object gbStartCondition: TGroupBox
    Left = 4
    Top = 50
    Width = 151
    Height = 68
    Caption = ' Start '
    TabOrder = 1
    object rbStartData: TRadioButton
      Left = 9
      Top = 17
      Width = 126
      Height = 17
      Caption = 'From Data Starting'
      Checked = True
      TabOrder = 0
      TabStop = True
      OnClick = UpdateControls
    end
    object rbStartColRow: TRadioButton
      Left = 9
      Top = 41
      Width = 76
      Height = 17
      Caption = 'Start Row'
      TabOrder = 1
      OnClick = UpdateControls
    end
    object edStartColRow: TEdit
      Left = 94
      Top = 39
      Width = 50
      Height = 21
      TabOrder = 2
    end
  end
  object rgDirection: TRadioGroup
    Left = 4
    Top = 119
    Width = 308
    Height = 46
    Caption = ' Direction '
    Columns = 2
    ItemIndex = 0
    Items.Strings = (
      'Down'
      'Up')
    TabOrder = 3
    OnClick = UpdateControls
  end
end
