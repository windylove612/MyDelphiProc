object fmQImport3ProgressDlg: TfmQImport3ProgressDlg
  Left = 292
  Top = 236
  ActiveControl = paInfo
  BorderIcons = []
  BorderStyle = bsDialog
  Caption = 'Importing...'
  ClientHeight = 224
  ClientWidth = 341
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  OldCreateOrder = True
  Position = poScreenCenter
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 13
  object bCancel: TButton
    Left = 132
    Top = 200
    Width = 75
    Height = 22
    Caption = 'Cancel'
    ModalResult = 1
    TabOrder = 0
    OnClick = bCancelClick
  end
  object paInfo: TPanel
    Left = 0
    Top = 0
    Width = 340
    Height = 196
    TabOrder = 1
    object bvProcessedValue: TBevel
      Left = 171
      Top = 29
      Width = 163
      Height = 20
    end
    object bvPervent: TBevel
      Left = 5
      Top = 149
      Width = 329
      Height = 18
    end
    object bvErrorsValue: TBevel
      Left = 254
      Top = 77
      Width = 80
      Height = 20
    end
    object bvDeletedValue: TBevel
      Left = 171
      Top = 77
      Width = 80
      Height = 20
    end
    object bvCommitted: TBevel
      Left = 5
      Top = 101
      Width = 163
      Height = 20
    end
    object bvCommittedValue: TBevel
      Left = 5
      Top = 125
      Width = 163
      Height = 20
    end
    object bvTimeValue: TBevel
      Left = 171
      Top = 124
      Width = 163
      Height = 20
    end
    object bvProcessed: TBevel
      Left = 171
      Top = 5
      Width = 163
      Height = 20
    end
    object laProcessed: TLabel
      Left = 174
      Top = 8
      Width = 157
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Caption = 'Processed'
    end
    object bvStateValue: TBevel
      Left = 5
      Top = 29
      Width = 163
      Height = 20
    end
    object laProcessedValue: TLabel
      Left = 175
      Top = 32
      Width = 155
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Layout = tlCenter
    end
    object bvState: TBevel
      Left = 5
      Top = 5
      Width = 163
      Height = 20
    end
    object laState: TLabel
      Left = 8
      Top = 8
      Width = 156
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Caption = 'Status'
    end
    object laStateValue: TLabel
      Left = 8
      Top = 32
      Width = 156
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Layout = tlCenter
    end
    object bvTime: TBevel
      Left = 171
      Top = 101
      Width = 163
      Height = 20
    end
    object laTime: TLabel
      Left = 174
      Top = 104
      Width = 156
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Caption = 'Time'
    end
    object laTimeValue: TLabel
      Left = 174
      Top = 127
      Width = 156
      Height = 13
      Alignment = taCenter
      AutoSize = False
    end
    object laPercent: TLabel
      Left = 9
      Top = 152
      Width = 321
      Height = 11
      Alignment = taCenter
      AutoSize = False
      Caption = '0%'
    end
    object laCommitted: TLabel
      Left = 8
      Top = 104
      Width = 156
      Height = 13
      Alignment = taCenter
      AutoSize = False
      Caption = 'Committed'
    end
    object laCommittedValue: TLabel
      Left = 9
      Top = 128
      Width = 156
      Height = 13
      Alignment = taCenter
      AutoSize = False
    end
    object bvInserted: TBevel
      Left = 5
      Top = 53
      Width = 80
      Height = 20
    end
    object bvUpdates: TBevel
      Left = 88
      Top = 53
      Width = 80
      Height = 20
    end
    object bvDeleted: TBevel
      Left = 171
      Top = 53
      Width = 80
      Height = 20
    end
    object bvErrors: TBevel
      Left = 254
      Top = 53
      Width = 80
      Height = 20
    end
    object laInserted: TLabel
      Left = 9
      Top = 56
      Width = 72
      Height = 14
      Alignment = taCenter
      AutoSize = False
      Caption = 'Inserted'
      Layout = tlCenter
    end
    object laUpdated: TLabel
      Left = 91
      Top = 56
      Width = 73
      Height = 14
      Alignment = taCenter
      AutoSize = False
      Caption = 'Updated'
      Layout = tlCenter
    end
    object laDeleted: TLabel
      Left = 174
      Top = 56
      Width = 73
      Height = 14
      Alignment = taCenter
      AutoSize = False
      Caption = 'Deleted'
      Layout = tlCenter
    end
    object laErrors: TLabel
      Left = 257
      Top = 56
      Width = 73
      Height = 14
      Alignment = taCenter
      AutoSize = False
      Caption = 'Errors'
      Layout = tlCenter
    end
    object bvInsertedValue: TBevel
      Left = 5
      Top = 77
      Width = 80
      Height = 20
    end
    object laInsertedValue: TLabel
      Left = 9
      Top = 80
      Width = 72
      Height = 14
      Alignment = taCenter
      AutoSize = False
      Layout = tlCenter
    end
    object bvUpdatedValue: TBevel
      Left = 88
      Top = 77
      Width = 80
      Height = 20
    end
    object laUpdatedValue: TLabel
      Left = 91
      Top = 80
      Width = 73
      Height = 14
      Alignment = taCenter
      AutoSize = False
      Layout = tlCenter
    end
    object laDeletedValue: TLabel
      Left = 175
      Top = 80
      Width = 73
      Height = 14
      Alignment = taCenter
      AutoSize = False
      Layout = tlCenter
    end
    object laErrorsValue: TLabel
      Left = 257
      Top = 80
      Width = 73
      Height = 14
      Alignment = taCenter
      AutoSize = False
      Layout = tlCenter
    end
    object prbImport: TProgressBar
      Left = 5
      Top = 174
      Width = 329
      Height = 16
      Step = 1
      TabOrder = 0
    end
  end
  object Timer: TTimer
    OnTimer = TimerTimer
    Top = 229
  end
end
