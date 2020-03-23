object fmQImport3Loading: TfmQImport3Loading
  Left = 384
  Top = 287
  BorderIcons = []
  BorderStyle = bsNone
  Caption = 'fmQImport3Loading'
  ClientHeight = 103
  ClientWidth = 288
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'MS Sans Serif'
  Font.Style = []
  FormStyle = fsStayOnTop
  OldCreateOrder = True
  Position = poScreenCenter
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object Bevel1: TBevel
    Left = 0
    Top = 0
    Width = 288
    Height = 103
    Align = alClient
    Shape = bsFrame
    Style = bsRaised
  end
  object pbLoadingInfo: TPaintBox
    Left = 8
    Top = 84
    Width = 272
    Height = 15
    OnPaint = pbLoadingInfoPaint
  end
  object laLoading_01: TLabel
    Left = 8
    Top = 6
    Width = 272
    Height = 15
    Alignment = taCenter
    AutoSize = False
    Caption = 'Loading...'
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -11
    Font.Name = 'MS Sans Serif'
    Font.Style = [fsBold]
    ParentFont = False
    Layout = tlCenter
  end
  object Animate1: TAnimate
    Left = 8
    Top = 22
    Width = 272
    Height = 60
    CommonAVI = aviCopyFiles
    StopFrame = 31
  end
end
