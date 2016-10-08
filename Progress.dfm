object ProgressDlg: TProgressDlg
  Left = 231
  Top = 159
  BorderIcons = [biMinimize, biMaximize]
  BorderStyle = bsDialog
  ClientHeight = 69
  ClientWidth = 436
  Color = clBtnFace
  Font.Charset = SHIFTJIS_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #65325#65331' '#65328#12468#12471#12483#12463
  Font.Style = []
  OldCreateOrder = False
  Visible = True
  OnShow = FormShow
  PixelsPerInch = 96
  TextHeight = 12
  object lblFileName: TLabel
    Left = 72
    Top = 16
    Width = 46
    Height = 12
    Caption = #65420#65383#65394#65433#21517' :'
  end
  object txtFileName: TLabel
    Left = 128
    Top = 16
    Width = 289
    Height = 12
    AutoSize = False
  end
  object lblCount: TLabel
    Left = 72
    Top = 40
    Width = 46
    Height = 12
    Caption = #65420#65383#65394#65433#25968' :'
  end
  object txtCount: TLabel
    Left = 128
    Top = 40
    Width = 18
    Height = 12
    Caption = '0/0'
  end
  object btnCancel: TButton
    Left = 328
    Top = 40
    Width = 89
    Height = 22
    Caption = #65399#65388#65437#65406#65433
    TabOrder = 0
    OnClick = btnCancelClick
  end
  object aniFindFile: TAnimate
    Left = 13
    Top = 16
    Width = 16
    Height = 16
    Active = True
    CommonAVI = aviFindFile
    StopFrame = 8
  end
end
