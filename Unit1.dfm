object Form1: TForm1
  Left = 192
  Top = 133
  Width = 366
  Height = 216
  Caption = #20803#12395#25147#12377'(Undo)'
  Color = clBtnFace
  Font.Charset = SHIFTJIS_CHARSET
  Font.Color = clWindowText
  Font.Height = -12
  Font.Name = #65325#65331' '#65328#12468#12471#12483#12463
  Font.Style = []
  OldCreateOrder = False
  OnCloseQuery = FormCloseQuery
  PixelsPerInch = 96
  TextHeight = 12
  object Button1: TButton
    Left = 48
    Top = 24
    Width = 257
    Height = 25
    Caption = #20803#12395#25147#12377#12434#23455#34892'(Undo)'
    TabOrder = 0
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 48
    Top = 72
    Width = 257
    Height = 25
    Caption = #20803#12395#25147#12377#12434#21462#24471'(Get Undo)'
    TabOrder = 1
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 48
    Top = 120
    Width = 257
    Height = 25
    Caption = #12501#12449#12452#12523#12434#38283#12367'(File Open Dialog)'
    TabOrder = 2
    OnClick = Button3Click
  end
  object OpenDialog1: TOpenDialog
    Left = 16
    Top = 120
  end
end
