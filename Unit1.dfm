object Form1: TForm1
  Left = 0
  Top = 0
  Caption = #1055#1091#1083#1100#1090
  ClientHeight = 412
  ClientWidth = 670
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesktopCenter
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  OnResize = FormResize
  PixelsPerInch = 96
  TextHeight = 13
  object Label1: TLabel
    Left = 8
    Top = 10
    Width = 154
    Height = 19
    Caption = #1042#1099#1073#1077#1088#1080#1090#1077' '#1086#1090#1088#1072#1089#1083#1100
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -16
    Font.Name = 'Tahoma'
    Font.Style = [fsBold]
    ParentFont = False
  end
  object DBLookupComboBox1: TDBLookupComboBox
    Left = 194
    Top = 8
    Width = 183
    Height = 21
    KeyField = #1053#1072#1079#1074#1072#1085#1080#1077
    ListFieldIndex = 1
    ListSource = Dm1.DataSource1
    TabOrder = 0
    OnClick = DBLookupComboBox1Click
  end
  object PageControl: TPageControl
    Left = -2
    Top = 40
    Width = 684
    Height = 382
    ActivePage = TabSheet1
    MultiLine = True
    TabOrder = 1
    object TabSheet1: TTabSheet
      Caption = #1058#1072#1073#1083#1080#1094#1072
      object Label2: TLabel
        Left = 2
        Top = 4
        Width = 184
        Height = 19
        Caption = #1042#1099#1073#1077#1088#1080#1090#1077' '#1087#1086#1082#1072#1079#1072#1090#1077#1083#1100
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -16
        Font.Name = 'Tahoma'
        Font.Style = [fsBold]
        ParentFont = False
      end
      object ComboBox1: TComboBox
        Left = 192
        Top = 2
        Width = 233
        Height = 21
        TabOrder = 0
        Text = 'ComboBox1'
        OnClick = ComboBox1Click
      end
      object StringGrid1: TStringGrid
        Left = 0
        Top = 29
        Width = 665
        Height = 120
        ColCount = 27
        TabOrder = 1
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1044#1080#1072#1075#1088#1072#1084#1084#1072
      ImageIndex = 1
      object Chart: TChart
        Left = 0
        Top = 0
        Width = 676
        Height = 354
        Title.Text.Strings = (
          'TChart')
        Align = alClient
        TabOrder = 0
        DefaultCanvas = 'TGDIPlusCanvas'
        PrintMargins = (
          15
          24
          15
          24)
        ColorPaletteIndex = 13
      end
    end
  end
  object Button1: TButton
    Left = 416
    Top = 9
    Width = 75
    Height = 25
    Caption = 'Button1'
    TabOrder = 2
    OnClick = Button1Click
  end
  object Button2: TButton
    Left = 520
    Top = 8
    Width = 75
    Height = 25
    Caption = 'Button2'
    TabOrder = 3
    OnClick = Button2Click
  end
end
