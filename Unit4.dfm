object Form4: TForm4
  Left = 0
  Top = 0
  Align = alClient
  BorderIcons = [biSystemMenu]
  Caption = #1055#1091#1083#1100#1090
  ClientHeight = 616
  ClientWidth = 1041
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  WindowState = wsMaximized
  OnCreate = FormCreate
  PixelsPerInch = 96
  TextHeight = 13
  object PageControlOsnova: TPageControl
    Left = 0
    Top = 0
    Width = 1041
    Height = 617
    ActivePage = TabSheet1
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = #1041#1072#1079#1072' '#1044#1072#1085#1085#1099#1093
      object PanelBD: TPanel
        Left = 0
        Top = 0
        Width = 185
        Height = 584
        TabOrder = 0
        object ButtonDimografia: TButton
          Left = 0
          Top = 32
          Width = 185
          Height = 33
          Caption = #1044#1077#1084#1086#1075#1088#1072#1092#1080#1103
          TabOrder = 0
          OnClick = ButtonDimografiaClick
        end
        object ButtonDinamic: TButton
          Left = 1
          Top = 0
          Width = 185
          Height = 33
          Caption = #1044#1080#1085#1072#1084#1080#1082#1072' '#1087#1086' '#1086#1090#1088#1072#1089#1083#1103#1084
          TabOrder = 1
          OnClick = ButtonDinamicClick
        end
      end
      object PageControlDinamic: TPageControl
        Left = 188
        Top = 0
        Width = 845
        Height = 589
        ActivePage = TabSheetDimografiaTable
        Align = alRight
        MultiLine = True
        TabOrder = 1
        Visible = False
        OnChange = PageControlDinamicChange
        ExplicitLeft = 191
        object TabSheetDinamicTable: TTabSheet
          Caption = #1058#1072#1073#1083#1080#1094#1072
          object Label2: TLabel
            Left = 2
            Top = 33
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
          object Label1: TLabel
            Left = 2
            Top = 5
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
          object ComboBoxDinamic: TComboBox
            Left = 192
            Top = 31
            Width = 233
            Height = 21
            TabOrder = 0
            OnClick = ComboBoxDinamicClick
          end
          object StringGridDinamic: TStringGrid
            Left = 0
            Top = 63
            Width = 665
            Height = 120
            Align = alCustom
            ColCount = 27
            TabOrder = 1
            ColWidths = (
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64)
          end
          object DBLookupComboBoxDinamic: TDBLookupComboBox
            Left = 192
            Top = 3
            Width = 233
            Height = 21
            KeyField = #1053#1072#1079#1074#1072#1085#1080#1077
            ListFieldIndex = 1
            ListSource = Dm1.DataSource1
            TabOrder = 2
            OnClick = DBLookupComboBoxDinamicClick
          end
        end
        object TabSheetDinamicChart: TTabSheet
          Caption = #1044#1080#1072#1075#1088#1072#1084#1084#1072
          ImageIndex = 1
          object ChartDinamic: TChart
            Left = 0
            Top = 0
            Width = 837
            Height = 561
            Title.Text.Strings = (
              'TChart')
            Align = alClient
            TabOrder = 0
            DefaultCanvas = 'TGDIPlusCanvas'
            PrintMargins = (
              15
              16
              15
              16)
            ColorPaletteIndex = 13
            object Series1: TFastLineSeries
              LinePen.Color = 10708548
              XValues.Name = 'X'
              XValues.Order = loAscending
              YValues.Name = 'Y'
              YValues.Order = loNone
            end
          end
        end
        object TabSheetDimografiaTable: TTabSheet
          Caption = #1058#1072#1073#1083#1080#1094#1072
          ImageIndex = 2
          object Label3: TLabel
            Left = 10
            Top = 5
            Width = 97
            Height = 19
            Caption = #1042#1099#1073#1077#1088#1080#1090#1077' ?'
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentFont = False
          end
          object StringGridDimografia: TStringGrid
            Left = 0
            Top = 35
            Width = 665
            Height = 120
            Align = alCustom
            ColCount = 27
            TabOrder = 0
            ColWidths = (
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64
              64)
          end
          object ComboBoxDimografia: TComboBox
            Left = 200
            Top = 3
            Width = 233
            Height = 21
            TabOrder = 1
            OnClick = ComboBoxDimografiaClick
          end
        end
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1057#1094#1077#1085#1072#1088#1080#1080
      ImageIndex = 1
    end
  end
  object MainMenu1: TMainMenu
    Left = 704
    Top = 65513
    object N11: TMenuItem
      Caption = #1060#1072#1081#1083#1099
      object N12: TMenuItem
        Caption = '1'
      end
      object N22: TMenuItem
        Caption = '2'
      end
    end
    object N21: TMenuItem
      Caption = #1044#1088' '#1087#1091#1085#1082#1090' '#1084#1077#1085#1102
    end
  end
end