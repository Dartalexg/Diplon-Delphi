object PultUpav: TPultUpav
  Left = 0
  Top = 0
  Align = alClient
  BorderIcons = [biSystemMenu]
  Caption = #1054#1094#1077#1085#1082#1072' '#1076#1086#1089#1090#1080#1078#1080#1084#1086#1089#1090#1080' '#1080#1085#1076#1080#1082#1072#1090#1086#1088#1086#1074' '#1088#1072#1079#1074#1080#1090#1080#1103' '#1089#1086#1094#1080#1072#1083#1100#1085#1086#1081' '#1089#1092#1077#1088#1099
  ClientHeight = 650
  ClientWidth = 1484
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
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object PageControlOsnova: TPageControl
    Left = 0
    Top = 0
    Width = 1484
    Height = 650
    ActivePage = TabSheet4
    Align = alClient
    TabOrder = 0
    OnChange = PageControlOsnovaChange
    object TabSheet1: TTabSheet
      Caption = #1041#1072#1079#1072' '#1044#1072#1085#1085#1099#1093
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object PanelBD: TPanel
        Left = 0
        Top = 0
        Width = 189
        Height = 584
        TabOrder = 0
        object ButtonDimografia: TButton
          Left = 0
          Top = 33
          Width = 182
          Height = 33
          Caption = #1044#1077#1084#1086#1075#1088#1072#1092#1080#1103
          TabOrder = 0
          OnClick = ButtonDimografiaClick
        end
        object ButtonDinamic: TButton
          Left = 0
          Top = 0
          Width = 182
          Height = 33
          Caption = #1044#1080#1085#1072#1084#1080#1082#1072' '#1087#1086' '#1086#1090#1088#1072#1089#1083#1103#1084
          TabOrder = 1
          OnClick = ButtonDinamicClick
        end
      end
      object PageControlDinamic: TPageControl
        Left = 188
        Top = 0
        Width = 1288
        Height = 622
        ActivePage = TabSheetDimografiaTable
        Align = alRight
        MultiLine = True
        TabOrder = 1
        Visible = False
        object TabSheetDinamicChart: TTabSheet
          Caption = #1044#1080#1072#1075#1088#1072#1084#1084#1072
          ImageIndex = 1
          ExplicitLeft = 0
          ExplicitTop = 0
          ExplicitWidth = 0
          ExplicitHeight = 0
          object ChartDinamic: TChart
            Left = 0
            Top = 0
            Width = 1280
            Height = 594
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
        object TabSheetDinamicTable: TTabSheet
          Caption = #1058#1072#1073#1083#1080#1094#1072
          ExplicitLeft = 0
          ExplicitTop = 0
          ExplicitWidth = 0
          ExplicitHeight = 0
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
            Width = 401
            Height = 21
            Style = csDropDownList
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
            Width = 401
            Height = 21
            KeyField = #1053#1072#1079#1074#1072#1085#1080#1077
            ListFieldIndex = 1
            ListSource = Dm1.DataSource1
            TabOrder = 2
            OnClick = DBLookupComboBoxDinamicClick
          end
        end
        object TabSheetDimografiaTable: TTabSheet
          Caption = #1058#1072#1073#1083#1080#1094#1072
          ImageIndex = 2
          ExplicitLeft = 0
          ExplicitTop = 0
          ExplicitWidth = 0
          ExplicitHeight = 0
          object Label3: TLabel
            Left = 3
            Top = 3
            Width = 248
            Height = 19
            Caption = #1042#1099#1073#1077#1088#1080#1090#1077' '#1074#1086#1079#1088#1072#1089#1090#1085#1091#1102' '#1075#1088#1091#1087#1087#1091
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
            Left = 258
            Top = 3
            Width = 401
            Height = 21
            Style = csDropDownList
            TabOrder = 1
            OnClick = ComboBoxDimografiaClick
          end
        end
      end
    end
    object TabSheet3: TTabSheet
      Caption = #1055#1091#1083#1100#1090
      ImageIndex = 2
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object ScrollBox1: TScrollBox
        Left = 0
        Top = 0
        Width = 1500
        Height = 622
        Align = alLeft
        TabOrder = 0
        object PultPanelScriptZP: TPanel
          Left = 1350
          Top = 0
          Width = 350
          Height = 601
          Align = alLeft
          TabOrder = 0
          object Label22: TLabel
            Left = 1
            Top = 1
            Width = 348
            Height = 19
            Align = alTop
            Alignment = taCenter
            Caption = #1057#1094#1077#1085#1072#1088#1080#1081' '#1088#1086#1089#1090#1072' '#1079#1072#1088#1087#1083#1072#1090' '#1074' '#1041#1057
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentFont = False
            ExplicitWidth = 248
          end
          object Label23: TLabel
            Left = 4
            Top = 30
            Width = 32
            Height = 18
            Caption = #1046#1050#1061
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label24: TLabel
            Left = 2
            Top = 59
            Width = 176
            Height = 18
            Caption = #1044#1086#1096#1082#1086#1083#1100#1085#1099#1077' '#1091#1095#1088#1077#1078#1076#1077#1085#1080#1103
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label25: TLabel
            Left = 4
            Top = 86
            Width = 140
            Height = 18
            Caption = #1054#1073#1097#1077#1077' '#1086#1073#1088#1072#1079#1086#1074#1072#1085#1080#1077
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label26: TLabel
            Left = 4
            Top = 110
            Width = 68
            Height = 18
            Caption = #1041#1086#1083#1100#1085#1080#1094#1099
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label27: TLabel
            Left = 4
            Top = 134
            Width = 88
            Height = 18
            Caption = #1055#1086#1083#1080#1082#1083#1080#1085#1080#1082#1080
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label28: TLabel
            Left = 4
            Top = 161
            Width = 65
            Height = 18
            Caption = #1050#1091#1083#1100#1090#1091#1088#1072
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label29: TLabel
            Left = 4
            Top = 185
            Width = 147
            Height = 18
            Caption = #1060#1080#1079#1080#1095#1077#1089#1082#1072#1103' '#1082#1091#1083#1100#1090#1091#1088#1072
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object BoxFizKeltScriptPoctZP: TComboBox
            Left = 188
            Top = 182
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 0
            Text = '2'
            OnClick = BoxFizKeltScriptPoctZPClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxKyltScriptPoctZP: TComboBox
            Left = 188
            Top = 157
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 1
            Text = '2'
            OnClick = BoxKyltScriptPoctZPClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxPoliclinScriptPoctZP: TComboBox
            Left = 188
            Top = 132
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 2
            Text = '2'
            OnClick = BoxPoliclinScriptPoctZPClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxBolnicScriptPoctZP: TComboBox
            Left = 188
            Top = 107
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 3
            Text = '2'
            OnClick = BoxBolnicScriptPoctZPClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxObheeObrozScriptPoctZP: TComboBox
            Left = 188
            Top = 82
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 4
            Text = '2'
            OnClick = BoxObheeObrozScriptPoctZPClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxHkolaScriptPoctZP: TComboBox
            Left = 188
            Top = 57
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 5
            Text = '2'
            OnClick = BoxHkolaScriptPoctZPClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxGKXScriptPoctZP: TComboBox
            Left = 188
            Top = 31
            Width = 133
            Height = 22
            Style = csDropDownList
            Enabled = False
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 6
            Items.Strings = (
              '1'
              '2'
              '3')
          end
        end
        object PultPanelScriptRostTarif: TPanel
          Left = 1000
          Top = 0
          Width = 350
          Height = 601
          Align = alLeft
          TabOrder = 1
          object Label14: TLabel
            Left = 1
            Top = 1
            Width = 348
            Height = 19
            Align = alTop
            Alignment = taCenter
            Caption = #1057#1094#1077#1085#1072#1088#1080#1080' '#1088#1086#1089#1090#1072' '#1090#1072#1088#1080#1092#1086#1074
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentFont = False
            ExplicitWidth = 211
          end
          object Label15: TLabel
            Left = 6
            Top = 30
            Width = 32
            Height = 18
            Caption = #1046#1050#1061
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label16: TLabel
            Left = 4
            Top = 59
            Width = 176
            Height = 18
            Caption = #1044#1086#1096#1082#1086#1083#1100#1085#1099#1077' '#1091#1095#1088#1077#1078#1076#1077#1085#1080#1103
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label17: TLabel
            Left = 6
            Top = 85
            Width = 140
            Height = 18
            Caption = #1054#1073#1097#1077#1077' '#1086#1073#1088#1072#1079#1086#1074#1072#1085#1080#1077
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label18: TLabel
            Left = 6
            Top = 110
            Width = 68
            Height = 18
            Caption = #1041#1086#1083#1100#1085#1080#1094#1099
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label19: TLabel
            Left = 6
            Top = 134
            Width = 88
            Height = 18
            Caption = #1055#1086#1083#1080#1082#1083#1080#1085#1080#1082#1080
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label20: TLabel
            Left = 6
            Top = 161
            Width = 65
            Height = 18
            Caption = #1050#1091#1083#1100#1090#1091#1088#1072
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label21: TLabel
            Left = 6
            Top = 185
            Width = 147
            Height = 18
            Caption = #1060#1080#1079#1080#1095#1077#1089#1082#1072#1103' '#1082#1091#1083#1100#1090#1091#1088#1072
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object BoxGKXScriptTarif: TComboBox
            Left = 208
            Top = 31
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 0
            Text = '2'
            OnClick = BoxGKXScriptTarifClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxHkolaScriptTarif: TComboBox
            Left = 208
            Top = 57
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 1
            Text = '2'
            OnClick = BoxHkolaScriptTarifClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxObheeObrozScriptTarif: TComboBox
            Left = 208
            Top = 82
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 2
            Text = '2'
            OnClick = BoxObheeObrozScriptTarifClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxBolnicScriptTarif: TComboBox
            Left = 208
            Top = 107
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 3
            Text = '2'
            OnClick = BoxBolnicScriptTarifClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxPoliclinScriptTarif: TComboBox
            Left = 208
            Top = 132
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 4
            Text = '2'
            OnClick = BoxPoliclinScriptTarifClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxKyltScriptTarif: TComboBox
            Left = 208
            Top = 157
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 5
            Text = '2'
            OnClick = BoxKyltScriptTarifClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxFizKeltScriptTarif: TComboBox
            Left = 208
            Top = 182
            Width = 133
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 6
            Text = '2'
            OnClick = BoxFizKeltScriptTarifClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
        end
        object PultPanelBudgetRegion: TPanel
          Left = 300
          Top = 0
          Width = 350
          Height = 601
          Align = alLeft
          TabOrder = 2
          object Label6: TLabel
            Left = 1
            Top = 1
            Width = 348
            Height = 19
            Align = alTop
            Alignment = taCenter
            Caption = #1050#1088#1072#1077#1074#1086#1081' '#1073#1102#1076#1078#1077#1090
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentFont = False
            ExplicitWidth = 146
          end
          object Label7: TLabel
            Left = 5
            Top = 32
            Width = 179
            Height = 18
            Caption = #1043#1086#1076' '#1089#1090#1072#1088#1090#1072' '#1087#1088#1086#1077#1082#1090#1086#1074' '#1052#1057#1041
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label8: TLabel
            Left = 6
            Top = 59
            Width = 223
            Height = 18
            Caption = 'C'#1094#1077#1085#1072#1088#1080#1081' '#1080#1085#1074#1077#1089#1090#1080#1094#1080#1081' '#1060#1041' '#1074' '#1052#1057#1041
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label9: TLabel
            Left = 6
            Top = 83
            Width = 221
            Height = 35
            AutoSize = False
            Caption = #1058#1077#1084#1087' '#1088#1086#1089#1090#1072' '#1089#1086#1073#1089#1090#1074#1077#1085#1085#1099#1093' '#1076#1086#1093#1086#1076#1086#1074' '#1086#1090' '#1089#1090#1072#1088#1086#1081' '#1101#1082#1086#1085#1086#1084#1080#1082#1080
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            WordWrap = True
          end
          object Label10: TLabel
            Left = 6
            Top = 124
            Width = 219
            Height = 36
            AutoSize = False
            Caption = #1057#1094#1077#1085#1072#1088#1080#1081' '#1090#1088#1072#1085#1089#1092#1077#1088#1090#1086#1074' '#1080' '#1080#1085#1074#1077#1089#1090#1080#1094#1080#1081
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            WordWrap = True
          end
          object BoxYearStartProject: TComboBox
            Left = 234
            Top = 29
            Width = 113
            Height = 22
            Align = alCustom
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 5
            ParentFont = False
            TabOrder = 0
            Text = '2010'
            OnClick = BoxYearStartProjectClick
            Items.Strings = (
              '2005'
              '2006'
              '2007'
              '2008'
              '2009'
              '2010'
              '2011'
              '2012'
              '2013'
              '2014'
              '2015'
              '2016'
              '2017'
              '2018'
              '2019'
              '2020')
          end
          object BoxScriptInvesticFB: TComboBox
            Left = 234
            Top = 55
            Width = 113
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 1
            Text = '2'
            OnClick = BoxScriptInvesticFBClick
            Items.Strings = (
              '1'
              '2'
              '3'
              '4'
              '5')
          end
          object BoxTempPoctDoxodOtStartEkonom: TComboBox
            Left = 233
            Top = 96
            Width = 113
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 2
            Text = '2'
            OnClick = BoxTempPoctDoxodOtStartEkonomClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxScriptTransferAndInvest: TComboBox
            Left = 234
            Top = 138
            Width = 113
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 3
            Text = '2'
            OnClick = BoxScriptTransferAndInvestClick
            Items.Strings = (
              '1'
              '2'
              '3'
              '4'
              '5')
          end
          object BitBtn2: TBitBtn
            Left = 195
            Top = 139
            Width = 33
            Height = 21
            Caption = ' '
            Kind = bkHelp
            NumGlyphs = 2
            TabOrder = 4
            OnClick = BitBtn2Click
          end
        end
        object PultPanelNONProductSfer: TPanel
          Left = 650
          Top = 0
          Width = 350
          Height = 601
          Align = alLeft
          TabOrder = 3
          object Label11: TLabel
            Left = 1
            Top = 1
            Width = 348
            Height = 19
            Align = alTop
            Alignment = taCenter
            Caption = #1053#1077#1087#1088#1086#1080#1079#1074#1086#1076#1089#1090#1074#1077#1085#1085#1072#1103' '#1089#1092#1077#1088#1072
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentFont = False
            ExplicitWidth = 234
          end
          object Label12: TLabel
            Left = 6
            Top = 32
            Width = 132
            Height = 18
            Caption = #1057#1094#1077#1085#1072#1088#1080#1081'  '#1076#1086#1083#1080' '#1053#1057
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object Label13: TLabel
            Left = 6
            Top = 59
            Width = 123
            Height = 37
            AutoSize = False
            Caption = #1057#1094#1077#1085#1072#1088#1080#1081' '#1074#1074#1086#1076#1072' '#1078#1080#1083#1100#1103' '#1079#1072' '#1089#1095#1077#1090' '#1085#1072#1089#1077#1083#1077#1085#1080#1103
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            WordWrap = True
          end
          object BoxScriptDoliNS: TComboBox
            Left = 160
            Top = 29
            Width = 184
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 0
            Text = '2'
            OnClick = BoxScriptDoliNSClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BoxScriptReadGilaZaCheat: TComboBox
            Left = 160
            Top = 75
            Width = 184
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 1
            Text = '2'
            OnClick = BoxScriptReadGilaZaCheatClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
        end
        object PultPanelScriptINFL: TPanel
          Left = 0
          Top = 0
          Width = 300
          Height = 601
          Align = alLeft
          TabOrder = 4
          object Label4: TLabel
            Left = 1
            Top = 1
            Width = 298
            Height = 19
            Align = alTop
            Alignment = taCenter
            Caption = #1057#1094#1077#1085#1072#1088#1080#1080' '#1080#1085#1092#1083#1103#1094#1080#1080
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -16
            Font.Name = 'Tahoma'
            Font.Style = [fsBold]
            ParentFont = False
            ExplicitWidth = 171
          end
          object Label5: TLabel
            Left = 4
            Top = 26
            Width = 66
            Height = 18
            Caption = #1057#1094#1077#1085#1072#1088#1080#1081
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -15
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
          end
          object BoxScriptINFL: TComboBox
            Left = 76
            Top = 29
            Width = 184
            Height = 22
            Style = csDropDownList
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -12
            Font.Name = 'Tahoma'
            Font.Style = []
            ItemIndex = 1
            ParentFont = False
            TabOrder = 0
            Text = '2'
            OnClick = BoxScriptINFLClick
            Items.Strings = (
              '1'
              '2'
              '3')
          end
          object BitBtn1: TBitBtn
            Left = 265
            Top = 27
            Width = 33
            Height = 25
            Caption = ' '
            Kind = bkHelp
            NumGlyphs = 2
            TabOrder = 1
            OnClick = BitBtn1Click
          end
        end
      end
    end
    object TabSheet4: TTabSheet
      Caption = #1044#1080#1085#1072#1084#1080#1082#1072' '#1086#1073#1077#1089#1087#1077#1095#1077#1085#1085#1086#1089#1090#1080
      ImageIndex = 3
      object PageControl1: TPageControl
        Left = 0
        Top = 0
        Width = 1476
        Height = 622
        ActivePage = TabSheet9
        Align = alClient
        MultiLine = True
        TabOrder = 0
        object TabSheet9: TTabSheet
          Caption = #1058#1072#1073#1083#1080#1094#1072
          ImageIndex = 2
          object StringGridDinamicObecpec: TStringGrid
            Left = 0
            Top = 35
            Width = 660
            Height = 204
            Align = alCustom
            ColCount = 27
            RowCount = 8
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
          object CheckBox1: TCheckBox
            Left = 276
            Top = 12
            Width = 185
            Height = 16
            Caption = #1054#1073#1097#1077#1077' '#1087#1086#1083#1085#1086#1077' '#1086#1073#1088#1072#1079#1086#1074#1072#1085#1080#1077
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 1
            OnClick = CheckBox2Click
          end
          object CheckBox2: TCheckBox
            Left = 3
            Top = 12
            Width = 95
            Height = 17
            Caption = #1046#1080#1083#1086#1081' '#1092#1086#1085#1076
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 2
            OnClick = CheckBox2Click
          end
          object CheckBox3: TCheckBox
            Left = 99
            Top = 12
            Width = 175
            Height = 17
            Caption = #1044#1086#1096#1082#1086#1083#1100#1085#1086#1077' '#1086#1073#1088#1072#1079#1086#1074#1072#1085#1080#1077
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 3
            OnClick = CheckBox2Click
          end
          object CheckBox4: TCheckBox
            Left = 464
            Top = 13
            Width = 80
            Height = 16
            Caption = #1041#1086#1083#1100#1085#1080#1094#1099
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 4
            OnClick = CheckBox2Click
          end
          object CheckBox5: TCheckBox
            Left = 542
            Top = 13
            Width = 98
            Height = 16
            Caption = #1055#1086#1083#1080#1082#1083#1080#1085#1080#1082#1080
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 5
            OnClick = CheckBox2Click
          end
          object CheckBox6: TCheckBox
            Left = 637
            Top = 13
            Width = 75
            Height = 16
            Caption = #1050#1091#1083#1100#1090#1091#1088#1072
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 6
            OnClick = CheckBox2Click
          end
          object CheckBox7: TCheckBox
            Left = 713
            Top = 13
            Width = 151
            Height = 16
            Caption = #1060#1080#1079#1080#1095#1077#1089#1082#1072#1103' '#1082#1091#1083#1100#1090#1091#1088#1072
            Font.Charset = DEFAULT_CHARSET
            Font.Color = clWindowText
            Font.Height = -13
            Font.Name = 'Tahoma'
            Font.Style = []
            ParentFont = False
            TabOrder = 7
            OnClick = CheckBox2Click
          end
          object Button1: TButton
            Left = 440
            Top = 328
            Width = 75
            Height = 25
            Caption = 'Button1'
            TabOrder = 8
            OnClick = Button1Click
          end
        end
        object TabSheet7: TTabSheet
          Caption = #1044#1080#1072#1075#1088#1072#1084#1084#1072
          ImageIndex = 1
          ExplicitLeft = 0
          ExplicitTop = 0
          ExplicitWidth = 0
          ExplicitHeight = 0
          object Chart1: TChart
            Left = 0
            Top = 0
            Width = 1468
            Height = 594
            Title.Font.Color = clBlack
            Title.Text.Strings = (
              'TChart')
            View3D = False
            Align = alClient
            TabOrder = 0
            DefaultCanvas = 'TGDIPlusCanvas'
            PrintMargins = (
              15
              16
              15
              16)
            ColorPaletteIndex = 13
            object FastLineSeries1: TFastLineSeries
              Marks.Children = <
                item
                  Shape.ShapeStyle = fosRectangle
                  Shape.Style = smsValue
                end>
              LinePen.Color = 10708548
              XValues.Name = 'X'
              XValues.Order = loAscending
              YValues.Name = 'Y'
              YValues.Order = loNone
            end
          end
        end
      end
    end
    object TabSheet5: TTabSheet
      Caption = #1054#1089#1074#1086#1077#1085#1080#1077' '#1052#1057#1041
      ImageIndex = 4
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
    end
    object TabSheet6: TTabSheet
      Caption = #1044#1086#1084#1086#1093#1086#1079#1103#1081#1089#1090#1074#1072
      ImageIndex = 5
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
    end
    object TabSheet8: TTabSheet
      Caption = #1054#1094#1077#1085#1082#1072
      ImageIndex = 6
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object StringGrid1: TStringGrid
        Left = 0
        Top = 0
        Width = 1476
        Height = 622
        Align = alClient
        ColCount = 4
        RowCount = 8
        TabOrder = 0
      end
    end
  end
  object MainMenu1: TMainMenu
    Left = 704
    Top = 65513
    object N11: TMenuItem
      Caption = #1060#1072#1081#1083#1099
      object N12: TMenuItem
        Caption = #1042#1099#1093#1086#1076
        OnClick = N12Click
      end
    end
    object N21: TMenuItem
      Caption = #1044#1088' '#1087#1091#1085#1082#1090' '#1084#1077#1085#1102
    end
  end
end
