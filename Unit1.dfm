object Form1: TForm1
  Left = 244
  Top = 110
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Form1'
  ClientHeight = 587
  ClientWidth = 1087
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  Position = poDesigned
  OnCreate = Form1Create
  OnDestroy = Form1Destroy
  PixelsPerInch = 96
  TextHeight = 13
  object PageControl1: TPageControl
    Left = 0
    Top = 1
    Width = 1097
    Height = 561
    ActivePage = TabSheet1
    Style = tsFlatButtons
    TabOrder = 0
    object TabSheet9: TTabSheet
      Caption = #1053#1072#1083#1072#1096#1090#1091#1074#1072#1085#1085#1103
      ImageIndex = 5
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object Label22: TLabel
        Left = 3
        Top = 160
        Width = 153
        Height = 13
        Caption = #1055#1077#1088#1077#1083#1110#1082' '#1091#1089#1110#1093' '#1085#1072#1074#1095#1072#1083#1100#1085#1080#1093' '#1088#1086#1082#1110#1074
      end
      object Button24: TButton
        Left = 3
        Top = 503
        Width = 154
        Height = 25
        Caption = #1042#1080#1076#1072#1083#1080#1090#1080' '#1073#1072#1079#1091' '#1076#1072#1085#1080#1093
        TabOrder = 0
      end
      object DBGrid10: TDBGrid
        Left = 3
        Top = 179
        Width = 154
        Height = 318
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid10CellClick
        OnDrawColumnCell = DBGrid10DrawColumnCell
        OnDblClick = DBGrid10DblClick
      end
      object GroupBox1: TGroupBox
        Left = 3
        Top = 3
        Width = 380
        Height = 151
        Caption = #1044#1086#1076#1072#1074#1072#1085#1085#1103' '#1073#1072#1079#1080' '#1085#1086#1074#1086#1075#1086' '#1085#1072#1074#1095#1072#1083#1100#1085#1086#1075#1086' '#1088#1086#1082#1091
        Color = clSkyBlue
        ParentBackground = False
        ParentColor = False
        TabOrder = 2
        object Label21: TLabel
          Left = 47
          Top = 51
          Width = 120
          Height = 13
          Caption = #1042#1082#1072#1078#1110#1090#1100' '#1085#1072#1074#1095#1072#1083#1100#1085#1080#1081' '#1088#1110#1082
        end
        object ComboBox5: TComboBox
          Left = 31
          Top = 70
          Width = 145
          Height = 22
          Style = csOwnerDrawFixed
          ItemIndex = 0
          TabOrder = 0
          Text = '2016 - 2017'
          Items.Strings = (
            '2016 - 2017'
            '2017 - 2018'
            '2018 - 2019'
            '2019 - 2020'
            '2020 - 2021'
            '2021 - 2022'
            '2022 - 2023'
            '2023 - 2024'
            '2024 - 2025'
            '2025 - 2026'
            '2026 - 2027')
        end
        object Button23: TButton
          Left = 222
          Top = 61
          Width = 145
          Height = 31
          Caption = #1057#1090#1074#1086#1088#1080#1090#1080' '#1073#1072#1079#1091' '#1076#1072#1085#1080#1093
          TabOrder = 1
          OnClick = Button23Click
        end
      end
      object GroupBox2: TGroupBox
        Left = 405
        Top = 3
        Width = 380
        Height = 151
        Caption = #1044#1086#1076#1072#1074#1072#1085#1085#1103' '#1096#1083#1103#1093#1091' '#1076#1086' '#1074#1093#1110#1076#1085#1086#1075#1086' Excel '#1060#1072#1081#1083#1091
        Color = clSkyBlue
        ParentBackground = False
        ParentColor = False
        TabOrder = 3
        object Label23: TLabel
          Left = 111
          Top = 47
          Width = 148
          Height = 13
          Caption = #1042#1082#1072#1078#1110#1090#1100' '#1096#1083#1103#1093' '#1076#1086' Excel '#1092#1072#1081#1083#1091
        end
        object Edit10: TEdit
          Left = 67
          Top = 66
          Width = 241
          Height = 21
          TabOrder = 0
        end
        object Button25: TButton
          Left = 98
          Top = 109
          Width = 169
          Height = 25
          Caption = #1042#1082#1072#1079#1072#1090#1080' '#1096#1083#1103#1093' '
          TabOrder = 1
          OnClick = Button25Click
        end
      end
      object GroupBox3: TGroupBox
        Left = 362
        Top = 179
        Width = 551
        Height = 367
        Caption = #1042#1110#1076#1082#1072#1090' '#1079#1084#1110#1085
        Color = clSkyBlue
        ParentBackground = False
        ParentColor = False
        TabOrder = 4
        object Label24: TLabel
          Left = 80
          Top = 19
          Width = 110
          Height = 13
          Caption = #1055#1086#1090#1086#1095#1085#1072' '#1073#1072#1079#1072' '#1076#1072#1085#1080#1093': '
        end
        object Label25: TLabel
          Left = 310
          Top = 109
          Width = 205
          Height = 13
          Caption = #1042#1082#1072#1078#1110#1090#1100' '#1091' '#1075#1086#1076#1080#1085#1072#1093' '#1110#1085#1090#1077#1088#1074#1072#1083' '#1079#1073#1077#1088#1077#1078#1077#1085#1085#1103
        end
        object Button26: TButton
          Left = 310
          Top = 36
          Width = 188
          Height = 25
          Caption = #1047#1073#1077#1088#1110#1075#1090#1080', '#1103#1082' '#1087#1086#1090#1086#1095#1085#1091' '#1073#1072#1079#1091' '#1076#1072#1085#1080#1093
          TabOrder = 0
          OnClick = Button26Click
        end
        object DBGrid11: TDBGrid
          Left = 16
          Top = 72
          Width = 265
          Height = 287
          TabOrder = 1
          TitleFont.Charset = DEFAULT_CHARSET
          TitleFont.Color = clWindowText
          TitleFont.Height = -11
          TitleFont.Name = 'Tahoma'
          TitleFont.Style = []
          OnDrawColumnCell = DBGrid11DrawColumnCell
        end
        object Edit7: TEdit
          Left = 55
          Top = 38
          Width = 153
          Height = 21
          TabOrder = 2
        end
        object Button29: TButton
          Left = 342
          Top = 247
          Width = 145
          Height = 25
          Caption = #1042#1110#1076#1082#1072#1090#1080#1090#1080' '#1073#1072#1079#1091' '#1076#1072#1085#1080#1093
          TabOrder = 3
          OnClick = Button29Click
        end
        object CheckBox1: TCheckBox
          Left = 333
          Top = 77
          Width = 159
          Height = 17
          Caption = #1059#1074#1110#1084#1082#1085#1091#1090#1080' '#1072#1074#1090#1086#1079#1073#1077#1088#1077#1078#1077#1085#1085#1103
          TabOrder = 4
          OnClick = CheckBox1Click
        end
        object SpinEdit1: TSpinEdit
          Left = 354
          Top = 130
          Width = 121
          Height = 22
          MaxLength = 3
          MaxValue = 200
          MinValue = 1
          TabOrder = 5
          Value = 1
        end
        object Button28: TButton
          Left = 342
          Top = 198
          Width = 145
          Height = 25
          Caption = #1047#1073#1077#1088#1110#1075#1090#1080' '#1073#1072#1079#1091' '#1076#1072#1085#1080#1093
          TabOrder = 6
          OnClick = Button28Click
        end
        object Button27: TButton
          Left = 342
          Top = 320
          Width = 147
          Height = 25
          Caption = #1042#1080#1076#1072#1083#1080#1090#1080' '#1079#1072#1087#1080#1089
          TabOrder = 7
          OnClick = Button27Click
        end
      end
    end
    object TabSheet1: TTabSheet
      Caption = #1043#1088#1091#1087#1080
      object Label12: TLabel
        Left = 38
        Top = 4
        Width = 33
        Height = 13
        Caption = #1055#1086#1096#1091#1082
      end
      object Button1: TButton
        Left = 295
        Top = 59
        Width = 173
        Height = 25
        Caption = #1042#1110#1076#1082#1088#1080#1090#1080' Excel '#1092#1072#1081#1083
        TabOrder = 0
        OnClick = Button1Click
      end
      object Panel1: TPanel
        Left = 124
        Top = 0
        Width = 964
        Height = 44
        Caption = #1043#1088#1091#1087#1072' '#1085#1077' '#1074#1080#1073#1088#1072#1085#1072
        Color = clActiveCaption
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -24
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentBackground = False
        ParentFont = False
        TabOrder = 1
      end
      object DBNavigator2: TDBNavigator
        Left = 118
        Top = 152
        Width = 960
        Height = 18
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        TabOrder = 2
      end
      object DBGrid1: TDBGrid
        Left = 0
        Top = 50
        Width = 118
        Height = 479
        TabOrder = 3
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid1CellClick
        OnDrawColumnCell = DBGrid1DrawColumnCell
      end
      object Edit6: TEdit
        Left = 3
        Top = 23
        Width = 115
        Height = 21
        TabOrder = 4
        OnChange = Edit6Change
      end
      object DBGrid13: TDBGrid
        Left = 118
        Top = 169
        Width = 960
        Height = 358
        Color = clBtnFace
        GradientEndColor = clSilver
        TabOrder = 5
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = DBGrid13DrawColumnCell
      end
      object Button30: TButton
        Left = 295
        Top = 90
        Width = 173
        Height = 25
        Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080
        TabOrder = 6
        OnClick = Button30Click
      end
      object RadioGroup2: TRadioGroup
        Left = 124
        Top = 50
        Width = 165
        Height = 96
        Caption = #1042#1080#1073#1110#1088' '#1089#1077#1084#1077#1089#1090#1088#1091
        Items.Strings = (
          #1055#1077#1088#1096#1080#1081' '#1089#1077#1084#1077#1089#1090#1088
          #1044#1088#1091#1075#1080#1081' '#1089#1077#1084#1077#1089#1090#1088
          #1059#1074#1077#1089#1100' '#1088#1110#1082)
        TabOrder = 7
      end
      object Button3: TButton
        Left = 295
        Top = 121
        Width = 173
        Height = 25
        Caption = #1055#1086#1082#1072#1079#1072#1090#1080' '#1087#1077#1088#1077#1083#1110#1082' '#1091#1089#1110#1093' '#1075#1088#1091#1087
        TabOrder = 8
        OnClick = Button3Click
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1042#1080#1082#1083#1072#1076#1072#1095#1110
      ImageIndex = 1
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object Label11: TLabel
        Left = 46
        Top = 9
        Width = 33
        Height = 13
        Caption = #1055#1086#1096#1091#1082
      end
      object Label13: TLabel
        Left = 360
        Top = 9
        Width = 26
        Height = 13
        Caption = #1044#1072#1090#1072
      end
      object Label14: TLabel
        Left = 360
        Top = 55
        Width = 30
        Height = 13
        Caption = #1043#1088#1091#1087#1072
      end
      object Label15: TLabel
        Left = 852
        Top = 9
        Width = 46
        Height = 13
        Caption = #8470' '#1083#1077#1085#1090#1080
      end
      object Label16: TLabel
        Left = 520
        Top = 9
        Width = 71
        Height = 13
        Caption = #1050#1080#1084' '#1079#1072#1084#1085#1102#1108#1084#1086
      end
      object Label17: TLabel
        Left = 520
        Top = 58
        Width = 44
        Height = 13
        Caption = #1055#1088#1077#1076#1084#1077#1090
      end
      object Label18: TLabel
        Left = 665
        Top = 9
        Width = 75
        Height = 13
        Caption = #1050#1086#1075#1086' '#1079#1072#1084#1110#1085#1103#1108#1084#1086
      end
      object Label19: TLabel
        Left = 678
        Top = 58
        Width = 44
        Height = 13
        Caption = #1055#1088#1077#1076#1084#1077#1090
      end
      object Label20: TLabel
        Left = 342
        Top = 101
        Width = 52
        Height = 13
        Caption = #1040#1091#1076#1080#1090#1086#1088#1110#1103
      end
      object Label26: TLabel
        Left = 511
        Top = 104
        Width = 69
        Height = 13
        Caption = #1053#1086#1084#1077#1088' '#1085#1072#1082#1072#1079#1091
      end
      object Label27: TLabel
        Left = 678
        Top = 104
        Width = 49
        Height = 13
        Caption = #1050#1086#1084#1077#1085#1090#1072#1088
      end
      object Edit2: TEdit
        Left = 0
        Top = 28
        Width = 121
        Height = 21
        TabOrder = 0
        OnChange = Edit2Change
      end
      object DBGrid4: TDBGrid
        Left = 0
        Top = 58
        Width = 118
        Height = 476
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = DBGrid4DrawColumnCell
      end
      object Button16: TButton
        Left = 976
        Top = 28
        Width = 97
        Height = 108
        Caption = #1044#1086#1076#1072#1090#1080
        TabOrder = 2
        OnClick = Button16Click
      end
      object ComboBox2: TComboBox
        Left = 807
        Top = 28
        Width = 146
        Height = 21
        Style = csDropDownList
        TabOrder = 3
        Items.Strings = (
          #1055#1077#1088#1096#1072
          #1044#1088#1091#1075#1072
          #1058#1088#1077#1090#1103
          #1063#1077#1090#1074#1077#1088#1090#1072
          #1055'`'#1103#1090#1072)
      end
      object ComboBox3: TComboBox
        Left = 807
        Top = 74
        Width = 146
        Height = 21
        Style = csDropDownList
        TabOrder = 4
        Items.Strings = (
          #1055#1077#1088#1096#1091
          #1044#1088#1091#1075#1091
          #1058#1088#1077#1090#1102
          #1063#1077#1090#1074#1077#1088#1090#1091
          #1055'`'#1103#1090#1091)
      end
      object DateTimePicker5: TDateTimePicker
        Left = 303
        Top = 28
        Width = 146
        Height = 21
        Date = 42526.829379166670000000
        Time = 42526.829379166670000000
        TabOrder = 5
        OnCloseUp = DateTimePicker5CloseUp
      end
      object DBGrid5: TDBGrid
        Left = 117
        Top = 176
        Width = 960
        Height = 358
        TabOrder = 6
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = DBGrid5DrawColumnCell
      end
      object DBLookupComboBox1: TDBLookupComboBox
        Left = 303
        Top = 74
        Width = 146
        Height = 21
        TabOrder = 7
        OnCloseUp = DBLookupComboBox1CloseUp
      end
      object DBLookupComboBox2: TDBLookupComboBox
        Left = 471
        Top = 28
        Width = 146
        Height = 21
        TabOrder = 8
      end
      object DBLookupComboBox3: TDBLookupComboBox
        Left = 471
        Top = 77
        Width = 146
        Height = 21
        TabOrder = 9
      end
      object DBLookupComboBox4: TDBLookupComboBox
        Left = 639
        Top = 28
        Width = 146
        Height = 21
        TabOrder = 10
      end
      object DBLookupComboBox5: TDBLookupComboBox
        Left = 639
        Top = 77
        Width = 146
        Height = 21
        TabOrder = 11
      end
      object Edit11: TEdit
        Left = 303
        Top = 120
        Width = 146
        Height = 21
        TabOrder = 12
      end
      object DBNavigator5: TDBNavigator
        Left = 117
        Top = 158
        Width = 960
        Height = 18
        TabOrder = 13
      end
      object Edit12: TEdit
        Left = 471
        Top = 123
        Width = 146
        Height = 21
        TabOrder = 14
      end
      object Edit13: TEdit
        Left = 639
        Top = 123
        Width = 146
        Height = 21
        TabOrder = 15
      end
      object RadioGroup1: TRadioGroup
        Left = 127
        Top = 28
        Width = 162
        Height = 108
        Caption = #1042#1080#1073#1110#1088' '#1090#1080#1087#1091' '#1079#1072#1084#1110#1085#1080
        Items.Strings = (
          #1047#1072#1084#1110#1085#1072
          #1055#1088#1077#1076#1084#1077#1090'-'#1087#1088#1077#1076#1084#1077#1090#1086#1084
          #1055#1086#1075#1086#1076#1080#1085#1085#1086
          #1047#1072#1084#1110#1085#1072' '#1074#1110#1076#1087#1088#1072#1094#1102#1074#1072#1085#1085#1103#1084)
        TabOrder = 16
        OnClick = RadioGroup1Click
      end
    end
    object TabSheet3: TTabSheet
      Caption = #1043#1088#1072#1092#1110#1082' '#1085#1072#1074#1095'. '#1087#1088#1086#1094#1077#1089#1091
      ImageIndex = 2
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object Label1: TLabel
        Left = 177
        Top = 231
        Width = 129
        Height = 13
        Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1087#1086#1095#1072#1090#1082#1086#1074#1091' '#1076#1072#1090#1091
      end
      object Label2: TLabel
        Left = 177
        Top = 287
        Width = 113
        Height = 13
        Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1082#1110#1085#1094#1077#1074#1091' '#1076#1072#1090#1091
      end
      object Label3: TLabel
        Left = 203
        Top = 333
        Width = 70
        Height = 13
        Caption = #1042#1080#1076' '#1085#1072#1074#1095#1072#1085#1085#1103
      end
      object Label4: TLabel
        Left = 203
        Top = 393
        Width = 55
        Height = 13
        Caption = #1050#1086#1084#1084#1077#1085#1090#1072#1088
      end
      object Label10: TLabel
        Left = 36
        Top = 3
        Width = 33
        Height = 13
        Caption = #1055#1086#1096#1091#1082
      end
      object DateTimePicker1: TDateTimePicker
        Left = 140
        Top = 306
        Width = 186
        Height = 21
        Date = 42525.558961805550000000
        Time = 42525.558961805550000000
        TabOrder = 0
      end
      object Edit3: TEdit
        Left = 140
        Top = 423
        Width = 186
        Height = 21
        TabOrder = 1
      end
      object ComboBox1: TComboBox
        Left = 140
        Top = 352
        Width = 186
        Height = 22
        Style = csOwnerDrawFixed
        TabOrder = 2
        Items.Strings = (
          #1055#1088#1072#1082#1090#1080#1082#1072
          #1050#1072#1085#1110#1082#1091#1083#1080
          #1057#1077#1089#1110#1103)
      end
      object Button9: TButton
        Left = 140
        Top = 458
        Width = 186
        Height = 53
        Caption = #1044#1086#1076#1072#1090#1080
        TabOrder = 3
        OnClick = Button9Click
      end
      object DBGrid6: TDBGrid
        Left = 0
        Top = 49
        Width = 118
        Height = 480
        TabOrder = 4
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid6CellClick
        OnDrawColumnCell = DBGrid6DrawColumnCell
        OnDblClick = DBGrid6DblClick
      end
      object DBGrid7: TDBGrid
        Left = 352
        Top = 128
        Width = 728
        Height = 402
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = DBGrid7DrawColumnCell
      end
      object Panel2: TPanel
        Left = 124
        Top = 0
        Width = 964
        Height = 43
        Caption = #1043#1088#1091#1087#1072' '#1085#1077' '#1074#1080#1073#1088#1072#1085#1072
        Color = clActiveCaption
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -24
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentBackground = False
        ParentFont = False
        TabOrder = 6
      end
      object DBNavigator7: TDBNavigator
        Left = 352
        Top = 104
        Width = 720
        Height = 18
        TabOrder = 7
      end
      object DateTimePicker2: TDateTimePicker
        Left = 140
        Top = 260
        Width = 186
        Height = 21
        Date = 42525.559036608790000000
        Time = 42525.559036608790000000
        TabOrder = 8
      end
      object ListBox2: TListBox
        Left = 140
        Top = 88
        Width = 186
        Height = 106
        ItemHeight = 13
        TabOrder = 9
        OnDblClick = ListBox2DblClick
      end
      object Edit5: TEdit
        Left = 0
        Top = 22
        Width = 118
        Height = 21
        TabOrder = 10
        OnChange = Edit5Change
      end
      object Button6: TButton
        Left = 140
        Top = 200
        Width = 186
        Height = 25
        Caption = #1054#1095#1080#1089#1090#1080#1090#1080
        TabOrder = 11
        OnClick = Button6Click
      end
      object Button32: TButton
        Left = 376
        Top = 64
        Width = 75
        Height = 25
        Caption = 'Button32'
        TabOrder = 12
      end
      object Button2: TButton
        Left = 140
        Top = 64
        Width = 186
        Height = 25
        Caption = #1044#1086#1076#1072#1090#1080' '#1091#1089#1110' '#1075#1088#1091#1087#1080
        TabOrder = 13
        OnClick = Button2Click
      end
    end
    object TabSheet4: TTabSheet
      Caption = #1056#1086#1079#1082#1083#1072#1076
      ImageIndex = 3
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object MonthCalendar1: TMonthCalendar
        Left = 3
        Top = 64
        Width = 191
        Height = 160
        Date = 42528.695179143520000000
        TabOrder = 0
      end
    end
    object TabSheet5: TTabSheet
      Caption = #1047#1072#1084#1110#1085#1080
      ImageIndex = 4
      ExplicitLeft = 0
      ExplicitTop = 0
      ExplicitWidth = 0
      ExplicitHeight = 0
      object Label7: TLabel
        Left = 163
        Top = 57
        Width = 139
        Height = 13
        Caption = #1043#1088#1091#1087#1080' '#1076#1083#1103' '#1076#1086#1076#1072#1074#1072#1085#1085#1103' '#1079#1072#1084#1110#1085
      end
      object Label6: TLabel
        Left = 15
        Top = 29
        Width = 87
        Height = 13
        Caption = #1055#1077#1088#1077#1083#1110#1082' '#1091#1089#1110#1093' '#1075#1088#1091#1087
      end
      object Label9: TLabel
        Left = 175
        Top = 296
        Width = 111
        Height = 13
        Caption = #1044#1077#1085#1100' '#1103#1082#1080#1084' '#1079#1072#1084#1110#1085#1102#1102#1090#1100
      end
      object Label8: TLabel
        Left = 199
        Top = 415
        Width = 49
        Height = 13
        Caption = #1050#1086#1084#1077#1085#1090#1072#1088
      end
      object Label5: TLabel
        Left = 175
        Top = 234
        Width = 111
        Height = 13
        Caption = #1044#1077#1085#1100' '#1103#1082#1080#1081' '#1079#1072#1084#1110#1085#1102#1102#1090#1100
      end
      object ListBox1: TListBox
        Left = 135
        Top = 100
        Width = 186
        Height = 97
        ItemHeight = 13
        TabOrder = 0
        OnDblClick = ListBox1DblClick
      end
      object DateTimePicker4: TDateTimePicker
        Left = 135
        Top = 315
        Width = 186
        Height = 21
        Date = 42526.483374201390000000
        Time = 42526.483374201390000000
        TabOrder = 1
      end
      object DateTimePicker3: TDateTimePicker
        Left = 135
        Top = 253
        Width = 186
        Height = 21
        Date = 42525.805654305560000000
        Time = 42525.805654305560000000
        TabOrder = 2
      end
      object RadioButton1: TRadioButton
        Left = 135
        Top = 350
        Width = 113
        Height = 17
        Caption = #1056#1086#1073#1086#1095#1072' '#1089#1091#1073#1086#1090#1072
        Checked = True
        TabOrder = 3
        TabStop = True
      end
      object RadioButton2: TRadioButton
        Left = 135
        Top = 381
        Width = 113
        Height = 17
        Caption = #1044#1077#1085#1100' '#1085#1072' '#1076#1077#1085#1100
        TabOrder = 4
      end
      object DBGrid9: TDBGrid
        Left = 352
        Top = 172
        Width = 731
        Height = 358
        Font.Charset = ANSI_CHARSET
        Font.Color = clWindowText
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 5
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnDrawColumnCell = DBGrid9DrawColumnCell
      end
      object Button10: TButton
        Left = 135
        Top = 477
        Width = 186
        Height = 39
        Caption = #1044#1086#1076#1072#1090#1080
        TabOrder = 6
        OnClick = Button10Click
      end
      object Button11: TButton
        Left = 135
        Top = 76
        Width = 186
        Height = 25
        Caption = #1044#1086#1076#1072#1090#1080' '#1091#1089#1110' '#1075#1088#1091#1087#1080
        TabOrder = 7
        OnClick = Button11Click
      end
      object Panel3: TPanel
        Left = 124
        Top = -1
        Width = 957
        Height = 43
        Caption = #1043#1088#1091#1087#1072' '#1085#1077' '#1074#1080#1073#1088#1072#1085#1072
        Color = clActiveCaption
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -24
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentBackground = False
        ParentFont = False
        TabOrder = 8
      end
      object Button12: TButton
        Left = 135
        Top = 196
        Width = 186
        Height = 25
        Caption = #1054#1095#1080#1089#1090#1080#1090#1080
        TabOrder = 9
        OnClick = Button12Click
      end
      object DBNavigator9: TDBNavigator
        Left = 352
        Top = 148
        Width = 720
        Height = 18
        TabOrder = 10
      end
      object Edit4: TEdit
        Left = 135
        Top = 434
        Width = 186
        Height = 21
        TabOrder = 11
      end
      object DBGrid8: TDBGrid
        Left = -4
        Top = 48
        Width = 118
        Height = 479
        TabOrder = 12
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid8CellClick
        OnDrawColumnCell = DBGrid8DrawColumnCell
        OnDblClick = DBGrid8DblClick
      end
    end
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 568
    Width = 1087
    Height = 19
    Panels = <
      item
        Text = #1044#1072#1090#1072
        Width = 100
      end
      item
        Text = #1042#1088#1077#1084#1103
        Width = 50
      end>
    OnHint = StatusBar1Hint
  end
  object Button17: TButton
    Left = 824
    Top = 254
    Width = 161
    Height = 25
    Caption = #1043#1088#1091#1087#1080
    TabOrder = 2
    Visible = False
  end
  object Button18: TButton
    Left = 824
    Top = 285
    Width = 161
    Height = 25
    Caption = 'Button18'
    TabOrder = 3
    Visible = False
  end
  object Button19: TButton
    Left = 824
    Top = 316
    Width = 161
    Height = 25
    Caption = 'Button19'
    TabOrder = 4
    Visible = False
  end
  object Button20: TButton
    Left = 824
    Top = 347
    Width = 161
    Height = 25
    Caption = 'Button20'
    TabOrder = 5
    Visible = False
  end
  object Button21: TButton
    Left = 824
    Top = 378
    Width = 161
    Height = 25
    Caption = 'Button21'
    TabOrder = 6
    Visible = False
  end
  object Button22: TButton
    Left = 824
    Top = 409
    Width = 161
    Height = 25
    Caption = 'Button22'
    TabOrder = 7
    Visible = False
  end
  object ADOConnection1: TADOConnection
    Left = 616
    Top = 65520
  end
  object DataSource1: TDataSource
    Left = 648
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 680
  end
  object ADOConnection2: TADOConnection
    Left = 616
  end
  object DataSource2: TDataSource
    Left = 648
  end
  object ADOQuery2: TADOQuery
    Parameters = <>
    Left = 680
  end
  object ADOConnection3: TADOConnection
    Left = 616
  end
  object DataSource3: TDataSource
    Left = 648
  end
  object ADOQuery3: TADOQuery
    Parameters = <>
    Left = 680
  end
  object ADOConnection4: TADOConnection
    Left = 616
  end
  object DataSource4: TDataSource
    Left = 648
  end
  object ADOQuery4: TADOQuery
    Parameters = <>
    Left = 680
  end
  object ADOConnection5: TADOConnection
    Left = 616
  end
  object DataSource5: TDataSource
    Left = 648
  end
  object ADOQuery5: TADOQuery
    Parameters = <>
    Left = 680
  end
  object OpenDialog1: TOpenDialog
    Left = 548
  end
  object ADOConnection7: TADOConnection
    Left = 764
  end
  object DataSource7: TDataSource
    Left = 796
  end
  object ADOQuery7: TADOQuery
    Parameters = <>
    Left = 828
  end
  object ADOConnection9: TADOConnection
    Left = 760
  end
  object DataSource9: TDataSource
    Left = 792
  end
  object ADOQuery9: TADOQuery
    Parameters = <>
    Left = 824
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 716
    Top = 6
  end
  object MainMenu1: TMainMenu
    Left = 504
    object N1: TMenuItem
      Caption = #1043#1086#1083#1086#1074#1085#1077' '#1084#1077#1085#1102
      OnClick = N1Click
    end
    object N7: TMenuItem
      Caption = #1053#1072#1083#1072#1096#1090#1091#1074#1072#1085#1085#1103
    end
    object N8: TMenuItem
      Caption = #1055#1088#1086' '#1087#1088#1086#1075#1088#1072#1084#1091
    end
    object N9: TMenuItem
      Caption = #1044#1086#1074#1110#1076#1082#1072
    end
    object N10: TMenuItem
      Caption = #1042#1080#1093#1110#1076
    end
  end
  object ADOConnection10: TADOConnection
    Left = 84
    Top = 457
  end
  object ADOQuery10: TADOQuery
    Parameters = <>
    Left = 44
    Top = 505
  end
  object DataSource10: TDataSource
    Left = 20
    Top = 465
  end
  object Timer2: TTimer
    Enabled = False
    OnTimer = Timer2Timer
    Left = 68
    Top = 268
  end
  object DataSource11: TDataSource
    Left = 192
    Top = 472
  end
  object ADOQuery11: TADOQuery
    Parameters = <>
    Left = 252
    Top = 500
  end
  object ADOConnection11: TADOConnection
    Left = 272
    Top = 448
  end
  object ADOConnection13: TADOConnection
    Left = 168
    Top = 248
  end
  object ADOQuery13: TADOQuery
    Parameters = <>
    Left = 176
    Top = 312
  end
  object DataSource13: TDataSource
    Left = 248
    Top = 296
  end
end
