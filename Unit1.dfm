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
    ActivePage = TabSheet9
    TabOrder = 0
    object TabSheet9: TTabSheet
      Caption = #1053#1072#1083#1072#1096#1090#1091#1074#1072#1085#1085#1103
      ImageIndex = 5
      object Label21: TLabel
        Left = 52
        Top = 16
        Width = 120
        Height = 13
        Caption = #1042#1082#1072#1078#1110#1090#1100' '#1085#1072#1074#1095#1072#1083#1100#1085#1080#1081' '#1088#1110#1082
      end
      object Label22: TLabel
        Left = 3
        Top = 133
        Width = 153
        Height = 13
        Caption = #1055#1077#1088#1077#1083#1110#1082' '#1091#1089#1110#1093' '#1085#1072#1074#1095#1072#1083#1100#1085#1080#1093' '#1088#1086#1082#1110#1074
      end
      object Label23: TLabel
        Left = 256
        Top = 16
        Width = 146
        Height = 13
        Caption = #1042#1082#1072#1078#1110#1090#1100' '#1085#1072#1074#1095#1072#1083#1100#1085#1080#1081' '#1089#1077#1084#1077#1089#1090#1088
      end
      object DateTimePicker7: TDateTimePicker
        Left = 44
        Top = 35
        Width = 65
        Height = 21
        Date = 42528.562290185180000000
        Format = 'yyyy'
        Time = 42528.562290185180000000
        DateMode = dmUpDown
        DoubleBuffered = False
        ParentDoubleBuffered = False
        TabOrder = 0
      end
      object DateTimePicker8: TDateTimePicker
        Left = 123
        Top = 35
        Width = 65
        Height = 21
        Date = 42528.564100289350000000
        Format = 'yyyy'
        Time = 42528.564100289350000000
        DateMode = dmUpDown
        TabOrder = 1
      end
      object Button23: TButton
        Left = 256
        Top = 88
        Width = 193
        Height = 81
        Caption = #1057#1090#1074#1086#1088#1080#1090#1080' '#1073#1072#1079#1091' '#1076#1072#1085#1080#1093
        TabOrder = 2
        OnClick = Button23Click
      end
      object Button24: TButton
        Left = 3
        Top = 463
        Width = 169
        Height = 25
        Caption = #1042#1080#1076#1072#1083#1080#1090#1080' '#1073#1072#1079#1091' '#1076#1072#1085#1080#1093
        TabOrder = 3
      end
      object DBGrid10: TDBGrid
        Left = 3
        Top = 152
        Width = 169
        Height = 305
        TabOrder = 4
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object RadioButton3: TRadioButton
        Left = 256
        Top = 35
        Width = 57
        Height = 17
        Caption = #1055#1077#1088#1096#1080#1081
        Checked = True
        TabOrder = 5
        TabStop = True
      end
      object RadioButton4: TRadioButton
        Left = 345
        Top = 35
        Width = 57
        Height = 17
        Caption = #1044#1088#1091#1075#1080#1081
        TabOrder = 6
      end
      object Button25: TButton
        Left = 272
        Top = 229
        Width = 177
        Height = 25
        Caption = #1042#1082#1072#1079#1072#1090#1080' Excel '#1092#1072#1081#1083
        TabOrder = 7
      end
    end
    object TabSheet1: TTabSheet
      Caption = #1043#1088#1091#1087#1080
      ExplicitWidth = 1081
      object Label12: TLabel
        Left = 38
        Top = 4
        Width = 33
        Height = 13
        Caption = #1055#1086#1096#1091#1082
      end
      object Button1: TButton
        Left = 124
        Top = 59
        Width = 145
        Height = 25
        Caption = #1042#1110#1076#1082#1088#1080#1090#1080' Exel '#1092#1072#1081#1083
        TabOrder = 0
        OnClick = Button1Click
      end
      object Button2: TButton
        Left = 460
        Top = 59
        Width = 145
        Height = 25
        Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1088#1086#1079#1082#1083#1072#1076
        TabOrder = 1
        OnClick = Button2Click
      end
      object Button3: TButton
        Left = 124
        Top = 121
        Width = 145
        Height = 25
        Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1095'/'#1079
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clBlack
        Font.Height = -11
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 2
        OnClick = Button3Click
      end
      object Button8: TButton
        Left = 124
        Top = 90
        Width = 145
        Height = 25
        Caption = #1044#1086#1076#1072#1090#1080' '#1088#1086#1079#1082#1083#1072#1076' (test)'
        TabOrder = 3
        OnClick = Button8Click
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
        TabOrder = 4
      end
      object DBNavigator2: TDBNavigator
        Left = 118
        Top = 152
        Width = 960
        Height = 18
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        TabOrder = 5
      end
      object DBGrid2: TDBGrid
        Left = 112
        Top = 176
        Width = 561
        Height = 354
        TabOrder = 6
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object DBNavigator1: TDBNavigator
        Left = 461
        Top = 90
        Width = 144
        Height = 20
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        Enabled = False
        TabOrder = 7
      end
      object Button5: TButton
        Left = 300
        Top = 90
        Width = 145
        Height = 25
        Caption = #1044#1086#1076#1072#1090#1080' '#1075#1088#1091#1087#1091
        TabOrder = 8
        OnClick = Button5Click
      end
      object Edit1: TEdit
        Left = 300
        Top = 63
        Width = 145
        Height = 21
        TabOrder = 9
      end
      object Button7: TButton
        Left = 300
        Top = 121
        Width = 145
        Height = 25
        Caption = #1042#1080#1076#1072#1083#1080#1090#1080' '#1075#1088#1091#1087#1091' (test)'
        TabOrder = 10
        OnClick = Button7Click
      end
      object DBGrid3: TDBGrid
        Left = 679
        Top = 176
        Width = 402
        Height = 354
        TabOrder = 11
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object DBGrid1: TDBGrid
        Left = 0
        Top = 50
        Width = 118
        Height = 479
        TabOrder = 12
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid1CellClick
      end
      object Edit6: TEdit
        Left = 3
        Top = 23
        Width = 115
        Height = 21
        TabOrder = 13
        OnChange = Edit6Change
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1042#1080#1082#1083#1072#1076#1072#1095#1110
      ImageIndex = 1
      ExplicitLeft = 1
      ExplicitTop = 28
      object Label11: TLabel
        Left = 49
        Top = 9
        Width = 33
        Height = 13
        Caption = #1055#1086#1096#1091#1082
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
        Top = 61
        Width = 118
        Height = 476
        TabOrder = 1
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object Button4: TButton
        Left = 140
        Top = 9
        Width = 146
        Height = 25
        Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1074#1080#1082#1083#1072#1076#1072#1095#1110#1074
        TabOrder = 2
        OnClick = Button4Click
      end
      object PageControl2: TPageControl
        Left = 120
        Top = 61
        Width = 961
        Height = 476
        ActivePage = TabSheet6
        TabOrder = 3
        object TabSheet6: TTabSheet
          Caption = #1047#1072#1084#1110#1085#1072
          ExplicitLeft = 0
          ExplicitTop = 0
          ExplicitWidth = 0
          ExplicitHeight = 0
          object Label13: TLabel
            Left = 57
            Top = 13
            Width = 26
            Height = 13
            Caption = #1044#1072#1090#1072
          end
          object Label14: TLabel
            Left = 53
            Top = 59
            Width = 30
            Height = 13
            Caption = #1043#1088#1091#1087#1072
          end
          object Label15: TLabel
            Left = 53
            Top = 105
            Width = 46
            Height = 13
            Caption = #8470' '#1083#1077#1085#1090#1080
          end
          object Label16: TLabel
            Left = 46
            Top = 183
            Width = 71
            Height = 13
            Caption = #1050#1080#1084' '#1079#1072#1084#1085#1102#1108#1084#1086
          end
          object Label17: TLabel
            Left = 49
            Top = 229
            Width = 44
            Height = 13
            Caption = #1055#1088#1077#1076#1084#1077#1090
          end
          object Label18: TLabel
            Left = 42
            Top = 267
            Width = 75
            Height = 13
            Caption = #1050#1086#1075#1086' '#1079#1072#1084#1110#1085#1103#1108#1084#1086
          end
          object Label19: TLabel
            Left = 49
            Top = 313
            Width = 44
            Height = 13
            Caption = #1055#1088#1077#1076#1084#1077#1090
          end
          object Label20: TLabel
            Left = 46
            Top = 370
            Width = 52
            Height = 13
            Caption = #1040#1091#1076#1080#1090#1086#1088#1110#1103
          end
          object DateTimePicker5: TDateTimePicker
            Left = 16
            Top = 32
            Width = 126
            Height = 21
            Date = 42526.829379166670000000
            Time = 42526.829379166670000000
            TabOrder = 0
          end
          object DBLookupComboBox1: TDBLookupComboBox
            Left = 16
            Top = 78
            Width = 126
            Height = 21
            TabOrder = 1
          end
          object ComboBox2: TComboBox
            Left = 16
            Top = 124
            Width = 126
            Height = 21
            Style = csDropDownList
            TabOrder = 2
            Items.Strings = (
              #1055#1077#1088#1096#1072
              #1044#1088#1091#1075#1072
              #1058#1088#1077#1090#1103
              #1063#1077#1090#1074#1077#1088#1090#1072
              #1055'`'#1103#1090#1072)
          end
          object DBLookupComboBox2: TDBLookupComboBox
            Left = 16
            Top = 202
            Width = 126
            Height = 21
            TabOrder = 3
          end
          object DBLookupComboBox3: TDBLookupComboBox
            Left = 16
            Top = 248
            Width = 126
            Height = 21
            TabOrder = 4
          end
          object DBLookupComboBox5: TDBLookupComboBox
            Left = 16
            Top = 332
            Width = 126
            Height = 21
            TabOrder = 5
          end
          object TEdit
            Left = 16
            Top = 389
            Width = 126
            Height = 21
            TabOrder = 6
          end
          object DBLookupComboBox4: TDBLookupComboBox
            Left = 16
            Top = 286
            Width = 126
            Height = 21
            TabOrder = 7
          end
          object DBGrid5: TDBGrid
            Left = 160
            Top = 78
            Width = 793
            Height = 371
            TabOrder = 8
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -11
            TitleFont.Name = 'Tahoma'
            TitleFont.Style = []
          end
          object DBNavigator5: TDBNavigator
            Left = 160
            Top = 43
            Width = 240
            Height = 18
            TabOrder = 9
          end
          object Button16: TButton
            Left = 16
            Top = 416
            Width = 126
            Height = 29
            Caption = #1044#1086#1076#1072#1090#1080
            TabOrder = 10
            OnClick = Button16Click
          end
          object ComboBox3: TComboBox
            Left = 16
            Top = 151
            Width = 126
            Height = 21
            Style = csDropDownList
            TabOrder = 11
            Items.Strings = (
              #1055#1077#1088#1096#1091
              #1044#1088#1091#1075#1091
              #1058#1088#1077#1090#1102
              #1063#1077#1090#1074#1077#1088#1090#1091
              #1055'`'#1103#1090#1091)
          end
          object Edit8: TEdit
            Left = 160
            Top = 16
            Width = 121
            Height = 21
            TabOrder = 12
            Text = 'Edit8'
          end
        end
        object TabSheet7: TTabSheet
          Caption = #1055#1088#1077#1076#1084#1077#1090' '#1085#1072' '#1087#1088#1077#1076#1084#1077#1090
          ImageIndex = 1
        end
        object TabSheet8: TTabSheet
          Caption = #1047#1072#1084#1110#1085#1072' '#1074#1110#1076#1087#1088#1072#1094#1102#1074#1072#1085#1085#1103#1084
          ImageIndex = 2
          ExplicitLeft = 0
          ExplicitTop = 0
          ExplicitWidth = 0
          ExplicitHeight = 0
          object ComboBox4: TComboBox
            Left = 3
            Top = 72
            Width = 121
            Height = 21
            TabOrder = 0
            Text = 'ComboBox4'
          end
          object DateTimePicker6: TDateTimePicker
            Left = 3
            Top = 29
            Width = 121
            Height = 21
            Date = 42527.576763888890000000
            Time = 42527.576763888890000000
            TabOrder = 1
          end
          object DBLookupComboBox6: TDBLookupComboBox
            Left = 3
            Top = 112
            Width = 121
            Height = 21
            TabOrder = 2
          end
          object DBLookupComboBox7: TDBLookupComboBox
            Left = 3
            Top = 139
            Width = 121
            Height = 21
            TabOrder = 3
          end
          object Edit7: TEdit
            Left = 3
            Top = 176
            Width = 121
            Height = 21
            TabOrder = 4
            Text = 'Edit7'
          end
        end
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
        Top = 145
        Width = 129
        Height = 13
        Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1087#1086#1095#1072#1090#1082#1086#1074#1091' '#1076#1072#1090#1091
      end
      object Label2: TLabel
        Left = 177
        Top = 199
        Width = 113
        Height = 13
        Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1082#1110#1085#1094#1077#1074#1091' '#1076#1072#1090#1091
      end
      object Label3: TLabel
        Left = 203
        Top = 261
        Width = 70
        Height = 13
        Caption = #1042#1080#1076' '#1085#1072#1074#1095#1072#1085#1085#1103
      end
      object Label4: TLabel
        Left = 203
        Top = 316
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
        Top = 218
        Width = 186
        Height = 21
        Date = 42525.558961805550000000
        Time = 42525.558961805550000000
        TabOrder = 0
      end
      object Edit3: TEdit
        Left = 140
        Top = 335
        Width = 186
        Height = 21
        TabOrder = 1
      end
      object ComboBox1: TComboBox
        Left = 140
        Top = 280
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
        Top = 372
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
        Top = 164
        Width = 186
        Height = 21
        Date = 42525.559036608790000000
        Time = 42525.559036608790000000
        TabOrder = 8
      end
      object ListBox2: TListBox
        Left = 140
        Top = 49
        Width = 186
        Height = 61
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
        Top = 101
        Width = 186
        Height = 25
        Caption = #1054#1095#1080#1089#1090#1080#1090#1080
        TabOrder = 11
        OnClick = Button6Click
      end
    end
    object TabSheet4: TTabSheet
      Caption = #1056#1086#1079#1082#1083#1072#1076
      ImageIndex = 3
      ExplicitWidth = 1081
      object MonthCalendar1: TMonthCalendar
        Left = 3
        Top = 64
        Width = 191
        Height = 160
        Date = 42528.537816817130000000
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
        OnDblClick = DBGrid8DblClick
      end
      object Button13: TButton
        Left = 352
        Top = 52
        Width = 193
        Height = 25
        Caption = #1055#1077#1088#1077#1083#1110#1082' '#1089#1091#1073#1086#1090' ('#1085#1077' '#1088#1072#1073')'
        TabOrder = 13
        OnClick = Button13Click
      end
      object Button14: TButton
        Left = 352
        Top = 83
        Width = 193
        Height = 25
        Caption = #1055#1077#1088#1077#1083#1110#1082' '#1076#1085#1110#1074' '#1079' '#1079#1072#1084#1110#1085#1072#1084#1080' ('#1085#1077' '#1088#1072#1073')'
        TabOrder = 14
        OnClick = Button14Click
      end
      object Button15: TButton
        Left = 551
        Top = 68
        Width = 193
        Height = 25
        Caption = 'C'#1082#1080#1085#1091#1090#1080' '#1092#1110#1083#1100#1090#1088#1080'('#1085#1077' '#1088#1072#1073')'
        TabOrder = 15
      end
      object Edit9: TEdit
        Left = 248
        Top = 30
        Width = 121
        Height = 21
        TabOrder = 16
        Text = 'Edit9'
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
    ExplicitLeft = 4
    ExplicitTop = 590
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
end
