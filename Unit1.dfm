object Form1: TForm1
  Left = 439
  Top = 219
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Form1'
  ClientHeight = 564
  ClientWidth = 895
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  Position = poDesigned
  OnCreate = Form1Create
  OnDestroy = Form1Destroy
  PixelsPerInch = 96
  TextHeight = 13
  object SpeedButton1: TSpeedButton
    Left = 32
    Top = 8
    Width = 23
    Height = 22
  end
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 887
    Height = 561
    ActivePage = TabSheet5
    TabOrder = 0
    object TabSheet1: TTabSheet
      Caption = #1043#1088#1091#1087#1080
      ExplicitLeft = -28
      ExplicitTop = 28
      object Button1: TButton
        Left = 0
        Top = 3
        Width = 145
        Height = 25
        Caption = #1042#1110#1076#1082#1088#1080#1090#1080' Exel '#1092#1072#1081#1083
        TabOrder = 0
        OnClick = Button1Click
      end
      object Button2: TButton
        Left = 151
        Top = 3
        Width = 146
        Height = 25
        Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1088#1086#1079#1082#1083#1072#1076
        TabOrder = 1
        OnClick = Button2Click
      end
      object Button3: TButton
        Left = 303
        Top = 3
        Width = 146
        Height = 25
        Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1095'/'#1079
        TabOrder = 2
        OnClick = Button3Click
      end
      object Button8: TButton
        Left = 151
        Top = 34
        Width = 146
        Height = 25
        Caption = #1053#1086#1074#1086#1077' '#1076#1086#1073#1072#1074#1083#1077#1085#1080#1077
        TabOrder = 3
        OnClick = Button8Click
      end
      object Panel1: TPanel
        Left = 150
        Top = 61
        Width = 570
        Height = 40
        Caption = #1043#1088#1091#1087#1072' '#1085#1077' '#1074#1080#1073#1088#1072#1085#1072
        Font.Charset = DEFAULT_CHARSET
        Font.Color = clWindowText
        Font.Height = -24
        Font.Name = 'Tahoma'
        Font.Style = []
        ParentFont = False
        TabOrder = 4
      end
      object DBNavigator2: TDBNavigator
        Left = 150
        Top = 107
        Width = 570
        Height = 20
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        TabOrder = 5
      end
      object DBGrid2: TDBGrid
        Left = 151
        Top = 133
        Width = 570
        Height = 294
        TabOrder = 6
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object DBNavigator1: TDBNavigator
        Left = 0
        Top = 92
        Width = 144
        Height = 19
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        Enabled = False
        TabOrder = 7
      end
      object Button5: TButton
        Left = 0
        Top = 61
        Width = 145
        Height = 25
        Caption = #1044#1086#1076#1072#1090#1080' '#1075#1088#1091#1087#1091
        TabOrder = 8
        OnClick = Button5Click
      end
      object Edit1: TEdit
        Left = 0
        Top = 34
        Width = 145
        Height = 21
        TabOrder = 9
      end
      object Button7: TButton
        Left = 0
        Top = 507
        Width = 145
        Height = 25
        Caption = 'delete_test'
        TabOrder = 10
        OnClick = Button7Click
      end
      object DBGrid3: TDBGrid
        Left = 151
        Top = 433
        Width = 570
        Height = 120
        TabOrder = 11
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object DBGrid1: TDBGrid
        Left = 3
        Top = 117
        Width = 126
        Height = 384
        TabOrder = 12
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid1CellClick
      end
    end
    object TabSheet2: TTabSheet
      Caption = #1042#1080#1082#1083#1072#1076#1072#1095#1110
      ImageIndex = 1
      ExplicitWidth = 745
      ExplicitHeight = 395
      object Edit2: TEdit
        Left = 3
        Top = 3
        Width = 146
        Height = 21
        TabOrder = 0
      end
      object Button6: TButton
        Left = 3
        Top = 30
        Width = 146
        Height = 25
        Caption = #1055#1086#1096#1091#1082
        TabOrder = 1
        OnClick = Button6Click
      end
      object DBGrid4: TDBGrid
        Left = 0
        Top = 61
        Width = 121
        Height = 434
        TabOrder = 2
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object DBGrid5: TDBGrid
        Left = 142
        Top = 78
        Width = 570
        Height = 434
        TabOrder = 3
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object Button4: TButton
        Left = 166
        Top = -1
        Width = 146
        Height = 25
        Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1074#1080#1082#1083#1072#1076#1072#1095#1110#1074
        TabOrder = 4
        OnClick = Button4Click
      end
    end
    object TabSheet3: TTabSheet
      Caption = #1043#1088#1072#1092#1110#1082' '#1085#1072#1074#1095'. '#1087#1088#1086#1094#1077#1089#1091
      ImageIndex = 2
      ExplicitLeft = -12
      ExplicitTop = 28
      object Label1: TLabel
        Left = 36
        Top = 3
        Width = 129
        Height = 13
        Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1087#1086#1095#1072#1090#1082#1086#1074#1091' '#1076#1072#1090#1091
      end
      object Label2: TLabel
        Left = 36
        Top = 49
        Width = 113
        Height = 13
        Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1082#1110#1085#1094#1077#1074#1091' '#1076#1072#1090#1091
      end
      object Label3: TLabel
        Left = 264
        Top = 3
        Width = 70
        Height = 13
        Caption = #1042#1080#1076' '#1085#1072#1074#1095#1072#1085#1085#1103
      end
      object Label4: TLabel
        Left = 264
        Top = 50
        Width = 55
        Height = 13
        Caption = #1050#1086#1084#1084#1077#1085#1090#1072#1088
      end
      object DateTimePicker1: TDateTimePicker
        Left = 3
        Top = 22
        Width = 186
        Height = 21
        Date = 42525.558961805550000000
        Time = 42525.558961805550000000
        TabOrder = 0
      end
      object Edit3: TEdit
        Left = 195
        Top = 69
        Width = 186
        Height = 21
        TabOrder = 1
        Text = 'Edit3'
      end
      object ComboBox1: TComboBox
        Left = 195
        Top = 22
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
        Left = 400
        Top = 23
        Width = 105
        Height = 67
        Caption = #1044#1086#1076#1072#1090#1080
        TabOrder = 3
        OnClick = Button9Click
      end
      object DBGrid6: TDBGrid
        Left = 3
        Top = 132
        Width = 126
        Height = 309
        TabOrder = 4
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid6CellClick
      end
      object DBGrid7: TDBGrid
        Left = 144
        Top = 152
        Width = 449
        Height = 308
        TabOrder = 5
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object Panel2: TPanel
        Left = 135
        Top = 105
        Width = 554
        Height = 41
        Caption = 'Panel2'
        TabOrder = 6
      end
      object DBNavigator7: TDBNavigator
        Left = 511
        Top = 64
        Width = 240
        Height = 25
        TabOrder = 7
      end
      object DateTimePicker2: TDateTimePicker
        Left = 3
        Top = 68
        Width = 186
        Height = 21
        Date = 42525.559036608790000000
        Time = 42525.559036608790000000
        TabOrder = 8
      end
      object DBLookupComboBox1: TDBLookupComboBox
        Left = 624
        Top = 200
        Width = 145
        Height = 21
        TabOrder = 9
      end
    end
    object TabSheet4: TTabSheet
      Caption = #1056#1086#1079#1082#1083#1072#1076
      ImageIndex = 3
      ExplicitLeft = 0
      ExplicitWidth = 433
      ExplicitHeight = 197
    end
    object TabSheet5: TTabSheet
      Caption = #1047#1072#1084#1077#1085#1099
      ImageIndex = 4
      object PageControl2: TPageControl
        Left = 0
        Top = 0
        Width = 793
        Height = 513
        ActivePage = TabSheet6
        TabOrder = 0
        object TabSheet6: TTabSheet
          Caption = #1057#1091#1073#1086#1090#1080
          ExplicitHeight = 389
          object Label5: TLabel
            Left = 24
            Top = 3
            Width = 138
            Height = 13
            Caption = #1042#1080#1073#1077#1088#1110#1090#1100' '#1089#1091#1073#1086#1090#1091' '#1076#1083#1103' '#1079#1072#1084#1110#1085#1080
          end
          object Label6: TLabel
            Left = 16
            Top = 69
            Width = 87
            Height = 13
            Caption = #1055#1077#1088#1077#1083#1110#1082' '#1091#1089#1110#1093' '#1075#1088#1091#1087
          end
          object Label7: TLabel
            Left = 135
            Top = 69
            Width = 139
            Height = 13
            Caption = #1043#1088#1091#1087#1080' '#1076#1083#1103' '#1076#1086#1076#1072#1074#1072#1085#1085#1103' '#1079#1072#1084#1110#1085
          end
          object Label8: TLabel
            Left = 264
            Top = 0
            Width = 49
            Height = 13
            Caption = #1050#1086#1084#1077#1085#1090#1072#1088
          end
          object DBGrid8: TDBGrid
            Left = 3
            Top = 88
            Width = 126
            Height = 370
            TabOrder = 0
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -11
            TitleFont.Name = 'Tahoma'
            TitleFont.Style = []
            OnCellClick = DBGrid8CellClick
            OnDblClick = DBGrid8DblClick
          end
          object DateTimePicker3: TDateTimePicker
            Left = 3
            Top = 21
            Width = 186
            Height = 21
            Date = 42525.805654305560000000
            Time = 42525.805654305560000000
            TabOrder = 1
            OnCloseUp = DateTimePicker3CloseUp
          end
          object Edit4: TEdit
            Left = 195
            Top = 19
            Width = 186
            Height = 21
            TabOrder = 2
            Text = 'Edit4'
          end
          object Button10: TButton
            Left = 387
            Top = 10
            Width = 135
            Height = 39
            Caption = #1044#1086#1076#1072#1090#1080
            TabOrder = 3
            OnClick = Button10Click
          end
          object DBGrid9: TDBGrid
            Left = 280
            Top = 88
            Width = 436
            Height = 297
            TabOrder = 4
            TitleFont.Charset = DEFAULT_CHARSET
            TitleFont.Color = clWindowText
            TitleFont.Height = -11
            TitleFont.Name = 'Tahoma'
            TitleFont.Style = []
          end
          object ListBox1: TListBox
            Left = 135
            Top = 88
            Width = 139
            Height = 339
            ItemHeight = 13
            TabOrder = 5
            OnDblClick = ListBox1DblClick
          end
          object DBNavigator9: TDBNavigator
            Left = 328
            Top = 57
            Width = 240
            Height = 25
            TabOrder = 6
          end
          object Panel3: TPanel
            Left = 567
            Top = 41
            Width = 185
            Height = 41
            Caption = 'Panel3'
            TabOrder = 7
          end
          object Button11: TButton
            Left = 135
            Top = 433
            Width = 139
            Height = 25
            Caption = #1044#1086#1076#1072#1090#1080' '#1091#1089#1110' '#1075#1088#1091#1087#1080
            TabOrder = 8
            OnClick = Button11Click
          end
        end
        object TabSheet7: TTabSheet
          Caption = #1044#1077#1085#1100' '#1085#1072' '#1076#1077#1085#1100
          ImageIndex = 1
          ExplicitLeft = 0
          ExplicitHeight = 389
        end
      end
    end
  end
  object DBListBox1: TDBListBox
    Left = 304
    Top = 400
    Width = 121
    Height = 97
    ItemHeight = 13
    TabOrder = 1
  end
  object ADOConnection1: TADOConnection
    Left = 616
    Top = 32
  end
  object DataSource1: TDataSource
    Left = 648
    Top = 32
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 680
    Top = 32
  end
  object ADOConnection2: TADOConnection
    Left = 616
    Top = 32
  end
  object DataSource2: TDataSource
    Left = 648
    Top = 32
  end
  object ADOQuery2: TADOQuery
    Parameters = <>
    Left = 680
    Top = 32
  end
  object ADOConnection3: TADOConnection
    Left = 616
    Top = 32
  end
  object DataSource3: TDataSource
    Left = 648
    Top = 32
  end
  object ADOQuery3: TADOQuery
    Parameters = <>
    Left = 680
    Top = 32
  end
  object ADOConnection4: TADOConnection
    Left = 616
    Top = 32
  end
  object DataSource4: TDataSource
    Left = 648
    Top = 32
  end
  object ADOQuery4: TADOQuery
    Parameters = <>
    Left = 680
    Top = 32
  end
  object ADOConnection5: TADOConnection
    Left = 616
    Top = 32
  end
  object DataSource5: TDataSource
    Left = 648
    Top = 32
  end
  object ADOQuery5: TADOQuery
    Parameters = <>
    Left = 680
    Top = 32
  end
  object OpenDialog1: TOpenDialog
    Left = 540
    Top = 32
  end
  object ADOConnection7: TADOConnection
    Left = 756
    Top = 32
  end
  object DataSource7: TDataSource
    Left = 756
    Top = 72
  end
  object ADOQuery7: TADOQuery
    Parameters = <>
    Left = 756
    Top = 120
  end
  object ADOConnection9: TADOConnection
    Left = 760
    Top = 32
  end
  object DataSource9: TDataSource
    Left = 760
    Top = 72
  end
  object ADOQuery9: TADOQuery
    Parameters = <>
    Left = 760
    Top = 120
  end
end
