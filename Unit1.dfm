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
  object PageControl1: TPageControl
    Left = 0
    Top = 0
    Width = 887
    Height = 561
    ActivePage = TabSheet1
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
    end
    object TabSheet4: TTabSheet
      Caption = #1056#1086#1079#1082#1083#1072#1076
      ImageIndex = 3
      ExplicitLeft = 0
      ExplicitWidth = 433
      ExplicitHeight = 197
    end
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
    Left = 748
    Top = 40
  end
  object DataSource7: TDataSource
    Left = 756
    Top = 88
  end
  object ADOQuery7: TADOQuery
    Parameters = <>
    Left = 756
    Top = 136
  end
end
