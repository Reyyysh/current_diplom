object Form1: TForm1
  Left = 0
  Top = 0
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Form1'
  ClientHeight = 561
  ClientWidth = 713
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  OldCreateOrder = False
  OnCreate = Form1Create
  OnDestroy = Form1Destroy
  PixelsPerInch = 96
  TextHeight = 13
  object TabControl1: TTabControl
    Left = 0
    Top = 0
    Width = 713
    Height = 25
    TabOrder = 0
    Tabs.Strings = (
      #1043#1088#1091#1087#1080
      #1042#1080#1082#1083#1072#1076#1072#1095#1110
      #1043#1088#1072#1092#1110#1082' '#1085#1072#1074#1095'.'#1087#1088#1086#1094#1077#1089#1091
      #1056#1086#1079#1082#1083#1072#1076)
    TabIndex = 0
    OnChange = TabControl1Change
  end
  object DBGrid1: TDBGrid
    Left = 8
    Top = 139
    Width = 121
    Height = 416
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnCellClick = DBGrid1CellClick
  end
  object DBGrid2: TDBGrid
    Left = 135
    Top = 147
    Width = 570
    Height = 294
    TabOrder = 2
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object DBGrid3: TDBGrid
    Left = 135
    Top = 435
    Width = 570
    Height = 120
    TabOrder = 3
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object DBGrid4: TDBGrid
    Left = 16
    Top = 201
    Width = 121
    Height = 434
    TabOrder = 4
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object DBGrid5: TDBGrid
    Left = 270
    Top = 273
    Width = 570
    Height = 434
    TabOrder = 5
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
  end
  object Button1: TButton
    Left = 8
    Top = 31
    Width = 145
    Height = 25
    Caption = #1042#1110#1076#1082#1088#1080#1090#1080' Exel '#1092#1072#1081#1083
    TabOrder = 6
    OnClick = Button1Click
  end
  object Edit1: TEdit
    Left = 8
    Top = 65
    Width = 121
    Height = 21
    TabOrder = 7
  end
  object Button2: TButton
    Left = 159
    Top = 31
    Width = 146
    Height = 25
    Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1088#1086#1079#1082#1083#1072#1076
    TabOrder = 8
    OnClick = Button2Click
  end
  object Button3: TButton
    Left = 311
    Top = 31
    Width = 146
    Height = 25
    Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1095'/'#1079
    TabOrder = 9
    OnClick = Button3Click
  end
  object Button4: TButton
    Left = 159
    Top = 31
    Width = 146
    Height = 25
    Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1074#1080#1082#1083#1072#1076#1072#1095#1110#1074
    TabOrder = 10
    OnClick = Button4Click
  end
  object Button5: TButton
    Left = 8
    Top = 90
    Width = 121
    Height = 25
    Caption = #1044#1086#1076#1072#1090#1080' '#1075#1088#1091#1087#1091
    TabOrder = 11
    OnClick = Button5Click
  end
  object DBNavigator1: TDBNavigator
    Left = 8
    Top = 121
    Width = 120
    Height = 19
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
    Enabled = False
    TabOrder = 12
  end
  object DBNavigator2: TDBNavigator
    Left = 135
    Top = 121
    Width = 570
    Height = 20
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
    TabOrder = 13
  end
  object Button6: TButton
    Left = 32
    Top = 132
    Width = 121
    Height = 25
    Caption = #1055#1086#1096#1091#1082
    TabOrder = 14
    OnClick = Button6Click
  end
  object Edit2: TEdit
    Left = 64
    Top = 92
    Width = 121
    Height = 21
    TabOrder = 15
  end
  object Panel1: TPanel
    Left = 135
    Top = 62
    Width = 570
    Height = 53
    Caption = #1043#1088#1091#1087#1072' '#1085#1077' '#1074#1080#1073#1088#1072#1085#1072
    Font.Charset = DEFAULT_CHARSET
    Font.Color = clWindowText
    Font.Height = -24
    Font.Name = 'Tahoma'
    Font.Style = []
    ParentFont = False
    TabOrder = 16
  end
  object Button7: TButton
    Left = 8
    Top = 530
    Width = 121
    Height = 25
    Caption = 'delete_test'
    TabOrder = 17
    OnClick = Button7Click
  end
  object Button8: TButton
    Left = 463
    Top = 31
    Width = 114
    Height = 25
    Caption = #1053#1086#1074#1086#1077' '#1076#1086#1073#1072#1074#1083#1077#1085#1080#1077
    TabOrder = 18
    OnClick = Button8Click
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
    Left = 584
    Top = 32
  end
end
