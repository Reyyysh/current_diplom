object Form1: TForm1
  Left = 439
  Top = 219
  BorderIcons = [biSystemMenu, biMinimize]
  BorderStyle = bsSingle
  Caption = 'Form1'
  ClientHeight = 543
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
      ExplicitLeft = 8
      ExplicitTop = 72
      ExplicitWidth = 745
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
        Height = 26
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
        Top = 92
        Width = 570
        Height = 20
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        TabOrder = 5
      end
      object DBGrid2: TDBGrid
        Left = 150
        Top = 117
        Width = 570
        Height = 294
        TabOrder = 6
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
      end
      object DBGrid1: TDBGrid
        Left = 0
        Top = 117
        Width = 145
        Height = 384
        TabOrder = 7
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
        OnCellClick = DBGrid1CellClick
      end
      object DBNavigator1: TDBNavigator
        Left = 0
        Top = 92
        Width = 144
        Height = 19
        VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbInsert, nbDelete]
        Enabled = False
        TabOrder = 8
      end
      object Button5: TButton
        Left = 0
        Top = 61
        Width = 145
        Height = 25
        Caption = #1044#1086#1076#1072#1090#1080' '#1075#1088#1091#1087#1091
        TabOrder = 9
        OnClick = Button5Click
      end
      object Edit1: TEdit
        Left = 0
        Top = 34
        Width = 145
        Height = 21
        TabOrder = 10
      end
      object Button7: TButton
        Left = 0
        Top = 507
        Width = 145
        Height = 25
        Caption = 'delete_test'
        TabOrder = 11
        OnClick = Button7Click
      end
      object DBGrid3: TDBGrid
        Left = 151
        Top = 417
        Width = 570
        Height = 120
        TabOrder = 12
        TitleFont.Charset = DEFAULT_CHARSET
        TitleFont.Color = clWindowText
        TitleFont.Height = -11
        TitleFont.Name = 'Tahoma'
        TitleFont.Style = []
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
      ExplicitWidth = 433
      ExplicitHeight = 197
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
    Left = 528
    Top = 32
  end
end
