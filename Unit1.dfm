object Form1: TForm1
  Left = 0
  Top = 0
  Caption = 'Form1'
  ClientHeight = 589
  ClientWidth = 1213
  Color = clBtnFace
  Font.Charset = DEFAULT_CHARSET
  Font.Color = clWindowText
  Font.Height = -11
  Font.Name = 'Tahoma'
  Font.Style = []
  Menu = MainMenu1
  OldCreateOrder = False
  OnCreate = FormCreate
  OnDestroy = FormDestroy
  PixelsPerInch = 96
  TextHeight = 13
  object DBNavigator1: TDBNavigator
    Left = 1193
    Top = 0
    Width = 20
    Height = 570
    Align = alRight
    Kind = dbnVertical
    TabOrder = 0
  end
  object DBGrid2: TDBGrid
    Left = 453
    Top = 0
    Width = 135
    Height = 570
    Align = alRight
    FixedColor = clWindow
    Options = [dgTitles, dgIndicator, dgColumnResize, dgColLines, dgRowLines, dgTabs, dgRowSelect, dgConfirmDelete, dgCancelOnExit, dgTitleClick, dgTitleHotTrack]
    ReadOnly = True
    TabOrder = 1
    TitleFont.Charset = DEFAULT_CHARSET
    TitleFont.Color = clWindowText
    TitleFont.Height = -11
    TitleFont.Name = 'Tahoma'
    TitleFont.Style = []
    OnCellClick = DBGrid2CellClick
  end
  object DBNavigator2: TDBNavigator
    Left = 0
    Top = 0
    Width = 160
    Height = 25
    VisibleButtons = [nbFirst, nbPrior, nbNext, nbLast, nbDelete, nbPost, nbCancel, nbRefresh]
    TabOrder = 2
    Visible = False
  end
  object StatusBar1: TStatusBar
    Left = 0
    Top = 570
    Width = 1213
    Height = 19
    Panels = <
      item
        Text = #1044#1072#1090#1072':'
        Width = 120
      end
      item
        Text = #1042#1088#1077#1084#1103':'
        Width = 120
      end>
    OnHint = StatusBar1Hint
  end
  object Panel2: TPanel
    Left = 588
    Top = 0
    Width = 605
    Height = 570
    Align = alRight
    TabOrder = 4
    object Splitter1: TSplitter
      Left = 1
      Top = 384
      Width = 603
      Height = 3
      Cursor = crVSplit
      Align = alTop
      ExplicitLeft = -4
      ExplicitTop = 263
    end
    object DBGrid1: TDBGrid
      Left = 1
      Top = 42
      Width = 603
      Height = 342
      Align = alTop
      DragMode = dmAutomatic
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -11
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
    end
    object Panel1: TPanel
      Left = 1
      Top = 501
      Width = 603
      Height = 68
      Align = alBottom
      BevelInner = bvLowered
      TabOrder = 1
      object Label2: TLabel
        Left = 22
        Top = 8
        Width = 34
        Height = 13
        Caption = #1055#1083#1072#1085' '#1041
      end
      object Label3: TLabel
        Left = 62
        Top = 8
        Width = 35
        Height = 13
        Caption = #1055#1083#1072#1085' '#1050
      end
      object Label4: TLabel
        Left = 109
        Top = 8
        Width = 49
        Height = 13
        Caption = #1050#1086#1084#1077#1085#1090#1072#1088
      end
      object Label5: TLabel
        Left = 218
        Top = 8
        Width = 30
        Height = 13
        Caption = #1043#1088#1091#1087#1072
      end
      object Label6: TLabel
        Left = 335
        Top = 8
        Width = 57
        Height = 13
        Caption = #1044#1080#1089#1094#1080#1087#1083#1110#1085#1072
      end
      object Label7: TLabel
        Left = 481
        Top = 8
        Width = 49
        Height = 13
        Caption = #1042#1080#1082#1083#1072#1076#1072#1095
      end
      object DBEdit1: TDBEdit
        Left = 15
        Top = 27
        Width = 41
        Height = 21
        DataSource = DataSource1
        TabOrder = 0
      end
      object DBEdit2: TDBEdit
        Left = 62
        Top = 27
        Width = 41
        Height = 21
        TabOrder = 1
      end
      object DBEdit3: TDBEdit
        Left = 109
        Top = 27
        Width = 67
        Height = 21
        TabOrder = 2
      end
      object DBLookupComboBox1: TDBLookupComboBox
        Left = 182
        Top = 27
        Width = 105
        Height = 21
        TabOrder = 3
      end
      object DBLookupComboBox2: TDBLookupComboBox
        Left = 293
        Top = 27
        Width = 143
        Height = 21
        TabOrder = 4
      end
      object DBLookupComboBox3: TDBLookupComboBox
        Left = 442
        Top = 27
        Width = 145
        Height = 21
        TabOrder = 5
      end
    end
    object Panel3: TPanel
      Left = 1
      Top = 1
      Width = 603
      Height = 41
      Align = alTop
      Caption = 'Panel3'
      Font.Charset = DEFAULT_CHARSET
      Font.Color = clWindowText
      Font.Height = -27
      Font.Name = 'Tahoma'
      Font.Style = []
      ParentFont = False
      TabOrder = 2
    end
    object DBGrid5: TDBGrid
      Left = 200
      Top = 387
      Width = 404
      Height = 114
      Align = alRight
      TabOrder = 3
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
    end
  end
  object GroupBox1: TGroupBox
    Left = 260
    Top = 0
    Width = 193
    Height = 570
    Align = alRight
    Caption = #1044#1086#1076#1072#1074#1072#1085#1085#1103
    TabOrder = 5
    object Label1: TLabel
      Left = 24
      Top = 31
      Width = 146
      Height = 13
      Caption = #1042#1074#1077#1076#1110#1090#1100' '#1087#1088#1110#1079#1074#1080#1097#1077' '#1074#1080#1082#1083#1072#1076#1072#1095#1072
    end
    object Label9: TLabel
      Left = 24
      Top = 124
      Width = 130
      Height = 13
      Caption = #1042#1074#1077#1076#1110#1090#1100' '#1085#1072#1079#1074#1091' '#1076#1080#1089#1094#1080#1087#1083#1110#1085#1080
    end
    object Label10: TLabel
      Left = 26
      Top = 219
      Width = 104
      Height = 13
      Caption = #1042#1074#1077#1076#1110#1090#1100' '#1085#1086#1084#1077#1088' '#1075#1088#1091#1087#1080
    end
    object Label11: TLabel
      Left = 26
      Top = 265
      Width = 85
      Height = 13
      Caption = #1042#1074#1077#1076#1110#1090#1100' ID '#1075#1088#1091#1087#1080
    end
    object Button2: TButton
      Left = 24
      Top = 311
      Width = 145
      Height = 25
      Caption = #1044#1086#1076#1072#1090#1080' '#1075#1088#1091#1087#1091
      TabOrder = 0
      OnClick = Button2Click
    end
    object Button5: TButton
      Left = 24
      Top = 170
      Width = 145
      Height = 25
      Caption = #1044#1086#1076#1072#1090#1080' '#1076#1080#1089#1094#1080#1087#1083#1110#1085#1091
      TabOrder = 1
      OnClick = Button5Click
    end
    object Button6: TButton
      Left = 24
      Top = 77
      Width = 145
      Height = 25
      Caption = #1044#1086#1076#1072#1090#1080' '#1074#1080#1082#1083#1072#1076#1072#1095#1072
      TabOrder = 2
      OnClick = Button6Click
    end
    object Edit1: TEdit
      Left = 24
      Top = 284
      Width = 145
      Height = 21
      TabOrder = 3
    end
    object Edit2: TEdit
      Left = 24
      Top = 238
      Width = 145
      Height = 21
      TabOrder = 4
    end
    object Edit3: TEdit
      Left = 24
      Top = 143
      Width = 145
      Height = 21
      TabOrder = 5
    end
    object Edit4: TEdit
      Left = 24
      Top = 50
      Width = 145
      Height = 21
      TabOrder = 6
    end
    object Button1: TButton
      Left = 24
      Top = 468
      Width = 145
      Height = 25
      Caption = #1047#1072#1074#1072#1085#1090#1072#1078#1080#1090#1080' '#1079' Excel'
      TabOrder = 7
      OnClick = Button1Click
    end
    object Button3: TButton
      Left = 24
      Top = 497
      Width = 145
      Height = 25
      Caption = #1044#1086#1076#1072#1090#1080
      TabOrder = 8
      OnClick = Button3Click
    end
    object Button7: TButton
      Left = 24
      Top = 381
      Width = 145
      Height = 68
      Caption = #1043#1077#1085#1077#1088#1091#1074#1072#1090#1080
      TabOrder = 9
      OnClick = Button7Click
    end
    object Button8: TButton
      Left = 24
      Top = 528
      Width = 145
      Height = 25
      Caption = #1044#1086#1076#1072#1090#1080' '#1088#1086#1079#1082#1083#1072#1076
      TabOrder = 10
      OnClick = Button8Click
    end
  end
  object Button4: TButton
    Left = 0
    Top = 48
    Width = 75
    Height = 25
    Caption = 'Button4'
    TabOrder = 6
    OnClick = Button4Click
  end
  object DateTimePicker1: TDateTimePicker
    Left = 0
    Top = 79
    Width = 186
    Height = 21
    Date = 42435.920827997690000000
    Time = 42435.920827997690000000
    TabOrder = 7
  end
  object Edit5: TEdit
    Left = 1
    Top = 31
    Width = 121
    Height = 21
    TabOrder = 8
    Text = 'Edit5'
  end
  object Panel4: TPanel
    Left = 96
    Top = 0
    Width = 164
    Height = 570
    Align = alRight
    TabOrder = 9
    object DBGrid3: TDBGrid
      Left = 16
      Top = 204
      Width = 142
      Height = 180
      TabOrder = 0
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
    end
    object DBGrid4: TDBGrid
      Left = 16
      Top = 390
      Width = 144
      Height = 180
      TabOrder = 1
      TitleFont.Charset = DEFAULT_CHARSET
      TitleFont.Color = clWindowText
      TitleFont.Height = -11
      TitleFont.Name = 'Tahoma'
      TitleFont.Style = []
    end
    object DBNavigator3: TDBNavigator
      Left = -1
      Top = 204
      Width = 20
      Height = 180
      Kind = dbnVertical
      TabOrder = 2
    end
    object DBNavigator4: TDBNavigator
      Left = 1
      Top = 390
      Width = 20
      Height = 180
      Kind = dbnVertical
      TabOrder = 3
    end
  end
  object ADOConnection1: TADOConnection
    Left = 576
    Top = 72
  end
  object ADOConnection2: TADOConnection
    Left = 576
    Top = 128
  end
  object DataSource2: TDataSource
    Left = 632
    Top = 128
  end
  object DataSource1: TDataSource
    Left = 632
    Top = 72
  end
  object ADOQuery1: TADOQuery
    Parameters = <>
    Left = 680
    Top = 72
  end
  object ADOQuery2: TADOQuery
    Parameters = <>
    Left = 680
    Top = 128
  end
  object ADOConnection3: TADOConnection
    Left = 568
    Top = 184
  end
  object DataSource3: TDataSource
    Left = 632
    Top = 184
  end
  object ADOQuery3: TADOQuery
    Parameters = <>
    Left = 688
    Top = 192
  end
  object ADOConnection4: TADOConnection
    Left = 560
    Top = 240
  end
  object DataSource4: TDataSource
    Left = 632
    Top = 240
  end
  object ADOQuery4: TADOQuery
    Parameters = <>
    Left = 688
    Top = 240
  end
  object MainMenu1: TMainMenu
    Left = 832
    Top = 128
    object N1: TMenuItem
      Caption = #1043#1088#1091#1087#1080
    end
    object N2: TMenuItem
      Caption = #1044#1080#1089#1094#1080#1087#1083#1110#1085#1080
    end
    object N3: TMenuItem
      Caption = #1042#1080#1082#1083#1072#1076#1072#1095#1110
    end
  end
  object Timer1: TTimer
    OnTimer = Timer1Timer
    Left = 832
    Top = 80
  end
  object OpenDialog1: TOpenDialog
    Left = 832
    Top = 176
  end
  object ADOConnection5: TADOConnection
    Left = 408
    Top = 416
  end
  object DataSource5: TDataSource
    Left = 464
    Top = 416
  end
  object ADOQuery5: TADOQuery
    Parameters = <>
    Left = 536
    Top = 424
  end
end
