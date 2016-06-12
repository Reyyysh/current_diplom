unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.ExtCtrls, Vcl.DBCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls, Data.Win.ADODB, ADOX_TLB, Vcl.Mask,
  Vcl.ToolWin, Vcl.ComCtrls, Vcl.Menus, ComObj, Vcl.Buttons,IniFiles,
  Vcl.Samples.Spin, Vcl.ColorGrd;

type
  TForm1 = class(TForm)

  //IniFile: TIniFile;
    ADOConnection1: TADOConnection;
    DataSource1: TDataSource;
    ADOQuery1: TADOQuery;
   // ADOConnection2: TADOConnection;
  //  DataSource2: TDataSource;
  //  ADOQuery2: TADOQuery;
   // ADOConnection3: TADOConnection;
   // DataSource3: TDataSource;
  //  ADOQuery3: TADOQuery;
    ADOConnection4: TADOConnection;
    DataSource4: TDataSource;
    ADOQuery4: TADOQuery;
    ADOConnection5: TADOConnection;
    DataSource5: TDataSource;
    ADOQuery5: TADOQuery;
    OpenDialog1: TOpenDialog;
    ADOConnection7: TADOConnection;
    DataSource7: TDataSource;
    ADOQuery7: TADOQuery;
    ADOConnection9: TADOConnection;
    DataSource9: TDataSource;
    ADOQuery9: TADOQuery;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    Label12: TLabel;
    Button1: TButton;
    Panel1: TPanel;
    DBNavigator2: TDBNavigator;
    DBGrid1: TDBGrid;
    Edit6: TEdit;
    TabSheet2: TTabSheet;
    Label11: TLabel;
    Edit2: TEdit;
    DBGrid4: TDBGrid;
    Label13: TLabel;
    Label14: TLabel;
    Label15: TLabel;
    Label16: TLabel;
    Label17: TLabel;
    Label18: TLabel;
    Label19: TLabel;
    Label20: TLabel;
    DateTimePicker5: TDateTimePicker;
    DBLookupComboBox1: TDBLookupComboBox;
    ComboBox2: TComboBox;
    DBLookupComboBox2: TDBLookupComboBox;
    DBLookupComboBox3: TDBLookupComboBox;
    DBLookupComboBox5: TDBLookupComboBox;
    DBLookupComboBox4: TDBLookupComboBox;
    DBGrid5: TDBGrid;
    DBNavigator5: TDBNavigator;
    Button16: TButton;
    ComboBox3: TComboBox;
    TabSheet3: TTabSheet;
    Label1: TLabel;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label10: TLabel;
    DateTimePicker1: TDateTimePicker;
    Edit3: TEdit;
    ComboBox1: TComboBox;
    Button9: TButton;
    DBGrid6: TDBGrid;
    DBGrid7: TDBGrid;
    Panel2: TPanel;
    DBNavigator7: TDBNavigator;
    DateTimePicker2: TDateTimePicker;
    ListBox2: TListBox;
    Edit5: TEdit;
    Button6: TButton;
    TabSheet4: TTabSheet;
    MonthCalendar1: TMonthCalendar;
    TabSheet5: TTabSheet;
    Label7: TLabel;
    Label6: TLabel;
    Label9: TLabel;
    Label8: TLabel;
    Label5: TLabel;
    ListBox1: TListBox;
    DateTimePicker4: TDateTimePicker;
    DateTimePicker3: TDateTimePicker;
    RadioButton1: TRadioButton;
    RadioButton2: TRadioButton;
    DBGrid9: TDBGrid;
    Button10: TButton;
    Button11: TButton;
    Panel3: TPanel;
    Button12: TButton;
    DBNavigator9: TDBNavigator;
    Edit4: TEdit;
    DBGrid8: TDBGrid;
    StatusBar1: TStatusBar;
    Timer1: TTimer;
    Button17: TButton;
    Button18: TButton;
    Button19: TButton;
    Button20: TButton;
    Button21: TButton;
    Button22: TButton;
    TabSheet9: TTabSheet;
    Label21: TLabel;
    Button23: TButton;
    Button24: TButton;
    DBGrid10: TDBGrid;
    Label22: TLabel;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    Edit10: TEdit;
    Button25: TButton;
    ADOConnection10: TADOConnection;
    ADOQuery10: TADOQuery;
    DataSource10: TDataSource;
    ComboBox5: TComboBox;
    GroupBox1: TGroupBox;
    Label23: TLabel;
    GroupBox2: TGroupBox;
    Edit11: TEdit;
    Timer2: TTimer;
    GroupBox3: TGroupBox;
    Edit7: TEdit;
    Button26: TButton;
    DBGrid11: TDBGrid;
    Button28: TButton;
    DataSource11: TDataSource;
    ADOQuery11: TADOQuery;
    ADOConnection11: TADOConnection;
    Label24: TLabel;
    Button29: TButton;
    CheckBox1: TCheckBox;
    Label25: TLabel;
    SpinEdit1: TSpinEdit;
    Button27: TButton;
    Label26: TLabel;
    Edit12: TEdit;
    Label27: TLabel;
    Edit13: TEdit;
    ADOConnection13: TADOConnection;
    ADOQuery13: TADOQuery;
    DataSource13: TDataSource;
    DBGrid13: TDBGrid;
    Button30: TButton;
    Button32: TButton;
    RadioGroup1: TRadioGroup;
    Button2: TButton;
    RadioGroup2: TRadioGroup;
    Button3: TButton;
    procedure Form1Destroy(Sender: TObject);
    procedure Form1Create(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure Button6Click(Sender: TObject);
    procedure DBGrid6CellClick(Column: TColumn);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);
    procedure DBGrid8DblClick(Sender: TObject);
    procedure ListBox1DblClick(Sender: TObject);
    procedure DBGrid8CellClick(Column: TColumn);
    procedure Button11Click(Sender: TObject);
    procedure DBGrid6DblClick(Sender: TObject);
    procedure ListBox2DblClick(Sender: TObject);
    procedure Edit2Change(Sender: TObject);
    procedure Edit5Change(Sender: TObject);
    procedure Button12Click(Sender: TObject);
    procedure Edit6Change(Sender: TObject);
    procedure Button13Click(Sender: TObject);
    procedure DBGrid9DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Button14Click(Sender: TObject);
    procedure DBGrid7DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure Button16Click(Sender: TObject);
    procedure Button17Click(Sender: TObject);
    procedure StatusBar1Hint(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure N1Click(Sender: TObject);
    procedure Button23Click(Sender: TObject);

    procedure Button25Click(Sender: TObject);
    procedure DBGrid10CellClick(Column: TColumn);
    procedure DBGrid10DblClick(Sender: TObject);
    procedure Button28Click(Sender: TObject);
    procedure Button29Click(Sender: TObject);
    procedure Button26Click(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure Timer2Timer(Sender: TObject);
    procedure Button27Click(Sender: TObject);
    procedure DBGrid2DrawDataCell(Sender: TObject; const [Ref] Rect: TRect;
      Field: TField; State: TGridDrawState);
    procedure Button30Click(Sender: TObject);
    procedure DBGrid13DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid1DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid4DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid10DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid11DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid5DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid6DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBGrid8DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
      DataCol: Integer; Column: TColumn; State: TGridDrawState);
    procedure DBLookupComboBox1CloseUp(Sender: TObject);
    procedure RadioGroup1Click(Sender: TObject);
    procedure DateTimePicker5CloseUp(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);


  private
    { Private declarations }
  public
    table:string;
    fexe: integer; // Проверка первого запука программы
    procedure DBGrid1Width();
   // procedure DBGrid2Width();
  //  procedure DBGrid3Width();
    procedure DBGrid4Width();
    procedure DBGrid5Width();
    procedure DBGrid6Width();
    procedure DBGrid7Width();
    procedure DBGrid8Width();
    procedure DBGrid9Width();
    procedure DBGrid10Width();
    procedure DBGrid11Width();
    procedure DBGrid13Width();
    procedure LoadForm();
    procedure LoadProSys();
    procedure BackupSave();
    procedure ViewGroup();
  end;

var
  Form1: TForm1;
  IniFile: TIniFile; // переменная файла конфугурации .ini
  DB:Catalog;
  Tables:Table;
  Columns:Column;
  excel: variant;
implementation


procedure TForm1.ViewGroup();
begin

end;


procedure TForm1.BackupSave();
var
  strNameDate, strNameTime, tI, strTable:string;
  i,y:integer;
begin
     CreateDir('Backup'); // Создание папки

     strNameTime:=FormatDateTime('tt',Now);
     for i := 1 to Length(strNameTime) do
  begin
    if strNameTime[i] = ':' then strNameTime[i] := '.';
  end;
      strTable:=edit7.Text+'.mdb';
   //strName:='2016-2017('+DateToStr(now)+' '+FormatDateTime('c',Now)+').mdb';
    strNameDate:=FormatDateTime('ddddd',Now);
  if CopyFile(PWideChar(strTable), PWideChar('Backup\'+edit7.Text+'('+strNameDate+'-'+strNameTime+').mdb'),true) then
  begin
     tI:=edit7.Text+'('+strNameDate+'-'+strNameTime+').mdb';


    ADOQuery11.SQL.Clear;
    ADOQuery11.SQL.Add('SELECT * FROM Backup');
    ADOQuery11.Active:=true;

         for y:=0 to Length(tI) do  // Удаление " " из диапазона
  begin
    Delete(tI,Pos(' ',tI),1);
  end;


    ADOQuery11.Insert;
    ADOQuery11.FieldByName('ID').AsString :=tI;
    ADOQuery11.FieldByName('Дата').AsString :=tI;
    ADOQuery11.Post;

    DBGrid11Width();

      ShowMessage('База даних успішно збережена!')
  end
  else ShowMessage('Помилка у зберіганні бази даних!');

end;



 procedure TForm1.LoadProSys();
 begin

   if FileExists('pro-system.mdb') then
  begin
    ADOConnection10.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pro-system.mdb';
    ADOConnection10.LoginPrompt:=false;
    ADOConnection10.Connected:=true;
    ADOQuery10.Connection:=ADOConnection10;
    ADOQuery10.SQL.Clear;
    ADOQuery10.SQL.Add('SELECT * FROM ListAcademicYear');
    ADOQuery10.Active:=true;
    DataSource10.DataSet:=ADOQuery10;
    DBGrid10.DataSource:=DataSource10;
   // DBNavigator10.DataSource:=DataSource10;
    DBGrid10.ReadOnly:=true;

    DBGrid10Width();



    ADOConnection11.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pro-system.mdb';
    ADOConnection11.LoginPrompt:=false;
    ADOConnection11.Connected:=true;
    ADOQuery11.Connection:=ADOConnection11;
    ADOQuery11.SQL.Clear;
    ADOQuery11.SQL.Add('SELECT * FROM Backup');
    ADOQuery11.Active:=true;
    DataSource11.DataSet:=ADOQuery11;
    DBGrid11.DataSource:=DataSource11;
   // DBNavigator10.DataSource:=DataSource10;
    DBGrid11.ReadOnly:=true;

     DBGrid11Width();




  end
  else
  begin
      DB:=CoCatalog.Create;
    DB.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pro-system.mdb');
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pro-system.mdb');


    //Создание таблицы для групп
    ADOConnection10.Connected:=false;
    Tables:=CoTable.Create;
    Tables.Name:='ListAcademicYear';
    Tables.ParentCatalog:=DB;
    DB.Tables.Append(Tables);
    Columns:=CoColumn.Create;
    with Columns do
    begin
      ParentCatalog:=DB;
      Name:='id';
      type_:=adInteger;
      Properties['Autoincrement'].Value:=True;
      Properties['Description'].Value:='Ключевое поле';
    end;
    Tables.Columns.Append('IDНавч_рік',adVarWChar,255);
    Tables.Columns.Append('Навч_рік',adVarWChar,255);

     ADOConnection10.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pro-system.mdb';
    ADOConnection10.LoginPrompt:=false;
    ADOConnection10.Connected:=true;
    ADOQuery10.Connection:=ADOConnection10;
    ADOQuery10.SQL.Clear;
    ADOQuery10.SQL.Add('SELECT * FROM ListAcademicYear');
    ADOQuery10.Active:=true;
    DataSource10.DataSet:=ADOQuery10;
    DBGrid10.DataSource:=DataSource10;
   // DBNavigator10.DataSource:=DataSource10;
    DBGrid10.ReadOnly:=true;

    DBGrid10Width();


     //Создание таблицы для групп
    ADOConnection11.Connected:=false;
    Tables:=CoTable.Create;
    Tables.Name:='Backup';
    Tables.ParentCatalog:=DB;
    DB.Tables.Append(Tables);
    Columns:=CoColumn.Create;
    with Columns do
    begin
      ParentCatalog:=DB;
      Name:='id';
      type_:=adInteger;
      Properties['Autoincrement'].Value:=True;
      Properties['Description'].Value:='Ключевое поле';
    end;
    Tables.Columns.Append('ID',adVarWChar,255);
    Tables.Columns.Append('Дата',adVarWChar,255);

    ADOConnection11.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pro-system.mdb';
    ADOConnection11.LoginPrompt:=false;
    ADOConnection11.Connected:=true;
    ADOQuery11.Connection:=ADOConnection11;
    ADOQuery11.SQL.Clear;
    ADOQuery11.SQL.Add('SELECT * FROM Backup');
    ADOQuery11.Active:=true;
    DataSource11.DataSet:=ADOQuery11;
    DBGrid11.DataSource:=DataSource11;
   // DBNavigator10.DataSource:=DataSource10;
    DBGrid11.ReadOnly:=true;

    DBGrid11Width();

    {
     //Создание таблицы для групп
    ADOConnection13.Connected:=false;
    Tables:=CoTable.Create;
    Tables.Name:='AllRecords';
    Tables.ParentCatalog:=DB;
    DB.Tables.Append(Tables);
    Columns:=CoColumn.Create;
    with Columns do
    begin
      ParentCatalog:=DB;
      Name:='id';
      type_:=adInteger;
      Properties['Autoincrement'].Value:=True;
      Properties['Description'].Value:='Ключевое поле';
    end;

    Tables.Columns.Append('Бютжет',adVarWChar,0);
    Tables.Columns.Append('Контракт',adVarWChar,0);
    Tables.Columns.Append('Комментар',adVarWChar,0);
    Tables.Columns.Append('Група',adVarWChar,255);
    Tables.Columns.Append('Дисципліна',adVarWChar,255);
    Tables.Columns.Append('Викладач',adVarWChar,255);
    Tables.Columns.Append('Color',adVarWChar,255);

    ADOConnection13.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=pro-system.mdb';
    ADOConnection13.LoginPrompt:=false;
    ADOConnection13.Connected:=true;
    ADOQuery13.Connection:=ADOConnection13;
    ADOQuery13.SQL.Clear;
    ADOQuery13.SQL.Add('SELECT * FROM AllRecords');
    ADOQuery13.Active:=true;
    DataSource13.DataSet:=ADOQuery13;
    DBGrid13.DataSource:=DataSource13;
   // DBNavigator10.DataSource:=DataSource10;
    DBGrid13.ReadOnly:=true;

   DBGrid13Width();
    }
  end;
 end;

{$R *.dfm}
procedure TForm1.LoadForm();
var
  i:integer;
  rsp:string;
begin
   ADOConnection1.Connected:=false;
   //ADOConnection2.Connected:=false;
  // ADOConnection3.Connected:=false;
   ADOConnection4.Connected:=false;
   ADOConnection5.Connected:=false;
   ADOConnection7.Connected:=false;
   ADOConnection9.Connected:=false;
   ADOConnection10.Connected:=false;
   ADOConnection11.Connected:=false;
  // ADOConnection13.Connected:=false;


   LoadProSys();



   //ADOQuery10.FindFirst;
   if ADOQuery10.RecordCount<>0 then
  table:=DBGrid10.Fields[0].Value;





  if FileExists(table+'.mdb') then
  begin

    DB:=CoCatalog.Create;
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb');



     ADOConnection13.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection13.LoginPrompt:=false;
    ADOConnection13.Connected:=true;
    ADOQuery13.Connection:=ADOConnection13;
    ADOQuery13.SQL.Clear;
    ADOQuery13.SQL.Add('SELECT * FROM AllRecords');
    ADOQuery13.Active:=true;
    DataSource13.DataSet:=ADOQuery13;
    DBGrid13.DataSource:=DataSource13;
   // DBNavigator10.DataSource:=DataSource10;

   DBGrid13Width();



    //Список групп
    ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection1.LoginPrompt:=false;
    ADOConnection1.Connected:=true;
    ADOQuery1.Connection:=ADOConnection1;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT * FROM grp');
    ADOQuery1.Active:=true;
    DataSource1.DataSet:=ADOQuery1;
    DBGrid1.DataSource:=DataSource1;
    DBGrid1.ReadOnly:=true;
  //  DBNavigator1.DataSource:=DataSource1;




    DBGrid6.DataSource:=DataSource1;
    DBGrid6.ReadOnly:=true;

    DBGrid8.DataSource:=DataSource1;
    DBGrid8.ReadOnly:=true;

    DBGrid1Width();
    DBGrid6Width();
    DBGrid8Width();







  if ADOQuery1.RecordCount<>0 then
  begin
   //Расписание группы
    Panel1.Caption:='Група '+DBGrid1.Fields[1].Value;

  {  ADOConnection2.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection2.LoginPrompt:=false;
    ADOConnection2.Connected:=true;
    ADOQuery2.Connection:=ADOConnection2;
    ADOQuery2.SQL.Clear;
    ADOQuery2.SQL.Add('SELECT * FROM '+DBGrid1.Fields[0].Value);
    ADOQuery2.Active:=true;
    DataSource2.DataSet:=ADOQuery2;

  //  DBGrid2.DataSource:=DataSource2;
  //  DBNavigator2.DataSource:=DataSource2;

    DBGrid2Width();


  {
    //Расписание числитель/знаменатель
    ADOConnection3.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection3.LoginPrompt:=false;
    ADOConnection3.Connected:=true;
    ADOQuery3.Connection:=ADOConnection3;
    ADOQuery3.SQL.Clear;
    rsp:=DBGrid1.Fields[0].Value+'rsp';
    ADOQuery3.SQL.Add('SELECT * FROM '+rsp);
    ADOQuery3.Active:=true;
   // DataSource3.DataSet:=ADOQuery3;
  //  DBGrid3.DataSource:=DataSource3;

    DBGrid3Width();

   }
    ADOConnection7.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection7.LoginPrompt:=false;
    ADOConnection7.Connected:=true;
    ADOQuery7.Connection:=ADOConnection7;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('SELECT * FROM '+DBGrid6.Fields[0].Value+'graf');
    ADOQuery7.Active:=true;
    DataSource7.DataSet:=ADOQuery7;
    DBGrid7.DataSource:=DataSource7;
    DBNavigator7.DataSource:=DataSource7;

    DBGrid7Width();

/////////////////////////////


    ADOConnection9.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection9.LoginPrompt:=false;
    ADOConnection9.Connected:=true;
    ADOQuery9.Connection:=ADOConnection9;
    ADOQuery9.SQL.Clear;
    ADOQuery9.SQL.Add('SELECT * FROM '+DBGrid8.Fields[0].Value+'subb');
    ADOQuery9.Active:=true;
    DataSource9.DataSet:=ADOQuery9;
    DBGrid9.DataSource:=DataSource9;
    DBNavigator9.DataSource:=DataSource9;
    DBGrid9Width();

   end;

     //Список преподавателей
    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection4; //Было адо 1
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT * FROM thr');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
    DBGrid4.ReadOnly:=true;

    DBGrid4Width();

  ////////////////////////////////////////////////
     ADOConnection5.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection5.LoginPrompt:=false;
    ADOConnection5.Connected:=true;
    ADOQuery5.Connection:=ADOConnection5;
    ADOQuery5.SQL.Clear;
    ADOQuery5.SQL.Add('SELECT * FROM zam');
    ADOQuery5.Active:=true;
    DataSource5.DataSet:=ADOQuery5;
    DBGrid5.DataSource:=DataSource5;
    DBNavigator5.DataSource:=DataSource5;

    DBGrid5Width();




  DBLookupComboBox1.DataField:= 'Група';
  DBLookupComboBox1.DataSource:= DataSource5;
  DBLookupComboBox1.KeyField:= 'Група';
  DBLookupComboBox1.ListField:= 'Група';
  DBLookupComboBox1.ListSource:= DataSource1;
 // DBLookupComboBox1.TabOrder:= 4;


  DBLookupComboBox2.DataField:= 'Ким_замінюємо';
  DBLookupComboBox2.DataSource:= DataSource5;
  DBLookupComboBox2.KeyField:= 'Викладач';
  DBLookupComboBox2.ListField:= 'Викладач';
  DBLookupComboBox2.ListSource:= DataSource4;


  DBLookupComboBox4.DataField:= 'Кого_замінюємо';
  DBLookupComboBox4.DataSource:= DataSource5;
  DBLookupComboBox4.KeyField:= 'Викладач';
  DBLookupComboBox4.ListField:= 'Викладач';
  DBLookupComboBox4.ListSource:= DataSource4;

  DBLookupComboBox3.DataField:= 'Предмет1';
  DBLookupComboBox3.DataSource:= DataSource5;
  DBLookupComboBox3.KeyField:= 'Дисципліна';
  DBLookupComboBox3.ListField:= 'Дисципліна';
  DBLookupComboBox3.ListSource:= DataSource13;


  DBLookupComboBox5.DataField:= 'Предмет2';
  DBLookupComboBox5.DataSource:= DataSource5;
  DBLookupComboBox5.KeyField:= 'Дисципліна';
  DBLookupComboBox5.ListField:= 'Дисципліна';
  DBLookupComboBox5.ListSource:= DataSource13;


  end
  else
  begin




  end;


   {
     ADOConnection1.Connected:=true;
   ADOConnection2.Connected:=true;
   ADOConnection3.Connected:=true;
   ADOConnection4.Connected:=true;
   ADOConnection5.Connected:=true;
   ADOConnection7.Connected:=true;
   ADOConnection9.Connected:=true;
   ADOConnection10.Connected:=true;
   }
end;




procedure TForm1.DBGrid1Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid1.Columns[0].Width := 0; //ширина полей
    DBGrid1.Columns[1].Width := 80; //ширина полей
end;

procedure TForm1.DBGrid2DrawDataCell(Sender: TObject; const [Ref] Rect: TRect;
  Field: TField; State: TGridDrawState);
begin
//
end;
{
procedure TForm1.DBGrid2Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid2.Columns[0].Width := 30; //ширина полей
    DBGrid2.Columns[1].Width := 40; //ширина полей
    DBGrid2.Columns[2].Width := 40; //ширина полей
    DBGrid2.Columns[3].Width := 70; //ширина полей
    DBGrid2.Columns[4].Width := 0; //ширина полей
    DBGrid2.Columns[5].Width := 120; //ширина полей
    DBGrid2.Columns[6].Width := 120; //ширина полей
    DBGrid2.Columns[7].Width := 80; //ширина полей
end;

procedure TForm1.DBGrid3Width();   //Функция для установки ширини полей таблицы
var
  i:integer;
begin
    for i := 0 to 23 do DBGrid3.Columns[i].Width := 15; //ширина полей
end;

}


procedure TForm1.DBGrid4DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  with DBGrid4.Canvas do
    begin
      if (gdFocused in State) then Brush.Color:=16502182 else
      if (DBGrid4.DataSource.DataSet.RecNo mod 2)=0
      then Brush.Color:=16777215 else Brush.Color:=16510175;
    end;

DBGrid4.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TForm1.DBGrid4Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid4.Columns[0].Width := 0; //ширина полей
    DBGrid4.Columns[1].Width := 80; //ширина полей
end;

procedure TForm1.DBGrid5DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  with DBGrid5.Canvas do
    begin
      if (gdFocused in State) then Brush.Color:=16502182 else
      if (DBGrid5.DataSource.DataSet.RecNo mod 2)=0
      then Brush.Color:=16777215 else Brush.Color:=16510175;
    end;

DBGrid5.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TForm1.DBGrid5Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid5.Columns[0].Width := 0; //ширина полей
    DBGrid5.Columns[1].Width := 70; //ширина полей
    DBGrid5.Columns[2].Width := 70; //ширина полей
    DBGrid5.Columns[3].Width := 70; //ширина полей
    DBGrid5.Columns[4].Width := 70; //ширина полей
    DBGrid5.Columns[5].Width := 70; //ширина полей
    DBGrid5.Columns[6].Width := 70; //ширина полей
    DBGrid5.Columns[7].Width := 70; //ширина полей
    DBGrid5.Columns[8].Width := 70; //ширина полей
    DBGrid5.Columns[9].Width := 70; //ширина полей
    DBGrid5.Columns[10].Width := 70; //ширина полей
    DBGrid5.Columns[10].Width := 70; //ширина полей
end;

procedure TForm1.DBGrid6Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid6.Columns[0].Width := 0; //ширина полей
    DBGrid6.Columns[1].Width := 80; //ширина полей
end;

procedure TForm1.DBGrid7Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid7.Columns[0].Width := 0; //ширина полей
    DBGrid7.Columns[1].Width := 80; //ширина полей
    DBGrid7.Columns[2].Width := 80; //ширина полей
    DBGrid7.Columns[3].Width := 80; //ширина полей
    DBGrid7.Columns[4].Width := 80; //ширина полей
end;

procedure TForm1.DBGrid8Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid8.Columns[0].Width := 0; //ширина полей
    DBGrid8.Columns[1].Width := 80; //ширина полей
end;

procedure TForm1.DBGrid9Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid9.Columns[0].Width := 0; //ширина полей
    DBGrid9.Columns[1].Width := 100; //ширина полей
    DBGrid9.Columns[2].Width := 100; //ширина полей
    DBGrid9.Columns[3].Width := 100; //ширина полей
    DBGrid9.Columns[4].Width := 120; //ширина полей
    DBGrid9.Columns[5].Width := 150; //ширина полей
end;

procedure TForm1.DBLookupComboBox1CloseUp(Sender: TObject);
begin

  ADOQuery13.SQL.Clear;
  ADOQuery13.SQL.Add('SELECT * FROM AllRecords WHERE Група = '+''''+DBGrid1.Fields[0].Value+'''');
  ADOQuery13.Active:=true;
  DBGrid13width();


end;

procedure TForm1.DateTimePicker5CloseUp(Sender: TObject);
begin
     ADOQuery5.SQL.Clear;
     ADOQuery5.SQL.Add('SELECT * FROM zam');
    ADOQuery5.Active:=true;
    DBgrid5width();

    ADOQuery5.Insert;
end;

procedure TForm1.DBGrid10CellClick(Column: TColumn);
begin

     ADOQuery1.SQL.Clear;
  //   ADOQuery2.SQL.Clear;
  //   ADOQuery3.SQL.Clear;
     ADOQuery4.SQL.Clear;
     ADOQuery5.SQL.Clear;
     ADOQuery7.SQL.Clear;
     ADOQuery9.SQL.Clear;
     ADOQuery13.SQL.Clear;







    table:=DBGrid10.Fields[0].Value;

    ADOConnection13.Connected:=false;

     ADOConnection13.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection13.LoginPrompt:=false;
    ADOConnection13.Connected:=true;
    ADOQuery13.Connection:=ADOConnection13;
    ADOQuery13.SQL.Clear;
    ADOQuery13.SQL.Add('SELECT * FROM AllRecords');
    ADOQuery13.Active:=true;
    DataSource13.DataSet:=ADOQuery13;
    DBGrid13.DataSource:=DataSource13;
   // DBNavigator10.DataSource:=DataSource10;

   DBGrid13Width();

   { DB:=CoCatalog.Create;
    DB.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb');
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb');
    }
    ADOConnection1.Connected:=false;

    ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection1.LoginPrompt:=false;
    ADOConnection1.Connected:=true;
    ADOQuery1.Connection:=ADOConnection1;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT * FROM grp');
    ADOQuery1.Active:=true;
    DataSource1.DataSet:=ADOQuery1;
    DBGrid1.DataSource:=DataSource1;
  //  DBNavigator1.DataSource:=DataSource1;
    DBGrid1.ReadOnly:=true;

    DBGrid6.DataSource:=DataSource1;    //Запрет на редактирование
    DBGrid6.ReadOnly:=true;

    DBGrid8.DataSource:=DataSource1;   //Запрет на редактирование
    DBGrid8.ReadOnly:=true;

    DBGrid1Width();
    DBGrid6Width();
    DBGrid8Width();


 ADOConnection4.Connected:=false;

ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection4;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT * FROM thr');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
    DBGrid4.ReadOnly:=true;

    ADOConnection1.Connected:=true;

  ADOConnection5.Connected:=false;

ADOConnection5.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection5.LoginPrompt:=false;
    ADOConnection5.Connected:=true;
    ADOQuery5.Connection:=ADOConnection5;
    ADOQuery5.SQL.Clear;
    ADOQuery5.SQL.Add('SELECT * FROM zam');
    ADOQuery5.Active:=true;
    DataSource5.DataSet:=ADOQuery5;
    DBGrid5.DataSource:=DataSource5;
    DBNavigator5.DataSource:=DataSource5;

    DBGrid4Width();
    DBGrid5Width();


    if ADOQuery13.RecordCount<>0 then
begin
    ADOConnection1.Connected:=false;
    ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection1.LoginPrompt:=false;
    ADOConnection1.Connected:=true;
    ADOQuery1.Connection:=ADOConnection1;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT DISTINCT(Група) FROM AllRecords');
    ADOQuery1.Active:=true;
    DataSource1.DataSet:=ADOQuery1;
    DBGrid1.DataSource:=DataSource1;
end;

 if ADOQuery13.RecordCount<>0 then
begin
    ADOConnection4.Connected:=false;
    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection4;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT DISTINCT(Викладач) FROM AllRecords');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
end;




end;

procedure TForm1.DBGrid10DblClick(Sender: TObject);
begin
  edit7.Text:=DBGrid10.Fields[1].Value;
end;

procedure TForm1.DBGrid10DrawColumnCell(Sender: TObject;
  const [Ref] Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
  with DBGrid10.Canvas do
    begin
      if (gdFocused in State) then Brush.Color:=16502182 else
      if (DBGrid10.DataSource.DataSet.RecNo mod 2)=0
      then Brush.Color:=16777215 else Brush.Color:=16510175;
    end;

DBGrid10.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TForm1.DBGrid10Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid10.Columns[0].Width := 0; //ширина полей
    DBGrid10.Columns[1].Width := 120; //ширина полей
end;

procedure TForm1.DBGrid11DrawColumnCell(Sender: TObject;
  const [Ref] Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin
with DBGrid11.Canvas do
    begin
      if (gdFocused in State) then Brush.Color:=16502182 else
      if (DBGrid11.DataSource.DataSet.RecNo mod 2)=0
      then Brush.Color:=16777215 else Brush.Color:=16510175;
    end;

DBGrid11.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TForm1.DBGrid11Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid11.Columns[0].Width := 0; //ширина полей
    DBGrid11.Columns[1].Width := 230; //ширина полей
end;

procedure TForm1.DBGrid13DrawColumnCell(Sender: TObject;
  const [Ref] Rect: TRect; DataCol: Integer; Column: TColumn;
  State: TGridDrawState);
begin



with DBGrid13.Canvas do
    begin
      if (gdFocused in State) then Brush.Color:=16502182 else
      if (DBGrid13.DataSource.DataSet.RecNo mod 2)=0
      then Brush.Color:=16777215 else Brush.Color:=16510175;
    end;

     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='1' then
    DBGrid13.Canvas.Brush.Color := $000000 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='2' then
    DBGrid13.Canvas.Brush.Color := $FFFFFF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='3' then
    DBGrid13.Canvas.Brush.Color := $0000FF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='4' then
    DBGrid13.Canvas.Brush.Color := $00FF00 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='5' then
    DBGrid13.Canvas.Brush.Color := $FF0000 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='6' then
    DBGrid13.Canvas.Brush.Color := $00FFFF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='7' then
    DBGrid13.Canvas.Brush.Color := $FF00FF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='8' then
    DBGrid13.Canvas.Brush.Color := $FFFF00 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='9' then
    DBGrid13.Canvas.Brush.Color := $000080 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='10' then
    DBGrid13.Canvas.Brush.Color := $008000 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='11' then
    DBGrid13.Canvas.Brush.Color := $800000 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='12' then
    DBGrid13.Canvas.Brush.Color := $008080 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='13' then
    DBGrid13.Canvas.Brush.Color := $800080 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='14' then
    DBGrid13.Canvas.Brush.Color := $808000 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='15' then
    DBGrid13.Canvas.Brush.Color := $c0c0c0 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='16' then
    DBGrid13.Canvas.Brush.Color := $808080 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='17' then
    DBGrid13.Canvas.Brush.Color := $FF9999 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='18' then
    DBGrid13.Canvas.Brush.Color := $663399 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='19' then
    DBGrid13.Canvas.Brush.Color := $CCFFFF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='20' then
    DBGrid13.Canvas.Brush.Color := $FFFFCC ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='21' then
    DBGrid13.Canvas.Brush.Color := $660066 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='22' then
    DBGrid13.Canvas.Brush.Color := $8080FF ;
    if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='23' then
    DBGrid13.Canvas.Brush.Color := $CC6600 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='24' then
    DBGrid13.Canvas.Brush.Color := $FFCCCC ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='25' then
    DBGrid13.Canvas.Brush.Color := $800000 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='26' then
    DBGrid13.Canvas.Brush.Color := $FF00FF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='27' then
    DBGrid13.Canvas.Brush.Color := $00FFFF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='28' then
    DBGrid13.Canvas.Brush.Color := $FFFF00 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='29' then
    DBGrid13.Canvas.Brush.Color := $800080 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='30' then
    DBGrid13.Canvas.Brush.Color := $000080 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='31' then
    DBGrid13.Canvas.Brush.Color := $808000 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='32' then
    DBGrid13.Canvas.Brush.Color := $FF0000 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='33' then
    DBGrid13.Canvas.Brush.Color := $FFCC00 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='34' then
    DBGrid13.Canvas.Brush.Color := $FFFFCC ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='35' then
    DBGrid13.Canvas.Brush.Color := $CCFFCC ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='36' then
    DBGrid13.Canvas.Brush.Color := $99FFFF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='37' then
    DBGrid13.Canvas.Brush.Color := $FFCC99 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='38' then
    DBGrid13.Canvas.Brush.Color := $CC99FF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='39' then
    DBGrid13.Canvas.Brush.Color := $FF99CC ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='40' then
    DBGrid13.Canvas.Brush.Color := $99CCFF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='41' then
    DBGrid13.Canvas.Brush.Color := $FF6633 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='42' then
    DBGrid13.Canvas.Brush.Color := $CCCC33 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='43' then
    DBGrid13.Canvas.Brush.Color := $00CC99 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='44' then
    DBGrid13.Canvas.Brush.Color := $00CCFF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='45' then
    DBGrid13.Canvas.Brush.Color := $0099FF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='46' then
    DBGrid13.Canvas.Brush.Color := $0066FF ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='47' then
    DBGrid13.Canvas.Brush.Color := $996666 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='48' then
    DBGrid13.Canvas.Brush.Color := $969696 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='49' then
    DBGrid13.Canvas.Brush.Color := $663300 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='50' then
    DBGrid13.Canvas.Brush.Color := $669933 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='51' then
    DBGrid13.Canvas.Brush.Color := $003300 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='52' then
    DBGrid13.Canvas.Brush.Color := $003333 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='53' then
    DBGrid13.Canvas.Brush.Color := $003399 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='54' then
    DBGrid13.Canvas.Brush.Color := $663399 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='55' then
    DBGrid13.Canvas.Brush.Color := $993333 ;
     if DBGrid13.DataSource.DataSet.FieldByName('Color').AsString='56' then
    DBGrid13.Canvas.Brush.Color := $999999 ;





     DBGrid13.DefaultDrawColumnCell(Rect,DataCol,Column,State);


end;

procedure TForm1.DBGrid13Width();   //Функция для установки ширини полей таблицы
var
  i:integer;
begin
    DBGrid13.Columns[0].Width := 47; //ширина полей
    DBGrid13.Columns[1].Width := 52; //ширина полей
    DBGrid13.Columns[2].Width := 85; //ширина полей
    DBGrid13.Columns[3].Width := 80; //ширина полей
    DBGrid13.Columns[4].Width := 110; //ширина полей
    DBGrid13.Columns[5].Width := 110; //ширина полей
    DBGrid13.Columns[6].Width := 50; //ширина полей
    DBGrid13.Columns[7].Width := 20; //ширина полей


    for i := 8 to 31 do DBGrid13.Columns[i].Width := 15; //ширина полей


end;



procedure TForm1.Button10Click(Sender: TObject);
var
  Result,test,all,zam,grup:string;
  i,y:integer;
begin


  if RadioButton1.Checked=true then
  begin
      if DayOfWeek(DateTimePicker3.Date)<>7 then
  begin


    ShowMessage('Виберіть суботу!');
    Exit;
  end;
  end;



  if ListBox1.Count=0 then
  begin
     ShowMessage('Групи не вибрані!');
    Exit;
  end;



 {case DayOfWeek(DateTimePicker3.Date) of
  1: Result := 'Воскресенье';
  2: Result := 'Понедельник';
  3: Result := 'Вторник';
  4: Result := 'Среда';
  5: Result := 'Четверг';
  6: Result := 'Пятница';
  7: Result := 'Суббота';
  end;
  Edit4.Text:=Result;  }
  zam:='';
  if RadioButton1.Checked=true then
  zam:='Робоча субота';
  if RadioButton2.Checked=true then
  zam:='День на день';





   for i:=0 to ListBox1.Items.Count-1 do
   begin
   test:=ListBox1.Items[i];
    grup:=test;
   all:=all+ListBox1.Items[i]+',';

    for y:=0 to Length(test) do  // Удаление "-" из диапазона
  begin
    Delete(test,Pos('-',test),1);
  end;

  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM '+test+'subb');
  ADOQuery9.Active:=true;

    ADOQuery9.Insert;
    ADOQuery9.FieldByName('Група').Value:=grup;
    ADOQuery9.FieldByName('День_який_замін').Value:=DateToStr(DateTimePicker3.Date);
    ADOQuery9.FieldByName('День_яким_замін').Value:=DateToStr(DateTimePicker4.Date);
    ADOQuery9.FieldByName('Вид_заміни').Value:=zam;
    ADOQuery9.FieldByName('Коментар').Value:=edit4.text;
    ADOQuery9.Post;

   end;

   //ShowMessage('Для груп: '+all+' була додана робітнича субота '+DateToStr(DateTimePicker3.Date));
   ListBox1.Clear;





  DBGrid9Width();


end;



procedure TForm1.Button11Click(Sender: TObject);
var
kol,i:integer;
begin
  kol:=ADOQuery1.RecordCount;


Form1.DBGrid8.DataSource.DataSet.First;


   for i:=1 to ADOQuery1.RecordCount do
   begin
    DBGrid8.Fields[0].Value;
   ListBox1.Items.add(DBGrid8.Fields[0].Value);

   Form1.DBGrid8.DataSource.DataSet.Next;
   end;


end;

procedure TForm1.Button12Click(Sender: TObject);
begin
 ListBox1.Clear;
end;

procedure TForm1.Button13Click(Sender: TObject);
var
  kol,i:integer;
  mycell:string;
begin
  kol:=ADOQuery1.RecordCount;


  ADOQuery1.FindFirst;
  ADOQuery9.SQL.Clear;
   ADOQuery9.SQL.Add('SELECT * FROM ТВ141subb FULL JOIN ТВ142subb');
   // );
  // ADOQuery9.SQL.Add('SELECT * FROM ТВ142subb WHERE Вид_заміни = ''Робоча субота''');
    ADOQuery9.Active:=true;

  {for i:=0 to kol do
    begin
  ADOQuery9.SQL.Clear;
   ADOQuery9.SQL.Add('SELECT * FROM '+DBGrid8.Fields[0].Value+'subb WHERE Вид_заміни = ''Робоча субота''');
  ADOQuery9.Active:=true;

  ADOQuery1.FindNext;
    end;




  //DBGrid8.DataSource.DataSet.RecNo;
 // DBGrid8.Fields[0].Value;

  Panel3.Caption:='Група '+DBGrid8.Fields[1].Value;


 { ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM ТВ15п1subb WHERE Вид_заміни = ''Робоча субота''');
  ADOQuery9.Active:=true; }

 // ADOQuery1.SQL.Clear;
 //  ADOQuery1.SQL.Add('SELECT * FROM Test WHERE names=''Andrey''');
 //  ADOQuery1.Active:=True;




  DBGrid9Width();



end;

procedure TForm1.Button14Click(Sender: TObject);
begin
    ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM ТВ15п1subb WHERE Вид_заміни = ''День на день''');
  ADOQuery9.Active:=true;

  DBGrid9Width();
end;

procedure TForm1.Button16Click(Sender: TObject);
begin

{    Tables.Columns.Append('Дата',adVarWChar,255);
    Tables.Columns.Append('Група',adVarWChar,255);
    Tables.Columns.Append('№_ленты',adVarWChar,255);
    Tables.Columns.Append('Ким_замінюємо',adVarWChar,255);
    Tables.Columns.Append('Предмет1',adVarWChar,255);
    Tables.Columns.Append('Кого_замінюємо',adVarWChar,255);
    Tables.Columns.Append('Предмет2',adVarWChar,255);
    Tables.Columns.Append('Аудиторія',adVarWChar,255);
    Tables.Columns.Append('Номер_наказу',adVarWChar,255);
    Tables.Columns.Append('Коментар',adVarWChar,255);
    Tables.Columns.Append('Тип_заміни',adVarWChar,255);

       DateTimePicker5.Enabled:=false;
   DBLookupComboBox1.Enabled:=false;
   DBLookupComboBox2.Enabled:=false;
   DBLookupComboBox3.Enabled:=false;
   DBLookupComboBox4.Enabled:=false;
   DBLookupComboBox5.Enabled:=false;
   comboBox2.Enabled:=false;
   comboBox3.Enabled:=false;
   edit11.Enabled:=false;
   edit12.Enabled:=false;
   edit13.Enabled:=false;


    }

 //  ADOQuery5.SQL.Clear;
  //   ADOQuery5.SQL.Add('SELECT * FROM zam');
  //  ADOQuery5.Active:=true;


  //  ADOQuery5.Insert;

    ADOQuery5.FieldByName('Дата').Value:=DateToStr(DateTimePicker5.Date);
    ADOQuery5.FieldByName('Група').Value:=DBLookupComboBox1.Text;
    ADOQuery5.FieldByName('№_ленти1').Value:=ComboBox2.Items[ComboBox2.ItemIndex];
    ADOQuery5.FieldByName('№_ленти2').Value:=ComboBox3.Items[ComboBox3.ItemIndex];
    ADOQuery5.FieldByName('Ким_замінюємо').Value:=DBLookupComboBox2.Text;
    ADOQuery5.FieldByName('Предмет1').Value:=DBLookupComboBox3.Text;
     ADOQuery5.FieldByName('Кого_замінюємо').Value:=DBLookupComboBox4.Text;
    ADOQuery5.FieldByName('Предмет2').Value:=DBLookupComboBox5.Text;
    ADOQuery5.FieldByName('Аудиторія').Value:=edit11.Text;
    ADOQuery5.FieldByName('Номер_наказу').Value:=edit12.Text;
    ADOQuery5.FieldByName('Коментар').Value:=edit13.Text;
    ADOQuery5.FieldByName('Тип_заміни').Value:=RadioGroup1.Items[RadioGroup1.ItemIndex];

    ADOQuery5.Post;


    DBgrid5width();


  {  DBLookupComboBox1.DataField:= 'Група';
  DBLookupComboBox1.DataSource:= DataSource5;
  DBLookupComboBox1.KeyField:= 'Група';
  DBLookupComboBox1.ListField:= 'Група';
  DBLookupComboBox1.ListSource:= DataSource1; }
end;

procedure TForm1.Button17Click(Sender: TObject);
begin
 PageControl1.Visible:=false;
end;

procedure TForm1.Button1Click(Sender: TObject); //Кнопка открытия Exel файла
begin
  excel:=CreateOleObject('Excel.Application');
  excel.WorkBooks.Open(edit10.text);
  excel.Visible:=true;
end;

procedure TForm1.Button23Click(Sender: TObject);
var
  y:integer;
  tableInput:string;
begin


    tableInput:=ComboBox5.Items[ComboBox5.ItemIndex];


    ADOQuery10.SQL.Clear;
    ADOQuery10.SQL.Add('SELECT * FROM ListAcademicYear');
    ADOQuery10.Active:=true;

         for y:=0 to Length(tableInput) do  // Удаление " " из диапазона
  begin
    Delete(tableInput,Pos(' ',tableInput),1);
  end;



     ADOQuery10.Insert;
    ADOQuery10.FieldByName('IDНавч_рік').AsString :=tableInput;
    ADOQuery10.FieldByName('Навч_рік').AsString :=tableInput;
    ADOQuery10.Post;

    DBGrid10Width();

    table:=DBGrid10.Fields[0].Value;









    DB:=CoCatalog.Create;
    DB.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb');
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb');


    //Создание таблицы для групп
    ADOConnection1.Connected:=false;
    Tables:=CoTable.Create;
    Tables.Name:='grp';
    Tables.ParentCatalog:=DB;
    DB.Tables.Append(Tables);
    Columns:=CoColumn.Create;
    with Columns do
    begin
      ParentCatalog:=DB;
      Name:='id';
      type_:=adInteger;
      Properties['Autoincrement'].Value:=True;
      Properties['Description'].Value:='Ключевое поле';
    end;
    Tables.Columns.Append('IDГруп',adVarWChar,255);
    Tables.Columns.Append('Група',adVarWChar,255);



    ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection1.LoginPrompt:=false;
    ADOConnection1.Connected:=true;
    ADOQuery1.Connection:=ADOConnection1;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT * FROM grp');
    ADOQuery1.Active:=true;
    DataSource1.DataSet:=ADOQuery1;
    DBGrid1.DataSource:=DataSource1;
  //  DBNavigator1.DataSource:=DataSource1;
    DBGrid1.ReadOnly:=true;

    DBGrid6.DataSource:=DataSource1;    //Запрет на редактирование
    DBGrid6.ReadOnly:=true;

    DBGrid8.DataSource:=DataSource1;   //Запрет на редактирование
    DBGrid8.ReadOnly:=true;

    DBGrid1Width();
    DBGrid6Width();
    DBGrid8Width();

    /////////////////////////////////////


    //Создание таблицы для преподавателей
   // ADOConnection1.Connected:=false;
    ADOConnection4.Connected:=false;

    Tables:=CoTable.Create;
    Tables.Name:='thr';
    Tables.ParentCatalog:=DB;
    DB.Tables.Append(Tables);
    Columns:=CoColumn.Create;
    with Columns do
    begin
      ParentCatalog:=DB;
      Name:='id';
      type_:=adInteger;
      Properties['Autoincrement'].Value:=True;
      Properties['Description'].Value:='Ключевое поле';
    end;
    Tables.Columns.Append('IDВикладача',adVarWChar,255);
    Tables.Columns.Append('Викладач',adVarWChar,255);


    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection4;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT * FROM thr');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
    DBGrid4.ReadOnly:=true;

    ADOConnection1.Connected:=true;


    /////////////////////////////////////////////////////////////////////////



    // Создание таблицы замен учителей
     ADOConnection5.Connected:=false;

    Tables:=CoTable.Create;
    Tables.Name:='zam';

    Tables.ParentCatalog:=DB;
    DB.Tables.Append(Tables);
    Columns:=CoColumn.Create;
    with Columns do
    begin
      ParentCatalog:=DB;
      Name:='id';
      type_:=adInteger;
      Properties['Autoincrement'].Value:=True;
      Properties['Description'].Value:='Ключевое поле';
    end;

    Tables.Columns.Append(Columns,0,0);
    Tables.Columns.Append('Дата',adVarWChar,255);
    Tables.Columns.Append('Група',adVarWChar,255);
    Tables.Columns.Append('№_ленти1',adVarWChar,255);
    Tables.Columns.Append('№_ленти2',adVarWChar,255);
    Tables.Columns.Append('Ким_замінюємо',adVarWChar,255);
    Tables.Columns.Append('Предмет1',adVarWChar,255);
    Tables.Columns.Append('Кого_замінюємо',adVarWChar,255);
    Tables.Columns.Append('Предмет2',adVarWChar,255);
    Tables.Columns.Append('Аудиторія',adVarWChar,255);
    Tables.Columns.Append('Номер_наказу',adVarWChar,255);
    Tables.Columns.Append('Коментар',adVarWChar,255);
    Tables.Columns.Append('Тип_заміни',adVarWChar,255);


    ADOConnection5.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
    ADOConnection5.LoginPrompt:=false;
    ADOConnection5.Connected:=true;
    ADOQuery5.Connection:=ADOConnection5;
    ADOQuery5.SQL.Clear;
    ADOQuery5.SQL.Add('SELECT * FROM zam');
    ADOQuery5.Active:=true;
    DataSource5.DataSet:=ADOQuery5;
    DBGrid5.DataSource:=DataSource5;
    DBNavigator5.DataSource:=DataSource5;





    ////////////////////////////////////////////////////////////////



    DBGrid4Width();
    DBGrid5Width();



    //Создание таблицы для групп
    ADOConnection13.Connected:=false;
    Tables:=CoTable.Create;
    Tables.Name:='AllRecords';
    Tables.ParentCatalog:=DB;
    DB.Tables.Append(Tables);
    Columns:=CoColumn.Create;
    with Columns do
    begin
      ParentCatalog:=DB;
      Name:='id';
      type_:=adInteger;
      Properties['Autoincrement'].Value:=True;
      Properties['Description'].Value:='Ключевое поле';
    end;

    Tables.Columns.Append('Бютжет',adVarWChar,0);
    Tables.Columns.Append('Контракт',adVarWChar,0);
    Tables.Columns.Append('Комментар',adVarWChar,0);
    Tables.Columns.Append('Група',adVarWChar,255);
    Tables.Columns.Append('Дисципліна',adVarWChar,255);
    Tables.Columns.Append('Викладач',adVarWChar,255);
    Tables.Columns.Append('Семестр',adVarWChar,255);
    Tables.Columns.Append('Color',adVarWChar,255);
    Tables.Columns.Append('ПнЧ1',adVarWChar,255);
    Tables.Columns.Append('ПнЧ2',adVarWChar,255);
    Tables.Columns.Append('ВтЧ1',adVarWChar,255);
    Tables.Columns.Append('ВтЧ2',adVarWChar,255);
    Tables.Columns.Append('СрЧ1',adVarWChar,255);
    Tables.Columns.Append('СрЧ2',adVarWChar,255);
    Tables.Columns.Append('ЧтЧ1',adVarWChar,255);
    Tables.Columns.Append('ЧтЧ2',adVarWChar,255);
    Tables.Columns.Append('ПтЧ1',adVarWChar,255);
    Tables.Columns.Append('ПтЧ2',adVarWChar,255);
    Tables.Columns.Append('СбЧ1',adVarWChar,255);
    Tables.Columns.Append('СбЧ2',adVarWChar,255);
    Tables.Columns.Append('НдЧ1',adVarWChar,255);
    Tables.Columns.Append('НдЧ2',adVarWChar,255);
    Tables.Columns.Append('ПнЗ1',adVarWChar,255);
    Tables.Columns.Append('ПнЗ2',adVarWChar,255);
    Tables.Columns.Append('ВтЗ1',adVarWChar,255);
    Tables.Columns.Append('ВтЗ2',adVarWChar,255);
    Tables.Columns.Append('СрЗ1',adVarWChar,255);
    Tables.Columns.Append('СрЗ2',adVarWChar,255);
    Tables.Columns.Append('ЧтЗ1',adVarWChar,255);
    Tables.Columns.Append('ЧтЗ2',adVarWChar,255);
    Tables.Columns.Append('ПтЗ1',adVarWChar,255);
    Tables.Columns.Append('ПтЗ2',adVarWChar,255);

    ADOConnection13.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection13.LoginPrompt:=false;
    ADOConnection13.Connected:=true;
    ADOQuery13.Connection:=ADOConnection13;
    ADOQuery13.SQL.Clear;
    ADOQuery13.SQL.Add('SELECT * FROM AllRecords');
    ADOQuery13.Active:=true;
    DataSource13.DataSet:=ADOQuery13;
    DBGrid13.DataSource:=DataSource13;
   // DBNavigator10.DataSource:=DataSource10;

   DBGrid13Width();


  end;



procedure TForm1.Button25Click(Sender: TObject);
begin
  if not OpenDialog1.Execute then exit;
  edit10.text:=OpenDialog1.FileName;


end;

procedure TForm1.Button26Click(Sender: TObject);
begin
  edit7.Text:=DBGrid10.Fields[1].Value;
end;

procedure TForm1.Button27Click(Sender: TObject);
begin

  if DeleteFile('Backup\'+DBGrid11.Fields[1].Value) then
  ShowMessage('Файл успішно видалено!')
  else ShowMessage('Помилка при видаленні!');

  DBGrid11.DataSource.DataSet.Delete;
end;

procedure TForm1.Button28Click(Sender: TObject);
begin
  BackupSave();
end;

procedure TForm1.Button29Click(Sender: TObject);
var
  strName,strTable:string;
begin

 if ADOQuery11.RecordCount<>0   then
 begin

   DB:=nil;
   table:='';



  ADOConnection1.Close;
 // ADOConnection2.Close;
  //ADOConnection3.Close;
  ADOConnection4.Close;
  ADOConnection5.Close;
  ADOConnection7.Close;
  ADOConnection9.Close;
  ADOConnection13.Close;

   strTable:=edit7.Text+'.mdb';

  DeleteFile(strTable);



  strName:='Backup\'+DBGrid11.Fields[1].Value;


 if CopyFile(PWideChar(strName),PWideChar(strTable),true) then
 begin
     ShowMessage('База даних успішно збережена!')
 end
 else
  ShowMessage('Помилка у зберіганні бази даних!');


 end
 else
 ShowMessage('Записів для відкату бази даних немає!');



end;

procedure TForm1.Button2Click(Sender: TObject);
var
  i:integer;
begin

Form1.DBGrid6.DataSource.DataSet.First;


   for i:=1 to ADOQuery1.RecordCount do
   begin
    DBGrid6.Fields[0].Value;
   ListBox2.Items.add(DBGrid8.Fields[0].Value);

   Form1.DBGrid6.DataSource.DataSet.Next;
   end;

end;

procedure TForm1.Button30Click(Sender: TObject);
var
  rsp,test,graf,subb,zam:string;
  y,x1,x2,add:string;
  nom,i:integer;
  z:integer;
begin


  table:=DBGrid10.Fields[0].Value;
 y:=excel.Selection.Address;



  for i:=0 to Length(y) do  // Удаление $ из диапазона
  begin
  Delete(y,Pos('$',y),1);
  end;



  x1:=Copy(y,2,Pos(':',y)-2);   // Удаление лишних знаков из диапазона
  x2:=Copy(y,Pos(':',y)+3,Length(y));







  nom:=strtoint(x2)-strtoint(x1);  // Количество записей


    for i:=0 to nom  do
    begin
    add:=inttostr(strtoint(x1)+i);  // Номер елемента



    if (excel.WorkBooks[1].WorkSheets[2].Range['C'+add].Value<>'')
    and (excel.WorkBooks[1].WorkSheets[2].Range['G'+add].Value<>'')
     then
     begin
       ADOQuery13.Insert;
      ADOQuery13.FieldByName('Група').Value :=excel.WorkBooks[1].WorkSheets[2].Range['C'+add].Value;
      ADOQuery13.FieldByName('Дисципліна').Value :=excel.WorkBooks[1].WorkSheets[2].Range['G'+add].Value;
      ADOQuery13.FieldByName('Викладач').Value :=excel.WorkBooks[1].WorkSheets[2].Range['BD'+add].Value;
      ADOQuery13.FieldByName('Бютжет').Value :='';
      ADOQuery13.FieldByName('Контракт').Value :='';
      ADOQuery13.FieldByName('Комментар').Value :='';
      ADOQuery13.FieldByName('Color').Value :=excel.WorkBooks[1].WorkSheets[2].Range['G'+add].Interior.ColorIndex;

      ADOQuery13.FieldByName('Семестр').Value :='-';



      ADOQuery13.FieldByName('ПнЧ1').Value :=' ';
      ADOQuery13.FieldByName('ПнЧ2').Value :=' ';
      ADOQuery13.FieldByName('ВтЧ1').Value :=' ';
      ADOQuery13.FieldByName('ВтЧ2').Value :=' ';
      ADOQuery13.FieldByName('СрЧ1').Value :=' ';
      ADOQuery13.FieldByName('СрЧ2').Value :=' ';
      ADOQuery13.FieldByName('ЧтЧ1').Value :=' ';
      ADOQuery13.FieldByName('ЧтЧ2').Value :=' ';
      ADOQuery13.FieldByName('ПтЧ1').Value :=' ';
      ADOQuery13.FieldByName('ПтЧ2').Value :=' ';
      ADOQuery13.FieldByName('СбЧ1').Value :=' ';
      ADOQuery13.FieldByName('СбЧ2').Value :=' ';
      ADOQuery13.FieldByName('НдЧ1').Value :=' ';
      ADOQuery13.FieldByName('НдЧ2').Value :=' ';
      ADOQuery13.FieldByName('ПнЗ1').Value :=' ';
      ADOQuery13.FieldByName('ПнЗ2').Value :=' ';
      ADOQuery13.FieldByName('ВтЗ1').Value :=' ';
      ADOQuery13.FieldByName('ВтЗ2').Value :=' ';
      ADOQuery13.FieldByName('СрЗ1').Value :=' ';
      ADOQuery13.FieldByName('СрЗ2').Value :=' ';
      ADOQuery13.FieldByName('ЧтЗ1').Value :=' ';
      ADOQuery13.FieldByName('ЧтЗ2').Value :=' ';
      ADOQuery13.FieldByName('ПтЗ1').Value :=' ';
      ADOQuery13.FieldByName('ПтЗ2').Value :=' ';

      ADOQuery13.Post;
     end;


    end;

     DBGrid13Width();


if ADOQuery13.RecordCount<>0 then
begin
    ADOConnection1.Connected:=false;
    ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection1.LoginPrompt:=false;
    ADOConnection1.Connected:=true;
    ADOQuery1.Connection:=ADOConnection1;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT DISTINCT(Група) FROM AllRecords');
    ADOQuery1.Active:=true;
    DataSource1.DataSet:=ADOQuery1;
    DBGrid1.DataSource:=DataSource1;
end;

 if ADOQuery13.RecordCount<>0 then
begin
    ADOConnection4.Connected:=false;
    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection4;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT DISTINCT(Викладач) FROM AllRecords');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
end;







     ADOConnection7.Connected:=false;


Form1.DBGrid1.DataSource.DataSet.First;

for i:=1 to ADOQuery1.RecordCount do
begin


  test:=DBGrid1.Fields[0].Value;



  for z:=0 to Length(test) do  // Удаление - из диапазона
  begin
    Delete(test,Pos('-',test),1);
  end;


  graf:=test+'graf';
  subb:=test+'subb';

   ADOConnection7.Connected:=false;

  //Создание таблицы групп
  Tables:=CoTable.Create;
  Tables.Name:=graf;

  Tables.ParentCatalog:=DB;
  DB.Tables.Append(Tables);
  Columns:=CoColumn.Create;
  with Columns do
  begin
    ParentCatalog:=DB;
    Name:='id';
    type_:=adInteger;
    Properties['Autoincrement'].Value:=True;
    Properties['Description'].Value:='Ключевое поле';
  end;

  Tables.Columns.Append(Columns,0,0);
  Tables.Columns.Append('Поч_дата',adVarWChar,255);
  Tables.Columns.Append('Кін_дата',adVarWChar,255);
  Tables.Columns.Append('Вид_навч',adVarWChar,255);
  Tables.Columns.Append('Комент',adVarWChar,255);
  Tables.Columns.Append('Група',adVarWChar,255);

  ADOConnection7.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
  ADOConnection7.LoginPrompt:=false;
  ADOConnection7.Connected:=true;
  ADOQuery7.Connection:=ADOConnection7;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('SELECT * FROM '+graf);
  ADOQuery7.Active:=true;
  DataSource7.DataSet:=ADOQuery7;
  DBGrid7.DataSource:=DataSource7;
  DBNavigator7.DataSource:=DataSource7;


  // Создание таблицы суббботы

  ADOConnection9.Connected:=false;

  Tables:=CoTable.Create;
  Tables.Name:=subb;

  Tables.ParentCatalog:=DB;
  DB.Tables.Append(Tables);
  Columns:=CoColumn.Create;
  with Columns do
  begin
    ParentCatalog:=DB;
    Name:='id';
    type_:=adInteger;
    Properties['Autoincrement'].Value:=True;
    Properties['Description'].Value:='Ключевое поле';
  end;

  Tables.Columns.Append(Columns,0,0);
  Tables.Columns.Append('Група',adVarWChar,255);
  Tables.Columns.Append('День_який_замін',adVarWChar,255);
  Tables.Columns.Append('День_яким_замін',adVarWChar,255);
  Tables.Columns.Append('Вид_заміни',adVarWChar,255);
  Tables.Columns.Append('Коментар',adVarWChar,255);




  ADOConnection9.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
  ADOConnection9.LoginPrompt:=false;
  ADOConnection9.Connected:=true;
  ADOQuery9.Connection:=ADOConnection9;
  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM '+subb);
  ADOQuery9.Active:=true;
  DataSource9.DataSet:=ADOQuery9;
  DBGrid9.DataSource:=DataSource9;
  DBNavigator9.DataSource:=DataSource9;


   Form1.DBGrid1.DataSource.DataSet.Next;

   button30.Enabled:=false;
  //////////////////////////////////////////////////////////////////
end;



   // WHERE Група = ''ТВ-15-1'''
  // Select distinct(ID) from ...



end;

procedure TForm1.Button3Click(Sender: TObject);
begin
  Panel1.Caption:='Перелік усіх груп';
  ADOQuery13.SQL.Clear;
  ADOQuery13.SQL.Add('SELECT * FROM AllRecords');
  ADOQuery13.Active:=true;
  DBGrid13width();
end;

procedure TForm1.Button6Click(Sender: TObject); //Кнопка поиска преподавателей
begin
  ListBox2.Clear;
end;

procedure TForm1.Button9Click(Sender: TObject);
var
  Result,test,all,zam,grup:string;
  i,y:integer;
begin





     for i:=0 to ListBox2.Items.Count-1 do
   begin
   test:=ListBox2.Items[i];
   grup:=test;
   all:=all+ListBox2.Items[i]+',';

    for y:=0 to Length(test) do  // Удаление "-" из диапазона
  begin
    Delete(test,Pos('-',test),1);
  end;

     ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('SELECT * FROM '+test+'graf');
    ADOQuery7.Active:=true;

    ADOQuery7.Insert;
    ADOQuery7.FieldByName('Поч_дата').Value:=DateToStr(DateTimePicker1.Date);
    ADOQuery7.FieldByName('Кін_дата').Value:=DateToStr(DateTimePicker2.Date);
    ADOQuery7.FieldByName('Вид_навч').Value:=ComboBox1.Items[ComboBox1.ItemIndex];
    ADOQuery7.FieldByName('Комент').Value:=Edit3.Text;
    ADOQuery7.FieldByName('Група').Value:=grup;
    ADOQuery7.Post;

   end;






  DBGrid7Width();

end;

procedure TForm1.CheckBox1Click(Sender: TObject);
begin
  if checkBox1.Checked=true then
  begin
    spinedit1.Enabled:=true;
    timer2.Enabled:=true;
  end
  else
  begin
    timer2.Enabled:=false;
    spinedit1.Enabled:=false;

  end;

end;

procedure TForm1.DBGrid1CellClick(Column: TColumn); //Выбор группы из списка для просмотра расписания
var
  i,nomSemestr: integer;

begin
  nomSemestr:=RadioGroup2.ItemIndex;

if nomSemestr=0 then
begin
  Panel1.Caption:='Група '+DBGrid1.Fields[0].Value;
  ADOQuery13.SQL.Clear;
  ADOQuery13.SQL.Add('SELECT * FROM AllRecords WHERE Група = '+''''+DBGrid1.Fields[0].Value+'''');
   ADOQuery13.SQL.Add('AND Семестр = ''1''');
  ADOQuery13.Active:=true;
  DBGrid13width();
end;

if nomSemestr=1 then
begin
  Panel1.Caption:='Група '+DBGrid1.Fields[0].Value;
  ADOQuery13.SQL.Clear;
  ADOQuery13.SQL.Add('SELECT * FROM AllRecords WHERE Група = '+''''+DBGrid1.Fields[0].Value+'''');
   ADOQuery13.SQL.Add('AND Семестр = ''2''');
  ADOQuery13.Active:=true;
  DBGrid13width();
end;

if nomSemestr=2 then
begin
  Panel1.Caption:='Група '+DBGrid1.Fields[0].Value;
  ADOQuery13.SQL.Clear;
  ADOQuery13.SQL.Add('SELECT * FROM AllRecords WHERE Група = '+''''+DBGrid1.Fields[0].Value+'''');
  ADOQuery13.Active:=true;
  DBGrid13width();
end;


end;

procedure TForm1.DBGrid1DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  with DBGrid1.Canvas do
    begin
      if (gdFocused in State) then Brush.Color:=16502182 else
      if (DBGrid1.DataSource.DataSet.RecNo mod 2)=0
      then Brush.Color:=16777215 else Brush.Color:=16510175;
    end;

DBGrid1.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TForm1.DBGrid6CellClick(Column: TColumn);
var
  i: integer;
  test,graf:string;
begin
//  DBGrid6.DataSource.DataSet.RecNo;
 // DBGrid6.Fields[0].Value;


 test:=DBGrid6.Fields[0].Value;


 for i:=0 to Length(test) do  // Удаление - из диапазона
  begin
    Delete(test,Pos('-',test),1);
  end;

   graf:=test+'graf';
  Panel2.Caption:='Група '+DBGrid6.Fields[0].Value;

   ADOConnection7.Connected:=false;


  ADOConnection7.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
  ADOConnection7.LoginPrompt:=false;
  ADOConnection7.Connected:=true;
  ADOQuery7.Connection:=ADOConnection7;
  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('SELECT * FROM '+graf);
  ADOQuery7.Active:=true;
  DataSource7.DataSet:=ADOQuery7;
  DBGrid7.DataSource:=DataSource7;
  DBNavigator7.DataSource:=DataSource7;


 // DBGrid6Width();
  DBGrid7Width();


end;

procedure TForm1.DBGrid6DblClick(Sender: TObject);
var
  i,check:integer;
begin
//ListBox1.Items.i:=DBGrid8.Fields[0].Value  ;

      if Listbox2.Count=2 then
    exit;

  if ListBox2.Items.Count=0 then
  ListBox2.Items.add(DBGrid6.Fields[0].Value);


   check:=0;
  for i:=0 to ListBox2.Items.Count-1 do
     begin
   if DBGrid6.Fields[0].Value=ListBox2.Items[i]  then
   begin
    check:=1;
    break;


   end
   else
    check:=0;


     end;

    if check=0 then  ListBox2.Items.add(DBGrid6.Fields[0].Value);





end;

procedure TForm1.DBGrid6DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  with DBGrid6.Canvas do
    begin
      if (gdFocused in State) then Brush.Color:=16502182 else
      if (DBGrid6.DataSource.DataSet.RecNo mod 2)=0
      then Brush.Color:=16777215 else Brush.Color:=16510175;
    end;

DBGrid6.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TForm1.DBGrid7DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin



  if DBGrid7.DataSource.DataSet.FieldByName('Вид_навч').AsString='Практика' then
    DBGrid7.Canvas.Brush.Color := $000DC10D
    else DBGrid1.Canvas.Brush.Color := $00BFFFBF;

  if DBGrid7.DataSource.DataSet.FieldByName('Вид_навч').AsString='Канікули' then
    DBGrid7.Canvas.Brush.Color := $0019A2EE
    else DBGrid1.Canvas.Brush.Color := $00BFFFBF;

  if DBGrid7.DataSource.DataSet.FieldByName('Вид_навч').AsString='Сесія' then
    DBGrid7.Canvas.Brush.Color := $0019A2EE
    else DBGrid1.Canvas.Brush.Color := $00BFFFBF;

    DBGrid7.DefaultDrawColumnCell(Rect,DataCol,Column,State);

end;

procedure TForm1.DBGrid8CellClick(Column: TColumn);
var
  i: integer;
  test,subb:string;
begin
//  DBGrid6.DataSource.DataSet.RecNo;
 // DBGrid6.Fields[0].Value;


 test:=DBGrid6.Fields[0].Value;


 for i:=0 to Length(test) do  // Удаление - из диапазона
  begin
    Delete(test,Pos('-',test),1);
  end;

   subb:=test+'subb';
  Panel3.Caption:='Група '+DBGrid8.Fields[0].Value;

   ADOConnection9.Connected:=false;


  ADOConnection9.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb;';
  ADOConnection9.LoginPrompt:=false;
  ADOConnection9.Connected:=true;
  ADOQuery9.Connection:=ADOConnection9;
  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM '+subb);
  ADOQuery9.Active:=true;
  DataSource9.DataSet:=ADOQuery9;
  DBGrid9.DataSource:=DataSource9;
  DBNavigator9.DataSource:=DataSource9;


 // DBGrid6Width();
  DBGrid9Width();


end;

procedure TForm1.DBGrid8DblClick(Sender: TObject);
var
  i,check:integer;
begin
//ListBox1.Items.i:=DBGrid8.Fields[0].Value  ;

if ListBox1.Items.Count=0 then
ListBox1.Items.add(DBGrid8.Fields[0].Value);


   check:=0;
  for i:=0 to ListBox1.Items.Count-1 do
     begin
   if DBGrid8.Fields[0].Value=ListBox1.Items[i]  then
   begin
    check:=1;
    break;


   end
   else
    check:=0;


     end;

    if check=0 then  ListBox1.Items.add(DBGrid8.Fields[0].Value);



end;

procedure TForm1.DBGrid8DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
  with DBGrid8.Canvas do
    begin
      if (gdFocused in State) then Brush.Color:=16502182 else
      if (DBGrid8.DataSource.DataSet.RecNo mod 2)=0
      then Brush.Color:=16777215 else Brush.Color:=16510175;
    end;

DBGrid8.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TForm1.DBGrid9DrawColumnCell(Sender: TObject; const [Ref] Rect: TRect;
  DataCol: Integer; Column: TColumn; State: TGridDrawState);
begin
 if DBGrid9.DataSource.DataSet.FieldByName('Вид_заміни').AsString='Робоча субота' then
    DBGrid9.Canvas.Brush.Color := $000DC10D
    else DBGrid1.Canvas.Brush.Color := $00BFFFBF;


 if DBGrid9.DataSource.DataSet.FieldByName('Вид_заміни').AsString='День на день' then
    DBGrid9.Canvas.Brush.Color := $0067C76B
    else DBGrid1.Canvas.Brush.Color := $00BFFFBF;


  DBGrid9.DefaultDrawColumnCell(Rect,DataCol,Column,State);
end;

procedure TForm1.Edit2Change(Sender: TObject);
begin
  if not ADOQuery4.Locate('Викладач',Edit2.Text,[loCaseInsensitive, loPartialKey])then
  //Message('Запись не найдена');
end;

procedure TForm1.Edit5Change(Sender: TObject);
begin
  if not ADOQuery1.Locate('Група',Edit5.Text,[loCaseInsensitive, loPartialKey])then
  //Message('Запись не найдена');
end;

procedure TForm1.Edit6Change(Sender: TObject);
begin
 if not ADOQuery1.Locate('Група',Edit6.Text,[loCaseInsensitive, loPartialKey])then
  //Message('Запись не найдена');
end;

procedure TForm1.Form1Create(Sender: TObject); //Создание и чтение файлов при запуске
var
  i:integer;
  rsp:string;
begin
   DateTimePicker5.Enabled:=false;
   DBLookupComboBox1.Enabled:=false;
   DBLookupComboBox2.Enabled:=false;
   DBLookupComboBox3.Enabled:=false;
   DBLookupComboBox4.Enabled:=false;
   DBLookupComboBox5.Enabled:=false;
   comboBox2.Enabled:=false;
   comboBox3.Enabled:=false;
   edit11.Enabled:=false;
   edit12.Enabled:=false;
   edit13.Enabled:=false;





      if checkBox1.Checked=true then
  begin
    spinedit1.Enabled:=true;

  end
  else
  begin
    spinedit1.Enabled:=false;
  end;


  IniFile:=TIniFile.Create(ExtractFilePath(ParamStr(0)) + 'Config.ini'); // Создание файла настроек

  edit10.text:=IniFile.ReadString('path','path-dir','');
  edit7.text:=IniFile.ReadString('bdBackup','path-dir','');
  Checkbox1.Checked:=IniFile.ReadBool('BackupCheck','path-dir',false);
  SpinEdit1.Value:=IniFile.ReadInteger('BackupTime','path-dir',1);


  LoadProSys();
  LoadForm();
  fexe:=1;




if ADOQuery10.RecordCount<>0 then
begin
    if ADOQuery13.RecordCount<>0 then
begin
    ADOConnection1.Connected:=false;
    ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection1.LoginPrompt:=false;
    ADOConnection1.Connected:=true;
    ADOQuery1.Connection:=ADOConnection1;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT DISTINCT(Група) FROM AllRecords');
    ADOQuery1.Active:=true;
    DataSource1.DataSet:=ADOQuery1;
    DBGrid1.DataSource:=DataSource1;
    DBGrid1.ReadOnly:=true;
end;

if ADOQuery13.RecordCount<>0 then
begin
    ADOConnection4.Connected:=false;
    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source='+table+'.mdb';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection4;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT DISTINCT(Викладач) FROM AllRecords');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
end;


end;





end;

procedure TForm1.Form1Destroy(Sender: TObject); //Закрытие файла при выходе из программы
begin
  excel:=Unassigned;
  IniFile.WriteString('bdBackup','path-dir',edit7.text);
  IniFile.WriteString('path','path-dir',edit10.text);
  IniFile.WriteBool('BackupCheck','path-dir',Checkbox1.Checked);
  IniFile.WriteInteger('BackupTime','path-dir',SpinEdit1.Value);


  IniFile.Free;   // Освобождения обьекта
end;

procedure TForm1.ListBox1DblClick(Sender: TObject);
begin
  ListBox1.Items.Delete(ListBox1.ItemIndex);
end;

procedure TForm1.ListBox2DblClick(Sender: TObject);
begin
 ListBox2.Items.Delete(ListBox2.ItemIndex);
end;

procedure TForm1.N1Click(Sender: TObject);
begin
  PageControl1.Visible:=false;
  Button17.Visible:=true;
  Button18.Visible:=true;
  Button19.Visible:=true;
  Button20.Visible:=true;
  Button21.Visible:=true;
  Button22.Visible:=true;
end;

procedure TForm1.RadioGroup1Click(Sender: TObject);
begin
     DateTimePicker5.Enabled:=true;
   DBLookupComboBox1.Enabled:=true;

     if radiogroup1.ItemIndex=0 then     // Заміна
  begin
     DBLookupComboBox2.Enabled:=true;
   DBLookupComboBox3.Enabled:=true;
   DBLookupComboBox4.Enabled:=true;
   DBLookupComboBox5.Enabled:=true;
   comboBox2.Enabled:=true;
   comboBox3.Enabled:=true;
   edit11.Enabled:=true;
   edit12.Enabled:=true;
   edit13.Enabled:=true;
   {Замена
 - день
-група
- № ленты (2)
- кем
- предмет
- кого
- предмет
- аудит
-коммент
}

      edit12.Enabled:=false;
    //
  end;

    if radiogroup1.ItemIndex=1 then    //Предмет-предметом
  begin
     DBLookupComboBox2.Enabled:=true;
   DBLookupComboBox3.Enabled:=true;
   DBLookupComboBox4.Enabled:=true;
   DBLookupComboBox5.Enabled:=true;
   comboBox2.Enabled:=true;
   comboBox3.Enabled:=true;
   edit11.Enabled:=true;
   edit12.Enabled:=true;
   edit13.Enabled:=true;

   {Замена предмет-предиетом
- дата
- группа
- лента
- кто
- предмет
- вместо кого
- предмет
- аудитория
- комментарий}
  comboBox3.Enabled:=false;
  edit12.Enabled:=false;
    //
  end;

    if radiogroup1.ItemIndex=2 then   // Погодинно
  begin
     DBLookupComboBox2.Enabled:=true;
   DBLookupComboBox3.Enabled:=true;
   DBLookupComboBox4.Enabled:=true;
   DBLookupComboBox5.Enabled:=true;
   comboBox2.Enabled:=true;
   comboBox3.Enabled:=true;
   edit11.Enabled:=true;
   edit12.Enabled:=true;
   edit13.Enabled:=true;
   {Замена погодинно
- дата
- группа
- лента
- кто
- предмет
- вместо кого
- предмет
- аудитория
- № приказа
- комментарий}
   comboBox3.Enabled:=false;
    //
  end;

    if radiogroup1.ItemIndex=3 then    // Заміна відпрцюванням
  begin
    {
    Замена отработкой:
- дата
- группа
- № ленты
- ким зам
- предмет
- аудитория
- комментарий}
   DBLookupComboBox2.Enabled:=true;
   DBLookupComboBox3.Enabled:=true;
   DBLookupComboBox4.Enabled:=true;
   DBLookupComboBox5.Enabled:=true;
   comboBox2.Enabled:=true;
   comboBox3.Enabled:=true;
   edit11.Enabled:=true;
   edit12.Enabled:=true;
   edit13.Enabled:=true;




  comboBox3.Enabled:=false;
  DBLookupComboBox4.Enabled:=false;
  DBLookupComboBox5.Enabled:=false;
  edit12.Enabled:=false;


  end;
end;

procedure TForm1.StatusBar1Hint(Sender: TObject);
begin
  StatusBar1.Panels[2].Text := Application.Hint;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
    StatusBar1.Panels[0].text:='Дата: '+DateToStr(now);
  StatusBar1.Panels[1].text:='Время: '+TimeToStr(now);
end;

procedure TForm1.Timer2Timer(Sender: TObject);
begin
  if CheckBox1.Checked=true then
  begin
    timer2.Interval:= SpinEdit1.Value*3600000;
    BackupSave();
  end;


end;

end.
