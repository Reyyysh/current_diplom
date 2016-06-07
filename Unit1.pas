unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.ExtCtrls, Vcl.DBCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls, Data.Win.ADODB, ADOX_TLB, Vcl.Mask,
  Vcl.ToolWin, Vcl.ComCtrls, Vcl.Menus, ComObj, Vcl.Buttons;

type
  TForm1 = class(TForm)
    ADOConnection1: TADOConnection;
    DataSource1: TDataSource;
    ADOQuery1: TADOQuery;
    ADOConnection2: TADOConnection;
    DataSource2: TDataSource;
    ADOQuery2: TADOQuery;
    ADOConnection3: TADOConnection;
    DataSource3: TDataSource;
    ADOQuery3: TADOQuery;
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
    Button2: TButton;
    Button3: TButton;
    Button8: TButton;
    Panel1: TPanel;
    DBNavigator2: TDBNavigator;
    DBGrid2: TDBGrid;
    DBNavigator1: TDBNavigator;
    Button5: TButton;
    Edit1: TEdit;
    Button7: TButton;
    DBGrid3: TDBGrid;
    DBGrid1: TDBGrid;
    Edit6: TEdit;
    TabSheet2: TTabSheet;
    Label11: TLabel;
    Edit2: TEdit;
    DBGrid4: TDBGrid;
    Button4: TButton;
    PageControl2: TPageControl;
    TabSheet6: TTabSheet;
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
    Edit8: TEdit;
    TabSheet7: TTabSheet;
    TabSheet8: TTabSheet;
    ComboBox4: TComboBox;
    DateTimePicker6: TDateTimePicker;
    DBLookupComboBox6: TDBLookupComboBox;
    DBLookupComboBox7: TDBLookupComboBox;
    Edit7: TEdit;
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
    Button13: TButton;
    Button14: TButton;
    Button15: TButton;
    Edit9: TEdit;
    StatusBar1: TStatusBar;
    Timer1: TTimer;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N7: TMenuItem;
    N8: TMenuItem;
    N9: TMenuItem;
    N10: TMenuItem;
    Button17: TButton;
    Button18: TButton;
    Button19: TButton;
    Button20: TButton;
    Button21: TButton;
    Button22: TButton;
    TabSheet9: TTabSheet;
    Label21: TLabel;
    DateTimePicker7: TDateTimePicker;
    DateTimePicker8: TDateTimePicker;
    Button23: TButton;
    Button24: TButton;
    DBGrid10: TDBGrid;
    Label22: TLabel;
    Label23: TLabel;
    RadioButton3: TRadioButton;
    RadioButton4: TRadioButton;
    Button25: TButton;
    procedure Button5Click(Sender: TObject);
    procedure Form1Destroy(Sender: TObject);
    procedure Form1Create(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
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
    procedure DBGrid1Width();
    procedure DBGrid2Width();
    procedure DBGrid3Width();
    procedure DBGrid4Width();
    procedure DBGrid5Width();
    procedure DBGrid6Width();
    procedure DBGrid7Width();
    procedure DBGrid8Width();
    procedure DBGrid9Width();


  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  DB:Catalog;
  Tables:Table;
  Columns:Column;
  excel: variant;
implementation

{$R *.dfm}

procedure TForm1.DBGrid1Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid1.Columns[0].Width := 0; //ширина полей
    DBGrid1.Columns[1].Width := 80; //ширина полей
end;

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

procedure TForm1.DBGrid4Width();   //Функция для установки ширини полей таблицы
begin
    DBGrid4.Columns[0].Width := 0; //ширина полей
    DBGrid4.Columns[1].Width := 80; //ширина полей
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
  ADOQuery1.FindFirst;
  edit9.Text:= DBGrid8.Fields[0].Value;

   for i:=0 to kol-1 do
   begin
    ADOQuery1.FindFirst;
    DBGrid8.Fields[0].Value;
 // ListBox1.Items.add(DBGrid1.);


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
  Edit8.text:=DBLookupComboBox1.KeyField;


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
  if not OpenDialog1.Execute then exit;
  excel.WorkBooks.Open(OpenDialog1.FileName);
  excel.Visible:=true;
end;

procedure TForm1.Button23Click(Sender: TObject);
begin
    DB:=CoCatalog.Create;
    DB.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');


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



    ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection1.LoginPrompt:=false;
    ADOConnection1.Connected:=true;
    ADOQuery1.Connection:=ADOConnection1;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT * FROM grp');
    ADOQuery1.Active:=true;
    DataSource1.DataSet:=ADOQuery1;
    DBGrid1.DataSource:=DataSource1;
    DBNavigator1.DataSource:=DataSource1;
    DBGrid1.ReadOnly:=true;

    DBGrid1Width();



    DBGrid6.DataSource:=DataSource1;    //Запрет на редактирование
    DBGrid6.ReadOnly:=true;

    DBGrid8.DataSource:=DataSource1;   //Запрет на редактирование
    DBGrid8.ReadOnly:=true;

    /////////////////////////////////////


    //Создание таблицы для преподавателей
    ADOConnection1.Connected:=false;
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


    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection1;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT * FROM thr');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
    DBGrid4.ReadOnly:=true;


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
    Tables.Columns.Append('№_ленты',adVarWChar,255);
    Tables.Columns.Append('Ким_заміняємо',adVarWChar,255);
    Tables.Columns.Append('Предмет1',adVarWChar,255);
    Tables.Columns.Append('Кого_заміняємо',adVarWChar,255);
    Tables.Columns.Append('Предмет2',adVarWChar,255);
    Tables.Columns.Append('Аудиторія',adVarWChar,255);


    ADOConnection5.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
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


  end;



procedure TForm1.Button2Click(Sender: TObject); //Кнопка добавления расписания группы
var
  y,x1,x2,add:string;
  nom,i:integer;
begin
  y:=excel.Selection.Address;

  for i:=0 to Length(y) do  // Удаление $ из диапазона
  begin
    Delete(y,Pos('$',y),1);
  end;

  x1:=Copy(y,2,Pos(':',y)-2);   // Удаление лишних знаков из диапазона
  x2:=Copy(y,Pos(':',y)+2,Length(y));

  nom:=strtoint(x2)-strtoint(x1);  // Количество записей
  for i:=0 to nom  do
  begin
    add:=inttostr(strtoint(x1)+i);  // Номер елемента
    ADOQuery2.Insert;
    ADOQuery2.FieldByName('План_Б').Value :=excel.WorkBooks[1].WorkSheets[1].Range['C'+add].Value;
    ADOQuery2.FieldByName('План_К').Value :=excel.WorkBooks[1].WorkSheets[1].Range['D'+add].Value;
    ADOQuery2.FieldByName('Коментар').Value :=excel.WorkBooks[1].WorkSheets[1].Range['E'+add].Value;
    ADOQuery2.FieldByName('Група').Value :=excel.WorkBooks[1].WorkSheets[1].Range['F'+add].Value;
    ADOQuery2.FieldByName('Дисципліна').Value :=excel.WorkBooks[1].WorkSheets[1].Range['G'+add].Value;
    ADOQuery2.FieldByName('Викладач').Value :=excel.WorkBooks[1].WorkSheets[1].Range['H'+add].Value;
    ADOQuery2.FieldByName('Семестр').Value :='-';
    ADOQuery2.Post;
  end;
end;

procedure TForm1.Button3Click(Sender: TObject); //Кнопка добавления расписания по числителю/знаменателю
var
  y,x1,x2,add:string;
  nom,i:integer;
begin
  y:=excel.Selection.Address;

  for i:=0 to Length(y) do  // Удаление $ из диапазона
  begin
    Delete(y,Pos('$',y),1);
  end;

  x1:=Copy(y,3,Pos(':',y)-3);   // Удаление лишних знаков из диапазона
  x2:=Copy(y,Pos(':',y)+3,Length(y));

  nom:=strtoint(x2)-strtoint(x1);  // Количество записей
  for i:=0 to nom  do
  begin
    add:=inttostr(strtoint(x1)+i);  // Номер елемента
    ADOQuery3.Insert;
    ADOQuery3.FieldByName('ПнЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['MZ'+add].Value;
    ADOQuery3.FieldByName('ПнЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NA'+add].Value;
    ADOQuery3.FieldByName('ВтЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NB'+add].Value;
    ADOQuery3.FieldByName('ВтЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NC'+add].Value;
    ADOQuery3.FieldByName('СрЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['ND'+add].Value;
    ADOQuery3.FieldByName('СрЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NE'+add].Value;
    ADOQuery3.FieldByName('ЧтЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NF'+add].Value;
    ADOQuery3.FieldByName('ЧтЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NG'+add].Value;
    ADOQuery3.FieldByName('ПтЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NH'+add].Value;
    ADOQuery3.FieldByName('ПтЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NI'+add].Value;
    ADOQuery3.FieldByName('СбЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NJ'+add].Value;
    ADOQuery3.FieldByName('СбЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NK'+add].Value;
    ADOQuery3.FieldByName('НдЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NL'+add].Value;
    ADOQuery3.FieldByName('НдЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NM'+add].Value;
    ADOQuery3.FieldByName('ПнЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NN'+add].Value;
    ADOQuery3.FieldByName('ПнЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NO'+add].Value;
    ADOQuery3.FieldByName('ВтЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NP'+add].Value;
    ADOQuery3.FieldByName('ВтЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NQ'+add].Value;
    ADOQuery3.FieldByName('СрЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NR'+add].Value;
    ADOQuery3.FieldByName('СрЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NS'+add].Value;
    ADOQuery3.FieldByName('ЧтЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NT'+add].Value;
    ADOQuery3.FieldByName('ЧтЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NU'+add].Value;
    ADOQuery3.FieldByName('ПтЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NV'+add].Value;
    ADOQuery3.FieldByName('ПтЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NW'+add].Value;
    ADOQuery3.Post;
  end;
end;

procedure TForm1.Button4Click(Sender: TObject); //Кнопка добавления списка преподавателей
var
  y,x1,x2,add:string;
  nom,i:integer;
begin
  y:=excel.Selection.Address;

  for i:=0 to Length(y) do  // Удаление $ из диапазона
  begin
    Delete(y,Pos('$',y),1);
  end;

  x1:=Copy(y,2,Pos(':',y)-2);   // Удаление лишних знаков из диапазона
  x2:=Copy(y,Pos(':',y)+2,Length(y));

  nom:=strtoint(x2)-strtoint(x1);  // Количество записей
  for i:=0 to nom  do
  begin
    add:=inttostr(strtoint(x1)+i);  // Номер елемента
    ADOQuery4.Insert;
    ADOQuery4.FieldByName('IDВикладача').Value :=excel.WorkBooks[1].WorkSheets[4].Range['A'+add].Value;
    ADOQuery4.FieldByName('Викладач').Value :=excel.WorkBooks[1].WorkSheets[4].Range['A'+add].Value;
    ADOQuery4.Post;
  end;
end;

procedure TForm1.Button5Click(Sender: TObject); //Добавление группы в список
var
  i:integer;
  rsp,test,graf,subb:string;
begin
  test:=Edit1.Text;
  for i:=0 to Length(test) do  // Удаление - из диапазона
  begin
    Delete(test,Pos('-',test),1);
  end;
  rsp:=test+'rsp';
  graf:=test+'graf';
  subb:=test+'subb';


  ADOConnection2.Connected:=false;
  ADOConnection3.Connected:=false;

  //Создание таблицы групп
  Tables:=CoTable.Create;
  Tables.Name:=test;

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
  Tables.Columns.Append('План_Б',adInteger,0);
  Tables.Columns.Append('План_К',adInteger,0);
  Tables.Columns.Append('Коментар',adVarWChar,255);
  Tables.Columns.Append('Група',adVarWChar,255);
  Tables.Columns.Append('Дисципліна',adVarWChar,255);
  Tables.Columns.Append('Викладач',adVarWChar,255);
  Tables.Columns.Append('Семестр',adVarWChar,255);

  ADOConnection2.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
  ADOConnection2.LoginPrompt:=false;
  ADOConnection2.Connected:=true;
  ADOQuery2.Connection:=ADOConnection2;
  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add('SELECT * FROM '+test);
  ADOQuery2.Active:=true;
  DataSource2.DataSet:=ADOQuery2;
  DBGrid2.DataSource:=DataSource2;
  DBNavigator2.DataSource:=DataSource2;

  //Создание таблицы расписания
  Tables:=CoTable.Create;
  Tables.Name:=rsp;

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

  ADOConnection3.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
  ADOConnection3.LoginPrompt:=false;
  ADOConnection3.Connected:=true;
  ADOQuery3.Connection:=ADOConnection3;
  ADOQuery3.SQL.Clear;
  ADOQuery3.SQL.Add('SELECT * FROM '+rsp);
  ADOQuery3.Active:=true;
  DataSource3.DataSet:=ADOQuery3;
  DBGrid3.DataSource:=DataSource3;
  //DBNavigator1.DataSource:=DataSource1;


  // Создание таблицы графика

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

  ADOConnection7.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
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
  Tables.Columns.Append('Дата',adVarWChar,255);
  Tables.Columns.Append('Комент',adVarWChar,255);

  ADOConnection9.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
  ADOConnection9.LoginPrompt:=false;
  ADOConnection9.Connected:=true;
  ADOQuery9.Connection:=ADOConnection9;
  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM '+subb);
  ADOQuery9.Active:=true;
  DataSource9.DataSet:=ADOQuery9;
  DBGrid9.DataSource:=DataSource9;
  DBNavigator9.DataSource:=DataSource9;


  //////////////////////////////////////////////////////////////////



  ADOQuery1.Insert;
  ADOQuery1.FieldByName('Група').AsString := Edit1.Text;
  ADOQuery1.FieldByName('IDГруп').AsString := test;
  ADOQuery1.Post;

  DBGrid2Width();
  DBGrid3Width();
  DBGrid7Width();
  DBGrid9Width();


  Edit1.Text:='';
end;

procedure TForm1.Button6Click(Sender: TObject); //Кнопка поиска преподавателей
begin
  ListBox2.Clear;
end;

procedure TForm1.Button7Click(Sender: TObject);
begin
  ADOConnection2.Connected:=true;
  ADOQuery2.Close;
  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add('DROP Table '+DBGrid1.Fields[0].Value);
  ADOQuery2.ExecSQL;;

  ADOConnection3.Connected:=true;
  ADOQuery3.Close;
  ADOQuery3.SQL.Clear;
  ADOQuery3.SQL.Add('DROP Table '+DBGrid1.Fields[0].Value+'rsp');
  ADOQuery3.ExecSQL;

  DBGrid1.DataSource.DataSet.Delete;


  //Таблица с группами из вкладки график
  //DBGrid6.DataSource.DataSet.Delete;
end;

procedure TForm1.Button8Click(Sender: TObject);
var
  rsp,test,graf,subb,zam:string;
  y,x1,x2,add:string;
  nom,i:integer;
  z:integer;
begin

 y:=excel.Selection.Address;
 // excel.WorkBooks[1].WorkSheets[1].Range['C14:H28'].Select;
 // y:='$C$14:$H$28';

  //Нельзя выделять диапазоны с двойными цифрами и одинарными буквами

  for i:=0 to Length(y) do  // Удаление $ из диапазона
  begin
  Delete(y,Pos('$',y),1);
  end;

  x1:=Copy(y,2,Pos(':',y)-2);   // Удаление лишних знаков из диапазона
  x2:=Copy(y,Pos(':',y)+2,Length(y));


 test:=excel.WorkBooks[1].WorkSheets[1].Range['G'+x1].Value;



  for i:=0 to Length(test) do  // Удаление - из диапазона
  begin
    Delete(test,Pos('-',test),1);
  end;
  rsp:=test+'rsp';
  graf:=test+'graf';
  subb:=test+'subb';

  ADOConnection2.Connected:=false;
  ADOConnection3.Connected:=false;

  //Создание таблицы групп
  Tables:=CoTable.Create;
  Tables.Name:=test;

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
  Tables.Columns.Append('План_Б',adInteger,0);
  Tables.Columns.Append('План_К',adInteger,0);
  Tables.Columns.Append('Коментар',adVarWChar,255);
  Tables.Columns.Append('Група',adVarWChar,255);
  Tables.Columns.Append('Дисципліна',adVarWChar,255);
  Tables.Columns.Append('Викладач',adVarWChar,255);
  Tables.Columns.Append('Семестр',adVarWChar,255);

  ADOConnection2.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
  ADOConnection2.LoginPrompt:=false;
  ADOConnection2.Connected:=true;
  ADOQuery2.Connection:=ADOConnection2;
  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add('SELECT * FROM '+test);
  ADOQuery2.Active:=true;
  DataSource2.DataSet:=ADOQuery2;
  DBGrid2.DataSource:=DataSource2;
  DBNavigator2.DataSource:=DataSource2;

  //Создание таблицы расписания
  Tables:=CoTable.Create;
  Tables.Name:=rsp;

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

  ADOConnection3.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
  ADOConnection3.LoginPrompt:=false;
  ADOConnection3.Connected:=true;
  ADOQuery3.Connection:=ADOConnection3;
  ADOQuery3.SQL.Clear;
  ADOQuery3.SQL.Add('SELECT * FROM '+rsp);
  ADOQuery3.Active:=true;
  DataSource3.DataSet:=ADOQuery3;
  DBGrid3.DataSource:=DataSource3;
  //DBNavigator1.DataSource:=DataSource1;

  // Создание таблицы графика

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

  ADOConnection7.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
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




  ADOConnection9.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
  ADOConnection9.LoginPrompt:=false;
  ADOConnection9.Connected:=true;
  ADOQuery9.Connection:=ADOConnection9;
  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM '+subb);
  ADOQuery9.Active:=true;
  DataSource9.DataSet:=ADOQuery9;
  DBGrid9.DataSource:=DataSource9;
  DBNavigator9.DataSource:=DataSource9;



  //////////////////////////////////////////////////////////////////



 


  ADOQuery1.Insert;
  ADOQuery1.FieldByName('Група').AsString := excel.WorkBooks[1].WorkSheets[1].Range['G'+x1].Value;
  ADOQuery1.FieldByName('IDГруп').AsString := test;
  ADOQuery1.Post;




  Edit1.Text:='';



   x1:=inttostr(strtoint(x1)+3);


  nom:=strtoint(x2)-strtoint(x1);  // Количество записей
    for i:=0 to nom  do
    begin
    add:=inttostr(strtoint(x1)+i);  // Номер елемента
  ADOQuery2.Insert;
  ADOQuery2.FieldByName('План_Б').Value :=excel.WorkBooks[1].WorkSheets[1].Range['C'+add].Value;
  ADOQuery2.FieldByName('План_К').Value :=excel.WorkBooks[1].WorkSheets[1].Range['D'+add].Value;
  ADOQuery2.FieldByName('Коментар').Value :=excel.WorkBooks[1].WorkSheets[1].Range['E'+add].Value;
  ADOQuery2.FieldByName('Група').Value :=excel.WorkBooks[1].WorkSheets[1].Range['F'+add].Value;
  ADOQuery2.FieldByName('Дисципліна').Value :=excel.WorkBooks[1].WorkSheets[1].Range['G'+add].Value;
  ADOQuery2.FieldByName('Викладач').Value :=excel.WorkBooks[1].WorkSheets[1].Range['H'+add].Value;
  ADOQuery2.FieldByName('Семестр').Value :='-';
  ADOQuery2.Post;

    end;

     DBGrid2Width();
     DBGrid3Width();
     DBGrid7Width();
     DBGrid9Width();

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

procedure TForm1.DBGrid1CellClick(Column: TColumn); //Выбор группы из списка для просмотра расписания
var
  i: integer;
begin
  DBGrid1.DataSource.DataSet.RecNo;
  DBGrid1.Fields[0].Value;

  Panel1.Caption:='Група '+DBGrid1.Fields[1].Value;

  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add('SELECT * FROM '+DBGrid1.Fields[0].Value);
  ADOQuery2.Active:=true;

  ADOQuery3.SQL.Clear;
  ADOQuery3.SQL.Add('SELECT * FROM '+DBGrid1.Fields[0].Value+'rsp');
  ADOQuery3.Active:=true;

  DBGrid2Width();   // Функция ширина рядков таблицы
  DBGrid3Width();   // Функция ширина рядков таблицы

end;

procedure TForm1.DBGrid6CellClick(Column: TColumn);
var
  i: integer;
begin
  DBGrid6.DataSource.DataSet.RecNo;
  DBGrid6.Fields[0].Value;

  Panel2.Caption:='Група '+DBGrid6.Fields[1].Value;

  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('SELECT * FROM '+DBGrid6.Fields[0].Value+'graf');
  ADOQuery7.Active:=true;


  DBGrid6Width();


end;

procedure TForm1.DBGrid6DblClick(Sender: TObject);
var
  i,check:integer;
begin
//ListBox1.Items.i:=DBGrid8.Fields[0].Value  ;

      if Listbox2.Count=2 then
    exit;

  if ListBox2.Items.Count=0 then
  ListBox2.Items.add(DBGrid6.Fields[1].Value);


   check:=0;
  for i:=0 to ListBox2.Items.Count-1 do
     begin
   if DBGrid6.Fields[1].Value=ListBox2.Items[i]  then
   begin
    check:=1;
    break;


   end
   else
    check:=0;


     end;

    if check=0 then  ListBox2.Items.add(DBGrid6.Fields[1].Value);





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
begin

  Panel3.Caption:='Група '+DBGrid8.Fields[1].Value;

  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM '+DBGrid8.Fields[0].Value+'subb');
  ADOQuery9.Active:=true;


  DBGrid9Width();


end;

procedure TForm1.DBGrid8DblClick(Sender: TObject);
var
  i,check:integer;
begin
//ListBox1.Items.i:=DBGrid8.Fields[0].Value  ;

if ListBox1.Items.Count=0 then
ListBox1.Items.add(DBGrid8.Fields[1].Value);


   check:=0;
  for i:=0 to ListBox1.Items.Count-1 do
     begin
   if DBGrid8.Fields[1].Value=ListBox1.Items[i]  then
   begin
    check:=1;
    break;


   end
   else
    check:=0;


     end;

    if check=0 then  ListBox1.Items.add(DBGrid8.Fields[1].Value);



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
  if FileExists('test.mdb') then
  begin
    DB:=CoCatalog.Create;
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');





    //Список групп
    ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection1.LoginPrompt:=false;
    ADOConnection1.Connected:=true;
    ADOQuery1.Connection:=ADOConnection1;
    ADOQuery1.SQL.Clear;
    ADOQuery1.SQL.Add('SELECT * FROM grp');
    ADOQuery1.Active:=true;
    DataSource1.DataSet:=ADOQuery1;


    DBGrid1.DataSource:=DataSource1;
    DBGrid1.ReadOnly:=true;

    DBGrid6.DataSource:=DataSource1;
    DBGrid6.ReadOnly:=true;

    DBGrid8.DataSource:=DataSource1;
    DBGrid8.ReadOnly:=true;





    DBNavigator1.DataSource:=DataSource1;






    //Расписание группы
    Panel1.Caption:='Група '+DBGrid1.Fields[1].Value;

    ADOConnection2.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection2.LoginPrompt:=false;
    ADOConnection2.Connected:=true;
    ADOQuery2.Connection:=ADOConnection2;
    ADOQuery2.SQL.Clear;
    ADOQuery2.SQL.Add('SELECT * FROM '+DBGrid1.Fields[0].Value);
    ADOQuery2.Active:=true;
    DataSource2.DataSet:=ADOQuery2;
    DBGrid2.DataSource:=DataSource2;
    DBNavigator2.DataSource:=DataSource2;




    //Расписание числитель/знаменатель
    ADOConnection3.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection3.LoginPrompt:=false;
    ADOConnection3.Connected:=true;
    ADOQuery3.Connection:=ADOConnection3;
    ADOQuery3.SQL.Clear;
    rsp:=DBGrid1.Fields[0].Value+'rsp';
    ADOQuery3.SQL.Add('SELECT * FROM '+rsp);
    ADOQuery3.Active:=true;
    DataSource3.DataSet:=ADOQuery3;
    DBGrid3.DataSource:=DataSource3;





    //Список преподавателей
    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection1;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT * FROM thr');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
    DBGrid4.ReadOnly:=true;
    DBGrid4.DataSource.DataSet.RecNo;



    ADOConnection7.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection7.LoginPrompt:=false;
    ADOConnection7.Connected:=true;
    ADOQuery7.Connection:=ADOConnection7;
    ADOQuery7.SQL.Clear;
    ADOQuery7.SQL.Add('SELECT * FROM '+DBGrid6.Fields[0].Value+'graf');
    ADOQuery7.Active:=true;
    DataSource7.DataSet:=ADOQuery7;
    DBGrid7.DataSource:=DataSource7;
    DBNavigator7.DataSource:=DataSource7;

/////////////////////////////



    ADOConnection9.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection9.LoginPrompt:=false;
    ADOConnection9.Connected:=true;
    ADOQuery9.Connection:=ADOConnection9;
    ADOQuery9.SQL.Clear;
    ADOQuery9.SQL.Add('SELECT * FROM '+DBGrid8.Fields[0].Value+'subb');
    ADOQuery9.Active:=true;
    DataSource9.DataSet:=ADOQuery9;
    DBGrid9.DataSource:=DataSource9;
    DBNavigator9.DataSource:=DataSource9;


  ////////////////////////////////////////////////
     ADOConnection5.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection5.LoginPrompt:=false;
    ADOConnection5.Connected:=true;
    ADOQuery5.Connection:=ADOConnection5;
    ADOQuery5.SQL.Clear;
    ADOQuery5.SQL.Add('SELECT * FROM zam');
    ADOQuery5.Active:=true;
    DataSource5.DataSet:=ADOQuery5;
    DBGrid5.DataSource:=DataSource5;
    DBNavigator5.DataSource:=DataSource5;




  DBLookupComboBox1.DataField:= 'Група';
  DBLookupComboBox1.DataSource:= DataSource5;
  DBLookupComboBox1.KeyField:= 'Група';
  DBLookupComboBox1.ListField:= 'Група';
  DBLookupComboBox1.ListSource:= DataSource1;
 // DBLookupComboBox1.TabOrder:= 4;


  DBLookupComboBox2.DataField:= 'Ким_заміняємо';
  DBLookupComboBox2.DataSource:= DataSource5;
  DBLookupComboBox2.KeyField:= 'Викладач';
  DBLookupComboBox2.ListField:= 'Викладач';
  DBLookupComboBox2.ListSource:= DataSource4;


    DBLookupComboBox4.DataField:= 'Кого_заміняємо';
  DBLookupComboBox4.DataSource:= DataSource5;
  DBLookupComboBox4.KeyField:= 'Викладач';
  DBLookupComboBox4.ListField:= 'Викладач';
  DBLookupComboBox4.ListSource:= DataSource4;


    DBGrid1Width();
    DBGrid2Width();
    DBGrid3Width();
    DBGrid4Width();
    DBGrid5Width();
    DBGrid7Width();
    DBGrid9Width();
  end
  else
  begin

  end;




end;

procedure TForm1.Form1Destroy(Sender: TObject); //Закрытие файла при выходе из программы
begin
  excel:=Unassigned;
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

procedure TForm1.StatusBar1Hint(Sender: TObject);
begin
  StatusBar1.Panels[2].Text := Application.Hint;
end;

procedure TForm1.Timer1Timer(Sender: TObject);
begin
    StatusBar1.Panels[0].text:='Дата: '+DateToStr(now);
  StatusBar1.Panels[1].text:='Время: '+TimeToStr(now);
end;

end.
