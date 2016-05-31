unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.ExtCtrls, Vcl.DBCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls, Data.Win.ADODB, ADOX_TLB, Vcl.Mask,
  Vcl.ToolWin, Vcl.ComCtrls, Vcl.Menus, ComObj, Unit2;

type
  TForm1 = class(TForm)
    DBNavigator1: TDBNavigator;
    Button2: TButton;
    ADOConnection1: TADOConnection;
    Edit1: TEdit;
    DBGrid2: TDBGrid;
    ADOConnection2: TADOConnection;
    DataSource2: TDataSource;
    DataSource1: TDataSource;
    ADOQuery1: TADOQuery;
    ADOQuery2: TADOQuery;
    DBNavigator2: TDBNavigator;
    Edit2: TEdit;
    Edit3: TEdit;
    Button5: TButton;
    Edit4: TEdit;
    Button6: TButton;
    ADOConnection3: TADOConnection;
    DataSource3: TDataSource;
    ADOQuery3: TADOQuery;
    DBGrid3: TDBGrid;
    DBNavigator3: TDBNavigator;
    DBGrid4: TDBGrid;
    DBNavigator4: TDBNavigator;
    ADOConnection4: TADOConnection;
    DataSource4: TDataSource;
    ADOQuery4: TADOQuery;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    Label2: TLabel;
    Label3: TLabel;
    Label4: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    DBLookupComboBox1: TDBLookupComboBox;
    DBLookupComboBox2: TDBLookupComboBox;
    DBLookupComboBox3: TDBLookupComboBox;
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    N1: TMenuItem;
    N2: TMenuItem;
    N3: TMenuItem;
    Timer1: TTimer;
    Panel1: TPanel;
    Button1: TButton;
    OpenDialog1: TOpenDialog;
    Button3: TButton;
    DBGrid1: TDBGrid;
    Panel2: TPanel;
    Panel3: TPanel;
    GroupBox1: TGroupBox;
    Label1: TLabel;
    Label9: TLabel;
    Label10: TLabel;
    Label11: TLabel;
    DBGrid5: TDBGrid;
    ADOConnection5: TADOConnection;
    DataSource5: TDataSource;
    ADOQuery5: TADOQuery;
    Button4: TButton;
    Button7: TButton;
    DateTimePicker1: TDateTimePicker;
    Edit5: TEdit;
    Splitter1: TSplitter;
    Panel4: TPanel;
    Button8: TButton;
    Button9: TButton;
    Edit6: TEdit;
    Button10: TButton;
    procedure FormCreate(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure DBGrid2CellClick(Column: TColumn);
    procedure Button5Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure StatusBar1Hint(Sender: TObject);
    procedure Timer1Timer(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
    procedure Button9Click(Sender: TObject);
    procedure Button10Click(Sender: TObject);






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
  excel: variant; // Переменная в которой создаётся объект EXCEL

implementation

{$R *.dfm}

//dataset.tables["blablabla"].rows[i]["ID"]
//DBGrid2.SelectedField.AsString

procedure TForm1.Button10Click(Sender: TObject);
begin

     if not ADOQuery1.Locate('Викладач',Edit6.Text,[loCaseInsensitive, loPartialKey])then
    ShowMessage('Запись не найдена');

end;

procedure TForm1.Button1Click(Sender: TObject);
begin
  excel:=CreateOleObject('Excel.Application');
  if not OpenDialog1.Execute then exit;
  excel.WorkBooks.Open(OpenDialog1.FileName);
    excel.Visible:=true;
  //  Panel2.Visible:=False;
   // Application.MainFormOnTaskbar := True;
   // Form1.Visible:=False;
end;


procedure TForm1.Button2Click(Sender: TObject);
var
i:integer;
rsp:string;
begin
    rsp:=Edit1.Text+'rsp';

    ADOConnection1.Connected:=false;
    ADOConnection5.Connected:=false;

    //Создание таблицы групп
   Tables:=CoTable.Create;
   Tables.Name:=Edit1.Text;

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
   Tables.Columns.Append('План Б',adInteger,0);
   Tables.Columns.Append('План К',adInteger,0);
   Tables.Columns.Append('Коментар',adVarWChar,255);
   Tables.Columns.Append('Група',adVarWChar,255);
   Tables.Columns.Append('Дисципліна',adVarWChar,255);
   Tables.Columns.Append('Викладач',adVarWChar,255);


   ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection1.LoginPrompt:=false;
   ADOConnection1.Connected:=true;
   ADOQuery1.Connection:=ADOConnection1;
   ADOQuery1.SQL.Clear;
   ADOQuery1.SQL.Add('SELECT * FROM '+Edit1.Text);
   ADOQuery1.Active:=true;
   DataSource1.DataSet:=ADOQuery1;
   DBGrid1.DataSource:=DataSource1;
   DBNavigator1.DataSource:=DataSource1;
   ////////////////////////////////////////////////////////


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


   ADOConnection5.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection5.LoginPrompt:=false;
   ADOConnection5.Connected:=true;
   ADOQuery5.Connection:=ADOConnection5;
   ADOQuery5.SQL.Clear;
   ADOQuery5.SQL.Add('SELECT * FROM '+rsp);
   ADOQuery5.Active:=true;
   DataSource5.DataSet:=ADOQuery5;
   DBGrid5.DataSource:=DataSource5;
   //DBNavigator1.DataSource:=DataSource1;

  ADOQuery2.Insert;
  ADOQuery2.FieldByName('Група').AsString := Edit2.Text;
  ADOQuery2.FieldByName('IDГруп').AsString := Edit1.Text;
  ADOQuery2.Post;

  DBGrid2.Columns[0].Width := 0; //ширина полей
  DBGrid2.Columns[1].Width := 100; //ширина полей




end;

procedure TForm1.Button3Click(Sender: TObject);
var
y,x1,x2,add:string;
nom,i:integer;
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


  nom:=strtoint(x2)-strtoint(x1);  // Количество записей
  for i:=0 to nom  do
    begin
    add:=inttostr(strtoint(x1)+i);  // Номер елемента
  ADOQuery1.Insert;
  ADOQuery1.FieldByName('План Б').Value :=excel.WorkBooks[1].WorkSheets[1].Range['C'+add].Value;
  ADOQuery1.FieldByName('План К').Value :=excel.WorkBooks[1].WorkSheets[1].Range['D'+add].Value;
  ADOQuery1.FieldByName('Коментар').Value :=excel.WorkBooks[1].WorkSheets[1].Range['E'+add].Value;
  ADOQuery1.FieldByName('Група').Value :=excel.WorkBooks[1].WorkSheets[1].Range['F'+add].Value;
  ADOQuery1.FieldByName('Дисципліна').Value :=excel.WorkBooks[1].WorkSheets[1].Range['G'+add].Value;
  ADOQuery1.FieldByName('Викладач').Value :=excel.WorkBooks[1].WorkSheets[1].Range['H'+add].Value;
  ADOQuery1.Post;

    end;

end;

procedure TForm1.Button4Click(Sender: TObject);
begin
 ADOConnection5.Connected:=false;


   Tables:=CoTable.Create;
   Tables.Name:='Test';

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





   ADOConnection5.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection5.LoginPrompt:=false;
   ADOConnection5.Connected:=true;
   ADOQuery5.Connection:=ADOConnection5;
   ADOQuery5.SQL.Clear;
   ADOQuery5.SQL.Add('SELECT * FROM Test');
   ADOQuery5.Active:=true;
   DataSource5.DataSet:=ADOQuery5;
   DBGrid5.DataSource:=DataSource5;
   //DBNavigator1.DataSource:=DataSource1;

end;

procedure TForm1.Button5Click(Sender: TObject);
begin
    ADOQuery3.Insert;
  ADOQuery3.FieldByName('Дисципліна').AsString := Edit3.Text;
  ADOQuery3.Post;
end;

procedure TForm1.Button6Click(Sender: TObject);
begin
      ADOQuery4.Insert;
  ADOQuery4.FieldByName('Викладач').AsString := Edit4.Text;
  ADOQuery4.Post;
end;

procedure TForm1.Button7Click(Sender: TObject);
var
Result:string;
i:integer;
begin
  case DayOfWeek(DateTimePicker1.Date) of
  1: Result := 'Воскресенье';
  2: Result := 'Понедельник';
  3: Result := 'Вторник';
  4: Result := 'Среда';
  5: Result := 'Четверг';
  6: Result := 'Пятница';
  7: Result := 'Суббота';
  end;
  Edit5.Text:=Result;


  excel:= CreateOleObject('Excel.Application');
 // if not OpenDialog1.Execute then exit;
  excel.WorkBooks.Add;          // Создание новой книги
  excel.WorkBooks[1].WorkSheets[1].Name := 'Розклад';     //Создание нового листа


   for i:=1 to 1000  do
    begin
      excel.WorkBooks[1].WorkSheets[1].Columns[i].ColumnWidth := 1;
    end;

    excel.WorkBooks[1].WorkSheets[1].Range['KL3:MX3'].Merge;
    excel.WorkBooks[1].WorkSheets[1].Range['KL3']:='Сетрябрь';

    excel.WorkBooks[1].WorkSheets[1].Range['KL4:KM4'].Merge;
    excel.WorkBooks[1].WorkSheets[1].Range['KL4']:='Чт';
    excel.WorkBooks[1].WorkSheets[1].Range['KL5:KM5'].Merge;
    excel.WorkBooks[1].WorkSheets[1].Range['KL5']:='1';







  //excel.WorkBooks.Open('Розклад.xls');
    excel.Visible:=true;
end;


procedure TForm1.Button8Click(Sender: TObject);
var
y,x1,x2,add:string;
nom,i:integer;
begin

  y:=excel.Selection.Address;
 // excel.WorkBooks[1].WorkSheets[1].Range['C14:H28'].Select;
 // y:='$C$14:$H$28';

  //Нельзя выделять диапазоны с двойными цифрами и одинарными буквами
  //Привязка есть к буквам

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
  ADOQuery5.Insert;
  {ADOQuery1.FieldByName('План Б').Value :=excel.WorkBooks[1].WorkSheets[1].Range['C'+add].Value;
  ADOQuery1.FieldByName('План К').Value :=excel.WorkBooks[1].WorkSheets[1].Range['D'+add].Value;
  ADOQuery1.FieldByName('Коментар').Value :=excel.WorkBooks[1].WorkSheets[1].Range['E'+add].Value;
  ADOQuery1.FieldByName('Група').Value :=excel.WorkBooks[1].WorkSheets[1].Range['F'+add].Value;
  ADOQuery1.FieldByName('Дисципліна').Value :=excel.WorkBooks[1].WorkSheets[1].Range['G'+add].Value;
  ADOQuery1.FieldByName('Викладач').Value :=excel.WorkBooks[1].WorkSheets[1].Range['H'+add].Value;}




  ADOQuery5.FieldByName('ПнЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['MZ'+add].Value;
  ADOQuery5.FieldByName('ПнЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NA'+add].Value;
  ADOQuery5.FieldByName('ВтЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NB'+add].Value;
  ADOQuery5.FieldByName('ВтЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NC'+add].Value;
  ADOQuery5.FieldByName('СрЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['ND'+add].Value;
  ADOQuery5.FieldByName('СрЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NE'+add].Value;
  ADOQuery5.FieldByName('ЧтЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NF'+add].Value;
  ADOQuery5.FieldByName('ЧтЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NG'+add].Value;
  ADOQuery5.FieldByName('ПтЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NH'+add].Value;
  ADOQuery5.FieldByName('ПтЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NI'+add].Value;
  ADOQuery5.FieldByName('СбЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NJ'+add].Value;
  ADOQuery5.FieldByName('СбЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NK'+add].Value;
  ADOQuery5.FieldByName('НдЧ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NL'+add].Value;
  ADOQuery5.FieldByName('НдЧ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NM'+add].Value;


  ADOQuery5.FieldByName('ПнЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NN'+add].Value;
  ADOQuery5.FieldByName('ПнЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NO'+add].Value;
  ADOQuery5.FieldByName('ВтЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NP'+add].Value;
  ADOQuery5.FieldByName('ВтЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NQ'+add].Value;
  ADOQuery5.FieldByName('СрЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NR'+add].Value;
  ADOQuery5.FieldByName('СрЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NS'+add].Value;
  ADOQuery5.FieldByName('ЧтЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NT'+add].Value;
  ADOQuery5.FieldByName('ЧтЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NU'+add].Value;
  ADOQuery5.FieldByName('ПтЗ1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NV'+add].Value;
  ADOQuery5.FieldByName('ПтЗ2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NW'+add].Value;
  ADOQuery5.Post;
    end;

end;

procedure TForm1.Button9Click(Sender: TObject);
var
y,x1,x2,add:string;
nom,i:integer;
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


  nom:=strtoint(x2)-strtoint(x1);  // Количество записей
  for i:=0 to nom  do
    begin
    add:=inttostr(strtoint(x1)+i);  // Номер елемента
  ADOQuery4.Insert;

  ADOQuery4.FieldByName('Викладач').Value :=excel.WorkBooks[1].WorkSheets[4].Range['A'+add].Value;
  ADOQuery4.Post;

    end;

end;

procedure TForm1.DBGrid2CellClick(Column: TColumn);
begin


    DBGrid2.DataSource.DataSet.RecNo;
    DBGrid2.Fields[0].Value;

   Panel3.Caption:='Номер групи: '+DBGrid2.Fields[1].Value;


   ADOQuery1.SQL.Clear;
   ADOQuery1.SQL.Add('SELECT * FROM '+DBGrid2.Fields[0].Value);
   ADOQuery1.Active:=true;

   ADOQuery5.SQL.Clear;
   ADOQuery5.SQL.Add('SELECT * FROM '+DBGrid2.Fields[0].Value+'rsp');
   ADOQuery5.Active:=true;


   DBGrid1.Columns[0].Width := 30; //ширина полей
   DBGrid1.Columns[1].Width := 40; //ширина полей
   DBGrid1.Columns[2].Width := 40; //ширина полей
  DBGrid1.Columns[3].Width := 80; //ширина полей
  DBGrid1.Columns[4].Width := 60; //ширина полей
  DBGrid1.Columns[5].Width := 200; //ширина полей
  DBGrid1.Columns[6].Width := 120; //ширина полей



   DBGrid5.Columns[0].Width := 15; //ширина полей
  DBGrid5.Columns[1].Width := 15; //ширина полей
  DBGrid5.Columns[2].Width := 15; //ширина полей
  DBGrid5.Columns[3].Width := 15; //ширина полей
  DBGrid5.Columns[4].Width := 15; //ширина полей
  DBGrid5.Columns[5].Width := 15; //ширина полей
  DBGrid5.Columns[6].Width := 15; //ширина полей
  DBGrid5.Columns[7].Width := 15; //ширина полей
  DBGrid5.Columns[8].Width := 15; //ширина полей
  DBGrid5.Columns[9].Width := 15; //ширина полей
  DBGrid5.Columns[10].Width := 15; //ширина полей
  DBGrid5.Columns[11].Width := 15; //ширина полей
  DBGrid5.Columns[12].Width := 15; //ширина полей
  DBGrid5.Columns[13].Width := 15; //ширина полей
  DBGrid5.Columns[14].Width := 15; //ширина полей
  DBGrid5.Columns[15].Width := 15; //ширина полей
  DBGrid5.Columns[16].Width := 15; //ширина полей
  DBGrid5.Columns[17].Width := 15; //ширина полей
  DBGrid5.Columns[18].Width := 15; //ширина полей
  DBGrid5.Columns[19].Width := 15; //ширина полей
  DBGrid5.Columns[20].Width := 15; //ширина полей
  DBGrid5.Columns[21].Width := 15; //ширина полей
  DBGrid5.Columns[22].Width := 15; //ширина полей
  DBGrid5.Columns[23].Width := 15; //ширина полей

end;

procedure TForm1.FormCreate(Sender: TObject);
var
i:integer;
rsp:string;
begin
   if FileExists('test.mdb') then
  begin
  DB:=CoCatalog.Create;
  DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');

  ADOConnection2.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection2.LoginPrompt:=false;
   ADOConnection2.Connected:=true;
   ADOQuery2.Connection:=ADOConnection2;
   ADOQuery2.SQL.Clear;
   ADOQuery2.SQL.Add('SELECT * FROM grp');
   ADOQuery2.Active:=true;
   DataSource2.DataSet:=ADOQuery2;
   DBGrid2.DataSource:=DataSource2;
   DBGrid2.ReadOnly:=true;
   DBNavigator2.DataSource:=DataSource2;



   DBGrid2.DataSource.DataSet.RecNo;
    DBGrid2.Fields[0].Value;

   Panel3.Caption:='Номер групи: '+DBGrid2.Fields[1].Value;

   ADOConnection1.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection1.LoginPrompt:=false;
   ADOConnection1.Connected:=true;
   ADOQuery1.Connection:=ADOConnection1;
   ADOQuery1.SQL.Clear;
   ADOQuery1.SQL.Add('SELECT * FROM '+DBGrid2.Fields[0].Value);
   ADOQuery1.Active:=true;
   DataSource1.DataSet:=ADOQuery1;
   DBGrid1.DataSource:=DataSource1;
   DBNavigator1.DataSource:=DataSource1;

  DBGrid2.Columns[0].Width := 0; //ширина полей
  DBGrid2.Columns[1].Width := 100; //ширина полей

   DBGrid1.Columns[0].Width := 30; //ширина полей
   DBGrid1.Columns[1].Width := 40; //ширина полей
   DBGrid1.Columns[2].Width := 40; //ширина полей
  DBGrid1.Columns[3].Width := 80; //ширина полей
  DBGrid1.Columns[4].Width := 60; //ширина полей
  DBGrid1.Columns[5].Width := 200; //ширина полей
  DBGrid1.Columns[6].Width := 120; //ширина полей



   ADOConnection3.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection3.LoginPrompt:=false;
   ADOConnection3.Connected:=true;
   ADOQuery3.Connection:=ADOConnection3;
   ADOQuery3.SQL.Clear;
   ADOQuery3.SQL.Add('SELECT * FROM subject');
   ADOQuery3.Active:=true;
   DataSource3.DataSet:=ADOQuery3;
   DBGrid3.DataSource:=DataSource3;
   DBGrid3.ReadOnly:=true;
   DBNavigator3.DataSource:=DataSource3;



   ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection4.LoginPrompt:=false;
   ADOConnection4.Connected:=true;
   ADOQuery4.Connection:=ADOConnection4;
   ADOQuery4.SQL.Clear;
   ADOQuery4.SQL.Add('SELECT * FROM teacher');
   ADOQuery4.Active:=true;
   DataSource4.DataSet:=ADOQuery4;
   DBGrid4.DataSource:=DataSource4;
   DBGrid4.ReadOnly:=true;
   DBNavigator4.DataSource:=DataSource4;



   DBEdit1.DataSource:=DataSource1;
   DBEdit1.DataField:='План Б';

   DBEdit2.DataSource:=DataSource1;
   DBEdit2.DataField:='План К';

   DBEdit3.DataSource:=DataSource1;
   DBEdit3.DataField:='Коментар';




  DBLookupComboBox1.DataField:= 'Група';
  DBLookupComboBox1.DataSource:= DataSource1;
  DBLookupComboBox1.KeyField:= 'Група';
  DBLookupComboBox1.ListField:= 'Група';
  DBLookupComboBox1.ListSource:= DataSource2;
  DBLookupComboBox1.TabOrder:= 4;


  DBLookupComboBox2.DataField:= 'Дисципліна';
  DBLookupComboBox2.DataSource:= DataSource1;
  DBLookupComboBox2.KeyField:= 'Дисципліна';
  DBLookupComboBox2.ListField:= 'Дисципліна';
  DBLookupComboBox2.ListSource:= DataSource3;
  DBLookupComboBox2.TabOrder:= 4;

  DBLookupComboBox3.DataField:= 'Викладач';
  DBLookupComboBox3.DataSource:= DataSource1;
  DBLookupComboBox3.KeyField:= 'Викладач';
  DBLookupComboBox3.ListField:= 'Викладач';
  DBLookupComboBox3.ListSource:= DataSource4;
  DBLookupComboBox3.TabOrder:= 4;



     ADOConnection5.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection5.LoginPrompt:=false;
   ADOConnection5.Connected:=true;
   ADOQuery5.Connection:=ADOConnection5;
   ADOQuery5.SQL.Clear;

   rsp:=DBGrid2.Fields[0].Value+'rsp';
   ADOQuery5.SQL.Add('SELECT * FROM '+rsp);
   ADOQuery5.Active:=true;
   DataSource5.DataSet:=ADOQuery5;
   DBGrid5.DataSource:=DataSource5;
   //DBNavigator1.DataSource:=DataSource1;


  DBGrid5.Columns[0].Width := 15; //ширина полей
  DBGrid5.Columns[1].Width := 15; //ширина полей
  DBGrid5.Columns[2].Width := 15; //ширина полей
  DBGrid5.Columns[3].Width := 15; //ширина полей
  DBGrid5.Columns[4].Width := 15; //ширина полей
  DBGrid5.Columns[5].Width := 15; //ширина полей
  DBGrid5.Columns[6].Width := 15; //ширина полей
  DBGrid5.Columns[7].Width := 15; //ширина полей
  DBGrid5.Columns[8].Width := 15; //ширина полей
  DBGrid5.Columns[9].Width := 15; //ширина полей
  DBGrid5.Columns[10].Width := 15; //ширина полей
  DBGrid5.Columns[11].Width := 15; //ширина полей
  DBGrid5.Columns[12].Width := 15; //ширина полей
  DBGrid5.Columns[13].Width := 15; //ширина полей
  DBGrid5.Columns[14].Width := 15; //ширина полей
  DBGrid5.Columns[15].Width := 15; //ширина полей
  DBGrid5.Columns[16].Width := 15; //ширина полей
  DBGrid5.Columns[17].Width := 15; //ширина полей
  DBGrid5.Columns[18].Width := 15; //ширина полей
  DBGrid5.Columns[19].Width := 15; //ширина полей
  DBGrid5.Columns[20].Width := 15; //ширина полей
  DBGrid5.Columns[21].Width := 15; //ширина полей
  DBGrid5.Columns[22].Width := 15; //ширина полей
  DBGrid5.Columns[23].Width := 15; //ширина полей



  end
  else
  begin
  DB:=CoCatalog.Create;
  DB.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');
  DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');


  //Создание таблицы для групп
   ADOConnection2.Connected:=false;
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

     //Создание таблицы для предметов
   ADOConnection3.Connected:=false;
   Tables:=CoTable.Create;
   Tables.Name:='subject';

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
   Tables.Columns.Append('Дисципліна',adVarWChar,255);

   //Создание таблицы для преводавателей
   ADOConnection4.Connected:=false;
   Tables:=CoTable.Create;
   Tables.Name:='teacher';

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
   Tables.Columns.Append('Викладач',adVarWChar,255);



   ADOConnection2.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection2.LoginPrompt:=false;
   ADOConnection2.Connected:=true;
   ADOQuery2.Connection:=ADOConnection2;
   ADOQuery2.SQL.Clear;
   ADOQuery2.SQL.Add('SELECT * FROM grp');
   ADOQuery2.Active:=true;
   DataSource2.DataSet:=ADOQuery2;
   DBGrid2.DataSource:=DataSource2;
   DBGrid2.ReadOnly:=true;
   DBNavigator2.DataSource:=DataSource2;

      ADOConnection3.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection3.LoginPrompt:=false;
   ADOConnection3.Connected:=true;
   ADOQuery3.Connection:=ADOConnection3;
   ADOQuery3.SQL.Clear;
   ADOQuery3.SQL.Add('SELECT * FROM subject');
   ADOQuery3.Active:=true;
   DataSource3.DataSet:=ADOQuery3;
   DBGrid3.DataSource:=DataSource3;
   DBGrid3.ReadOnly:=true;
   DBNavigator3.DataSource:=DataSource3;

         ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
   ADOConnection4.LoginPrompt:=false;
   ADOConnection4.Connected:=true;
   ADOQuery4.Connection:=ADOConnection4;
   ADOQuery4.SQL.Clear;
   ADOQuery4.SQL.Add('SELECT * FROM teacher');
   ADOQuery4.Active:=true;
   DataSource4.DataSet:=ADOQuery4;
   DBGrid4.DataSource:=DataSource4;
   DBGrid4.ReadOnly:=true;
   DBNavigator4.DataSource:=DataSource4;


  end;

end;

procedure TForm1.FormDestroy(Sender: TObject);
begin

     excel := Unassigned;
   // excel.Workbooks.Close;
  //закрываем Excel
 // excel.Application.quit;
  //освобождаем интерфейсы

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
