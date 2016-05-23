unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Data.DB, Vcl.Grids,
  Vcl.DBGrids, Data.Win.ADODB, Vcl.ExtCtrls, Vcl.DBCtrls, Vcl.Mask, ComObj, ADOX_TLB ;

 // Vcl.ExtCtrls, Vcl.DBCtrls,
//ADOX_TLB, Vcl.Mask,
 // Vcl.ToolWin, Vcl.ComCtrls, Vcl.Menus, ComObj, Unit2;

type
  TForm1 = class(TForm)
    ADOConnection1: TADOConnection;
    ADOQuery1: TADOQuery;
    DataSource1: TDataSource;
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    ADOConnection2: TADOConnection;
    ADOQuery2: TADOQuery;
    DataSource2: TDataSource;
    DBGrid3: TDBGrid;
    DBGrid4: TDBGrid;
    DBGrid5: TDBGrid;
    ADOConnection3: TADOConnection;
    ADOQuery3: TADOQuery;
    DataSource3: TDataSource;
    ADOConnection4: TADOConnection;
    ADOQuery4: TADOQuery;
    DataSource4: TDataSource;
    ADOConnection5: TADOConnection;
    ADOQuery5: TADOQuery;
    DataSource5: TDataSource;
    DBNavigator1: TDBNavigator;
    DBNavigator2: TDBNavigator;
    DBNavigator3: TDBNavigator;
    DBNavigator4: TDBNavigator;
    DBNavigator5: TDBNavigator;
    DBEdit1: TDBEdit;
    DBEdit2: TDBEdit;
    DBEdit3: TDBEdit;
    DBLookupComboBox1: TDBLookupComboBox;
    DBLookupComboBox2: TDBLookupComboBox;
    DBLookupComboBox3: TDBLookupComboBox;
    Panel3: TPanel;
    procedure FormCreate(Sender: TObject);
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

procedure TForm1.FormCreate(Sender: TObject);
var
i:integer;
rsp:string;
begin
   if FileExists('H:\test.mdb') then
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

end.
