unit Unit1;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.ExtCtrls, Vcl.DBCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls, Data.Win.ADODB, ADOX_TLB, Vcl.Mask,
  Vcl.ToolWin, Vcl.ComCtrls, Vcl.Menus, ComObj, Vcl.Buttons;

type
  TForm1 = class(TForm)
    DBGrid1: TDBGrid;
    DBGrid2: TDBGrid;
    DBGrid3: TDBGrid;
    DBGrid4: TDBGrid;
    DBGrid5: TDBGrid;
    Button1: TButton;
    Edit1: TEdit;
    Button2: TButton;
    Button3: TButton;
    Button4: TButton;
    Button5: TButton;
    DBNavigator1: TDBNavigator;
    DBNavigator2: TDBNavigator;
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
    Button6: TButton;
    Edit2: TEdit;
    Panel1: TPanel;
    Button7: TButton;
    Button8: TButton;
    PageControl1: TPageControl;
    TabSheet1: TTabSheet;
    TabSheet2: TTabSheet;
    TabSheet3: TTabSheet;
    TabSheet4: TTabSheet;
    DateTimePicker1: TDateTimePicker;
    DateTimePicker2: TDateTimePicker;
    Label1: TLabel;
    Label2: TLabel;
    Edit3: TEdit;
    OpenDialog1: TOpenDialog;
    ComboBox1: TComboBox;
    Label3: TLabel;
    Label4: TLabel;
    DBGrid6: TDBGrid;
    Button9: TButton;
    DBGrid7: TDBGrid;
    ADOConnection7: TADOConnection;
    DataSource7: TDataSource;
    ADOQuery7: TADOQuery;
    Panel2: TPanel;
    DBNavigator7: TDBNavigator;
    TabSheet5: TTabSheet;
    PageControl2: TPageControl;
    TabSheet6: TTabSheet;
    TabSheet7: TTabSheet;
    SpeedButton1: TSpeedButton;
    DBGrid8: TDBGrid;
    DateTimePicker3: TDateTimePicker;
    Button10: TButton;
    Edit4: TEdit;
    DBGrid9: TDBGrid;
    DBLookupComboBox1: TDBLookupComboBox;
    ListBox1: TListBox;
    ADOConnection9: TADOConnection;
    DataSource9: TDataSource;
    ADOQuery9: TADOQuery;
    DBNavigator9: TDBNavigator;
    Panel3: TPanel;
    Label5: TLabel;
    Label6: TLabel;
    Label7: TLabel;
    Label8: TLabel;
    Button11: TButton;
    DBListBox1: TDBListBox;
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
    procedure DateTimePicker3CloseUp(Sender: TObject);
    procedure Button11Click(Sender: TObject);
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

procedure TForm1.Button10Click(Sender: TObject);
var
  Result,test,all:string;
  i,y:integer;
begin


 {case DayOfWeek(DateTimePicker3.Date) of
  1: Result := '�����������';
  2: Result := '�����������';
  3: Result := '�������';
  4: Result := '�����';
  5: Result := '�������';
  6: Result := '�������';
  7: Result := '�������';
  end;
  Edit4.Text:=Result;  }


   for i:=0 to ListBox1.Items.Count-1 do
   begin
   test:=ListBox1.Items[i];
   all:=all+ListBox1.Items[i]+',';

    for y:=0 to Length(test) do  // �������� "-" �� ���������
  begin
    Delete(test,Pos('-',test),1);
  end;

  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM '+test+'subb');
  ADOQuery9.Active:=true;

    ADOQuery9.Insert;
    ADOQuery9.FieldByName('�����').Value:=ListBox1.Items[i];
    ADOQuery9.FieldByName('����').Value:=DateToStr(DateTimePicker3.Date);
    ADOQuery9.FieldByName('������').Value:=edit4.text;
    ADOQuery9.Post;

   end;

   ShowMessage('��� ����: '+all+' ���� ������ �������� ������ '+DateToStr(DateTimePicker3.Date));
   ListBox1.Clear;





  DBGrid9.Columns[0].Width := 80; //������ �����
  DBGrid9.Columns[1].Width := 80; //������ �����
  DBGrid9.Columns[2].Width := 80; //������ �����


end;

procedure TForm1.Button11Click(Sender: TObject);
var
kol,i:integer;
begin
  kol:=ADOQuery1.RecordCount;  

   for i:=0 to kol-1 do
   begin
  //  DBGrid1.SelectedIndex:=3;
  DBGrid8.Fields[1].Value;
 // ListBox1.Items.add(DBGrid1.);


   end;


end;

procedure TForm1.Button1Click(Sender: TObject); //������ �������� Exel �����
begin
  excel:=CreateOleObject('Excel.Application');
  if not OpenDialog1.Execute then exit;
  excel.WorkBooks.Open(OpenDialog1.FileName);
  excel.Visible:=true;
end;

procedure TForm1.Button2Click(Sender: TObject); //������ ���������� ���������� ������
var
  y,x1,x2,add:string;
  nom,i:integer;
begin
  y:=excel.Selection.Address;

  for i:=0 to Length(y) do  // �������� $ �� ���������
  begin
    Delete(y,Pos('$',y),1);
  end;

  x1:=Copy(y,2,Pos(':',y)-2);   // �������� ������ ������ �� ���������
  x2:=Copy(y,Pos(':',y)+2,Length(y));

  nom:=strtoint(x2)-strtoint(x1);  // ���������� �������
  for i:=0 to nom  do
  begin
    add:=inttostr(strtoint(x1)+i);  // ����� ��������
    ADOQuery2.Insert;
    ADOQuery2.FieldByName('���� �').Value :=excel.WorkBooks[1].WorkSheets[1].Range['C'+add].Value;
    ADOQuery2.FieldByName('���� �').Value :=excel.WorkBooks[1].WorkSheets[1].Range['D'+add].Value;
    ADOQuery2.FieldByName('��������').Value :=excel.WorkBooks[1].WorkSheets[1].Range['E'+add].Value;
    ADOQuery2.FieldByName('�����').Value :=excel.WorkBooks[1].WorkSheets[1].Range['F'+add].Value;
    ADOQuery2.FieldByName('���������').Value :=excel.WorkBooks[1].WorkSheets[1].Range['G'+add].Value;
    ADOQuery2.FieldByName('��������').Value :=excel.WorkBooks[1].WorkSheets[1].Range['H'+add].Value;
    ADOQuery2.Post;
  end;
end;

procedure TForm1.Button3Click(Sender: TObject); //������ ���������� ���������� �� ���������/�����������
var
  y,x1,x2,add:string;
  nom,i:integer;
begin
  y:=excel.Selection.Address;

  for i:=0 to Length(y) do  // �������� $ �� ���������
  begin
    Delete(y,Pos('$',y),1);
  end;

  x1:=Copy(y,3,Pos(':',y)-3);   // �������� ������ ������ �� ���������
  x2:=Copy(y,Pos(':',y)+3,Length(y));

  nom:=strtoint(x2)-strtoint(x1);  // ���������� �������
  for i:=0 to nom  do
  begin
    add:=inttostr(strtoint(x1)+i);  // ����� ��������
    ADOQuery3.Insert;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['MZ'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NA'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NB'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NC'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['ND'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NE'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NF'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NG'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NH'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NI'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NJ'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NK'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NL'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NM'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NN'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NO'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NP'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NQ'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NR'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NS'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NT'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NU'+add].Value;
    ADOQuery3.FieldByName('���1').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NV'+add].Value;
    ADOQuery3.FieldByName('���2').Value :=excel.WorkBooks[1].WorkSheets[1].Range['NW'+add].Value;
    ADOQuery3.Post;
  end;
end;

procedure TForm1.Button4Click(Sender: TObject); //������ ���������� ������ ��������������
var
  y,x1,x2,add:string;
  nom,i:integer;
begin
  y:=excel.Selection.Address;

  for i:=0 to Length(y) do  // �������� $ �� ���������
  begin
    Delete(y,Pos('$',y),1);
  end;

  x1:=Copy(y,2,Pos(':',y)-2);   // �������� ������ ������ �� ���������
  x2:=Copy(y,Pos(':',y)+2,Length(y));

  nom:=strtoint(x2)-strtoint(x1);  // ���������� �������
  for i:=0 to nom  do
  begin
    add:=inttostr(strtoint(x1)+i);  // ����� ��������
    ADOQuery4.Insert;
    ADOQuery4.FieldByName('ID���������').Value :=excel.WorkBooks[1].WorkSheets[4].Range['A'+add].Value;
    ADOQuery4.FieldByName('��������').Value :=excel.WorkBooks[1].WorkSheets[4].Range['A'+add].Value;
    ADOQuery4.Post;
  end;
end;

procedure TForm1.Button5Click(Sender: TObject); //���������� ������ � ������
var
  i:integer;
  rsp,test,graf,subb:string;
begin
  test:=Edit1.Text;
  for i:=0 to Length(test) do  // �������� - �� ���������
  begin
    Delete(test,Pos('-',test),1);
  end;
  rsp:=test+'rsp';
  graf:=test+'graf';
  subb:=test+'subb';


  ADOConnection2.Connected:=false;
  ADOConnection3.Connected:=false;

  //�������� ������� �����
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
    Properties['Description'].Value:='�������� ����';
  end;

  Tables.Columns.Append(Columns,0,0);
  Tables.Columns.Append('���� �',adInteger,0);
  Tables.Columns.Append('���� �',adInteger,0);
  Tables.Columns.Append('��������',adVarWChar,255);
  Tables.Columns.Append('�����',adVarWChar,255);
  Tables.Columns.Append('���������',adVarWChar,255);
  Tables.Columns.Append('��������',adVarWChar,255);

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

  //�������� ������� ����������
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
    Properties['Description'].Value:='�������� ����';
  end;

  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);

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


  // �������� ������� �������

  ADOConnection7.Connected:=false;

  //�������� ������� �����
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
    Properties['Description'].Value:='�������� ����';
  end;

  Tables.Columns.Append(Columns,0,0);
  Tables.Columns.Append('��� ����',adVarWChar,255);
  Tables.Columns.Append('ʳ� ����',adVarWChar,255);
  Tables.Columns.Append('��� ����',adVarWChar,255);
  Tables.Columns.Append('������',adVarWChar,255);
  Tables.Columns.Append('�����',adVarWChar,255);

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



  DBGrid7.Columns[0].Width := 0; //������ �����
  DBGrid7.Columns[1].Width := 80; //������ �����
  DBGrid7.Columns[2].Width := 80; //������ �����
  DBGrid7.Columns[3].Width := 80; //������ �����
  DBGrid7.Columns[4].Width := 80; //������ �����




  // �������� ������� ��������

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
    Properties['Description'].Value:='�������� ����';
  end;

  Tables.Columns.Append(Columns,0,0);
  Tables.Columns.Append('�����',adVarWChar,255);
  Tables.Columns.Append('����',adVarWChar,255);
  Tables.Columns.Append('������',adVarWChar,255);

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




  DBGrid9.Columns[0].Width := 80; //������ �����
  DBGrid9.Columns[1].Width := 80; //������ �����
  DBGrid9.Columns[2].Width := 80; //������ �����


  //////////////////////////////////////////////////////////////////



  ADOQuery1.Insert;
  ADOQuery1.FieldByName('�����').AsString := Edit1.Text;
  ADOQuery1.FieldByName('ID����').AsString := test;
  ADOQuery1.Post;

  DBGrid1.Columns[0].Width := 0; //������ �����
  DBGrid1.Columns[1].Width := 80; //������ �����

  //������� � ������ �� ������� �������
  DBGrid6.Columns[0].Width := 0; //������ �����
  DBGrid6.Columns[1].Width := 80; //������ �����


    //������� � ������ �� ������� �������
  DBGrid8.Columns[0].Width := 0; //������ �����
  DBGrid8.Columns[1].Width := 80; //������ �����

  DBGrid2.Columns[0].Width := 30; //������ �����
  DBGrid2.Columns[1].Width := 40; //������ �����
  DBGrid2.Columns[2].Width := 40; //������ �����
  DBGrid2.Columns[3].Width := 70; //������ �����
  DBGrid2.Columns[4].Width := 0; //������ �����
  DBGrid2.Columns[5].Width := 220; //������ �����
  DBGrid2.Columns[6].Width := 120; //������ �����

  for i := 0 to 23 do DBGrid3.Columns[i].Width := 15; //������ �����

  Edit1.Text:='';
end;

procedure TForm1.Button6Click(Sender: TObject); //������ ������ ��������������
begin
  if not ADOQuery4.Locate('��������',Edit2.Text,[loCaseInsensitive, loPartialKey])then
  ShowMessage('������ �� �������');
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


  //������� � �������� �� ������� ������
  //DBGrid6.DataSource.DataSet.Delete;
end;

procedure TForm1.Button8Click(Sender: TObject);
var
  rsp,test,graf,subb:string;
  y,x1,x2,add:string;
  nom,i:integer;
  z:integer;
begin

 y:=excel.Selection.Address;
 // excel.WorkBooks[1].WorkSheets[1].Range['C14:H28'].Select;
 // y:='$C$14:$H$28';

  //������ �������� ��������� � �������� ������� � ���������� �������

  for i:=0 to Length(y) do  // �������� $ �� ���������
  begin
  Delete(y,Pos('$',y),1);
  end;

  x1:=Copy(y,2,Pos(':',y)-2);   // �������� ������ ������ �� ���������
  x2:=Copy(y,Pos(':',y)+2,Length(y));


 test:=excel.WorkBooks[1].WorkSheets[1].Range['G'+x1].Value;



  for i:=0 to Length(test) do  // �������� - �� ���������
  begin
    Delete(test,Pos('-',test),1);
  end;
  rsp:=test+'rsp';
  graf:=test+'graf';
  subb:=test+'subb';

  ADOConnection2.Connected:=false;
  ADOConnection3.Connected:=false;

  //�������� ������� �����
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
    Properties['Description'].Value:='�������� ����';
  end;

  Tables.Columns.Append(Columns,0,0);
  Tables.Columns.Append('���� �',adInteger,0);
  Tables.Columns.Append('���� �',adInteger,0);
  Tables.Columns.Append('��������',adVarWChar,255);
  Tables.Columns.Append('�����',adVarWChar,255);
  Tables.Columns.Append('���������',adVarWChar,255);
  Tables.Columns.Append('��������',adVarWChar,255);

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

  //�������� ������� ����������
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
    Properties['Description'].Value:='�������� ����';
  end;

  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);
  Tables.Columns.Append('���1',adVarWChar,255);
  Tables.Columns.Append('���2',adVarWChar,255);

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

  // �������� ������� �������

  ADOConnection7.Connected:=false;

  //�������� ������� �����
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
    Properties['Description'].Value:='�������� ����';
  end;

  Tables.Columns.Append(Columns,0,0);
  Tables.Columns.Append('��� ����',adVarWChar,255);
  Tables.Columns.Append('ʳ� ����',adVarWChar,255);
  Tables.Columns.Append('��� ����',adVarWChar,255);
  Tables.Columns.Append('������',adVarWChar,255);
  Tables.Columns.Append('�����',adVarWChar,255);

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

  //�������� ������� ����������


  DBGrid7.Columns[0].Width := 0; //������ �����
  DBGrid7.Columns[1].Width := 80; //������ �����
  DBGrid7.Columns[2].Width := 80; //������ �����
  DBGrid7.Columns[3].Width := 80; //������ �����
  DBGrid7.Columns[4].Width := 80; //������ �����

  // �������� ������� ��������

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
    Properties['Description'].Value:='�������� ����';
  end;

  Tables.Columns.Append(Columns,0,0);
  Tables.Columns.Append('�����',adVarWChar,255);
  Tables.Columns.Append('����',adVarWChar,255);
  Tables.Columns.Append('������',adVarWChar,255);

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




  DBGrid9.Columns[0].Width := 80; //������ �����
  DBGrid9.Columns[1].Width := 80; //������ �����
  DBGrid9.Columns[2].Width := 80; //������ �����


  //////////////////////////////////////////////////////////////////


  ADOQuery1.Insert;
  ADOQuery1.FieldByName('�����').AsString := excel.WorkBooks[1].WorkSheets[1].Range['G'+x1].Value;
  ADOQuery1.FieldByName('ID����').AsString := test;
  ADOQuery1.Post;

  DBGrid1.Columns[0].Width := 0; //������ �����
  DBGrid1.Columns[1].Width := 80; //������ �����

  // ���� � ��� �� ������� ������
  DBGrid6.Columns[0].Width := 0; //������ �����
  DBGrid6.Columns[1].Width := 80; //������ �����

    // ���� � ��� �� ������� �������
  DBGrid8.Columns[0].Width := 0; //������ �����
  DBGrid8.Columns[1].Width := 80; //������ �����

  DBGrid2.Columns[0].Width := 30; //������ �����
  DBGrid2.Columns[1].Width := 40; //������ �����
  DBGrid2.Columns[2].Width := 40; //������ �����
  DBGrid2.Columns[3].Width := 70; //������ �����
  DBGrid2.Columns[4].Width := 0; //������ �����
  DBGrid2.Columns[5].Width := 220; //������ �����
  DBGrid2.Columns[6].Width := 120; //������ �����

  for i := 0 to 23 do DBGrid3.Columns[i].Width := 15; //������ �����

  Edit1.Text:='';



   x1:=inttostr(strtoint(x1)+3);


  nom:=strtoint(x2)-strtoint(x1);  // ���������� �������
    for i:=0 to nom  do
    begin
    add:=inttostr(strtoint(x1)+i);  // ����� ��������
  ADOQuery2.Insert;
  ADOQuery2.FieldByName('���� �').Value :=excel.WorkBooks[1].WorkSheets[1].Range['C'+add].Value;
  ADOQuery2.FieldByName('���� �').Value :=excel.WorkBooks[1].WorkSheets[1].Range['D'+add].Value;
  ADOQuery2.FieldByName('��������').Value :=excel.WorkBooks[1].WorkSheets[1].Range['E'+add].Value;
  ADOQuery2.FieldByName('�����').Value :=excel.WorkBooks[1].WorkSheets[1].Range['F'+add].Value;
  ADOQuery2.FieldByName('���������').Value :=excel.WorkBooks[1].WorkSheets[1].Range['G'+add].Value;
  ADOQuery2.FieldByName('��������').Value :=excel.WorkBooks[1].WorkSheets[1].Range['H'+add].Value;
  ADOQuery2.Post;

    end;

end;

procedure TForm1.Button9Click(Sender: TObject);
begin

    ADOQuery7.Insert;
    ADOQuery7.FieldByName('��� ����').Value:=DateToStr(DateTimePicker1.Date);
    ADOQuery7.FieldByName('ʳ� ����').Value:=DateToStr(DateTimePicker2.Date);
    ADOQuery7.FieldByName('��� ����').Value:=ComboBox1.ItemIndex;
    ADOQuery7.FieldByName('������').Value:=Edit3.Text;
    ADOQuery7.FieldByName('�����').Value:=DBGrid6.Fields[1].Value;
    ADOQuery7.Post;

end;

procedure TForm1.DateTimePicker3CloseUp(Sender: TObject);
begin
  if DayOfWeek(DateTimePicker3.Date)<>7 then
  ShowMessage('������� ������!');
end;

procedure TForm1.DBGrid1CellClick(Column: TColumn); //����� ������ �� ������ ��� ��������� ����������
var
  i: integer;
begin
  DBGrid1.DataSource.DataSet.RecNo;
  DBGrid1.Fields[0].Value;

  Panel1.Caption:='����� '+DBGrid1.Fields[1].Value;

  ADOQuery2.SQL.Clear;
  ADOQuery2.SQL.Add('SELECT * FROM '+DBGrid1.Fields[0].Value);
  ADOQuery2.Active:=true;

  ADOQuery3.SQL.Clear;
  ADOQuery3.SQL.Add('SELECT * FROM '+DBGrid1.Fields[0].Value+'rsp');
  ADOQuery3.Active:=true;

  DBGrid2.Columns[0].Width := 0; //������ �����
  DBGrid2.Columns[1].Width := 40; //������ �����
  DBGrid2.Columns[2].Width := 40; //������ �����
  DBGrid2.Columns[3].Width := 70; //������ �����
  DBGrid2.Columns[4].Width := 0; //������ �����
  DBGrid2.Columns[5].Width := 220; //������ �����
  DBGrid2.Columns[6].Width := 120; //������ �����

  for i := 0 to 23 do DBGrid3.Columns[i].Width := 15; //������ �����
end;

procedure TForm1.DBGrid6CellClick(Column: TColumn);
var
  i: integer;
begin
  DBGrid6.DataSource.DataSet.RecNo;
  DBGrid6.Fields[0].Value;

  Panel2.Caption:='����� '+DBGrid6.Fields[1].Value;

  ADOQuery7.SQL.Clear;
  ADOQuery7.SQL.Add('SELECT * FROM '+DBGrid6.Fields[0].Value+'graf');
  ADOQuery7.Active:=true;


  DBGrid7.Columns[0].Width := 0; //������ �����
  DBGrid7.Columns[1].Width := 80; //������ �����
  DBGrid7.Columns[2].Width := 80; //������ �����
  DBGrid7.Columns[3].Width := 80; //������ �����
  DBGrid7.Columns[4].Width := 80; //������ �����


end;

procedure TForm1.DBGrid8CellClick(Column: TColumn);
var
  i: integer;
begin
  DBGrid8.DataSource.DataSet.RecNo;
  DBGrid8.Fields[0].Value;

  Panel3.Caption:='����� '+DBGrid8.Fields[1].Value;

  ADOQuery9.SQL.Clear;
  ADOQuery9.SQL.Add('SELECT * FROM '+DBGrid8.Fields[0].Value+'subb');
  ADOQuery9.Active:=true;


  DBGrid9.Columns[0].Width := 80; //������ �����
  DBGrid9.Columns[1].Width := 80; //������ �����
  DBGrid9.Columns[2].Width := 80; //������ �����


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

procedure TForm1.Form1Create(Sender: TObject); //�������� � ������ ������ ��� �������
var
  i:integer;
  rsp:string;
begin
  if FileExists('test.mdb') then
  begin
    DB:=CoCatalog.Create;
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');
    //������ �����
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

    DBGrid1.DataSource.DataSet.RecNo;
    DBGrid1.Fields[0].Value;

    DBGrid6.DataSource.DataSet.RecNo;
    DBGrid6.Fields[0].Value;

    DBGrid8.DataSource.DataSet.RecNo;
    DBGrid8.Fields[0].Value;

    DBGrid1.Columns[0].Width := 0; //������ �����
    DBGrid1.Columns[1].Width := 80; //������ �����

    DBGrid6.Columns[0].Width := 0; //������ �����
    DBGrid6.Columns[1].Width := 80; //������ �����

    DBGrid8.Columns[0].Width := 0; //������ �����
    DBGrid8.Columns[1].Width := 80; //������ �����
    //���������� ������
    Panel1.Caption:='����� '+DBGrid1.Fields[1].Value;
    ADOConnection2.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb;';
    ADOConnection2.LoginPrompt:=false;
    ADOConnection2.Connected:=true;
    ADOQuery2.Connection:=ADOConnection1;
    ADOQuery2.SQL.Clear;
    ADOQuery2.SQL.Add('SELECT * FROM '+DBGrid1.Fields[0].Value);
    ADOQuery2.Active:=true;
    DataSource2.DataSet:=ADOQuery2;
    DBGrid2.DataSource:=DataSource2;
    DBNavigator2.DataSource:=DataSource2;

    DBGrid2.Columns[0].Width := 30; //������ �����
    DBGrid2.Columns[1].Width := 40; //������ �����
    DBGrid2.Columns[2].Width := 40; //������ �����
    DBGrid2.Columns[3].Width := 70; //������ �����
    DBGrid2.Columns[4].Width := 0; //������ �����
    DBGrid2.Columns[5].Width := 220; //������ �����
    DBGrid2.Columns[6].Width := 120; //������ �����
    //���������� ���������/�����������
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
    for i := 0 to 23 do DBGrid3.Columns[i].Width := 15; //������ �����
    //������ ��������������
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
    DBGrid4.Fields[0].Value;

    DBGrid4.Columns[0].Width := 0; //������ �����
    DBGrid4.Columns[1].Width := 80; //������ �����


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

    DBGrid7.Columns[0].Width := 0; //������ �����
    DBGrid7.Columns[1].Width := 80; //������ �����
    DBGrid7.Columns[2].Width := 80; //������ �����
    DBGrid7.Columns[3].Width := 80; //������ �����
    DBGrid7.Columns[4].Width := 80; //������ �����





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

    DBGrid9.Columns[0].Width := 80; //������ �����
    DBGrid9.Columns[1].Width := 80; //������ �����
    DBGrid9.Columns[2].Width := 80; //������ �����






  end
  else
  begin
    DB:=CoCatalog.Create;
    DB.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=test.mdb');
    //�������� ������� ��� �����
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
      Properties['Description'].Value:='�������� ����';
    end;
    Tables.Columns.Append('ID����',adVarWChar,255);
    Tables.Columns.Append('�����',adVarWChar,255);
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
    //�������� ������� ��� ��������������
    ADOConnection1.Connected:=false;
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
      Properties['Description'].Value:='�������� ����';
    end;
    Tables.Columns.Append('ID���������',adVarWChar,255);
    Tables.Columns.Append('��������',adVarWChar,255);
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
  end;
  {
  if FileExists('teacher.mdb') then
  begin
    DB:=CoCatalog.Create;
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=teacher.mdb');
    //������ ��������������
    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=teacher.mdb;';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection1;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT * FROM thr');
    ADOQuery4.Active:=true;
    DataSource1.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
    DBGrid4.ReadOnly:=true;
    DBNavigator1.DataSource:=DataSource4;
    DBGrid4.DataSource.DataSet.RecNo;
    DBGrid4.Fields[0].Value;

    DBGrid1.Columns[0].Width := 0; //������ �����
    DBGrid1.Columns[1].Width := 80; //������ �����
  end
  else
  begin
    DB:=CoCatalog.Create;
    DB.Create('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=teacher.mdb');
    DB.Set_ActiveConnection('Provider=Microsoft.Jet.OLEDB.4.0;Data Source=teacher.mdb');
    //�������� ������� ��� ��������������
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
      Properties['Description'].Value:='�������� ����';
    end;
    Tables.Columns.Append('ID���������',adVarWChar,255);
    Tables.Columns.Append('��������',adVarWChar,255);
    ADOConnection4.ConnectionString:='Provider=Microsoft.Jet.OLEDB.4.0;Data Source=teacher.mdb;';
    ADOConnection4.LoginPrompt:=false;
    ADOConnection4.Connected:=true;
    ADOQuery4.Connection:=ADOConnection4;
    ADOQuery4.SQL.Clear;
    ADOQuery4.SQL.Add('SELECT * FROM thr');
    ADOQuery4.Active:=true;
    DataSource4.DataSet:=ADOQuery4;
    DBGrid4.DataSource:=DataSource4;
    DBGrid4.ReadOnly:=true;
  end;
  }
  //bControl1Change(TabControl1);




end;

procedure TForm1.Form1Destroy(Sender: TObject); //�������� ����� ��� ������ �� ���������
begin
  excel:=Unassigned;
end;

procedure TForm1.ListBox1DblClick(Sender: TObject);
begin
  ListBox1.Items.Delete(ListBox1.ItemIndex);
end;

end.
