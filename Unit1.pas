unit Unit1;

interface

uses
  {Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.Grids, Vcl.DBGrids,
  Vcl.ComCtrls, Vcl.StdCtrls, Data.Win.ADODB, Vcl.ExtCtrls, Vcl.DBCtrls;}
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Data.DB, Vcl.ExtCtrls, Vcl.DBCtrls,
  Vcl.Grids, Vcl.DBGrids, Vcl.StdCtrls, Data.Win.ADODB, ADOX_TLB, Vcl.Mask,
  Vcl.ToolWin, Vcl.ComCtrls, Vcl.Menus, ComObj;

type
  TForm1 = class(TForm)
    TabControl1: TTabControl;
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
    OpenDialog1: TOpenDialog;
    Button6: TButton;
    Edit2: TEdit;
    Panel1: TPanel;
    Button7: TButton;
    Button8: TButton;
    procedure Button5Click(Sender: TObject);
    procedure Form1Destroy(Sender: TObject);
    procedure Form1Create(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure DBGrid1CellClick(Column: TColumn);
    procedure TabControl1Change(Sender: TObject);
    procedure Button3Click(Sender: TObject);
    procedure Button4Click(Sender: TObject);
    procedure Button6Click(Sender: TObject);
    procedure Button7Click(Sender: TObject);
    procedure Button8Click(Sender: TObject);
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

procedure TForm1.Button1Click(Sender: TObject); //������ �������� Exel �����
begin
  excel:=CreateOleObject('Excel.Application');
  if not OpenDialog1.Execute then exit;
  excel.WorkBooks.Open(OpenDialog1.FileName);
  excel.Visible:=true;
end;

function TextToTranslit(Text: string): string;
var i:integer;
begin
 for i:=1 to Length(text)*3 do
 begin
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('a',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('b',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('v',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('g',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('d',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('e',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('e',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('zh',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('z',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('i',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('y',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('k',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('l',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('m',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('n',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('o',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('p',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('r',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('s',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('t',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('u',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('f',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('h',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('c',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('ch',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('sh',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('sch',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('yi',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('ye',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('yu',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('ya',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('A',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('B',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('V',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('G',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('D',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('E',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('E',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Zh',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Z',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('I',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Y',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('K',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('L',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('M',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('N',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('O',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('P',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('R',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('S',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('T',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('U',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('F',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('H',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('C',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Ch',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Sh',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Sch',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Yi',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Ye',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Yu',text,i); end;
   if copy(text,i,1)='�' then begin delete(text,i,1); insert('Ya',text,i); end;
   if copy(text,i,1)=' ' then begin delete(text,i,1); insert('',text,i); end;
   if copy(text,i,1)='-' then begin delete(text,i,1); insert('',text,i); end;
 end;
 result:=Text;
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
  rsp,test:string;
begin
  test:=Edit1.Text;
  for i:=0 to Length(test) do  // �������� $ �� ���������
  begin
    Delete(test,Pos('-',test),1);
  end;
  rsp:=test+'rsp';

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

  ADOQuery1.Insert;
  ADOQuery1.FieldByName('�����').AsString := Edit1.Text;
  ADOQuery1.FieldByName('ID����').AsString := test;
  ADOQuery1.Post;

  DBGrid1.Columns[0].Width := 0; //������ �����
  DBGrid1.Columns[1].Width := 80; //������ �����

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

  DBGrid1.DataSource.DataSet.Delete
end;

procedure TForm1.Button8Click(Sender: TObject);
var
  rsp,test:string;
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

  ADOQuery1.Insert;
  ADOQuery1.FieldByName('�����').AsString := excel.WorkBooks[1].WorkSheets[1].Range['G'+x1].Value;
  ADOQuery1.FieldByName('ID����').AsString := test;
  ADOQuery1.Post;

  DBGrid1.Columns[0].Width := 0; //������ �����
  DBGrid1.Columns[1].Width := 80; //������ �����

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

  DBGrid2.Columns[0].Width := 30; //������ �����
  DBGrid2.Columns[1].Width := 40; //������ �����
  DBGrid2.Columns[2].Width := 40; //������ �����
  DBGrid2.Columns[3].Width := 70; //������ �����
  DBGrid2.Columns[4].Width := 0; //������ �����
  DBGrid2.Columns[5].Width := 220; //������ �����
  DBGrid2.Columns[6].Width := 120; //������ �����

  for i := 0 to 23 do DBGrid3.Columns[i].Width := 15; //������ �����
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
    DBNavigator1.DataSource:=DataSource1;
    DBGrid1.DataSource.DataSet.RecNo;
    DBGrid1.Fields[0].Value;

    DBGrid1.Columns[0].Width := 0; //������ �����
    DBGrid1.Columns[1].Width := 80; //������ �����
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
  TabControl1Change(TabControl1);
end;

procedure TForm1.Form1Destroy(Sender: TObject); //�������� ����� ��� ������ �� ���������
begin
  excel:=Unassigned;
end;

procedure TForm1.TabControl1Change(Sender: TObject); //�������
begin
  Panel1.Visible:=false;
  Edit1.Visible:=false;
  Edit2.Visible:=false;
  Button1.Visible:=false;
  Button2.Visible:=false;
  Button3.Visible:=false;
  Button4.Visible:=false;
  Button5.Visible:=false;
  Button6.Visible:=false;
  Button7.Visible:=false;
  DBGrid1.Visible:=false;
  DBGrid2.Visible:=false;
  DBGrid3.Visible:=false;
  DBGrid4.Visible:=false;
  DBGrid5.Visible:=false;
  DBNavigator1.Visible:=false;
  DBNavigator2.Visible:=false;
  if TabControl1.TabIndex=0 then
  begin
    ADOQuery1.Active:=true;
    Panel1.Visible:=true;
    Edit1.Visible:=true;
    Button1.Visible:=true;
    Button2.Visible:=true;
    Button3.Visible:=true;
    Button5.Visible:=true;
    Button7.Visible:=true;
    DBGrid1.Visible:=true;
    DBGrid2.Visible:=true;
    DBGrid3.Visible:=true;
    DBNavigator1.Visible:=true;
    DBNavigator2.Visible:=true;
  end;
  if TabControl1.TabIndex=1 then
  begin
    ADOQuery4.Active:=true;
    Edit2.Visible:=true;
    Button1.Visible:=true;
    Button4.Visible:=true;
    Button6.Visible:=true;
    DBGrid4.Visible:=true;
    DBGrid5.Visible:=true;
  end;
end;

end.
