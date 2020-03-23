unit fmProcGetDes;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, RzLabel, RzButton, Mask, RzEdit, DB, DBTables,unConnectDB,IniFiles,
  ExtCtrls, RzPanel, RzRadGrp;

const  C1 = 21469;  C2 = 12347;
const  KeyWord = 26493;

type
  TForm1 = class(TForm)
    RzEdit1: TRzEdit;
    RzButton1: TRzButton;
    Database1: TDatabase;
    Query1: TQuery;
    rgLX: TRzRadioGroup;
    RzButton2: TRzButton;
    RzEdit2: TRzEdit;
    RzButton3: TRzButton;
    RzEdit3: TRzEdit;
    RzLabel1: TRzLabel;
    procedure RzButton1Click(Sender: TObject);
    procedure RzButton2Click(Sender: TObject);
    procedure RzButton3Click(Sender: TObject);
  private
    { Private declarations }
    sUser,sPsw,sServer:string;
    function Decrypt(const S: String): String;
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

function TForm1.Decrypt(const S: String): String;
var
  I: byte;
  Key: Word;
begin
  Key := Keyword;
  SetLength(Result,Length(S));
  for I := 1 to Length(S) do
  begin
    Result[I] := char(byte(S[I]) xor (Key shr 8));
    Key := Word((byte(S[I]) + Key) * C1 + C2);
  end;
end;

procedure TForm1.RzButton1Click(Sender: TObject);
begin
  with Query1 do
  begin
    if AnsiLowerCase(Database1.DriverName)='oracle' then
      SQL.Text := 'select CONTENT from BFPUB8.XTXX where ID=150'
    else
      SQL.Text := 'select CONTENT from BFPUB.XTXX where ID=150';
    Open;
    if not IsEmpty then
      RzEdit1.Text := Decrypt(TrimRight(FieldByName('CONTENT').AsString))
    else
    begin
      RzEdit1.Text:='';
      ShowMessage('没有设置系统密钥！');
    end;
  end;
end;

procedure TForm1.RzButton2Click(Sender: TObject);
begin
  if rgLX.ItemIndex<0 then
  begin
    showmessage('请选择数据库类型');
    exit;
  end;
  if rgLX.ItemIndex=0 then
    ConnectDB(Database1,'DB_SYS.ini')
  else
    ConnectDB(Database1,'DB_ORA.ini');
  showmessage('连接数据库成功');
end;

procedure TForm1.RzButton3Click(Sender: TObject);
begin
  with Query1 do
  begin
    if AnsiLowerCase(Database1.DriverName)='oracle' then
      SQL.Text := 'select B.LOGIN_PASSWORD from BFPUB8.RYXX A,BFPUB8.XTCZY B where A.PERSON_ID=B.PERSON_ID and RYDM=:RYDM'
    else
      SQL.Text := 'select B.LOGIN_PASSWORD from BFPUB.RYXX A,BFPUB.XTCZY B where A.PERSON_ID=B.PERSON_ID and RYDM=:RYDM';
    ParamByName('RYDM').AsString:=RzEdit2.Text;
    Open;
    if not IsEmpty then
      RzEdit3.Text := Decrypt(TrimRight(FieldByName('LOGIN_PASSWORD').AsString))
    else
    begin
      RzEdit3.Text:='';
      ShowMessage('没有改操作员！');
    end;
  end;
end;

end.
