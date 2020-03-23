unit Decrypt_PL;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs,unConnectDB, DB, DBTables, StdCtrls, ComCtrls, Grids, DBGrids,
  RzDBGrid, RzButton,unDes, RzLabel, RzEdit,unDBTools,IniFiles;

const
  DesKey_C = 'PdfSsoDi';

type
  TfmProcDecrypt = class(TForm)
    GroupBox1: TGroupBox;
    Memo1: TMemo;
    Button1: TButton;
    Database1: TDatabase;
    tmpQry: TQuery;
    qrySave: TQuery;
    DBGrid1: TRzDBGrid;
    tblSFZ: TTable;
    dsSFZ: TDataSource;
    Label1: TLabel;
    RzButton1: TRzButton;
    SaveDialog: TSaveDialog;
    tblSFZHYK_NO: TStringField;
    tblSFZCDNR: TStringField;
    RzMemo1: TRzMemo;
    RzLabel1: TRzLabel;
    RzLabel2: TRzLabel;
    procedure FormCreate(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure RzButton1Click(Sender: TObject);
  private
    sUser,sPsw,sServer:string;
    { Private declarations }
  //内部函数
    function getDES_XT: TDesKey;
    function Encrypt_DES(const S: string; DesKey: TDesKey): string;
    function Decrypt_DES(const S: string; DesKey: TDesKey): string;
    function CardEncrypt_DES(const S: string; DesKey: TDesKey): string;
    function CardDecrypt_DES(const S: string; DesKey: TDesKey): string;
    procedure WriteDebugLog(sLog: string);
  public
    { Public declarations }
  end;

var
  fmProcDecrypt: TfmProcDecrypt;

implementation

{$R *.dfm}

procedure TfmProcDecrypt.FormCreate(Sender: TObject);
var
  ini: TIniFile;
  iIP,iDBname:string;
  i,j,iPort: integer;
  sIP,sIP1,sDBName:string;
begin
  //读取配置文件
  ini := TIniFile.Create('.\DB.INI');
  try
    sUser := ini.ReadString('DB','USER','');
    sPsw := ini.ReadString('DB','PASSWORD','');
    sServer := ini.ReadString('DB','SERVER','');
  finally
    ini.Free ;
  end;

  i:=pos(',',sServer);
  if i = 0 then
    Exit
  else
  begin
    iIP := Trim(Copy(sServer, 1, i - 1));
    sServer:=Trim(Copy(sServer, i+1, 1000));
    j:=pos(',',sServer);
    if j = 0 then
      Exit
    else
    begin
      iPort := StrToInt(Trim(Copy(sServer, 1, j - 1)));
      iDBname:=Trim(Copy(sServer, j+1, 1000))
    end;
  end;
  sIP:= iIP+','+IntToStr(iPort)+','+iDBname;
  sIP1:= iIP+','+IntToStr(iPort);
  SetDBConfigureFile('CRM_PLJM',sIP,sIP1);
  OpenDatabase(Database1,'CRM_PLJM',iDBname,sUser,sPsw);

  //ConnectDB(Database1,'DB.ini');
  tblSFZ.TableName:='xx';
  tblSFZ.CreateTable;
  tblSFZ.Open;
end;

procedure TfmProcDecrypt.Button1Click(Sender: TObject);
var
  s:string;
begin
  with tmpQry do
  begin
    SQL.Text := Memo1.Text;
    Open();
  end;
  tblSFZ.Close;
  tblSFZ.BatchMove(tmpQry,batCopy);
  tblSFZ.Open;

  showmessage('执行完成！请生成文件！');
end;

procedure TfmProcDecrypt.FormDestroy(Sender: TObject);
begin
  tblSFZ.close;
  tblSFZ.DeleteTable;
end;

procedure TfmProcDecrypt.RzButton1Click(Sender: TObject);
var
  TxtFile: TextFile;
  str: string;
begin
  str:='';
  tblSFZ.First;
  while not tblSFZ.Eof do
  begin
    str:=str+tblSFZHYK_NO.AsString+','+CARDDecrypt_DES(tblSFZCDNR.AsString,getDES_XT)+#13#10;
    tblSFZ.Next;
  end;
  WriteDebugLog(str);
end;

procedure TfmProcDecrypt.WriteDebugLog(sLog: string);
var
  LogFile: TextFile;
  sFileName,str: string;
begin
  sFileName := '.\CDNR.log';
  try
    str := sLog;
    AssignFile(LogFile, sFileName);
    if FileExists(sFileName) then
      Rewrite(LogFile)
    else
      Rewrite(LogFile);
    Writeln(LogFile, str);
    Flush(LogFile);
    CloseFile(LogFile);
  finally
  end;
end;

function TfmProcDecrypt.CardDecrypt_DES(const S: string;
  DesKey: TDesKey): string;
var
  i, len: integer;
  S2, TmpStr: string;
begin
  S2 := Decrypt_DES(S, DesKey);
  len := length(S2);
  SetLength(tmpstr, len * 2);
  for i := 1 to Len do
  begin
    tmpstr[2 * i] := char(((Ord(S2[len - i + 1]) shr 4) and $F) + $30);
    tmpstr[2 * i - 1] := char((Ord(S2[len - i + 1]) and $F) + $30);
  end;
  if tmpstr[1] = ':' then
    Result := Copy(tmpstr, 2, 1000)
  else
    Result := tmpstr;
end;

function TfmProcDecrypt.CardEncrypt_DES(const S: string;
  DesKey: TDesKey): string;
var
  i, len1, len2: integer;
  s1, s2, tmpstr: string;
  output: array of Byte;
begin
  len1 := Length(S);
  if Odd(len1) then
  begin
    tmpstr := ':' + S;
    Inc(Len1);
  end
  else
    tmpstr := S;
  len2 := len1 div 2;
  SetLength(s1, Len2);
  for i := 1 to Len2 do
    s1[i] := char(((Ord(tmpstr[Len1 - 2 * (i - 1)]) - $30) shl 4) or (Ord(tmpstr[Len1 - 2 * (i - 1) - 1]) - $30));
  Result := Encrypt_DES(S1, DesKey)
end;

function TfmProcDecrypt.Decrypt_DES(const S: string;
  DesKey: TDesKey): string;
var
  i, len: integer;
  s1, s2: string;
  input: array of Byte;
begin
  len := length(S) div 2;
  setlength(input, len);
  s1 := '';
  if MyCheckHexStr(S) then
  begin
    for i := 0 to len - 1 do
    begin
      input[i] := MyHexCharToByte(S[2 * i + 1], S[2 * i + 2]);
    end;
    DesDecrypt(DesKey, input, s1);
  end;
  Result := s1;
end;

function TfmProcDecrypt.Encrypt_DES(const S: string;
  DesKey: TDesKey): string;
var
  i, len: integer;
  s1, s2: string;
  output: array of Byte;
begin
  i := (length(S) + 7) div 8;
  setlength(output, i * 8);
  DesEncrypt(DesKey, S, Output);
  Result := ByteTostr(Output);
end;

function TfmProcDecrypt.getDES_XT: TDesKey;
var
  i: integer;
  DesKey: TDesKey;
begin
  for i := 1 to length(DesKey_C) do
  begin
    DesKey[i - 1] := ord(DesKey_C[i]);
  end;
  Result := DesKey;
end;

end.
