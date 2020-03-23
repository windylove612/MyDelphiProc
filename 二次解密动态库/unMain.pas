unit unMain;

interface

uses StrUtils,SysUtils,unDES,dialogs;

const
  DesKey_C = 'PdfSsoDi';

  function GetDecrypt_DES(pCODE:PChar;var pREAL,pMsg:PChar):boolean;stdcall;
  //内部函数
  function getDES_XT: TDesKey;
  function Encrypt_DES(const S: string; DesKey: TDesKey): string;
  function Decrypt_DES(const S: string; DesKey: TDesKey): string;
  function CardEncrypt_DES(const S: string; DesKey: TDesKey): string;
  function CardDecrypt_DES(const S: string; DesKey: TDesKey): string;
  procedure Log(pStr: string);
  function GetCompactDate(pRQ: TDateTime): string;
  procedure InitPChar(var pPch: PChar; iLen: Integer);

implementation

function GetDecrypt_DES(pCODE:PChar;var pREAL,pMsg:PChar):boolean;stdcall;
begin
  Result := False;
  InitPChar(pMsg, 100);
  InitPChar(pREAL, 100);
  try
    if pCODE='' then
    begin
      StrPCopy(pMsg,'验证码不能为空！');
      Log('GetDecrypt_DES 错误:'+pMsg);
      Result := False;
      exit;
    end;
    StrPCopy(pREAL,CARDDecrypt_DES(pCODE,getDES_XT));
    StrPCopy(pMsg,'成功');
    Result := true;
  except
    on E: Exception do
    begin
      StrPCopy(pMsg, '操作失败，' + E.Message );
      Log('GetDecrypt_DES 错误:' + pMsg);
      Result := False;
    end;
  end;
end;

procedure Log(pStr: string);
var
  dat: TextFile;
  sFile: string;
begin
  sFile := GetCompactDate(Date) + '.log';
  try
    if not DirectoryExists(getCurrentDir + '\log') then
      MkDir(getCurrentDir + '\log');
    Assignfile(dat, getCurrentDir + '\log\' + sFile);
    if FileExists(getCurrentDir + '\log\' + sFile) then
      Append(dat)
    else
      Rewrite(dat);
    Writeln(dat, DateTimeToStr(Now) + '：' + pStr);
  finally
    Closefile(dat);
  end;
end;

function GetCompactDate(pRQ: TDateTime): string;
begin
  Result := FormatDateTime('yyyymmdd', pRQ);
end;

function getDES_XT: TDesKey;
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

function Encrypt_DES(const S: string; DesKey: TDesKey): string;
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

function Decrypt_DES(const S: string; DesKey: TDesKey): string;
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

function CardEncrypt_DES(const S: string; DesKey: TDesKey): string;
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

function CardDecrypt_DES(const S: string; DesKey: TDesKey): string;
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

procedure InitPChar(var pPch: PChar; iLen: Integer);
begin
  pPch := stralloc(iLen);
end;

end.
