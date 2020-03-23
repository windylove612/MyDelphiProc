unit Unit1;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Registry, ExtCtrls,IPHlpApi, Iprtrmib, IpTypes, IpFunctions,
   TLHelp32, IPExport , WinSock, inifiles, IdBaseComponent,
  IdComponent, IdIPWatch, Mask;

type
  TForm1 = class(TForm)
    lbl1: TLabel;
    edt2: TEdit;
    lbl2: TLabel;
    btn1: TButton;
    btn2: TButton;
    lbl3: TLabel;
    lbl4: TLabel;
    lbl5: TLabel;
    lbl6: TLabel;
    edt3: TEdit;
    edt4: TEdit;
    edt5: TEdit;
    edt6: TEdit;
    txt1: TStaticText;
    cbb1: TComboBox;
    lbl7: TLabel;
    edt1: TEdit;
    lbl8: TLabel;
    edt7: TEdit;
    chk1: TCheckBox;
    txt2: TStaticText;
    lbl9: TLabel;
    cbb2: TComboBox;
    procedure cbb1Change(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btn2Click(Sender: TObject);
    procedure btn1Click(Sender: TObject);
//    procedure btn3Click(Sender: TObject);
//    procedure btn3Click(Sender: TObject);
  private
    procedure GetIPPart;
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;
  CardList: Tstringlist;

implementation

uses Unit2;

{$R *.dfm}

//重启机器

procedure PcReBoot;
var
  hdlProcessHandle: Cardinal;
  hdlTokenHandle: Cardinal;
  tmpLuid: Int64;
  tkp: TOKEN_PRIVILEGES;
  tkpNewButIgnored: TOKEN_PRIVILEGES;
  lBufferNeeded: Cardinal;
  Privilege: array[0..0] of _LUID_AND_ATTRIBUTES;
begin
  hdlProcessHandle := GetCurrentProcess;
  OpenProcessToken(hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES or TOKEN_QUERY), hdlTokenHandle);
  LookupPrivilegeValue('', 'SeShutdownPrivilege', tmpLuid);
  Privilege[0].Luid := tmpLuid;
  Privilege[0].Attributes := SE_PRIVILEGE_ENABLED;
  tkp.PrivilegeCount := 1; // One privilege to set
  tkp.Privileges[0] := Privilege[0];
  AdjustTokenPrivileges(hdlTokenHandle,
    False,
    tkp,
    Sizeof(tkpNewButIgnored),
    tkpNewButIgnored,
    lBufferNeeded);
  ExitWindowsEx((EWX_REBOOT or EWX_FORCE), $FFFF);
end;



procedure TForm1.cbb1Change(Sender: TObject);
var
  sServiceRegKey, sNetCardDriverName, sTCPIPRegKey: string;
  regRootKey: TRegistry;
  buf: PChar;

begin

  regRootKey := TRegistry.Create;
  sNetCardDriverName := CardList.Strings[cbb1.ItemIndex];
  sServiceRegKey := '\SYSTEM\CurrentControlSet\Services\';
  sTCPIPRegKey := sServiceRegKey + sNetCardDriverName + '\Parameters\Tcpip';
  regRootKey.RootKey := HKEY_LOCAL_MACHINE;
  if regRootKey.OpenKeyReadOnly('\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName') then begin
    edt3.Text := regRootKey.ReadString('ComputerName');
  end;
  if regRootKey.OpenKeyReadOnly('\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\' + sNetCardDriverName) then begin
    GetMem(buf, 255);

    regRootKey.ReadBinaryData('IPAddress', buf^, 255);
    edt4.Text := StrPas(buf);

    regRootKey.ReadBinaryData('SubNetMask', buf^, 255);
    Edt5.Text := StrPas(buf);

    regRootKey.ReadBinaryData('DefaultGateway', buf^, 255);
    Edt1.Text := StrPas(buf);
    FreeMem(buf);
    Edt6.Text := regRootKey.ReadString('NameServer');
    if regRootKey.ReadInteger('EnableDHCP') = 1 then
      chk1.Checked := True
    else
      chk1.Checked := False;
  end;
  regRootKey.CloseKey;
  regRootKey.RootKey := HKEY_LOCAL_MACHINE;
  regRootKey.OpenKey('SYSTEM\CurrentControlSet\Control\Network\{4D36E972-E325-11CE-BFC1-08002BE10318}\' + sNetCardDriverName + '\Connection\',True);
  Edt7.Text := regRootKey.ReadString('Name');
  regRootKey.Free;

end;

procedure TForm1.FormShow(Sender: TObject);
var
  PAdapter: PipAdapterInfo;
  OutBufLen: ULONG;
  T3 : string;
begin
  try
  T3 := inifile.ReadString('配置','ipauto','');
  CardList := Tstringlist.Create;
  CardList.Clear;
  cbb1.Items.Clear;
  VVGetAdaptersInfo(PAdapter, OutBufLen);
    while PAdapter <> nil do begin
    CardList.Add(Trim(PAdapter.AdapterName));
    cbb1.Items.Add(Trim(PAdapter.Description));
     PAdapter := PAdapter.Next;
    end;
    cbb1.ItemIndex := 0 ;   //网卡读取完毕
    except
      Application.MessageBox('读取网卡信息失败，请检查网卡是否存在或启用！',
        '提示', MB_OK + MB_ICONINFORMATION);
  end;
  cbb1Change(Self); //显示首选信息
  GetIPPart;
  if T3 = 'NO' then
  begin
    Self.cbb2.Enabled := False;
  end;
end;

procedure WriteIp(Key: Hkey; Name, Value: string);
var
  I, j: Integer;
  buf: array[0..255] of Char;
begin
  Value := Value + #0;
  j := Length(Value) + 1;
  for I := 0 to j - 1 do
    buf[I] := Value[I + 1];
  RegSetValueEx(Key, PChar(Name), 0, REG_MULTI_SZ, @buf, j);
end;

{ 字符转化为16进制数据 }
function strtohex(v:string):integer;
const HEX:array['A'..'F'] of integer=(10,11,12,13,14,15);
var
int,i:integer;
begin
Int:=0;
for i:=1 to Length(v) do
if v[i] < 'A' then Int:=Int*16+ord(v[i])-48
else Int:=Int*16+HEX[v[i]];
strtohex:=int;
end;



procedure TForm1.btn2Click(Sender: TObject);
begin
Form2.ShowModal;
GetIPPart;
end;

procedure TForm1.btn1Click(Sender: TObject);
var
  regRootKey: TRegistry;
  I,j,T5: Integer;
  T,T1,T2,T3,T4:string;
  str: string;
  tmp: TStringList;
  tmp1: string;
begin
//读取配置文件中的参数
  Form2.edt1.Text := inifile.ReadString('配置','计算机名','');
  Form2.edt2.Text := inifile.ReadString('配置','IP段','');
  Form2.edt3.Text := inifile.ReadString('配置','子网掩码','');
  Form2.edt4.Text := inifile.ReadString('配置','默认网关','');
  Form2.edt5.Text := inifile.ReadString('配置','DNS','');
  T2 := inifile.ReadString('配置','DHCP','');
  T3 := inifile.ReadString('配置','ipauto','');
  T4 := inifile.ReadString('配置','IP偏移','');
  T5 := inifile.ReadInteger('配置','偏移',0);
  I := 0;
//读取完毕
//判断输入是否合法
  if  Trim(Self.edt2.Text) = '' then
  begin
      Application.MessageBox('自动编号框不能为空，请输入对应的计算机编号！', '警告',
        MB_OK + MB_ICONSTOP);
        Exit;
  end;

 try
  try
    j := StrToInt(edt2.Text);
  except
     Application.MessageBox('计算名只能为数字，请重新输入！', '警告', MB_OK +
       MB_ICONSTOP);
     edt2.Text := '';
     edt2.SetFocus;
     Exit;
  end;
//判断输入是否合法结束

  if Self.cbb2.Enabled = True then
    begin
    if Self.cbb2.Text = '' then
     begin
       Application.MessageBox('你未手动选择对应的网段功能,请重新选择！', '提示',
     MB_OK + MB_ICONINFORMATION);
     Exit;
     end;
    end;
  tmp := TStringList.Create;
  str := inifile.ReadString('配置','IP段','');
  if cbb2.Text <> '' then
   begin
     tmp.Text := StringReplace(str, '.', #13, [rfReplaceAll]);
     tmp1 := tmp.Strings[0] + '.' + tmp.Strings[1] + '.' + cbb2.Text + '.';
   end;

//导入CDKEY
  if j > 254 then
    j := j mod 254;
  T := 'CDKEY'+ IntToStr(j);
  T := inifile.ReadString('CSKEY',T,'');
  T1 := inifile.ReadString('配置','CSKEY','');

  regRootKey := TRegistry.Create;
  regRootKey.RootKey := HKEY_CURRENT_USER;
  if T1 = 'Yes'   then
  begin
    regRootKey.OpenKey('Software\Valve\CounterStrike\Settings\',true);
    regRootKey.WriteString('Key',T);
    regRootKey.CloseKey;
  end;
//导入CDKEY结束

//修改计算机名称开始
  regRootKey.CloseKey;
  regRootKey.RootKey := HKEY_LOCAL_MACHINE;
  if regRootKey.OpenKey('\SYSTEM\ControlSet001\Control\ComputerName\ComputerName\', True) then
   begin
    regRootKey.WriteString('ComputerName',Form2.edt1.Text + Self.edt2.Text);
     if regRootKey.OpenKey('\SYSTEM\ControlSet001\Services\Tcpip\Parameters', True) then
      begin
        regRootKey.WriteString('NV Hostname',  Form2.edt1.Text + Self.edt2.Text);
       if regRootKey.OpenKey('\SYSTEM\CurrentControlSet\Control\ComputerName\ComputerName', True) then
        begin
          regRootKey.WriteString('ComputerName', Form2.edt1.Text + Self.edt2.Text);
          if regRootKey.OpenKey('\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters', True) then
            begin
              regRootKey.WriteString('NV Hostname',Form2.edt1.Text + Self.edt2.Text);
            end;
        end;
      end;
  end;
//修改计算机名结束

//修改IP开始
  if regRootKey.OpenKey('\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\' + CardList.Strings[cbb1.ItemIndex], True) then
  begin
    if T3= 'NO' then
    begin
      if T4 = 'Yes'  then
      begin
      T5:=StrToInt(edt2.Text) + T5;
      regRootKey.WriteString('NameServer', Form2.edt5.Text);
      WriteIp(regRootKey.CurrentKey,'IPAddress',Trim(Form2.edt2.Text + IntToStr(T5)));
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway', Form2.edt4.Text);
      WriteIp(regRootKey.CurrentKey, 'SubNetMask', Form2.edt3.Text);
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway', Form2.edt4.Text);
      end;
     if  T4 = 'NO'  then
     begin
      regRootKey.WriteString('NameServer', Form2.edt5.Text);
      WriteIp(regRootKey.CurrentKey,'IPAddress',Form2.edt2.Text + Self.edt2.Text);
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway', Form2.edt4.Text);
      WriteIp(regRootKey.CurrentKey, 'SubNetMask', Form2.edt3.Text);
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway', Form2.edt4.Text);
    end;
   end;
    if T3 = 'Yes' then
     begin
     if T4 = 'Yes' then
     begin
      T5:=StrToInt(edt2.Text) + T5;
      regRootKey.WriteString('NameServer', Form2.edt5.Text);
      WriteIp(regRootKey.CurrentKey,'IPAddress',Trim(tmp1+ IntToStr(T5)));
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway', Form2.edt4.Text);
      WriteIp(regRootKey.CurrentKey, 'SubNetMask', Form2.edt3.Text);
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway', Form2.edt4.Text);
     end;
     if T4 = 'NO' then
     begin
      regRootKey.WriteString('NameServer', Form2.edt5.Text);
      WriteIp(regRootKey.CurrentKey,'IPAddress',tmp1 + Self.edt2.Text);
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway', Form2.edt4.Text);
      WriteIp(regRootKey.CurrentKey, 'SubNetMask', Form2.edt3.Text);
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway', Form2.edt4.Text);
     end;
   end;
   i := i+ 1;
   regRootKey.CloseKey;
  end;

// 判断是否启用DHCP
  regRootKey.RootKey := HKEY_LOCAL_MACHINE;
  if T2 = 'Yes'  then
  begin
    regRootKey.OpenKey('\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\' + CardList.Strings[cbb1.ItemIndex], True);
    regRootKey.WriteInteger('EnableDHCP',1);
    regRootKey.WriteString('NameServer','');
    WriteIp(regRootKey.CurrentKey, 'IPAddress', '0.0.0.0');
    WriteIp(regRootKey.CurrentKey, 'SubNetMask','0.0.0.0');
    WriteIp(regRootKey.CurrentKey, 'DefaultGateway','');
    regRootKey.CloseKey;
    end;
//判断IPX同步更新是否开启，如开启则修改，否则跳出
     T1 := inifile.ReadString('配置','IPX','');
     regRootKey.RootKey := HKEY_LOCAL_MACHINE;
 if T1 = 'Yes' then
 begin
   regRootKey.OpenKey('SYSTEM\CurrentControlSet\Services\NwlnkIpx\Parameters',True);
   regRootKey.WriteInteger('VirtualNetworkNumber',strtohex(Self.edt2.Text));
   regRootKey.CloseKey;
 end;
//IPX结束
  if  I = 1  then
  begin
    if Application.MessageBox('恭喜，IP修改完成，按确定后将重启计算机,否则请取消！',
       '提示', MB_OKCANCEL + MB_ICONINFORMATION) = IDOK then
     begin
      PcReBoot;
     end else
     begin
      regRootKey.CloseKey;
      regRootKey.Free;
      Exit;
     end;
 end  else
Application.MessageBox('出现错误！请检查注册表项是否可写，或输入的编号是否正确！', 
  '警告', MB_OK + MB_ICONSTOP);
  regRootKey.CloseKey;
  regRootKey.Free;
  finally
end;
end;

function GetIp(str: string): string;
var
  i: Integer;
  R :string;
begin
  i := pos('.',str);
  r:=str;
  if i > 0 then
  begin
    str := copy(str,i + 1, length(str));
    i := pos('.',str);
    if i > 0 then
    begin
      str := copy(str,i + 1, length(str));
      i := pos('.',str);
      if i > 0 then
      begin
        str := copy(str,1,i -1);
      end;
    end;
  end;
   Result := str;
end;


{读IP段}
procedure TForm1.GetIPPart;
var
  R,R1 : string;
  T1 : Integer;
  ii,ij,it: Integer;
begin
  cbb2.Items.Clear;
  R := GetIp(inifile.ReadString('配置','IP段',''));
  R1 := GetIp(inifile.ReadString('配置','EndIP',''));

  ii := StrToIntDef(r,0);
  ij := StrToIntDef(R1,0);

  R := IntToStr(ii);
  if ii > ij then
  begin
    it := ij;
    ij := ii;
    ii := it;
  end;
//ShowMessage(Format('%d/%d',[ii,ij]));
  for it := ii -1 to ij -1 do
  begin
    cbb2.Items.Add(IntToStr(it + 1));
  end;
end;
{结束}


end.

