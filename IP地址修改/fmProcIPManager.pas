unit fmProcIPManager;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, DB, DBTables, Grids, DBGrids, RzDBGrid, StdCtrls, RzCmboBx,
  Mask, RzEdit, RzLabel, ExtCtrls, RzPanel, RzButton, BMPBTN,LGetAdapterInfo,IniFiles,IpHlpApi,
  Registry;

type
  TfmIPManager = class(TForm)
    RzPanel1: TRzPanel;
    RzLabel1: TRzLabel;
    RzLabel2: TRzLabel;
    edtMC: TRzEdit;
    ComWK: TRzComboBox;
    RzLabel3: TRzLabel;
    edtIP: TRzEdit;
    RzLabel4: TRzLabel;
    edtYM: TRzEdit;
    RzLabel5: TRzLabel;
    edtWG: TRzEdit;
    RzLabel6: TRzLabel;
    edtMS: TRzEdit;
    RzDBGrid1: TRzDBGrid;
    RzPanel2: TRzPanel;
    tblDZ: TTable;
    dsDZ: TDataSource;
    tblDZMC: TStringField;
    tblDZWKMC: TStringField;
    tblDZIP: TStringField;
    tblDZWG: TStringField;
    tblDZYM: TStringField;
    tblDZDNS: TStringField;
    btnAdd: TBmpBtn;
    btnXG: TBmpBtn;
    btnDel: TBmpBtn;
    btnSave: TBmpBtn;
    btnCancel: TBmpBtn;
    btnClose: TBmpBtn;
    BmpBtn6: TBmpBtn;
    RzLabel7: TRzLabel;
    edtBY: TRzEdit;
    tblDZBYDNS: TStringField;
    procedure FormCreate(Sender: TObject);
    procedure btnAddClick(Sender: TObject);
    procedure btnXGClick(Sender: TObject);
    procedure btnDelClick(Sender: TObject);
    procedure btnCancelClick(Sender: TObject);
    procedure btnCloseClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure dsDZDataChange(Sender: TObject; Field: TField);
    procedure BmpBtn1Click(Sender: TObject);
    procedure BmpBtn6Click(Sender: TObject);
  private
    { Private declarations }
    sFlag:integer;
    procedure WriteLog(sLog: string);
    procedure WriteLog_ALL(sLog: string);
    procedure LoadConfig;
    procedure EnableButton(val:boolean);
    procedure WriteIp(Key: Hkey; Name, Value: string);
  public
    { Public declarations }
  end;

var
  fmIPManager: TfmIPManager;

implementation

{$R *.dfm}

procedure TfmIPManager.FormCreate(Sender: TObject);
var
  i :Integer;
  aa:TList;
begin
  tblDZ.TableName:='xx';
  tblDZ.CreateTable;
  tblDZ.Open;
  sFlag:=0;
  //找网卡
  ComWK.Items.Clear;
  ComWK.Values.Clear;
  aa :=GetAdapterInfo;
  for i := 0 to aa.Count - 1 do
  begin
    ComWK.Items.Add(TAdapterInfo(aa.Items[i]).Description);
    ComWK.Values.Add(TAdapterInfo(aa.Items[i]).AdapterName);
  end;
  EnableButton(true);
  //取已存文件
  LoadConfig;
end;


procedure TfmIPManager.btnAddClick(Sender: TObject);
begin
  tblDZ.Append;
  sFlag:=0;
  EnableButton(false);
end;

procedure TfmIPManager.btnXGClick(Sender: TObject);
begin
  tblDZ.Edit;
  sFlag:=1;
  EnableButton(false);
end;

procedure TfmIPManager.btnDelClick(Sender: TObject);
var
  sNR:string;
begin
  if not tblDZ.Active then exit;
  if tblDZ.RecordCount<=0 then exit;
  if Application.MessageBox(Pchar('是否确定删除？'),Pchar('注意'),MB_YESNO)<>IDYES then
    exit;
  tblDZ.Delete;
  tblDZ.Close;
  tblDZ.Open;
  sNR:='';
  tblDZ.First;
  while not tblDZ.Eof do
  begin
    sNR:=sNR+'['+edtMC.Text+']'+#13#10;
    sNR:=sNR+'网卡='+ComWK.Items[ComWK.ItemIndex]+#13#10;
    sNR:=sNR+'IP地址='+edtIP.Text+#13#10;
    sNR:=sNR+'子网掩码='+edtYM.Text+#13#10;
    sNR:=sNR+'默认网关='+edtWG.Text+#13#10;
    sNR:=sNR+'首选DNS='+edtMS.Text+#13#10;
    sNR:=sNR+'备用DNS='+edtBY.Text+#13#10;
    sNR:=sNR+#13#10;
    tblDZ.Next;
  end;
  WriteLog_ALL(sNR);
  LoadConfig;
  EnableButton(true);
  sFlag:=0;
end;

procedure TfmIPManager.btnCancelClick(Sender: TObject);
begin
  tblDZ.Cancel;
  sFlag:=0;
  EnableButton(true);
end;

procedure TfmIPManager.btnCloseClick(Sender: TObject);
begin
  close;
end;

procedure TfmIPManager.btnSaveClick(Sender: TObject);
var
  sNR:string;
  sSecList: TStringList;
  sSection:string;
  Ini: TIniFile;
  i:integer;
begin
  if edtMC.Text='' then
  begin
    showmessage('请输入名称！');
    exit;
  end;
  if ComWK.ItemIndex<0 then
  begin
    showmessage('请选择网卡！');
    exit;
  end;
  {if edtIP.Text='' then
  begin
    showmessage('请输入IP地址！');
    exit;
  end;
  if edtYM.Text='' then
  begin
    showmessage('请输入子网掩码！');
    exit;
  end;
  if edtWG.Text='' then
  begin
    showmessage('请输入网关！');
    exit;
  end;  }
  ini := TIniFile.Create('.\Config.Ini');
  sSecList := TStringList.Create();
  Ini.ReadSections(sSecList);
  try
    for i:=0 to sSecList.Count-1 do
    begin
      sSection:=sSecList[i];
      if edtMC.Text=sSection then
      begin
        showmessage('名称重复！');
        exit;
      end;
    end;
  finally
    Ini.Free;
  end;

  if sFlag=0 then
  begin
    tblDZ.Edit;
    tblDZMC.AsString:=edtMC.Text;
    tblDZWKMC.AsString:=ComWK.Items[ComWK.ItemIndex];
    tblDZIP.AsString:=edtIP.Text;
    tblDZWG.AsString:=edtWG.Text;
    tblDZYM.AsString:=edtYM.Text;
    tblDZDNS.AsString:=edtMS.Text;
    tblDZBYDNS.AsString:=edtBY.Text;
    tblDZ.Post;
    sNR:='';
    sNR:=sNR+'['+edtMC.Text+']'+#13#10;
    sNR:=sNR+'网卡='+ComWK.Items[ComWK.ItemIndex]+#13#10;
    sNR:=sNR+'IP地址='+edtIP.Text+#13#10;
    sNR:=sNR+'子网掩码='+edtYM.Text+#13#10;
    sNR:=sNR+'默认网关='+edtWG.Text+#13#10;
    sNR:=sNR+'首选DNS='+edtMS.Text+#13#10;
    sNR:=sNR+'备用DNS='+edtBY.Text+#13#10;
    sNR:=sNR+#13#10;
    WriteLog(sNR);
  end
  else
  begin
    sNR:='';
    tblDZMC.AsString:=edtMC.Text;
    tblDZWKMC.AsString:=ComWK.Values[ComWK.ItemIndex];
    tblDZIP.AsString:=edtIP.Text;
    tblDZWG.AsString:=edtWG.Text;
    tblDZYM.AsString:=edtYM.Text;
    tblDZDNS.AsString:=edtMS.Text;
    tblDZBYDNS.AsString:=edtBY.Text;
    tblDZ.Post;
    tblDZ.First;
    while not tblDZ.Eof do
    begin
      sNR:=sNR+'['+tblDZMC.AsString+']'+#13#10;
      sNR:=sNR+'网卡='+tblDZWKMC.AsString+#13#10;
      sNR:=sNR+'IP地址='+tblDZIP.AsString+#13#10;
      sNR:=sNR+'子网掩码='+tblDZYM.AsString+#13#10;
      sNR:=sNR+'默认网关='+tblDZWG.AsString+#13#10;
      sNR:=sNR+'首选DNS='+tblDZDNS.AsString+#13#10;
      sNR:=sNR+'备用DNS='+tblDZBYDNS.AsString+#13#10;
      sNR:=sNR+#13#10;
      tblDZ.Next;
    end;
    WriteLog_ALL(sNR);
  end;
  LoadConfig;
  EnableButton(true);
  sFlag:=0;
end;

procedure TfmIPManager.FormDestroy(Sender: TObject);
begin
  if not tblDZ.Exists then exit;
  if tblDZ.Active then
    tblDZ.Active:=false;
  tblDZ.DeleteTable;
end;

procedure TfmIPManager.WriteLog(sLog: string);
var
  LogFile: TextFile;
  sFileName,str: string;
begin
  sFileName := '.\Config.Ini';
  try
    str := sLog;
    AssignFile(LogFile, sFileName);
    if FileExists(sFileName) then
      Append(LogFile)
    else
      Rewrite(LogFile);
    Writeln(LogFile, str);
    Flush(LogFile);
    CloseFile(LogFile);
  finally
  end;
end;

procedure TfmIPManager.LoadConfig;
var
  Ini: TIniFile;
  sSecList: TStringList;
  sSection:string;
  i:integer;
begin
  ini := TIniFile.Create('.\Config.Ini');                //程序配置文件

  sSecList := TStringList.Create();
  Ini.ReadSections(sSecList);

  tblDZ.Close;
  tblDZ.EmptyTable;
  tblDZ.Open;

  try
    for i:=0 to sSecList.Count-1 do
    begin
      sSection:=sSecList[i];
      tblDZ.Append;
      tblDZMC.AsString:=sSection;
      tblDZWKMC.AsString:=Ini.ReadString(sSection, '网卡', '');
      tblDZIP.AsString:=Ini.ReadString(sSection, 'IP地址', '');
      tblDZYM.AsString:=Ini.ReadString(sSection, '子网掩码', '');
      tblDZWG.AsString:=Ini.ReadString(sSection, '默认网关', '');
      tblDZDNS.AsString:=Ini.ReadString(sSection, '首选DNS', '');
      tblDZBYDNS.AsString:=Ini.ReadString(sSection, '备用DNS', '');
      tblDZ.Post;
    end;
  finally
    Ini.Free;
  end;
end;

procedure TfmIPManager.dsDZDataChange(Sender: TObject; Field: TField);
begin
  if not btnAdd.Enabled then exit;
  edtMC.Text:=tblDZMC.AsString;
  ComWK.ItemIndex:=ComWK.Items.IndexOf(tblDZWKMC.AsString);
  edtIP.Text:=tblDZIP.AsString;
  edtWG.Text:=tblDZWG.AsString;
  edtYM.Text:=tblDZYM.AsString;
  edtMS.Text:=tblDZDNS.AsString;
  edtBY.Text:=tblDZBYDNS.AsString;
end;

procedure TfmIPManager.EnableButton(val: boolean);
begin
  btnAdd.Enabled:=val;
  btnXG.Enabled:=val;
  btnDel.Enabled:=val;
  btnSave.Enabled:=not val;
  btnCancel.Enabled:=not val;
  edtMC.ReadOnly:=val;
  ComWK.ReadOnly:=val;
  edtIP.ReadOnly:=val;
  edtYM.ReadOnly:=val;
  edtWG.ReadOnly:=val;
  edtMS.ReadOnly:=val;
  edtBY.ReadOnly:=val;
end;

procedure TfmIPManager.WriteLog_ALL(sLog: string);
var
  LogFile: TextFile;
  sFileName,str: string;
begin
  sFileName := '.\Config.Ini';
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

procedure TfmIPManager.BmpBtn1Click(Sender: TObject);
var
  NetShTxt:string;
begin
  if tblDZ.RecordCount<=0 then exit;
  if edtIP.Text='' then
  begin
    NetShTxt := '';
        winexec(pchar(NetShTxt), sw_hide);
  end;
end;

procedure TfmIPManager.BmpBtn6Click(Sender: TObject);
var
  regRootKey: TRegistry;
  sNameServer:string;
  i:integer;
begin
  //修改IP开始      //判断是否启用DHCP
  if tblDZBYDNS.IsNull then
    sNameServer:=tblDZDNS.AsString
  else
    sNameServer:=tblDZDNS.AsString + ',' +tblDZBYDNS.AsString;

  i:=0;  
  regRootKey := TRegistry.Create;
  try
    if not tblDZIP.IsNull then
    begin
      regRootKey.RootKey := HKEY_LOCAL_MACHINE;
      if regRootKey.OpenKey('\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\' + ComWK.Values[ComWK.ItemIndex], True) then
      begin
        regRootKey.WriteInteger('EnableDHCP',0);
        regRootKey.WriteString('NameServer',sNameServer);
        WriteIp(regRootKey.CurrentKey,'IPAddress',tblDZIP.AsString);
        WriteIp(regRootKey.CurrentKey, 'DefaultGateway', tblDZWG.AsString);
        WriteIp(regRootKey.CurrentKey, 'SubNetMask', tblDZYM.AsString);
        regRootKey.CloseKey;
      end;
    end
    else
    begin
      regRootKey.RootKey := HKEY_LOCAL_MACHINE;
      regRootKey.OpenKey('\SYSTEM\CurrentControlSet\Services\Tcpip\Parameters\Interfaces\' + ComWK.Values[ComWK.ItemIndex], True);
      regRootKey.WriteInteger('EnableDHCP',1);
      regRootKey.WriteString('NameServer','');
      WriteIp(regRootKey.CurrentKey, 'IPAddress', '0.0.0.0');
      WriteIp(regRootKey.CurrentKey, 'SubNetMask','0.0.0.0');
      WriteIp(regRootKey.CurrentKey, 'DefaultGateway','');
      regRootKey.CloseKey;
    end;
    i := i+ 1;
    if i = 1 then
      showmessage('恭喜，IP修改完成！')
    else
      showmessage('出现错误！请检查注册表项是否可写，或输入的编号是否正确！');
  finally
    regRootKey.Free;
  end;
end;

procedure TfmIPManager.WriteIp(Key: Hkey; Name, Value: string);
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

end.
