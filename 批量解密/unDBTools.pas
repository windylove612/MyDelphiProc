unit unDBTools;

{$IF NOT DEFINED(ORACLE) and NOT DEFINED(SYBASE)}
  请定义使用的数据库系统：ORACLE 或 SYBASE;
{$IFEND}
interface

uses
  DB, DBTables,StdCtrls;

const
  {$IF DEFINED(ORACLE)}
  DefaultDSNName = 'BFV800.CYC.COM';            //数据库配置文件连接名称
  {$ELSEIF DEFINED(SYBASE)}
  DefaultDSNName = 'BFV800';                    //数据库配置文件连接名称
  {$IFEND}
  LocalSettingSection = 'LocalSetting';
  IniFileBFSYSTEM     = 'BFSYSTEM.INI';
  IniFileBFLOCAL      = 'BFLOCAL.INI';

function GetCurrentDBSection() :string;
function GetDatabaseDefine(const ServerSection:string; var sServerName, sServerAddress,sDatabaseName,sUserName,sServerAddress1:string) :Boolean;
procedure FillDatabaseList(const ComboBox:TCustomComboBox);
procedure SaveCurDatabaseSel(const sSelDBSection:string);
procedure OpenDatabase(const DB:TDatabase; const ServerName,DatabaseName,UserName,Password:string; const bLargeBlobSize:Boolean=False);
procedure SetDBConfigureFile(const ServerName,sServerAddress:string;sServerAddress1:string='');

implementation

uses
  SysUtils, Classes, Forms, Controls, WinTypes,IniFiles,registry, StrUtils;

const
  PrevSelServer       = 'Select Server';
  identDSNName = 'DSN';
  identServerConnectString = 'SERVER';
  identServerConnectString1 = 'SERVER1';
  identDatabaseName   = 'DATABASE';
  identUserName       = 'USER';

function GetCurrentDBSection():string;
begin
  Result := '';
  with TIniFile.Create(IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName))+IniFileBFLOCAL) do
  begin
    try
      Result := ReadString(LocalSettingSection,PrevSelServer,'');
    finally
      Free;
    end;
  end;
end;

{-------------------------------------------------------------------------------
   用可选择的数据库填充组合框
-------------------------------------------------------------------------------}
procedure FillDatabaseList(const ComboBox:TCustomComboBox);
var
  ini     :TiniFile;
  tmpstr  :string;
  i       :integer;
  S       :TStringList;
begin
  S := TStringList.Create;
  try
    ini := TIniFile.Create(IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName))+IniFileBFSYSTEM);   //程序配置文件
    try
      Ini.ReadSections(S);
    finally
      ini.Free;
    end;
    for i:=0 to S.Count-1 do
      if Uppercase(S[i])<>'DOWNLOAD' then
        ComboBox.Items.Add(S[i]);
  finally
    S.Free;
  end;

  tmpstr := GetCurrentDBSection();
  if Length(tmpstr)>0 then
  begin
    for i:=0 to ComboBox.Items.Count-1 do
    begin
      if ComboBox.Items[i] = tmpstr then
      begin
        ComboBox.ItemIndex := i;
        break;
      end;
    end;
  end;
end;

procedure SaveCurDatabaseSel(const sSelDBSection:string);
var
  ini :TIniFile;
begin
  ini := TIniFile.Create(IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName))+IniFileBFLOCAL);
  try
    ini.WriteString(LocalSettingSection,PrevSelServer, sSelDBSection);
  finally
    ini.Free;
  end;
end;

procedure OpenDatabase(const DB:TDatabase; const ServerName,DatabaseName,UserName,Password:string; const bLargeBlobSize:Boolean);
begin
  DB.Close;
  DB.AliasName := '';
  {$IF DEFINED(ORACLE)}
  DB.DriverName := 'ORACLE';
  {$ELSEIF DEFINED(SYBASE)}
  DB.DriverName := 'SYBASE';
  {$IFEND}
  DB.Params.Clear;
  DB.Params.Add('USER NAME='+UserName);
  DB.Params.Add('PASSWORD='+Password);
  DB.Params.Add('SERVER NAME='+ServerName);
  {$IFDEF ORACLE}
  DB.Params.Add('ENABLE INTEGERS=TRUE');
  {$ELSEIF DEFINED(SYBASE)}
  DB.Params.Add('DATABASE NAME='+DatabaseName);
  {$IFEND}
  if bLargeBlobSize then
      DB.Params.Add('BLOB SIZE= 1000');
  DB.Open;
end;


function GetDatabaseDefine(const ServerSection:string; var sServerName,sServerAddress,sDatabaseName,sUserName,sServerAddress1:string) :Boolean;
  var
    ini :TIniFile;
    S   :string;
    i   :integer;
begin
  Result := False;
  ini := TIniFile.Create(IncludeTrailingPathDelimiter(ExtractFilePath(Application.ExeName))+IniFileBFSYSTEM);   //程序配置文件
  try
    sServerName := ini.ReadString(ServerSection, identDSNName, DefaultDSNName);
    sServerAddress := ini.ReadString(ServerSection, identServerConnectString, '');
    sServerAddress1 := ini.ReadString(ServerSection, identServerConnectString1, '');
    sDatabaseName := ini.ReadString(ServerSection, identDatabaseName,'');
    sUserName     := ini.ReadString(ServerSection, identUserName, sDatabaseName); //'BFAPP8');
    if (Length(sServerAddress)>0) and ((Length(sDatabaseName)>0) or (Length(sUserName)>0)) then
    begin
      Result := True;
    end;
  finally
    ini.Free;
  end;
end;

{$IFDEF ORACLE}
procedure SetDBConfigureFile(const ServerName, sServerAddress:string;sServerAddress1:string='');
var
  DSNName :string;
  sTCPID, sPORT, sSID :string;
  sTCPID1, sPORT1, sSID1 :string;
  Strings : TStringList;
  i, j : integer;

  CfgPath,CfgFile :string;
  S               :string;
  iLength         :integer;
  iCountLeft      :integer;  //左括号的数量
  iCountRight     :integer;  //左括号的数量
  LastHome:string;

  function GetOracle10Path :string;
  begin
  end;
begin
  CfgPath := '';
  CfgPath := GetEnvironmentVariable('TNS_ADMIN');
  if Length(CfgPath)>0 then
  begin
    CfgFile := IncludeTrailingPathDelimiter(CfgPath)+'tnsnames.ora';
    if not FileExists(CfgFile) then
      CfgPath := '';
  end;

  if Length(CfgPath)=0 then
  begin
    with TRegistry.Create do
    try
      RootKey := HKEY_LOCAL_MACHINE;
      //--检查是否为 Oracle 10g -----------
      if OpenKeyReadOnly('\SOFTWARE\ORACLE') then
        CfgPath := ReadString('VOBHOME2.0');

      if Length(CfgPath)=0 then
      begin
        for i:=1 to 9 do
          if OpenKeyReadOnly('\SOFTWARE\ORACLE\KEY_OraClient10g_home'+IntToStr(i)) then
          begin
            CfgPath := ReadString('ORACLE_HOME');
            break;
          end
      end;

      if Length(CfgPath)=0 then
      begin
        for i:=1 to 9 do
          if OpenKeyReadOnly('\SOFTWARE\ORACLE\KEY_OraDb10g_home'+IntToStr(i)) then
          begin
            CfgPath := ReadString('ORACLE_HOME');
            break;
          end
      end;

      if Length(CfgPath)=0 then  //Version before 10g
      begin
        LastHome := '0';
        if OpenKeyReadOnly('\SOFTWARE\ORACLE\ALL_HOMES') then
        begin
          LastHome := ReadString('LAST_HOME');
          if OpenKeyReadOnly('\SOFTWARE\ORACLE\ALL_HOMES\ID'+LastHome) then
            CfgPath := ReadString('PATH');
        end;
      end;
    finally
      free;
    end;
    if FileExists( CfgPath+'\network\admin\tnsnames.ora') then
      CfgFile := CfgPath+'\network\admin\tnsnames.ora'
    else
      CfgFile := CfgPath+'\net80\admin\tnsnames.ora';
  end;

  Strings := TStringList.Create;
  try
    Strings.CommaText := sServerAddress;
    if Strings.Count=3 then
    begin
      sTCPID := Strings[0];
      sPORT := Strings[1];
      sSID := Strings[2];
    end
    else
      raise Exception.Create(sServerAddress+'不是一个有效的数据库服务器地址');

    Strings.Clear;
    Strings.CommaText := sServerAddress1;
    if Strings.Count=2 then
    begin
      sTCPID1 := Strings[0];
      sPORT1 := Strings[1];
    end
    else
      raise Exception.Create(sServerAddress1+'不是一个有效的数据库服务器地址');

    Strings.Clear;
    Strings.LoadFromFile(CfgFile);
    DSNName := AnsiUpperCase(trim(ServerName));
    iLength := Length(DSNName);
    i := 0;
    while (i<Strings.Count) do
    begin
      S := AnsiUpperCase(Trim(Strings[i]));
      if (Length(S)>iLength) and (Copy(S,1,iLength)=DSNName) then
        break;
      Inc(i);
    end;
    if i<Strings.Count then       //已有定义，需要删除
    begin
      iCountLeft  := 0;  //左括号的数量
      iCountRight := 0;  //左括号的数量
      while i<Strings.Count do
      begin
        S := Strings[i];
        Strings.Delete(i);
        for j:=1 to Length(S) do
        begin
          if S[j]='(' then Inc(iCountLeft);
          if S[j]=')' then Inc(iCountRight);
        end;
        if (iCountLeft>0) and (iCountLeft=iCountRight) then
          break;
      end;
    end;

    //Strings.Add(DSNName + '.CYC.COM =');
    Strings.Add(DSNName + ' =');
    Strings.Add('  (DESCRIPTION =');
    //Strings.Add('    (ADDRESS_LIST=');
    Strings.Add('      (ADDRESS = (PROTOCOL = TCP)(HOST = '+sTCPID+')(PORT = '+sPORT+'))');
    if (sTCPID1 <>'') and (sPORT1 <>'') Then 
       Strings.Add('      (ADDRESS = (PROTOCOL = TCP)(HOST = '+sTCPID1+')(PORT = '+sPORT1+'))');
    Strings.Add('     (LOAD_BALANCE = yes) ');
   // Strings.Add('    )');
    Strings.Add('    (CONNECT_DATA =');
    Strings.Add('     (SERVER = DEDICATED)');
    Strings.Add('      (SERVICE_NAME = '+sSID+')');
    Strings.Add('  (FAILOVER_MODE =');
    Strings.Add('    (TYPE = SELECT)');
    Strings.Add('    (METHOD = BASIC)');
    Strings.Add('    (RETRIES = 180)');
    Strings.Add('    (DELAY = 5)');
    Strings.Add('      )');

    //Strings.Add('      (SID = '+sSID+')');       //考虑对8.0的支持，使用 SID，以后应该使用SERVICE_NAME (8.1以上版本)
    Strings.Add('    )');
    Strings.Add('  )');
    Strings.SaveToFile(CfgFile);
  finally
    Strings.Free;
  end;
end;
{$ENDIF}

{$IFDEF SYBASE}
procedure SetDBConfigureFile(const ServerName, sServerAddress:string);
var
  SybPath :string;
  iniSyb  :TIniFile;
  sConnectString :string;
begin
  SybPath := GetEnvironmentVariable('SYBASE');
  if Length(SybPath)=0 then
    with TRegistry.Create do
    begin
      try
        RootKey := HKEY_LOCAL_MACHINE;
        if OpenKeyReadOnly('\SOFTWARE\SYBASE\SETUP') then
          SybPath := ReadString('sybase');
      finally
        free;
      end;
    end;

  if Length(SybPath)=0 then
    SybPath := 'c:\sybase';

  sConnectString := trim(sServerAddress);
  if UpperCase(Copy(sConnectString,1,3))<>'TCP' then
    sConnectString := 'TCP,'+sConnectString;
  iniSyb := TIniFile.Create(SybPath+'\ini\SQL.INI');   //SYBASE配置文件
  try
    //重建SYBASE连接
    if iniSyb.SectionExists(ServerName) then
    begin
      iniSyb.DeleteKey(ServerName,'master');
      iniSyb.DeleteKey(ServerName,'query');
    end;
    iniSyb.WriteString(ServerName,'query',sConnectString);
    iniSyb.UpdateFile;
  finally
    iniSyb.Free;
  end;
end;
{$ENDIF}

end.

