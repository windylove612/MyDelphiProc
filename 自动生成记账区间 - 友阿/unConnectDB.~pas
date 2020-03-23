unit unConnectDB;

interface

uses Classes, Windows, SysUtils, Dialogs, DBTables, IniFiles, Registry;

procedure ConnectDB(DB: TDatabase; sIniName: string; sSection: string = '');

implementation

{
配置文件样例
[Sybase];Sybase方式，区分在于query有无TCP字段
query=TCP,192.168.111.4,5000
DATABASE=V65_ZD
USER=BFCRM
PASSWORD=1111111
[Oracle];标准Oracle方式，没有DATABASE
query=ORCL
DATABASE=
USER=BFCRM
PASSWORD=1111111
[ODBC];ODBC方式，支持DB2/Oracle
query=XAJHJXC_ZD
DATABASE=V75_ZD
USER=BFCRM
PASSWORD=111111
}

procedure ConnectDB(DB: TDatabase; sIniName: string; sSection: string = '');
var
  SybPath: string;
  Ini: TIniFile;
  sIP, sDBName, sUser, sPasswd, sAliasName: string;
  sSecList: TStringList;
begin
  ini := TIniFile.Create('.\' + sIniName);                  //程序配置文件
  if sSection = '' then
  begin
    sSecList := TStringList.Create();
    Ini.ReadSections(sSecList);
    sSection := sSecList[0];
  end;
  try
    sIP := Ini.ReadString(sSection, 'query', '');
    sDBName := Ini.ReadString(sSection, 'DATABASE', '');
    sUser := Ini.ReadString(sSection, 'USER', '');
    sPasswd := Ini.ReadString(sSection, 'PASSWORD', '');
  finally
    Ini.Free;
  end;
  DB.Connected := False;
  if pos('TCP', sIP) > 0 then                               //Sybase方式
  begin
    SybPath := 'c:\sybase';
    sAliasName := 'tmpConnectDB';
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
    try
      Ini := TIniFile.Create(SybPath + '\ini\SQL.INI');     //重建SYBASE连接
      if Ini.SectionExists(sAliasName) then
      begin
        Ini.DeleteKey(sAliasName, 'master');
        Ini.DeleteKey(sAliasName, 'query');
      end;
      Ini.WriteString(sAliasName, 'master', sIP);
      Ini.WriteString(sAliasName, 'query', sIP);
      Ini.UpdateFile;
    finally
      Ini.Free;
    end;
    DB.DriverName := 'SYBASE';
    DB.Params.Values['SERVER NAME'] := sAliasName;
    DB.Params.Values['DATABASE NAME'] := sDBName;
    DB.Params.Values['USER NAME'] := sUser;
    DB.Params.Values['PASSWORD'] := sPasswd;
  end
  else if sDBName = '' then
  begin                                                     //Oracle方式
    DB.Params.Clear();
    DB.DriverName := 'ORACLE';
    DB.Params.Values['USER NAME'] := sUser;
    DB.Params.Values['PASSWORD'] := sPasswd;
    DB.Params.Values['SERVER NAME']:=sIP;
//    DB.Params.Add('ENABLE INTEGERS=TRUE');
    DB.Params.Values['ENABLE INTEGERS']:='TRUE';
  end
  else                                                      //普通ODBC方式
  begin
    DB.Params.Clear();
    DB.AliasName := sIP;
    DB.Params.Values['ODBC DSN'] := sIP;
    DB.Params.Values['DATABASE NAME'] := sDBName;
    DB.Params.Values['USER NAME'] := sUser;
    DB.Params.Values['PASSWORD'] := sPasswd;
  end;
  try
    DB.Connected := True;
  except
    on E: Exception do
    begin
      ShowMessage('连接失败:' + #$D#$A + E.Message);
    end;
  end;
end;

end.

