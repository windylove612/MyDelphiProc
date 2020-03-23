unit ADO_QImport3Access;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.Classes,
    Data.Win.ADODB,
    Data.DB,
    System.IniFiles,
  {$ELSE}
    Classes,
    ADODb,
    DB,
    IniFiles,
  {$ENDIF}
  QImport3,
  QImport3StrTypes;

type
  TStringArray = Array Of String;
  TFieldTypes = Array Of TFieldType;
  TFieldSizes = Array Of Integer;
  TQImportAccessSourceType = (isTable, isSQL);
  TQImportAccessVersion = (avAccess97_2003, avAccess2007);

  TADO_QImport3Access = class(TQImport3)
  private
    FSQL: TStrings;
    FTableName: string;
    FPassword,
    FLogin: string;
    FSourceType: TQImportAccessSourceType;
    FADO: TADOQuery;
    FSkipCounter: integer;
    FVersion: TQImportAccessVersion;
    procedure SetSQL(const Value: TStrings);
    function GetConnectionString(): string;
  protected
    procedure StartImport; override;
    function CheckCondition: boolean; override;
    function Skip: boolean; override;
    procedure FillImportRow; override;
    function ImportData: TQImportResult; override;
    procedure ChangeCondition; override;
    procedure FinishImport; override;
    procedure DoLoadConfiguration(IniFile: TIniFile); override;
    procedure DoSaveConfiguration(IniFile: TIniFile); override;
    procedure DoAfterSetFileName; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    procedure GetTableNames(List: TqiStrings);{$IFDEF QI_UNICODE}overload;
    procedure GetTableNames(List: TStrings); overload;{$ENDIF}
    procedure GetFieldNames(var FieldNames, FieldTypesString: TStringArray; var
        FieldTypes: TFieldTypes; var FieldSizes: TFieldSizes);
  published
    property FileName;
    property SkipFirstRows default 0;
    property TableName: string read FTableName write FTableName;
    property SQL: TStrings read FSQL write SetSQL;
    property SourceType: TQImportAccessSourceType read FSourceType
      write FSourceType default isTable;
    property Login: String read FLogin write FLogin;
    property Password: string read FPassword write FPassword;
  end;

  function IsEncryptedAccessFile(const AccessFile: qiString): Boolean;

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
    System.Variants,
    System.Win.ComObj,
    Winapi.ActiveX,
    Vcl.Dialogs,
  {$ELSE}
    SysUtils,
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
    ComObj,
    ActiveX,
    Dialogs,
  {$ENDIF}
  QImport3Common;

const
  SelectFromTable = 'select * from [%s]';
  SelectFromTableFieldNames = 'select * from [%s] where 1 = 0';
  ConnectionString97_2003 = 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source=%s';
  PasswordString97_2003 = ';Jet OLEDB:Database Password=%s';
  LoginString = ';User ID=%s';

  ConnectionString2007 = 'Provider=Microsoft.ACE.OLEDB.12.0;Data Source=%s';
  PasswordString2007 = ';Password=%s';
  ExtAccess2007 = '.accdb';

function BuildConnectionString(const Version: TQImportAccessVersion;
  const UserName, Password: qiString): qiString;
begin
  if Version = avAccess97_2003 then
  begin
    Result := ConnectionString97_2003;
    if UserName <> EmptyStr then
      Result := Result + Format(LoginString, [UserName]);
    if Password <> EmptyStr then
      Result := Result + Format(PasswordString97_2003, [Password]);
  end
  else
    Result := ConnectionString2007 + Format(PasswordString2007, [Password]);
end;

function IsEncryptedAccessFile(const AccessFile: qiString): Boolean;
const
  ErrCode = -2147217843;
var
  Connection: TADOConnection;
  Version: TQImportAccessVersion;
begin
  Result := False;
  Connection := TADOConnection.Create(nil);
  try
    Connection.LoginPrompt := False;
    if ExtractFileExt(AccessFile) = ExtAccess2007 then
      Version := avAccess2007
    else
      Version := avAccess97_2003;

    Connection.ConnectionString := Format(BuildConnectionString(Version, '', ''),
      [AccessFile]);
    try
      Connection.Connected := True;
    except
      on E: EOleException do
        Result := E.ErrorCode = ErrCode;
    end;
  finally
    Connection.Free;
  end;
end;  

{ TADO_QImport3Access }

constructor TADO_QImport3Access.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
  FSourceType := isTable;
  FSQL := TStringList.Create;
  FPassword := EmptyStr;
  FLogin := EmptyStr;
  FVersion := avAccess97_2003;
end;

destructor TADO_QImport3Access.Destroy;
begin
  FSQL.Free;
  inherited;
end;

procedure TADO_QImport3Access.DoAfterSetFileName;
begin
  inherited;
  if ExtractFileExt(FileName) = ExtAccess2007 then
    FVersion := avAccess2007
  else
    FVersion := avAccess97_2003;
end;

procedure TADO_QImport3Access.StartImport;
var
  connStr: string;
begin
  connStr := GetConnectionString();

  FADO := TADOQuery.Create(nil);
  FADO.CursorType := ctOpenForwardOnly;
  FADO.CursorLocation := clUseServer;
  FADO.LockType := ltReadOnly;
  FADO.ConnectionString := Format(connStr, [FileName]);

  if SourceType = isSQL
    then FADO.SQL.Assign(SQL)
    else FADO.SQL.Text := Format(SelectFromTable, [TableName]);
  FADO.Open;

  FSkipCounter := SkipFirstRows;
  if FSkipCounter < 0 then FSkipCounter := 0;
end;

function TADO_QImport3Access.CheckCondition: boolean;
begin
  Result := not FADO.Eof;
end;

function TADO_QImport3Access.Skip: boolean;
begin
  Result := FSkipCounter > 0;
end;

procedure TADO_QImport3Access.FillImportRow;
var
  i, k: Integer;
  SField: TField;
{$IFDEF QI_UNICODE}
  fieldValue: Variant;
{$ENDIF}
  p: Pointer;
  mapValue: qiString;
  FieldValueBlob: qiString;
begin
  FImportRow.ClearValues;
  RowIsEmpty := True;
  for i := 0 to FImportRow.Count - 1 do
  begin
    if FImportRow.MapNameIdxHash.Search(FImportRow[i].Name, p) then
    begin
      k := Integer(p);
{$IFDEF VCL7}
      mapValue := Map.ValueFromIndex[k];
{$ELSE}
      mapValue := Map.Values[FImportRow[i].Name];
{$ENDIF}
      if Pos('=', mapValue) > 0 then
        mapValue := Copy(mapValue, 1, Pos('=', mapValue) - 1);
      SField := FADO.FindField(mapValue);
      if Assigned(SField) then
      begin
{$IFDEF QI_UNICODE}
        if SField.IsBlob then
        begin
          FieldValueBlob := FADO.FieldByName(mapValue).AsString;
          fieldValue := FieldValueBlob;
        end
        else
          fieldValue := FADO.Recordset.Fields[mapValue].Value;

        if IsCSV and (SField.DataType in [ftDate, ftTime, ftDateTime]) then //for MySQL
          fieldValue := FormatDateTime('yyyy-mm-dd hh:mm:ss', fieldValue);

        if VarIsNull(fieldValue) or VarIsClear(fieldValue) then
        begin
          fieldValue := '';
          if Formats.NullValues.Count = 0 then
            FImportRow.SetValue(Map.Names[k], fieldValue, SField.IsBlob)
          else
            FImportRow.SetValue(Map.Names[k], Formats.NullValues[0], SField.IsBlob);
        end
        else
        begin
          if AutoTrimValue and VarIsStr(fieldValue) then
            fieldValue := Trim(fieldValue);

          RowIsEmpty := RowIsEmpty and (VarToStr(fieldValue) = '');

          if SField.IsBlob then
            FImportRow.SetValue(Map.Names[k], FieldValueBlob, SField.IsBlob)
          else
            FImportRow.SetValue(Map.Names[k], fieldValue, SField.IsBlob);
        end;
{$ELSE}
        FImportRow.SetValue(Map.Names[k], SField.AsString, SField.IsBlob);
{$ENDIF}
      end
      else
        RowIsEmpty := RowIsEmpty and True;
    end;
    DoUserDataFormat(FImportRow[i]);
  end;
end;

function TADO_QImport3Access.ImportData: TQImportResult;
begin
  Result := qirOk;
  try
    try
      if Canceled  and not CanContinue then
      begin
        Result := qirBreak;
        Exit;
      end;

      DataManipulation;

    except
      on E:Exception do begin
        try
          DestinationCancel;
        except
        end;
        DoImportError(E);
        Result := qirContinue;
        Exit;
      end;
    end;
  finally
    if (not IsCSV) and (CommitRecCount > 0) and not CommitAfterDone and
       (
        ((ImportedRecs + ErrorRecs) > 0)
        and ((ImportedRecs + ErrorRecs) mod CommitRecCount = 0)
       )
    then
      DoNeedCommit;
    if (ImportRecCount > 0) and
       ((ImportedRecs + ErrorRecs) mod ImportRecCount = 0) then
      Result := qirBreak;
  end;
end;

procedure TADO_QImport3Access.ChangeCondition;
begin
  FADO.Next;
  if FSkipCounter > 0 then Dec(FSkipCounter);
end;

procedure TADO_QImport3Access.FinishImport;
begin
  try
    if not Canceled and not IsCSV then
    begin
      if CommitAfterDone then
        DoNeedCommit
      else if (CommitRecCount > 0) and ((ImportedRecs + ErrorRecs) mod CommitRecCount > 0) then
        DoNeedCommit;
    end;
  finally
    {$IFDEF VCL6}
      FADO.Close;
    {$ELSE}
    try
      FADO.Close;
    except
    end;
    {$ENDIF}
    if Assigned(FADO) then FADO.Free;
  end;
end;

function TADO_QImport3Access.GetConnectionString: string;
begin
  Result := BuildConnectionString(FVersion, FLogin, FPassword);
end;

procedure TADO_QImport3Access.GetFieldNames(var FieldNames, FieldTypesString:
    TStringArray; var FieldTypes: TFieldTypes; var FieldSizes: TFieldSizes);

  function FieldTypeToStr(ADataType: TFieldType;
                          const AFieldSize: Integer): string;
  begin
    case ADataType of
      ftBlob,
      {$IFDEF VCL10}
        ftWideMemo,
      {$ENDIF}
      ftMemo:
        Result := 'Memo';
      ftWideString,
      ftString,
      ftGuid:
        Result := Format('Text(%d)', [AFieldSize]);
      ftAutoInc:
        Result := 'AutoNumber';
      ftSmallint,
      ftInteger,
      ftWord,
      ftLargeInt,
      ftFloat,
      {$IFDEF VCL6}ftFMTBcd,{$ENDIF}
      ftBCD:
        Result := 'Number';
      ftBoolean:
        Result := 'Yes/No';
      ftCurrency:
        Result := 'Currency';
      ftDate,
      ftTime,
      {$IFDEF VCL6}ftTimeStamp,{$ENDIF}
      ftDateTime:
        Result := 'Date/Time';
    else
      Result := 'Unknown';
    end;
  end;

var
  ADO: TADOQuery;
  connStr: string;
  i: Integer;
  fd: TFieldDef;
begin
  connStr := GetConnectionString();

  ADO := TADOQuery.Create(nil);
  try
    ADO.ConnectionString := Format(connStr, [FileName]);
    ADO.CursorType := ctOpenForwardOnly;
    ADO.CursorLocation := clUseServer;
    ADO.LockType := ltReadOnly;
    if SourceType = isSQL then
      ADO.SQL.Assign(Self.SQL)
    else
      ADO.SQL.Text := Format(SelectFromTableFieldNames, [TableName]);

    if ADO.SQL.Text <> '' then
    begin
      ADO.Open;
      try
        SetLength(FieldNames, ADO.FieldCount);
        SetLength(FieldTypesString, ADO.FieldCount);
        SetLength(FieldTypes, ADO.FieldCount);
        SetLength(FieldSizes, ADO.FieldCount);
        for i := 0 to ADO.FieldCount - 1 do
        begin
          fd := ADO.FieldDefList.FieldDefs[i];
          if fd <> nil then
          begin
            FieldNames[I] := Fd.Name;
            FieldTypesString[I] := FieldTypeToStr(fd.DataType, fd.Size);
            FieldTypes[I] := Fd.DataType;
            FieldSizes[I] := Fd.Size;
          end
          else
          begin
            FieldNames[I] := EmptyStr;
            FieldTypesString[I] := EmptyStr;
            FieldTypes[I] := ftUnknown;
            FieldSizes[I] := -1;
          end;  
        end;
      finally
        ADO.Close;
      end;
    end;
  finally
    ADO.Free;
  end;
end;

procedure TADO_QImport3Access.GetTableNames(List: TqiStrings);
var
  ADO: TADOConnection;
  connStr: string;
  Names: TStrings;
begin
  connStr := GetConnectionString();

  ADO := TADOConnection.Create(nil);
  try
    ADO.LoginPrompt := false;
    ADO.ConnectionString := Format(connStr, [FileName]);
    ADO.Open;
    Names := TStringList.Create;
    try
      ADO.GetTableNames(Names, false);
      List.Assign(Names);
    finally
      ADO.Close;
      Names.Free;
    end;
  finally
    ADO.Free;
  end;
end;

{$IFDEF QI_UNICODE}
procedure TADO_QImport3Access.GetTableNames(List: TStrings);
var
  ADO: TADOConnection;
  connStr: WideString;
  Names: TStrings;
begin
  connStr := GetConnectionString();

  ADO := TADOConnection.Create(nil);
  try
    ADO.LoginPrompt := false;
    ADO.ConnectionString := Format(connStr, [FileName]);
    ADO.Open;
    Names := TStringList.Create;
    try
      ADO.GetTableNames(Names, false);
      List.Assign(Names);
    finally
      ADO.Close;
      Names.Free;
    end;
  finally
    ADO.Free;
  end;
end;
{$ENDIF}

procedure TADO_QImport3Access.SetSQL(const Value: TStrings);
begin
  FSQL.Assign(Value);
end;

//**************************************************************************
//
//  Access Password Encrypt/Decrypt
//
//**************************************************************************

function SimpleXOR(const text: string): string;
const
  key = #9#8#7#6#5#4#3#2#1#0;
var
  longkey: string;
  i: integer;
  toto: char;
begin
  Result := '';
  for i := 0 to (length(text) div length(key)) do
    longkey := longkey + key;
  for i := 1 to length(text) do
  begin
    toto := chr((ord(text[i]) xor ord(longkey[i])));
    result := result + toto;
  end;
end;

function adler32(const buf : string; len : Cardinal) : Cardinal;
var
  s1, s2: Cardinal;
  I: Integer;
begin
  s1 := Cardinal(1);
  s2 := Cardinal(0);
  for I := 0 to len - 1 do
  begin
    s1 := (s1 + Ord(buf[i+1]))  mod Cardinal(65521);
    s2 := (s2 + s1)             mod Cardinal(65521);
  end;
  Result := (s2 shl 16) + s1;
end;

type
  TAdlerRec = packed record
    case Integer of
      0: (b1, b2, b3, b4: Byte);
      1: (Adler: Cardinal);
  end;

function AdlerStr(adler: Cardinal): string;
var
  a: TAdlerRec;
begin
  a.Adler := adler;
  Result := Char(a.b1)+Char(a.b2)+Char(a.b3)+Char(a.b4);
end;

function AdlerNum(str: string): Cardinal;
var
  a: TAdlerRec;
begin
  a.b1 := Byte(str[1]);
  a.b2 := Byte(str[2]);
  a.b3 := Byte(str[3]);
  a.b4 := Byte(str[4]);
  Result := a.adler;
end;

function PasswordDecrypt(const HexStr: string): string;
var
  Pass, S: string;
  a: Cardinal;
begin
  Result := HexStr;
  if HexStr = '' then
    Exit;

  if HexToString(HexStr,Pass) then
  begin
    if Length(Pass) > 4 then
    begin
      s := Copy(Pass,5,length(Pass)-4);
      a := adler32( S, Length(s));
      if AdlerNum(Pass) = a then
      begin
        Result := SimpleXOR(s);
      end;
    end;
  end;
end;

function PasswordEncrypt(const Password: string): string;
var
  pass, prefix: string;
  a: Cardinal;
begin
  Result := '';
  if Password = '' then
    Exit;

  pass := SimpleXOR(Password);
  a := adler32( Pass, Length(Pass));
  prefix := AdlerStr( a );

  Result := StringToHex(prefix + pass);
end;

//**************************************************************************

procedure TADO_QImport3Access.DoLoadConfiguration(IniFile: TIniFile);
var
  AStrings: TStrings;
  i : Integer;
begin
  inherited;
  with IniFile do
  begin
    SkipFirstRows := ReadInteger(ACCESS_OPTIONS, ACCESS_SKIP_LINES, SkipFirstRows);
    SourceType := TQImportAccessSourceType(ReadInteger(ACCESS_OPTIONS, ACCESS_SOURCE_TYPE, Integer(SourceType)));
    TableName := ReadString(ACCESS_OPTIONS, ACCESS_TABLE_NAME, TableName);
    Password := PasswordDecrypt(ReadString(ACCESS_OPTIONS, ACCESS_PASSWORD, EmptyStr));
    AStrings := TStringList.Create;
    try
      AStrings.Clear;
      SQL.Clear;
      ReadSection(ACCESS_SQL, AStrings);
      for i := 0 to AStrings.Count - 1 do
        SQL.Add( ReadString(ACCESS_SQL, AStrings[i], EmptyStr) );
    finally
      AStrings.Free;
    end;
  end;
end;

procedure TADO_QImport3Access.DoSaveConfiguration(IniFile: TIniFile);
var
  i : Integer;
begin
  inherited;
  with IniFile do
  begin
    WriteInteger(ACCESS_OPTIONS, ACCESS_SKIP_LINES, SkipFirstRows);
    WriteInteger(ACCESS_OPTIONS, ACCESS_SOURCE_TYPE, Integer(SourceType));
    WriteString(ACCESS_OPTIONS, ACCESS_TABLE_NAME, TableName);
    WriteString(ACCESS_OPTIONS, ACCESS_PASSWORD, PasswordEncrypt( Password ));
    EraseSection(ACCESS_SQL);
    for i := 0 to SQL.Count - 1 do
      WriteString(ACCESS_SQL, Format('%s%.3d',[ACCESS_SQL_LINE,i+1]), SQL.Strings[i]);
  end;
end;


initialization
  CoInitialize(nil);

finalization
  CoUninitialize;

end.
