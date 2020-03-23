unit QImport3DBF;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.IniFiles,
    System.Classes,
    Winapi.Windows,
  {$ELSE}
    Classes,
    Windows,
    IniFiles,
  {$ENDIF}
    QImport3,
    QImport3DBFFile,
    QImport3EZDSLHsh,
    QImport3StrTypes;

type
  TDBFCharSet = (
    dcsNone, dcsLatin1, dcsArmscii8, dcsAscii, dcsCp850, dcsCp852, dcsCp866,
    dcsCp1250, dcsCp1251, dcsCp1256, dcsCp1257, dcsDec8, dcsGeostd8, dcsGreek,
    dcsHebrew, dcsHp8, dcsKeybcs2, dcsKoi8r, dcsKoi8u, dcsLatin2, dcsLatin5,
    dcsLatin7, dcsMacce, dcsMacroman, dcsSwe7, dcsUtf8, {dcsUtf16, dcsUtf32,}
    dcsLatin3, dcsLatin4, dcsLatin6, dcsLatin8, dcsIso8859_5, dcsIso8859_6,
    dcsCp1026, dcsCp1254, dcsCp1255, dcsCp1258, dcsCp437, dcsCp500, dcsCp737, dcsCp855,
    dcsCp856, dcsCp857, dcsCp860, dcsCp862, dcsCp863, dcsCp864, dcsCp865, dcsCp869,
    dcsCp874, dcsCp875, dcsIceland,
    dcsBig5, dcsKSC5601, dcsEUC, dcsGB2312, dcsSJIS_0208, dcsLatin9, dcsLatin13,
    dcsCp1252, dcsCp1253, dcsCp775, dcsCp858);

  TQImport3DBF = class(TQImport3)
  private
    FDBF: TDBFRead;
    FSourceFieldsHash: THashTable;

    FSkipDeleted: boolean;
    FCharSet: TDBFCharSet;
    function DBFStrToDateTime(const Str: string): TDateTime;
  protected
    function AllowImportRowComplete: Boolean; override;
  protected
    procedure StartImport; override;
    function CheckCondition: boolean; override;
    function Skip: boolean; override;
    procedure FillImportRow; override;
    function ImportData: TQImportResult; override;
    procedure ChangeCondition; override;
    procedure FinishImport; override;

    procedure BeforeImport; override;
    procedure AfterImport; override;

    procedure DoLoadConfiguration(IniFile: TIniFile); override;
    procedure DoSaveConfiguration(IniFile: TIniFile); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property SkipDeleted: boolean read FSkipDeleted write FSkipDeleted
      default true;
    property SkipFirstRows default 0;
    property CharSet: TDBFCharSet read FCharSet
      write FCharSet default dcsNone;
    property FileName;
  end;

implementation

uses
  {$IFDEF VCL16}
    Data.Db,
    System.SysUtils,
  {$ELSE}
    Db,
    SysUtils,
  {$ENDIF}
  QImport3Common;

const
  cDBFSystemCP: array[TDBFCharSet] of Integer =
    (
        -1, 1252, -1, 20127, 850, 852, 866, 1250, 1251, 1256, 1257, -1, -1, 28597,
      28598, -1, -1, 20866, 21866, 870, 1026, -1, 10029, 10000, 20107, CP_UTF8, {1200, 12000,}
      28593, -1, -1, -1, 28595, 28596,
      1026, 1254, 1255, 1258, 437, 500, 737, 855, 856, 857, 860, 862, 863, 864, 865, 869,
      874, 875, 871,
      950, 949, -1, 936, -1, 28599, 28594, 1252, 1253, 775, 858
    );
    

{ TQImport3DBF }

function TQImport3DBF.AllowImportRowComplete: Boolean;
begin
  Result := not (FSkipDeleted and FDBF.Deleted);
end;

procedure TQImport3DBF.BeforeImport;
begin
  FDBF := TDBFRead.Create(FileName);
  FTotalRecCount := FDBF.RecordCount;
  inherited;
end;

procedure TQImport3DBF.AfterImport;
begin
  if Assigned(FDBF) then FDBF.Free;
  inherited;
end;

procedure TQImport3DBF.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do                             
  begin
    SkipFirstRows := ReadInteger(DBF_OPTIONS, DBF_SKIP_LINES, SkipFirstRows);
    CharSet := TDBFCharSet(ReadInteger(DBF_OPTIONS, DBF_CHARSET, Integer(CharSet)));
  end;
end;

procedure TQImport3DBF.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    WriteInteger(DBF_OPTIONS, DBF_SKIP_LINES, SkipFirstRows);
    WriteInteger(DBF_OPTIONS, DBF_CHARSET, Integer(CharSet));
  end;
end;

constructor TQImport3DBF.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
  FSkipDeleted := true;
  FCharSet := dcsNone;
end;

procedure TQImport3DBF.StartImport;
var
  i: integer;
begin
  FSourceFieldsHash := THashTable.Create(False);
  FSourceFieldsHash.TableSize := Map.Count;
  for i := 0 to Map.Count - 1 do
{$IFDEF VCL7}
    FSourceFieldsHash.Insert(Map.ValueFromIndex[i], Pointer(i));
{$ELSE}
    FSourceFieldsHash.Insert(Map.Values[Map.Names[i]], Pointer(i));
{$ENDIF}
end;

function TQImport3DBF.CheckCondition: boolean;
begin
  Result := not FDBF.EOF;
end;

function TQImport3DBF.Skip: boolean;
begin
  Result := false;
end;

procedure TQImport3DBF.FillImportRow;
var
  i, j, l: integer;
  dataStr: String;
  value: WideString;
  p: Pointer;
begin
  FImportRow.ClearValues;
  RowIsEmpty := True;
  for i := 0 to FDBF.FieldCount - 1 do
  begin
    value := '';
    dataStr := String(FDBF.GetData(i));
    if (FCharSet <> dcsNone) and (dataStr <> EmptyStr) then
    {$IFDEF VCL6}
      if FCharSet = dcsUtf8 then
        dataStr := UTF8Decode(AnsiString(dataStr))
      else 
      {$ENDIF}
      begin
        l := MultiByteToWideChar(cDBFSystemCP[FCharSet], 0, PAnsiChar(AnsiString(dataStr)), Length(dataStr), nil, 0);
        SetLength(value, l);
        MultiByteToWideChar(cDBFSystemCP[FCharSet], 0, PAnsiChar(AnsiString(dataStr)), Length(dataStr), PWideChar(value), l);
        dataStr := value;
      end;
    if FSourceFieldsHash.Search({$IFDEF VCL12}string{$ENDIF}(FDBF.FieldName[i]), p) then
    begin
      j := Integer(p);
      if (FDBF.FieldType[i] = dftDate) and (dataStr <> '') then
        dataStr := {$IFDEF VCL12}String{$ENDIF}(DateTimeToStr(DBFStrToDateTime(
          {$IFDEF VCL12}string{$ENDIF}(dataStr))));
      if AutoTrimValue then
        dataStr := Trim(dataStr);
      RowIsEmpty := RowIsEmpty and (dataStr = '');
      FImportRow.SetValue(Map.Names[j], qiString(dataStr), False);
      DoUserDataFormat(FImportRow.ColByName(Map.Names[j]));
    end
  end;

  for i := 0 to FImportRow.Count - 1 do
  begin
    if not FImportRow.MapNameIdxHash.Search(FImportRow[i].Name, p) then
      DoUserDataFormat(FImportRow[i]);
  end;
end;

function TQImport3DBF.ImportData: TQImportResult;
begin
  Result := qirOk;
  try
    try
      if Canceled  and not CanContinue then begin
        Result := qirBreak;
        Exit;
      end;
      if (SkipFirstRows > 0) and (FDBF.RecNo <= SkipFirstRows) then begin
        DestinationCancel;
        Result := qirContinue;
        Exit;
      end;

      if FSkipDeleted and FDBF.Deleted then begin
        Result := qirContinue;
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

procedure TQImport3DBF.ChangeCondition;
begin
  //
end;

procedure TQImport3DBF.FinishImport;
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
    FSourceFieldsHash.Free;
  end;
end;

function TQImport3DBF.DBFStrToDateTime(const Str: string): TDateTime;
var
  Year, Month, Day: word;
begin
  try
    Year := StrToIntDef(Copy(Str, 1, 4), 0);
    Month := StrToIntDef(Copy(Str, 5, 2), 0);
    Day := StrToIntDef(Copy(Str, 7, 2), 0);
    Result := EncodeDate(Year, Month, Day);
  except
    Result := 0;
  end;
end;

end.
