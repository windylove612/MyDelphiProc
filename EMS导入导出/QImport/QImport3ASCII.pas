unit QImport3ASCII;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    {$IFDEF QI_UNICODE}
      System.WideStrings,
      QImport3GpTextFile,
    {$ENDIF}
    Winapi.Windows,
    System.Classes,
    System.SysUtils,
    System.IniFiles,
  {$ELSE}
    {$IFDEF QI_UNICODE}
      {$IFDEF VCL10}
        WideStrings,
      {$ELSE}
        QImport3WideStrings,
      {$ENDIF}
      QImport3GpTextFile,
    {$ENDIF}
    Windows,
    Classes,
    SysUtils,
    IniFiles,
  {$ENDIF}
  QImport3,
  QImport3EZDSLHsh,
  QImport3StrTypes;

type
  TQImport3ASCII = class(TQImport3)
  private
    FComma: AnsiChar;
    FQuote: AnsiChar;
    FEncoding: TQICharsetType;
    FCounter: Integer;
{$IFDEF QI_UNICODE}
    FBuffStrW: WideString;
    FFileW: TGpTextFile;
    FColumnsW: TWideStrings;
{$ELSE}
    FBuffStr: AnsiString;
    FFile: TextFile;
    FColumns: TStrings;
{$ENDIF}
    FColumnsHash: THashTable;
    FPosDataHash: THashTable;

    function HasComma: Boolean;
    procedure ReadColumns( const Str: qiString; AStrings: TqiStrings;
      AColumnsHash: THashTable);
  protected
    function CheckProperties: Boolean; override;

    procedure StartImport; override;
    function CheckCondition: Boolean; override;
    function Skip: Boolean; override;
    procedure FillImportRow; override;
    function ImportData: TQImportResult; override;
    procedure ChangeCondition; override;
    procedure FinishImport; override;

    procedure DoLoadConfiguration(IniFile: TIniFile); override;
    procedure DoSaveConfiguration(IniFile: TIniFile); override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property FileName;
    property SkipFirstRows default 0;
    property Comma: AnsiChar read FComma write FComma;
    property Quote: AnsiChar read FQuote write FQuote;
    property Encoding: TQICharsetType read FEncoding write FEncoding
      default ctWinDefined;
  end;

function EndQuote(const Value:
  {$IFDEF QI_UNICODE}
    WideString
  {$ELSE}
    AnsiString
  {$ENDIF}; const Comma, Quote: AnsiChar): Boolean;
  
function NewLineInValue(const Value:
  {$IFDEF QI_UNICODE}
    WideString
  {$ELSE}
    AnsiString
  {$ENDIF}; const Comma, Quote: AnsiChar): Boolean;
{$IFDEF QI_UNICODE}
function CsvReadLn(const CsvFile: TGpTextFile; const Comma, Quote: AnsiChar): WideString;
{$ENDIF}

implementation

uses
  {$IFDEF VCL16}
    Data.DB,
    System.StrUtils,
  {$ELSE}
    DB,
    {$IFDEF VCL7}
      StrUtils,
    {$ENDIF}
  {$ENDIF}
  QImport3StrIDs,
  QImport3Common,
  QImport3WideStrUtils,
  QImport3EzdslBse;

type
  PPosData = ^TPosData;
  TPosData = record
    Pos: Integer;
    Len: Integer;
  end;

  PStrData = ^TStrData;
  TStrData = record
    Str: qiString;
  end;

{$WARNINGS OFF}
function EndQuote(const Value:
  {$IFDEF QI_UNICODE}
    WideString
  {$ELSE}
    AnsiString
  {$ENDIF}; const Comma, Quote: AnsiChar): Boolean;

var
  PosCQ,
  PosQC: Integer;
begin
  PosCQ := Pos(Comma + Quote, Value);
  PosQC := Pos(Char(Quote), Value);
  Result := (PosQC < PosCQ) or ((PosQC > 0) and (PosCQ = 0)) or
    (PosCQ + 1 = PosQC);
end;
{$WARNINGS ON}

function NewLineInValue(const Value:
  {$IFDEF QI_UNICODE}
    WideString
  {$ELSE}
    AnsiString
  {$ENDIF}; const Comma, Quote: AnsiChar): Boolean;
  {$IFNDEF VCL7}
  function PosEx(const SubStr, S: string; Offset: Cardinal = 1): Integer;
  var
    I,X: Integer;
    Len, LenSubStr: Integer;
  begin
    if Offset = 1 then
      Result := Pos(SubStr, S)
    else
    begin
      I := Offset;
      LenSubStr := Length(SubStr);
      Len := Length(S) - LenSubStr + 1;
      while I <= Len do
      begin
        if S[I] = SubStr[1] then
        begin
          X := 1;
          while (X < LenSubStr) and (S[I + X] = SubStr[X + 1]) do
            Inc(X);
          if (X = LenSubStr) then
          begin
            Result := I;
            exit;
          end;
        end;
        Inc(I);
      end;
      Result := 0;
    end;
  end;
  {$ENDIF}
var
  PosCQ,
  PosCQLast,
  ValueLength,
  PosLastQuote,
  AllQuotes,
  I: Integer;
  CommaQuote: string;
  LastCommaInQuote,
  NonQuoted: Boolean;
const
  NewLineChars = #13#10;
begin
  CommaQuote := String(Comma + Quote);
  PosCQLast := 0;
  PosCQ := PosEx(CommaQuote, Value, 1);
  while PosCQ > 0 do
  begin
    PosCQLast := PosCQ;
    PosCQ := PosEx(CommaQuote, Value, PosCQ + 1);
  end;

  ValueLength := Length(Value);
  AllQuotes := ValueLength - Length(QIStringReplace(Value, String(Quote), '', [rfReplaceAll]));
  PosLastQuote := PosEx(String(Quote), Value, PosCQLast + 2);
  NonQuoted := True;
  if (PosLastQuote > 0) then
  while PosLastQuote > 0 do
    if ((PosLastQuote + 1) <= ValueLength) and
      (AnsiChar(Value[PosLastQuote + 1]) = Quote) then
    begin
      PosLastQuote := PosEx(String(Quote), Value, PosLastQuote + 2);
      NonQuoted := True;
    end
    else
    begin
      PosLastQuote := 0;
      NonQuoted := False;
    end
  else begin
  //нестандартные ситуации (последний CQ он же QC)
    NonQuoted := not ((PosCQLast > 0) and (AnsiChar(Value[PosCQLast + 2]) = Comma) and (AllQuotes mod 2 = 0));
  end;

  Result := (PosCQLast > 0) and NonQuoted;
  if Result then
  begin
    LastCommaInQuote := False;
    for I := 1 to PosCQLast do
      if AnsiChar(Value[I]) = Quote then
        LastCommaInQuote := not LastCommaInQuote;
     Result := not LastCommaInQuote;
  end;
  //выходим из зацикливания, если что-то из пользовательского форматирования не учли
  if Result then
     Result := PosEx(NewLineChars, Value, 1) = 0;
end;

{$IFDEF QI_UNICODE}
function CsvReadLn(const CsvFile: TGpTextFile; const Comma, Quote: AnsiChar): WideString;
const
  NewLineChars = #13#10;
var
  TmpValue,
  ValueForSearchNewLine: WideString;
begin
  TmpValue := '';
  Result := CsvFile.Readln;
  ValueForSearchNewLine := Result;
  while NewLineInValue(ValueForSearchNewLine, Comma, Quote) do
  repeat
    TmpValue := TmpValue + Result + NewLineChars;
    Result := CsvFile.ReadLn;
    ValueForSearchNewLine := TmpValue + Result;
  until EndQuote(Result, Comma, Quote) or CsvFile.EOF;

  Result := TmpValue + Result;
end;
{$ENDIF}

procedure DisposePosHashData(AData: Pointer);
begin
  Dispose(AData);
end;

procedure DisposeColumnsHashData(AData: Pointer);
begin
  Finalize(PStrData(AData)^);
  Dispose(AData);
end;

{ TQImport3ASCII }

constructor TQImport3ASCII.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
  FComma := AnsiChar(GetListSeparator);
  FQuote := '"';
  FEncoding := ctWinDefined;
end;

function TQImport3ASCII.CheckProperties: Boolean;
var
  i, j: Integer;
begin
  Result := inherited CheckProperties;
  if HasComma then // example field=2
  begin
    for i := 0 to Map.Count - 1 do
      try
        StrToInt(Map.Values[Map.Names[i]]);
      except
        on E: EConvertError do
          raise EQImportError.Create(QImportLoadStr(QIE_MapMissing));
      end;
  end
  else begin // example: field=3;5
    for i := 0 to Map.Count - 1 do
    begin
      j := Pos(';', Map.Values[Map.Names[i]]);
      if j = 0 then
        raise EQImportError.Create(QImportLoadStr(QIE_MapMissing))
      else begin
        try
          StrToInt(Copy(Map.Values[Map.Names[i]], 1, j - 1));
        except
          on E:EConvertError do
            raise EQImportError.Create(QImportLoadStr(QIE_MapMissing));
        end;
        try
          StrToInt(Copy(Map.Values[Map.Names[i]], j + 1, Length(Map.Values[Map.Names[i]]) - j));
        except
          on E:EConvertError do
            raise EQImportError.Create(QImportLoadStr(QIE_MapMissing));
        end;
      end;
    end;
  end
end;

procedure TQImport3ASCII.StartImport;
var
  i, k: Integer;
  pi: Pointer;
  str: string;
  posData: PPosData;
begin
{$IFDEF QI_UNICODE}
  FFileW := TGpTextFile.CreateEx(FileName, FILE_ATTRIBUTE_NORMAL,
    GENERIC_READ, FILE_SHARE_READ);
  FFileW.TryReadUpCRLF := True;
  FFileW.Reset;
  FFileW.Codepage := QICharsetToCodepage(FEncoding);
  FColumnsW := TWideStringList.Create;
{$ELSE}
  AssignFile(FFile, FileName);
  Reset(FFile);
  FColumns := TStringList.Create;
{$ENDIF}
  FColumnsHash := THashTable.Create(True);

  FColumnsHash.TableSize := Map.Count;
  FColumnsHash.DisposeData := DisposeColumnsHashData;

  FPosDataHash := THashTable.Create(True);
  FPosDataHash.TableSize := Map.Count;
  FPosDataHash.DisposeData := DisposePosHashData;

  for i := 0 to Map.Count - 1 do
  begin
    if FImportRow.MapNameIdxHash.Search(Map.Names[i], pi) then
    begin
      k := Integer(pi);
{$IFDEF VCL7}
      str := Map.ValueFromIndex[k];
{$ELSE}
      str := Map.Values[Map.Names[k]];
{$ENDIF}
      New(posData);
      posData^.Pos := StrToIntDef(Copy(str, 1, Pos(';', str) - 1), 0);
      posData^.Len := StrToIntDef(Copy(str, Pos(';', str) + 1, Length(str)), 0);
      FPosDataHash.Insert(Map.Names[i], posData);
    end;
  end;

  FCounter := 0;
end;

function TQImport3ASCII.CheckCondition: Boolean;
begin
{$IFDEF QI_UNICODE}
  Result := not FFileW.Eof;
{$ELSE}
  Result := not Eof(FFile);
{$ENDIF}  
end;

function TQImport3ASCII.Skip: Boolean;

  function BuffEmpty: Boolean;
  begin
  {$IFDEF QI_UNICODE}
    Result := Trim(FBuffStrW) = '';
  {$ELSE}
    Result := Trim(FBuffStr) = '';
  {$ENDIF}
  end;
  
  {$IFNDEF QI_UNICODE}
  const
    NewLine = #13#10;

  var
    TmpValue,
    ValueForSearchNewLine: AnsiString;
  {$ENDIF}

begin
{$IFDEF QI_UNICODE}

  if not HasComma then
  begin
    FBuffStrW := FFileW.ReadLn;
    ReplaceTabs(FBuffStrW)
  end
  else
    FBuffStrW := CsvReadLn(FFileW, Comma, Quote);
{$ELSE}
  Readln(FFile, FBuffStr);
  if not HasComma then
    ReplaceTabs(FBuffStr)
  else
  begin
    TmpValue := '';
    ValueForSearchNewLine := FBuffStr;
    while NewLineInValue(ValueForSearchNewLine, Comma, Quote) do
    repeat
      TmpValue := TmpValue + FBuffStr + NewLine;

      Readln(FFile, FBuffStr);
      ValueForSearchNewLine := TmpValue + FBuffStr;
    until EndQuote(FBuffStr, Comma, Quote) or Eof(FFile);

    FBuffStr := TmpValue + FBuffStr;
  end;
{$ENDIF}
  Result := (SkipFirstRows > 0) and (FCounter < SkipFirstRows);
  if Result then
    Inc(FCounter);
  Result := Result or BuffEmpty;
end;

procedure TQImport3ASCII.FillImportRow;
var
  p: Pointer;
  j, k: Integer;
  ps: PStrData;
begin
  FImportRow.ClearValues;
  FColumnsHash.Empty();

{$IFDEF QI_UNICODE}
  ReadColumns(FBuffStrW, FColumnsW, FColumnsHash);
{$ELSE}
  ReadColumns(FBuffStr, FColumns, FColumnsHash);
{$ENDIF}
  RowIsEmpty := True;
  for j := 0 to FImportRow.Count - 1 do
  begin
    if FImportRow.MapNameIdxHash.Search(FImportRow[j].Name, p) then
    begin
      k := Integer(p);
{$IFDEF QI_UNICODE}
      try
        if HasComma then
{$IFDEF VCL7}
          FBuffStrW := FColumnsW[StrToInt(Map.ValueFromIndex[k]) - 1]
{$ELSE}
          FBuffStrW := FColumnsW[StrToInt(Map.Values[Map.Names[k]]) - 1]
{$ENDIF}
        else begin
          ps := FColumnsHash.Examine(FImportRow[j].Name);
          if ps <> nil then
            FBuffStrW := ps^.Str;
        end;
      except
        FBuffStrW := '';
      end;
      if AutoTrimValue then
        FBuffStrW := Trim(FBuffStrW);
      RowIsEmpty := RowIsEmpty and (FBuffStrW = '');
      FImportRow.SetValue(Map.Names[k], FBuffStrW, False);
{$ELSE}
      try
        if HasComma then
{$IFDEF VCL7}
          FBuffStr := FColumns[StrToInt(Map.ValueFromIndex[k]) - 1]
{$ELSE}
          FBuffStr := FColumns[StrToInt(Map.Values[Map.Names[k]]) - 1]
{$ENDIF}
        else begin
          ps := FColumnsHash.Examine(FImportRow[j].Name);
          if ps <> nil then
            FBuffStr := ps^.Str;
        end;
      except
        FBuffStr := '';
      end;
      FImportRow.SetValue(Map.Names[k], FBuffStr, False);
{$ENDIF}
    end;
    DoUserDataFormat(FImportRow[j]);
  end;
end;

function TQImport3ASCII.ImportData: TQImportResult;
begin
  Result := qirOk;
  try
    try
      if Canceled and not CanContinue then
      begin
        Result := qirBreak;
        Exit;
      end;

      DataManipulation;

    except
      on E:Exception do
      begin
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

procedure TQImport3ASCII.ChangeCondition;
begin
end;

procedure TQImport3ASCII.FinishImport;
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
{$IFDEF QI_UNICODE}
    if Assigned(FColumnsW) then
      FColumnsW.Free;
    FFileW.Free;
{$ELSE}
    if Assigned(FColumns) then
      FColumns.Free;
    CloseFile(FFile);
{$ENDIF}
    FColumnsHash.Free;
    FPosDataHash.Free;
  end;
end;

procedure TQImport3ASCII.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    SkipFirstRows := ReadInteger(ASCII_OPTIONS, ASCII_SKIP_LINES, SkipFirstRows);
    Encoding := TQICharsetType(ReadInteger(ASCII_OPTIONS, ASCII_ENCODING, Integer(Encoding)));
    Comma := Str2Char(ReadString(ASCII_OPTIONS, ASCII_COMMA,
      Char2Str(Comma)), Comma);
    Quote := Str2Char(ReadString(ASCII_OPTIONS, ASCII_QUOTE,
      Char2Str( Quote)), Quote);
  end;
end;

procedure TQImport3ASCII.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    WriteInteger(ASCII_OPTIONS, ASCII_SKIP_LINES, SkipFirstRows);
    WriteInteger(ASCII_OPTIONS, ASCII_ENCODING, Integer(Encoding));
    WriteString(ASCII_OPTIONS, ASCII_COMMA, Char2Str(Comma));
    WriteString(ASCII_OPTIONS, ASCII_QUOTE, Char2Str(Quote));
  end;
end;

function TQImport3ASCII.HasComma: Boolean;
begin
  Result := FComma <> #0;
end;

procedure TQImport3ASCII.ReadColumns(
  const Str: qiString;
  AStrings: TqiStrings;
  AColumnsHash: THashTable);
var
  P, L, i: Integer;
  pi: Pointer;
  ps: PStrData;
begin
  if HasComma then
  begin
    CSVStringToStrings(Str, Quote, Comma, AStrings);
  end
  else begin
    for i := 0 to Map.Count - 1 do
    begin
      if FPosDataHash.Search(Map.Names[i], pi) then
      begin
        P := PPosData(pi)^.Pos;
        L := PPosData(pi)^.Len;
        New(ps);
        ps^.Str := Copy(Str, P + 1, L);
        AColumnsHash.Insert(Map.Names[i], ps);
      end;
    end;
  end;
end;

end.
