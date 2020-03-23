unit QImport3DBFFile;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.SysUtils,
    System.Classes,
    Winapi.Windows;
  {$ELSE}
    Classes,
    SysUtils,
    Windows;
  {$ENDIF}

{----------------------------------------------------------------------------}
const
  dBaseIII     = $03;
  dBaseIIIMemo = $83;
  dBaseIVMemo  = $8B;
  dBaseIVSQL   = $63;
  FoxPro       = $05;
  FoxProMemo   = $F5;
  VisualFoxPro = $30;   

  dftString  = 'C'; // AnsiChar
  dftBoolean = 'L'; // boolean
  dftNumber  = 'N'; // number
  dftMemo    = 'M'; // memo
  dftFloat   = 'F'; // float -- not in DBaseIII
  dftDate    = 'D'; // date

  dbfFileEmpty     = 100;
  dbfENoFileName   = 101;
  dbfEFileNotFound = 102;
  dbfECreate       = 103;
  dbfEOpen         = 104;
  dbfESeek         = 105;
  dbfERead         = 106;
  dbfEWrite        = 107;
  dbfEOutOfMemory  = 108;
  dbfEDestruct     = 109;
  dbfECreateList   = 110;
  dbfEOutOfNumber  = 111;

  crEFileNotFound  = 200;
  crEOutOfMemory  = 201;

type
  RFileHeader = packed record { *** Первая запись заголовка ***  L=32 }
   {+0} DBType, {  7 6 5 4 3 2 1 0
                                                  ¦ ¦ ¦ ¦ ¦ ¦ ¦ L- 
                                                  ¦ ¦ ¦ ¦ ¦ ¦ L--- ¦ Номер версии
                                                  ¦ ¦ ¦ ¦ ¦ L----- 
                                                  ¦ ¦ ¦ ¦ L------- Файл DBT [в DBASE IV]
                                                  ¦ ¦ ¦ L--------- ???
                                                  ¦ ¦ L-----------  Флаг SQL [только в DBASE IV]
                                                  ¦ L------------- 
                                                  L--------------- Подключение файла DBT  }
   {+1} Year, { Год последней модификации }
   {+2} Month, { Месяц ---------//-------- }
   {+3} Day: Byte; { День  ---------//-------- }
   {+4} RecCount: LongInt; { Число записей включая удаленные}
   {+8} HeaderSize: Word; { Положение первой записи}
   {+10} RecordSize: Longint; { Длина записи }
   {+14} FDelayTrans: Byte; { Флаг задержки транзакции }
   {+15} Reserve2: array[1..13] of Byte; { Резервные байты }
   {+28} FlagMDX: Byte; { Флаг подключения множествен-
                                                ного индекса 1 - есть файл
                                                .MDX, 0 - нет }
   {+29} Reserve3: array[1..3] of Byte; { Резерв }
  end;

type
  RMemoHeader = packed record { *** Пеpвый блок DBT - файла *** L = 512 }
    NextBlock: DWORD;
    BlockSize: DWORD;
    Reserve: array[1..504] of Byte;
  end;

  RFPTMemoHeader = packed record // Foxpro memo file header
    NextBlock: DWORD;
    Unused: WORD;
    BlockSize: WORD;
    Reserve: array[1..504] of byte;
  end;

  RFPTMemoBlockHeader = packed record // Foxpro memo file block header
    BlockSignature: DWORD;
    MemoLength: DWORD;
  end;

type
  TFieldName = array[1..10] of AnsiChar;

  RDescriptor = packed record { *** Элемент (дескриптор) заголовка *** L= 32 }
    {+0} FieldName: TFieldName; { Имя поля }
    {+10} FieldEnd: AnsiChar; { Теpминиpующий байт 00h }
    {+11} FieldType: AnsiChar; { Тип поля C,D,L,N,M}
    {+12} FieldDisp: LongInt; { Смещение поля внутри записи }
    {+16} FieldLen, { Длина поля }
    {+17} FieldDec: Byte; { Длина поля после точки }
    {+18} A1: array[1..13] of Byte; { Резерв }
    {+31} FlagTagMDX: Byte; { Флаг тега многоиндексного
                                                 файла MDX, 1 - поле индек-
                                                 сированно, 0 - нет }
  end;

{ TDescriptor class declaration }

  TDescriptor = class { *** Элемент (дескриптор) заголовка *** }
  private
    FieldName: TFieldName; { Имя поля }
    FieldType: AnsiChar; { Тип поля C,D,L,N,M}
    FieldDisp: LongInt; { Смещение поля внутри записи }
    FieldLen, { Длина поля }
    FieldDec: Byte; { Длина поля после точки }
    FlagTagMDX: Byte; { Флаг тега многоиндексного файла MDX }
  public
    constructor Create(DRec: RDescriptor);
    procedure GetData(var DRec: RDescriptor);
    function GetName: AnsiString;
    function GetType: AnsiChar;
  end;

  TBaseElem = class
  private
    ElName: String[12];
    ElType: AnsiChar;
    ElLen: Integer;
    ElDec: Integer;
  public
    constructor Create(const AName: AnsiString; AType: AnsiChar; ALen: Integer; ADec: Integer);
    destructor Destroy; override;
  end;

 {class TBaseList declaration }

  TBaseList = class
  private
    InitF: boolean;
    cbResult: Longint;
    List: TList;
    function GetElem(Index: Integer): TBaseElem;
    function GetCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Add(ABaseElem: TBaseElem);
    property Count: Integer read GetCount;
    property Elems[Index: Integer]: TBaseElem read GetElem;
  end;

{ class TDBFRead declaration }

  TMemoType = (mtNone, mtDBT, mtFPT);

  TDBFRead = class
  private
    FlagInit: Byte;
    DBFName: string;
    Res: Longint;
    MemoFile: TFileStream;
    MemoHeader: RMemoHeader;
    FPTMemoHeader: RFPTMemoHeader;
    DBFHandle: Integer; { Hомеp хэндлеpа файла базы данных }
    DBFHeader: RFileHeader; { Буфер главной записи заголовка файла }
    Descriptor: RDescriptor; { Буфер дескриптора записи }
    DList: TList; { Список указателей дескрипторов}
    FRecNo: Integer;
    FDeleted: boolean;
    FMemoType: TMemoType;
    function GetMemoData(BlockNo: Longint): AnsiString;
    function StrBuffToNum(const ABuff: AnsiString): Integer;
  protected
    function GetVersion: byte;
    function GetRCount: Integer;
    function GetFCount: Integer;
    function GetEof: boolean;
    function GetFieldName(Index: Integer): AnsiString;
    function GetFieldType(Index: Integer): AnsiChar;
    function GetFieldLength(Index: Integer): Integer;
    function GetFieldDec(Index: Integer): Integer;
    function GetFieldDisp(Index: Integer): Integer;
  public
    constructor Create(const DBName: string);
    destructor Destroy; override;
    procedure ErrorHandle(Error: Longint); virtual;

    function GetNumField(const Name: AnsiString): Integer;
    function GetType(const Name: AnsiString): AnsiChar;
    function GetLength(const Name: AnsiString): Integer;
    function GetDecName(const Name: AnsiString): Integer;
    function GetDisp(const Name: AnsiString): Integer;

    function GetData(FieldNum: Integer): AnsiString;

    property Version: byte read GetVersion;
    property RecordCount: Integer read GetRCount;
    property FieldCount: Integer read GetFCount;
    property RecNo: Integer read FRecNo;
    property Deleted: boolean read FDeleted;
    property Eof: boolean read GetEof;
    property FileName: string read DBFName;

    property FieldName[Index: Integer]: AnsiString read GetFieldName;
    property FieldType[Index: Integer]: AnsiChar read GetFieldType;
    property FieldLength[Index: Integer]: Integer read GetFieldLength;
    property FieldDec[Index: Integer]: Integer read GetFieldDec;
    property FieldDisp[Index: Integer]: Integer read GetFieldDisp;
    property MemoType: TMemoType read FMemoType write FMemoType;
  end;

const

  erInvalidHandle = $06; { Hевеpный HANDLE-p файла }
  erSharedViolation = $20; { Ошибка pазделения файла }
  erLockingViolation = $21; { Ошибка блокиpовки файла }
  erSharingBufferOverflov = $24; { Пеpеполнение стека SHARE }

function DBFTypeToStr(DBFType: Char): string;

implementation

uses
  QImport3Common;

function ExtractFName(const FileName: TFileName):string;
var
{$IFDEF VCL5}
  fpath, fname: string;
{$ELSE}
  i: integer;
{$ENDIF}
begin
  Result := '';
{$IFDEF VCL5}
  fpath := ExtractFilePath(FileName);
  fname := ExtractFileName(FileName);
  Delete(fname, Pos('.', fname), MaxInt);
  Result := fpath + fname;
{$ELSE}
  i := 1;
  while (i < Length(FileName)) and (FileName[i] <> '.') do
  begin
    Result := Result + FileName[i];
    Inc(i);
  end;
{$ENDIF}
end;

function DBFTypeToStr(DBFType: char): string;
begin
  case DBFType of
    dftNumber: Result := 'Number';
    dftString: Result := 'AnsiString';
    dftBoolean: Result := 'Boolean';
    dftMemo: Result := 'Memo';
    dftFloat: Result := 'Float';
    dftDate: Result := 'Date';
  else Result := 'Unknown';
  end;
end;

function TDBFRead.GetRCount: Integer;
begin
  Result := DBFHeader.RecCount;
end;

function TDBFRead.GetVersion: byte;
begin
  Result := DBFHeader.DBType;
end;

function TDBFRead.GetFCount: Integer;
begin
  if Assigned(DList) then Result := DList.Count
  else Result := -1;
end;

function TDBFRead.GetData(FieldNum: Integer): AnsiString;
var
  a: byte;
  k: Integer;
  buf: AnsiString;
  fType: AnsiChar;
  mBlock: Integer;
begin
  Result := '';
  k := FieldLength[FieldNum];
  fType := FieldType[FieldNum];
  SetLength(buf, k);
  if FieldNum = 0 then
  begin
    FileRead(DBFHandle, a, 1); // skip first symbol of record
    FDeleted := a = $2A;
    inc(FRecNo);
  end;
  FileRead(DBFHandle, buf[1], k);
  if fType <> 'M' then
    if (DBFHeader.DBType = VisualFoxPro) And (fType = 'I') then
      Result := AnsiString(IntToStr(StrBuffToNum(buf)))
    else
      Result := buf
  else
  begin
    if DBFHeader.DBType = VisualFoxPro then
      mBlock := StrBuffToNum(buf)
    else
      mBlock := StrToIntDef({$IFDEF VCL12}string{$ENDIF}(buf), 0);
    if mBlock > 0 then
      Result := GetMemoData(mBlock)
  end;
end;

function TDBFRead.GetMemoData(BlockNo: Longint): AnsiString;
const
  BufSize = 512;
  EndMemoFlag: AnsiChar = #26;
var
  Buffer: AnsiString;
  TempBuff: array[1..512] of AnsiChar;
  EndFlag: Boolean;
  BlockStart: Integer;
  EndMemoFlagPos: Integer;
  FPTMemoBlockHeader: RFPTMemoBlockHeader;
  BlockSignatureArray, MemoLengthArray: array[0..3] of byte;
  dw: DWORD; i: Integer;
  ReadBufSize,
  CopyBuffSize: Word;
begin
  if not Assigned(MemoFile) then
  begin
    Result := QIGetEmptyStr;
    Exit;
  end;

  try
    if FMemoType = mtDBT then
    begin
      Buffer := '';
      MemoFile.Seek(BlockNo * BufSize, soFromBeginning);
      repeat
        ReadBufSize := MemoFile.Read(TempBuff, BufSize);
        EndFlag := ReadBufSize < BufSize;
        EndMemoFlagPos := Pos(EndMemoFlag, AnsiString(TempBuff));
        EndFlag := EndFlag or (EndMemoFlagPos > 0) and
          (EndMemoFlagPos < (BufSize - 1)) and
          (TempBuff[EndMemoFlagPos + 1] in [EndMemoFlag, #0]);
        if EndMemoFlagPos > 0 then
          CopyBuffSize := EndMemoFlagPos - 1
        else
          CopyBuffSize := ReadBufSize;
        Buffer := Buffer + Copy(TempBuff, 1, CopyBuffSize);
      until EndFlag;

      {$WARNINGS OFF}
        Result := StringReplace(Buffer, #141, #13, [rfReplaceAll]);
      {$WARNINGS OFF}
    end
    else
    begin // .fpt file
      BlockStart := BlockNo * FPTMemoHeader.BlockSize;
      MemoFile.Seek(BlockStart, soFromBeginning);
      MemoFile.Read(BlockSignatureArray[0], 4);
      dw := 0;
      for i := 0 to 3 do
      begin
        dw := dw + BlockSignatureArray[i] shl (Abs(i - 3) * 8);
      end;
      FPTMemoBlockHeader.BlockSignature := dw;
      MemoFile.Read(MemoLengthArray[0], 4);
      dw := 0;
      for i := 0 to 3 do
      begin
        dw := dw + MemoLengthArray[i] shl (Abs(i - 3) * 8);
      end;
      FPTMemoBlockHeader.MemoLength := dw;
      if FPTMemoBlockHeader.BlockSignature = 0 then
      begin
        Result := '(Picture)';
        exit;
      end;
      SetLength(Buffer, FPTMemoBlockHeader.MemoLength);
      MemoFile.Read(Buffer[1], FPTMemoBlockHeader.MemoLength);
      Result := Buffer;
    end;
  except
    Result := 'Error!!!';
  end;
end;

function TDBFRead.GetEof: boolean;
begin
  Result := FRecNo > RecordCount;
end;

constructor TDescriptor.Create;
begin
  FieldName := DRec.FieldName;
  FieldType := UpCase(DRec.FieldType);
  FieldDisp := DRec.FieldDisp;
  FieldLen := DRec.FieldLen;
  FieldDec := DRec.FieldDec;
  FlagTagMDX := DRec.FlagTagMDX;
end;

procedure TDescriptor.GetData;
begin
  DRec.FieldName := FieldName;
  DRec.FieldEnd := #0;
  DRec.FieldType := FieldType;
  DRec.FieldDisp := FieldDisp;
  DRec.FieldLen := FieldLen;
  DRec.FieldDec := FieldDec;
  DRec.FlagTagMDX := FlagTagMDX;
end;

function TDescriptor.GetName;
var
  s: AnsiString;
  i: Integer;
begin
  s := '';
  for i := 1 to 10 do
  begin
    if FieldName[i] = #0 then
      Break
    else
      s := s + FieldName[i];
  end;
  Result := {$IFDEF VCL12}AnsiString{$ENDIF}(AnsiUpperCase(Trim( {$IFDEF VCL12}string{$ENDIF}(S))));
end;

function TDescriptor.GetType;
begin
  Result := UpCase(FieldType);
end;

constructor TBaseElem.Create(const AName: AnsiString; AType: AnsiChar; ALen: Integer; ADec: Integer);
begin
  inherited Create;
  ElName := AName;
  ElType := AType;
  ElLen := ALen;
  ElDec := ADec;
end;

destructor TBaseElem.Destroy;
begin
  ElName := '';
  ElType := #0;
  ElLen := 0;
  ElDec := 0;
  inherited Destroy;
end;

constructor TBaseList.Create;
begin
  List := TList.Create;
  InitF := True;
  cbResult := 0;
end;

destructor TBaseList.Destroy;
var
  i: Integer;
begin
  if InitF then
  begin
    if Assigned(List) then
    begin
      for i := 0 to List.Count - 1 do
      begin
        if Assigned(List[i]) then
          TBaseElem(List[i]).Free;
      end;
      List.Free;
    end;
  end;
  cbResult := 0;
  inherited Destroy;
end;

procedure TBaseList.Add(ABaseElem: TBaseElem);
begin
  if Assigned(ABaseElem) then
    List.Add(ABaseElem);
end;

function TBaseList.GetElem(Index: Integer): TBaseElem;
begin
  Result := nil;
  if InitF then
    if (Index < 0) or (Index < List.Count) then
      Result := TBaseElem(List.Items[Index])
    else
      Result := nil;
end;

function TBaseList.GetCount: Integer;
begin
  if InitF then
    cbResult := List.Count
  else
    cbResult := 0;
  Result := cbResult;
end;

constructor TDBFRead.Create;
var
  i: Integer;
  CurrentDisp: Word;
  PD: TDescriptor;
  FPath, FName, FExt: string;

  MemoFileName: string;

  NextBlockArray: array[0..3] of byte;
  BlockSizeArray: array[0..1] of byte;
  dw: DWORD; w: WORD;

begin
  DList := nil;
  FPath := ExtractFilePath(DBName);
  FName := ExtractFileName(DBName);
  FExt := ExtractFileExt(DBName);

  if FExt = '' then
    DBFName := FPath + FName + '.DBF'
  else if AnsiUpperCase(FExt) <> '.DBF' then
    DbfName := FPath + ChangeFileExt(FName, '.DBF')
  else
    DBFName := FPath + FName;

  DBFHandle := FileOpen(DBFName, fmOpenRead); // open the file
  if DBFHandle = -1 then
  begin
    FlagInit := 0;
    Res := dbfEOpen;
    ErrorHandle(Res);
    Exit;
  end;

  Res := FileRead(DBFHandle, DBFHeader, 32);
  if Res = 0 then
  begin
    FlagInit := 0;
    Res := dbfFileEmpty;
    Exit;
  end
  else
    if Res = -1 then
    begin
      FlagInit := 0;
      FileClose(DBFHandle);
      Res := dbfERead;
      ErrorHandle(Res);
      Exit;
    end
    else
      if Res <> SizeOf(DBFHeader) then
      begin
        FlagInit := 0;
        FileClose(DBFHandle);
        Res := dbfEDestruct;
        ErrorHandle(Res);
        Exit;
      end;

  CurrentDisp := 1;

  DList := TList.Create;
  for i := 1 to Pred(DBFHeader.HeaderSize div 32) do
  begin
    Res := FileRead(DBFHandle, Descriptor, 32);
    if Res = -1 then
    begin
      FlagInit := 0;
      if Assigned(DList) then
      begin
        DList.Free;
        DList := nil;
      end;
      FileClose(DBFHandle);
      Res := dbfERead;
      ErrorHandle(Res);
      Exit
    end;
    with Descriptor do
    begin
      FieldDisp := CurrentDisp;
      Inc(CurrentDisp, FieldLen);
    end;
    PD := TDescriptor.Create(Descriptor);
    DList.Add(PD);
  end;
  FlagInit := 123;
  FRecNo := 1;
  FileSeek(DBFHandle, DBFHeader.HeaderSize, 0);
  FMemoType := mtNone;
  if DBFHeader.DBType = $83 then
  begin
    MemoFileName := ExtractFName(DBName) + '.dbt';
    if FileExists(MemoFileName) then
    begin
      MemoFile := TFileStream.Create(MemoFileName, fmOpenRead);
      MemoFile.Seek(0, soFromBeginning);
      MemoFile.ReadBuffer(MemoHeader, SizeOf(RMemoHeader));
      FMemoType := mtDBT;
    end
  end
  else begin
    MemoFileName := ExtractFName(DBName) + '.fpt';
    if FileExists(MemoFileName) then
    begin
      MemoFile := TFileStream.Create(MemoFileName, fmOpenRead);
      MemoFile.Seek(0, soFromBeginning);
      MemoFile.Read(NextBlockArray[0], 4);
      dw := 0;
      for i := 0 to 3 do
      begin
        dw := dw + NextBlockArray[i] shl (Abs(i - 3) * 8); //(i - 3) pai;
      end;
      FPTMemoHeader.NextBlock := dw;
      MemoFile.Read(BlockSizeArray[0], 2);
      FPTMemoHeader.Unused := 0;
      MemoFile.Read(w, 2);
      FPTMemoHeader.BlockSize := Swap(w);
      FillChar(FPTMemoHeader.Reserve, SizeOf(FPTMemoHeader.Reserve), 0);
      FMemoType := mtFPT;
    end
  end;
end;

destructor TDBFRead.Destroy;
var
  i: Integer;
begin
  if Assigned(DList) then
  begin
    for i := 0 to DList.Count - 1 do
      if Assigned(DList[i]) then
        TDescriptor(DList[i]).Free;
    DList.Free;
  end;
  FileClose(DBFHandle);
  FlagInit := 0;
  if Assigned(MemoFile) then MemoFile.Free;
  inherited Destroy;
end;

procedure TDBFRead.ErrorHandle;
var
  S: string;
begin
  case Error of
    dbfFileEmpty: S := 'File is empty!';
    dbfENoFileName: S := 'File name not assigned!';
    dbfEFileNotFound: S := 'File not found!';
    dbfECreate: S := 'File create error!';
    dbfEOpen: S := 'File open error';
    dbfESeek: S := 'File seek error!';
    dbfERead: S := 'File read error!';
    dbfEWrite: S := 'File write error!';
    dbfEOutOfMemory: S := 'Out of memory!';
    dbfEDestruct: S := 'Table file is invalid!';
    dbfECreateList: S := 'Error creating field descriptor list!';
    dbfEOutOfNumber: S := 'Out of range!';
    else S := 'Unknown error!';
  end;
  raise Exception.Create(S);
end;

function TDBFRead.GetNumField(const Name: AnsiString): Integer;
var
  S0: AnsiString;
  i, ei: Integer;
begin
  ei := DList.Count - 1;
  if (FlagInit = 123) then
  begin
    S0 := {$IFDEF VCL12}AnsiString{$ENDIF}(AnsiUpperCase(Trim( {$IFDEF VCL12}string{$ENDIF}(Name))));
    if S0 <> '' then
      for i := 0 to ei do
        with DList do
        begin
          if S0 = TDescriptor(Items[i]).GetName then
          begin
            Result := I; Exit;
          end;
        end;
  end;
  Result := -1;
end;

function TDBFRead.GetFieldName;
begin
  if (Index > -1) and (Index <= DList.Count - 1) then
    Result := TDescriptor(DList.Items[Index]).GetName
  else Result := '';
end;

function TDBFRead.GetFieldType(Index: Integer): AnsiChar;
begin
  Result := #0;
  if (FlagInit <> 123) then Exit;
  if (Index < 0) or (Index > FieldCount - 1) then
  begin
    Result := #0; Exit;
  end;
  Result := TDescriptor(DList.Items[Index]).GetType;
end;

function TDBFRead.GetType(const Name: AnsiString): AnsiChar;
var
  NF: Integer;
begin
  Result := #0;
  if (FlagInit <> 123) then Exit;
  NF := GetNumField(Name);
  if NF = -1 then Exit;
  Result := TDescriptor(DList.Items[NF]).GetType;
end;

function TDBFRead.GetFieldLength(Index: Integer): Integer;
begin
  Result := 0;
  if (FlagInit <> 123) then Exit;
  if Index > DList.Count - 1 then Exit;
  Result := TDescriptor(DList.Items[Index]).FieldLen;
end;

function TDBFRead.GetLength(const Name: AnsiString): Integer;
var
  NF: Integer;
begin
  Result := 0;
  if (FlagInit <> 123) then Exit;
  NF := GetNumField(Name);
  if NF = -1 then Exit;
  Result := TDescriptor(DList.Items[NF]).FieldLen;
end;

function TDBFRead.GetDecName(const Name: AnsiString): Integer;
var
  NF: Integer;
begin
  Result := 0;
  if (FlagInit <> 123) then Exit;
  NF := GetNumField(Name);
  if NF = -1 then Exit;
  Result := TDescriptor(DList.Items[NF]).FieldDec;
end;

function TDBFRead.GetFieldDec(Index: Integer): Integer;
begin
  Result := -1;
  if (FlagInit <> 123) then Exit;
  if (Index < 0) or (Index > DList.Count - 1) then
    ErrorHandle(dbfEOutOfNumber);
  Result := TDescriptor(DList.Items[Index]).FieldDec;
end;

function TDBFRead.GetFieldDisp(Index: Integer): Integer;
begin
  Result := -1;
  if (FlagInit <> 123) then Exit;
  if (Index < 0) or (Index > DList.Count - 1) then
    ErrorHandle(dbfEOutOfNumber);
  Result := TDescriptor(DList.Items[Index]).FieldDisp;
end;

function TDBFRead.GetDisp(const Name: AnsiString): Integer;
var
  NF: Integer;
begin
  Result := 0;
  if (FlagInit <> 123) then Exit;
  NF := GetNumField(Name);
  if NF = -1 then Exit;
  Result := TDescriptor(DList.Items[NF]).FieldDisp;
end;

function TDBFRead.StrBuffToNum(const ABuff: AnsiString): Integer;
var
  i: Integer;
  b: Byte;
begin
  Result := 0;
  for i := 1 to Length(ABuff) do
  begin
    b := Byte(ABuff[i]);
    Result := Result + (b shl ((i - 1) * 8));
  end;
end;

end.
