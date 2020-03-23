unit QImport3XLSFile;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.Classes,
    System.SysUtils,
    System.Variants,
  {$ELSE}
    Classes,
    SysUtils,
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
  {$ENDIF}
  QImport3XLSCommon;

type
  TxlsSection = class;
  TbiffXFList = class;
  TbiffFormatList = class;
  TbiffSSTList = class;

  TbiffContinue = class;

  TbiffRecord = class
  private
    FSection: TxlsSection;
    FID: word;
    FDataSize: word;
    FData: PByteArray;
    FContinue: TbiffContinue;
    FOnDestroy: TNotifyEvent;

    function GetXFList: TbiffXFList;
    function GetFormatList: TbiffFormatList;
    function GetSSTList: TbiffSSTList;
  public
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); virtual;
    destructor Destroy; override;
    procedure AddContinue(const Continue: TbiffContinue);

    property Section: TxlsSection read FSection;
    property XFList: TbiffXFList read GetXFList;
    property FormatList: TbiffFormatList read GetFormatList;
    property SSTList: TbiffSSTList read GetSSTList;

    property ID: word read FID;
    property DataSize: word read FDataSize write FDataSize;
    property Data: PByteArray read FData;

    property Continue: TbiffContinue read FContinue;

    property OnDestroy: TNotifyEvent read FOnDestroy write FOnDestroy;
  end;

  TbiffContinue = class(TbiffRecord);

  TbiffBOF = class(TbiffRecord)
  private
    function GetBOFType: word;
  public
    property BOFType: word read GetBOFType;
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); override;
  end;

  TbiffEOF = class(TbiffRecord);

  TbiffBoundSheet = class(TbiffRecord)
  private
    function GetName: WideString;
    function GetOptionFlags: word;
  public
    property Name: WideString read GetName;
    property OptionFlags: word read GetOptionFlags;
  end;

  TbiffString = class(TbiffRecord)
  private
    function GetValue: WideString;
  public
    property Value: WideString read GetValue;
  end;

  TbiffColRow = class(TbiffRecord)
  private
    function GetCol: word;
    procedure SetCol(Value: word);
    function GetRow: word;
    procedure SetRow(Value: word);
  public
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); override;

    property Col: word read GetCol write SetCol;
    property Row: word read GetRow write SetRow;
  end;

  TbiffCellType = (bctString, bctBoolean, bctNumeric, bctDateTime, bctUnknown);

  TxlsWorkSheet = class;

  TbiffCell = class(TbiffColRow)
  private
    FWorkSheet: TxlsWorkSheet;
    function GetXFIndex: word;
    procedure SetXFIndex(Value: word);
    function GetFormatIndex: word;
    function GetCellName: string;
  protected
    function GetCellType: TbiffCellType; virtual;
    function GetIsFormula: boolean; virtual;
    function GetIsString: boolean;
    function GetIsBoolean: boolean;
    function GetIsFloat: boolean;
    function GetIsDateTime: boolean;
    function GetIsDateOnly: boolean;
    function GetIsTimeOnly: boolean;
    function GetIsVariant: boolean;
    function GetAsString: WideString; virtual;
    procedure SetAsString(const Value: WideString); virtual;
    function GetAsBoolean: boolean; virtual;
    procedure SetAsBoolean(Value: boolean); virtual;
    function GetAsFloat: double; virtual;
    procedure SetAsFloat(Value: double); virtual;
    function GetAsDateTime: TDateTime; virtual;
    procedure SetAsDateTime(Value: TDateTime); virtual;
    function GetAsVariant: variant; virtual;
    procedure SetAsVariant(Value: variant); virtual;
  public
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); override;

    property WorkSheet: TxlsWorkSheet read FWorkSheet;
    property XFIndex: word read GetXFIndex write SetXFIndex;
    property FormatIndex: word read GetFormatIndex;

    property CellType: TbiffCellType read GetCellType;

    property IsFormula: boolean read GetIsFormula;
    property IsString: boolean read GetIsString;
    property IsBoolean: boolean read GetIsBoolean;
    property IsFloat: boolean read GetIsFloat;
    property IsDateTime: boolean read GetIsDateTime;
    property IsDateOnly: boolean read GetIsDateOnly;
    property IsTimeOnly: boolean read GetIsTimeOnly;
    property IsVariant: boolean read GetIsVariant;

    property AsString: WideString read GetAsString write SetAsString;
    property AsBoolean: boolean read GetAsBoolean write SetAsBoolean;
    property AsFloat: double read GetAsFloat write SetAsFloat;
    property AsDateTime: TDateTime read GetAsDateTime write SetAsDateTime;
    property AsVariant: variant read GetAsVariant write SetAsVariant;
    property Value: variant read GetAsVariant write SetAsVariant;
    property CellName: string read GetCellName;
  end;

  TbiffBlank = class(TbiffCell);

  TbiffBoolErr = class(TbiffCell)
  protected
    function GetCellType: TbiffCellType; override;

    function GetAsBoolean: boolean; override;
    procedure SetAsBoolean(Value: boolean); override;
    function GetAsVariant: variant; override;
    procedure SetAsVariant(Value: variant); override;
    function GetAsString: WideString; override;
    procedure SetAsString(const Value: WideString); override;
  end;

  TbiffNumber = class(TbiffCell)
  protected
    function GetCellType: TbiffCellType; override;

    function GetAsFloat: double; override;
    procedure SetAsFloat(Value: double); override;
    function GetAsDateTime: TDateTime; override;
    procedure SetAsDateTime(Value: TDateTime); override;
    function GetAsVariant: variant; override;
    procedure SetAsVariant(Value: variant); override;
    function GetAsString: WideString; override;
    procedure SetAsString(const Value: WideString); override;
  end;

  TbiffRK = class(TbiffCell)
  protected
    function GetCellType: TbiffCellType; override;

    function GetAsFloat: double; override;
    procedure SetAsFloat(Value: double); override;
    function GetAsDateTime: TDateTime; override;
    procedure SetAsDateTime(Value: TDateTime); override;
    function GetAsVariant: variant; override;
    procedure SetAsVariant(Value: variant); override;
    function GetAsString: WideString; override;
    procedure SetAsString(const Value: WideString); override;
  end;

  TbiffLabel = class(TbiffCell)
  protected
    function GetCellType: TbiffCellType; override;
    function GetAsString: WideString; override;
    function GetAsVariant: Variant; override;
  end;

  TxlsCharSize = 1..2;

  TxlsString = class
    FIsWideStr: boolean;

    FStrLen: word;
    FOptionFlags: byte;
    FWideData: WideString;
    FShortData: AnsiString;

    FRTFNumber: word;
    FRTFData: PByteArray;
    FFarEastDataSize: Longword;
    FFarEastData: PByteArray;

    function GetLenOfLen: byte;
    function GetHasWideChar: boolean;
    function GetHasRichText: boolean;
    function GetFarEast: boolean;
    function GetValue: WideString;
  public
    constructor CreateR(IsWideStr: boolean; var ARecord: TbiffRecord;
      var Offset: integer);
    constructor CreateWS(IsWideStr: boolean; const Str: WideString);
    destructor Destroy; override;
    function Compare(Str: TxlsString): integer; //-1 if less, 0 if equal, 1 if more
    property OptionFlags: byte read FOptionFlags;
    property ShortData: AnsiString read FShortData;
    property WideData: WideString read FWideData;

    property LenOfLen: byte read GetLenOfLen;
    property HasWideChar: boolean read GetHasWideChar;
    property HasRichText: boolean read GetHasRichText;
    property HasFarEast: boolean read GetFarEast;
    property Value: WideString read GetValue;
  end;

  TxlsSSTEntry = class (TObject)
  private
    FRefCount: integer;
    FValue: TxlsString;
    FOnDestroy: TNotifyEvent;

  public
    constructor CreateXS(Str: TxlsString);
    constructor CreateWS(Str: WideString);
    destructor Destroy; override;

    procedure IncRef;
    procedure DecRef;

    property RefCount: integer read FRefCount;
    property Value: TxlsString read FValue;

    property OnDestroy: TNotifyEvent read FOnDestroy write FOnDestroy;
  end;

  TbiffLabelSST = class(TbiffCell{, IInterface})
  private
    FSSTEntry: TxlsSSTEntry;
    procedure DestroySSTEntry(Sender: TObject);
  protected
    function GetCellType: TbiffCellType; override;

    function GetAsString: WideString; override;
    procedure SetAsString(const Value: WideString); override;
    function GetAsVariant: variant; override;
    procedure SetAsVariant(Value: variant); override;
  public
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); override;
    destructor Destroy; override;
  end;

  ETokenException = class(Exception)
  private
    FToken: integer;
  public
    constructor Create(Token: integer; const Msg: string);
    property Token: integer read FToken;
  end;

  TbiffFormula = class(TbiffCell)
  private
    FValue: variant;

    function GetIsExp: boolean;
    function GetKey: cardinal;

    procedure ArrangeSharedFormulas;
  protected
    function GetIsFormula: boolean; override;
    function GetAsVariant: variant; override;
    procedure SetAsVariant(Value: variant); override;
    function GetAsString: WideString; override;
    procedure SetAsString(const Value: WideString); override;
  public
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); override;
    procedure MixShared(SharedData: PByteArray; SharedDataSize: integer);

    property IsExp: boolean read GetIsExp;
    property Key: cardinal read GetKey;
  end;

  TbiffShrFmla = class(TbiffRecord)
  private
    function GetFirstRow: word;
    function GetLastRow: word;
    function GetFirstCol: word;
    function GetLastCol: word;
    function GetKey: integer;
  public
    property FirstRow: word read GetFirstRow;
    property LastRow: word read GetLastRow;
    property FirstCol: word read GetFirstCol;
    property LastCol: word read GetLastCol;
    property Key: integer read GetKey;
  end;

  TbiffName = class(TbiffRecord)
  private
    function GetName: WideString;
    function GetNameLength: byte;
    function GetNameSize: integer;
    function GetOptionFlags: byte;
    function GetRow1: integer;
    function GetRow2: integer;
    function GetCol1: integer;
    function GetCol2: integer;
  public
    property Name: WideString read GetName;
    property NameLength: byte read GetNameLength;
    property NameSize: integer read GetNameSize;
    property OptionFlags: byte read GetOptionFlags;
    property Row1: integer read GetRow1;
    property Row2: integer read GetRow2;
    property Col1: integer read GetCol1;
    property Col2: integer read GetCol2;
  end;

  TbiffMultiple = class(TbiffRecord)
  protected
    FCol: integer;
    function GetEOF: boolean; virtual; abstract;
    function GetCell: TbiffCell; virtual; abstract;
  public
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); override;
    property Eof: boolean read GetEOF;
    property Cell: TbiffCell read GetCell;
  end;

  TbiffMulBlank = class(TbiffMultiple)
  protected
    function GetEOF: boolean; override;
    function GetCell: TbiffCell; override;
  end;

  TbiffMulRK = class(TbiffMultiple)
  protected
    function GetEOF: boolean; override;
    function GetCell: TbiffCell; override;
  end;

  TbiffFont = class(TbiffRecord);

  TbiffXF = class(TbiffRecord)
  private
    function GetFormatIndex: word;
    procedure SetFormatIndex(Value: word);
  public
    property FormatIndex: word read GetFormatIndex
      write SetFormatIndex;
  end;

  TbiffFormat = class(TbiffRecord)
  private
    FID: word;
    FValue: WideString;
  public
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); override;

    property ID: word read FID;
    property Value: WideString read FValue;
  end;

  TbiffSST = class(TbiffRecord)
  private
    FCount: integer;
  public
    property Count: integer read FCount;
    constructor Create(Section: TxlsSection; ID, DataSize: word;
      Data: PByteArray); override;
  end;

  TxlsWorkbook = class;

  TxlsList = class(TList)
  private
    FWorkbook: TxlsWorkbook;
    function GetItems(Index: integer): TObject;
    procedure SetItems(Index: integer; Value: TObject);
  public
    function Add(Item: TObject): integer;
    constructor Create(Workbook: TxlsWorkbook);
    procedure Clear; reintroduce; override;
    procedure Delete(Index: integer);
    {$IFDEF VCL5}
    function Extract(Item: TObject): TObject;
    {$ENDIF}
    function First: TObject;
    function IndexOf(Item: TObject): integer;
    procedure Insert(Index: integer; Item: TObject);
    function Last: TObject;
    function Remove(Item: TObject): integer;

    property Items[Index: integer]: TObject read GetItems write SetItems;
    property Workbook: TxlsWorkbook read FWorkbook;
  end;

  TbiffRecordList = class(TxlsList)
  private
    FTotalSize: integer;
    function GetItems(Index: integer): TbiffRecord;
    procedure SetItems(Index: integer; Value: TbiffRecord);
  public
    function Add(Item: TbiffRecord): integer;
    procedure Insert(Index: integer; Item: TbiffRecord);
    procedure CorrectSize(Delta: integer);
    procedure RecalculateTotalSize;

    property Items[Index: integer]: TbiffRecord read GetItems
      write SetItems; default;
  end;

  TbiffColRowList = class(TbiffRecordList)
  private
    function GetItems(Index: integer): TbiffColRow;
    procedure SetItems(Index: integer; Value: TbiffColRow);
  public
    property Items[Index: integer]: TbiffColRow read GetItems write SetItems; default;
    function Add(Item: TbiffColRow): integer;
    procedure Insert(Index: integer; Item: TbiffColRow);
  end;

  TbiffColRowListClass = class of TbiffColRowList;

  TbiffCellList = class(TbiffColRowList)
  private
    FSorted: boolean;
    function GetItems(Index: integer): TbiffCell;
    procedure SetItems(Index: integer; Value: TbiffCell);
    procedure OnDestroyItem(Sender: TObject);
  protected
    procedure SetMinAndMaxCells(Item: TbiffCell); virtual;
    procedure SetColRowNumber(Item: TbiffCell); virtual;
    property Sorted: boolean read FSorted write FSorted;
  public
    procedure Clear; override;
    function Add(Item: TbiffCell): integer;
    procedure Insert(Index: integer; Item: TbiffCell);

    property Items[Index: integer]: TbiffCell read GetItems
      write SetItems; default;
  end;

  TxlsRow = class(TbiffCellList)
  private
    FRowNumber: integer;
    FMinCol: integer;
    FMaxCol: integer;
  protected
    procedure SetMinAndMaxCells(Item: TbiffCell); override;
    procedure SetColRowNumber(Item: TbiffCell); override;
  public
    constructor Create(Workbook: TxlsWorkbook);
    function Find(Col: integer; var Index: integer): boolean;
    procedure Sort;

    property RowNumber: integer read FRowNumber;
    property MinCol: integer read FMinCol;
    property MaxCol: integer read FMaxCol;
  end;

  TxlsCol = class(TbiffCellList)
  private
    FColNumber: integer;
    FMinRow: integer;
    FMaxRow: integer;
  protected
    procedure SetMinAndMaxCells(Item: TbiffCell); override;
    procedure SetColRowNumber(Item: TbiffCell); override;
  public
    constructor Create(Workbook: TxlsWorkbook);
    function Find(Row: integer; var Index: integer): boolean;
    procedure Sort;

    property ColNumber: integer read FColNumber;
    property MinRow: integer read FMinRow;
    property MaxRow: integer read FMaxRow;
  end;

  TbiffShrFmlaList = class(TbiffRecordList)
  private
    FSorted: boolean;

    function GetItems(Index: integer): TbiffShrFmla;
    procedure SetItems(Index: integer; Value: TbiffShrFmla);
  public
    function Add(Item: TbiffShrFmla): integer;
    procedure Insert(Index: integer; Item: TbiffShrFmla);
    function Find(Key: integer; var Index: integer): boolean;
    procedure Sort;

    property Items[Index: integer]: TbiffShrFmla read GetItems
      write SetItems; default;
  end;

  TbiffNameList = class(TbiffRecordList)
  private
    function GetItems(Index: integer): TbiffName;
    procedure SetItems(Index: integer; Value: TbiffName);
  public
    function Add(Item: TbiffName): integer;
    procedure Insert(Index: integer; Item: TbiffName);

    property Items[Index: integer]: TbiffName read GetItems
      write SetItems; default;
  end;

  TbiffBoundSheetList = class(TbiffRecordList)
  private
    function GetItems(Index: integer): TbiffBoundSheet;
    procedure SetItems(Index: integer; Value: TbiffBoundSheet);
    function GetName(Index: integer): WideString;
  public
    function Add(Item: TbiffBoundSheet): integer;
    procedure Insert(Index: integer; Item: TbiffBoundSheet);

    property Items[Index: integer]: TbiffBoundSheet read GetItems
      write SetItems; default;
    property Name[Index: integer]: WideString read GetName;
  end;

  TbiffXFList = class(TbiffRecordList)
  private
    function GetItems(Index: integer): TbiffXF;
    procedure SetItems(Index: integer; Value: TbiffXF);
  public
    function Add(Item: TbiffXF): integer;
    procedure Insert(Index: integer; Item: TbiffXF);

    property Items[Index: integer]: TbiffXF read GetItems
      write SetItems; default;
  end;

  TbiffFormatList = class(TbiffRecordList)
  private
    FSorted: boolean;

    function GetItems(Index: integer): TbiffFormat;
    procedure SetItems(Index: integer; Value: TbiffFormat);
    function GetFormat(ID: integer): WideString;
  public
    function Add(Item: TbiffFormat): integer;
    procedure Insert(Index: integer; Item: TbiffFormat);
    function Find(ID: integer; var Index: integer): boolean;
    procedure Sort;

    property Items[Index: integer]: TbiffFormat read GetItems write SetItems;
    property Format[Index: integer]: WideString read GetFormat; default;
  end;

  TbiffSSTList = class(TxlsList)
  private
    function GetItems(Index: integer): TxlsSSTEntry;
    procedure SetItems(Index: integer; Value: TxlsSSTEntry);
  public
    function Add(Item: TxlsSSTEntry): integer;
    procedure Insert(Index: integer; Item: TxlsSSTEntry);
    function Find(Str: TxlsString; var Index: integer): boolean;
    procedure Sort;
    procedure Load(SST: TbiffSST);
    function AddString(const Str: WideString): integer;

    property Items[Index: integer]: TxlsSSTEntry read GetItems write SetItems; default;
  end;

  TxlsRowList = class(TxlsList)
  private
    FSorted: boolean;
    function GetItems(Index: integer): TxlsRow;
    procedure SetItems(Index: integer; Value: TxlsRow);
  public
    constructor Create(Workbook: TxlsWorkbook);
    function Add(Row: TxlsRow): integer;
    procedure Insert(Index: integer; Row: TxlsRow);
    function Find(Row: integer; var Index: integer): boolean;
    procedure Sort;

    property Items[Index: integer]: TxlsRow read GetItems write SetItems; default;
  end;

  TxlsColList = class(TxlsList)
  private
    FSorted: boolean;
    function GetItems(Index: integer): TxlsCol;
    procedure SetItems(Index: integer; Value: TxlsCol);
  public
    constructor Create(Workbook: TxlsWorkbook);
    function Add(Col: TxlsCol): integer;
    procedure Insert(Index: integer; Col: TxlsCol);
    function Find(Col: integer; var Index: integer): boolean;
    procedure Sort;

    property Items[Index: integer]: TxlsCol read GetItems write SetItems; default;
  end;

  TxlsSection = class
  private
    FWorkbook: TxlsWorkbook;
    FBOF: TbiffBOF;
    FEOF: TbiffEOF;

    FOnDestroy: TNotifyEvent;
  public
    constructor Create(Workbook: TxlsWorkbook);
    destructor Destroy; override;
    procedure Load(Stream: TStream; BOF: TbiffBOF); virtual; abstract;
    procedure Clear;

    property Workbook: TxlsWorkbook read FWorkbook;
    property Bof: TbiffBOF read FBOF write FBOF;
    property Eof: TbiffEOF read FEOF write FEOF;

    property OnDestroy: TNotifyEvent read FOnDestroy write FOnDestroy;
  end;

  TxlsGlobals = class(TxlsSection)
  private
    FBoundSheetList: TbiffBoundSheetList;
    FNameList: TbiffNameList;
    FSSTList: TbiffSSTList;
    FXFList: TbiffXFList;
    FFormatList: TbiffFormatList;

    procedure FixFormats;
  public
    constructor Create(Workbook: TxlsWorkbook);
    destructor Destroy; override;
    procedure Clear;
    procedure Load(Stream: TStream; BOF: TbiffBOF); override;

    property BoundSheetList: TbiffBoundSheetList read FBoundSheetList;
    property NameList: TbiffNameList read FNameList;
    property SSTList: TbiffSSTList read FSSTList;
    property XFList: TbiffXFList read FXFList;
    property FormatList: TbiffFormatList read FFormatList;
  end;

  TxlsSheet = class(TxlsSection)
  private
    FIndex: integer;
    function GetName: WideString;
  public
    constructor Create(Workbook: TxlsWorkbook); virtual;
    property Name: WideString read GetName;
  end;

  TxlsSheetList = class(TxlsList)
  private
    function GetItems(Index: integer): TxlsSheet;
    procedure SetItems(Index: integer; Value: TxlsSheet);
  public
    function Add(Item: TxlsSheet): integer;
    procedure Insert(Index: integer; Item: TxlsSheet);

    property Items[Index: integer]: TxlsSheet read GetItems
      write SetItems; default;
  end;

  TxlsWorkSheet = class(TxlsSheet)
  private
    FRows: TxlsRowList;
    FCols: TxlsColList;
    FShrFmlaList: TbiffShrFmlaList;
    function GetRowCount: integer;
    function GetColCount: integer;
    function GetCells(Row, Col: integer): TbiffCell;
  public
    constructor Create(Workbook: TxlsWorkbook); override;
    destructor Destroy; override;
    procedure Clear;
    procedure Load(AStream: TStream; ABof: TbiffBOF); override;
    procedure LoadRows(AStream: TStream; ABof: TbiffBOF; ARowCount: Cardinal);
    procedure AddCell(Cell: TbiffCell);
    procedure AddMultiple(Multiple: TbiffMultiple);
    procedure FixFormulas;

    property Rows: TxlsRowList read FRows;
    property Cols: TxlsColList read FCols;
    property ShrFmlaList: TbiffShrFmlaList read FShrFmlaList;
    property RowCount: integer read GetRowCount;
    property ColCount: integer read GetColCount;
    property Cells[Row, Col: integer]: TbiffCell read GetCells;
  end;

  TxlsWorkSheetList = class(TxlsList)
  private
    function GetItems(Index: integer): TxlsWorkSheet;
    procedure SetItems(Index: integer; Value: TxlsWorkSheet);
    procedure OnDestroyItem(Sender: TObject);
  public
    function Add(Item: TxlsWorkSheet): integer;
    procedure Insert(Index: integer; Item: TxlsWorkSheet);
    function IndexOfName(const Name: WideString): integer;

    property Items[Index: integer]: TxlsWorkSheet read GetItems
      write SetItems; default;
  end;

  TxlsFile = class;

  TxlsWorkbook = class
  private
    FExcelFile: TxlsFile;
    FGlobals: TxlsGlobals;
    FSheets: TxlsSheetList;
    FWorkSheets: TxlsWorkSheetList;
  public
    constructor Create(ExcelFile: TxlsFile);
    destructor Destroy; override;
    procedure Load(Stream: TStream);
    procedure Clear;

    property ExcelFile: TxlsFile read FExcelFile;

    property Globals: TxlsGlobals read FGlobals;
    property Sheets: TxlsSheetList read FSheets;
    property WorkSheets: TxlsWorkSheetList read FWorkSheets;
  end;

  TFunctionEvent = procedure(Sender: TObject; const FunctionName: AnsiString;
    Arguments: Variant; var Result: Variant) of object;

  TxlsFile = class
  private
    FFileName: string;
    FStream: TMemoryStream;
    FWorkbook: TxlsWorkbook;
    FLoaded: Boolean;
    FRowCount: Cardinal; // for MapEditor

    FOnFunction: TFunctionEvent;

    procedure Open;
    procedure Close;
    procedure OpenStream;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Load;
    procedure LoadRows(ARowCount: Cardinal);

    procedure Clear;

    property Loaded: boolean read FLoaded;

    property FileName: string read FFileName write FFileName;
    property Workbook: TxlsWorkbook read FWorkbook;

    property OnFunction: TFunctionEvent read FOnFunction write FOnFunction;
  end;

implementation

uses
  {$IFDEF VCL16}
    System.Win.ComObj,
    Winapi.ActiveX,
    Winapi.Windows,
    Vcl.AxCtrls,
    Vcl.Dialogs,
    {$IFDEF VCL17}
      System.Types,
    {$ENDIF}
  {$ELSE}
    ComObj,
    ActiveX,
    Windows,
    AxCtrls,
    Dialogs,
  {$ENDIF}
  QImport3XLSConsts,
  QImport3XLSUtils,
  QImport3Common;

function CorrectDate(const Value: TDateTime): TDateTime;
//correct date for 01/01/1900
begin
  if Trunc(Value) = 1 then
    Result := Value + 1
  else
    Result := Value;
end;    

{ TbiffRecord }

constructor TbiffRecord.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
begin
  inherited Create;
  FSection := Section;
  FID := ID;
  FDataSize := DataSize;
  FData := Data;
end;

destructor TbiffRecord.Destroy;
begin
  if Assigned(FOnDestroy) then FOnDestroy(Self);
  if Assigned(FData) then FreeMem(FData);
  if Assigned(FContinue) then begin
    ObjFreeAndNil(FContinue);
  end;
  inherited;
end;

procedure TbiffRecord.AddContinue(const Continue: TbiffContinue);
begin
  if Assigned(FContinue) then
    raise ExlsFileError.Create(sInvalidContinue);
  FContinue := Continue;
end;

function TbiffRecord.GetXFList: TbiffXFList;
begin
  Result := Section.Workbook.Globals.XFList;
end;

function TbiffRecord.GetFormatList: TbiffFormatList;
begin
  Result := Section.Workbook.Globals.FormatList;
end;

function TbiffRecord.GetSSTList: TbiffSSTList;
begin
  Result := Section.Workbook.Globals.SSTList;
end;

{ TbiffBOF }

constructor TbiffBOF.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
begin
  inherited;
  if GetWord(Data, 0) <> BIFF_BOF_VER then
    raise Exception.Create(sInvalidVersion);
end;

function TbiffBOF.GetBOFType: word;
begin
  Result := GetWord(Data, 2);
end;

{ TbiffBoundSheet }

function TbiffBoundSheet.GetName: WideString;
var
  Str: TxlsString;
  Offset: integer;
  R: TbiffRecord;
begin
  Offset := 6;
  R := Self;
  Str := TxlsString.CreateR(false, R, Offset);
  try
    Result := Str.Value;
  finally
    Str.Free;
  end;
end;

function TbiffBoundSheet.GetOptionFlags: word;
begin
  Result := GetWord(Data, 4);
end;

{ TbiffString }

function TbiffString.GetValue: WideString;
var
  Str: TxlsString;
  Tmp: TbiffRecord;
  Offset: integer;
begin
  Tmp := Self;
  Offset := 0;
  Str := TxlsString.CreateR(true, Tmp, Offset);
  try
    Result := Str.Value;
  finally
    Str.Free;
  end;
end;

{ TbiffColRow }

constructor TbiffColRow.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
begin
  inherited;
  if DataSize < 4 then
    raise ExlsFileError.CreateFmt(sWrongExcelRecord, [ID]);
end;

function TbiffColRow.GetCol: word;
begin
  Result := GetWord(Data, 2);
end;

procedure TbiffColRow.SetCol(Value: word);
begin
  SetWord(Data, 2, Value);
end;

function TbiffColRow.GetRow: word;
begin
  Result := GetWord(Data, 0);
end;

procedure TbiffColRow.SetRow(Value: word);
begin
  SetWord(Data, 0, Value);
end;

{ TbiffCell }

constructor TbiffCell.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
begin
  inherited;
  if Section is TxlsWorkSheet
    then FWorkSheet := Section as TxlsWorkSheet
    else FWorkSheet := nil;
end;

function TbiffCell.GetXFIndex: word;
begin
  Result := GetWord(Data, 4);
end;

procedure TbiffCell.SetXFIndex(Value: word);
begin
  SetWord(Data, 4, Value);
end;

function TbiffCell.GetFormatIndex: word;
begin
  Result := Section.Workbook.Globals.XFList[XFIndex].FormatIndex;
end;

function TbiffCell.GetCellName: string;
begin
  Result := Col2Letter(Col) + Row2Number(Row);
end;

function TbiffCell.GetCellType: TbiffCellType;
begin
  Result := bctUnknown;
end;

function TbiffCell.GetIsFormula: boolean;
begin
  Result := false;
end;

function TbiffCell.GetIsString: boolean;
begin
  Result := GetCellType = bctString;
end;

function TbiffCell.GetIsBoolean: boolean;
begin
  Result := GetCellType = bctBoolean;
end;

function TbiffCell.GetIsFloat: boolean;
begin
  Result := GetCellType = bctNumeric;
end;

function TbiffCell.GetIsDateTime: boolean;
begin
  Result := GetCellType = bctDateTime;
end;

function TbiffCell.GetIsDateOnly: boolean;
begin
  Result := (GetCellType = bctDateTime) and (Frac(GetAsDateTime) = 0);
end;

function TbiffCell.GetIsTimeOnly: boolean;
begin
  Result := (GetCellType = bctDateTime) and (Int(GetAsDateTime) = 0);
end;

function TbiffCell.GetIsVariant: boolean;
begin
  Result := GetCellType = bctUnknown;
end;

function TbiffCell.GetAsString: WideString;
begin
  Result := EmptyStr;
end;

procedure TbiffCell.SetAsString(const Value: WideString);
begin
  raise ExlsFileError.CreateFmt(sCellAccessError, [CellName, 'String']);
end;

function TbiffCell.GetAsBoolean: boolean;
begin
  raise ExlsFileError.CreateFmt(sCellAccessError, [CellName, 'Boolean']);
end;

procedure TbiffCell.SetAsBoolean(Value: boolean);
begin
  raise ExlsFileError.CreateFmt(sCellAccessError, [CellName, 'Boolean']);
end;

function TbiffCell.GetAsFloat: double;
begin
  raise ExlsFileError.CreateFmt(sCellAccessError, [CellName, 'Float']);
end;

procedure TbiffCell.SetAsFloat(Value: double);
begin
  raise ExlsFileError.CreateFmt(sCellAccessError, [CellName, 'Float']);
end;

function TbiffCell.GetAsDateTime: TDateTime;
begin
  raise ExlsFileError.CreateFmt(sCellAccessError, [CellName, 'DateTime']);
end;

procedure TbiffCell.SetAsDateTime(Value: TDateTime);
begin
  raise ExlsFileError.CreateFmt(sCellAccessError, [CellName, 'DateTime']);
end;

function TbiffCell.GetAsVariant: variant;
begin
  Result := NULL;
end;

procedure TbiffCell.SetAsVariant(Value: variant);
begin
  raise ExlsFileError.CreateFmt(sCellAccessError, [CellName, 'Variant']);
end;

{ TbiffBoolErr }

function TbiffNumber.GetCellType: TbiffCellType;
begin
  if CellIsDateTime(Self)
    then Result := bctDateTime
    else Result := bctNumeric;
end;

function TbiffBoolErr.GetCellType: TbiffCellType;
begin
  Result := bctBoolean;
end;

function TbiffBoolErr.GetAsBoolean: boolean;
begin
  if Data[6] = 0
    then Result := false
    else Result := true
end;

procedure TbiffBoolErr.SetAsBoolean(Value: boolean);
begin
  Data[7] := 0;
  if Value
    then Data[6] := 1
    else Data[6] := 0;
end;

function TbiffBoolErr.GetAsVariant: variant;
begin
  if Data[7] = 0
    then  Result := GetAsBoolean
    else Result := ErrcodeToString(Data[6]);
end;

procedure TbiffBoolErr.SetAsVariant(Value: variant);
begin
  case VarType(Value) of
    varBoolean: SetAsBoolean(Value);
    varOleStr,
    varString: begin
      Data[7] := 1;
      Data[6] := StringToErrcode(Value);
    end;
    varNull: raise ExlsFileError.CreateFmt(sInvalidCellValue, ['']);

    else raise Exception.CreateFmt(sInvalidCellValue, [VarAsType(Value, varString)]);
  end;
end;

function TbiffBoolErr.GetAsString: WideString;
begin
  if GetAsBoolean
    then Result := 'true'
    else Result := 'false';
end;

procedure TbiffBoolErr.SetAsString(const Value: WideString);
begin
  if AnsiCompareText(Value, 'true') = 0
    then SetAsBoolean(true)
    else SetAsBoolean(false);
end;

{ TbiffNumber }

function TbiffNumber.GetAsFloat: double;
var
  D: double;
begin
  Move(Data[6], D, SizeOf(d));
  Result := D;
end;

procedure TbiffNumber.SetAsFloat(Value: double);
var
  D: double;
begin
  D := Value;
  Move(D, Data[6], SizeOf(D));
end;

function TbiffNumber.GetAsDateTime: TDateTime;
begin
  Result := CorrectDate(GetAsFloat);
end;

procedure TbiffNumber.SetAsDateTime(Value: TDateTime);
begin
  SetAsFloat(Value);
end;

function TbiffNumber.GetAsVariant: variant;
begin
  Result := GetAsFloat;
end;

procedure TbiffNumber.SetAsVariant(Value: variant);
begin
  SetAsFloat(Value);
end;

function TbiffNumber.GetAsString: WideString;
begin
  if GetCellType = bctNumeric then
    Result := FloatToStr(GetAsFloat)
  else begin
    if IsDateOnly then
      Result := DateToStr(GetAsDateTime)
    else if IsTimeOnly then
      Result := TimeToStr(GetAsDateTime)
    else Result := DateTimeToStr(GetAsDateTime);
  end;
end;

procedure TbiffNumber.SetAsString(const Value: WideString);
var
  D: double;
begin
  D := StrToFloat(Value);
  SetAsFloat(D);
end;

{ TbiffRK }

function TbiffRK.GetCellType: TbiffCellType;
begin
  if CellIsDateTime(Self)
    then Result := bctDateTime
    else Result := bctNumeric;
end;

function TbiffRK.GetAsFloat: double;
var
  RK, PD: ^longint;
  D: double;
begin
  RK := @(Data[6]);
  if RK^ and $2 = $2 then // integer
    if RK^ and (1 shl 31) = (1 shl 31) // negative
      then D := not (not (RK^) shr 2)
      else D := RK^ shr 2
  else begin
    PD := @D;
    PD^ := 0;
    Inc(PD);
    PD^ := RK^ and $FFFFFFFC;
  end;

  Result := D;
  if RK^ and $1 = $1 then Result := Result / 100;
end;

procedure TbiffRK.SetAsFloat(Value: double);
var
  RK: ^longint;
begin
  RK := @(Data[6]);
  if not EncodeRK(Value, RK^) then
    raise ExlsFileError.CreateFmt(sInvalidCellValue, [FloatToStr(Value)]);
end;

function TbiffRK.GetAsDateTime: TDateTime;
begin
  Result := CorrectDate(GetAsFloat);
end;

procedure TbiffRK.SetAsDateTime(Value: TDateTime);
begin
  SetAsFloat(Value);
end;

function TbiffRK.GetAsVariant: variant;
begin
  Result := GetAsFloat;
end;

procedure TbiffRK.SetAsVariant(Value: variant);
begin
  SetAsFloat(Value);
end;

function TbiffRK.GetAsString: WideString;
begin
  if GetCellType = bctNumeric then
    Result := FloatToStr(GetAsFloat)
  else begin
    if IsDateOnly then
      Result := DateToStr(GetAsDateTime)
    else if IsTimeOnly then
      Result := TimeToStr(GetAsDateTime)
    else Result := DateTimeToStr(GetAsDateTime);
  end;
end;

procedure TbiffRK.SetAsString(const Value: WideString);
var
  D: double;
  DT: TDateTime;
begin
  if GetCellType = bctNumeric then begin
    D := StrToFloat(Value);
    SetAsFloat(D);
  end
  else begin
    DT := StrToDateTime(Value);
    SetAsDateTime(DT);
  end;
end;

{ TxlsString }

constructor TxlsString.CreateR(IsWideStr: boolean; var ARecord: TbiffRecord;
  var Offset: integer);
var
  ByteLength: byte;
  RealOptionFlags: byte;
  DestPos: integer;
begin
  inherited Create;
  FIsWideStr := IsWideStr;
  if not FIsWideStr then begin
    ReadMem(ARecord, Offset, LenOfLen, @ByteLength);
    FStrLen := ByteLength
  end
  else
    ReadMem(ARecord, Offset, LenOfLen, @FStrLen);

  ReadMem(ARecord, Offset, SizeOf(FOptionFlags), @FOptionFlags);
  RealOptionFlags := FOptionFlags;

  if HasRichText
    then ReadMem(ARecord, Offset, SizeOf(FRTFNumber), @FRTFNumber)
    else FRTFNumber := 0;

  if HasFarEast
    then ReadMem(ARecord, Offset, SizeOf(FFarEastDataSize), @FFarEastDataSize)
    else FFarEastDataSize := 0;

  DestPos := 0;
  SetLength(FShortData, FStrLen);
  SetLength(FWideData, FStrLen);
  ReadStr(ARecord, Offset, FShortData, FWideData, FOptionFlags, RealOptionFlags,
    DestPos, FStrLen);
  if (Integer(HasWideChar) + 1) = 1
    then FWideData := EmptyStr
    else FShortData := QIGetEmptyStr;

  if FRTFNumber > 0 then begin
    GetMem(FRTFData, 4 * FRTFNumber);
    ReadMem(ARecord, Offset, 4 * FRTFNumber, FRTFData);
  end;

  if FFarEastDataSize > 0 then begin
    GetMem(FFarEastData, FFarEastDataSize);
    ReadMem(ARecord, Offset, FFarEastDataSize, FFarEastData)
  end;
end;

constructor TxlsString.CreateWS(IsWideStr: boolean; const Str: WideString);
begin
  inherited Create;
  FIsWideStr := IsWideStr;
  if (not FIsWideStr) and (Length(Str) > $FF) then
    raise ExlsFileError.Create(sInvalidStringRecord);

  FStrLen := Length(Str);

  FOptionFlags := 0;
  if IsWide(Str) then FOptionFlags := 1;

  FRTFNumber := 0;
  FFarEastDataSize := 0;

  if not GetHasWideChar
    then FShortData := WideStringToStringNoCodePage(Str)
    else FWideData := Str;
end;

destructor TxlsString.Destroy;
begin
  if Assigned(FFarEastData) then FreeMem(FFarEastData);
  if Assigned(FRTFData) then FreeMem(FRTFData);
  inherited;
end;

function TxlsString.GetLenOfLen: byte;
begin
  Result := Byte(FIsWideStr) + 1;
end;

function TxlsString.GetHasWideChar: boolean;
begin
  Result :=  not (FOptionFlags and $1 = 0);
end;

function TxlsString.GetFarEast: boolean;
begin
  Result := FOptionFlags and $4 = $4;
end;

function TxlsString.GetHasRichText: boolean;
begin
  Result := FOptionFlags and $8 = $8;
end;

function TxlsString.GetValue: WideString;
begin
  if not GetHasWideChar
    then Result := StringToWideStringNoCodePage(FShortData)
    else Result := WideData;
end;

function TxlsString.Compare(Str: TxlsString): integer;
begin
  if LenOfLen < Str.LenOfLen then begin
    Result := -1;
    Exit;
  end
  else if LenOfLen > Str.LenOfLen then begin
    Result := 1;
    Exit;
  end;

  if FOptionFlags < Str.OptionFlags then begin
    Result := -1;
    Exit;
  end
  else if FOptionFlags > Str.OptionFlags then begin
    Result := 1;
    Exit;
  end;

  if not GetHasWideChar
    then Result := CompareStr(String(FShortData), String(Str.ShortData))
    else Result:= CompareWideStr(FWideData, Str.WideData);
end;

{ TxlsSSTEntry }

constructor TxlsSSTEntry.CreateXS(Str: TxlsString);
begin
  inherited Create;
  FValue := Str;
  FRefCount := 0;
end;

constructor TxlsSSTEntry.CreateWS(Str: WideString);
begin
  CreateXS( TxlsString.CreateWS(true, Str) );
end;

destructor TxlsSSTEntry.Destroy;
begin
  if Assigned(FValue) then FValue.Free;
  //if Assigned(FOnDestroy) then FOnDestroy(Self);  
  inherited Destroy;
end;

procedure TxlsSSTEntry.IncRef;
begin
//  Inc(FRefCount);
end;

procedure TxlsSSTEntry.DecRef;
begin
//  Dec(FRefCount);
end;

{ TbiffLabelSST }

function TbiffLabelSST.GetCellType: TbiffCellType;
begin
  Result := bctString;
end;

var cc: integer = 0;

constructor TbiffLabelSST.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
var
  l: integer;
begin
  Inc(cc);
  inherited Create(Section, ID, DataSize,Data);
  l := GetInteger(Data, 6);
  if l >= SSTList.Count then
    raise ExlsFileError.Create(sExcelInvalid);

  FSSTEntry := SSTList[l];
  FSSTEntry.IncRef;
  FSSTEntry.OnDestroy := DestroySSTEntry;
end;

destructor TbiffLabelSST.Destroy;
begin
  Dec(cc);
  if Assigned(FSSTEntry) then
  begin
    FSSTEntry.DecRef;
//    if FSSTEntry.RefCount = 0 then
//      FSSTEntry.Free;
  end;
  inherited;
end;

function TbiffLabelSST.GetAsString: WideString;
begin
  Result := FSSTEntry.Value.Value;
end;

procedure TbiffLabelSST.SetAsString(const Value: WideString);
var
  OldSSTEntry: TxlsSSTEntry;
begin
  OldSSTEntry := FSSTEntry;
  FSSTEntry := SSTList[SSTList.AddString(Value)];
  if Assigned(OldSSTEntry) then OldSSTEntry.DecRef;
end;

function TbiffLabelSST.GetAsVariant: variant;
begin
  Result := GetAsString;
end;

procedure TbiffLabelSST.SetAsVariant(Value: variant);
begin
  SetAsString(Value);
end;

procedure TbiffLabelSST.DestroySSTEntry(Sender: TObject);
begin
  //FSSTEntry := nil;
end;

{ ETokenException }

constructor ETokenException.Create(Token: integer; const Msg: string);
begin
  FToken := Token;
  inherited CreateFmt({sBadToken}Msg, [FToken]);
end;

{ TbiffFormula }

constructor TbiffFormula.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
var
  D: double;
begin
  inherited;
  FValue := NULL;
  if GetWord(Data, 12) <> $FFFF then begin // numeric
    Move(Data[6], D, SizeOf(d));
    FValue := D;
  end else begin
    case Data[6] of
      0: FValue := QIGetEmptyStr; // string
      1: FValue := Data[8] = 1; // boolean
      //2 is error. we can't codify this on a variant.
    end;
  end;

  FillChar(Data^[6], 8, 0);
  Data^[6] := 2; //error value
  SetWord(Data, 12, $FFFF);
  FillChar(Data^[16], 4, 0);

  Data^[14] := Data^[14] or 2;
end;

procedure TbiffFormula.MixShared(SharedData: PByteArray;
  SharedDataSize: integer);
var
  NewDataSize: integer;
begin
  // Note: This method changes the size of the record without notifying
  NewDataSize := DataSize - 5 + SharedDataSize - 8;
  ReallocMem(FData, NewDataSize);
  DataSize := NewDataSize;
  Move(SharedData[8], Data[20], SharedDataSize - 8);

  try
    ArrangeSharedFormulas;
  except
    on E: ETokenException do
      raise Exception.CreateFmt(sBadFormula, [Row + 1, Col + 1, E.Token]);
    else raise;
  end;
end;

function TbiffFormula.GetIsFormula: boolean;
begin
  Result := true;
end;

function TbiffFormula.GetAsVariant: variant;
begin
  Result := FValue;
end;

procedure TbiffFormula.SetAsVariant(Value: variant);
begin
  FValue := Value;
end;

function TbiffFormula.GetAsString: WideString;
begin
  Result := VarToStr(GetAsVariant);
end;

procedure TbiffFormula.SetAsString(const Value: WideString);
begin
  SetAsVariant(Value);
end;

function TbiffFormula.GetIsExp: boolean;
begin
  Result := (DataSize = 27) and (GetWord(Data, 20) = 5) and (Data[22] = 1);
end;

function TbiffFormula.GetKey: cardinal;
begin
  Result := 0;
  if GetIsExp then
    Result := GetWord(Data, 23) or (GetWord(Data, 25) shl 16);
end;

procedure TbiffFormula.ArrangeSharedFormulas;
var
  S, F: integer;
  Token: byte;

  procedure ArrangeOperand;
  var
    AbsRef: boolean;
  begin
    if Token in [ptgRefN, ptgRefNV, ptgRefNA,
                 ptgAreaN, ptgAreaNV, ptgAreaNA] then
    begin
      Dec(Data[S], 8);
      Token := Data[S];
    end;

    Inc(S);

    if Token in [ptgRef3d, ptgRef3dV, ptgRef3dA,
                 ptgArea3d, ptgArea3dV, ptgArea3dA,
                 ptgRefErr3d, ptgRefErr3dV, ptgRefErr3dA,
                 ptgAreaErr3d, ptgAreaErr3dV, ptgAreaErr3dA] then begin
      Inc(S, 2);
    end;

    case Token of
      ptgArray,
      ptgArrayV, ptgArrayA: Inc(S, 7);

      ptgName,
      ptgNameV, ptgNameA: Inc(S, 4);

      ptgNameX,
      ptgNameXV, ptgNameXA: Inc(S, 6);

      ptgRef,
      ptgRefV, ptgRefA,
      ptgRefErr,
      ptgRefErrV, ptgRefErrA,
      ptgRef3d,
      ptgRef3dV, ptgRef3dA,
      ptgRefErr3d,
      ptgRefErr3dV, ptgRefErr3dA: begin
        // row defined absolutely
        AbsRef := (GetWord(Data, S + 2) and $8000) <> $8000;
        if not AbsRef then Data[S] := Data[S] + Row;
        // col defined absolutely
        AbsRef := (GetWord(Data, S + 2) and $4000) <> $4000;
        if not AbsRef then Data[S + 2] := Data[S + 2] + Col;
        Inc(S, 4);
      end;

      ptgRefN,
      ptgRefNV, ptgRefNA: Inc(S, 4);

      ptgArea,
      ptgAreaV, ptgAreaA,
      ptgAreaErr,
      ptgAreaErrV, ptgAreaErrA,
      ptgArea3d,
      ptgArea3dV, ptgArea3dA,
      ptgAreaErr3d,
      ptgAreaErr3dV, ptgAreaErr3dA: begin
        AbsRef := GetWord(Data, S + 4) and $8000 <> $8000;
        if not AbsRef then Data[S] := Data[S] + Row;
        AbsRef:= GetWord(Data, S + 4) and $4000 <> $4000;
        if not AbsRef then Data[S + 4] := Data[S + 4] + Col;

        AbsRef:= GetWord(Data, S + 6) and $8000 <> $8000;
        if not AbsRef then Data[S + 2] := Data[S + 2] + Row;;
        AbsRef:= GetWord(Data, S + 6) and $4000 <> $4000;
        if not AbsRef then Data[S + 6] := Data[S + 6] + Col;

        Inc(S, 8);
      end;

      ptgAreaN,
      ptgAreaNV, ptgAreaNA: Inc(S, 8);

      else raise ETokenException.Create(Token, sBadToken);
    end;
  end;

  procedure ArrangeTable;
  begin
  end;

begin
  S := 22;
  F := S + GetWord(Data, 20);
  while S < F do begin
    Token := Data[S];
    case Token of
      ptgUplus..ptgParen,
      ptgAdd..ptgRange,
      ptgMissArg: Inc(S);

      ptgStr: Inc(S, 1 + GetStrLen(false, Data, S + 1, false, 0));

      ptgErr, ptgBool: Inc(S, 1 + 1);

      ptgInt, ptgFunc,
      ptgFuncV, ptgFuncA: Inc(S, 1 + 2);

      ptgFuncVar,
      ptgFuncVarV, ptgFuncVarA: Inc(S, 1 + 3);

      ptgNum: Inc(S, 1 + 8);

      ptgAttr: begin
        if (Data[S + 1] and $04) = $04 then
          Inc(S, (GetWord(Data, S + 2) + 1) * 2);
        Inc(S, 1 + 3);
      end;

      ptgArray,
      ptgName..ptgArea,
      ptgRefErr..ptgAreaN,
      ptgNameX..ptgAreaErr3d,
      ptgArrayV,
      ptgNameV..ptgAreaV,
      ptgRefErrV..ptgAreaNV,
      ptgNameXV..ptgAreaErr3dV,
      ptgNameA..ptgAreaA,
      ptgRefErrA..ptgAreaNA,
      ptgNameXA..ptgAreaErr3dA: ArrangeOperand;

      ptgTbl: ArrangeTable;
      else raise ETokenException.Create(Token, sBadToken);
    end;
  end;
end;

{ TbiffShrFmla }

function TbiffShrFmla.GetFirstRow: word;
begin
  Result := GetWord(Data, 0);
end;

function TbiffShrFmla.GetLastRow: word;
begin
  Result := GetWord(Data, 2);
end;

function TbiffShrFmla.GetFirstCol: word;
begin
  Result := Data[4];
end;

function TbiffShrFmla.GetLastCol: word;
begin
  Result := Data[5];
end;

function TbiffShrFmla.GetKey: integer{longword};
begin
  Result := GetWord(Data, 0) or Data[4] shl 16;
end;

{ TbiffName }

function TbiffName.GetName: WideString;
var
  Str: AnsiString;
begin
  if (GetOptionFlags and $01) = 1 then begin
    SetLength(Result, GetNameLength);
    Move(Data[15], Result[1], GetNameLength * 2);
  end else begin
    SetLength(Str, GetNameLength);
    Move(Data[15], Str[1], GetNameLength);
    Result := WideString(Str);
  end;
end;

function TbiffName.GetNameLength: byte;
begin
  Result := Data[3];
end;

function TbiffName.GetNameSize: integer;
begin
  Result := GetStrLen(false , Data, 14, true, NameLength);
end;

function TbiffName.GetOptionFlags: byte;
begin
  Result := Data[14];
end;

function TbiffName.GetRow1: integer;
begin
  if Data[14 + NameSize] in tkArea3d
    then Result := GetWord(Data, 15 + 2 + NameSize)
    else Result := -1;
end;

function TbiffName.GetRow2: integer;
begin
  if Data[14 + NameSize] in tkArea3d
    then Result := GetWord(Data, 15 + 4 + NameSize)
    else Result := -1;
end;

function TbiffName.GetCol1: integer;
begin
  if Data[14 + NameSize] in tkArea3d
    then Result := GetWord(Data, 15 + 6 + NameSize)
    else Result := -1;
end;

function TbiffName.GetCol2: integer;
begin
  if Data[14 + NameSize] in tkArea3d
    then Result := GetWord(Data, 15 + 8 + NameSize)
    else Result := -1;
end;

{ TbiffMultiple }

constructor TbiffMultiple.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
begin
  inherited;
  FCol := 0;
end;

{ TbiffMulBlank }

function TbiffMulBlank.GetEOF: boolean;
begin
  Result := 4 + (FCol + 1) * SizeOf(Word) >= DataSize;
end;

function TbiffMulBlank.GetCell: TbiffCell;
var
  NewData: PByteArray;
  NewDataSize: integer;
begin
  NewDataSize := 6;
  GetMem(NewData, NewDataSize);
  try
    SetWord(NewData, 0, GetWord(Data, 0));
    SetWord(NewData, 2, GetWord(Data, 2) + FCol);
    SetWord(NewData, 4, GetWord(Data, 4 + FCol * SizeOf(Word)));

    Result := TbiffBlank.Create(Section, BIFF_BLANK, NewDataSize, NewData);
    Inc(FCol);
  except
    FreeMem(NewData);
    raise;
  end;
end;

{ TbiffMulRK }

type
  P_RK = ^T_RK;
  T_RK = packed record
    XF: word;
    RK: longint;
  end;

function TbiffMulRK.GetEOF: boolean;
begin
  Result := 4 + (FCol + 1) * SizeOf(T_RK) >= DataSize;
end;

function TbiffMulRK.GetCell: TbiffCell;
var
  NewData: PByteArray;
  NewDataSize: integer;
  RK1, RK2: P_RK;
begin
  NewDataSize := 10;
  GetMem(NewData, NewDataSize);
  try
    SetWord(NewData, 0, GetWord(Data, 0));
    SetWord(NewData, 2, GetWord(Data, 2) + FCol);
    RK1 := P_RK(@(Data[4 + FCol * SizeOf(T_RK)]));
    RK2 := P_RK(@(NewData[4]));
    RK2^ := RK1^;

    Result := TbiffRK.Create(Section, BIFF_RK, NewDataSize, NewData);
    Inc(FCol);
  except
    FreeMem(NewData);
    raise;
  end;
end;

{ TbiffXF }

function TbiffXF.GetFormatIndex: word;
begin
  Result := PBIFF_XF(Data).FormatIndex;
end;

procedure TbiffXF.SetFormatIndex(Value: word);
begin
  PBIFF_XF(Data).FormatIndex := Value;
end;

{ TbiffFormat }

constructor TbiffFormat.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
var
  MySelf: TbiffRecord;
  TempPos, StrLen: integer;
  Str: AnsiString;
  WStr: WideString;
  OptionFlags, RealOptionFlags: byte;
  DestPos: integer;
begin
  inherited;
  FID := GetWord(Data, 0);

  TempPos := 5;
  MySelf := Self;
  DestPos := 0;
  OptionFlags := Data[4];
  RealOptionFlags := OptionFlags;
  StrLen := GetWord(Data, 2);
  SetLength(Str, StrLen);
  SetLength(WStr, StrLen);
  ReadStr(MySelf, TempPos, Str, WStr, OptionFlags, RealOptionFlags, DestPos,
    StrLen);
  if (OptionFlags and $1) = 0
    then FValue := StringToWideStringNoCodePage(Str)
    else FValue := WStr;
end;

{ TbiffSST }

constructor TbiffSST.Create(Section: TxlsSection; ID, DataSize: word;
  Data: PByteArray);
begin
  inherited;
  FCount := GetInteger(Data, 4);
end;

{ TxlsList }

function TxlsList.GetItems(Index: integer): TObject;
begin
  Result := TObject(inherited Items[Index]);
end;

procedure TxlsList.SetItems(Index: integer; Value: TObject);
begin
  inherited Items[Index] := Value;
end;

function TxlsList.Add(Item: TObject): integer;
begin
  Result := inherited Add(Item);
end;

constructor TxlsList.Create(Workbook: TxlsWorkbook);
begin
  inherited Create;
  FWorkbook := Workbook;
end;

procedure TxlsList.Clear;
var
  I: Integer;
begin
  for I := Count - 1 downto 0 do
    Delete(I);
  inherited Clear;
end;

procedure TxlsList.Delete(Index: integer);
var
  O: TObject;
begin
  if Assigned(Items[Index]) then
  begin
    if Items[Index] is TxlsList then
      TxlsList(Items[Index]).Clear;
    O := Items[Index];
    ObjFreeAndNil( O );
  end;
end;

{$IFDEF VCL5}
function TxlsList.Extract(Item: TObject): TObject;
begin
  Result := TObject(inherited Extract(Item));
end;
{$ENDIF}

function TxlsList.First: TObject;
begin
  Result := TObject(inherited First);
end;

function TxlsList.IndexOf(Item: TObject): integer;
begin
  Result := inherited IndexOf(Item);
end;

procedure TxlsList.Insert(Index: integer; Item: TObject);
begin
  inherited Insert(Index, Item);
end;

function TxlsList.Last: TObject;
begin
  Result := TObject(inherited Last);
end;

function TxlsList.Remove(Item: TObject): integer;
begin
  Result := inherited Remove(Item);
end;

{ TbiffRecordList }

function TbiffRecordList.Add(Item: TbiffRecord): integer;
begin
  Result := inherited Add(Item);
end;

procedure TbiffRecordList.Insert(Index: integer; Item: TbiffRecord);
begin
  inherited Insert(Index, Item);
end;

procedure TbiffRecordList.CorrectSize(Delta: integer);
begin
  Inc(FTotalSize, Delta);
end;

procedure TbiffRecordList.RecalculateTotalSize;
var
  i: integer;
begin
  FTotalSize := 0;
  for i := 0 to Count - 1 do
    Inc(FTotalSize, Items[i].FDataSize);
end;

function TbiffRecordList.GetItems(Index: integer): TbiffRecord;
begin
  Result := TbiffRecord(inherited GetItems(Index));
end;

procedure TbiffRecordList.SetItems(Index: integer; Value: TbiffRecord);
begin
  inherited SetItems(Index, Value);
end;

{ TxlsRowList }

constructor TxlsRowList.Create(Workbook: TxlsWorkbook);
begin
  FSorted := false;
end;

function TxlsRowList.GetItems(Index: integer): TxlsRow;
begin
  Result := TxlsRow(inherited Items[Index]);
end;

procedure TxlsRowList.SetItems(Index: integer; Value: TxlsRow);
begin
  inherited Items[Index] := Value;
end;

function TxlsRowList.Add(Row: TxlsRow): integer;
begin
  Result := inherited Add(Row);
  FSorted := false;
end;

procedure TxlsRowList.Insert(Index: integer; Row: TxlsRow);
begin
  inherited Insert(Index, Row);
end;

function TxlsRowList.Find(Row: integer; var Index: integer): boolean;
var
 L, H, I, C: Integer;
begin
  if not FSorted then Sort;
  Result := false;
  L := 0;
  H := Count - 1;
  while L <= H do begin
    I := (L + H) shr 1;
    if Items[i].RowNumber < Row then
      C := -1
    else if Items[i].RowNumber > Row then
      C := 1
    else C := 0;
    if C < 0 then L := I + 1
    else begin
      H := I - 1;
      if C = 0 then begin
        Result := true;
        L := I;
      end;
    end;
  end;
  Index := L;
end;

function CompareRowNumber(Item1, Item2: Pointer): integer;
begin
  if TxlsRow(Item1).RowNumber < TxlsRow(Item2).RowNumber then
    Result := -1
  else if TxlsRow(Item1).RowNumber > TxlsRow(Item2).RowNumber then
    Result := 1
  else Result := 0;
end;

procedure TxlsRowList.Sort;
begin
  inherited Sort(CompareRowNumber);
  FSorted := true;
end;

{ TxlsColList }

constructor TxlsColList.Create(Workbook: TxlsWorkbook);
begin
  inherited;
  FSorted := false;
end;

function TxlsColList.GetItems(Index: integer): TxlsCol;
begin
  Result := TxlsCol(inherited Items[Index]);
end;

procedure TxlsColList.SetItems(Index: integer; Value: TxlsCol);
begin
  inherited Items[Index] := Value;
end;

function TxlsColList.Add(Col: TxlsCol): integer;
begin
  Result := inherited Add(Col);
  FSorted := false;
end;

procedure TxlsColList.Insert(Index: integer; Col: TxlsCol);
begin
  inherited Insert(Index, Col);
end;

function TxlsColList.Find(Col: integer; var Index: integer): boolean;
var
 L, H, I, C: Integer;
begin
  if not FSorted then Sort;
  Result := false;
  L := 0;
  H := Count - 1;
  while L <= H do begin
    I := (L + H) shr 1;
    if Items[i].ColNumber < Col then
      C := -1
    else if Items[i].ColNumber > Col then
      C := 1
    else C := 0;
    if C < 0 then L := I + 1
    else begin
      H := I - 1;
      if C = 0 then begin
        Result := true;
        L := I;
      end;
    end;
  end;
  Index := L;
end;

function CompareColNumber(Item1, Item2: Pointer): integer;
begin
  if TxlsCol(Item1).ColNumber < TxlsCol(Item2).ColNumber then
    Result := -1
  else if TxlsCol(Item1).ColNumber > TxlsCol(Item2).ColNumber then
    Result := 1
  else Result := 0;
end;

procedure TxlsColList.Sort;
begin
  inherited Sort(CompareColNumber);
  FSorted := true;
end;

{ TbiffColRowList }

function TbiffColRowList.GetItems(Index: integer): TbiffColRow;
begin
  Result := TbiffColRow(inherited Items[Index]);
end;

procedure TbiffColRowList.SetItems(Index: integer; Value: TbiffColRow);
begin
  inherited Items[Index] := Value;
end;

function TbiffColRowList.Add(Item: TbiffColRow): integer;
begin
  Result := inherited Add(Item);
end;

procedure TbiffColRowList.Insert(Index: integer; Item: TbiffColRow);
begin
  inherited Insert(Index, Item);
end;

{ TbiffCellList }

function TbiffCellList.GetItems(Index: integer): TbiffCell;
begin
  Result := TbiffCell(inherited Items[Index]);
end;

procedure TbiffCellList.SetItems(Index: integer; Value: TbiffCell);
begin
  inherited Items[Index] := Value;
end;

procedure TbiffCellList.SetMinAndMaxCells(Item: TbiffCell);
begin
// do nothing
end;

procedure TbiffCellList.SetColRowNumber(Item: TbiffCell);
begin
// do nothing
end;

function TbiffCellList.Add(Item: TbiffCell): integer;
begin
  Result := inherited Add(Item);
  FSorted := false;
  Item.OnDestroy := OnDestroyItem;

  SetMinAndMaxCells(Item);
  if Count = 1 then SetColRowNumber(Item);
end;

procedure TbiffCellList.Clear;
begin
  inherited Clear;
end;

procedure TbiffCellList.Insert(Index: integer; Item: TbiffCell);
begin
  inherited Insert(Index, Item);
  Item.OnDestroy := OnDestroyItem;
end;

procedure TbiffCellList.OnDestroyItem(Sender: TObject);
var
  i: integer;
begin
  for i := Count - 1 downto 0 do
    if Items[i] = Sender then begin
      Remove(Items[i]);
      Break;
    end;
end;

{ TxlsRow }

constructor TxlsRow.Create(Workbook: TxlsWorkbook);
begin
  inherited;
  FRowNumber := -1;
  FMinCol := -1;
  FMaxCol := -1;
end;

procedure TxlsRow.SetMinAndMaxCells(Item: TbiffCell);
begin
  if (FMinCol = -1) or (Item.Col < FMinCol) then
    FMinCol := Item.Col;
  if (FMaxCol = -1) or (Item.Col > FMaxCol) then
    FMaxCol := Item.Col;
end;

procedure TxlsRow.SetColRowNumber(Item: TbiffCell);
begin
  FRowNumber := Item.Row;
end;

function TxlsRow.Find(Col: integer; var Index: integer): boolean;
var
 L, H, I, C: Integer;
begin
  if not Sorted then Sort;
  Result := false;
  L := 0;
  H := Count - 1;
  while L <= H do begin
    I := (L + H) shr 1;
    if Items[i].Col < Col then
      C := -1
    else if Items[i].Col > Col then
      C := 1
    else C := 0;
    if C < 0 then L := I + 1
    else begin
      H := I - 1;
      if C = 0 then begin
        Result := true;
        L := I;
      end;
    end;
  end;
  Index := L;
end;

function CompareCellCols(Item1, Item2: Pointer): integer;
begin
  if TbiffCell(Item1).Col < TbiffCell(Item2).Col then
    Result := -1
  else if TbiffCell(Item1).Col > TbiffCell(Item2).Col then
    Result := 1
  else Result := 0;
end;

procedure TxlsRow.Sort;
begin
  inherited Sort(CompareCellCols);
  Sorted := true;
end;

{ TxlsCol }

constructor TxlsCol.Create(Workbook: TxlsWorkbook);
begin
  inherited;
  FColNumber := -1;
  FMinRow := -1;
  FMaxRow := -1;
end;

procedure TxlsCol.SetMinAndMaxCells(Item: TbiffCell);
begin
  if (FMinRow = -1) or (Item.Row < FMinRow) then
    FMinRow := Item.Row;
  if (FMaxRow = -1) or (Item.Row > FMaxRow) then
    FMaxRow := Item.Row;
end;

procedure TxlsCol.SetColRowNumber(Item: TbiffCell);
begin
  FColNumber := Item.Col;
end;

function TxlsCol.Find(Row: integer; var Index: integer): boolean;
var
 L, H, I, C: Integer;
begin
  if not Sorted then Sort;
  Result := false;
  L := 0;
  H := Count - 1;
  while L <= H do begin
    I := (L + H) shr 1;
    if Items[i].Row < Row then
      C := -1
    else if Items[i].Row > Row then
      C := 1
    else C := 0;
    if C < 0 then L := I + 1
    else begin
      H := I - 1;
      if C = 0 then begin
        Result := true;
        L := I;
      end;
    end;
  end;
  Index := L;
end;

function CompareCellRows(Item1, Item2: Pointer): integer;
begin
  if TbiffCell(Item1).Row < TbiffCell(Item2).Row then
    Result := -1
  else if TbiffCell(Item1).Row > TbiffCell(Item2).Row then
    Result := 1
  else Result := 0;
end;

procedure TxlsCol.Sort;
begin
  inherited Sort(CompareCellRows);
  Sorted := true;
end;

{ TbiffShrFmlaList }

function TbiffShrFmlaList.GetItems(Index: integer): TbiffShrFmla;
begin
  Result := TbiffShrFmla(inherited Items[Index]);
end;

procedure TbiffShrFmlaList.SetItems(Index: integer; Value: TbiffShrFmla);
begin
  inherited Items[Index] := Value;
end;

function TbiffShrFmlaList.Add(Item: TbiffShrFmla): integer;
begin
  Result := inherited Add(Item);
  FSorted := false;
end;

procedure TbiffShrFmlaList.Insert(Index: integer; Item: TbiffShrFmla);
begin
  inherited Insert(Index, Item);
end;

function TbiffShrFmlaList.Find(Key: integer; var Index: integer): boolean;
Var
 L, H, I, C: Integer;
begin
  if not FSorted then Sort;
  Result := false;
  L := 0;
  H := Count - 1;
  while L <= H do begin
    I := (L + H) shr 1;
    if Items[i].Key < Key then
      C := -1
    else if Items[i].Key > Key then
      C := 1
    else C := 0;
    if C < 0 then L := I + 1
    else begin
      H := I - 1;
      if C = 0 then begin
        Result := true;
        L := I;
      end;
    end;
  end;
  Index := L;
end;

function CompareFormulaKey(Item1, Item2: Pointer): integer;
begin
  if TbiffShrFmla(Item1).Key < TbiffShrFmla(Item2).Key then
    Result := -1
  else if TBiffShrFmla(Item1).Key > TBiffShrFmla(Item2).Key then
    Result := 1
  else Result := 0;
end;

procedure TbiffShrFmlaList.Sort;
begin
  inherited Sort(CompareFormulaKey);
  FSorted := true;
end;

{ TbiffNameList }

function TbiffNameList.GetItems(Index: integer): TbiffName;
begin
  Result := TbiffName(inherited Items[Index]);
end;

procedure TbiffNameList.SetItems(Index: integer; Value: TbiffName);
begin
  inherited Items[Index] := Value;
end;

function TbiffNameList.Add(Item: TbiffName): integer;
begin
  Result := inherited Add(Item);
end;

procedure TbiffNameList.Insert(Index: integer; Item: TbiffName);
begin
  inherited Insert(Index, Item);
end;

{ TbiffBoundSheetList }

function TbiffBoundSheetList.GetItems(Index: integer): TbiffBoundSheet;
begin
  Result := TbiffBoundSheet(inherited Items[Index]);
end;

procedure TbiffBoundSheetList.SetItems(Index: integer; Value: TbiffBoundSheet);
begin
  inherited Items[Index] := Value;
end;

function TbiffBoundSheetList.Add(Item: TbiffBoundSheet): integer;
begin
  Result := inherited Add(Item);
end;

procedure TbiffBoundSheetList.Insert(Index: integer; Item: TbiffBoundSheet);
begin
  inherited Insert(Index, Item);
end;

function TbiffBoundSheetList.GetName(Index: integer): WideString;
begin
  Result := Items[Index].Name;
end;

{ TbiffXFList }

function TbiffXFList.GetItems(Index: integer): TbiffXF;
begin
  Result := TbiffXF(inherited Items[Index]);
end;

procedure TbiffXFList.SetItems(Index: integer; Value: TbiffXF);
begin
  inherited Items[Index] := Value;
end;

function TbiffXFList.Add(Item: TbiffXF): integer;
begin
  Result := inherited Add(Item);
end;

procedure TbiffXFList.Insert(Index: integer; Item: TbiffXF);
begin
  inherited Insert(Index, Item);
end;

{ TbiffFormatList }

function TbiffFormatList.GetItems(Index: integer): TbiffFormat;
begin
  Result := TbiffFormat(inherited Items[Index]);
end;

procedure TbiffFormatList.SetItems(Index: integer; Value: TbiffFormat);
begin
  inherited Items[Index] := Value;
end;

function TbiffFormatList.GetFormat(ID: integer): WideString;
var
  Index: integer;
begin
  if Find(ID, Index) then
    Result := Items[Index].Value
  else if (ID >= Low(InternalNumberFormats)) and (ID <= High(InternalNumberFormats)) then
    Result := InternalNumberFormats[ID]
  else Result := EmptyStr;
end;

function TbiffFormatList.Add(Item: TbiffFormat): integer;
begin
  Result := inherited Add(Item);
  FSorted := false;
end;

procedure TbiffFormatList.Insert(Index: integer; Item: TbiffFormat);
begin
  inherited Insert(Index, Item);
end;

function TbiffFormatList.Find(ID: integer; var Index: integer): boolean;
var
 L, H, I, C: Integer;
begin
  if not FSorted then Sort;
  Result := false;
  L := 0;
  H := Count - 1;
  while L <= H do begin
    I := (L + H) shr 1;
    if Items[i].ID < ID then
      C := -1
    else if Items[i].ID > ID then
      C := 1
    else C := 0;
    if C < 0 then L := I + 1 else begin
      H := I - 1;
      if C = 0 then begin
        Result := true;
        L := I;
      end;
    end;
  end;
  Index := L;
end;

function CompareFormat(Item1, Item2: Pointer): integer;
begin
  if TbiffFormat(Item1).ID < TbiffFormat(Item2).ID then
    Result := -1
  else if TbiffFormat(Item1).ID > TbiffFormat(Item2).ID then
    Result := 1
  else Result := 0;
end;

procedure TbiffFormatList.Sort;
begin
  inherited Sort(CompareFormat);
  FSorted := true;
end;

{ TbiffSSTList }

function TbiffSSTList.GetItems(Index: integer): TxlsSSTEntry;
begin
  Result := TxlsSSTEntry(inherited Items[Index]);
end;

procedure TbiffSSTList.SetItems(Index: integer; Value: TxlsSSTEntry);
begin
  inherited Items[Index] := Value;
end;

function TbiffSSTList.Add(Item: TxlsSSTEntry): integer;
begin
  Result := inherited Add(Item);
end;

procedure TbiffSSTList.Insert(Index: integer; Item: TxlsSSTEntry);
begin
  inherited Insert(Index, Item);
end;

function TbiffSSTList.Find(Str: TxlsString; var Index: integer): boolean;
var
  L, H, I, C: integer;
begin
  Result := false;
  L := 0;
  H := Count - 1;
  while L <= H do begin
    I := (L + H) shr 1;
    C := Items[I].Value.Compare(Str);
    if C < 0 then
      L := I + 1
    else begin
      H := I - 1;
      if C = 0 then begin
        Result := true;
        L := I;
      end;
    end;
  end;
  Index := L;
end;

function CompareSSTEntries(Item1, Item2: Pointer): integer;
begin
  Result := TxlsSSTEntry(Item1).Value.Compare(TxlsSSTEntry(Item2).Value);
end;

procedure TbiffSSTList.Sort;
begin
  inherited Sort(CompareSSTEntries)
end;

procedure TbiffSSTList.Load(SST: TbiffSST);
var
  i, Offset: integer;
  Str: TxlsString;
  Tmp: TbiffRecord;
begin
  Offset := 8;
  Tmp := SST;
  for i := 0 to SST.Count - 1 do
  begin
    Str := TxlsString.CreateR(true, Tmp, Offset);
    Add( TxlsSSTEntry.CreateXS(Str) );
  end;
end;

function TbiffSSTList.AddString(const Str: WideString): integer;
var
  S: TxlsString;
  Entr: TxlsSSTEntry;
begin
  S := TxlsString.CreateWS(true, Str);
  try
    if Find(S, Result) then
      Items[Result].IncRef
    else begin
      Entr := TxlsSSTEntry.CreateWS(Str);
      Entr.IncRef;
      Insert(Result, Entr);
    end;
  finally
    S.Free;
  end;
end;

{ TxlsSection }

constructor TxlsSection.Create(Workbook: TxlsWorkbook);
begin
  inherited Create;
  FBOF := nil;
  FEOF := nil;
  FWorkbook := Workbook;
end;

destructor TxlsSection.Destroy;
begin
  if Assigned(FOnDestroy) then FOnDestroy(Self);
  Clear;
  inherited;
end;

procedure TxlsSection.Clear;
begin
  if Assigned(FBOF) then ObjFreeAndNil(FBOF);
  if Assigned(FEOF) then ObjFreeAndNil(FEOF);
end;

{ TxlsGlobals }

constructor TxlsGlobals.Create(Workbook: TxlsWorkbook);
begin
  inherited;
  FBoundSheetList := TbiffBoundSheetList.Create(Workbook);
  FNameList := TbiffNameList.Create(Workbook);
  FSSTList := TbiffSSTList.Create(Workbook);
  FXFList := TbiffXFList.Create(Workbook);
  FFormatList := TbiffFormatList.Create(Workbook);
end;

destructor TxlsGlobals.Destroy;
begin
  Clear;
  FBoundSheetList.Free;
  FNameList.Free;
  FSSTList.Free;
  FXFList.Free;
  FFormatList.Free;
  inherited;
end;

procedure TxlsGlobals.Clear;
begin
  inherited Clear;
  if Assigned(FBoundSheetList) then FBoundSheetList.Clear;
  if Assigned(FNameList) then FNameList.Clear;
  if Assigned(FSSTList) then FSSTList.Clear;
  if Assigned(FXFList) then FXFList.Clear;
  if Assigned(FFormatList) then FFormatList.Clear;
end;

procedure TxlsGlobals.Load(Stream: TStream; BOF: TbiffBOF);
var
  Header: TBIFF_Header;
  R: TbiffRecord;
begin
  Clear;
  repeat
    if (Stream.Read(Header, SizeOf(Header)) <> SizeOf(Header)) then
      raise ExlsFileError.Create(sExcelInvalid);

    R := LoadRecord(Self, Stream, Header);
    try
      case R.ID of
        BIFF_BOF       : raise ExlsFileError.Create(sExcelInvalid);
        BIFF_BOUNDSHEET:
          if R is TbiffBoundSheet then
            FBoundSheetList.Add(R as TbiffBoundSheet)
          else ObjFreeAndNil(R);
        BIFF_NAME      :
          if R is TbiffName then FNameList.Add(R as TbiffName)
          else ObjFreeAndNil(R);
        BIFF_XF        :
          if R is TbiffXF then FXFList.Add(R as TbiffXF)
          else ObjFreeAndNil(R);
        BIFF_FORMAT    :
          if R is TbiffFormat then FFormatList.Add(R as TbiffFormat)
          else ObjFreeAndNil(R);
        BIFF_EOF       :
          if R is TbiffEOF then EOF := R as TbiffEOF
          else ObjFreeAndNil(R);
        BIFF_SST       :
          if R is TbiffSST then begin
            FSSTList.Load(R as TbiffSST);
            ObjFreeAndNil(R);
          end
          else ObjFreeAndNil(R);
      else
        ObjFreeAndNil(R);
      end;
    except
      if Assigned(R) then ObjFreeAndNil(R);
      raise;
    end;

  until Header.ID = BIFF_EOF;

  FixFormats;

  if Assigned(Self.FBOF) then ObjFreeAndNil(Self.FBOF);
  Self.FBOF := BOF;
end;

procedure TxlsGlobals.FixFormats;
var
  i, j: integer;
begin
  for i := 0 to FXFList.Count - 1 do
    for j := 0 to FFormatList.Count - 1 do
      if FXFList[i].FormatIndex = FFormatList.Items[j].ID then
        FXFList[i].FormatIndex := j + High(InternalNumberFormats);
end;

{ TxlsSheet }

constructor TxlsSheet.Create(Workbook: TxlsWorkbook);
begin
  inherited;
  FIndex := -1;
end;

function TxlsSheet.GetName: WideString;
begin
  Result := Workbook.Globals.BoundSheetList[FIndex].Name;
end;

{ TxlsSheetList }

function TxlsSheetList.GetItems(Index: integer): TxlsSheet;
begin
  Result := TxlsSheet(inherited Items[Index]);
end;

procedure TxlsSheetList.SetItems(Index: integer; Value: TxlsSheet);
begin
  inherited Items[Index] := Value;
end;

function TxlsSheetList.Add(Item: TxlsSheet): integer;
begin
  Result := inherited Add(Item);
  Item.FIndex := Result;
end;

procedure TxlsSheetList.Insert(Index: integer; Item: TxlsSheet);
begin
  inherited Insert(Index, Item);
end;

{ TxlsWorkSheet }

constructor TxlsWorkSheet.Create(Workbook: TxlsWorkbook);
begin
  inherited Create(Workbook);
  FRows := TxlsRowList.Create(Workbook);
  FCols := TxlsColList.Create(Workbook);
  FShrFmlaList := TbiffShrFmlaList.Create(Workbook);
end;

destructor TxlsWorkSheet.Destroy;
begin
  //Clear;
  if Assigned(FRows) then ObjFreeAndNil(FRows);
  if Assigned(FCols) then ObjFreeAndNil(FCols);
  if Assigned(FShrFmlaList) then ObjFreeAndNil(FShrFmlaList);
  inherited Destroy;
end;

procedure TxlsWorkSheet.Clear;
begin
  if Assigned(FRows) then FRows.Clear;
  if Assigned(FCols) then FCols.Clear;
  if Assigned(FShrFmlaList) then FshrFmlaList.Clear;
  inherited Clear;
end;

procedure TxlsWorkSheet.Load(AStream: TStream; ABof: TbiffBOF);
begin
  Self.LoadRows(AStream, ABof, 0);
end;

procedure TxlsWorkSheet.LoadRows(AStream: TStream; ABof: TbiffBOF; ARowCount: Cardinal);
var
  Header: TBIFF_Header;
  R: TbiffRecord;
  FLastFormula: TbiffFormula;
  Del : Boolean;
  I: Cardinal;

  procedure ReadRec(IsAdd: Boolean);
  begin
    if (AStream.Read(Header, SizeOf(Header)) <> SizeOf(Header)) then
      raise ExlsFileError.Create(sExcelInvalid);

    R := LoadRecord(Self, AStream, Header);
    del := True;
    try
      if (R is TbiffFormula) and IsAdd then
      begin
        FLastFormula := R as TbiffFormula;
        del := False;
      end;

      if (R is TbiffBOF) then
        raise ExlsFileError.Create(sExcelInvalid)
      else if (R is TbiffCell) and IsAdd then begin
        AddCell(R as TbiffCell);
        del := False;
      end
      else if (R is TbiffMultiple) and IsAdd then
        AddMultiple(R as TbiffMultiple)
      else if (R is TbiffShrFmla) and IsAdd then begin
        FShrFmlaList.Add(R as TbiffShrFmla);
        del := False;
      end
      else if (R is TbiffString) and IsAdd then
      begin
        if not Assigned(FLastFormula) then
          raise ExlsFileError.Create(sExcelInvalid)
        else
          FLastFormula.Value := (R as TbiffString).Value;
      end
      else if (R is TbiffEOF) then begin
        Eof := (R as TbiffEOF);
        del := False;
      end;
    except
      if Assigned(R) then R.Free;
      raise;
    end;
    if del then R.Free;
  end;

begin
  Clear;
  FLastFormula := nil;
  I := 0;
  if ARowCount > 0 then
    repeat
      if (I <= ARowCount) then
      begin
        ReadRec(True);
        if {igorp ticket 31044}(not del) and {/igorp} (R is TbiffCell) and (I < (R as TbiffCell).Row) and ( I <= ARowCount) then
          Inc(I);
      end
      else
        ReadRec(False);
    until Header.ID = BIFF_EOF
  else
    repeat
      ReadRec(True);
    until Header.ID = BIFF_EOF;
  FixFormulas;
  Self.Bof := ABof;
end;

procedure TxlsWorkSheet.AddCell(Cell: TbiffCell);
var
  i: integer;
begin
  for i := FRows.Count to Cell.Row do
    FRows.Add(TxlsRow.Create(FWorkbook));
  FRows[Cell.Row].Add(Cell);
  for i := FCols.Count to Cell.Col do
    FCols.Add(TxlsCol.Create(FWorkbook));
  FCols[Cell.Col].Add(Cell);
end;

procedure TxlsWorkSheet.AddMultiple(Multiple: TbiffMultiple);
var
  Cell: TbiffCell;
begin
  while not Multiple.EOF do begin
    Cell := Multiple.Cell;
    AddCell(Cell);
  end;
end;

procedure TxlsWorkSheet.FixFormulas;
var
  i, j, Index: integer;
  Formula: TbiffFormula;
begin
  for i := 0 to FRows.Count - 1 do begin
    for j := 0 to FRows[i].Count - 1 do begin
      if FRows[i][j] is TbiffFormula then begin
        Formula := FRows[i][j] as TbiffFormula;
        if not Formula.IsExp then Continue;
        if not ShrFmlaList.Find(Formula.Key, Index) then
          raise Exception.Create(sShrFmlaNotFound);
        Formula.MixShared(ShrFmlaList[Index].Data,
          ShrFmlaList[Index].DataSize);
      end;
    end;
    FRows[i].RecalculateTotalSize;
  end;
  for i := 0 to FCols.Count - 1 do
    FCols[i].RecalculateTotalSize;
end;

function TxlsWorkSheet.GetRowCount: integer;
begin
  Result := FRows.Count;
end;

function TxlsWorkSheet.GetColCount: integer;
begin
  Result := FCols.Count;
end;

function TxlsWorkSheet.GetCells(Row, Col: integer): TbiffCell;
var
  Index: integer;
begin
  Result := nil;

  if (Row < 0) or (Row > MAX_ROW_COUNT) then
    raise ExlsFileError.CreateFmt(sInvalidRow, [Row]);
  if (Col < 0) or (Col > MAX_COL_COUNT) then
    raise ExlsFileError.CreateFmt(sInvalidCol, [Col]);

  if (Row >= GetRowCount) or (Col >= GetColCount) then Exit;

  if FRows[Row].Find(Col, Index) then
    Result := FRows[Row][Index];

  if not Assigned(Result) or (Result.Col <> Col) or (Result.Row <> Row) then
  begin
    Result := nil;
    if FCols[Col].Find(Row, Index) then
    begin
      Result :=FCols[Col][Index];

      if (Result.Col <> Col) or (Result.Row <> Row) then
        Result := nil;
    end;
  end;
end;

{ TxlsWorkSheetList }

function TxlsWorkSheetList.GetItems(Index: integer): TxlsWorkSheet;
begin
  Result := TxlsWorkSheet(inherited Items[Index]);
end;

procedure TxlsWorkSheetList.SetItems(Index: integer; Value: TxlsWorkSheet);
begin
  inherited Items[Index] := Value;
end;

function TxlsWorkSheetList.Add(Item: TxlsWorkSheet): integer;
begin
  Result := inherited Add(Item);
  Items[Result].OnDestroy := OnDestroyItem;
end;

procedure TxlsWorkSheetList.Insert(Index: integer; Item: TxlsWorkSheet);
begin
  inherited Insert(Index, Item);
end;

function TxlsWorkSheetList.IndexOfName(const Name: WideString): integer;
var
  i: integer;
begin
  Result := -1;
  for i := 0 to Count - 1 do
  begin
    if AnsiCompareText(Name, Items[i].Name) = 0 then
    begin
      Result := i;
      Break;
    end;
  end;
end;

procedure TxlsWorkSheetList.OnDestroyItem(Sender: TObject);
var
  i: integer;
begin
  for i := Count - 1 downto 0 do
    if Items[i] = Sender then begin
      Remove(Items[i]);
      Break;
    end;
end;

{ TxlsWorkbook }

constructor TxlsWorkbook.Create(ExcelFile: TxlsFile);
begin
  inherited Create;
  FExcelFile := ExcelFile;
  FGlobals := TxlsGlobals.Create(Self);
  FSheets := TxlsSheetList.Create(Self);
  FWorkSheets := TxlsWorkSheetList.Create(Self);
end;

destructor TxlsWorkbook.Destroy;
begin
  FSheets.Clear;
  FSheets.Free;
  FWorkSheets.Clear;
  FWorkSheets.Free;
  FGlobals.Clear;
  FGlobals.Free;
  inherited;
end;

procedure TxlsWorkbook.Load(Stream: TStream);
var
  Header: TBIFF_Header;
  R: TbiffRecord;
  WorkSheet: TxlsWorkSheet;
begin
  FSheets.Clear;
  FGlobals.Clear;
  Stream.Seek(soFromBeginning, 0);
  while (Stream.Read(Header, SizeOf(Header)) = SizeOf(Header)) and
        ((Header.ID <> BIFF_EOF) and (Header.ID <> 0)) do
  begin
    R := LoadRecord(FGlobals, Stream, Header);
    if (Header.ID = BIFF_BOF) or (Header.ID = BIFF_BOF_3) then
      case (R as TbiffBOF).BOFType of
        BIFF_BOF_GLOBALS  :
          FGlobals.Load(Stream, R as TbiffBOF);
        BIFF_BOF_WORKSHEET: begin
          WorkSheet := TxlsWorkSheet.Create(Self);
          (FSheets[FSheets.Add(WorkSheet)] as TxlsWorkSheet).LoadRows(Stream, R as TbiffBOF, FExcelFile.FRowCount);
          FWorkSheets.Add(WorkSheet);
        end;
        BIFF_BOF_CHART: begin
        end;
        else raise ExlsFileError.Create(sExcelInvalid);
      end
    else raise ExlsFileError.Create(sExcelInvalid);
  end;
  FGlobals.SSTList.Sort;
end;

procedure TxlsWorkbook.Clear;
begin
  FGlobals.Clear;
  FSheets.Clear;
  FWorkSheets.Clear;
end;

{ TxlsFile }

constructor TxlsFile.Create;
begin
  inherited;
  FWorkbook := TxlsWorkbook.Create(Self);
  FLoaded := false;
end;

destructor TxlsFile.Destroy;
begin
  Clear;
  FWorkbook.Free;
  inherited;
end;

procedure TxlsFile.Load;
begin
  LoadRows(0);
end;

procedure TxlsFile.LoadRows( ARowCount: Cardinal);
begin
  Open;
  try
    FRowCount := ARowCount;
    FWorkbook.Load(FStream);
    FLoaded := true;
  finally
    Close;
  end;
end;


procedure TxlsFile.Clear;
begin
  FWorkbook.Clear;
  FLoaded := false;
end;

procedure TxlsFile.Open;
begin
  if FFileName = EmptyStr then
    raise ExlsFileError.Create(sFileNameNotDefined);
  if not FileExists(FFileName) then
    raise ExlsFileError.CreateFmt(sFileDoesNotExist, [FFileName]);
  OpenStream;
end;

procedure TxlsFile.OpenStream;
const
  ReadStorageOptions = STGM_DIRECT or STGM_SHARE_EXCLUSIVE or STGM_READ;
var
  WC: PWideChar;
  IsWorkbook: boolean;
  Enum: IEnumStatStg;
  Stat: TSTATSTG;
  OLEInputStream: TOleStream;
  FileInputStream: TFileStream;
  FStorage: IStorage;
  FOleStream: IStream;
begin
  GetMem(WC, 512);
  try
    StringToWideChar(FFileName, WC, 512);
    if StgIsStorageFile(WC) = S_OK then
    begin
      IsWorkBook := false;
      OleCheck(StgOpenStorage(WC, nil, STGM_READ or STGM_SHARE_DENY_WRITE,
                              nil, 0, FStorage));
      OleCheck(FStorage.EnumElements(0, nil, 0, Enum));

      while Enum.Next(1, Stat, nil) = S_OK do
      begin
        IsWorkbook := Stat.pwcsName = 'Workbook';
        if IsWorkbook then
        begin
          OleCheck(FStorage.OpenStream('Workbook', nil, ReadStorageOptions, 0,
            FOleStream));
          if not Assigned(FStream) then
            FStream := TMemoryStream.Create;
          OLEInputStream := TOleStream.Create(FOleStream);
          try
            FStream.CopyFrom(OLEInputStream, OLEInputStream.Size);
          finally
            OLEInputStream.Free;
          end;
          Break;
        end;
      end;

      if not IsWorkBook then
        raise ExlsFileError.CreateFmt(sFileIsNotExcelWorkbook, [FFileName]);
    end
    else begin
      if not Assigned(FStream) then
        FStream := TMemoryStream.Create;
      FileInputStream := TFileStream.Create(FFileName, fmOpenRead);
      try
        FStream.CopyFrom(FileInputStream, FileInputStream.Size);
      finally
        FileInputStream.Free;
      end;
    end;
  finally
    FreeMem(WC);
  end;
end;

procedure TxlsFile.Close;
begin
  if Assigned(FStream) then ObjFreeAndNil(FStream);
end;

{ TbiffLabel }
function TbiffLabel.GetAsString: WideString;
var
  Str: TxlsString;
  Offset: integer;
  Tmp: TbiffRecord;
  begin
  Tmp := Self;
  Offset := 6;
  Str := TxlsString.CreateR(True, Tmp, Offset);
  try
    Result := Str.Value;
  finally
    Str.Free;
  end;
end;

function TbiffLabel.GetAsVariant: variant;
begin
  Result := GetAsString;
end;

function TbiffLabel.GetCellType: TbiffCellType;
begin
  Result := bctString;
end;

end.
