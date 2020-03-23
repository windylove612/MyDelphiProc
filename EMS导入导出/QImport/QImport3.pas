unit QImport3;

{$R QIResStr.res}
{$R QIEULA.res}

{$I QImport3VerCtrl.Inc}

{$IFDEF VCL6}
  {$WARN UNIT_PLATFORM OFF}
{$ENDIF}

interface

uses
  {$IFDEF VCL16}
    System.Classes,
    Data.DB,
    System.SysUtils,
    System.IniFiles,
    {$IFNDEF NOGUI}
      Vcl.DBGrids,
      Vcl.ComCtrls,
      Vcl.Grids,
    {$ENDIF}
  {$ELSE}
    Classes,
    DB,
    SysUtils,
    IniFiles,
    {$IFNDEF NOGUI}
      DBGrids,
      ComCtrls,
      Grids,
    {$ENDIF}
  {$ENDIF}
  QImport3EZDSLHsh,
  {$IFDEF USESCRIPT}
  QImport3ScriptEngine,
  {$ENDIF}
  QImport3StrTypes;

type
  TQImport3 = class;
  TAllowedImport = (aiXLS, aiDBF, aiXML, aiTXT, aiCSV, aiAccess, aiHTML,aiXMLDoc,
    aiXlsx, aiDocx, aiODS, aiODT);
  TAllowedImports = set of TAllowedImport;

{$IFDEF QI_UNICODE}
  WideException = class(Exception)
  private
    FWideMessage: WideString;
  public
    constructor Create(const Msg: WideString);
    constructor CreateFmt(const Msg: WideString; const Args: array of const);
    property Message: WideString read FWideMessage write FWideMessage;
  end;
{$ENDIF}

  TQICharsetType = (
    ctWinDefined, ctLatin1, ctArmscii8, ctAscii, ctCp850, ctCp852, ctCp866,
    ctCp1250, ctCp1251, ctCp1256, ctCp1257, ctDec8, ctGeostd8, ctGreek,
    ctHebrew, ctHp8, ctKeybcs2, ctKoi8r, ctKoi8u, ctLatin2, ctLatin5,
    ctLatin7, ctMacce, ctMacroman, ctSwe7, ctUtf8, ctUtf16, ctUtf32,
    // unique in postrgesql
    ctLatin3, ctLatin4, ctLatin6, ctLatin8, ctIso8859_5, ctIso8859_6,
    //unique in db2
    ctCp1026, ctCp1254, ctCp1255, ctCp1258, ctCp437, ctCp500, ctCp737, ctCp855,
    ctCp856, ctCp857, ctCp860, ctCp862, ctCp863, ctCp864, ctCp865, ctCp869,
    ctCp874, ctCp875, ctIceland,
    //unique in IB/FB
    ctBig5, ctKSC5601, ctEUC, ctGB2312, ctSJIS_0208, ctLatin9, ctLatin13,
    ctCp1252, ctCp1253, ctCp775, ctCp858 );

  TQuoteAction = (qaNone, qaAdd, qaRemove);
  TQImportCharCase = (iccNone, iccUpper, iccLower, iccUpperFirst, iccUpperFirstWord);
  TQImportCharSet = (icsNone, icsAnsi, icsOem);

  TLocalizeEvent = procedure(StringID: Integer; var ResultString: String) of object;

  TQImportLocale = class(TObject)
  private
    FDllHandle: Cardinal;
    FLoaded: Boolean;
    FOnLocalize: TLocalizeEvent;
    FIDEMode: Boolean;
  public
    constructor Create;
    function LoadStr(ID: Integer): String;
    procedure LoadDll(const Name: string);
    procedure UnloadDll;
    property OnLocalize: TLocalizeEvent read FOnLocalize write FOnLocalize;
  end;

  TQImportFormats = class(TPersistent)
  private
    FDecimalSeparator: Char;
    FThousandSeparator: Char;
    FShortDateFormat: string;
    FLongDateFormat: string;
    FDateSeparator: Char;
    FShortTimeFormat: string;
    FLongTimeFormat: string;
    FTimeSeparator: Char;
    FBooleanTrue: TStrings;
    FBooleanFalse: TStrings;
    FNullValues: TStrings;
    FOldDecimalSeparator: Char;
    FOldThousandSeparator: Char;
    FOldShortDateFormat: string;
    FOldLongDateFormat: string;
    FOldDateSeparator: Char;
    FOldShortTimeFormat: string;
    FOldLongTimeFormat: string;
    FOldTimeSeparator: Char;

    procedure SetBooleanTrue(const Value: TStrings);
    procedure SetBooleanFalse(const Value: TStrings);
    procedure SetNullValues(const Value: TStrings);
  protected
    procedure DefineProperties(Filer: TFiler); override;

    procedure ReadShortDateFormat(Reader: TReader);
    procedure ReadLongDateFormat(Reader: TReader);
    procedure ReadShortTimeFormat(Reader: TReader);
    procedure ReadLongTimeFormat(Reader: TReader);

    procedure WriteShortDateFormat(Writer: TWriter);
    procedure WriteLongDateFormat(Writer: TWriter);
    procedure WriteShortTimeFormat(Writer: TWriter);
    procedure WriteLongTimeFormat(Writer: TWriter);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    procedure StoreFormats;
    procedure RestoreFormats;
    procedure ApplyParams;
  published
    property DecimalSeparator: Char read FDecimalSeparator
      write FDecimalSeparator stored True;
    property ThousandSeparator: Char read FThousandSeparator
      write FThousandSeparator stored True;
    property ShortDateFormat: string read FShortDateFormat
      write FShortDateFormat stored False;
    property LongDateFormat: string read FLongDateFormat
      write FLongDateFormat stored False;
    property DateSeparator: Char read FDateSeparator write FDateSeparator
      stored True;
    property ShortTimeFormat: string read FShortTimeFormat
      write FShortTimeFormat stored False;
    property LongTimeFormat: string read FLongTimeFormat
      write FLongTimeFormat stored False;
    property TimeSeparator: Char read FTimeSeparator write FTimeSeparator
      stored True;
    property BooleanTrue: TStrings read FBooleanTrue write SetBooleanTrue;
    property BooleanFalse: TStrings read FBooleanFalse write SetBooleanFalse;
    property NullValues: TStrings read FNullValues write SetNullValues;
  end;

  TQImportReplacement = class(TCollectionItem)
  private
    FTextToFind: qiString;
    FReplaceWith: qiString;
    FIgnoreCase: Boolean;
  protected
    function  GetDisplayName: string; override;
  public
    constructor Create(Collection: TCollection); override;
    procedure Assign(Source: TPersistent); override;
  published
    property TextToFind: qiString
      read FTextToFind write FTextToFind;
    property ReplaceWith: qiString
      read FReplaceWith write FReplaceWith;
    property IgnoreCase: Boolean read FIgnoreCase
      write FIgnoreCase default False;
  end;

  TQImportReplacements = class(TCollection)
  private
    FHolder: TPersistent;
    function GetItem(Index: integer): TQImportReplacement;
    procedure SetItem(Index: integer; Replacement: TQImportReplacement);
  protected
    function GetOwner: TPersistent; override;
  public
    property Holder: TPersistent read FHolder;
    constructor Create(Holder: TPersistent);
    function Add: TQImportReplacement;
    property Items[Index: integer]: TQImportReplacement read GetItem
      write SetItem; default;
    function ItemExists(
      const ATextToFind, AReplaceWith: qiString;
      AIgnoreCase: Boolean): Boolean;
  end;

  TQImportFunctions = (ifNone, ifDate, ifTime, ifDateTime, ifLongFileName,
    ifShortFileName);

  TQImportFieldFormat = class(TCollectionItem)
  private
    FFieldName: string;
    FGeneratorValue: Int64;
    FGeneratorStep: Integer;

    FConstantValue: qiString;
    FNullValue: qiString;
    FDefaultValue: qiString;
    FLeftQuote: qiString;
    FRightQuote: qiString;

    FQuoteAction: TQuoteAction;
    FCharCase: TQImportCharCase;
    FCharSet: TQImportCharSet;
    FReplacements: TQImportReplacements;
    FFunctions: TQImportFunctions;
    FScript: TqiStrings;

    function IsConstant: Boolean;
    function IsNull: Boolean;
    function IsDefault: Boolean;
    function IsLeftQuote: Boolean;
    function IsRightQuote: Boolean;

    procedure SetReplacements(const Value: TQImportReplacements);
  protected
    function GetDisplayName: string; override;
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    procedure Assign(Source: TPersistent); override;
    function IsDefaultValues: boolean;
  published
    property FieldName: string read FFieldName write FFieldName;
    property GeneratorValue: Int64 read FGeneratorValue
      write FGeneratorValue default 0;
    property GeneratorStep: integer read FGeneratorStep
      write FGeneratorStep default 0;

    property ConstantValue: qiString
      read FConstantValue write FConstantValue stored IsConstant;
    property NullValue: qiString
      read FNullValue write FNullValue stored IsNull;
    property DefaultValue: qiString
      read FDefaultValue write FDefaultValue stored IsDefault;
    property LeftQuote: qiString
      read FLeftQuote write FLeftQuote stored IsLeftQuote;
    property RightQuote: qiString
      read FRightQuote write FRightQuote stored IsRightQuote;

    property QuoteAction: TQuoteAction read FQuoteAction
      write FQuoteAction default qaNone;
    property CharCase: TQImportCharCase read FCharCase
      write FCharCase default iccNone;
    property CharSet: TQImportCharSet read FCharSet
      write FCharSet default icsNone;
    property Replacements: TQImportReplacements read FReplacements
      write SetReplacements;
    property Functions: TQImportFunctions read FFunctions write FFunctions;
    {$IFNDEF USESCRIPT}
    protected
    {$ENDIF}
    property Script: TqiStrings read FScript write FScript;
  end;

  TQImportFieldFormats = class(TCollection)
  private
    FHolder: TComponent;
    function GetItem(Index: integer): TQImportFieldFormat;
    procedure SetItem(Index: integer; FieldFormat: TQImportFieldFormat);
  protected
    function GetOwner: TPersistent; override;
  public
    property Holder: TComponent read FHolder;
    constructor Create(AHolder: TComponent);
    function Add: TQImportFieldFormat;
    function IndexByName(const FieldName: string): integer;
    property Items[Index: integer]: TQImportFieldFormat read GetItem write SetItem; default;
  end;

  TQImportCol = class
  private
    FValue: qiString;
    FName: string;
    FIsBinary: Boolean;
    FColumnIndex: Integer;
  public
    constructor Create;
    property Name: string read FName write FName;
    property Value: qiString read FValue write FValue;
    property IsBinary: Boolean read FIsBinary;
  end;

  TQImportRow = class(TList)
  private
    FMapNameIdxHash: THashTable;
    FColHash: THashTable;
    FQImport: TQImport3;

    function Get(Index: Integer): TQImportCol;
    procedure Put(Index: Integer; const Value: TQImportCol);
  public
    constructor Create(AImport: TQImport3);
    destructor Destroy; override;
    function Add(const AName: string): TQImportCol;
    procedure Clear; {$IFNDEF VCL3} override; {$ENDIF}
    procedure Delete(Index: integer);
    function First: TQImportCol;
    procedure SetValue(const AName, AValue: qiString;
      AIsBinary: Boolean);
    procedure ClearValues;
    function Last: TQImportCol;
    function IndexOf(Item: TQImportCol): Integer;

    function ColByName(const AName: string): TQImportCol;

    property QImport: TQImport3 read FQImport;
    property Items[Index: Integer]: TQImportCol read Get write Put; default;
    property MapNameIdxHash: THashTable read FMapNameIdxHash;
  end;

  TQImportGenerator = class
  private
    FName: string;
    FValue: Int64;
    FStep: Int64;
    FIsFirstRequest: Boolean;
  public
    constructor Create; 
    function GetNewValue: Int64;
    property Name: string read FName write FName;
    property Value: Int64 read FValue write FValue;
    property Step: Int64 read FStep write FStep;
  end;

  TQImportGenerators = class(TList)
  private
    function Get(Index: Integer): TQImportGenerator;
    procedure Put(Index: Integer; const Value: TQImportGenerator);
  public
    destructor Destroy; override;
    function Add(const AName: string; AValue, AStep: Int64): TQImportGenerator;
    procedure Delete(Index: integer);
    function GetNewValue(const AName: string): Int64;
    function GenByName(const AName: string): TQImportGenerator;

    property Items[Index: Integer]: TQImportGenerator read Get write Put; default;
  end;

  TQImportFieldType = (iftUnknown, iftString, iftInteger, iftBoolean,
    iftDouble, iftCurrency, iftDateTime, iftBytes);
  EQImportError = class(Exception);
  TQImportAddType = (qatAppend, qatInsert);
  TQImportDestination = (qidDataSet, qidDBGrid, qidListView, qidStringGrid,
    qidUserDefined);
  TQImportResult = (qirOk, qirContinue, qirBreak);
  TQImportMode = (qimInsertAll, qimInsertNew, qimUpdate, qimUpdateOrInsert,
    qimDelete, qimDeleteOrInsert);
  TQImportAction = (qiaNone, qiaInsert, qiaUpdate, qiaDelete);

  TImportCancelEvent = procedure(Sender: TObject;
    var Continue: boolean) of object;
  TImportBeforePostEvent = procedure(Sender: TObject;
    Row: TQImportRow; var Accept: boolean) of object;
  TUserDefinedImportEvent = procedure(Sender: TObject;
    Row: TQImportRow) of object;
  TImportAfterPostEvent = procedure(Sender: TObject;
    Row: TQImportRow) of object;
  TImportLoadTemplateEvent = procedure(Sender: TObject;
    const FileName: string) of object;
  TDestinationLocateEvent = procedure(Sender: TObject; KeyColumns: TStrings;
    Row: TQImportRow; var KeyFields: string; var KeyValues: Variant) of object;
  TSetCharsetTypeEvent = procedure(Sender: TObject; const Charset: AnsiString) of object;
  TWideStringToCharsetEvent = procedure(Sender: TObject;
    const SourceStr: WideString; var EncodedStr: AnsiString) of object;
  {$IFDEF USESCRIPT}
    TScriptExecuteError = procedure (Sender: TObject;
      const Error: TScriptErrorMsg; const ColName: String;
      var RaiseError: Boolean; const ErrorMessage: qiString) of object;
  {$ENDIF}

  TQImport3 = class(TComponent)
  private
    {$IFDEF USESCRIPT}
    FScriptEngine: TQImport3ScriptEngine;
    {$ENDIF}
    FLastError: qiString;
    FDataSet: TDataSet;
{$IFNDEF NOGUI}
    FDBGrid: TDBGrid;
    FListView: TListView;
    FStringGrid: TStringGrid;
    FGridCaptionRow: integer;
    FGridStartRow: integer;
{$ENDIF}
    FAutoTrimValue: Boolean;
    FImportEmptyRows: Boolean;
    FRowIsEmpty: Boolean;

    FFileName: string;
    FErrors: TqiStrings;
    FMap: TStrings;
    FImportRecCount: integer;
    FCommitRecCount: integer;
    FCommitAfterDone: boolean;
    FErrorLog: boolean;
    FErrorLogFileName: string;
    FRewriteErrorLogFile: boolean;
    FShowErrorLog: boolean;
    FErrorLogFS: TFileStream;
//    FSQLLog: boolean;
//    FSQLLogFileName: string;
//    FSQLLogFileRewrite: boolean;
//    FSQL: TFileStream;
    FSkipFirstRows: integer;
    FSkipFirstCols: integer;
    FImportedRecs: integer;
    FCanceled: boolean;
    FFormats: TQImportFormats;
    FFieldFormats: TQImportFieldFormats;
    FAddType: TQImportAddType;
    FImportDestination: TQImportDestination;
    FImportMode: TQImportMode;
    FKeyColumns: TStrings;
    FCurrentLineNumber: Integer;

    FIsCSV: boolean;
    FLastAction: TQImportAction;
    FPrevAction: TQImportAction;

    FStream: TStream;
    FComma: AnsiChar;
    FQuote: AnsiChar;

    FOnBeforeImport: TNotifyEvent;
    FOnAfterImport: TNotifyEvent;
    FOnImportRecord: TNotifyEvent;
    FOnImportError: TNotifyEvent;
    FOnImportErrorAdv: TNotifyEvent;
    FOnNeedCommit: TNotifyEvent;
    FOnImportCancel: TImportCancelEvent;
    FOnBeforePost: TImportBeforePostEvent;
    FOnAfterPost: TImportAfterPostEvent;
//    FOnGetSQLIdentifier: TImportSQLIdentifierEvent;
    FOnUserDefinedImport: TUserDefinedImportEvent;
    FOnImportRowComplete: TUserDefinedImportEvent;
    FOnDestinationLocate: TDestinationLocateEvent;
    {$IFDEF USESCRIPT}
      FOnScriptExecuteError: TScriptExecuteError;
    {$ENDIF}

    FAbout: string;
    FVersion: string;
    FMappedColumns: TStrings;
    FAllowDuplicates: Boolean;
    procedure SetDataSet(const Value: TDataSet);
{$IFNDEF NOGUI}
    procedure SetDBGrid(const Value: TDBGrid);
    procedure SetListView(const Value: TListView);
    procedure SetStringGrid(const Value: TStringGrid);
{$ENDIF}
    procedure SetKeyColumns(const Value: TStrings);

    procedure SetFileName(const Value: string);
    procedure SetMap(const Value: TStrings);
    function GetErrorRecs: integer;
    procedure SetFormats(const Value: TQImportFormats);
    procedure SetFieldFormats(const Value: TQImportFieldFormats);
  private
{$IFNDEF NOGUI}
    FCurrListItem: TListItem;
    FCurrStrGrRow: Integer;
    FFindPos: Integer;
    FAddPos: Integer;
{$ENDIF}
    procedure InitializeImportRow;
  protected
    FTotalRecCount: integer;
    FImportRow: TQImportRow;
    FImportGenerators: TQImportGenerators;

    property IsCSV: boolean read FIsCSV;

    procedure Notification(AComponent: TComponent;
      Operation: TOperation); override;

    procedure DoImport;
    procedure DoUserDefinedImport; virtual;
    procedure DoImportRowComplete; virtual;

    procedure BeforeImport; virtual;
    procedure StartImport; virtual; abstract;
    function CheckCondition: boolean; virtual; abstract;
    function Skip: boolean; virtual; abstract;
    procedure FillImportRow; virtual; abstract;
    function ImportData: TQImportResult; virtual; abstract;
    procedure DataManipulation;
    procedure ChangeCondition; virtual; abstract;
    procedure FinishImport; virtual; abstract;
    procedure AfterImport; virtual;

    procedure DoAfterSetFileName; virtual;
    function CheckProperties: Boolean; virtual;
    procedure DoUserDataFormat(Col: TQImportCol);
    function CanContinue: boolean;
    function AllowImportRowComplete: Boolean; virtual;


    function StringToField(Field: TField;
      const Str: qiString;
      AIsBinary: Boolean): string;
    function PrepareImportValue(const AValue: qiString; const AIsBinary: Boolean;
      const AFieldType: TFieldType): Variant;

    procedure DoLoadConfiguration(IniFile: TIniFile); virtual;
    procedure DoSaveConfiguration(IniFile: TIniFile); virtual;

    property SkipFirstRows: integer read FSkipFirstRows write FSkipFirstRows;
    property SkipFirstCols: integer read FSkipFirstCols write FSkipFirstCols;
    property RowIsEmpty: Boolean read FRowIsEmpty write FRowIsEmpty;
  private
    {$IFDEF USESCRIPT}
    procedure SetScriptEngine(const Value: TQImport3ScriptEngine);
    {$ENDIF}
    function StringToBoolean(const Str: string): Boolean;
    function StrIsNull(const Str: string): Boolean;
  protected
    function GetFloatValue(const AValue: qiString; const AFieldType: TFieldType):
        Variant; virtual;
    function GetIntegerValue(const AValue: qiString; const AFieldType: TFieldType):
        Variant; virtual;
    function GetDateTimeValue(const AValue: qiString): Variant; virtual;
    function GetStringValue(const AValue: qiString; const AFieldType: TFieldType): Variant; virtual;
    function GetBitesValue(const AValue: qiString): Variant; virtual;
    procedure DestinationInsert;
    procedure DestinationEdit;
    procedure DestinationDelete;
    procedure DestinationSetValues;
    procedure DestinationPost;
    procedure DestinationCancel;
    function  DestinationFindColumn(const ColName: string): integer;
    procedure DestinationDisableControls;
    procedure DestinationEnableControls;
    procedure CheckDestination;

    function DestinationFindByKey: boolean;
    function DestinationFindByFields: Boolean;
    function DestinationColCount: integer;
    function DestinationColName(Index: integer): string;

    procedure DoBeginImport; dynamic;
    function DoBeforePost: Boolean; dynamic;
    procedure DoAfterPost; dynamic;
    procedure DoImportRecord; dynamic;
    procedure DoImportError(Error: Exception); dynamic;
    procedure WriteErrorLog(const ErrorMsg: string);
    procedure DoNeedCommit; dynamic;
    procedure DoEndImport; dynamic;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
    function Execute: boolean;
    procedure ImportToCSV(Stream: TStream; Comma, Quote: AnsiChar);
    procedure Cancel;
    procedure LoadConfiguration(const FileName: string);
    procedure SaveConfiguration(const FileName: string);
    property Errors: TqiStrings read FErrors;
    property ImportedRecs: integer read FImportedRecs;
    property ErrorRecs: integer read GetErrorRecs;
    property Canceled: boolean read FCanceled;
    property LastAction: TQImportAction read FLastAction write FLastAction;
    property TotalRecCount: integer read FTotalRecCount;
    property FileName: string read FFileName write SetFileName;
    property CurrentLineNumber: Integer read FCurrentLineNumber
      write FCurrentLineNumber;
    property LastError: qiString read FLastError write FLastError;
  published
    property About: string read FAbout write FAbout;
    property AutoTrimValue: Boolean read FAutoTrimValue write FAutoTrimValue
      default False;
    property ImportEmptyRows: Boolean read FImportEmptyRows
      write FImportEmptyRows default True;
    property Version: string read FVersion write FVersion;

    property DataSet: TDataSet read FDataSet write SetDataSet;
{$IFNDEF NOGUI}
    property DBGrid: TDBGrid read FDBGrid write SetDBGrid;
    property ListView: TListView read FListView write SetListView;
    property StringGrid: TStringGrid read FStringGrid write SetStringGrid;
    property GridCaptionRow: integer read FGridCaptionRow
      write FGridCaptionRow default -1;
    property GridStartRow: integer read FGridStartRow
      write FGridStartRow default -1;
{$ENDIF}

    property ImportDestination: TQImportDestination read FImportDestination
      write FImportDestination default qidDataSet;
    property ImportMode: TQImportMode read FImportMode
      write FImportMode default qimInsertAll;

    property Map: TStrings read FMap write SetMap;
    property Formats: TQImportFormats read FFormats write SetFormats;
    property FieldFormats: TQImportFieldFormats read FFieldFormats
      write SetFieldFormats;
    property ErrorLog: boolean read FErrorLog write FErrorLog default false;
    property ErrorLogFileName: string read FErrorLogFileName
      write FErrorLogFileName;
    property RewriteErrorLogFile: boolean read FRewriteErrorLogFile
      write FRewriteErrorLogFile default true;
    property ShowErrorLog: boolean read FShowErrorLog
      write FShowErrorLog default false;
//    property SQLLog: boolean read FSQLLog write FSQLLog default false;
//    property SQLLogFileName: string read FSQLLogFileName
//      write FSQLLogFileName;
//    property SQLLogFileRewrite: boolean read FSQLLogFileRewrite
//      write FSQLLogFileRewrite default true;
    property ImportRecCount: integer read FImportRecCount
      write FImportRecCount default 0;
    property CommitRecCount: integer read FCommitRecCount
      write FCommitRecCount default 1000;
    property CommitAfterDone: boolean read FCommitAfterDone
      write FCommitAfterDone default False;
    property AddType: TQImportAddType read FAddType
      write FAddType default qatAppend;
    property KeyColumns: TStrings read FKeyColumns write SetKeyColumns;
    property AllowDuplicates: Boolean read FAllowDuplicates
      write FAllowDuplicates default True;

    property OnBeforeImport: TNotifyEvent read FOnBeforeImport
      write FOnBeforeImport;
    property OnAfterImport: TNotifyEvent read FOnAfterImport
      write FOnAfterImport;
    property OnImportRecord: TNotifyEvent read FOnImportRecord
      write FOnImportRecord;
    property OnImportError: TNotifyEvent read FOnImportError
      write FOnImportError;
    property OnImportErrorAdv: TNotifyEvent read FOnImportErrorAdv
      write FOnImportErrorAdv;
    property OnNeedCommit: TNotifyEvent read FOnNeedCommit
      write FOnNeedCommit;
    property OnImportCancel: TImportCancelEvent read FOnImportCancel
      write FOnImportCancel;
    property OnBeforePost: TImportBeforePostEvent read FOnBeforePost
      write FOnBeforePost;
    property OnAfterPost: TImportAfterPostEvent read FOnAfterPost
      write FOnafterPost;
    property OnUserDefinedImport: TUserDefinedImportEvent
      read FOnUserDefinedImport write FOnUserDefinedImport;
    property OnImportRowComplete: TUserDefinedImportEvent
      read FOnImportRowComplete write FOnImportRowComplete;
//    property OnGetSQLIdentifier: TImportSQLIdentifierEvent
//      read FOnGetSQLIdentifier write FOnGetSQLIdentifier;
    property OnDestinationLocate: TDestinationLocateEvent read
      FOnDestinationLocate write FOnDestinationLocate;
    {$IFDEF USESCRIPT}
    property ScriptEngine: TQImport3ScriptEngine read FScriptEngine
      write SetScriptEngine;
    property OnScriptExecuteError: TScriptExecuteError read FOnScriptExecuteError
      write FOnScriptExecuteError;
    {$ENDIF}
  private
    FCustomImportMode: Boolean;
    FCustomImportError: Boolean;
    FTempFileCharset: AnsiString;
    FOnSetCharsetType: TSetCharsetTypeEvent;
    FOnWideStringToCharset: TWideStringToCharsetEvent;
  public
    property CustomImportMode: Boolean read FCustomImportMode
      write FCustomImportMode;
    property ImportRow: TQImportRow read FImportRow;
    property TempFileCharset: AnsiString read FTempFileCharset write
      FTempFileCharset;
    property OnSetCharsetType: TSetCharsetTypeEvent
      read FOnSetCharsetType write FOnSetCharsetType;
    property OnWideStringToCharset: TWideStringToCharsetEvent
      read FOnWideStringToCharset write FOnWideStringToCharset;
  end;

function GetImportFieldType(Field: TField): TQImportFieldType;
procedure StrReplace(var S: string; const Search, Replace: string;
  Flags: TReplaceFlags);

function QImportLocale: TQImportLocale;
function QImportLoadStr(ID: Integer): string;
function StringToHex(const Str: String): String;
function HexToString(const HexStr: string; out Str: string): Boolean;

implementation

uses
  {$IFDEF VCL16}
    Winapi.Windows,
    Vcl.FileCtrl,
    Winapi.ShellAPI,
    System.Variants,
  {$ELSE}
    Windows,
    FileCtrl,
    ShellAPI,
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
  {$ENDIF}
  {$IFDEF ADVANCED_DATA_IMPORT_TRIAL_VERSION}
    fuQImport3About,
  {$ENDIF}
  QImport3Common,
  QImport3StrIDs,
  QImport3WideStrUtils;

const
  LF = #13#10;
  sStreamMustBeAssigned = 'Stream must be assigned!';
  DefDecimalSeparator = '.';
var
  Locale: TQImportLocale = nil;

function StringToHex(const Str: String): String;
const
	HexDigits: array[0..15] of Char = '0123456789abcdef';
var
	I: Integer;
	B: Byte;
begin
	Result := '';
	for I := 1 to Length(Str) do
	begin
		B := Ord(Str[I]);
		Result := Result + HexDigits[B shr 4] + HexDigits[B and 15];
	end;
end;

{$WARNINGS OFF}
function HexToString(const HexStr: string; out Str: string): Boolean;
var
	I: Integer;
	V, B: Byte;
  Res: string;
begin
  Str := '';
  Result := Length(HexStr) mod 2 = 0;
  if not Result then
    Exit;
  V := 0;
  Res := '';
  for I := 1 to Length(HexStr) do
  begin
    case HexStr[I] of
      '0': B := 0;
      '1': B := 1;
      '2': B := 2;
      '3': B := 3;
      '4': B := 4;
      '5': B := 5;
      '6': B := 6;
      '7': B := 7;
      '8': B := 8;
      '9': B := 9;
      'a', 'A': B := 10;
      'b', 'B': B := 11;
      'c', 'C': B := 12;
      'd', 'D': B := 13;
      'e', 'E': B := 14;
      'f', 'F': B := 15;
      else begin
        Result := False;
        Exit;
      end;
    end;
    V := V shl 4 + B;
    if I mod 2 = 0 then
    begin
      Res := Res + Chr(V);
      V := 0;
    end;
  end;
  Str := Res;
end;
{$WARNINGS ON}

{$IFDEF ADVANCED_DATA_IMPORT_TRIAL_VERSION}
function IsIDERuning: Boolean;
begin
  Result := (FindWindow('TAppBuilder', nil) <> 0) or
            (FindWindow('TPropertyInspector', nil) <> 0) or
            (FindWindow('TAlignPalette', nil) <> 0);
end;
{$ENDIF}

{$IFDEF QI_UNICODE}
procedure _WideFmtStr(var Result: WideString; const Format: WideString;
  const Args: array of const);
var
  Len, BufLen: Integer;
  Buffer: array[0..4095] of WideChar;
begin
  BufLen := SizeOf(Buffer) div 2;
  if Length(Format) < (BufLen - BufLen div 4) then
    Len := WideFormatBuf(Buffer, BufLen - 1, Pointer(Format)^, Length(Format), Args)
  else
  begin
    BufLen := Length(Format);
    Len := BufLen;
  end;
  if Len >= BufLen - 1 then
  begin
    while Len >= BufLen - 1 do
    begin
      Inc(BufLen, BufLen);
      Result := '';          // prevent copying of existing data, for speed
      SetLength(Result, BufLen);
      Len := WideFormatBuf(Pointer(Result)^, BufLen - 1, Pointer(Format)^,
        Length(Format), Args);
    end;
    SetLength(Result, Len);
  end
  else
    SetString(Result, Buffer, Len);
end;

function WideFormatStr(const _Format: WideString; const Args: array of const): WideString;
begin
  _WideFmtStr(Result,_Format, Args);
end;
{$ENDIF}

function FieldTypeToImportFieldType(const FieldType: TFieldType):
    TQImportFieldType;
begin
  case FieldType of
    ftBlob,
    ftMemo,
    //igorp
    ftOraClob,
    ftOraBlob,
    //\igorp
    {$IFNDEF VCL3}
    ftWideString,
    {$IFDEF VCL10}
    ftWideMemo,
    {$ENDIF}
    {$ENDIF}
    ftString,
    ftGUID: Result := iftString;
    // ayz
    ftVarBytes,
    ftBytes: Result := iftBytes;
    //\ayz
    ftSmallint,
    ftInteger,
    {ab}ftAutoInc,{/ab}
{$IFNDEF VCL3}
    ftLargeInt,
{$ENDIF}
{$IFDEF VCL12}
    ftByte,
{$ENDIF}
    ftWord: Result := iftInteger;
    ftBoolean: Result := iftBoolean;
    ftFloat,
    ftBCD
{$IFDEF VCL6}
    , ftFMTBcd
{$ENDIF}
    : Result := iftDouble;
    ftCurrency: Result := iftCurrency;
    ftDate,
    ftTime,
    ftDateTime
{$IFDEF VCL6}
    , ftTimeStamp
{$ENDIF}
    : Result := iftDateTime;
  else
    Result := iftUnknown;
  end;
end;

function GetImportFieldType(Field: TField): TQImportFieldType;
begin
  Result := FieldTypeToImportFieldType(Field.DataType);
end;

procedure StrReplaceCS(var S: string; const Search, Replace: string; Flags: TReplaceFlags);
var
  ResultStr: string;
  SourcePtr: PChar;
  SourceMatchPtr: PChar;
  SearchMatchPtr: PChar;
  ResultPtr: PChar;
  SearchLength,
  ReplaceLength,
  ResultLength: Integer;
  C: Char;
begin
  SearchLength := Length(Search);
  ReplaceLength := Length(Replace);

  if Length(Search) >= ReplaceLength then
    ResultLength := Length(S)
  else
    ResultLength := ((Length(S) div Length(Search)) + 1) * Length(Replace);
  SetLength(ResultStr, ResultLength);
  ResultPtr := PChar(ResultStr);
  SourcePtr := PChar(S);
  C := Search[1];

  while True do
  begin
    while (SourcePtr^ <> C) and (SourcePtr^ <> #0) do
    begin
      ResultPtr^ := SourcePtr^;
      Inc(ResultPtr);
      Inc(SourcePtr);
    end;

    if SourcePtr^ = #0 then
      Break
    else
    begin

      SourceMatchPtr := SourcePtr + 1;
      SearchMatchPtr := PChar(Search) + 1;
      while (SourceMatchPtr^ = SearchMatchPtr^) and (SearchMatchPtr^ <> #0) do
      begin
        Inc(SourceMatchPtr);
        Inc(SearchMatchPtr);
      end;

      if SearchMatchPtr^ = #0 then
      begin
        Move((@Replace[1])^, ResultPtr^, ReplaceLength);
        Inc(SourcePtr, SearchLength);
        Inc(ResultPtr, ReplaceLength);

        if not (rfReplaceAll in Flags) then
        begin
          while SourcePtr^ <> #0 do
          begin
            ResultPtr^ := SourcePtr^;
            Inc(ResultPtr);
            Inc(SourcePtr);
          end;
          Break;
        end;
      end
      else
      begin
        ResultPtr^ := SourcePtr^;
        Inc(ResultPtr);
        Inc(SourcePtr);
      end;
    end;
  end;

  ResultPtr^ := #0;
  S := ResultStr;
  SetLength(S, StrLen(PChar(S)));
end;

procedure StrReplaceCI(var S: string; Search, Replace: string; Flags: TReplaceFlags);
var
  ResultStr: string;
  SourcePtr: PChar;
  SourceMatchPtr: PChar;
  SearchMatchPtr: PChar;
  ResultPtr: PChar;
  SearchLength,
  ReplaceLength,
  ResultLength: Integer;
  C: Char;
begin
  Search := AnsiUpperCase(Search);
  SearchLength := Length(Search);
  ReplaceLength := Length(Replace);

  if Length(Search) >= ReplaceLength then
    ResultLength := Length(S)
  else
    ResultLength := ((Length(S) div Length(Search)) + 1) * Length(Replace);
  SetLength(ResultStr, ResultLength);

  ResultPtr := PChar(ResultStr);
  SourcePtr := PChar(S);
  C := Search[1];

  while True do
  begin
    while (UpCase(SourcePtr^) <> C) and (SourcePtr^ <> #0) do
    begin
      ResultPtr^ := SourcePtr^;
      Inc(ResultPtr);
      Inc(SourcePtr);
    end;

    if SourcePtr^ = #0 then
      Break
    else
    begin
      SourceMatchPtr := SourcePtr + 1;
      SearchMatchPtr := PChar(Search) + 1;
      while (UpCase(SourceMatchPtr^) = SearchMatchPtr^) and (SearchMatchPtr^ <> #0) do
      begin
        Inc(SourceMatchPtr);
        Inc(SearchMatchPtr);
      end;

      if SearchMatchPtr^ = #0 then
      begin
        Move((@Replace[1])^, ResultPtr^, ReplaceLength);
        Inc(SourcePtr, SearchLength);
        Inc(ResultPtr, ReplaceLength);

        if not (rfReplaceAll in Flags) then
        begin
          while SourcePtr^ <> #0 do
          begin
            ResultPtr^ := SourcePtr^;
            Inc(ResultPtr);
            Inc(SourcePtr);
          end;
          Break;
        end;
      end
      else
      begin
        ResultPtr^ := SourcePtr^;
        Inc(ResultPtr);
        Inc(SourcePtr);
      end;
    end;
  end;

  ResultPtr^ := #0;
  S := ResultStr;
  SetLength(S, StrLen(PChar(S)));
end;

procedure StrReplace(var S: string; const Search, Replace: string;
  Flags: TReplaceFlags);
begin
  if (S <> '') and (Search <> '') then
  begin
    if rfIgnoreCase in Flags then
      StrReplaceCI(S, Search, Replace, Flags)
    else
      StrReplaceCS(S, Search, Replace, Flags);
  end;
end;

{$IFDEF QI_UNICODE}
procedure WideStrReplaceCI(var S: WideString; Search, Replace: WideString; Flags: TReplaceFlags);
var
  ResultStr: WideString;
  SourcePtr: PWideChar;
  SourceMatchPtr: PWideChar;
  SearchMatchPtr: PWideChar;
  ResultPtr: PWideChar;
  SearchLength,
  ReplaceLength,
  ResultLength: Integer;
  C: WideChar;
  LowerCaseFlag: Boolean;
begin
  Search := QIUpperCase(Search);
  SearchLength := Length(Search);
  ReplaceLength := Length(Replace);

  if Length(Search) >= ReplaceLength then
    ResultLength := Length(S)
  else
    ResultLength := ((Length(S) div Length(Search)) + 1) * Length(Replace);
  SetLength(ResultStr, ResultLength);

  ResultPtr := PWideChar(ResultStr);
  SourcePtr := PWideChar(S);
  C := Search[1];

  while True do
  begin
    LowerCaseFlag := False;
    if IsCharLowerW(SourcePtr^) then
    begin
      CharUpperBuffW(SourcePtr, 1);
      LowerCaseFlag := True;
    end;
    while (SourcePtr^ <> C) and (SourcePtr^ <> #0) do
    begin
      if LowerCaseFlag then
        CharLowerBuffW(SourcePtr, 1);

      ResultPtr^ := SourcePtr^;
      Inc(ResultPtr);
      Inc(SourcePtr);

      LowerCaseFlag := False;
      if IsCharLowerW(SourcePtr^) then
      begin
        CharUpperBuffW(SourcePtr, 1);
        LowerCaseFlag := True;
      end;
    end;
    if LowerCaseFlag then
      CharLowerBuffW(SourcePtr, 1);

    if SourcePtr^ = #0 then
      Break
    else
    begin
      SourceMatchPtr := SourcePtr + 1;
      SearchMatchPtr := PWideChar(Search) + 1;

      LowerCaseFlag := False;
      if IsCharLowerW(SourceMatchPtr^) then
      begin
        CharUpperBuffW(SourceMatchPtr, 1);
        LowerCaseFlag := True;
      end;
      while (SourceMatchPtr^ = SearchMatchPtr^) and (SearchMatchPtr^ <> #0) do
      begin
        if LowerCaseFlag then
          CharLowerBuffW(SourceMatchPtr, 1);

        Inc(SourceMatchPtr);
        Inc(SearchMatchPtr);

        LowerCaseFlag := False;
        if IsCharLowerW(SourceMatchPtr^) then
        begin
          CharUpperBuffW(SourceMatchPtr, 1);
          LowerCaseFlag := True;
        end;
      end;
      if LowerCaseFlag then
        CharLowerBuffW(SourceMatchPtr, 1);

      if SearchMatchPtr^ = #0 then
      begin
        Move((@Replace[1])^, ResultPtr^, ReplaceLength * 2);
        Inc(SourcePtr, SearchLength);
        Inc(ResultPtr, ReplaceLength);

        if not (rfReplaceAll in Flags) then
        begin
          while SourcePtr^ <> #0 do
          begin
            ResultPtr^ := SourcePtr^;
            Inc(ResultPtr);
            Inc(SourcePtr);
          end;
          Break;
        end;
      end
      else
      begin
        ResultPtr^ := SourcePtr^;
        Inc(ResultPtr);
        Inc(SourcePtr);
      end;
    end;
  end;

  ResultPtr^ := #0;
  S := ResultStr;
  SetLength(S, Length(S));
end;

procedure WideStrReplaceCS(var S: WideString; const Search, Replace: WideString; Flags: TReplaceFlags);
var
  ResultStr: WideString;
  SourcePtr: PWideChar;
  SourceMatchPtr: PWideChar;
  SearchMatchPtr: PWideChar;
  ResultPtr: PWideChar;
  SearchLength,
  ReplaceLength,
  ResultLength: Integer;
  C: WideChar;
begin
  SearchLength := Length(Search);
  ReplaceLength := Length(Replace);

  if Length(Search) >= ReplaceLength then
    ResultLength := Length(S)
  else
    ResultLength := ((Length(S) div Length(Search)) + 1) * Length(Replace);
  SetLength(ResultStr, ResultLength);

  ResultPtr := PWideChar(ResultStr);
  SourcePtr := PWideChar(S);
  C := Search[1];

  while True do
  begin
    while (SourcePtr^ <> C) and (SourcePtr^ <> #0) do
    begin
      ResultPtr^ := SourcePtr^;
      Inc(ResultPtr);
      Inc(SourcePtr);
    end;

    if SourcePtr^ = #0 then
      Break
    else
    begin

      SourceMatchPtr := SourcePtr + 1;
      SearchMatchPtr := PWideChar(Search) + 1;
      while (SourceMatchPtr^ = SearchMatchPtr^) and (SearchMatchPtr^ <> #0) do
      begin
        Inc(SourceMatchPtr);
        Inc(SearchMatchPtr);
      end;

      if SearchMatchPtr^ = #0 then
      begin
        Move((@Replace[1])^, ResultPtr^, ReplaceLength * 2);
        Inc(SourcePtr, SearchLength);
        Inc(ResultPtr, ReplaceLength);

        if not (rfReplaceAll in Flags) then
        begin
          while SourcePtr^ <> #0 do
          begin
            ResultPtr^ := SourcePtr^;
            Inc(ResultPtr);
            Inc(SourcePtr);
          end;
          Break;
        end;
      end
      else
      begin
        ResultPtr^ := SourcePtr^;
        Inc(ResultPtr);
        Inc(SourcePtr);
      end;
    end;
  end;

  ResultPtr^ := #0;
  S := ResultStr;
  SetLength(S, Length(S));
end;

procedure WideStrReplace(var S: WideString; const Search, Replace: WideString;
  Flags: TReplaceFlags);
begin
  if (S <> '') and (Search <> '') then
  begin
    if rfIgnoreCase in Flags then
      WideStrReplaceCI(S, Search, Replace, Flags)
    else
      WideStrReplaceCS(S, Search, Replace, Flags);
  end;
end;
{$ENDIF}

function QImportLocale: TQImportLocale;
begin
  if Locale = nil then
    Locale := TQImportLocale.Create;
  Result := Locale;
end;

function QImportLoadStr(ID: Integer): string;
begin
  Result := QImportLocale.LoadStr(ID);
end;

{ TQImportFormats }

constructor TQImportFormats.Create;
begin
  inherited;
  FDecimalSeparator := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator;
  FThousandSeparator := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ThousandSeparator;
  FShortDateFormat := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ShortDateFormat;
  FLongDateFormat := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}LongDateFormat;
  FDateSeparator := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}DateSeparator;
  FShortTimeFormat := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ShortTimeFormat;
  FLongTimeFormat := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}LongTimeFormat;
  FTimeSeparator := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}TimeSeparator;
  FBooleanTrue := TStringList.Create;
  FBooleanTrue.Add(QImportLoadStr(QID_BooleanTrue));
  FBooleanFalse := TStringList.Create;
  FBooleanFalse.Add(QImportLoadStr(QID_BooleanFalse));
  FNullValues := TStringList.Create;
  FNullValues.Add(QImportLoadStr(QID_NullValue));
end;

procedure TQImportFormats.DefineProperties(Filer: TFiler);
begin
  inherited DefineProperties(Filer);

  Filer.DefineProperty('ShortDateFormat', ReadShortDateFormat,
    WriteShortDateFormat, True);
  Filer.DefineProperty('LongDateFormat', ReadLongDateFormat,
    WriteLongDateFormat, True);
  Filer.DefineProperty('ShortTimeFormat', ReadShortTimeFormat,
    WriteShortTimeFormat, True);
  Filer.DefineProperty('LongTimeFormat', ReadLongTimeFormat,
    WriteLongTimeFormat, True);
end;

destructor TQImportFormats.Destroy;
begin
  FBooleanTrue.Free;
  FBooleanFalse.Free;
  FNullValues.Free;
  inherited;
end;

procedure TQImportFormats.Assign(Source: TPersistent);
begin
  if Source is TQImportFormats then
  begin
    DecimalSeparator := (Source as TQImportFormats).DecimalSeparator;
    ThousandSeparator := (Source as TQImportFormats).ThousandSeparator;
    ShortDateFormat := (Source as TQImportFormats).ShortDateFormat;
    LongDateFormat := (Source as TQImportFormats).LongDateFormat;
    DateSeparator := (Source as TQImportFormats).DateSeparator;
    ShortTimeFormat := (Source as TQImportFormats).ShortTimeFormat;
    LongTimeFormat := (Source as TQImportFormats).LongTimeFormat;
    TimeSeparator := (Source as TQImportFormats).TimeSeparator;
    BooleanTrue := (Source as TQImportFormats).BooleanTrue;
    BooleanFalse := (Source as TQImportFormats).BooleanFalse;
    NullValues := (Source as TQImportFormats).NullValues;
    Exit;
  end;
  inherited Assign(Source);
end;

procedure TQImportFormats.StoreFormats;
begin
  FOldDecimalSeparator := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator;
  FOldThousandSeparator := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ThousandSeparator;
  FOldShortDateFormat := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ShortDateFormat;
  FOldLongDateFormat := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}LongDateFormat;
  FOldDateSeparator := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}DateSeparator;
  FOldShortTimeFormat := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ShortTimeFormat;
  FOldLongTimeFormat := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}LongTimeFormat;
  FOldTimeSeparator := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}TimeSeparator;
end;

procedure TQImportFormats.RestoreFormats;
begin
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator := FOldDecimalSeparator;
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ThousandSeparator := FOldThousandSeparator;
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ShortDateFormat := FOldShortDateFormat;
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}LongDateFormat := FOldLongDateFormat;
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}DateSeparator := FOldDateSeparator;
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ShortTimeFormat := FOldShortTimeFormat;
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}LongTimeFormat := FOldLongTimeFormat;
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}TimeSeparator := FOldTimeSeparator;
end;

procedure TQImportFormats.ApplyParams;
begin
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator := Char(FDecimalSeparator);
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ThousandSeparator := Char(FThousandSeparator);
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ShortDateFormat := String(FShortDateFormat);
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}LongDateFormat := String(FLongDateFormat);
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}DateSeparator := Char(FDateSeparator);
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}ShortTimeFormat := String(FShortTimeFormat);
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}LongTimeFormat := String(FLongTimeFormat);
  {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.{$IFDEF VCL17}FormatSettings.{$ENDIF}TimeSeparator := Char(FTimeSeparator);
end;

procedure TQImportFormats.ReadLongDateFormat(Reader: TReader);
begin
  FLongDateFormat := Reader.ReadString;
end;

procedure TQImportFormats.ReadLongTimeFormat(Reader: TReader);
begin
  FLongTimeFormat := Reader.ReadString;
end;

procedure TQImportFormats.ReadShortDateFormat(Reader: TReader);
begin
  FShortDateFormat := Reader.ReadString;
end;

procedure TQImportFormats.ReadShortTimeFormat(Reader: TReader);
begin
  FShortTimeFormat := Reader.ReadString;
end;

procedure TQImportFormats.WriteLongDateFormat(Writer: TWriter);
begin
  Writer.WriteString(FLongDateFormat);
end;

procedure TQImportFormats.WriteLongTimeFormat(Writer: TWriter);
begin
  Writer.WriteString(FLongTimeFormat);
end;

procedure TQImportFormats.WriteShortDateFormat(Writer: TWriter);
begin
  Writer.WriteString(FShortDateFormat);
end;

procedure TQImportFormats.WriteShortTimeFormat(Writer: TWriter);
begin
  Writer.WriteString(FShortTimeFormat);
end;

procedure TQImportFormats.SetBooleanTrue(const Value: TStrings);
begin
  FBooleanTrue.Assign(Value);
end;

procedure TQImportFormats.SetBooleanFalse(const Value: TStrings);
begin
  FBooleanFalse.Assign(Value);
end;

procedure TQImportFormats.SetNullValues(const Value: TStrings);
begin
  FNullValues.Assign(Value);
end;

{ TQImport3 }

constructor TQImport3.Create(AOwner: TComponent);
begin
  inherited;
{$IFNDEF NOGUI}
  FGridCaptionRow := -1;
  FGridStartRow := -1;
{$ENDIF}

  FErrors := TqiStringList.Create;
  FMap := TStringList.Create;
  FFormats := TQImportFormats.Create;
  FFieldFormats := TQImportFieldFormats.Create(Self);
  FImportRow := TQImportRow.Create(Self);
  FImportGenerators := TQImportGenerators.Create;

  FImportRecCount := 0;
  FCommitRecCount := 1000;
  FCommitAfterDone := False;
  FSkipFirstRows := 0;
  FSkipFirstCols := 0;
  FImportedRecs := 0;
  FCanceled := false;
  FErrorLog := false;
  FErrorLogFileName:= 'error.log';
  FRewriteErrorLogFile := true;
  FShowErrorLog := false;
//  FSQLLog := false;
//  FSQLLogFileRewrite := true;
//  FSQL := nil;
  FTotalRecCount := 0;
  FAddType := qatInsert;
  FImportDestination := qidDataSet;
  FImportMode := qimInsertAll;
  FKeyColumns := TStringList.Create;
  FMappedColumns := TStringList.Create;

  FIsCSV := False;
  FLastAction := qiaNone;
  FPrevAction := qiaNone;

  FAllowDuplicates := True;

  FCustomImportMode := False;
  FAutoTrimValue := False;
  FImportEmptyRows := True;
end;

destructor TQImport3.Destroy;
begin
  FImportGenerators.Free;
  FImportRow.Free;
  FFieldFormats.Free;
  FFormats.Free;
  FMap.Free;
  FErrors.Free;
  FKeyColumns.Free;
  FMappedColumns.Free;
  inherited;
end;

function TQImport3.Execute: boolean;
begin
  {$IFDEF ADVANCED_DATA_IMPORT_TRIAL_VERSION}
    if not IsIDERuning then
      ShowAboutForm;
  {$ENDIF}
  Result := CheckProperties;

  BeforeImport;
  try
    DoBeginImport;
    DoImport;
  finally
    try
      DoEndImport
    finally
      AfterImport;
    end;
  end;
end;

procedure TQImport3.ImportToCSV(Stream: TStream; Comma, Quote: AnsiChar);
begin
  {$IFDEF ADVANCED_DATA_IMPORT_TRIAL_VERSION}
    if not IsIDERuning then
      ShowAboutForm;
  {$ENDIF}
  FIsCSV := True;
  try
    CheckProperties;
    if not Assigned(Stream) then
      raise Exception.Create(sStreamMustBeAssigned);
    FStream := Stream;
    FComma := Comma;
    FQuote := Quote;
    BeforeImport;
    DoBeginImport;
    DoImport;
  finally
    try
      DoEndImport;
      AfterImport;
    finally
      FIsCSV := False;
    end;
  end;
end;

procedure TQImport3.Cancel;
begin
  FCanceled := true;
end;

procedure TQImport3.LoadConfiguration(const FileName: string);
var
  FIniFile: TIniFile;
begin
  if not FileExists(FileName) then Exit;
  FIniFile := TIniFile.Create(FileName);
  try
    DoLoadConfiguration(FIniFile);
  finally
    FIniFile.Free;
  end;
end;

procedure TQImport3.SaveConfiguration(const FileName: string);
var
  FIniFile: TIniFile;
begin
  if Trim(FileName) = EmptyStr then Exit;
  FIniFile := TIniFile.Create(FileName);
  try
    DoSaveConfiguration(FIniFile);
  finally
    FIniFile.Free;
  end;
end;

procedure TQImport3.DoLoadConfiguration(IniFile: TIniFile);
var
  i, l, k: integer;
  AStrings: TStrings;
  SectionName,
  Str: string;
  FF: TQImportFieldFormat;
  R: TQImportReplacement;
  TextToFind, ReplaceWith,
  TextToFindHex, ReplaceWithHex,
  TextToFindHexRes, ReplaceWithHexRes: string;
  ignoreCase: Boolean;
  SectionNames: TStrings;
begin
  AStrings := TStringList.Create;
  try
    Self.FileName := IniFile.ReadString(QI_BASE, QI_FILE_NAME, Self.FileName);
    AStrings.Clear;
    Map.Clear;
    IniFile.ReadSection(QI_MAP, AStrings);
    for i := 0 to AStrings.Count - 1 do
      Map.Values[AStrings[i]] := IniFile.ReadString(QI_MAP, AStrings[i], EmptyStr);
    Formats.DecimalSeparator := IniFile.ReadString(BASE_FORMATS, BF_DECIMAL_SEPARATOR,
      Formats.DecimalSeparator)[1];
    Formats.ThousandSeparator := IniFile.ReadString(BASE_FORMATS,
      BF_THOUSAND_SEPARATOR, Formats.ThousandSeparator)[1];
    Formats.ShortDateFormat := IniFile.ReadString(BASE_FORMATS, BF_SHORT_DATE_FORMAT,
      Formats.ShortDateFormat);
    Formats.LongDateFormat := IniFile.ReadString(BASE_FORMATS, BF_LONG_DATE_FORMAT,
      Formats.LongDateFormat);
    Formats.ShortTimeFormat := IniFile.ReadString(BASE_FORMATS, BF_SHORT_TIME_FORMAT,
      Formats.ShortTimeFormat);
    Formats.LongTimeFormat := IniFile.ReadString(BASE_FORMATS, BF_LONG_TIME_FORMAT,
      Formats.LongTimeFormat);
    AStrings.Clear;
    Formats.BooleanTrue.Clear;
    IniFile.ReadSection(BOOLEAN_TRUE, AStrings);
    for i := 0 to AStrings.Count - 1 do
      Formats.BooleanTrue.Add(AStrings[i]);
    AStrings.Clear;
    Formats.BooleanFalse.Clear;
    IniFile.ReadSection(BOOLEAN_FALSE, AStrings);
    for i := 0 to AStrings.Count - 1 do
      Formats.BooleanFalse.Add(AStrings[i]);
    AStrings.Clear;
    Formats.NullValues.Clear;
    IniFile.ReadSection(NULL_VALUES, AStrings);
    for i := 0 to AStrings.Count - 1 do
      Formats.NullValues.Add(AStrings[i]);
    AStrings.Clear;
    FieldFormats.Clear;


    SectionNames := TStringList.Create;
    try
      if ImportDestination = qidUserDefined  then
      begin
        IniFile.ReadSections(SectionNames);
        for i := SectionNames.Count - 1 downto 0 do
          if Pos(DATA_FORMATS, SectionNames[i]) = 0 then
            SectionNames.Delete(i);
      end
      else
        for i := 0 to DestinationColCount - 1 do
          SectionNames.Add(DATA_FORMATS + DestinationColName(i));

      for i := 0 to SectionNames.Count - 1 do
      begin
        SectionName := SectionNames[i];
        IniFile.ReadSection(SectionName, AStrings);
        if AStrings.Count = 0 then
          Continue;
        with FieldFormats.Add do
        begin
          if ImportDestination = qidUserDefined  then
            FieldName := Copy(SectionName, Length(DATA_FORMATS) + 1, 
              Length(SectionName) - Length(DATA_FORMATS))
          else
            FieldName := DestinationColName(i);
            
          GeneratorValue := IniFile.ReadInteger(SectionName, DF_GENERATOR_VALUE, 0);
          GeneratorStep := IniFile.ReadInteger(SectionName, DF_GENERATOR_STEP, 0);
          ConstantValue := IniFile.ReadString(SectionName, DF_CONSTANT_VALUE, EmptyStr);
          NullValue := IniFile.ReadString(SectionName, DF_NULL_VALUE, EmptyStr);
          DefaultValue := IniFile.ReadString(SectionName, DF_DEFAULT_VALUE, EmptyStr);
          LeftQuote := IniFile.ReadString(SectionName, DF_LEFT_QUOTE, EmptyStr);
          RightQuote := IniFile.ReadString(SectionName, DF_RIGHT_QUOTE, EmptyStr);
          QuoteAction := TQuoteAction(IniFile.ReadInteger(SectionName, DF_QUOTE_ACTION, 0));
          CharCase := TQImportCharCase(IniFile.ReadInteger(SectionName, DF_CHAR_CASE, 0));
          CharSet := TQImportCharSet(IniFile.ReadInteger(SectionName, DF_CHAR_SET, 0));
          Functions := TQImportFunctions(IniFile.ReadInteger(SectionName, QI_FUNCTION, 0));
        end;
      end;
    finally
      SectionNames.Free;
    end;

    // replacements
    IniFile.ReadSections(AStrings);
    for i := 0 to AStrings.Count - 1 do
    begin
      if Pos(QI_REPLACEMENTS, AStrings[i]) = 1 then
      begin
        l := Length(AStrings[i]);
        while QImport3Common.CharInSet(AStrings[i][l],
        ['0', '1', '2', '3', '4', '5', '6', '7', '8', '9'])
        do Dec(l);
        str := Copy(AStrings[i], Length(QI_REPLACEMENTS) + 1,
               l - Length(QI_REPLACEMENTS) - 1);
        k := FieldFormats.IndexByName(str);
        if k < 0 then
        begin
          FF := FieldFormats.Add;
          FF.FieldName := str;
        end else
          FF := FieldFormats[k];
        TextToFind := IniFile.ReadString(AStrings[i], QI_RP_TEXT_TO_FIND, EmptyStr);
        ReplaceWith := IniFile.ReadString(AStrings[i], QI_RP_REPLACE_WITH, EmptyStr);
        TextToFindHex := IniFile.ReadString(AStrings[i], QI_RP_TEXT_TO_FIND_HEX, EmptyStr);
        ReplaceWithHex := IniFile.ReadString(AStrings[i], QI_RP_REPLACE_WITH_HEX, EmptyStr);
        if QICompareText(TextToFindHex, EmptyStr) <> 0 then
        begin
          if HexToString(TextToFindHex, TextToFindHexRes) then
            if CompareText(TextToFindHexRes, TextToFind) <> 0 then
              TextToFind := TextToFindHexRes;
        end;
        if QICompareText(ReplaceWithHex, EmptyStr) <> 0 then
        begin
          if HexToString(ReplaceWithHex, ReplaceWithHexRes) then
            if CompareText(ReplaceWithHexRes, ReplaceWith) <> 0 then
              ReplaceWith := ReplaceWithHexRes;
        end;
        ignoreCase := IniFile.ReadBool(AStrings[i], QI_RP_IGNORE_CASE, False);
        if (k < 0) or (not FF.Replacements.ItemExists(
            textToFind, replaceWith, ignoreCase)) then
        begin
          R := FF.Replacements.Add;
          R.TextToFind := textToFind;
          R.ReplaceWith := replaceWith;
          R.IgnoreCase := ignoreCase;
        end;
      end;
    end;

    CommitAfterDone := IniFile.ReadBool(QI_BASE, QI_COMMIT_AFTER_DONE,
      CommitAfterDone);
    CommitRecCount := IniFile.ReadInteger(QI_BASE, QI_COMMIT_REC_COUNT,
      CommitRecCount);
    ImportRecCount := IniFile.ReadInteger(QI_BASE, QI_IMPORT_REC_COUNT,
      ImportRecCount);
    ErrorLog := IniFile.ReadBool(QI_BASE, QI_ENABLE_ERROR_LOG, ErrorLog);
    ErrorLogFileName := IniFile.ReadString(QI_BASE, QI_ERROR_LOG_FILE_NAME,
      ErrorLogFileName);
    RewriteErrorLogFile := IniFile.ReadBool(QI_BASE, QI_REWRITE_ERROR_LOG_FILE,
      RewriteErrorLogFile);
    ShowErrorLog := IniFile.ReadBool(QI_BASE, QI_SHOW_ERROR_LOG, ShowErrorLog);
//      SQLLog := ReadBool(QI_BASE, QI_ENABLE_SQL_LOG, SQLLog);
//      SQLLogFileName := ReadString(QI_BASE, QI_SQL_LOG_FILE_NAME,
//        SQLLogFileName);
//      SQLLogFileRewrite := ReadBool(QI_BASE, QI_SQL_LOG_FILE_REWRITE,
//        SQLLogFileRewrite);

    ImportDestination := TQImportDestination(IniFile.ReadInteger(QI_BASE,
      QI_IMPORT_DESTINATION, Integer(qidDataSet)));
    ImportMode := TQImportMode(IniFile.ReadInteger(QI_BASE, QI_IMPORT_MODE,
      Integer(qimInsertAll)));
    AStrings.Clear;
    KeyColumns.Clear;
    IniFile.ReadSection(QI_KEY_COLUMNS, AStrings);
    for i := 0 to AStrings.Count - 1 do
      KeyColumns.Add(AStrings[i]);
    {$IFNDEF NOGUI}
    GridCaptionRow := IniFile.ReadInteger(QI_BASE, QI_GRID_CAPTION_ROW, -1);
    GridStartRow := IniFile.ReadInteger(QI_BASE, QI_GRID_START_ROW, -1);
    {$ENDIF}
    ImportEmptyRows := IniFile.ReadBool(QI_BASE, QI_IMPORT_EMPTY_ROWS, True);
    AutoTrimValue := IniFile.ReadBool(QI_BASE, QI_AUTO_TRIM_VALUE, False);
  finally
    AStrings.Free;
  end;
end;

procedure TQImport3.DoSaveConfiguration(IniFile: TIniFile);
var
  i,j : integer;
  str: string;
begin
  with IniFile do
  begin
    WriteString(QI_BASE, QI_FILE_NAME, Self.FileName);
    EraseSection(QI_MAP);
    for i := 0 to Map.Count - 1 do
      WriteString(QI_MAP, Map.Names[i], Map.Values[Map.Names[i]]);
    WriteString(BASE_FORMATS, BF_DECIMAL_SEPARATOR, Formats.DecimalSeparator);
    WriteString(BASE_FORMATS, BF_THOUSAND_SEPARATOR, Formats.ThousandSeparator);
    WriteString(BASE_FORMATS, BF_SHORT_DATE_FORMAT, Formats.ShortDateFormat);
    WriteString(BASE_FORMATS, BF_LONG_DATE_FORMAT, Formats.LongDateFormat);
    WriteString(BASE_FORMATS, BF_SHORT_TIME_FORMAT, Formats.ShortTimeFormat);
    WriteString(BASE_FORMATS, BF_LONG_TIME_FORMAT, Formats.LongTimeFormat);
    EraseSection(BOOLEAN_TRUE);
    for i := 0 to Formats.BooleanTrue.Count - 1 do
      WriteString(BOOLEAN_TRUE, Formats.BooleanTrue[i], EmptyStr);
    EraseSection(BOOLEAN_FALSE);
    for i := 0 to Formats.BooleanFalse.Count - 1 do
      WriteString(BOOLEAN_FALSE, Formats.BooleanFalse[i], EmptyStr);
    EraseSection(NULL_VALUES);
    for i := 0 to Formats.NullValues.Count - 1 do
      WriteString(NULL_VALUES, Formats.NullValues[i], EmptyStr);
    for i := 0 to FieldFormats.Count - 1 do begin
      str := DATA_FORMATS + FieldFormats[i].FieldName;
      if Trim(str) = EmptyStr then Continue;
      WriteString(str, DF_GENERATOR_VALUE, FieldFormats[i].FieldName);
      WriteInteger(str, DF_GENERATOR_VALUE, FieldFormats[i].GeneratorValue);
      WriteInteger(str, DF_GENERATOR_STEP, FieldFormats[i].GeneratorStep);
      WriteString(str, DF_CONSTANT_VALUE, FieldFormats[i].ConstantValue);
      WriteString(str, DF_NULL_VALUE, FieldFormats[i].NullValue);
      WriteString(str, DF_DEFAULT_VALUE, FieldFormats[i].DefaultValue);
      WriteString(str, DF_LEFT_QUOTE, FieldFormats[i].LeftQuote);
      WriteString(str, DF_RIGHT_QUOTE, FieldFormats[i].RightQuote);
      WriteInteger(str, DF_QUOTE_ACTION, Integer(FieldFormats[i].QuoteAction));
      WriteInteger(str, DF_CHAR_CASE, Integer(FieldFormats[i].CharCase));
      WriteInteger(str, DF_CHAR_SET, Integer(FieldFormats[i].CharSet));
      WriteInteger(str, QI_FUNCTION, Integer(FieldFormats[i].Functions));

      // Replacements
      str := QI_REPLACEMENTS + AnsiUpperCase(FieldFormats[i].FieldName);
      for j := 0 to FieldFormats[i].Replacements.Count - 1 do
      begin
        WriteString(Format('%s_%d', [str, j]), QI_RP_TEXT_TO_FIND,
          FieldFormats[i].Replacements[j].TextToFind);
        WriteString(Format('%s_%d', [str, j]), QI_RP_REPLACE_WITH,
          FieldFormats[i].Replacements[j].ReplaceWith);
        WriteString(Format('%s_%d', [str, j]), QI_RP_TEXT_TO_FIND_HEX,
          StringToHex(FieldFormats[i].Replacements[j].TextToFind));
        WriteString(Format('%s_%d', [str, j]), QI_RP_REPLACE_WITH_HEX,
          StringToHex(FieldFormats[i].Replacements[j].ReplaceWith));
        WriteBool(Format('%s_%d', [str, j]), QI_RP_IGNORE_CASE,
          FieldFormats[i].Replacements[j].IgnoreCase);
      end;
    end;
    WriteBool(QI_BASE, QI_COMMIT_AFTER_DONE, CommitAfterDone);
    WriteInteger(QI_BASE, QI_COMMIT_REC_COUNT, CommitRecCount);
    WriteInteger(QI_BASE, QI_IMPORT_REC_COUNT, ImportRecCount);
    WriteBool(QI_BASE, QI_ENABLE_ERROR_LOG, ErrorLog);
    WriteString(QI_BASE, QI_ERROR_LOG_FILE_NAME, ErrorLogFileName);
    WriteBool(QI_BASE, QI_REWRITE_ERROR_LOG_FILE, RewriteErrorLogFile);
    WriteBool(QI_BASE, QI_SHOW_ERROR_LOG, ShowErrorLog);
//    WriteBool(QI_BASE, QI_ENABLE_SQL_LOG, SQLLog);
//    WriteString(QI_BASE, QI_SQL_LOG_FILE_NAME, SQLLogFileName);
//    WriteBool(QI_BASE, QI_SQL_LOG_FILE_REWRITE, SQLLogFileRewrite);

    WriteInteger(QI_BASE, QI_IMPORT_DESTINATION, Integer(ImportDestination));
    WriteInteger(QI_BASE, QI_IMPORT_MODE, Integer(ImportMode));
    EraseSection(QI_KEY_COLUMNS);
    for i := 0 to KeyColumns.Count - 1 do
      WriteString(QI_KEY_COLUMNS, KeyColumns[i], EmptyStr);
    {$IFNDEF NOGUI}
    WriteInteger(QI_BASE, QI_GRID_CAPTION_ROW, GridCaptionRow);
    WriteInteger(QI_BASE, QI_GRID_START_ROW, GridStartRow);
    {$ENDIF}
    WriteBool(QI_BASE, QI_IMPORT_EMPTY_ROWS, ImportEmptyRows);
    WriteBool(QI_BASE, QI_AUTO_TRIM_VALUE, AutoTrimValue);
  end;
end;

procedure TQImport3.DestinationInsert;
{$IFNDEF NOGUI}
var
  i: Integer;
{$ENDIF}
begin
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet:
      case AddType of
        qatAppend: DataSet.Append;
        qatInsert: DataSet.Insert;
      end;
    {$IFNDEF NOGUI}
    qidDbGrid:
      case AddType of
        qatAppend: DBGrid.DataSource.DataSet.Append;
        qatInsert: DBGrid.DataSource.DataSet.Insert;
      end;
    qidListView: begin
      FCurrListItem := ListView.Items.Add;
      for i := 1 to ListView.Columns.Count - 1 do
        FCurrListItem.SubItems.Add(EmptyStr);
    end;
    qidStringGrid: begin
      if ImportMode = qimInsertAll then
      begin
          FAddPos := FCurrStrGrRow;
          if (FCurrStrGrRow = StringGrid.RowCount) then
            StringGrid.RowCount := StringGrid.RowCount + 1;
      end else begin
        if ((FPrevAction = qiaNone) and (FCurrStrGrRow = FGridStartRow))
        and (StringGrid.RowCount = FGridStartRow+1) then
          FAddPos := FCurrStrGrRow
        else begin
          FAddPos := StringGrid.RowCount;
          StringGrid.RowCount := StringGrid.RowCount + 1;
        end;
      end;
    end;
    {$ENDIF}
  end;
  FLastAction := qiaInsert;
end;

procedure TQImport3.DestinationEdit;
begin
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: DataSet.Edit;
    {$IFNDEF NOGUI}
    qidDbGrid: DBGrid.DataSource.DataSet.Edit;
    {$ENDIF}
  end;
  FLastAction := qiaUpdate;
end;

procedure TQImport3.DestinationDelete;
{$IFNDEF NOGUI}
var
  i, k: integer;
{$ENDIF}
begin
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: DataSet.Delete;
    {$IFNDEF NOGUI}
    qidDbGrid: DBGrid.DataSource.DataSet.Delete;
    qidListView:
      if Assigned(FCurrListItem) then FCurrListItem.Delete;
    qidStringGrid: begin
      //dee
      k := FFindPos;
      if (k > -1) and (k < StringGrid.RowCount) then
      begin
        for i := k to StringGrid.RowCount - 2 do
        begin
          StringGrid.Rows[i].Assign(StringGrid.Rows[i+1]);
        end;
        StringGrid.Rows[StringGrid.RowCount-1].Clear;
//        if (GridCaptionRow > -1) and (StringGrid.RowCount-1 <> GridCaptionRow) then
        StringGrid.RowCount:= StringGrid.RowCount - 1;
      end;
      //\dee
    end;
    {$ENDIF}
  end;
  FLastAction := qiaDelete;
end;

procedure TQImport3.DestinationSetValues;
var
  i{$IFNDEF NOGUI}, k{$ENDIF}: Integer;
  wstr: WideString;
  str: AnsiString;
  errorMsg: string;
begin
  {$IFNDEF NOGUI}
  k := -2;
  {$ENDIF}
  for i := 0 to FImportRow.Count - 1 do
  begin
    if FIsCSV then
    begin
      wstr := FImportRow[i].Value;

      if QIPos(',', wstr) > 0 then
        wstr := QIStringReplace(wstr, ',', '.', [rfReplaceAll]);

      if FQuote <> #0 then
        wstr := QIQuotedStr(wstr, qiChar(FQuote));

      if i = FImportRow.Count - 1 then
        wstr := wstr + LF
      else
        wstr := wstr + qiChar(FComma);

      if Assigned(FOnWideStringToCharset) then
        FOnWideStringToCharset(Self, wstr, str)
      else
        str := AnsiString(wstr);

      FStream.Write(str[1], Length(str));
      Continue;
    end;

    case ImportDestination of
      qidDataSet:
        begin
          errorMsg := StringToField(DataSet.Fields[FImportRow[i].FColumnIndex],
            FImportRow[i].Value, FImportRow[i].IsBinary);
          if errorMsg <> EmptyStr then
            WriteErrorLog(errorMsg);
        end;
{$IFNDEF NOGUI}
      qidDBGrid:
        begin
          errorMsg := StringToField(DBGrid.Columns[FImportRow[i].FColumnIndex].Field,
            FImportRow[i].Value, FImportRow[i].IsBinary);
          if errorMsg <> EmptyStr then
            WriteErrorLog(errorMsg);
        end;
      qidListView:
        if Assigned(FCurrListItem) then
        begin
          if FImportRow[i].FColumnIndex = 0 then
            FCurrListItem.Caption := FImportRow[i].Value
          else FCurrListItem.SubItems[FImportRow[i].FColumnIndex - 1] :=
            FImportRow[i].Value;
        end;
      qidStringGrid: begin
        //dee
        if k = -2 then
        begin
          if (FLastAction = qiaUpdate) then
            k := FFindPos
          else if (FLastAction = qiaInsert) then
            k := FAddPos;
        end;
        if (k > -1) and (k < StringGrid.RowCount)then
          StringGrid.Cells[FImportRow[i].FColumnIndex, k] :=
            FImportRow[i].Value;
        //\dee
      end;
{$ENDIF}
    end;
  end;
end;

procedure TQImport3.DestinationPost;
begin
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet:
      begin
        if DataSet.State in [dsInsert, dsEdit] then
          DataSet.Post;
      end;
{$IFNDEF NOGUI}
    qidDbGrid:
      if DBGrid.DataSource.DataSet.State in [dsInsert, dsEdit] then
        DBGrid.DataSource.DataSet.Post;
    qidStringGrid:
      Inc(FCurrStrGrRow);
{$ENDIF}
  end;
end;

procedure TQImport3.DestinationCancel;
begin
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: if DataSet.State in [dsInsert, dsEdit] then DataSet.Cancel;
    {$IFNDEF NOGUI}
    qidDbGrid: if DBGrid.DataSource.DataSet.State in [dsInsert, dsEdit] then
      DBGrid.DataSource.DataSet.Cancel;
    {$ENDIF}
  end;
end;

function TQImport3.DestinationFindColumn(const ColName: string): integer;
var
  Field: TField;
{$IFNDEF NOGUI}
  i: integer;
{$ENDIF}
begin
  Result := -1;
  //if FIsCSV then Exit;
  case ImportDestination of
    qidDataSet: begin
      Field := DataSet.FindField(ColName);
      if Assigned(Field) then
        Result := Field.Index;
    end;
    {$IFNDEF NOGUI}
    qidDBGrid:
      for i := 0 to DBGrid.Columns.Count - 1 do
        if AnsiCompareText(DBGrid.Columns[i].Title.Caption, ColName) = 0 then
        begin
          Result := i;
          Exit;
        end;
    qidListView:
      for i := 0 to ListView.Columns.Count - 1 do
        if AnsiCompareText(ListView.Columns[i].Caption, ColName) = 0 then
        begin
          Result := i;
          Exit;
        end;
    qidStringGrid: begin
      i := StrToIntDef(ColName, -1);
      if i > -1 then
      begin
        if  i < StringGrid.ColCount then Result := i;
      end
      else begin
        if GridCaptionRow > -1 then
          for i := 0 to StringGrid.ColCount - 1 do
            if AnsiCompareStr(StringGrid.Cells[i, GridCaptionRow], ColName) = 0 then
            begin
              Result := i;
              Exit;
            end;
      end;
    end;
    {$ENDIF}
  end;
end;

procedure TQImport3.DestinationDisableControls;
begin
  if FIsCSV then Exit;
  case ImportDestination of
    qidDataSet: DataSet.DisableControls;
    {$IFNDEF NOGUI}
    qidDBGrid: DBGrid.DataSource.DataSet.DisableControls;
    qidListView: ListView.Items.BeginUpdate;
    {$ENDIF}
  end;
end;

procedure TQImport3.DestinationEnableControls;
begin
  if FIsCSV then Exit;
  case ImportDestination of
    qidDataSet: DataSet.EnableControls;
    {$IFNDEF NOGUI}
    qidDBGrid: DBGrid.DataSource.DataSet.EnableControls;
    qidListView: ListView.Items.EndUpdate;
    {$ENDIF}
  end;
end;

procedure TQImport3.CheckDestination;
begin
  QImportCheckDestination(IsCSV, ImportDestination, DataSet
    {$IFNDEF NOGUI}, DBGrid, ListView,
 StringGrid{$ENDIF});
end;

function TQImport3.DestinationFindByKey: boolean;
var
  Keys{$IFNDEF NOGUI}, strValue{$ENDIF}: string;
  Values: Variant;
  i{$IFNDEF NOGUI}, j{$ENDIF}: Integer;
  Col: TQImportCol;
  {$IFNDEF NOGUI}Flag: Boolean; {$ENDIF}
begin
  Result := false;
  if FIsCSV then Exit;
  case ImportDestination of
    qidDataSet:
      if KeyColumns.Count > 0 then
      begin
        Keys := EmptyStr;
        if KeyColumns.Count > 1 then
          Values := VarArrayCreate([0, KeyColumns.Count - 1], varVariant);
        for i := 0 to KeyColumns.Count - 1 do
        begin
          Col := FImportRow.ColByName(KeyColumns[i]);
          if Assigned(Col) then
          begin
            Keys := Keys + Col.Name;
            if KeyColumns.Count > 1 then
              Values[i] := Col.Value
            else
              Values := Col.Value;
          end;
          if i < KeyColumns.Count - 1 then
            Keys := Keys + ';';
        end;
        if Assigned(FOnDestinationLocate) then
          FOnDestinationLocate(Self, KeyColumns, FImportRow, Keys, Values);
        try
          Result := DataSet.Locate(Keys, Values, [loCaseInsensitive])
        except
          Result := False;
        end;
      end;
    {$IFNDEF NOGUI}
    qidDBGrid:
      if KeyColumns.Count > 0 then
      begin
        Keys := EmptyStr;
        if KeyColumns.Count > 1 then
          Values := VarArrayCreate([0, KeyColumns.Count - 1], varVariant);
        for i := 0 to KeyColumns.Count - 1 do
        begin
          Col := FImportRow.ColByName(KeyColumns[i]);
          if Assigned(Col) then
          begin
            Keys := Keys + DBGrid.Columns[Col.FColumnIndex].Field.FieldName;
            if KeyColumns.Count > 1 then
              Values[i] := Col.Value
            else
              Values := Col.Value;
          end;
          if i < KeyColumns.Count - 1 then Keys := Keys + ';';
        end;
        if Assigned(FOnDestinationLocate) then
          FOnDestinationLocate(Self, KeyColumns, FImportRow, Keys, Values);
        Result := DBGrid.DataSource.DataSet.Locate(Keys, Values, [loCaseInsensitive])
      end;
    qidListView: begin
      for i := 0 to ListView.Items.Count - 1 do
      begin
        Flag := true;
        for j := 0 to KeyColumns.Count - 1 do
        begin
          Col := FImportRow.ColByName(KeyColumns[j]);
          if Assigned(Col) then
          begin
            strValue := Col.Value;
            if Col.FColumnIndex = 0 then
              Flag := AnsiCompareText(strValue, ListView.Items[i].Caption) = 0
            else
              Flag := AnsiCompareText(strValue, ListView.Items[i].SubItems[Col.FColumnIndex - 1]) = 0;
              if not Flag then
                Break;
          end
          else
            Exit;
        end;
        if Flag then
        begin
          Result := true;
          FCurrListItem := ListView.Items[i];
          Exit;
        end;
      end;
    end;
    qidStringGrid: begin
      FFindPos := -1;
      for i := 0 to StringGrid.RowCount - 1 do
      begin
        if (GridCaptionRow > -1) and (i = GridCaptionRow) then
          Continue;
        Flag := true;
        for j := 0 to KeyColumns.Count - 1 do
        begin
          Col := FImportRow.ColByName(KeyColumns[j]);
          if Assigned(Col) then
          begin
            strValue := Col.Value;
            Flag := AnsiCompareText(strValue, StringGrid.Cells[Col.FColumnIndex, i]) = 0;
          if not Flag then
            Break;
          end
          else
            Exit;
        end;
        if Flag then
        begin
          Result := true;
          // dee
          FFindPos := i;
          Exit;
        end;
      end;
    end;
    {$ENDIF}
  end;
end;

function TQImport3.DestinationFindByFields: Boolean;
var
  i{$IFNDEF NOGUI}, j{$ENDIF}: Integer;
  fieldNames{$IFNDEF NOGUI}, strValue{$ENDIF}: string;
  values: Variant;
  col: TQImportCol;
  {$IFNDEF NOGUI}flag: Boolean; {$ENDIF}
begin
  Result := False;

  if FAllowDuplicates then Exit;

  if FIsCSV then Exit;
  case ImportDestination of
    qidDataSet:
      if FMappedColumns.Count > 0 then
      begin
        fieldNames := EmptyStr;
        if FMappedColumns.Count > 1 then
          values := VarArrayCreate([0, FMappedColumns.Count - 1], varVariant);
        for i := 0 to FMappedColumns.Count - 1 do
        begin
          col := FImportRow.ColByName(FMappedColumns[i]);
          if Assigned(col) then
          begin
            fieldNames := fieldNames + col.Name;
            if FMappedColumns.Count > 1 then
              Values[i] := col.Value
            else
              Values := col.Value;
          end;
          if i < FMappedColumns.Count - 1 then
            fieldNames := fieldNames + ';';
        end;
        try
          Result := DataSet.Locate(fieldNames, values, [])
        except
          Result := False;
        end;
      end;
    {$IFNDEF NOGUI}
    qidDBGrid:
      if FMappedColumns.Count > 0 then
      begin
        fieldNames := EmptyStr;
        if FMappedColumns.Count > 1 then
          values := VarArrayCreate([0, FMappedColumns.Count - 1], varVariant);
        for i := 0 to FMappedColumns.Count - 1 do
        begin
          col := FImportRow.ColByName(FMappedColumns[i]);
          if Assigned(col) then
          begin
            fieldNames := fieldNames + DBGrid.Columns[col.FColumnIndex].Field.FieldName;
            if FMappedColumns.Count > 1 then
              Values[i] := col.Value
            else
              Values := col.Value;
          end;
          if i < FMappedColumns.Count - 1 then
            fieldNames := fieldNames + ';';
        end;
        Result := DBGrid.DataSource.DataSet.Locate(fieldNames, values, []);
      end;
    qidListView:
      for i := 0 to ListView.Items.Count - 1 do
      begin
        flag := False;
        for j := 0 to FMappedColumns.Count - 1 do
        begin
          col := FImportRow.ColByName(FMappedColumns[j]);
          if Assigned(col) then
          begin
            strValue := Col.Value;
            if col.FColumnIndex = 0 then
              flag := AnsiCompareStr(strValue, ListView.Items[i].Caption) = 0
            else
              flag := AnsiCompareStr(strValue, ListView.Items[i].SubItems[col.FColumnIndex - 1]) = 0;
            if not flag then
              Break;
          end
          else
            Exit;
        end;
        if flag then
        begin
          Result := True;
          FCurrListItem := ListView.Items[i];
          Exit;
        end;
      end;
    qidStringGrid: begin

      FFindPos := -1;
      for i := 0 to StringGrid.RowCount - 1 do
      begin
        if (GridCaptionRow > -1) and (i = GridCaptionRow) then
          Continue;
        Flag := False;
        for j := 0 to FMappedColumns.Count - 1 do
        begin
          col := FImportRow.ColByName(FMappedColumns[j]);
          if Assigned(col) then
          begin
            strValue := Col.Value;
            flag := AnsiCompareStr(strValue, StringGrid.Cells[col.FColumnIndex, i]) = 0
          end
          else
            Exit;
          if not flag then
            Break;
        end;
        if flag then
        begin
          Result := True;
          FFindPos := i;
          Exit;
        end;
      end;
    end;
    {$ENDIF}
  end;
end;

function TQImport3.DestinationColCount: integer;
begin
  Result := QImportDestinationColCount(IsCSV, ImportDestination, DataSet
    {$IFNDEF NOGUI}, DBGrid, ListView, StringGrid{$ENDIF});
end;

function TQImport3.DestinationColName(Index: integer): string;
begin
  Result := QImportDestinationColName(IsCSV, ImportDestination, DataSet,
    {$IFNDEF NOGUI}DBGrid, ListView, StringGrid, GridCaptionRow,{$ENDIF} Index);
end;

procedure TQImport3.DoBeginImport;
begin
  if Assigned(FOnBeforeImport) then
    FOnBeforeImport(Self);
end;

function TQImport3.DoBeforePost: Boolean;
begin
  Result := true;
  if Assigned(FOnBeforePost) then
    FOnBeforePost(Self, FImportRow, Result);
end;

procedure TQImport3.DoAfterPost;
begin
  if Assigned(FOnAfterPost) then
    FOnAfterPost(Self, FImportRow);
end;

procedure TQImport3.DoImportRecord;
begin
  Inc(FImportedRecs);
  if Assigned(FOnImportRecord) then
    FOnImportRecord(Self);
end;

procedure TQImport3.DoImportRowComplete;
begin
  if Assigned(FOnImportRowComplete) then
    FOnImportRowComplete(Self, FImportRow);
end;

procedure TQImport3.DoImportError(Error: Exception);
begin
  FCustomImportError := True;
  {$IFDEF QI_UNICODE}
  if Error is WideException then
    LastError := WideFormatStr(QImportLoadStr(QIW_ImportErrorFormat),
                  [FormatDateTime(Formats.ShortDateFormat + ' ' +
                   Formats.ShortTimeFormat, Now), FCurrentLineNumber, WideException(Error).Message])
  else
  {$ENDIF}
  LastError := string (Format(QImportLoadStr(QIW_ImportErrorFormat),
                [FormatDateTime(Formats.ShortDateFormat + ' ' +
                 Formats.ShortTimeFormat, Now), FCurrentLineNumber, Error.Message]));

  if FErrorLog then
  begin
    FErrors.Add(LastError);
    if Assigned(FErrorLogFS) then
    begin
      lastError := lastError + LF;
      FErrorLogFS.Write(AnsiString(lastError)[1], Length(lastError)*SizeOf(AnsiChar));
    end;
  end;

  Inc(FCurrentLineNumber);

  if Assigned(FOnImportError) then
    FOnImportError(Self);
end;

procedure TQImport3.WriteErrorLog(const ErrorMsg: string);
begin
  LastError := string( Format(QImportLoadStr(QIW_ImportErrorFormat),
                  [FormatDateTime(Formats.ShortDateFormat + ' ' +
                   Formats.ShortTimeFormat, Now), FCurrentLineNumber, ErrorMsg]));
  if FErrorLog then
  begin
    if Assigned(FErrorLogFS) then
    begin
      lastError := lastError + ansiString(LF);
      FErrorLogFS.Write(AnsiString(lastError)[1], Length(lastError)*SizeOf(AnsiChar));
    end;
  end;

  if Assigned(FOnImportErrorAdv) then
    FOnImportErrorAdv(Self);
end;

procedure TQImport3.DoNeedCommit;
begin
  if Assigned(FOnNeedCommit) then
    FOnNeedCommit(Self);
end;

procedure TQImport3.DoEndImport;
begin
  if Assigned(FOnAfterImport) then FOnAfterImport(Self);
end;

procedure TQImport3.Notification(AComponent: TComponent; Operation: TOperation);
begin
  inherited;
  if (AComponent = FDataSet) and (Operation = opRemove)
    then FDataSet := nil;
  {$IFNDEF NOGUI}
  if (AComponent = FDBGrid) and (Operation = opRemove)
    then FDBGrid := nil;
  if (AComponent = FListView) and (Operation = opRemove)
    then FListView := nil;
  if (AComponent = FStringGrid) and (Operation = opRemove)
    then FStringGrid := nil;
  {$ENDIF}
  {$IFDEF USESCRIPT}
  if (AComponent = FScriptEngine) and (Operation = opRemove) then
    FScriptEngine := nil;
  {$ENDIF}
end;

function TQImport3.PrepareImportValue(const AValue: qiString;
  const AIsBinary: Boolean; const AFieldType: TFieldType): Variant;
var
  ImportFieldType: TQImportFieldType;
begin
  ImportFieldType := FieldTypeToImportFieldType(AFieldType);
  if StrIsNull(Trim(AValue)) or ((not (ImportFieldType in [iftString])) and
      (Trim(AValue) = EmptyStr)) then
  begin
    Result := Null;
    Exit;
  end;

  case ImportFieldType of
    iftString:
      Result := GetStringValue(AValue, AFieldType);
    iftBytes:
      Result := GetBitesValue(AValue);
    iftInteger:
      Result := GetIntegerValue(AValue, AFieldType);
    iftBoolean:
      Result := StringToBoolean(AValue);
    iftDouble:
      Result := GetFloatValue(AValue, AFieldType);
    iftCurrency:
      Result := StrToCurr(GetNumericString(AValue));
    iftDateTime:
      Result := GetDateTimeValue(AValue);
    else
      raise Exception.Create(QImportLoadStr(QIE_UnknownFieldType));
  end;
end;

procedure TQImport3.InitializeImportRow;

  function IsKeyColumn(const AColumnName: string): Boolean;
  var
    i: Integer;
  begin
    Result := False;
    for i := 0 to KeyColumns.Count - 1 do
    begin
      if CompareStr(KeyColumns[i], AColumnName) = 0 then
      begin
        Result := True;
        Break;
      end;
    end;
  end;

var
  i, k: integer;
  Col: TQImportCol;
begin
  FImportRow.Clear;
  FImportRow.MapNameIdxHash.TableSize := FMap.Count;
  FMappedColumns.Clear;

  for i := 0 to FMap.Count - 1 do
  begin
    k := DestinationFindColumn(FMap.Names[i]);
    if (k = -1) and (FImportDestination <> qidUserDefined) then
      raise EQImportError.CreateFmt(QImportLoadStr(QIE_FieldNotFound), [FMap.Names[i]]);
    with FImportRow.Add(FMap.Names[i]) do
      FColumnIndex := k;
    FImportRow.MapNameIdxHash.Insert(FMap.Names[i], Pointer(i));
    if not IsKeyColumn(FMap.Names[i]) then
      FMappedColumns.Add(FMap.Names[i]);
  end;

  for i := 0 to FFieldFormats.Count - 1 do
  begin
    Col := FImportRow.ColByName(FFieldFormats.Items[i].FieldName);
    if not Assigned(Col) and ((FFieldFormats[i].GeneratorStep <> 0) or
      (FFieldFormats[i].ConstantValue <> '')) then
    begin
      k := DestinationFindColumn(FFieldFormats.Items[i].FieldName);
      if k > -1 then
      begin
        Col := FImportRow.Add(FFieldFormats.Items[i].FieldName);
        Col.FColumnIndex := k;
      end;
    end;

    if Assigned(Col) and (FFieldFormats[i].GeneratorStep <> 0) then
      FImportGenerators.Add(FFieldFormats.Items[i].FieldName,
        FFieldFormats[i].GeneratorValue, FFieldFormats[i].GeneratorStep);
  end;
end;

procedure TQImport3.BeforeImport;
var
  ErrorLogMode: word;
  str: AnsiString;
  path, ELFN: string;
{var
  SQLLogMode: word;}
begin
  if (FImportRecCount > 0)  and (ImportRecCount < FTotalRecCount)
    then FTotalRecCount := FImportRecCount;

  FLastAction := qiaNone;
  FPrevAction := qiaNone;
  FImportedRecs := 0;
  FErrors.Clear;
  FCanceled := false;
  FFormats.StoreFormats;
  FFormats.ApplyParams;
  DestinationDisableControls;
  FCurrentLineNumber := 1;

  {$IFNDEF NOGUI}
  FCurrListItem := nil;

  FCurrStrGrRow := FGridStartRow;
  if (ImportDestination = qidStringGrid) and Assigned(FStringGrid) then
  begin
    if FCurrStrGrRow = -1 then
      FCurrStrGrRow := FStringGrid.RowCount;
    if FStringGrid.RowCount <= FGridStartRow then
      FStringGrid.RowCount := FGridStartRow + 1;
  end;
  {$ENDIF}

  InitializeImportRow;

  if FErrorLog and (FErrorLogFileName <> EmptyStr) then
  begin
    if FRewriteErrorLogFile then
      ErrorLogMode := fmCreate
    else
      ErrorLogMode := fmOpenReadWrite;

    ELFN := FErrorLogFileName;
    path := ExtractFilePath(ELFN);
    if path = EmptyStr then
      ELFN := ExtractFilePath(ParamStr(0)) + ELFN
    else if not DirectoryExists(path) then
      ForceDirectories(path);

    FErrorLogFS := TFileStream.Create(ELFN, ErrorLogMode);
    FErrorLogFS.Position := FErrorLogFS.Size;

    str := AnsiString(Format(QImportLoadStr(QIW_ErrorLogStarted),
      [FormatDateTime(Formats.ShortDateFormat + ' ' + Formats.ShortTimeFormat,
        Now)]) + LF + LF);
    FErrorLogFS.Write( AnsiString(str)[1], Length(str)*SizeOf(AnsiChar));
  end;

{  if FSQLLog and (SQLLogFileName <> EmptyStr) then begin
    if FSQLLogFileRewrite
      then SQLLogMode := fmCreate
      else SQLLogMode := fmOpenReadWrite;
    FSQL := TFileStream.Create(SQLLogFileName, SQLLogMode);
    FSQL.Position := FSQL.Size;
  end;}
end;

procedure TQImport3.DoImport;
var
  ImpRes: TQImportResult;
begin
  if FIsCSV then
  begin
    if Assigned(FOnSetCharsetType) then
      FOnSetCharsetType(Self, FTempFileCharset);
  end;

  StartImport;
  try
    while CheckCondition do
    begin
      FLastAction := qiaNone;
      FCustomImportError := False;
      if not Skip then
      begin
        FillImportRow;
        if FImportDestination = qidUserDefined then
          DoUserDefinedImport
        else begin
          if AllowImportRowComplete then
            DoImportRowComplete;
          if FImportEmptyRows or not FRowIsEmpty then
            ImpRes := ImportData
          else
          begin
            ImpRes := qirContinue;
            DoImportRecord;
            Inc(FCurrentLineNumber);
          end;
          case ImpRes of
            qirBreak: Break;
            qirContinue:
              begin
                ChangeCondition;
                Continue;
              end
          end;
        end;
      end
      else
        Inc(FCurrentLineNumber);
      ChangeCondition;
      Sleep(0);
    end;
  finally
    FinishImport;
  end;
end;

procedure TQImport3.DataManipulation;
begin
  if IsCSV then
  begin
    DestinationSetValues;
  end
  else begin
    if DoBeforePost then
    begin
      case ImportMode of
        qimInsertAll:
          DestinationInsert;
        qimInsertNew:
          if not DestinationFindByKey and not DestinationFindByFields then
            DestinationInsert;
        qimUpdate:
          if DestinationFindByKey then
            DestinationEdit;
        qimUpdateOrInsert:
          if DestinationFindByKey then
            DestinationEdit
          else begin
            if not DestinationFindByFields then
              DestinationInsert;
          end;
        qimDelete:
          if DestinationFindByKey then
            DestinationDelete;
        qimDeleteOrInsert:
          if DestinationFindByKey then
            DestinationDelete
          else begin
            if not DestinationFindByFields then
              DestinationInsert;
          end;
      end;
    end;

    if not FCustomImportMode then
    begin
      if (LastAction in [qiaInsert, qiaUpdate]) then
      begin
        DestinationSetValues;
        DestinationPost;
      end;
    end;
    if not FCustomImportError then
      DoAfterPost;
  end;
  if not FCustomImportError then
  begin
    DoImportRecord;
    Inc(FCurrentLineNumber);
  end;
  FPrevAction := FLastAction;
end;

procedure TQImport3.AfterImport;
var
  i: integer;
  str: AnsiString;
//  Path, FLogName: string;
begin
  FFormats.RestoreFormats;

  for i := FImportRow.Count - 1 downto 0 do
    FImportRow.Delete(i);
//  if FImportGenerators.Count > 0 then
  for i := FImportGenerators.Count - 1 downto 0 do
    FImportGenerators.Delete(i);

  DestinationEnableControls;

  // Saving ErrorLog To File
{  if (FErrors.Count > 0) and (FErrorLogFileName <> EmptyStr) then begin
    FLogName := FErrorLogFileName;
    Path := ExtractFilePath(FLogName);
    if Path = EmptyStr then begin
      GetDir(0, Path);
      Path := IncludeTrailingBackSlash(Path);
      FLogName := Path + FLogName;
    end;
    if not DirectoryExists(Path) then
      ForceDirectories(Path);
    if DirectoryExists(Path) then
      FErrors.SaveToFile(FLogName);
  end;}

  if Assigned(FErrorLogFS) then
  begin
    str := AnsiString(Format(QImportLoadStr(QIW_ErrorLogFinished),
      [FormatDateTime(Formats.ShortDateFormat + ' ' + Formats.ShortTimeFormat,
        Now)]) + LF);
    if FErrors.Count > 0 then
      str := LF + str;
    FErrorLogFS.Write(AnsiString(str)[1], Length(str)*SizeOf(AnsiChar));

    if FErrors.Count > 0
      then str := AnsiString(Format(QImportLoadStr(QIW_SomeErrorsFound),
                         [FErrors.Count]) + LF + LF)
      else str := AnsiString(QImportLoadStr(QIW_NoErrorsFound) + LF + LF);
    FErrorLogFS.Write(AnsiString(str)[1], Length(str)*SizeOf(AnsiChar));

    FErrorLogFS.Free;
    FErrorLogFS := nil;

    if ShowErrorLog {and (FErrors.Count > 0)} then
      ShellExecute(0, 'open', PChar(FErrorLogFileName), '', '', SW_SHOWNORMAL);
    end;

{  if Assigned(FSQL) then begin
    FSQL.Free;
    FSQL := nil;
  end;}
end;

procedure TQImport3.DoAfterSetFileName;
begin
//
end;

function TQImport3.CheckProperties: Boolean;
var
  i: integer;
begin
  Result := false;
  CheckDestination;
  if FileName = EmptyStr then
    raise EQImportError.Create(QImportLoadStr(QIE_NoFileName));
  if not FileExists(FileName) then
    raise EQImportError.CreateFmt(QImportLoadStr(QIE_FileNotExists), [FFileName]);
  if (FMap.Count = 0) then
    raise EQImportError.Create(QImportLoadStr(QIE_MappingEmpty));

  if not (IsCSV or (FImportDestination = qidUserDefined)) then
  begin
    for i := 0 to FMap.Count - 1 do
    begin
      if DestinationFindColumn(FMap.Names[i]) = -1 then
        raise EQImportError.CreateFmt(QImportLoadStr(QIE_FieldNotFound), [FMap.Names[i]]);
    end;
  end;

  if ImportMode <> qimInsertAll then
  begin
    if KeyColumns.Count = 0 then
      raise EQImportError.Create(QImportLoadStr(QIE_KeyColumnsNotDefined));
    for i := 0 to KeyColumns.Count - 1 do
      if DestinationFindColumn(KeyColumns[i]) = -1 then
        raise EQImportError.CreateFmt(QImportLoadStr(QIE_KeyColumnNotFound), [KeyColumns[i]]);
  end;

  {$IFNDEF NOGUI}
  if (FImportDestination = qidStringGrid) and (FGridStartRow > -1)
  and (FGridCaptionRow > FGridStartRow) then
    raise EQImportError.Create(QImportLoadStr(QIE_GridCaptionPos));
  {$ENDIF}

  if not Result then
    Result := True;
end;

procedure TQImport3.DoUserDataFormat(Col: TQImportCol);
const
  DefInputValueFlagQuote = '%';

  procedure DoReplace(const FieldFormat: TQImportFieldFormat);
var
    TextToFind,
    ReplaceWith,
    ColValue: qiString;
    I: Integer;
  begin
    // replacements
    for I := 0 to FieldFormat.Replacements.Count - 1 do
    begin
      textToFind := FieldFormat.Replacements[I].TextToFind;
      textToFind := QIStringReplace(textToFind, '#10', #10, [rfReplaceAll]);
      textToFind := QIStringReplace(textToFind, '#13', #13, [rfReplaceAll]);
      replaceWith := FieldFormat.Replacements[I].ReplaceWith;
      replaceWith := QIStringReplace(replaceWith, '#10', #10, [rfReplaceAll]);
      replaceWith := QIStringReplace(replaceWith, '#13', #13, [rfReplaceAll]);
      colValue := Col.Value;
      if FieldFormat.Replacements[I].IgnoreCase then
        colValue := QIStringReplace(colValue, textToFind, replaceWith,
          [rfReplaceAll, rfIgnoreCase])
      else
        colValue := QIStringReplace(colValue, textToFind, replaceWith,
          [rfReplaceAll]);
      Col.Value := colValue;
    end;
  end;

var
  i: integer;
  p: PAnsiChar;
  Gen: TQImportGenerator;
  AllowReplace: Boolean;
  s, value: string;
  {$IFDEF USESCRIPT}
  RaiseError: Boolean;
  ErrorMessage: qiString;
  {$ENDIF}
begin
  if not Assigned(Col) then Exit;
  i := FFieldFormats.IndexByName(Col.Name);
  if i = -1 then Exit;

  AllowReplace := True;
  //Script
  {$IFDEF USESCRIPT}
  if (FFieldFormats[i].Script.Text <> '') and Assigned(FScriptEngine) then
  begin
    FScriptEngine.Script := FFieldFormats[i].Script.Text;
    FScriptEngine.InputValue := Col.Value;
    FScriptEngine.InputValueFlag := DefInputValueFlagQuote + Col.Name +
      DefInputValueFlagQuote;

    if FScriptEngine.Execute then
      Col.Value := FScriptEngine.ResultValue
    else
      if Assigned(OnScriptExecuteError) then
      begin
        RaiseError := False;
        OnScriptExecuteError(Self, FScriptEngine.LastError, Col.Name,
          RaiseError, ErrorMessage);
        if RaiseError then
          raise Exception.Create(ErrorMessage);
      end;
    DoReplace(FFieldFormats[i]);
  end
  else
  {$ENDIF}
  begin
  // generator
  Gen := FImportGenerators.GenByName(Col.Name);
  if Assigned(Gen) then
  begin
    Col.Value := IntToStr(Gen.GetNewValue);
    Exit;
  end;
  // default value
  if (FFieldFormats[i].DefaultValue <> '') and
     (FFieldFormats[i].NullValue <> '') and

     (((FFieldFormats[i].NullValue = '''''') and (Col.Value = '')) or
      (QICompareText(Col.Value, FFieldFormats[i].NullValue) = 0)) then
  begin
    if FFieldFormats[i].DefaultValue = '''''' then
      FFieldFormats[i].DefaultValue := '';
    Col.Value := FFieldFormats[i].DefaultValue;
    AllowReplace := False;
  end;
  // constant value
  if (FFieldFormats[i].ConstantValue <> '') then
  begin
    if FFieldFormats[i].ConstantValue = '''''' then
      Col.Value := ''
    else
      Col.Value := FFieldFormats[i].ConstantValue;
    AllowReplace := False;
  end;
    //function
    case FFieldFormats[i].Functions of
      ifDate:
        Col.Value := FormatDateTime(Formats.ShortDateFormat, Date);
      ifTime:
        Col.Value := FormatDateTime(Formats.LongTimeFormat, Time);
      ifDateTime:           
        Col.Value := FormatDateTime(Formats.ShortDateFormat + #32 +
          Formats.LongTimeFormat, Now);
      ifLongFileName:
        Col.Value := FileName;
      ifShortFileName:
        Col.Value := ExtractFileName(FileName);
    end;
    if AllowReplace then
      DoReplace(FFieldFormats[i]);
  // quote action
  case FFieldFormats[i].QuoteAction of
    qaAdd: Col.Value := QIAddQuote(Col.Value, FFieldFormats[i].LeftQuote,
      FFieldFormats[i].RightQuote);
    qaRemove: Col.Value := QIRemoveQuote(Col.Value, FFieldFormats[i].LeftQuote,
      FFieldFormats[i].RightQuote);
  end;
  // AnsiChar case
  case FFieldFormats[i].CharCase of
    iccUpper: Col.Value := QIUpperCase(Col.Value);
    iccLower: Col.Value := QILowerCase(Col.Value);
    iccUpperFirst: Col.Value := QIUpperFirst(Col.Value);
    iccUpperFirstWord: Col.Value := QIUpperFirstWord(Col.Value);
  end;
  // AnsiChar set
  if Col.Value <> '' then
  begin
    case FFieldFormats[i].CharSet of
      icsAnsi:
        begin
          s := Col.Value;
          SetLength(value, Length(s));
          if Length(value) > 0 then
            {$IFDEF VCL12}
            OemToCharBuff(PAnsiChar(PChar(s)), PWideChar(value), Length(value));
            {$ELSE}
            OemToCharBuff(PChar(s), PChar(value), Length(value));
            {$ENDIF}
          Col.Value := value;
        end;
      icsOem:
        begin
          s := Col.Value;
          GetMem(p, Length(s));
          try
            CharToOemBuff(PChar(@s[1]), p, Length(s));
            Col.Value := qiString(p);
          finally
            FreeMem(p)
          end;
        end;
    end;
  end;
end;
end;

procedure TQImport3.DoUserDefinedImport;
begin
  if Assigned(FOnUserDefinedImport) then
    FOnUserDefinedImport(Self, FImportRow);
end;

function TQImport3.CanContinue: boolean;
begin
  Result := not FCanceled;
  if Assigned(FOnImportCancel) then FOnImportCancel(Self, Result);
  FCanceled := not Result;
end;

function TQImport3.AllowImportRowComplete: boolean;
begin
  Result := True;
end;

function TQImport3.GetBitesValue(const AValue: qiString): Variant;
var
  TempValue: string;
  TempByte: PByte;
  I: Integer;
begin
  TempValue := AValue;
  Result := VarArrayCreate([0, Length(TempValue) - 1], varByte);
  TempByte := VarArrayLock(Result);
  try
    for I := 1 to Length(TempValue) do
    begin
      TempByte^ := Ord(TempValue[I]);
      Inc(TempByte);
    end;
  finally
    VarArrayUnLock(Result);
  end;
end;

function TQImport3.GetDateTimeValue(const AValue: qiString): Variant;
var
  DateTimeValue: TDateTime;
begin
  if AValue <> EmptyStr then
  begin
    if TryStrToDateTime(AValue, DateTimeValue) then
      Result := DateTimeValue
    else
      raise Exception.CreateFmt(QImportLoadStr(QIE_WrongDateFormat), [AValue]);
  end
  else
    Result := NULL;
end;

procedure TQImport3.SetDataSet(const Value: TDataSet);
begin
  if FDataSet <> Value then FDataSet := Value;
end;

{$IFNDEF NOGUI}
procedure TQImport3.SetDBGrid(const Value: TDBGrid);
begin
  if FDBGrid <> Value then FDBGrid := Value;
end;

procedure TQImport3.SetListView(const Value: TListView);
begin
  if FListView <> Value then FListView := Value;
end;

{$IFDEF USESCRIPT}
procedure TQImport3.SetScriptEngine(const Value: TQImport3ScriptEngine);
begin
  if FScriptEngine <> Value then
  begin
    if Assigned(FScriptEngine) then
      FScriptEngine.RemoveFreeNotification(Self);
    FScriptEngine := Value;
    if Assigned(FScriptEngine) then
      FScriptEngine.FreeNotification(Self);
  end;
end;
{$ENDIF}

procedure TQImport3.SetStringGrid(const Value: TStringGrid);
begin
  if FStringGrid <> Value then FStringGrid := Value;
end;
{$ENDIF}

procedure TQImport3.SetKeyColumns(const Value: TStrings);
begin
  FKeyColumns.Assign(Value);
end;

procedure TQImport3.SetFileName(const Value: string);
begin
  if FFileName <> Value then begin
    FFileName := Value;
    DoAfterSetFileName;
  end;
end;

procedure TQImport3.SetMap(const Value: TStrings);
begin
  FMap.Assign(Value);
end;

function TQImport3.GetErrorRecs: integer;
begin
  Result := FErrors.Count;
end;

function TQImport3.GetFloatValue(const AValue: qiString; const AFieldType:
    TFieldType): Variant;
var
  TempValue: string;
begin
  begin
    if AValue <> '' then
    begin
      TempValue := GetNumericString(AValue);
      if DefDecimalSeparator <> Formats.DecimalSeparator then
        TempValue := StringReplace(TempValue, DefDecimalSeparator,
          Formats.DecimalSeparator, []);
      if {$IFDEF VCL6}(AFieldType <> ftFmtBCD) or {$ENDIF}
        (Length(TempValue) <= 15)
      then
        Result := StrToFloat(TempValue)
      else
        Result := TempValue; // necessary for numbers with large precision
    end
    else
      Result := NULL;
  end;
end;

function TQImport3.GetIntegerValue(const AValue: qiString; const AFieldType:
    TFieldType): Variant;
var
  TempValue: string;
begin
  try
    if AValue <> '' then
    begin
      TempValue := GetNumericString(AValue);

      if Formats.DecimalSeparator <> DefDecimalSeparator then
        TempValue := StringReplace(TempValue, DefDecimalSeparator,
          Formats.DecimalSeparator, []);
      Result := TempValue;
    {$IFDEF VCL4}
      if AFieldType = ftLargeint then
        Result := VarAsType(Result, {$IFDEF VCL6}varInt64{$ELSE}varInteger{$ENDIF})
      else
    {$ENDIF}
        Result := VarAsType(Result, varInteger);
    end
    else
      Result := NULL;
  except
    if FFormats.FBooleanTrue.IndexOf(AValue) >= 0 then
      Result := 1
    else
      if FFormats.FBooleanFalse.IndexOf(AValue) > 0 then
        Result := 0
    else
      raise;
  end
end;

function TQImport3.GetStringValue(const AValue: qiString; const AFieldType: TFieldType): Variant;
begin
  Result := AValue;
end;

procedure TQImport3.SetFormats(const Value: TQImportFormats);
begin
  FFormats.Assign(Value);
end;

procedure TQImport3.SetFieldFormats(const Value: TQImportFieldFormats);
begin
  FFieldFormats.Assign(Value);
end;

function TQImport3.StringToBoolean(const Str: string): Boolean;
begin
  Result := FFormats.FBooleanTrue.IndexOf(Str) > -1;

  if not Result and (FFormats.BooleanFalse.IndexOf(Str) < 0) then
    raise Exception.Create(Format(QImportLoadStr(QIE_IsNotBoolean), [Str]));
end;

function TQImport3.StringToField(Field: TField;
  const Str: qiString;
  AIsBinary: Boolean): string;
begin
  try
    if Assigned(Field) then
      Field.Value := PrepareImportValue(Str, AIsBinary, Field.DataType);
    Result := '';
  except
    on E: Exception do
    begin
      Field.Value := Null;
      Result := E.Message;
    end;
  end;
end;

function TQImport3.StrIsNull(const Str: string): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to FFormats.NullValues.Count - 1 do
  begin
    Result := QICompareText(Str, FFormats.NullValues[i]) = 0;
    if Result then Exit;
  end;
end;

{$IFDEF QI_UNICODE}
//function TQImport3.StringToFieldW(Field: TField; const Str: WideString;
//  AIsBinary: Boolean): AnsiString;
//
//  function StrIsNull(const Str: AnsiString): Boolean;
//  var
//    i: Integer;
//  begin
//    Result := False;
//    for i := 0 to FFormats.NullValues.Count - 1 do
//    begin
//      Result := AnsiCompareText(Str, FFormats.NullValues[i]) = 0;
//      if Result then Exit;
//    end;
//  end;
//
//  function StringToBoolean(const Str: AnsiString): Boolean;
//  begin
//    Result := FFormats.FBooleanTrue.IndexOf(Str) > -1;
//  end;
//
//var
//  DT: TDateTime;
//  tmpStr: AnsiString;
//begin
//  Result := '';
//  try
//    if Assigned(Field) then
//    begin
//      if StrIsNull(Trim(Str)) then
//      begin
//        Field.AsVariant := NULL;
//        Exit;
//      end;
//      case GetImportFieldType(Field) of
//        iftString:
//          begin
//            Field.Value := Str;
//          end;
//        iftInteger:
//          begin
//            try
//              {$IFDEF VCL4}
//              if Field is TLargeIntField then
//                (Field as TLargeIntField).AsLargeInt := Trunc(StrToFloat(GetNumericString(Str)))
//              else {$ENDIF}
//                Field.AsInteger := Trunc(StrToFloat(GetNumericString(Str)));
//            except
//              if FFormats.FBooleanTrue.IndexOf(Str) > -1 then
//                Field.AsInteger := 1
//              else if FFormats.FBooleanFalse.IndexOf(Str) > -1 then
//                Field.AsInteger := 0
//              else
//                raise;
//            end
//          end;
//        iftBoolean: Field.AsBoolean := StringToBoolean(Str);
//        iftDouble:
//          begin
//            if Str <> '' then
//            begin
//              tmpStr := StringReplace(Str, '.', DecimalSeparator, []);
//              Field.AsFloat := StrToFloat(GetNumericString(tmpStr))
//            end
//            else
//              Field.AsVariant := NULL;
//          end;
//        iftCurrency:
//          Field.AsCurrency := StrToCurr(GetNumericString(Str));
//        iftDateTime:
//          begin
//            if Str <> '' then
//            begin
//              if TryStrToDateTime(Str, DT) then
//                Field.AsDateTime := DT
//              else begin
//                Field.AsVariant := NULL;
//                raise Exception.CreateFmt(QImportLoadStr(QIE_WrongDateFormat), [Str]);
//              end;
//            end
//            else
//              Field.AsVariant := NULL;
//          end;
//        else
//          raise Exception.Create(QImportLoadStr(QIE_UnknownFieldType));
//      end;
//    end;
//  except
//    on E: Exception do
//      Result := E.Message;
//  end;
//end;
{$ENDIF}


{ TQImportCol }

constructor TQImportCol.Create;
begin
  inherited;
  FColumnIndex := -1;
  FName := EmptyStr;
  FValue := EmptyStr;
  FIsBinary := false;
end;

{ TQImportRow }

constructor TQImportRow.Create(AImport: TQImport3);
begin
  inherited Create;
  FMapNameIdxHash := THashTable.Create(False);
  FColHash := THashTable.Create(False);
  FQImport := AImport;
end;

destructor TQImportRow.Destroy;
begin
  inherited;
  FColHash.Free;
  FMapNameIdxHash.Free;
end;

function TQImportRow.Add(const AName: string): TQImportCol;
begin
  Result := TQImportCol.Create;
  Result.Name := Trim(AName);
  inherited Add(Result);
  FColHash.Insert(AName, Result);
end;

procedure TQImportRow.Clear;
var
  i: Integer;
begin
  FMapNameIdxHash.Empty;
  FColHash.Empty;
  for i := Count - 1 downto 0 do
    Delete(i);
  inherited;
end;

procedure TQImportRow.Delete(Index: integer);
begin
  TQImportCol(Items[Index]).Free;
  inherited Delete(Index);
end;

function TQImportRow.First: TQImportCol;
begin
  Result := TQImportCol(inherited First);
end;

procedure TQImportRow.SetValue(const AName, AValue: qiString;
  AIsBinary: Boolean);
var
  p: Pointer;
  i: Integer;
begin
  if FMapNameIdxHash.Search(AName, p) then
  begin
    i := Integer(p);
    if AIsBinary then
    begin
      Items[i].Value := AValue;
      Items[i].FIsBinary := True;
    end
    else
      Items[i].Value := AValue;
  end;
end;

procedure TQImportRow.ClearValues;
var
  i: integer;
begin
  for i := 0 to Count - 1 do
    Items[i].Value := EmptyStr
end;

function TQImportRow.Last: TQImportCol;
begin
  Result := TQImportCol(inherited Last);
end;

function TQImportRow.IndexOf(Item: TQImportCol): Integer;
begin
  Result := inherited IndexOf(Item);
end;

function TQImportRow.ColByName(const AName: string): TQImportCol;
var
  p: Pointer;
begin
  Result := nil;
  if FColHash.Search(AName, p) then
    Result := TQImportCol(p);
end;

function TQImportRow.Get(Index: Integer): TQImportCol;
begin
  Result := TQImportCol(inherited Get(Index));
end;

procedure TQImportRow.Put(Index: Integer; const Value: TQImportCol);
begin
  inherited Put(Index, Value);
end;

{ TQImportReplacement }

constructor TQImportReplacement.Create(Collection: TCollection);
begin
  inherited;
  FIgnoreCase := False;
end;

procedure TQImportReplacement.Assign(Source: TPersistent);
begin
  if Source is TQImportReplacement then
  begin
    TextToFind := (Source as TQImportReplacement).TextToFind;
    ReplaceWith := (Source as TQImportReplacement).ReplaceWith;
    IgnoreCase := (Source as TQImportReplacement).IgnoreCase;
    Exit;
  end;
  inherited;
end;

function TQImportReplacement.GetDisplayName: string;
begin
  Result := Format('%s - %s', [FTextToFind, FReplaceWith]);
end;

{ TQImportReplacements }

constructor TQImportReplacements.Create(Holder: TPersistent);
begin
  inherited Create(TQImportReplacement);
  FHolder := Holder;
end;

function TQImportReplacements.Add: TQImportReplacement;
begin
  Result := (inherited Add) as TQImportReplacement;
end;

function TQImportReplacements.GetOwner: TPersistent;
begin
  Result := FHolder;
end;

function TQImportReplacements.GetItem(Index: integer): TQImportReplacement;
begin
  Result := inherited Items[Index] as TQImportReplacement;
end;

procedure TQImportReplacements.SetItem(Index: integer;
  Replacement: TQImportReplacement);
begin
  inherited Items[Index] := Replacement;
end;

function TQImportReplacements.ItemExists(
  const ATextToFind, AReplaceWith: qiString;
  AIgnoreCase: Boolean): Boolean;
var
  i: Integer;
begin
  Result := False;
  for i := 0 to Count - 1 do
  begin
    if Items[i] <> nil then
    begin
      Result := (QICompareStr(Items[i].TextToFind, ATextToFind) = 0) and
        (QICompareStr(Items[i].ReplaceWith, AReplaceWith) = 0) and
        (Items[i].IgnoreCase = AIgnoreCase);
      if Result then
        Break;
    end;
  end;
end;

{ TQImportGenerators }

destructor TQImportGenerators.Destroy;
var
  i: integer;
begin
  for i := Count - 1 downto 0 do Delete(i);
  inherited;
end;

function TQImportGenerators.Add(const AName: string; AValue, AStep: Int64): TQImportGenerator;
begin
  Result := TQImportGenerator.Create;
  Result.Name := Trim(AName);
  Result.Value := AValue;
  Result.Step := AStep;
  inherited Add(Result);
end;

procedure TQImportGenerators.Delete(Index: integer);
begin
  TQImportGenerator(Items[Index]).Free;
  inherited Delete(Index);
end;

function TQImportGenerators.Get(Index: Integer): TQImportGenerator;
begin
  Result := TQImportGenerator(inherited Get(Index));
end;

function TQImportGenerators.GetNewValue(const AName: string): Int64;
var
  i: integer;
begin
  Result := 0;
  for i := 0 to Count - 1 do
    if AnsiCompareText(Items[i].Name, Trim(AName)) = 0 then
    begin
      Result := Items[i].GetNewValue;
      Exit;
    end;
end;

procedure TQImportGenerators.Put(Index: Integer; const Value: TQImportGenerator);
begin
  inherited Put(Index, Value);
end;

function TQImportGenerators.GenByName(const AName: string): TQImportGenerator;
var
  i: integer;
begin
  Result := nil;
  for i := 0 to Count - 1 do
    if AnsiCompareText(Items[i].Name, Trim(AName)) = 0 then begin
      Result := Items[i];
      Exit;
    end;
end;

{ TQImportGenerator }

constructor TQImportGenerator.Create;
begin
  inherited;
  FIsFirstRequest := True;
end;

function TQImportGenerator.GetNewValue: Int64;
begin
  if not FIsFirstRequest then
    Inc(FValue, FStep);

  Result := FValue;
  FIsFirstRequest := False;
end;

{ TQImportFieldFormat }

constructor TQImportFieldFormat.Create(Collection: TCollection);
begin
  inherited;
  FFieldName := '';
  FGeneratorValue := 0;
  FGeneratorStep := 0;

  FConstantValue := '';
  FNullValue := '';
  FDefaultValue := '';
  FLeftQuote := '';
  FRightQuote := '';

  FQuoteAction := qaNone;
  FCharCase := iccNone;
  FCharSet := icsNone;
  FReplacements := TQImportReplacements.Create(Self);
  FScript := TqiStringList.Create;
end;

destructor TQImportFieldFormat.Destroy;
begin
  FScript.Free;
  FReplacements.Free;
  inherited;
end;

function TQImportFieldFormat.GetDisplayName: string;
begin
  Result := inherited GetDisplayName;
end;

function TQImportFieldFormat.IsConstant: boolean;
begin
  Result := FConstantValue <> '';
end;

function TQImportFieldFormat.IsNull: boolean;
begin
  Result := FNullValue <> '';
end;

function TQImportFieldFormat.IsDefault: boolean;
begin
  Result := FDefaultValue <> '';
end;

function TQImportFieldFormat.IsLeftQuote: boolean;
begin
  Result := FLeftQuote <> '';
end;

function TQImportFieldFormat.IsRightQuote: boolean;
begin
  Result := FRightQuote <> '';
end;

procedure TQImportFieldFormat.Assign(Source: TPersistent);
begin
  if Source is TQImportFieldFormat then
  begin
    FieldName := (Source as TQImportFieldFormat).FieldName;
    GeneratorValue := (Source as TQImportFieldFormat).GeneratorValue;
    GeneratorStep := (Source as TQImportFieldFormat).GeneratorStep;

    ConstantValue := (Source as TQImportFieldFormat).ConstantValue;
    NullValue := (Source as TQImportFieldFormat).NullValue;
    DefaultValue := (Source as TQImportFieldFormat).DefaultValue;
    LeftQuote := (Source as TQImportFieldFormat).LeftQuote;
    RightQuote := (Source as TQImportFieldFormat).RightQuote;

    QuoteAction := (Source as TQImportFieldFormat).QuoteAction;
    CharCase := (Source as TQImportFieldFormat).CharCase;
    CharSet := (Source as TQImportFieldFormat).CharSet;
    Replacements := (Source as TQImportFieldFormat).Replacements;
    Functions := (Source as TQImportFieldFormat).Functions;
    Script.Text := (Source as TQImportFieldFormat).Script.Text;
    Exit;
  end;
  inherited;
end;

function TQImportFieldFormat.IsDefaultValues: boolean;
begin
  Result := (GeneratorValue = 0) and
            (GeneratorStep = 0) and

            (ConstantValue = '') and
            (NullValue = '') and
            (DefaultValue = '') and
            (LeftQuote = '') and
            (RightQuote = '') and

            (QuoteAction = qaNone) and
            (CharCase = iccNone) and
            (CharSet = icsNone) and
            (Replacements.Count = 0) and
            (Functions = ifNone) and
            (Script.Text = '');
end;

procedure TQImportFieldFormat.SetReplacements(const Value: TQImportReplacements);
begin
  FReplacements.Assign(Value);
end;

{ TQImportFieldFormats }

constructor TQImportFieldFormats.Create(AHolder: TComponent);
begin
  inherited Create(TQImportFieldFormat);
  FHolder := AHolder;
end;

function TQImportFieldFormats.Add: TQImportFieldFormat;
begin
  Result := TQImportFieldFormat(inherited Add);
end;

function TQImportFieldFormats.GetItem(Index: integer): TQImportFieldFormat;
begin
  Result := TQImportFieldFormat(inherited Items[Index]);
end;

function TQImportFieldFormats.GetOwner: TPersistent;
begin
  Result := FHolder;
end;

procedure TQImportFieldFormats.SetItem(Index: integer; FieldFormat: TQImportFieldFormat);
begin
  inherited SetItem(Index, TCollectionItem(FieldFormat));
end;

function TQImportFieldFormats.IndexByName(const FieldName: string): integer;
var
  i: integer;
begin
  Result := -1;
  for i := 0 to Count - 1 do
  begin
    if AnsiCompareText(FieldName, Items[i].FieldName) = 0 then
    begin
      Result := i;
      Exit;
    end;
  end;
end;

{ TQImportLocale }

constructor TQImportLocale.Create;
begin
  FIDEMode := AnsiUpperCase(ExtractFileName(ParamStr(0))) = 'DELPHI32.EXE';
end;

procedure TQImportLocale.LoadDll(const Name: string);
begin
  if FLoaded then
    UnloadDll;
  FDllHandle := LoadLibrary(PChar(Name));
  FLoaded := FDllHandle <> HINSTANCE_ERROR;
end;

function TQImportLocale.LoadStr(ID: Integer): string;
var
  Buffer: array[0..1023] of Char;
  Handle: THandle;
begin
  if Assigned(FOnLocalize) then
  begin
    Result := '';
    FOnLocalize(ID, Result);
    if Result <> '' then
      Exit;
  end;

  if FLoaded then
    Handle := FDllHandle
  else
    Handle := HInstance;

  if FIDEMode then
    Result := {$IFDEF VCL16}System.SysUtils{$ELSE}SysUtils{$ENDIF}.LoadStr(ID)
  else
    SetString(Result, Buffer, LoadString(Handle, ID, Buffer, SizeOf(Buffer)));
end;

procedure TQImportLocale.UnloadDll;
begin
  if FLoaded then
    FreeLibrary(FDllHandle);
  FLoaded := False;
end;

{$IFDEF QI_UNICODE}

{ WideException }

constructor WideException.Create(const Msg: WideString);
begin
  FWideMessage := Msg;
end;

constructor WideException.CreateFmt(const Msg: WideString;
  const Args: array of const);
begin
  FWideMessage := WideFormatStr(Msg, Args);
end;

{$ENDIF}

initialization

finalization
  Locale.Free;
  Locale := nil;

end.
