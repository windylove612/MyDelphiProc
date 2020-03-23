unit QImport3Common;

{$I QImport3VerCtrl.Inc}

{$IFDEF VCL6}
  {$WARN SYMBOL_PLATFORM OFF}
{$ENDIF}

interface

uses
  {$IFDEF VCL16}
    {$IFDEF QI_UNICODE}
      System.WideStrings,
      QImport3WideStringGrid,
    {$ENDIF}
    System.Classes,
    System.IniFiles,
    Data.DB,
    Winapi.Windows,
    Vcl.Graphics,
    Soap.EncdDecd,
    {$IFNDEF NOGUI}
      Vcl.Grids,
      Vcl.Controls,
      Vcl.DBGrids,
      Vcl.ComCtrls,
    {$ENDIF}
  {$ELSE}
    {$IFDEF QI_UNICODE}
      {$IFDEF VCL10}
        WideStrings,
      {$ELSE}
        QImport3WideStrings,
      {$ENDIF}
      QImport3WideStringGrid,
    {$ENDIF}
    Classes,
    IniFiles,
    {$IFDEF VCL6}
    EncdDecd,
    {$ENDIF}
    DB,
    Windows,
    Graphics,
    {$IFNDEF NOGUI}
      Grids,
      Controls,
      DBGrids,
      ComCtrls,
    {$ENDIF}
  {$ENDIF}
  QImport3,
  QImport3XLSMapParser,
  QImport3StrTypes;

const
  //---------------------
  //  Don't localize it !
  //---------------------
  QI_PALETTE_PAGE = 'EMS Advanced Import 3';
  QI_ABOUT = '(About EMS AdvancedImport 3)';
{$IFDEF CS}
  QI_REG_URL = 'http://www.componentsource.com';
{$ELSE}
  QI_REG_URL = 'http://www.sqlmanager.net/products/tools/advancedimport/buy';
{$ENDIF}

  QI_WIZARD_HELP = 'AImportWizard.hlp';

  QI_FULL_PRODUCT_NAME = 'Advanced Data Import for Delphi(R) and C++ Builder(R) (%s)';
  QI_COPYRIGHT = '(C) 1999-2013  EMS Database Management Solutions, Ltd.  All rights reserved.';
  QI_VERSION = '3.9.0.2';

  QIW_FIRST_STEP = 'FIRST_STEP';
    QIW_FILE_NAME = 'FileName';
    QIW_IMPORT_TYPE = 'ImportType';
    QIW_GO_TO_LAST_PAGE = 'GoToLastPage';
    QIW_AUTO_SAVE_TEMPLATE = 'AutoSaveTemplate';
  QIW_XLSX_MAP = 'XLSX_MAP';
  QIW_XLSX_OPTIONS = 'XLSX_OPTIONS';
    QIW_XLSX_SKIP_LINES = 'SkipLines';
    QIW_XLSX_SHEET_NAME = 'SheetName';
    QIW_XLSX_NEED_FILLMERGE = 'NeedFillMerge';
    QIW_XLSX_LOAD_HIDDENSHEET = 'LoadHiddenSheet';
  QIW_DOCX_MAP = 'DOCX_MAP';
    QIW_DOCX_SKIP_LINES = 'SkipLines';
    QIW_DOCX_TABLE_NUMBER = 'TableNumber';
  QIW_DOCX_OPTIONS = 'DOCX_OPTIONS';
  QIW_ODS_MAP = 'ODS_MAP';
  QIW_ODS_OPTIONS = 'ODS_OPTIONS';
    QIW_ODS_SKIP_LINES = 'SkipLines';
    QIW_ODS_SHEET_NAME = 'SheetName';
  QIW_ODT_MAP = 'ODT_MAP';
  QIW_ODT_OPTIONS = 'ODT_OPTIONS';
    QIW_ODT_SKIP_LINES = 'SkipLines';
    QIW_ODT_SHEET_NAME = 'SheetName';
    QIW_ODT_USE_HEADER = 'UseHeader';
  QIW_XLS_MAP = 'XLS_MAP';
  QIW_XLS_OPTIONS = 'XLS_OPTIONS';
    QIW_XLS_SKIP_COLS = 'SkipCols';
    QIW_XLS_SKIP_ROWS = 'SkipRows';
  QIW_TXT_MAP = 'TXT_MAP';
  QIW_TXT_OPTIONS = 'ASCII_OPTIONS';
    QIW_TXT_SKIP_LINES = 'SkipLines';
    QIW_TXT_ENCODING_OPTION = 'Encoding';
  QIW_CSV_MAP = 'CSV_MAP';
  QIW_CSV_OPTIONS = 'CSV_OPTIONS';
    QIW_CSV_DELIMITER = 'Delimiter';
    QIW_CSV_SKIP_LINES = 'SkipLines';
    QIW_CSV_QUOTE = 'Quote';
  QIW_DBF_MAP = 'DBF_MAP';
  QIW_DBF_OPTIONS = 'DBF_OPTIONS';
    QIW_DBF_SKIP_DELETED = 'SkipDeleted';
  QIW_HTML_MAP = 'HTML_MAP';
  QIW_HTML_OPTIONS = 'HTML_OPTIONS';
    QIW_HTML_SKIP_LINES = 'SkipLines';
  QIW_XML_MAP = 'XML_MAP';
  QIW_XML_OPTIONS = 'XML_OPTIONS';
    QIW_XML_WRITE_ON_FLY = 'WriteOnFly';
  QIW_MDB_MAP = 'ACCESS_MAP';
  QIW_MDB_OPTIONS = 'ACCESS_OPTIONS';
    QIW_MDB_PASSWORD = 'Password';
    QIW_MDB_SOURCETYPE = 'SourceType';
    QIW_MDB_TABLENAME = 'TableName';
  QIW_MDB_QUERY = 'ACCESS_SQL';
    QIW_MDB_SQL_LINE = 'Line_';
  QIW_BASE_FORMATS = 'BASE_FORMATS';
    QIW_BF_DECIMAL_SEPARATOR = 'DecimalSeparator';
    QIW_BF_THOUSAND_SEPARATOR = 'ThousandSeparator';
    QIW_BF_SHORT_DATE_FORMAT = 'ShortDateFormat';
    QIW_BF_LONG_DATE_FORMAT = 'LongDateFormat';
    QIW_BF_DATE_SEPARATOR = 'DateSeparator';
    QIW_BF_SHORT_TIME_FORMAT = 'ShortTimeFormat';
    QIW_BF_LONG_TIME_FORMAT = 'LongTimeFormat';
    QIW_BF_TIME_SEPARATOR = 'TimeSeparator';
    //QIW_BF_LEFT_QUOTATION = 'LeftQuotation';
    //QIW_BF_RIGHT_QUOTATION = 'RightQuotation';
    //QIW_BF_QUOTATION_ACTION = 'QuotationAction';
  QIW_BOOLEAN_TRUE = 'BOOLEAN_TRUE';
  QIW_BOOLEAN_FALSE = 'BOOLEAN_FALSE';
  QIW_NULL_VALUES = 'NULL_VALUES';
  QIW_XML_DOC_MAP = 'XML_DOC_MAP';
    QIW_XML_DOC_MAP_LINE = 'QIW_XML_DOC_MAP_LINE_';
  QIW_XML_DOC_OPTIONS = 'XML_DOC_OPTIONS';
    QIW_XML_DOC_XPATH = 'QIW_XML_DOC_XPATH';
    QIW_XML_DOC_DATALOCATION = 'QIW_XML_DOC_DATALOCATION';
    QIW_XML_DOC_SKIPLINES = 'QIW_XML_DOC_SKIPLINES';
  QIW_DATA_FORMATS = 'DATA_FORMATS_';
    QIW_DF_GENERATOR_VALUE = 'GeneratorValue';
    QIW_DF_GENERATOR_STEP = 'GeneratorStep';
    QIW_DF_CONSTANT_VALUE = 'ConstantValue';
    QIW_DF_NULL_VALUE = 'NullValue';
    QIW_DF_DEFAULT_VALUE = 'DefaultValue';
    QIW_DF_LEFT_QUOTE = 'LeftQuote';
    QIW_DF_RIGHT_QUOTE = 'RightQuote';
    QIW_DF_QUOTE_ACTION = 'QuoteAction';
    QIW_DF_CHAR_CASE = 'CharCase';
    QIW_DF_CHAR_SET = 'CharSet';
  QIW_REPLACEMENTS = 'REPLACEMENTS_';
    QIW_RP_TEXT_TO_FIND = 'TextToFind';
    QIW_RP_REPLACE_WITH = 'ReplaceWith';
    QIW_RP_IGNORE_CASE = 'IgnoreCase';
  QIW_ITEM = 'ITEM_';
  QIW_LAST_STEP = 'LAST_STEP';
    QIW_COMMIT_AFTER_DONE = 'CommitAfterDone';
    QIW_COMMIT_REC_COUNT = 'CommitRecCount';
    QIW_IMPORT_REC_COUNT = 'ImportRecCount';
    QIW_CLOSE_AFTER_IMPORT = 'CloseAfterImport';
    QIW_ENABLE_ERROR_LOG = 'EnableErrorLog';
    QIW_ERROR_LOG_FILE_NAME = 'ErrorLogFileName';
    QIW_REWRITE_ERROR_LOG_FILE = 'RewriteErrorLogFile';
    QIW_SHOW_ERROR_LOG = 'ShowErrorLog';
    QIW_ENABLE_SQL_LOG = 'EnableSQLLog';
    QIW_SQL_LOG_FILE_NAME = 'SQLLogFileName';
    QIW_REWRITE_SQL_LOG_FILE = 'RewriteSQLLogFile';
    QIW_IMPORT_MODE = 'ImportMode';
    QIW_ADD_TYPE = 'AddType';
    QIW_KEY_COLUMNS = 'KeyColumns';
    QIW_AUTO_TRIM_VALUE = 'AutoTrimValue';
    QIW_IMPORT_EMPTY_ROWS = 'ImportEmptyRows';


  QI_BASE = 'BASE';
    QI_FILE_NAME = 'FileName';
    QI_IMPORT_TYPE = 'ImportType';
    QI_IMPORT_DESTINATION = 'ImportDestination';
    QI_IMPORT_MODE = 'ImportMode';
    QI_MAP = 'MAP';
    QI_KEY_COLUMNS = 'KeyColumns';
    QI_GRID_CAPTION_ROW = 'GridCaptionRow';
    QI_GRID_START_ROW = 'GridStartRow';
    QI_COMMIT_AFTER_DONE = 'CommitAfterDone';
    QI_COMMIT_REC_COUNT = 'CommitRecCount';
    QI_IMPORT_REC_COUNT = 'ImportRecCount';
    QI_ENABLE_ERROR_LOG = 'EnableErrorLog';
    QI_ERROR_LOG_FILE_NAME = 'ErrorLogFileName';
    QI_REWRITE_ERROR_LOG_FILE = 'RewriteErrorLogFile';
    QI_SHOW_ERROR_LOG = 'ShowErrorLog';
    QI_IMPORT_EMPTY_ROWS = 'ImportEmptyRows';
    QI_AUTO_TRIM_VALUE = 'AutoTrimValue';

  QI_REPLACEMENTS = 'REPLACEMENTS_';
    QI_RP_TEXT_TO_FIND = 'TextToFind';
    QI_RP_REPLACE_WITH = 'ReplaceWith';
    QI_RP_TEXT_TO_FIND_HEX = 'TextToFindHex';
    QI_RP_REPLACE_WITH_HEX = 'ReplaceWithHex';
    QI_RP_IGNORE_CASE = 'IgnoreCase';

    {QI_ENABLE_SQL_LOG = 'EnableSQLLog';
    QI_SQL_LOG_FILE_NAME = 'SQLLogFileName';
    QI_SQL_LOG_FILE_REWRITE = 'SQLLogFileRewrite';}
  BASE_FORMATS = 'BASE_FORMATS';
    BF_DECIMAL_SEPARATOR = 'DecimalSeparator';
    BF_THOUSAND_SEPARATOR = 'ThousandSeparator';
    BF_SHORT_DATE_FORMAT = 'ShortDateFormat';
    BF_LONG_DATE_FORMAT = 'LongDateFormat';
    BF_SHORT_TIME_FORMAT = 'ShortTimeFormat';
    BF_LONG_TIME_FORMAT = 'LongTimeFormat';
    //BF_LEFT_QUOTATION = 'LeftQuotation';
    //BF_RIGHT_QUOTATION = 'RightQuotation';
    BF_QUOTATION_ACTION = 'QuotationAction';
  BOOLEAN_TRUE = 'BOOLEAN_TRUE';
  BOOLEAN_FALSE = 'BOOLEAN_FALSE';
  NULL_VALUES = 'NULL_VALUES';
  DATA_FORMATS = 'DATA_FORMATS_';
    DF_GENERATOR_VALUE = 'GeneratorValue';
    DF_GENERATOR_STEP = 'GeneratorStep';
    DF_CONSTANT_VALUE = 'ConstantValue';
    DF_NULL_VALUE = 'NullValue';
    DF_DEFAULT_VALUE = 'DefaultValue';
    DF_LEFT_QUOTE = 'LeftQuote';
    DF_RIGHT_QUOTE = 'RightQuote';
    DF_QUOTE_ACTION = 'QuoteAction';
    DF_CHAR_CASE = 'CharCase';
    DF_CHAR_SET = 'CharSet';
    QI_FUNCTION = 'Function';

  XLS_OPTIONS = 'XLS_OPTIONS';
    XLS_SKIP_COLS = 'SkipCols';
    XLS_SKIP_ROWS = 'SkipRows';
  ASCII_OPTIONS = 'ASCII_OPTIONS';
    ASCII_SKIP_LINES = 'SkipLines';
    ASCII_COMMA = 'Comma';
    ASCII_Quote = 'Quote';
    ASCII_ENCODING = 'Encoding';
  DBF_OPTIONS = 'DBF_OPTIONS';
    DBF_SKIP_LINES = 'SkipLines';
    DBF_CHARSET = 'CharSet';
  HTML_OPTIONS = 'HTML_OPTIONS';
    HTML_SKIP_LINES = 'SkipLines'; {!HTML}
  XML_OPTIONS = 'XML_OPTIONS';
    XML_SKIP_LINES = 'SkipLines';
    XML_WRITE_ON_FLY = 'WriteOnFly';
  DATA_SET_OPTIONS = 'DATA_SET_OPTIONS';
    DATA_SET_SKIP_LINES = 'SkipLines';
  ODS_OPTIONS = 'ODS_OPTIONS';
    ODS_SKIP_LINES = 'SkipRows';
    ODS_SHEET_NAME = 'SheetName';
    ODS_NOT_EXPAND_MERGED_VALUE = 'NotExpandMergedValue';
  ODT_OPTIONS = 'ODT_OPTIONS';
    ODT_SKIP_LINES = 'SkipRows';
    ODT_SHEET_NAME = 'SheetName';
    ODT_USE_HEADER = 'UseHeader';
    ODT_COMPLEX_TABLE = 'ComplexTable';
  XLSX_OPTIONS = 'XLSX_OPTIONS';
    XLSX_SKIP_LINES = 'SkipRows';
    XLSX_SHEET_NAME = 'SheetName';
    XLSX_LOAD_HIDDEN_SHEET = 'LoadHiddenSheet';
    XLSX_NEED_FILL_MERGE = 'NeedFillMerge';
  DOCX_OPTIONS = 'DOCX_OPTIONS';
    DOCX_SKIP_LINES = 'SkipRows';
    DOCX_TABLE_NUMBER = 'TableNumber';
    DOCX_NEED_FILLMERGE = 'NeedFillMerge';
  ACCESS_OPTIONS = 'ACCESS_OPTIONS';
    ACCESS_SKIP_LINES = 'SkipRows';
    ACCESS_TABLE_NAME = 'TableName';
    ACCESS_SQL = 'SQL';
    ACCESS_SQL_LINE = 'Line';
    ACCESS_SOURCE_TYPE = 'SourceType';
    ACCESS_PASSWORD = 'Password';

  XLS_EXT = '.xls';
  XLSX_EXT = '.xlsx';
  DOCX_EXT = '.docx';
  ODS_EXT = '.ods';
  ODT_EXT = '.odt';
  DBF_EXT = '.dbf';
  DB_EXT  = '.db';
  HTM_EXT = '.htm';
  HTML_EXT = '.html';
  XML_EXT = '.xml';
  TXT_EXT = '.txt';
  CSV_EXT = '.csv';
  MDB_EXT = '.mdb';

const
  PathDelim  = {$IFDEF WIN} '\'; {$ELSE} '/'; {$ENDIF}

  SShortMonthNameJan = 'Jan';
  SShortMonthNameFeb = 'Feb';
  SShortMonthNameMar = 'Mar';
  SShortMonthNameApr = 'Apr';
  SShortMonthNameMay = 'May';
  SShortMonthNameJun = 'Jun';
  SShortMonthNameJul = 'Jul';
  SShortMonthNameAug = 'Aug';
  SShortMonthNameSep = 'Sep';
  SShortMonthNameOct = 'Oct';
  SShortMonthNameNov = 'Nov';
  SShortMonthNameDec = 'Dec';

  SLongMonthNameJan = 'January';
  SLongMonthNameFeb = 'February';
  SLongMonthNameMar = 'March';
  SLongMonthNameApr = 'April';
  SLongMonthNameMay = 'May';
  SLongMonthNameJun = 'June';
  SLongMonthNameJul = 'July';
  SLongMonthNameAug = 'August';
  SLongMonthNameSep = 'September';
  SLongMonthNameOct = 'October';
  SLongMonthNameNov = 'November';
  SLongMonthNameDec = 'December';

  SShortDayNameSun = 'Sun';
  SShortDayNameMon = 'Mon';
  SShortDayNameTue = 'Tue';
  SShortDayNameWed = 'Wed';
  SShortDayNameThu = 'Thu';
  SShortDayNameFri = 'Fri';
  SShortDayNameSat = 'Sat';

  SLongDayNameSun = 'Sunday';
  SLongDayNameMon = 'Monday';
  SLongDayNameTue = 'Tuesday';
  SLongDayNameWed = 'Wednesday';
  SLongDayNameThu = 'Thursday';
  SLongDayNameFri = 'Friday';
  SLongDayNameSat = 'Saturday';

  SystemCTNames: array[TQICharsetType] of Integer =
    (
        -1, 1252, -1, 20127, 850, 852, 866, 1250, 1251, 1256, 1257, -1, -1, 28597,
      28598, -1, -1, 20866, 21866, 870, 1026, -1, 10029, 10000, 20107, CP_UTF8, 1200, 12000,
      // unique in postrgesql
      28593, -1, -1, -1, 28595, 28596,
      //unique in db2
      1026, 1254, 1255, 1258, 437, 500, 737, 855, 856, 857, 860, 862, 863, 864, 865, 869,
      874, 875, 871,
      //unique in IB/FB
      950, 949, -1, 936, -1, 28599, 28594, 1252, 1253, 775, 858
    );

  QICharsetTypeNames: array[TQICharsetType] of string =
    (
      'Windows default',
      'latin1 (ISO 8859-1 West European)',
      'armscii8 (ARMSCII-8 Armenian)',
      'ascii (US ASCII)',
      'cp850 (DOS West European)',
      'cp852 (DOS Central European)',
      'cp866 (DOS Russian)',
      'cp1250 (Windows Central European)',
      'cp1251 (Windows Cyrillic)',
      'cp1256 (Windows Arabic)',
      'cp1257 (Windows Baltic)',
      'dec8 (DEC West European)',
      'geostd8 (GEOSTD8 Georgian)',
      'greek (ISO 8859-7 Greek)',
      'hebrew (ISO 8859-8 Hebrew)',
      'hp8 (HP West European)',
      'keybcs2 (DOS Kamenicky Czech-Slova)',
      'koi8r (KOI8-R Relcom Russian)',
      'koi8u (KOI8-U Ukrainian)',
      'latin2 (ISO 8859-2 Central European',
      'latin5 (ISO 8859-9 Turkish)',
      'latin7 ((ISO 8859-13 Baltic)',
      'macce (Mac Central European)',
      'macroman (Mac West European)',
      'swe7 (7bit Swedish)',
      'Unicode (UTF-8)',
      'Unicode (UTF-16)',
      'Unicode (UTF-32)',
      'latin3 (ISO 8859-3)',
      'latin4 (ISO 8859-4)',
      'latin6 (ISO 8859-10)',
      'latin8 (ISO 8859-14)',
      'iso8859-5 (Cyrillic)',
      'iso8859-6 (Arabic)',

      'CCSID 1026 - EBCDIC Turkey Latin-5',
      'Codepage 1254 - Windows Turkish',
      'Codepage 1255 - Windows Hebrew',
      'Codepage 1258 - Windows Vietnamese',
      'Codepage 437 - US, Europe',
      'CCSID 500 - EBCDIC Latin-1',
      'Codepage 737 - WIN Greece',
      'Codepage 855 - OS2 Cyrillic',
      'Codepage 856 - Hebrew',
      'Codepage 857 - OS2 Turkey',
      'Codepage 860 - Portuguese',
      'Codepage 862 - OS2 Hebrew',
      'Codepage 863 - Canadian-French',
      'Codepage 864 - OS2 Arabic',
      'Codepage 865 - Norway, Denmark',
      'Codepage 869 - OS2 Greece',
      'Codepage 874 - OS2/AIX Thailand',
      'CCSID 875 - EBCDIC Greece',
      'CCSID 871 - EBCDIC Iceland',
      'Big5',
      'Korean (Unified Hangeul)',
      'EUC Japanese',
      'Simplified Chinese (Hong Kong, PRC)',
      'Japanese',
      'iso8859-9 (Turkish)',
      'iso8859-4 (Baltic)',
      'Codepage 1252 - Western Europe, Amerika',
      'Codepage 1253 - Modern Greek',
      'Codepage 775 - Baltic',
      'Codepage 858 - Latin-1'
    );

var
  DefShortMonthNames: array[1..12] of string = (SShortMonthNameJan,
    SShortMonthNameFeb, SShortMonthNameMar, SShortMonthNameApr,
    SShortMonthNameMay, SShortMonthNameJun, SShortMonthNameJul,
    SShortMonthNameAug, SShortMonthNameSep, SShortMonthNameOct,
    SShortMonthNameNov, SShortMonthNameDec);

  DefLongMonthNames: array[1..12] of string = (SLongMonthNameJan,
    SLongMonthNameFeb, SLongMonthNameMar, SLongMonthNameApr,
    SLongMonthNameMay, SLongMonthNameJun, SLongMonthNameJul,
    SLongMonthNameAug, SLongMonthNameSep, SLongMonthNameOct,
    SLongMonthNameNov, SLongMonthNameDec);

function StringToChar(const S: AnsiString; const Default: AnsiChar): AnsiChar;
function PadString(const S: AnsiString; Chr: AnsiChar; Len: integer): AnsiString;
function Str2Char(const Str: string; Default: Char): Char; overload;
{$IFDEF VCL12}
function Str2Char(const Str: string; Default: AnsiChar): AnsiChar; overload;
{$ENDIF}
function Char2Str(Chr: Char): string; overload;
{$IFDEF VCL12}
function Char2Str(Chr: AnsiChar): string; overload;
{$ENDIF}

{$IFDEF QI_UNICODE}
function QIAddQuote(const S, LeftQuote, RightQuote: WideString): WideString;
function QIRemoveQuote(const S, LeftQuote, RightQuote: WideString): WideString;
function QIUpperFirst(const S: WideString): WideString;
function QIUpperFirstWord(const S: WideString): WideString;
{$ELSE}
function QIAddQuote(const S, LeftQuote, RightQuote: AnsiString): AnsiString;
function QIRemoveQuote(const S, LeftQuote, RightQuote: AnsiString): AnsiString;
function QIUpperFirst(const S: AnsiString): AnsiString;
function QIUpperFirstWord(const S: AnsiString): AnsiString;
{$ENDIF}

{$IFDEF QI_UNICODE}
procedure CSVStringToStrings(
  const Str: WideString; Quote, Separator: AnsiChar; AStrings: TWideStrings);
{$ELSE}
procedure CSVStringToStrings(
  const Str: AnsiString; Quote, Separator: AnsiChar; AStrings: TStrings);
{$ENDIF}

procedure ReplaceTabs(var Str: qiString);

// StringGrid routines
{$IFNDEF NOGUI}
procedure GridDrawCell(Grid: TqiStringGrid;
  const SheetName: string; SheetNumber, ACol, ARow: integer; Rect: TRect;
  State: TGridDrawState; DefinedRanges: TMapRow; SkipCols, SkipRows: integer;
  IsEditing: boolean; Selection: TMapRow);
procedure GridMouseDown(Grid: TqiStringGrid;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer;  FirstCol, LastCol, FirstRow, LastRow,
  SkipRows, SkipCols: integer; var CurrSel: string);
procedure GridSelectCell(Grid: TqiStringGrid;
  ACol, ARow: integer; Shift: TShiftState; var CurrSel: string);
procedure GridMoveCurrCell(Grid: TqiStringGrid;
  ACol, ARow: integer);
procedure GridSelCell(Grid: TqiStringGrid;
  RowNo, ColNo: integer; NeedClear: boolean; var CurrSel: string);
procedure GridSelRow(Grid: TqiStringGrid;
  FirstCol, LastCol, RowNo, StartCol, SkipRows, SkipCols: integer;
  NeedClear: boolean; var CurrSel: string);
procedure GridSelCol(Grid: TqiStringGrid;
  FirstRow, LastRow, ColNo, StartRow, SkipRows, SkipCols: integer;
  NeedClear: boolean; var CurrSel: string);
procedure GridFillFixedCells(Grid: TqiStringGrid);
//procedure GridClearDataCell(Grid: TStringGrid);
{$ENDIF}

// ImportDestination routines
procedure QImportCheckDestination(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF});
function QImportIsDestinationActive(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
  ListView: TListView;
  StringGrid: TStringGrid{$ENDIF}): boolean;
procedure QImportIsDestinationOpen(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF});
procedure QImportIsDestinationClose(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF});
function QImportDestinationColCount(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF}): integer;
function QImportDestinationColName(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet;
  {$IFNDEF NOGUI} DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid;
 GridCaptionRow,{$ENDIF} Index: integer): string;
function QImportDestinationAssigned(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF}): boolean;
function QImportDestinationFindColumn(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet;
  {$IFNDEF NOGUI}DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid; GridCaptionRow: integer;
{$ENDIF}
  const ColName: string): integer;

//dee function GetListSeparator: AnsiChar;
function GetListSeparator: Char;
procedure ClearIniFile(IniFile: TIniFile);
function IncludePathDelimiter(const S: string): string;
function GetNumericString(const Str: string): string;
function TryStrToDateTime(const Str: string; var DateTime: TDateTime): Boolean;


// 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009
// 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009
// 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009
type
  TSysCharSet = set of AnsiChar;
function CharInSet(C: AnsiChar; const CharSet: TSysCharSet): Boolean; overload;
function CharInSet(C: WideChar; const CharSet: TSysCharSet): Boolean; overload; 
{$IFDEF VCL12}
function UpperCase(const S: AnsiString): AnsiString; overload;
{$ENDIF}
// 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009 2009

{$IFNDEF VCL5}
type
  TReplaceFlags = set of (rfReplaceAll, rfIgnoreCase);

function StringReplace(const S, OldPattern, NewPattern: AnsiString;
  Flags: TReplaceFlags): AnsiString;
{$ENDIF}

function DirExists(const Directory: string): Boolean;

function FormatStrToDateTime(ADateTimeStr, AFormatStr: string): TDateTime;

function MinimizeName(const Filename: string; Canvas: TCanvas;
  MaxLen: Integer): string;

function GetColIdFromColIndex(ColIndex: string): Integer;

function IsUnicodeQICharset(Charset: TQICharsetType): boolean;
function QICharsetToCodepage(const Value: TQICharsetType): Integer;

{$IFDEF QI_UNICODE}
procedure FillCharsetTypeList(List: TWideStrings);
{$ELSE}
procedure FillCharsetTypeList(List: TStrings);
{$ENDIF}
function GetFieldDataType(const Value: string): TFieldType;

function QIGetEmptyStr: AnsiString;

function FileToBase64(AFileName: string): string;

{$IFNDEF VCL6}
function DecodeBase64(const CinLine: string): string;
function EncodeBase64(const inStr: string): string;
{$ENDIF}

function QIDecodeBase64(AValue: qiString): string;
function QIEncodeBase64(AValue: qiString): string;

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
    {$IFDEF VCL17}
      System.UITypes,
    {$ENDIF}
  {$ELSE}
    SysUtils,
  {$ENDIF}
  QImport3XLSUtils,
  QImport3StrIDs,
  QImport3WideStrUtils;

function QIGetEmptyStr: AnsiString;
begin
  Result := {$IFDEF VCL12} EmptyAnsiStr {$ELSE} EmptyStr {$ENDIF};
end;

function CharInSet(C: AnsiChar; const CharSet: TSysCharSet): Boolean;
begin
  Result := C in CharSet;
end;

function CharInSet(C: WideChar; const CharSet: TSysCharSet): Boolean;
begin
  Result := (C < #$0100) and (AnsiChar(C) in CharSet);
end;
{$IFDEF VCL12}
function UpperCase(const S: AnsiString): AnsiString;
var
  I: Integer;
  P: PAnsiChar;
begin
  Result := S;
  if Result <> '' then
  begin
    P := PAnsiChar(Pointer(Result));
    for I := Length(Result) downto 1 do
    begin
      case P^ of
        'a'..'z':
          P^ := AnsiChar(Word(P^) xor $0020);
      end;
      Inc(P);
    end;
  end;
end;
{$ENDIF}

function IsUnicodeQICharset(Charset: TQICharsetType): boolean;
begin
  Result := (Charset = ctUtf8) or (Charset = ctUtf16) or (Charset = ctUtf32);
end;

function QICharsetToCodepage(const Value: TQICharsetType): Integer;
begin
  if Value = ctWinDefined then
    Result := GetACP
  else
    Result := SystemCtNames[Value];
end;

{$IFDEF QI_UNICODE}
procedure FillCharsetTypeList(List: TWideStrings);
{$ELSE}
procedure FillCharsetTypeList(List: TStrings);
{$ENDIF}
var
  I: TQICharsetType;
begin
  List.Clear;
  for I := Low(TQICharsetType) to High(TQICharsetType) do
    if QICharsetTypeNames[I] <> '' then
      List.AddObject(QICharsetTypeNames[I], TObject(I));
  if List.InheritsFrom(TStringList) then
    TStringList(List).Sort;
end;

function GetColIdFromColIndex(ColIndex: string): Integer;
begin
  Result := 0;
  if ColIndex <> '' then
    case Length(ColIndex) of
      1: Result := Ord(ColIndex[1]) - 64;
      2: Result := (Ord(ColIndex[1]) - 64)*26 + (Ord(ColIndex[2]) - 64);
      3: Result := (Ord(ColIndex[1]) - 64)*676 + ((Ord(ColIndex[2]) - 64)*26 + Ord(ColIndex[3]) - 64);
    end;
end;

procedure GetMonthDayNames;
var
  I: Integer;
  DefaultLCID: LCID;

  function LocalGetLocaleStr(LocaleType, Index: Integer;
    const DefValues: array of string): string;
  begin
    Result := GetLocaleStr(DefaultLCID, LocaleType, '');
    if Result = '' then
      Result := DefValues[Index];
  end;

begin
  DefaultLCID := GetThreadLocale;
  for I := 1 to 12 do
  begin
    {$IFDEF VCL17}FormatSettings.{$ENDIF}ShortMonthNames[I] :=
      LocalGetLocaleStr(LOCALE_SABBREVMONTHNAME1 + I - 1,
      I - Low(DefShortMonthNames), DefShortMonthNames);
    {$IFDEF VCL17}FormatSettings.{$ENDIF}LongMonthNames[I] :=
      LocalGetLocaleStr(LOCALE_SMONTHNAME1 + I - 1,
      I - Low(DefLongMonthNames), DefLongMonthNames);
  end;
end;

function StringToChar(const S: AnsiString; const Default: AnsiChar): AnsiChar;
var
  i: integer;
  tmp: AnsiString;
begin
  Result := Default;
  if Length(S) = 1 then Result := S[1]
  else if (Length(S) > 1) and (S[1] = '#') then
  begin
    i := -1;
    tmp := Copy(S, 2, Length(S) - 1);
    try
     i := StrToInt( {$IFDEF VCL12}string{$ENDIF}(tmp));
    except
    end;
    if i > -1 then Result := AnsiChar(Chr(i));
  end;
end;

function PadString(const S: AnsiString; Chr: AnsiChar; Len: integer): AnsiString;
var
  i: integer;
begin
  Result := S;
  if Length(S) >= Len then Exit;
  for i := Length(S) to Len do
    Result := Result + Chr;
end;

function Char2Str(Chr: Char): string;
begin
  if QImport3Common.CharInSet(Chr,[#33..#127])
    then Result := Chr
    else Result := Format('#%d', [Ord(Chr)]);
end;

{$IFDEF VCL12}
function Char2Str(Chr: AnsiChar): string;
begin
  if QImport3Common.CharInSet(Chr,[#33..#127])
    then Result := string(Chr)
    else Result := Format('#%d', [Ord(Chr)]);
end;
{$ENDIF}

function Str2Char(const Str: string; Default: Char): Char;
begin
  Result := Default;
  if Str <> EmptyStr then
  begin
    if Length(Str) = 1 then Result := Str[1]
    else if Str[1] = '#'
      then Result := Chr(StrToIntDef(Copy(Str, 2, Length(Str)), 0));
  end;
end;

{$IFDEF VCL12}
function Str2Char(const Str: string; Default: AnsiChar): AnsiChar;
begin
  Result := Default;
  if Str <> EmptyStr then
  begin
    if Length(Str) = 1 then Result := AnsiChar(Str[1])
    else if Str[1] = '#'
      then Result := AnsiChar(Chr(StrToIntDef(Copy(Str, 2, Length(Str)), 0)));
  end;
end;
{$ENDIF}

{$IFNDEF NOGUI}
procedure GridDrawCell(Grid: TqiStringGrid;
  const SheetName: string; SheetNumber, ACol, ARow: integer; Rect: TRect;
  State: TGridDrawState; DefinedRanges: TMapRow; SkipCols, SkipRows: integer;
  IsEditing: boolean; Selection: TMapRow);
var
  X, Y, i: integer;
  CellNeighbours: TCellNeighbours;
begin
  if gdFixed in State then
  begin
    X := Rect.Left + (Rect.Right - Rect.Left - Grid.Canvas.TextWidth(Grid.Cells[ACol, ARow])) div 2;
    Y := Rect.Top + (Rect.Bottom - Rect.Top - Grid.Canvas.TextHeight(Grid.Cells[ACol, ARow])) div 2;
    if ((ACol = Grid.Col) and (ARow = 0)) or
       ((ARow = Grid.Row) and (ACol = 0)) then
    begin
      Grid.Canvas.Font.Style := Grid.Canvas.Font.Style + [fsBold];
      Grid.Canvas.Font.Color := clHighlightText;
      Grid.Canvas.Brush.Color := clBtnShadow;
    end
    else begin
      Grid.Canvas.Font.Style := Grid.Canvas.Font.Style - [fsBold];
      Grid.Canvas.Font.Color := clWindowText;
      Grid.Canvas.Brush.Color := clBtnFace;
    end;
    Grid.Canvas.FillRect(Rect);
    Grid.Canvas.TextOut(X, Y + 1, Grid.Cells[ACol, ARow]);
    Exit;
  end;
  if (gdSelected in State) and not (gdFocused in State) then
  begin
    Grid.DefaultDrawing := false;
    Grid.Canvas.Brush.Color := clWindow;
    Grid.Canvas.FillRect(Rect);
    Grid.Canvas.Font.Color := clWindowText;
    Grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      Grid.Cells[ACol, ARow]);
    DrawFocusRect(Grid.Canvas.Handle, Rect);
  end;
  if not IsEditing then
  begin
    for i := 0 to DefinedRanges.Count - 1 do
    begin
      if not (gdFixed in State) and (DefinedRanges.Count > 0) and Assigned(Grid.Parent) and
        CellInRange(DefinedRanges[i], SheetName, SheetNumber, ACol, ARow) and
         {((Range <> Range.MapRow[0]) or}
          (((SkipRows = 0) or (ARow > SkipRows)) and
           ((SkipCols = 0) or (ACol > SkipCols))){)} then
      begin
        Grid.Canvas.Brush.Color := 12639424;
        Grid.Canvas.FillRect(Rect);
        Grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, Grid.Cells[ACol, ARow]);
        if (gdSelected in State) and not (gdFocused in State) then
        begin
          DrawFocusRect(Grid.Canvas.Handle, Rect);
        end;

        Grid.Canvas.Pen.Color := clGreen;
        Grid.Canvas.Pen.Width := 3;

        CellNeighbours := GetCellNeighbours(DefinedRanges, SheetName,
          SheetNumber, ACol, ARow);

        case DefinedRanges[i].RangeType of
          rtCol: begin
            if not (cnLeft in CellNeighbours) then
            begin
              Grid.Canvas.MoveTo(Rect.Left, Rect.Top);
              Grid.Canvas.LineTo(Rect.Left, Rect.Bottom);
            end;
            if not (cnRight in CellNeighbours) then
            begin
              Grid.Canvas.MoveTo(Rect.Right - 1, Rect.Top);
              Grid.Canvas.LineTo(Rect.Right - 1, Rect.Bottom);
            end;
            if ((DefinedRanges[i].Direction = rdDown) and (DefinedRanges[i].Col1 > 0) and
                (DefinedRanges[i].Row1 > 0) and (DefinedRanges[i].Col1 = ACol) and
                (DefinedRanges[i].Row1 = ARow)) or
               ((DefinedRanges[i].Direction = rdUp) and (DefinedRanges[i].Col2 > 0) and
                (DefinedRanges[i].Row2 > 0) and (DefinedRanges[i].Col2 = ACol) and
                (DefinedRanges[i].Row2 = ARow)) then
                begin
              if not (cnTop in CellNeighbours) then
              begin
                Grid.Canvas.MoveTo(Rect.Left, Rect.Top);
                Grid.Canvas.LineTo(Rect.Right, Rect.Top);
              end;
            end
            else if ((DefinedRanges[i].Direction = rdDown) and (DefinedRanges[i].Col2 > 0) and
                     (DefinedRanges[i].Row2 > 0) and (DefinedRanges[i].Col2 = ACol) and
                     (DefinedRanges[i].Row2 = ARow)) or
                    ((DefinedRanges[i].Direction = rdUp) and (DefinedRanges[i].Col1 > 0) and
                     (DefinedRanges[i].Row1 > 0) and (DefinedRanges[i].Col1 = ACol) and
                     (DefinedRanges[i].Row1 = ARow)) then
            begin
              if not (cnBottom in CellNeighbours) then
              begin
                Grid.Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
                Grid.Canvas.LineTo(Rect.Right, Rect.Bottom - 1);
              end;
            end;
          end;
          rtRow: begin
            if not (cnTop in CellNeighbours) then
            begin
              Grid.Canvas.MoveTo(Rect.Left, Rect.Top);
              Grid.Canvas.LineTo(Rect.Right, Rect.Top);
            end;
            if not (cnBottom in CellNeighbours) then
            begin
              Grid.Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
              Grid.Canvas.LineTo(Rect.Right, Rect.Bottom - 1);
            end;
            if ((DefinedRanges[i].Direction = rdDown) and (DefinedRanges[i].Col1 > 0) and
                (DefinedRanges[i].Row1 > 0) and (DefinedRanges[i].Col1 = ACol) and
                (DefinedRanges[i].Row1 = ARow)) or
               ((DefinedRanges[i].Direction = rdUp) and (DefinedRanges[i].Col2 > 0) and
                (DefinedRanges[i].Row2 > 0) and (DefinedRanges[i].Col2 = ACol) and
                (DefinedRanges[i].Row2 = ARow)) then
            begin
              if not (cnLeft in CellNeighbours) then
              begin
                Grid.Canvas.MoveTo(Rect.Left, Rect.Top);
                Grid.Canvas.LineTo(Rect.Left, Rect.Bottom);
              end;
            end
            else if ((DefinedRanges[i].Direction = rdDown) and (DefinedRanges[i].Col2 > 0) and
                     (DefinedRanges[i].Row2 > 0) and (DefinedRanges[i].Col2 = ACol) and
                     (DefinedRanges[i].Row2 = ARow)) or
                    ((DefinedRanges[i].Direction = rdUp) and (DefinedRanges[i].Col1 > 0) and
                     (DefinedRanges[i].Row1 > 0) and (DefinedRanges[i].Col1 = ACol) and
                     (DefinedRanges[i].Row1 = ARow)) then
            begin
              if not (cnRight in CellNeighbours) then
              begin
                Grid.Canvas.MoveTo(Rect.Right - 1, Rect.Top);
                Grid.Canvas.LineTo(Rect.Right - 1, Rect.Bottom);
              end;
            end;
          end;
          rtCell: begin
            if not (cnTop in CellNeighbours) then
            begin
              Grid.Canvas.MoveTo(Rect.Left, Rect.Top);
              Grid.Canvas.LineTo(Rect.Right, Rect.Top);
            end;
            if not (cnBottom in CellNeighbours) then
            begin
              Grid.Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
              Grid.Canvas.LineTo(Rect.Right, Rect.Bottom - 1);
            end;
            if not (cnLeft in CellNeighbours) then
            begin
              Grid.Canvas.MoveTo(Rect.Left, Rect.Top);
              Grid.Canvas.LineTo(Rect.Left, Rect.Bottom);
            end;
            if not (cnRight in CellNeighbours) then
            begin
              Grid.Canvas.MoveTo(Rect.Right - 1, Rect.Top);
              Grid.Canvas.LineTo(Rect.Right - 1, Rect.Bottom);
            end;
          end;
        end;
      end;
    end;
  end;
  if IsEditing and CellInRow(Selection, SheetName, SheetNumber, ACol, ARow) then
  begin
    Grid.Canvas.Brush.Color := 16776176;
    Grid.Canvas.FillRect(Rect);
    Grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, Grid.Cells[ACol, ARow]);

    CellNeighbours := GetCellNeighbours(Selection, SheetName, SheetNumber,
      ACol, ARow);
    Grid.Canvas.Pen.Color := clBlue;
    Grid.Canvas.Pen.Width := 3;

    if not (cnLeft in CellNeighbours) then
    begin
      Grid.Canvas.MoveTo(Rect.Left, Rect.Top);
      Grid.Canvas.LineTo(Rect.Left, Rect.Bottom);
    end;
    if not (cnRight in CellNeighbours) then
    begin
      Grid.Canvas.MoveTo(Rect.Right - 1, Rect.Top);
      Grid.Canvas.LineTo(Rect.Right - 1, Rect.Bottom);
    end;
    if not (cnTop in CellNeighbours) then
    begin
      Grid.Canvas.MoveTo(Rect.Left , Rect.Top);
      Grid.Canvas.LineTo(Rect.Right, Rect.Top);
    end;
    if not (cnBottom in CellNeighbours) then
    begin
      Grid.Canvas.MoveTo(Rect.Left, Rect.Bottom - 1);
      Grid.Canvas.LineTo(Rect.Right, Rect.Bottom - 1);
    end;
  end;
  Grid.DefaultDrawing := true;
end;

procedure GridMoveCurrCell(Grid: TqiStringGrid;
  ACol, ARow: integer);
var
  OnSelectCell: TSelectCellEvent;
begin
  OnSelectCell := Grid.OnSelectCell;
  Grid.OnSelectCell := nil;
  try
    Grid.Col := ACol;
    Grid.Row := ARow;
  finally
    Grid.OnSelectCell := OnSelectCell;
  end;
end;

procedure GridSelCell(Grid: TqiStringGrid;
  RowNo, ColNo: integer; NeedClear: boolean; var CurrSel: string);
begin
  if NeedClear then CurrSel := EmptyStr;
  CurrSel := CurrSel + Col2Letter(ColNo) + Row2Number(RowNo) + ';';
  GridMoveCurrCell(Grid, ColNo, RowNo);
end;

procedure GridSelRow(Grid: TqiStringGrid;
  FirstCol, LastCol, RowNo, StartCol, SkipRows, SkipCols: integer;
  NeedClear: boolean; var CurrSel: string);
begin
  if NeedClear then CurrSel := EmptyStr;
  CurrSel := CurrSel + Col2Letter(FirstCol) + Row2Number(RowNo) + '-' +
    Col2Letter(LastCol) + Row2Number(RowNo);
  GridMoveCurrCell(Grid, FirstCol + SkipCols, RowNo);
end;

procedure GridSelCol(Grid: TqiStringGrid;
  FirstRow, LastRow, ColNo, StartRow, SkipRows, SkipCols: integer;
  NeedClear: boolean; var CurrSel: string);
begin
  if NeedClear then CurrSel := EmptyStr;
  CurrSel := CurrSel + Col2Letter(ColNo) + Row2Number(FirstRow) + '-' +
    Col2Letter(ColNo) + Row2Number(LastRow);
  GridMoveCurrCell(Grid, ColNo, FirstRow + SkipRows);
end;

procedure GridMouseDown(Grid: TqiStringGrid;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer; FirstCol, LastCol, FirstRow, LastRow,
  SkipRows, SkipCols: integer; var CurrSel: string);
var
  ARow, ACol, i: integer;
  OnDrawCell: TDrawCellEvent;
begin
  Grid.MouseToCell(X, Y, ACol, ARow);
  // simple select col or row
  if Shift = [ssLeft] then
  begin
    if (ACol = 0) and (ARow > 0) then
      GridSelRow(Grid, FirstCol, LastCol, ARow, 0,
        SkipRows, SkipCols, true, CurrSel)
    else if (ARow = 0) and (ACol > 0) then
      GridSelCol(Grid, FirstRow, LastRow, ACol, 0,
        SkipRows, SkipCols, true, CurrSel);
  end;
  // ctrl select col or row
  if Shift = [ssLeft, ssCtrl] then
  begin
    if (ACol = 0) and (ARow > 0) then
      GridSelRow(Grid, FirstCol, LastCol, ARow, 0,
        SkipRows, SkipCols, false, CurrSel)
    else if (ARow = 0) and (ACol > 0) then
      GridSelCol(Grid, FirstRow, LastRow, ACol, 0,
        SkipRows, SkipCols, false, CurrSel);
  end;
  // shift select col or row
  if Shift = [ssLeft, ssShift] then
  begin
    OnDrawCell := Grid.OnDrawCell;
    Grid.OnDrawCell := nil;
    try
      if (ACol = 0) {(ACol = StartCol)} and (ARow > 0) then
        if Grid.Row <= ARow then
          for i := Grid.Row to ARow do
            GridSelRow(Grid, FirstCol, LastCol, i, 0, SkipRows, SkipCols,
              false, CurrSel)
        else for i := Grid.Row downto ARow do
          GridSelRow(Grid, FirstCol, LastCol, i, 0, SkipRows, SkipCols,
            false, CurrSel)
      else if (ARow = 0){(ARow = StartRow)} and (ACol > 0) then
        if Grid.Col <= ACol then
          for i := Grid.Col to ACol do
            GridSelCol(grid, FirstRow, LastRow, i, 0, SkipRows, SkipCols,
              false, CurrSel)
        else for i := Grid.Col downto ACol do
          GridSelCol(Grid, FirstRow, LastRow, i, 0, SkipRows, SkipCols,
            false, CurrSel)
    finally
      Grid.OnDrawCell := OnDrawCell;
      Grid.Repaint;
    end;
  end;
  Grid.Repaint;
end;

procedure GridSelectCell(Grid: TqiStringGrid;
  ACol, ARow: integer; Shift: TShiftState; var CurrSel: string);
var
  i: integer;
begin
  if Shift = [] then CurrSel := EmptyStr;
  // select simple cell
  if Shift = [ssCtrl] then
  begin
{      i := FCells.IndexOf(ACol, ARow);
      while i > -1 do
      begin
        FCells.Delete(i);
        editCells.Text := FCells.OptimalString;
        if FCells.Count > 0 then FCells.Sort(0, FCells.Count - 1, CompareByColRow);
        xlsGrid.Repaint;
        i := FCells.IndexOf(ACol, ARow);
        if i = -1 then Exit;
      end;}
    GridSelCell(Grid, ARow, ACol, false, CurrSel);
  end;
  // select range of cells
  if Shift = [ssShift] then
  begin
    if (Grid.Col = ACol) and (Grid.Row <> ARow) then
      if Grid.Row < ARow then
        for i := Grid.Row to ARow do GridSelCell(Grid, i, ACol, false, CurrSel)
      else for i := Grid.Row downto ARow do
             GridSelCell(Grid, i, ACol, false, CurrSel);

    if (Grid.Row = ARow) and (Grid.Col <> ACol) then
      if Grid.Col < ACol then
        for i := Grid.Col to ACol do GridSelCell(Grid, ARow, i, false, CurrSel)
      else for i := Grid.Col downto ACol do
             GridSelCell(Grid, ARow, i, false, CurrSel);
  end;
end;

procedure GridFillFixedCells(Grid: TqiStringGrid);
var
  i: integer;
begin
  Grid.ColWidths[0] := 25;
  for i := 0 to Grid.ColCount - 2 do
    Grid.Cells[i + 1, 0] := Col2Letter(i + 1);
  for i := 1 to Grid.RowCount - 1 do
    Grid.Cells[0, i] := Row2Number(i);
end;

{procedure GridClearDataCell(Grid: TStringGrid);
var
  i, j: integer;
begin
  for i := 1 to Grid.ColCount - 1 do
    for j := 1 to Grid.RowCount - 1 do
      Grid.Cells[i, j] := QIGetEmptyStr;
end;}
{$ENDIF}

{$IFDEF QI_UNICODE}
procedure CSVStringToStrings(
  const Str: WideString; Quote, Separator: AnsiChar; AStrings: TWideStrings);
{$ELSE}
procedure CSVStringToStrings(
  const Str: AnsiString; Quote, Separator: AnsiChar; AStrings: TStrings);
{$ENDIF}
var
  i: Integer;
  is_quote, in_quote, is_first, is_separator, quote_in_quote, wait_separator: Boolean;
  sss: qiString;
begin
  if Separator = #0 then Exit;

  AStrings.Clear;

  is_first := true;
  in_quote := false;
  quote_in_quote := false;
  wait_separator := false;

  for i := 1 to Length(Str) do
  begin
    if is_first then
    begin
      sss := '';
      in_quote := (Quote <> #0) and (Str[i] = qiChar(Quote));
      is_first := false;
      if in_quote then Continue;
    end;

    is_separator := Str[i] = qiChar(Separator);
    if is_separator and not in_quote then
    begin
      AStrings.Add(sss);
      is_first := true;
      sss := '';
      Continue;
    end;

    if wait_separator then
      if is_separator then
      begin
        wait_separator := false;
        is_first := true;
        AStrings.Add(sss);
        Continue;
      end
      else begin
        Continue;
      end;

    is_quote := (Quote <> #0) and (Str[i] = qiChar(Quote));

    if quote_in_quote then
    begin
      if is_quote then
      begin
        sss := sss + qiChar(Quote);
        quote_in_quote := false;
        Continue;
      end
      else if is_separator then
      begin
        quote_in_quote := false;
        in_quote := false;
        is_first := true;
        AStrings.Add(sss);
        sss := '';
        Continue;
      end
      else begin
        wait_separator := true;
        quote_in_quote := false;
        Continue;
      end;
    end;

    if is_quote and quote_in_quote then
    begin
      sss := sss + qiChar(Quote);
      Continue;
    end;

    if is_quote and in_quote then
    begin
      quote_in_quote := True;
      Continue;
    end;

    sss := sss + Str[i];
  end;
  AStrings.Add(sss);
end;

procedure ReplaceTabs(var Str: qiString);
var
  sp: WideString;
  i, d, r: Integer;
begin
  i := 1;
  while i <= Length(Str) do
  begin
    if Str[i] = #9 then
    begin
      if (i mod 8) = 0  then
        r := i mod 9
      else
        r := i mod 8;
      d := 9 - r;
      sp := StringOfChar(' ', d);
      QIDelete(Str, i, 1);
      QIInsert(sp, Str, i);
      i := i + d;
      Continue;
    end;
    Inc(i);
  end;
end;

function GetListSeparator: Char;
begin
  Result := GetLocaleChar(GetThreadLocale, LOCALE_SLIST, ',');
end;

procedure ClearIniFile(IniFile: TIniFile);
var
  AStrings: TStrings;
  i: integer;
begin
  AStrings := TStringList.Create;
  try
    IniFile.ReadSections(AStrings);
    for i := 0 to AStrings.Count - 1 do
      IniFile.EraseSection(AStrings[i]);
  finally
    AStrings.Free;
  end;
end;

// ImportDestination routines
procedure QImportCheckDestination(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF});
begin
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet:
      if not Assigned(DataSet) then
        raise EQImportError.Create(QImportLoadStr(QIE_NoDataSet));
{$IFNDEF NOGUI}
    qidDBGrid:
      if not Assigned(DBGrid) then
        raise EQImportError.Create(QImportLoadStr(QIE_NoDBGrid));
    qidListView:
      if not Assigned(ListView) then
        raise EQImportError.Create(QImportLoadStr(QIE_NoListView));
    qidStringGrid:
      if not Assigned(StringGrid) then
        raise EQImportError.Create(QImportLoadStr(QIE_NoStringGrid));
{$ENDIF}
  end;
end;

function QImportIsDestinationActive(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF}): boolean;
begin
  Result := false;
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: Result := DataSet.Active;
    {$IFNDEF NOGUI}
    qidDBGrid:
      Result := Assigned(DBGrid.DataSource) and Assigned(DBGrid.DataSource.DataSet)and
        DBGrid.DataSource.DataSet.Active;
    else Result := true;
    {$ENDIF}
  end;
end;

procedure QImportIsDestinationOpen(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF});
begin
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: DataSet.Open;
    {$IFNDEF NOGUI}
    qidDBGrid:
      if Assigned(DBGrid.DataSource) and
         Assigned(DBGrid.DataSource.DataSet) then
        DBGrid.DataSource.DataSet.Open;
    {$ENDIF}
  end;
end;

procedure QImportIsDestinationClose(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF});
begin
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: DataSet.Close;
    {$IFNDEF NOGUI}
    qidDBGrid:
      if Assigned(DBGrid.DataSource) and
         Assigned(DBGrid.DataSource.DataSet) then
        DBGrid.DataSource.DataSet.Close;
    {$ENDIF}
  end;
end;

function QImportDestinationColCount(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF}): integer;
begin
  Result := 0;
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: Result := DataSet.FieldCount;
    {$IFNDEF NOGUI}
    qidDBGrid: Result := DBGrid.Columns.Count;
    qidListView: Result := ListView.Columns.Count;
    qidStringGrid: Result := StringGrid.ColCount;
    {$ENDIF}
  end;
end;

function QImportDestinationColName(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet;
  {$IFNDEF NOGUI}DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid;
 GridCaptionRow,{$ENDIF} Index: integer): string;
begin
  Result := EmptyStr;
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: Result := DataSet.Fields[Index].FieldName;
    {$IFNDEF NOGUI}
    qidDBGrid: Result := DBGrid.Columns[Index].Title.Caption;
    qidListView: Result := ListView.Columns[Index].Caption;
    qidStringGrid:
      if (GridCaptionRow > -1) and (GridCaptionRow < StringGrid.RowCount) then
        Result := StringGrid.Cells[Index, GridCaptionRow]
      else Result := IntToStr(Index);
    {$ENDIF}
  end;
end;

function QImportDestinationAssigned(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet
  {$IFNDEF NOGUI}; DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid{$ENDIF}): boolean;
begin
  Result := false;
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet:    Result := Assigned(DataSet);
    {$IFNDEF NOGUI}
    qidDBGrid:     Result := Assigned(DBGrid);
    qidListView:   Result := Assigned(ListView);
    qidStringGrid: Result := Assigned(StringGrid);
    {$ENDIF}
  end;
end;

function QImportDestinationFindColumn(IsCSV: boolean;
  ImportDestination: TQImportDestination; DataSet: TDataSet;
  {$IFNDEF NOGUI}DBGrid: TDBGrid;
 ListView: TListView;
  StringGrid: TStringGrid; GridCaptionRow: integer;
{$ENDIF}
  const ColName: string): integer;
var
  Field: TField;
{$IFNDEF NOGUI}
  i: integer;
{$ENDIF}
begin
  Result := -1;
  if IsCSV then Exit;
  case ImportDestination of
    qidDataSet: begin
      Field := DataSet.FindField(ColName);
      if Assigned(Field)  then
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

function IncludePathDelimiter(const S: string): string;
begin
  Result := S;
  if not IsPathDelimiter(Result, Length(Result)) then
    Result := Result + PathDelim;
end;

function GetNumericString(const Str: string): string;
var
  p: integer;
begin
  Result := Str;
  p := Pos({$IFDEF VCL17}FormatSettings.{$ENDIF}ThousandSeparator, Result);
  while p > 0 do
  begin
    Delete(Result, p, 1);
    p := Pos({$IFDEF VCL17}FormatSettings.{$ENDIF}ThousandSeparator, Result);
  end;
end;

function FormatStrToDateTime(ADateTimeStr, AFormatStr: string): TDateTime;

  function NormalizeSubString(const ASubStr: string): string;
  var
    i: Integer;
  begin
    Result := '';
    for i := 1 to Length(ASubStr) do
    begin
      if StrToIntDef(ASubStr[i], -1) <> -1 then
        Result := Result + ASubStr[i];
    end;
  end;

var
  i, j: Integer;
  st: TSystemTime;
  fChar, subStr: string;
begin
  Result := 0;

  if (Length(ADateTimeStr) = 0) or (Length(AFormatStr) < 5) then Exit;

  AFormatStr := UpperCase(AFormatStr);

  st.wHour := 0;
  st.wMinute := 0;
  st.wSecond := 0;
  st.wMilliSeconds := 0;

  while Length(AFormatStr) > 0 do
  begin
    fChar := Copy(AFormatStr, 1, 1);

    if fChar = 'Y' then
    begin
      i := Pos('YYYY', AFormatStr);
      if (i > 0) and (i < Length(ADateTimeStr)) then
      begin
        subStr := Copy(ADateTimeStr, i, 4);
        st.wYear := StrToIntDef(subStr, 0);
        if st.wYear = 0 then
        begin
          subStr := NormalizeSubString(subStr);
          st.wYear := StrToIntDef(subStr, 0);
          if st.wYear = 0 then Exit;
        end;
        if (Length(ADateTimeStr) < Length(subStr)) or
          (Length(AFormatStr) < 4) then Exit;
        Delete(ADateTimeStr, 1, Length(subStr));
        Delete(AFormatStr, 1, 4);
      end
      else begin
        i := Pos('YY', AFormatStr);
        if (i > 0) and (i < Length(ADateTimeStr)) then
        begin
          subStr := Copy(ADateTimeStr, i, 2);
          st.wYear := StrToIntDef(subStr, 0);
          if st.wYear = 0 then
          begin
            subStr := NormalizeSubString(subStr);
            st.wYear := StrToIntDef(subStr, 0);
            if st.wYear = 0 then Exit;
          end;
          if st.wYear > 69 then
            st.wYear := 1900 + st.wYear
          else
            st.wYear := 2000 + st.wYear;
          if (Length(ADateTimeStr) < Length(subStr)) or
            (Length(AFormatStr) < 2) then Exit;
          Delete(ADateTimeStr, 1, Length(subStr));
          Delete(AFormatStr, 1, 2);
        end
        else begin
          i := Pos('Y', AFormatStr);
          if (i > 0) and (i <= Length(ADateTimeStr)) then
          begin
            subStr := Copy(ADateTimeStr, i, 1);
            st.wYear := StrToIntDef(subStr, 0);
            if st.wYear = 0 then
            begin
              subStr := NormalizeSubString(subStr);
              st.wYear := StrToIntDef(subStr, 0);
              if st.wYear = 0 then Exit;
            end;
            st.wYear := 2000 + st.wYear;
            if (Length(ADateTimeStr) < Length(subStr)) or
              (Length(AFormatStr) < 1) then Exit;
            Delete(ADateTimeStr, 1, Length(subStr));
            Delete(AFormatStr, 1, 1);
          end
          else
            Exit;
        end;
      end;
    end
    else if fChar = 'M' then
    begin
      i := Pos('MMMM', AFormatStr);
      if (i > 0) and (i < Length(ADateTimeStr)) then
      begin
        for j := 1 to 12 do
        begin
          if Pos(UpperCase(DefLongMonthNames[j]), UpperCase(ADateTimeStr)) > 0 then
          begin
            st.wMonth := j;
            if (Length(ADateTimeStr) < Length(DefLongMonthNames[j])) or
              (Length(AFormatStr) < 4) then Exit;
            Delete(ADateTimeStr, 1, Length(DefLongMonthNames[j]));
            Delete(AFormatStr, 1, 4);
            Break
          end
          else if j = 12 then
            Exit;
        end;
      end
      else begin
        i := Pos('MMM', AFormatStr);
        if (i > 0) and (i < Length(ADateTimeStr)) then
        begin
          for j := 1 to 12 do
          begin
            if Pos(UpperCase(DefShortMonthNames[j]), UpperCase(ADateTimeStr)) > 0 then
            begin
              st.wMonth := j;
              if (Length(ADateTimeStr) < Length(DefShortMonthNames[j])) or
                (Length(AFormatStr) < 3) then Exit;
              Delete(ADateTimeStr, 1, Length(DefShortMonthNames[j]));
              Delete(AFormatStr, 1, 3);
              Break
            end
            else if j = 12 then
              Exit;
          end;
        end
        else begin
          i := Pos('MM', AFormatStr);
          if (i > 0) and (i < Length(ADateTimeStr)) then
          begin
            subStr := Copy(ADateTimeStr, i, 2);
            st.wMonth := StrToIntDef(subStr, 0);
            if st.wMonth = 0 then
            begin
              subStr := NormalizeSubString(subStr);
              st.wMonth := StrToIntDef(subStr, 0);
              if st.wMonth = 0 then Exit;
            end;
            if (Length(ADateTimeStr) < Length(subStr)) or
              (Length(AFormatStr) < 2) then Exit;
            Delete(ADateTimeStr, 1, Length(subStr));
            Delete(AFormatStr, 1, 2)
          end
          else begin
            i := Pos('M', AFormatStr);
            if (i > 0) and (i <= Length(ADateTimeStr)) then
            begin
              subStr := Copy(ADateTimeStr, i, 1);
              st.wMonth := StrToIntDef(subStr, 0);
              if st.wMonth = 0 then
              begin
                subStr := NormalizeSubString(subStr);
                st.wMonth := StrToIntDef(subStr, 0);
                if st.wMonth = 0 then Exit;
              end;
              if (Length(ADateTimeStr) < Length(subStr)) or
                (Length(AFormatStr) < 1) then Exit;
              Delete(ADateTimeStr, 1, Length(subStr));
              Delete(AFormatStr, 1, 1)
            end
            else
              Exit;
          end;
        end;
      end;
    end
    else if fChar = 'D' then
    begin
      i := Pos('DD', AFormatStr);
      if (i > 0) and (i < Length(ADateTimeStr)) then
      begin
        subStr := Copy(ADateTimeStr, i, 2);
        st.wDay := StrToIntDef(subStr, 0);
        if st.wDay = 0 then
        begin
          subStr := NormalizeSubString(subStr);
          st.wDay := StrToIntDef(subStr, 0);
          if st.wDay = 0 then Exit;
        end;
        if (Length(ADateTimeStr) < Length(subStr)) or
          (Length(AFormatStr) < 2) then Exit;
        Delete(ADateTimeStr, 1, Length(subStr));
        Delete(AFormatStr, 1, 2);
      end
      else begin
        i := Pos('D', AFormatStr);
        if (i > 0) and (i <= Length(ADateTimeStr)) then
        begin
          subStr := Copy(ADateTimeStr, i, 1);
          st.wDay := StrToIntDef(subStr, 0);
          if st.wDay = 0 then
          begin
            subStr := NormalizeSubString(subStr);
            st.wDay := StrToIntDef(subStr, 0);
            if st.wDay = 0 then Exit;
          end;
          if (Length(ADateTimeStr) < Length(subStr)) or
            (Length(AFormatStr) < 1) then Exit;
          Delete(ADateTimeStr, 1, Length(subStr));
          Delete(AFormatStr, 1, 1);
        end
        else
          Exit;
      end;
    end
    else if fChar = 'H' then
    begin
      i := Pos('HH', AFormatStr);
      if (i > 0) and (i < Length(ADateTimeStr)) then
      begin
        subStr := Copy(ADateTimeStr, i, 2);
        st.wHour := StrToIntDef(subStr, 0);
        if st.wHour = 0 then
        begin
          subStr := NormalizeSubString(subStr);
          if StrToIntDef(subStr, -1) = -1 then Exit;
          st.wHour := StrToIntDef(subStr, 0);
        end;
        if (Length(ADateTimeStr) < Length(subStr)) or
          (Length(AFormatStr) < 2) then Exit;
        Delete(ADateTimeStr, 1, Length(subStr));
        Delete(AFormatStr, 1, 2);
      end
      else begin
        i := Pos('H', AFormatStr);
        if (i > 0) and (i <= Length(ADateTimeStr)) then
        begin
          subStr := Copy(ADateTimeStr, i, 1);
          st.wHour := StrToIntDef(subStr, 0);
          if st.wHour = 0 then
          begin
            subStr := NormalizeSubString(subStr);
            if StrToIntDef(subStr, -1) = -1 then Exit;
            st.wHour := StrToIntDef(subStr, 0);
          end;
          if (Length(ADateTimeStr) < Length(subStr)) or
            (Length(AFormatStr) < 1) then Exit;
          Delete(ADateTimeStr, 1, Length(subStr));
          Delete(AFormatStr, 1, 1);
        end
        else
          Exit;
      end;
    end
    else if fChar = 'N' then
    begin
      i := Pos('NN', AFormatStr);
      if (i > 0) and (i < Length(ADateTimeStr)) then
      begin
        subStr := Copy(ADateTimeStr, i, 2);
        st.wMinute := StrToIntDef(subStr, 0);
        if st.wMinute = 0 then
        begin
          subStr := NormalizeSubString(subStr);
          if StrToIntDef(subStr, -1) = -1 then Exit;
          st.wMinute := StrToIntDef(subStr, 0);
        end;
        if (Length(ADateTimeStr) < Length(subStr)) or
          (Length(AFormatStr) < 2) then Exit;
        Delete(ADateTimeStr, 1, Length(subStr));
        Delete(AFormatStr, 1, 2);
      end
      else begin
        i := Pos('N', AFormatStr);
        if (i > 0) and (i <= Length(ADateTimeStr)) then
        begin
          subStr := Copy(ADateTimeStr, i, 1);
          st.wMinute := StrToIntDef(subStr, 0);
          if st.wMinute = 0 then
          begin
            subStr := NormalizeSubString(subStr);
            if StrToIntDef(subStr, -1) = -1 then Exit;
            st.wMinute := StrToIntDef(subStr, 0);
          end;
          if (Length(ADateTimeStr) < Length(subStr)) or
            (Length(AFormatStr) < 1) then Exit;
          Delete(ADateTimeStr, 1, Length(subStr));
          Delete(AFormatStr, 1, 1);
        end
        else
          Exit;
      end;
    end
    else if fChar = 'S' then
    begin
      i := Pos('SS', AFormatStr);
      if (i > 0) and (i < Length(ADateTimeStr)) then
      begin
        subStr := Copy(ADateTimeStr, i, 2);
        st.wSecond := StrToIntDef(subStr, 0);
        if st.wSecond = 0 then
        begin
          subStr := NormalizeSubString(subStr);
          if StrToIntDef(subStr, -1) = -1 then Exit;
          st.wSecond := StrToIntDef(subStr, 0);
        end;
        if (Length(ADateTimeStr) < Length(subStr)) or
          (Length(AFormatStr) < 2) then Exit;
        Delete(ADateTimeStr, 1, Length(subStr));
        Delete(AFormatStr, 1, 2);
      end
      else begin
        i := Pos('S', AFormatStr);
        if (i > 0) and (i <= Length(ADateTimeStr)) then
        begin
          subStr := Copy(ADateTimeStr, i, 1);
          st.wSecond := StrToIntDef(subStr, 0);
          if st.wSecond = 0 then
          begin
            subStr := NormalizeSubString(subStr);
            if StrToIntDef(subStr, -1) = -1 then Exit;
            st.wSecond := StrToIntDef(subStr, 0);
          end;
          if (Length(ADateTimeStr) < Length(subStr)) or
            (Length(AFormatStr) < 1) then Exit;
          Delete(ADateTimeStr, 1, Length(subStr));
          Delete(AFormatStr, 1, 1);
        end
        else
          Exit;
      end;
    end
    else if fChar = 'Z' then
    begin
      i := Pos('ZZZ', AFormatStr);
      if (i > 0) and (i < Length(ADateTimeStr)) then
      begin
        subStr := Copy(ADateTimeStr, i, 3);
        st.wMilliseconds := StrToIntDef(subStr, 0);
        if st.wMilliseconds = 0 then     
        begin
          subStr := NormalizeSubString(subStr);
          if StrToIntDef(subStr, -1) = -1 then Exit;
          st.wMilliseconds := StrToIntDef(subStr, 0);
        end;
        if (Length(ADateTimeStr) < Length(subStr)) or
          (Length(AFormatStr) < 3) then Exit;
        Delete(ADateTimeStr, 1, Length(subStr));
        Delete(AFormatStr, 1, 3);
      end;
    end
    else if fChar = 'A' then
    begin
      i := Pos('AP', AFormatStr);
      if (i > 0) and (i < Length(ADateTimeStr)) then
      begin
        subStr := Copy(ADateTimeStr, i, 2);
        if (subStr[1] = 'P') or (subStr[1] = 'p') then
        begin
          if st.wHour < 12 then
            st.wHour := st.wHour + 12;
        end
        else if st.wHour = 12 then
          st.wHour := 0;
        if (Length(ADateTimeStr) < Length(subStr)) or
          (Length(AFormatStr) < 2) then Exit;
        Delete(ADateTimeStr, 1, Length(subStr));
        Delete(AFormatStr, 1, 2);
      end
      else
        Exit;
    end
    else begin
      if (Length(ADateTimeStr) < 1) or
        (Length(AFormatStr) < 1) then Exit;
      if AFormatStr[1] <> ADateTimeStr[1] then Exit;
      Delete(ADateTimeStr, 1, 1);
      Delete(AFormatStr, 1, 1);
    end;
  end;

  try
    Result := SystemTimeToDateTime(st);
  except
    Result := 0;
  end;
end;

function TryStrToDateTime(const Str: string; var DateTime: TDateTime): Boolean;
var
  DT: TDateTime;
  i: integer;
  s: string;
begin
  Result := True;
  DateTime := 0;
  try
    DT := StrToDateTime(Str);
  except
    DT := FormatStrToDateTime(Str, {$IFDEF VCL17}FormatSettings.{$ENDIF}ShortDateFormat);
    if DT = 0 then
    try
      i := Pos(Chr($A0), Str);
      if i > 0 then
      begin
        s := Str;
        s[i] := Chr($20);
        DT := StrToDateTime(s);
      end else
        raise Exception.Create(''); // next try
    except
      DT := FormatStrToDateTime(Str, {$IFDEF VCL17}FormatSettings.{$ENDIF}ShortDateFormat);
      if DT = 0 then
      try
        DT := StrToFloat(Str);
      except
        Result := False;
      end;
    end;
  end;
  if Result then
    DateTime := DT;
end;

{$IFNDEF VCL5}
function StringReplace(const S, OldPattern, NewPattern: AnsiString;
  Flags: TReplaceFlags): AnsiString;
var
  SearchStr, Patt, NewStr: AnsiString;
  Offset: Integer;
begin
  if rfIgnoreCase in Flags then
  begin
    SearchStr := AnsiUpperCase(S);
    Patt := AnsiUpperCase(OldPattern);
  end else
  begin
    SearchStr := S;
    Patt := OldPattern;
  end;
  NewStr := S;
  Result := '';
  while SearchStr <> '' do
  begin
    Offset := AnsiPos(Patt, SearchStr);
    if Offset = 0 then
    begin
      Result := Result + NewStr;
      Break;
    end;
    Result := Result + Copy(NewStr, 1, Offset - 1) + NewPattern;
    NewStr := Copy(NewStr, Offset + Length(OldPattern), MaxInt);
    if not (rfReplaceAll in Flags) then
    begin
      Result := Result + NewStr;
      Break;
    end;
    SearchStr := Copy(SearchStr, Offset + Length(Patt), MaxInt);
  end;
end;
{$ENDIF}

function DirExists(const Directory: string): Boolean;
var
  Code: Integer;
begin
  Code := GetFileAttributes(PChar(Directory));
  Result := (Code <> -1) and (FILE_ATTRIBUTE_DIRECTORY and Code <> 0);
end;

procedure CutFirstDirectory(var S: string);
var
  Root: Boolean;
  P: Integer;
begin
  if S = '\' then
    S := ''
  else
  begin
    if S[1] = '\' then
    begin
      Root := True;
      Delete(S, 1, 1);
    end
    else
      Root := False;
    if S[1] = '.' then
      Delete(S, 1, 4);
    P := Pos('\',S);
    if P <> 0 then
    begin
      Delete(S, 1, P);
      S := '...\' + S;
    end
    else
      S := '';
    if Root then
      S := '\' + S;
  end;
end;

function MinimizeName(const Filename: string; Canvas: TCanvas;
  MaxLen: Integer): string;
var
  Drive: string;
  Dir: string;
  Name: string;
begin
  Result := FileName;
  Dir := ExtractFilePath(Result);
  Name := ExtractFileName(Result);

  if (Length(Dir) >= 2) and (Dir[2] = ':') then
  begin
    Drive := Copy(Dir, 1, 2);
    Delete(Dir, 1, 2);
  end
  else
    Drive := '';
  while ((Dir <> '') or (Drive <> '')) and (Canvas.TextWidth(Result) > MaxLen) do
  begin
    if Dir = '\...\' then
    begin
      Drive := '';
      Dir := '...\';
    end
    else if Dir = '' then
      Drive := ''
    else
      CutFirstDirectory(Dir);
    Result := Drive + Dir + Name;
  end;
end;

{$IFDEF QI_UNICODE}
function QIAddQuote(const S, LeftQuote, RightQuote: WideString): WideString;
begin
  Result := LeftQuote + S + RightQuote;
end;

function QIRemoveQuote(const S, LeftQuote, RightQuote: WideString): WideString;
var
  l: Integer;
begin
  Result := S;
  l := Length(LeftQuote);
  if QICompareStr(Copy(Result, 1, l), LeftQuote) = 0 then
    QIDelete(Result, 1, l);
  l := Length(RightQuote);
  if QICompareStr(Copy(Result, Length(Result) - l + 1, l), RightQuote) = 0 then
    QIDelete(Result, Length(Result) - l + 1, l);
end;

function QIUpperFirst(const S: WideString): WideString;
begin
  if Length(S) <> 0 then
    Result := QIUpperCase(S[1]) + QILowerCase(Copy(S, 2, Length(S) - 1))
  else
    Result := S;
end;

function QIUpperFirstWord(const S: WideString): WideString;
var
  spaceFlag: Boolean;
  resultPtr: PWideChar;
begin
  Result := S;
  resultPtr := PWideChar(Result);
  spaceFlag := False;
  while lstrlenW(resultPtr) > 0 do
  begin
    if lstrlenW(resultPtr) = Length(Result) then
    begin
      if resultPtr^ <> ' ' then
        CharUpperBuffW(resultPtr, 1)
      else
        spaceFlag := True;
    end
    else begin
      if resultPtr^ = ' ' then
        spaceFlag := True
      else if spaceFlag then
      begin
        CharUpperBuffW(resultPtr, 1);
        spaceFlag := False;
      end;
    end;
    Inc(resultPtr);
  end;
end;
{$ELSE}
function QIAddQuote(const S, LeftQuote, RightQuote: AnsiString): AnsiString;
begin
  Result := LeftQuote + S + RightQuote;
end;

function QIRemoveQuote(const S, LeftQuote, RightQuote: AnsiString): AnsiString;
var
  l: Integer;
begin
  Result := S;
  l := Length(LeftQuote);
  if AnsiCompareStr(Copy(Result, 1, l), LeftQuote) = 0 then
    Delete(Result, 1, l);
  l := Length(RightQuote);
  if AnsiCompareStr(Copy(Result, Length(Result) - l + 1, l), RightQuote) = 0 then
    Delete(Result, Length(Result) - l + 1, l);
end;

function QIUpperFirst(const S: AnsiString): AnsiString;
begin
  if Length(S) <> 0 then
    Result := AnsiUpperCase(S[1]) + AnsiLowerCase(Copy(S, 2, Length(S) - 1))
  else
    Result := S;
end;

function QIUpperFirstWord(const S: AnsiString): AnsiString;
var
  spaceFlag: Boolean;
  resultPtr: PChar;
begin
  Result := S;
  resultPtr := PChar(Result);
  spaceFlag := False;
  while lstrlen(resultPtr) > 0 do
  begin
    if lstrlen(resultPtr) = Length(Result) then
    begin
      if resultPtr^ <> ' ' then
         CharUpperBuff(resultPtr, 1)
      else
        spaceFlag := True;
    end
    else begin
      if resultPtr^ = ' ' then
        spaceFlag := True
      else if spaceFlag then
      begin
        CharUpperBuff(resultPtr, 1);
        spaceFlag := False;
      end;
    end;
    Inc(resultPtr);
  end;
end;
{$ENDIF}

function GetFieldDataType(const Value: string): TFieldType;

  function IsInteger: boolean;                          
  begin
    Result := True;
    try
      StrToInt(Value);
    except
      Result := False;
    end;
  end;

  function IsFloat: boolean;
  begin
    Result := True;
    try
      StrToFloat(Value);
    except
      Result := False;
    end;
  end;

  function IsTime: boolean;
  begin
    Result := True;
    try
      StrToTime(Value);
    except
      Result := False;
    end;
  end;

  function IsDate: boolean;
  begin
    Result := True;
    try
      StrToDate(Value);
    except
      Result := False;
    end;
  end;

  function IsDateTime: boolean;
  begin
    Result := True;
    try
      StrToDateTime(Value);
    except
      Result := False;
    end;
  end;

begin
  if Value <> '' then
  begin
    if Length(Value) < 255 then
    begin
      if (AnsiLowerCase(Value) = 'true') or (AnsiLowerCase(Value) = 'false') then
        Result := ftBoolean
      else if IsInteger then
        Result := ftInteger
      else if IsFloat then
        Result := ftFloat
      else if IsTime then
        Result := ftTime
      else if IsDate then
        Result := ftDate
      else if IsDateTime then
        Result := ftDateTime
      else
        Result := ftString;
    end else
      Result := ftMemo;
  end else
    Result := ftString;
end;

function FileToBase64(AFileName: string): string;
var
  F: TFileStream;
  S: TStringStream;
begin
  F := TFileStream.Create(AFileName, fmOpenRead or fmShareDenyWrite);
  S := TStringStream.Create('');
  try
    F.Position := 0;
    S.Position := 0;
    S.CopyFrom(F, F.Size);
    Result := QIEncodeBase64(S.DataString);
  finally
    F.Free;
    S.Free;
  end;
end;


{$IFNDEF VCL6}

 function EncodeBase64(const inStr: string): string;

  function Encode_Byte(b: Byte): char;
  const
    Base64Code: string[64] =
      'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/';
  begin
    Result := Base64Code[(b and $3F)+1];
  end;

var
  i: Integer;
begin
  i := 1;
  Result := '';
  while i <= Length(InStr) do
  begin
    Result := Result + Encode_Byte(Byte(inStr[i]) shr 2);
    Result := Result + Encode_Byte((Byte(inStr[i]) shl 4) or (Byte(inStr[i+1]) shr 4));
    if i+1 <= Length(inStr) then
      Result := Result + Encode_Byte((Byte(inStr[i+1]) shl 2) or (Byte(inStr[i+2]) shr 6))
    else
      Result := Result + '=';
    if i+2 <= Length(inStr) then
      Result := Result + Encode_Byte(Byte(inStr[i+2]))
    else
      Result := Result + '=';
    Inc(i, 3);
  end;
end;

// Base64 decoding
function DecodeBase64(const CinLine: string): string;
const
  RESULT_ERROR = -2;
var
  inLineIndex: Integer;
  c: Char;
  x: SmallInt;
  c4: Word;
  StoredC4: array[0..3] of SmallInt;
  InLineLength: Integer;
begin
  Result := '';
  inLineIndex := 1;
  c4 := 0;
  InLineLength := Length(CinLine);

  while inLineIndex <= InLineLength do
  begin
    while (inLineIndex <= InLineLength) and (c4 < 4) do
    begin
      c := CinLine[inLineIndex];
      case c of
        '+'     : x := 62;
        '/'     : x := 63;
        '0'..'9': x := Ord(c) - (Ord('0')-52);
        '='     : x := -1;
        'A'..'Z': x := Ord(c) - Ord('A');
        'a'..'z': x := Ord(c) - (Ord('a')-26);
      else
        x := RESULT_ERROR;
      end;
      if x <> RESULT_ERROR then
      begin
        StoredC4[c4] := x;
        Inc(c4);
      end;
      Inc(inLineIndex);
    end;

    if c4 = 4 then
    begin
      c4 := 0;
      Result := Result + Char((StoredC4[0] shl 2) or (StoredC4[1] shr 4));
      if StoredC4[2] = -1 then Exit;
      Result := Result + Char((StoredC4[1] shl 4) or (StoredC4[2] shr 2));
      if StoredC4[3] = -1 then Exit;
      Result := Result + Char((StoredC4[2] shl 6) or (StoredC4[3]));
    end;
  end;
end;

{$ENDIF}

function QIDecodeBase64(AValue: qiString): string;
begin
  {$IFDEF VCL6}
  Result := {$IFDEF VCL16}Soap.{$ENDIF}EncdDecd.DecodeString(String(AValue));
  {$ELSE}
   Result := DecodeBase64(AValue);
  {$ENDIF}
end;

function QIEncodeBase64(AValue: qiString): string;
begin
  {$IFDEF VCL6}
  Result := {$IFDEF VCL16}Soap.{$ENDIF}EncdDecd.EncodeString(String(AValue));
  {$ELSE}
   Result := EncodeBase64(AValue);
  {$ENDIF}
end;


initialization
  GetMonthDayNames;

end.
