unit QImport3XLSCommon;
{$I QImport3VerCtrl.Inc}
interface

uses
  {$IFDEF VCL16}
    System.SysUtils;
  {$ELSE}
    SysUtils;
  {$ENDIF}

type
  ExlsFileError = class(Exception);

const
  MAX_COL_COUNT          = 255;
  MAX_ROW_COUNT          = 65535; {zero based}
  MAX_SHEET_COUNT        = 250;
  XLSEDITOR_MAX_ROW_COUNT  = 25;

  //------

  MAX_RECORD_DATA_SIZE   = 8224;

  //------

  BIFF_BOF_VER           = $0600;

  //------

  BIFF_BOF_GLOBALS       = $0005;
  BIFF_BOF_WORKSHEET     = $0010;
  BIFF_BOF_CHART         = $0020;

  //------

  BIFF_ROW               = $0208;

  BIFF_FORMULA           = $0006;
  BIFF_EOF               = $000A;
  BIFF_NAME              = $0018;
  BIFF_CONTINUE          = $003C;
  BIFF_BOUNDSHEET        = $0085;
  BIFF_MULRK             = $00BD;
  BIFF_MULBLANK          = $00BE;
  BIFF_XF                = $00E0;
  BIFF_SST               = $00FC;
  BIFF_LABELSST          = $00FD;
  BIFF_LABEL             = $0204;

  BIFF_BLANK             = $0201;
  BIFF_NUMBER            = $0203;
  BIFF_BOOLERR           = $0205;
  BIFF_STRING            = $0207;
  BIFF_RK                = $027E;
  BIFF_FORMAT            = $041E;
  BIFF_SHRFMLA           = $04BC;

  BIFF_BOF               = $0809;
  BIFF_BOF_3             = $0209;

  //------

  BOOL_ERR_ID_NULL       = $00;
  BOOL_ERR_ID_DIV_ZERO   = $07;
  BOOL_ERR_ID_VALUE      = $0F;
  BOOL_ERR_ID_REF        = $17;
  BOOL_ERR_ID_NAME       = $1D;
  BOOL_ERR_ID_NUM        = $24;
  BOOL_ERR_ID_NA         = $2A;

  BOOL_ERR_STR_NULL      = '#NULL!';
  BOOL_ERR_STR_DIV_ZERO  = '#DIV/0!';
  BOOL_ERR_STR_VALUE     = '#VALUE!';
  BOOL_ERR_STR_REF       = '#REF!';
  BOOL_ERR_STR_NAME      = '#NAME?';
  BOOL_ERR_STR_NUM       = '#NUM!';
  BOOL_ERR_STR_NA        = '#N/A';

type
  TBOOL_ERR_STRINGS = array[0..6] of WideString;

const
  BOOL_ERR_STRINGS: TBOOL_ERR_STRINGS =
    (BOOL_ERR_STR_NULL, BOOL_ERR_STR_DIV_ZERO, BOOL_ERR_STR_VALUE,
     BOOL_ERR_STR_REF, BOOL_ERR_STR_NAME, BOOL_ERR_STR_NUM, BOOL_ERR_STR_NA);

const
  ptgExp         = $01;
  ptgTbl         = $02;
  ptgAdd         = $03;
  ptgSub         = $04;
  ptgMul         = $05;
  ptgDiv         = $06;
  ptgPower       = $07;
  ptgConcat      = $08;
  ptgLT	         = $09;
  ptgLE	         = $0A;
  ptgEQ	         = $0B;
  ptgGE	         = $0C;
  ptgGT	         = $0D;
  ptgNE	         = $0E;
  ptgIsect       = $0F;
  ptgUnion       = $10;
  ptgRange       = $11;
  ptgUplus       = $12;
  ptgUminus      = $13;
  ptgPercent     = $14;
  ptgParen       = $15;
  ptgMissArg     = $16;
  ptgStr         = $17;
  ptgAttr        = $19;
  ptgSheet       = $1A;
  ptgEndSheet    = $1B;
  ptgErr         = $1C;
  ptgBool        = $1D;
  ptgInt         = $1E;
  ptgNum         = $1F;
  ptgArray       = $20;
  ptgFunc        = $21;
  ptgFuncVar     = $22;
  ptgName        = $23;
  ptgRef         = $24;
  ptgArea        = $25;
  ptgMemArea     = $26;
  ptgMemErr      = $27;
  ptgMemNoMem    = $28;
  ptgMemFunc     = $29;
  ptgRefErr      = $2A;
  ptgAreaErr     = $2B;
  ptgRefN        = $2C;
  ptgAreaN       = $2D;
  ptgMemAreaN    = $2E;
  ptgMemNoMemN   = $2F;
  ptgNameX       = $39;
  ptgRef3d       = $3A;
  ptgArea3d      = $3B;
  ptgRefErr3d    = $3C;
  ptgAreaErr3d   = $3D;
  ptgArrayV      = $40;
  ptgFuncV       = $41;
  ptgFuncVarV    = $42;
  ptgNameV       = $43;
  ptgRefV        = $44;
  ptgAreaV       = $45;
  ptgMemAreaV    = $46;
  ptgMemErrV     = $47;
  ptgMemNoMemV   = $48;
  ptgMemFuncV    = $49;
  ptgRefErrV     = $4A;
  ptgAreaErrV    = $4B;
  ptgRefNV       = $4C;
  ptgAreaNV      = $4D;
  ptgMemAreaNV   = $4E;
  ptgMemNoMemNV  = $4F;
  ptgFuncCEV     = $58;
  ptgNameXV      = $59;
  ptgRef3dV      = $5A;
  ptgArea3dV     = $5B;
  ptgRefErr3dV   = $5C;
  ptgAreaErr3dV  = $5D;
  ptgArrayA      = $60;
  ptgFuncA       = $61;
  ptgFuncVarA    = $62;
  ptgNameA       = $63;
  ptgRefA        = $64;
  ptgAreaA       = $65;
  ptgMemAreaA    = $66;
  ptgMemErrA     = $67;
  ptgMemNoMemA   = $68;
  ptgMemFuncA    = $69;
  ptgRefErrA     = $6A;
  ptgAreaErrA    = $6B;
  ptgRefNA       = $6C;
  ptgAreaNA      = $6D;
  ptgMemAreaNA   = $6E;
  ptgMemNoMemNA  = $6F;
  ptgFuncCEA     = $78;
  ptgNameXA      = $79;
  ptgRef3dA      = $7A;
  ptgArea3dA     = $7B;
  ptgRefErr3dA   = $7C;
  ptgAreaErr3dA  = $7D;

  tkBinary       = [ptgAdd..ptgRange];
  tkUnary        = [ptgUplus..ptgParen];
  tkArea3D       = [ptgArea3d, ptgArea3dV, ptgArea3dA];


LETTERS = 'ABCDEFGHIJKLMNOPQRSTUVWXYZ';

InternalNumberFormats: array[0..49] of WideString = (
'',
'0',
'0.00',
'#,##0',
'#,##0.00',
'_($#,##0_);($#,##0)',
'_($#,##0_);[Red]($#,##0)',
'_($#,##0.00_);($#,##0.00)',
'_($#,##0.00_);[Red]($#,##0.00)',
'0%',
'0.00%',
'0.00E+00',
'# ?/?',
'# ??/??',
'm/d/yy',
'd-mmm-y',
'd-mmm',
'mmmm-yy',
'h:mm AM/PM',
'h:mm:ss AM/PM',
'h:mm',
'h:mm:SS',
'm/d/yy h:mm',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'',
'(#,##0_);(#,##0)',
'(#,##0_);[Red](#,##0)',
'(#,##0.00_);(#,##0.00)',
'(#,##0.00_);[Red](#,##0.00)',
'_(* #,##0_);_(* (#,##0);_(* "-"_);_(@_)',
'_($* #,##0_);_($* (#,##0);_($* "-"_);_(@_)',
'_(* #,##0.00_);_(* (#,##0.00);_(* "-"??_);_(@_)',
'_($* #,##0.00_);_($* (#,##0.00);_($* "-"??_);_(@_)',
'mm:ss',
'[h]:mm:ss',
'mm:ss.0',
'# #0.0E+0',
'@');

type
  PBIFF_Header = ^TBIFF_Header;
  TBIFF_Header = packed record
    ID    : word;
    Length: word;
  end;

  PBIFF_XF = ^TBIFF_XF;
  TBIFF_XF = packed record
    FontIndex  : word;
    FormatIndex: word;
    Data1      : word;
    Data2      : word;
    Data3      : word;
    Data4      : word;
    Data5      : word;
    Data6      : longint;
    Data7      : word;
  end;

  PBIFF_PTGRef = ^TBIFF_PTGRef;
  TBIFF_PTGRef = packed record
    Row: word;
    Col: word;
  end;

  PBIFF_PTGRef3D = ^TBIFF_PTGRef3D;
  TBIFF_PTGRef3D = packed record
    Index: word;
    Row: word;
    Col: word;
  end;

  PBIFF_PTGArea = ^TBIFF_PTGArea;
  TBIFF_PTGArea = packed record
    Row1: word;
    Row2: word;
    Col1: word;
    Col2: word;
  end;

  PBIFF_PTGArea3D = ^TBIFF_PTGArea3D;
  TBIFF_PTGArea3D = packed record
    Index: word;
    Row1: word;
    Row2: word;
    Col1: word;
    Col2: word;
  end;

  PBIFF_PTGName = ^TBIFF_PTGName;
  TBIFF_PTGName = packed record
    NameIndex: word;
    Reserved: word;
  end;

  PBIFF_PTGNameX = ^TBIFF_PTGNameX;
  TBIFF_PTGNameX = packed record
    ExternSheetIndex: word;
    ExternNameIndex: word; // 1 based
    Reserved: word;
  end;

  TXLS_FUNCTION = packed record
    Min: byte;
    Max: byte;
//dee    Name: string;
    Name: AnsiString;
  end;

  TXLS_FUNCTION_ID = (fidCount, fidIf, fidIsNa, fidIsError, fidSum, fidAverage,
    fidMin, fidMax, fidRow, fidColumn, fidNA, fidNPV, fidSTDEV, fidDollar,
    fidFixed, fidSin, fidCos, fidTan, fidATan, fidPI, fidSqrt, fidExp, fidLN,
    fidLog10, fidAbs, fidInt);

  TXLS_FUNCTIONS = array[0..367] of TXLS_FUNCTION;

const
  XLS_FUNCTIONS: TXLS_FUNCTIONS = (
  {000} (Min: 01; Max: 99; Name: 'COUNT'),
  {001} (Min: 02; Max: 03; Name: 'IF'),
  {002} (Min: 01; Max: 01; Name: 'ISNA'),
  {003} (Min: 01; Max: 01; Name: 'ISERROR'),
  {004} (Min: 01; Max: 99; Name: 'SUM'),
  {005} (Min: 01; Max: 99; Name: 'AVERAGE'),
  {006} (Min: 01; Max: 99; Name: 'MIN'),
  {007} (Min: 01; Max: 99; Name: 'MAX'),
  {008} (Min: 00; Max: 01; Name: 'ROW'),
  {009} (Min: 00; Max: 01; Name: 'COLUMN'),
  {010} (Min: 01; Max: 01; Name: 'NA'),
  {011} (Min: 01; Max: 99; Name: 'NPV'),
  {012} (Min: 01; Max: 99; Name: 'STDEV'),
  {013} (Min: 01; Max: 02; Name: 'DOLLAR'),
  {014} (Min: 01; Max: 03; Name: 'FIXED'),
  {015} (Min: 01; Max: 01; Name: 'SIN'),
  {016} (Min: 01; Max: 01; Name: 'COS'),
  {017} (Min: 01; Max: 01; Name: 'TAN'),
  {018} (Min: 01; Max: 01; Name: 'ATAN'),
  {019} (Min: 00; Max: 00; Name: 'PI'),
  {020} (Min: 01; Max: 01; Name: 'SQRT'),
  {021} (Min: 01; Max: 01; Name: 'EXP'),
  {022} (Min: 01; Max: 01; Name: 'LN'),
  {023} (Min: 01; Max: 01; Name: 'LOG10'),
  {024} (Min: 01; Max: 01; Name: 'ABS'),
  {025} (Min: 01; Max: 01; Name: 'INT'),
  {026} (Min: 01; Max: 01; Name: 'SIGN'),
  {027} (Min: 02; Max: 02; Name: 'ROUND'),
  {028} (Min: 03; Max: 03; Name: 'LOOKUP'),
  {029} (Min: 01; Max: 03; Name: 'INDEX'),
  {030} (Min: 02; Max: 02; Name: 'REPT'),
  {031} (Min: 01; Max: 03; Name: 'MID'),
  {032} (Min: 01; Max: 01; Name: 'LEN'),
  {033} (Min: 01; Max: 01; Name: 'VALUE'),
  {034} (Min: 00; Max: 00; Name: 'TRUE'),
  {035} (Min: 00; Max: 00; Name: 'FALSE'),
  {036} (Min: 01; Max: 99; Name: 'AND'),
  {037} (Min: 01; Max: 99; Name: 'OR'),
  {038} (Min: 01; Max: 01; Name: 'NOT'),
  {039} (Min: 02; Max: 02; Name: 'MOD'),
  {040} (Min: 03; Max: 03; Name: 'DCOUNT'),
  {041} (Min: 03; Max: 03; Name: 'DSUM'),
  {042} (Min: 03; Max: 03; Name: 'DAVERAGE'),
  {043} (Min: 03; Max: 03; Name: 'DMIN'),
  {044} (Min: 03; Max: 03; Name: 'DMAX'),
  {045} (Min: 03; Max: 03; Name: 'DSTDEV'),
  {046} (Min: 01; Max: 99; Name: 'VAR'),
  {047} (Min: 03; Max: 03; Name: 'DVAR'),
  {048} (Min: 02; Max: 02; Name: 'TEXT'),
  {049} (Min: 01; Max: 04; Name: 'LINEST'),
  {050} (Min: 01; Max: 04; Name: 'TREND'),
  {051} (Min: 01; Max: 04; Name: 'LOGEST'),
  {052} (Min: 01; Max: 04; Name: 'GROWTH'),
  {053} (Min: 01; Max: 01; Name: 'GOTO'),
  {054} (Min: 01; Max: 01; Name: 'HALT'),
  {055} (Min: 01; Max: 01; Name: 'RETURN'),
  {056} (Min: 02; Max: 05; Name: 'PV'),
  {057} (Min: 02; Max: 05; Name: 'FV'),
  {058} (Min: 02; Max: 05; Name: 'NPER'),
  {059} (Min: 03; Max: 05; Name: 'PMT'),
  {060} (Min: 03; Max: 05; Name: 'RATE'),
  {061} (Min: 03; Max: 03; Name: 'MIRR'),
  {062} (Min: 01; Max: 02; Name: 'IRR'),
  {063} (Min: 00; Max: 00; Name: 'RAND'),
  {064} (Min: 02; Max: 03; Name: 'MATCH'),
  {065} (Min: 03; Max: 03; Name: 'DATE'),
  {066} (Min: 03; Max: 03; Name: 'TIME'),
  {067} (Min: 01; Max: 01; Name: 'DAY'),
  {068} (Min: 01; Max: 01; Name: 'MONTH'),
  {069} (Min: 01; Max: 01; Name: 'YEAR'),
  {070} (Min: 01; Max: 02; Name: 'WEEKDAY'),
  {071} (Min: 01; Max: 01; Name: 'HOUR'),
  {072} (Min: 01; Max: 01; Name: 'MINUTE'),
  {073} (Min: 01; Max: 01; Name: 'SECOND'),
  {074} (Min: 00; Max: 00; Name: 'NOW'),
  {075} (Min: 01; Max: 01; Name: 'AREAS'),
  {076} (Min: 01; Max: 01; Name: 'ROWS'),
  {077} (Min: 01; Max: 01; Name: 'COLUMNS'),
  {078} (Min: 02; Max: 04; Name: 'OFFSET'),
  {079} (Min: 01; Max: 01; Name: 'ABSREF'),
  {080} (Min: 01; Max: 01; Name: 'RELREF'),
  {081} (Min: 01; Max: 01; Name: 'ARGUMENT'),
  {082} (Min: 02; Max: 03; Name: 'SEARCH'),
  {083} (Min: 01; Max: 01; Name: 'TRANSPOSE'),
  {084} (Min: 01; Max: 01; Name: 'ERROR'),
  {085} (Min: 01; Max: 01; Name: 'STEP'),
  {086} (Min: 01; Max: 01; Name: 'TYPE'),
  {087} (Min: 01; Max: 01; Name: 'ECHO'),
  {088} (Min: 01; Max: 01; Name: 'SET.NAME'),
  {089} (Min: 01; Max: 01; Name: 'CALLER'),
  {090} (Min: 01; Max: 01; Name: 'DEREF'),
  {091} (Min: 01; Max: 01; Name: 'WINDOWS'),
  {092} (Min: 01; Max: 01; Name: 'SERIES'),
  {093} (Min: 01; Max: 01; Name: 'DOCUMENTS'),
  {094} (Min: 01; Max: 01; Name: 'ACTIVE.CELL'),
  {095} (Min: 01; Max: 01; Name: 'SELECTION'),
  {096} (Min: 01; Max: 01; Name: 'RESULT'),
  {097} (Min: 02; Max: 02; Name: 'ATAN2'),
  {098} (Min: 01; Max: 01; Name: 'ASIN'),
  {099} (Min: 01; Max: 01; Name: 'ACOS'),
  {100} (Min: 02; Max: 99; Name: 'CHOOSE'),
  {101} (Min: 03; Max: 04; Name: 'HLOOKUP'),
  {102} (Min: 03; Max: 04; Name: 'VLOOKUP'),
  {103} (Min: 01; Max: 01; Name: 'LINKS'),
  {104} (Min: 01; Max: 01; Name: 'INPUT'),
  {105} (Min: 01; Max: 01; Name: 'ISREF'),
  {106} (Min: 01; Max: 01; Name: 'GET.FORMULA'),
  {107} (Min: 01; Max: 01; Name: 'GET.NAME'),
  {108} (Min: 01; Max: 01; Name: 'SET.VALUE'),
  {109} (Min: 01; Max: 02; Name: 'LOG'),
  {110} (Min: 01; Max: 01; Name: 'EXEC'),
  {111} (Min: 01; Max: 01; Name: 'CHAR'),
  {112} (Min: 01; Max: 01; Name: 'LOWER'),
  {113} (Min: 01; Max: 01; Name: 'UPPER'),
  {114} (Min: 01; Max: 01; Name: 'PROPER'),
  {115} (Min: 01; Max: 02; Name: 'LEFT'),
  {116} (Min: 01; Max: 02; Name: 'RIGHT'),
  {117} (Min: 02; Max: 02; Name: 'EXACT'),
  {118} (Min: 01; Max: 01; Name: 'TRIM'),
  {119} (Min: 04; Max: 04; Name: 'REPLACE'),
  {120} (Min: 03; Max: 04; Name: 'SUBSTITUTE'),
  {121} (Min: 01; Max: 01; Name: 'CODE'),
  {122} (Min: 01; Max: 01; Name: 'NAMES'),
  {123} (Min: 01; Max: 01; Name: 'DIRECTORY'),
  {124} (Min: 02; Max: 03; Name: 'FIND'),
  {125} (Min: 02; Max: 02; Name: 'CELL'),
  {126} (Min: 01; Max: 01; Name: 'ISERR'),
  {127} (Min: 01; Max: 01; Name: 'ISTEXT'),
  {128} (Min: 01; Max: 01; Name: 'ISNUMBER'),
  {129} (Min: 01; Max: 01; Name: 'ISBLANK'),
  {130} (Min: 01; Max: 01; Name: 'T'),
  {131} (Min: 01; Max: 01; Name: 'N'),
  {132} (Min: 01; Max: 01; Name: 'FOPEN'),
  {133} (Min: 01; Max: 01; Name: 'FCLOSE'),
  {134} (Min: 01; Max: 01; Name: 'FSIZE'),
  {135} (Min: 01; Max: 01; Name: 'FREADLN'),
  {136} (Min: 01; Max: 01; Name: 'FREAD'),
  {137} (Min: 01; Max: 01; Name: 'FWRITELN'),
  {138} (Min: 01; Max: 01; Name: 'FWRITE'),
  {139} (Min: 01; Max: 01; Name: 'FPOS'),
  {140} (Min: 01; Max: 01; Name: 'DATEVALUE'),
  {141} (Min: 01; Max: 01; Name: 'TIMEVALUE'),
  {142} (Min: 03; Max: 03; Name: 'SLN'),
  {143} (Min: 03; Max: 03; Name: 'SYD'),
  {144} (Min: 03; Max: 05; Name: 'DDB'),
  {145} (Min: 01; Max: 01; Name: 'GET.DEF'),
  {146} (Min: 01; Max: 01; Name: 'REFTEXT'),
  {147} (Min: 01; Max: 01; Name: 'TEXTREF'),
  {148} (Min: 01; Max: 02; Name: 'INDIRECT'),
  {149} (Min: 01; Max: 03; Name: 'REGISTER'),
  {150} (Min: 01; Max: 99; Name: 'CALL'),
  {151} (Min: 01; Max: 01; Name: 'ADD.BAR'),
  {152} (Min: 01; Max: 01; Name: 'ADD.MENU'),
  {153} (Min: 01; Max: 01; Name: 'ADD.COMMAND'),
  {154} (Min: 01; Max: 01; Name: 'ENABLE.COMMAND'),
  {155} (Min: 01; Max: 01; Name: 'CHECK.COMMAND'),
  {156} (Min: 01; Max: 01; Name: 'RENAME.COMMAND'),
  {157} (Min: 01; Max: 01; Name: 'SHOW.BAR'),
  {158} (Min: 01; Max: 01; Name: 'DELETE.MENU'),
  {159} (Min: 01; Max: 01; Name: 'DELETE.COMMAND'),
  {160} (Min: 01; Max: 01; Name: 'GET.CHART.ITEM'),
  {161} (Min: 01; Max: 01; Name: 'DIALOG.BOX'),
  {162} (Min: 01; Max: 01; Name: 'CLEAN'),
  {163} (Min: 01; Max: 01; Name: 'MDETERM'),
  {164} (Min: 01; Max: 01; Name: 'MINVERSE'),
  {165} (Min: 01; Max: 01; Name: 'MMULT'),
  {166} (Min: 01; Max: 01; Name: 'FILES'),
  {167} (Min: 04; Max: 06; Name: 'IPMT'),
  {168} (Min: 04; Max: 06; Name: 'PPMT'),
  {169} (Min: 01; Max: 99; Name: 'COUNTA'),
  {170} (Min: 01; Max: 01; Name: 'CANCEL.KEY'),
  {171} (Min: 01; Max: 01; Name: 'FOR'),
  {172} (Min: 01; Max: 01; Name: 'WHILE'),
  {173} (Min: 01; Max: 01; Name: 'BREAK'),
  {174} (Min: 01; Max: 01; Name: 'NEXT'),
  {175} (Min: 01; Max: 01; Name: 'INITIATE'),
  {176} (Min: 01; Max: 01; Name: 'REQUEST'),
  {177} (Min: 01; Max: 01; Name: 'POKE'),
  {178} (Min: 01; Max: 01; Name: 'EXECUTE'),
  {179} (Min: 01; Max: 01; Name: 'TERMINATE'),
  {180} (Min: 01; Max: 01; Name: 'RESTART'),
  {181} (Min: 01; Max: 01; Name: 'HELP'),
  {182} (Min: 01; Max: 01; Name: 'GET.BAR'),
  {183} (Min: 01; Max: 99; Name: 'PRODUCT'),
  {184} (Min: 01; Max: 01; Name: 'FACT'),
  {185} (Min: 01; Max: 01; Name: 'GET.CELL'),
  {186} (Min: 01; Max: 01; Name: 'GET.WORKSPACE'),
  {187} (Min: 01; Max: 01; Name: 'GET.WINDOW'),
  {188} (Min: 01; Max: 01; Name: 'GET.DOCUMENT'),
  {189} (Min: 03; Max: 03; Name: 'DPRODUCT'),
  {190} (Min: 01; Max: 01; Name: 'ISNONTEXT'),
  {191} (Min: 01; Max: 01; Name: 'GET.NOTE'),
  {192} (Min: 01; Max: 01; Name: 'NOTE'),
  {193} (Min: 01; Max: 99; Name: 'STDEVP'),
  {194} (Min: 01; Max: 99; Name: 'VARP'),
  {195} (Min: 03; Max: 03; Name: 'DSTDEVP'),
  {196} (Min: 03; Max: 03; Name: 'DVARP'),
  {197} (Min: 01; Max: 02; Name: 'TRUNC'),
  {198} (Min: 01; Max: 01; Name: 'ISLOGICAL'),
  {199} (Min: 03; Max: 03; Name: 'DCOUNTA'),
  {200} (Min: 01; Max: 01; Name: 'DELETE.BAR'),
  {201} (Min: 01; Max: 01; Name: 'UNREGISTER'),
  {202} (Min: 01; Max: 01; Name: '202'),
  {203} (Min: 01; Max: 01; Name: '203'),
  {204} (Min: 01; Max: 01; Name: 'USDOLLAR'),
  {205} (Min: 01; Max: 01; Name: 'FINDB'),
  {206} (Min: 01; Max: 01; Name: 'SEARCHB'),
  {207} (Min: 01; Max: 01; Name: 'REPLACEB'),
  {208} (Min: 01; Max: 01; Name: 'LEFTB'),
  {209} (Min: 01; Max: 01; Name: 'RIGHTB'),
  {210} (Min: 01; Max: 01; Name: 'MIDB'),
  {211} (Min: 01; Max: 01; Name: 'LENB'),
  {212} (Min: 01; Max: 02; Name: 'ROUNDUP'),
  {213} (Min: 01; Max: 02; Name: 'ROUNDDOWN'),
  {214} (Min: 01; Max: 01; Name: 'ASC'),
  {215} (Min: 01; Max: 01; Name: 'DBCS'),
  {216} (Min: 02; Max: 03; Name: 'RANK'),
  {217} (Min: 01; Max: 01; Name: '217'),
  {218} (Min: 01; Max: 01; Name: '218'),
  {219} (Min: 02; Max: 05; Name: 'ADDRESS'),
  {220} (Min: 02; Max: 03; Name: 'DAYS360'),
  {221} (Min: 00; Max: 00; Name: 'TODAY'),
  {222} (Min: 05; Max: 07; Name: 'VDB'),
  {223} (Min: 01; Max: 01; Name: 'ELSE'),
  {224} (Min: 01; Max: 01; Name: 'ELSE.IF'),
  {225} (Min: 01; Max: 01; Name: 'END.IF'),
  {226} (Min: 01; Max: 01; Name: 'FOR.CELL'),
  {227} (Min: 01; Max: 99; Name: 'MEDIAN'),
  {228} (Min: 02; Max: 99; Name: 'SUMPRODUCT'),
  {229} (Min: 01; Max: 01; Name: 'SINH'),
  {230} (Min: 01; Max: 01; Name: 'COSH'),
  {231} (Min: 01; Max: 01; Name: 'TANH'),
  {232} (Min: 01; Max: 01; Name: 'ASINH'),
  {233} (Min: 01; Max: 01; Name: 'ACOSH'),
  {234} (Min: 01; Max: 01; Name: 'ATANH'),
  {235} (Min: 03; Max: 03; Name: 'DGET'),
  {236} (Min: 01; Max: 01; Name: 'CREATE.OBJECT'),
  {237} (Min: 01; Max: 01; Name: 'VOLATILE'),
  {238} (Min: 01; Max: 01; Name: 'LAST.ERROR'),
  {239} (Min: 01; Max: 01; Name: 'CUSTOM.UNDO'),
  {240} (Min: 01; Max: 01; Name: 'CUSTOM.REPEAT'),
  {241} (Min: 01; Max: 01; Name: 'FORMULA.CONVERT'),
  {242} (Min: 01; Max: 01; Name: 'GET.LINK.INFO'),
  {243} (Min: 01; Max: 01; Name: 'TEXT.BOX'),
  {244} (Min: 01; Max: 01; Name: 'INFO'),
  {245} (Min: 01; Max: 01; Name: 'GROUP'),
  {246} (Min: 01; Max: 01; Name: 'GET.OBJECT'),
  {247} (Min: 04; Max: 05; Name: 'DB'),
  {248} (Min: 01; Max: 01; Name: 'PAUSE'),
  {249} (Min: 01; Max: 01; Name: '249'),
  {250} (Min: 01; Max: 01; Name: '250'),
  {251} (Min: 01; Max: 01; Name: 'RESUME'),
  {252} (Min: 02; Max: 02; Name: 'FREQUENCY'),
  {253} (Min: 01; Max: 01; Name: 'ADD.TOOLBAR'),
  {254} (Min: 01; Max: 01; Name: 'DELETE.TOOLBAR'),
  {255} (Min: 01; Max: 01; Name: '255'),
  {256} (Min: 01; Max: 01; Name: 'RESET.TOOLBAR'),
  {257} (Min: 01; Max: 01; Name: 'EVALUATE'),
  {258} (Min: 01; Max: 01; Name: 'GET.TOOLBAR'),
  {259} (Min: 01; Max: 01; Name: 'GET.TOOL'),
  {260} (Min: 01; Max: 01; Name: 'SPELLING.CHECK'),
  {261} (Min: 01; Max: 01; Name: 'ERROR.TYPE'),
  {262} (Min: 01; Max: 01; Name: 'APP.TITLE'),
  {263} (Min: 01; Max: 01; Name: 'WINDOW.TITLE'),
  {264} (Min: 01; Max: 01; Name: 'SAVE.TOOLBAR'),
  {265} (Min: 01; Max: 01; Name: 'ENABLE.TOOL'),
  {266} (Min: 01; Max: 01; Name: 'PRESS.TOOL'),
  {267} (Min: 01; Max: 01; Name: 'REGISTER.ID'),
  {268} (Min: 01; Max: 01; Name: 'GET.WORKBOOK'),
  {269} (Min: 01; Max: 99; Name: 'AVEDEV'),
  {270} (Min: 03; Max: 05; Name: 'BETADIST'),
  {271} (Min: 01; Max: 01; Name: 'GAMMALN'),
  {272} (Min: 01; Max: 01; Name: 'BETAINV'),
  {273} (Min: 03; Max: 05; Name: 'BINOMDIST'),
  {274} (Min: 01; Max: 01; Name: 'CHIDIST'),
  {275} (Min: 02; Max: 02; Name: 'CHIINV'),
  {276} (Min: 02; Max: 02; Name: 'COMBIN'),
  {277} (Min: 03; Max: 03; Name: 'CONFIDENCE'),
  {278} (Min: 03; Max: 03; Name: 'CRITBINOM'),
  {279} (Min: 01; Max: 01; Name: 'EVEN'),
  {280} (Min: 03; Max: 03; Name: 'EXPONDIST'),
  {281} (Min: 03; Max: 03; Name: 'FDIST'),
  {282} (Min: 03; Max: 03; Name: 'FINV'),
  {283} (Min: 01; Max: 01; Name: 'FISHER'),
  {284} (Min: 01; Max: 01; Name: 'FISHERINV'),
  {285} (Min: 02; Max: 02; Name: 'FLOOR'),
  {286} (Min: 04; Max: 04; Name: 'GAMMADIST'),
  {287} (Min: 03; Max: 03; Name: 'GAMMAINV'),
  {288} (Min: 02; Max: 02; Name: 'CEILING'),
  {289} (Min: 04; Max: 04; Name: 'HYPGEOMDIST'),
  {290} (Min: 03; Max: 03; Name: 'LOGNORMDIST'),
  {291} (Min: 03; Max: 03; Name: 'LOGINV'),
  {292} (Min: 03; Max: 03; Name: 'NEGBINOMDIST'),
  {293} (Min: 04; Max: 04; Name: 'NORMDIST'),
  {294} (Min: 01; Max: 01; Name: 'NORMSDIST'),
  {295} (Min: 03; Max: 03; Name: 'NORMINV'),
  {296} (Min: 01; Max: 01; Name: 'NORMSINV'),
  {297} (Min: 01; Max: 01; Name: 'STANDARDIZE'),
  {298} (Min: 01; Max: 01; Name: 'ODD'),
  {299} (Min: 02; Max: 02; Name: 'PERMUT'),
  {300} (Min: 03; Max: 03; Name: 'POISSON'),
  {301} (Min: 03; Max: 03; Name: 'TDIST'),
  {302} (Min: 04; Max: 04; Name: 'WEIBULL'),
  {303} (Min: 02; Max: 02; Name: 'SUMXMY2'),
  {304} (Min: 02; Max: 02; Name: 'SUMX2MY2'),
  {305} (Min: 02; Max: 02; Name: 'SUMX2PY2'),
  {306} (Min: 02; Max: 02; Name: 'CHITEST'),
  {307} (Min: 02; Max: 02; Name: 'CORREL'),
  {308} (Min: 02; Max: 02; Name: 'COVAR'),
  {309} (Min: 03; Max: 03; Name: 'FORECAST'),
  {310} (Min: 02; Max: 02; Name: 'FTEST'),
  {311} (Min: 02; Max: 02; Name: 'INTERCEPT'),
  {312} (Min: 02; Max: 02; Name: 'PEARSON'),
  {313} (Min: 02; Max: 02; Name: 'RSQ'),
  {314} (Min: 02; Max: 02; Name: 'STEYX'),
  {315} (Min: 02; Max: 02; Name: 'SLOPE'),
  {316} (Min: 04; Max: 04; Name: 'TTEST'),
  {317} (Min: 03; Max: 04; Name: 'PROB'),
  {318} (Min: 01; Max: 99; Name: 'DEVSQ'),
  {319} (Min: 01; Max: 99; Name: 'GEOMEAN'),
  {320} (Min: 01; Max: 99; Name: 'HARMEAN'),
  {321} (Min: 01; Max: 99; Name: 'SUMSQ'),
  {322} (Min: 01; Max: 99; Name: 'KURT'),
  {323} (Min: 01; Max: 99; Name: 'SKEW'),
  {324} (Min: 02; Max: 03; Name: 'ZTEST'),
  {325} (Min: 02; Max: 02; Name: 'LARGE'),
  {326} (Min: 02; Max: 02; Name: 'SMALL'),
  {327} (Min: 02; Max: 02; Name: 'QUARTILE'),
  {328} (Min: 02; Max: 02; Name: 'PERCENTILE'),
  {329} (Min: 02; Max: 03; Name: 'PERCENTRANK'),
  {330} (Min: 01; Max: 99; Name: 'MODE'),
  {331} (Min: 02; Max: 02; Name: 'TRIMMEAN'),
  {332} (Min: 02; Max: 02; Name: 'TINV'),
  {333} (Min: 01; Max: 01; Name: '333'),
  {334} (Min: 01; Max: 01; Name: 'MOVIE.COMMAND'),
  {335} (Min: 01; Max: 01; Name: 'GET.MOVIE'),
  {336} (Min: 01; Max: 99; Name: 'CONCATENATE'),
  {337} (Min: 02; Max: 02; Name: 'POWER'),
  {338} (Min: 01; Max: 01; Name: 'PIVOT.ADD.DATA'),
  {339} (Min: 01; Max: 01; Name: 'GET.PIVOT.TABLE'),
  {340} (Min: 01; Max: 01; Name: 'GET.PIVOT.FIELD'),
  {341} (Min: 01; Max: 01; Name: 'GET.PIVOT.ITEM'),
  {342} (Min: 01; Max: 01; Name: 'RADIANS'),
  {343} (Min: 01; Max: 01; Name: 'DEGREES'),
  {344} (Min: 02; Max: 99; Name: 'SUBTOTAL'),
  {345} (Min: 02; Max: 03; Name: 'SUMIF'),
  {346} (Min: 02; Max: 02; Name: 'COUNTIF'),
  {347} (Min: 01; Max: 01; Name: 'COUNTBLANK'),
  {348} (Min: 01; Max: 01; Name: 'SCENARIO.GET'),
  {349} (Min: 01; Max: 01; Name: 'OPTIONS.LISTS.GET'),
  {350} (Min: 01; Max: 01; Name: 'ISPMT'),
  {351} (Min: 01; Max: 01; Name: 'DATEDIF'),
  {352} (Min: 01; Max: 01; Name: 'DATESTRING'),
  {353} (Min: 01; Max: 01; Name: 'NUMBERSTRING'),
  {354} (Min: 01; Max: 02; Name: 'ROMAN'),
  {355} (Min: 01; Max: 01; Name: 'OPEN.DIALOG'),
  {356} (Min: 01; Max: 01; Name: 'SAVE.DIALOG'),
  {357} (Min: 01; Max: 01; Name: 'VIEW.GET'),
  {358} (Min: 02; Max: 99; Name: 'GETPIVOTDATA'),
  {359} (Min: 01; Max: 02; Name: 'HYPERLINK'),
  {360} (Min: 01; Max: 01; Name: 'PHONETIC'),
  {361} (Min: 01; Max: 99; Name: 'AVERAGEA'),
  {362} (Min: 01; Max: 99; Name: 'MAXA'),
  {363} (Min: 01; Max: 99; Name: 'MINA'),
  {364} (Min: 01; Max: 99; Name: 'STDEVPA'),
  {365} (Min: 01; Max: 99; Name: 'VARPA'),
  {366} (Min: 01; Max: 99; Name: 'STDEVA'),
  {367} (Min: 01; Max: 99; Name: 'VARA'));

implementation

//uses XLSConsts;

{ ExlsTokenError }

{constructor ExlsTokenError.Create(Token: integer);
begin
  FToken := Token;
  inherited CreateFmt(sBadToken, [Token]);
end;}

end.
