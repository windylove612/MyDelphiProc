unit QImport3XLSMapParser;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    System.Classes,
  {$ELSE}
    Classes,
  {$ENDIF}
  QImport3XLSFile;

type
  TRangeType = (rtCol, rtRow, rtCell, rtUnknown);
  TRangeDirection = (rdDown, rdUp, rdUnknown);
  TSheetIDType = (sitUnknown, sitName, sitNumber);

  TMapRow = class;

  TMapRange = class
  private
    FMapRow: TMapRow;
    FXLSFile: TxlsFile;

    FDirection: TRangeDirection;
    FSheetIDType: TSheetIDType;
    FSheetNumber: integer;
    FSheetName: WideString;

    FRow1: integer;
    FCol1: integer;
    FRow2: integer;
    FCol2: integer;

    FRangeType: TRangeType;
    FLength: integer;

    procedure Arrange;
    function GetHasSheet: boolean;
    function GetAsString: string;
    function GetSkipFirstRows: integer;
    function GetSkipFirstCols: integer;
  protected
    procedure UpdateDirection;
  public
    constructor Create(MapRow: TMapRow);
    procedure Update;
    procedure Assign(Range: TMapRange);

    property MapRow: TMapRow read FMapRow write FMapRow;
    property XLSFile: TxlsFile read FXLSFile;

    property RangeType: TRangeType read FRangeType;
    property Direction: TRangeDirection read FDirection write FDirection;
    property HasSheet: boolean read GetHasSheet;
    property SheetIDType: TSheetIDType read FSheetIDType write FSheetIDType;
    property SheetNumber: integer read FSheetNumber write FSheetNumber; // 1 based
    property SheetName: WideString read FSheetName write FSheetName;

    property Row1: integer read FRow1 write FRow1; // 1 based
    property Col1: integer read FCol1 write FCol1; // 1 based
    property Row2: integer read FRow2 write FRow2; // 1 based
    property Col2: integer read FCol2 write FCol2; // 1 based

    property Length: integer read FLength;
    property AsString: string read GetAsString;

    property SkipFirstRows: integer read GetSkipFirstRows;
    property SkipFirstCols: integer read GetSkipFirstCols;
  end;

  TMapRowList = class;

  TMapRow = class(TList)
  private
    FMapRowList: TMapRowList;
    FXLSFile: TxlsFile;
    FLength: integer;
    function GetItems(Index: integer): TMapRange;
    procedure SetItems(Index: integer; Value: TMapRange);
    function GetAsString: string;
    procedure SetAsString(const Value: string);
    function GetSkipFirstRows: integer;
    function GetSkipFirstCols: integer;
  public
    constructor Create(MapRowList: TMapRowList);
    function Add(Item: TMapRange): integer;
    procedure Delete(Index: integer);
    procedure Update;
    function GetCellValue(AbsoluteIndex: integer): WideString;
    procedure Optimize;
    function IndexOfRange(const RangeStr: string): integer;

    property MapRowList: TMapRowList read FMapRowList;
    property XLSFile: TxlsFile read FXLSFile;

    property Items[Index: integer]: TMapRange read GetItems
      write SetItems; default;
    property Length: integer read FLength;
    property AsString: string read GetAsString write SetAsString;

    property SkipFirstRows: integer read GetSkipFirstRows;
    property SkipFirstCols: integer read GetSkipFirstCols;
  end;

  TMapRowList = class(TList)
  private
    FXLSFile: TxlsFile;
    FMaxRow: integer;
    FMinRow: integer;
    FSkipFirstCols: integer;
    FSkipFirstRows: integer;
    function GetItems(Index: integer): TMapRow;
    procedure SetItems(Index: integer; Value: TMapRow);
  public
    constructor Create(XLSFile: TxlsFile);
    destructor Destroy; override;

    function Add(Item: TMapRow): integer;
    procedure Delete(Index: integer);
    procedure Update;

    property XLSFile: TxlsFile read FXLSFile;
    property Items[Index: integer]: TMapRow read GetItems
      write SetItems; default;
    property MaxRow: integer read FMaxRow;
    property MinRow: integer read FMinRow;

    property SkipFirstRows: integer read FSkipFirstRows write FSkipFirstRows;
    property SkipFirstCols: integer read FSkipFirstCols write FSkipFirstCols;
  end;

  TCellNeighbour = (cnLeft, cnRight, cnTop, cnBottom);
  TCellNeighbours = set of TCellNeighbour;

  function ParseMapString(const MapString: string; MapRow: TMapRow): boolean;
  procedure ParseCellString(const CellString: string; var Col, Row: integer);
  procedure ParseColString(const ColString: string; var Col: integer);
  procedure ParseRowString(const RowString: string; var Row: integer);
  procedure ParseSheetNumber(const SheetNumber: string; var Sheet: integer);
  procedure ParseSheetName(const SheetName: string; var Sheet: string);

  function CellInRange(Range: TMapRange; const SheetName: string;
    SheetNumber, Col, Row: integer): boolean;
  function CellInRow(MapRow: TMapRow; const SheetName: string;
    SheetNumber, Col, Row: integer): boolean;
  function GetCellNeighbours(MapRow: TMapRow; const SheetName: string;
    SheetNumber, Col, Row: integer): TCellNeighbours;
  procedure RemoveCellFromRow(MapRow: TMapRow; const SheetName: string;
    SheetNumber, Col, Row: integer);

  procedure Str2ColRow(const Str: string; var ACol, ARow: integer);
  procedure Str2Range(const Str: string; var ACol1, ARow1, ACol2, ARow2: integer);
  function GetRangeType(const Str: string): TRangeType;
  function OptimizeString(const Str: string): string;
  function SkipFirstRows(const Str: string; Rows: integer): string;
  function SkipFirstCols(const Str: string; Cols: integer): string;
  function CheckRange(const Str: string): string;

const
  MAX_COL = 256;
  MAX_ROW = 65536;

const
  RANGE_DELIMITER     = ';';
  ARRAY_DELIMITER     = '-';
  SHEET_START         = '[';
  SHEET_FINISH        = ']';
  SHEET_NUMBER        = ':';
  ILLEGAL_IN_SHEET    = ':\/?*[]';

  COLSTART  = 'COLSTART';
  COLFINISH = 'COLFINISH';
  ROWSTART  = 'ROWSTART';
  ROWFINISH = 'ROWFINISH';

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
    System.Math,
    Winapi.Windows,
  {$ELSE}
    SysUtils,
    Math,
    {$IFDEF VCL9}
      Windows,
    {$ENDIF}
  {$ENDIF}
  QImport3XLSUtils;

type
  TSymbolType = (stUnknown, stLetter, stNumber, stRange, stArray,
    stSheetStart, stSheetFinish, stSheetNumber);

const
  sUnknownSymbol          = 'The %s symbol at position %d is unknown';
  sUnexpectedSymbol       = 'The symbol %s at position %d is unexpected';
  sUnexpectedKeyword      = 'Expect %s but %s found';
  sIllegalSheetChar       = 'Illegal char %s in sheet name at position %d';
  sLetterExpected         = 'Letter expected but %s found (position %d)';
  sNumberExpected         = 'Number expected but %s found (position %d)';
  sLetterOrNumberExpected = 'Letter or number expected but %s found (position %d)';
  sColIsOutOfRange        = 'Column %d is out of range [%s..%s]';
  sRowIsOutOfRange        = 'Row %d is out of range [%d..%d]';
  sRangeFail              = 'Range %s is fail. It must be COL or ROW';
  sUnexpectedEndOfRange   = 'Unexpected end of range';
  sSheetNotFound          = 'Sheet with name %s is not found in the source file';
  sRowNotFound            = 'Row %d not found in the sheet %s';
  sColNotFound            = 'Col %d not found in the sheet %s';
  sCellIsEmpty            = 'Cell is empty';
  sSoLongCellDefinition   = 'So long cell definition';
  sColIsEmpty             = 'Col is empty';
  sSoLongColDefinition    = 'So long col definition';
  sRowIsEmpty             = 'Row is empty';
  sSoLongRowDefinition    = 'So long row definition';
  sSheetIsEmpty           = 'Sheet is empty';
  sSoLongSheetDefinition  = 'So long sheet definition';

function IsLetter(Ch: char): boolean;
begin
  Result := Pos(Ch, LETTERS) > 0;
end;

function IsNumber(Ch: char): boolean;
begin
  Result := Pos(Ch, NUMBERS) > 0;
end;

function IsRange(Ch: char): boolean;
begin
  Result := Ch = RANGE_DELIMITER;
end;

function IsArray(Ch: char): boolean;
begin
  Result := Ch = ARRAY_DELIMITER;
end;

function IsSheetStart(Ch: char): boolean;
begin
  Result := Ch = SHEET_START;
end;

function IsSheetFinish(Ch: char): boolean;
begin
  Result := Ch = SHEET_FINISH;
end;

function IsSheetNumber(Ch: char): boolean;
begin
  Result := Ch = SHEET_NUMBER;
end;

function IsIllegalInSheet(Ch: char): boolean;
begin
  Result := Pos(Ch, ILLEGAL_IN_SHEET) > 0;
end;

function IsColStart(const Str: string): boolean;
begin
  Result := AnsiCompareText(Str, COLSTART) = 0;
end;

function IsColFinish(const Str: string): boolean;
begin
  Result := AnsiCompareText(Str, COLFINISH) = 0;
end;

function IsRowStart(const Str: string): boolean;
begin
  Result := AnsiCompareText(Str, ROWSTART) = 0;
end;

function IsRowFinish(const Str: string): boolean;
begin
  Result := AnsiCompareText(Str, ROWFINISH) = 0;
end;

function IsKeyword(const Str: string): boolean;
begin
  Result := IsColStart(Str) or IsColFinish(Str) or
    IsRowStart(Str) or IsRowFinish(Str);
end;

function IsColKeyword(const Str: string): boolean;
begin
  Result := IsColStart(Str) or IsColFinish(Str);
end;

function IsRowKeyword(const Str: string): boolean;
begin
  Result := IsRowStart(Str) or IsRowFinish(Str);
end;

function GetSymbolType(Ch: char): TSymbolType;
begin
  if IsLetter(Ch) then
    Result := stLetter
  else if IsNumber(Ch) then
    Result := stNumber
  else if IsRange(Ch) then
    Result := stRange
  else if IsArray(Ch) then
    Result := stArray
  else if IsSheetStart(Ch) then
    Result := stSheetStart
  else if IsSheetFinish(Ch) then
    Result := stSheetFinish
  else if IsSheetNumber(Ch) then
    Result := stSheetNumber
  else Result := stUnknown;
end;

procedure CheckRowNumber(Row: integer);
begin
  if (Row <= 0) or (Row > MAX_ROW) then
    raise Exception.CreateFmt(sRowIsOutOfRange, [Row, 1, MAX_ROW]);
end;

procedure CheckColNumber(Col: integer);
begin
  if (Col <= 0) or (Col > MAX_COL) then
    raise Exception.CreateFmt(sColIsOutOfRange, [Col, 'A', 'AZ']);
end;


function ParseMapString(const MapString: string; MapRow: TMapRow): boolean;
type
  TState = 0..14;
var
  State: TState;
  SheetFlag: boolean;
  RangeType: TRangeType;
  i: integer;
  Str, Buf: string;
  MapRange: TMapRange;
  SymbolType: TSymbolType;
  Ch: char;
  T: integer;
begin
  Result := true;
  Str := Trim(MapString);
  if Str = EmptyStr then Exit;
  //for i := Length(Str) downto 1 do if Str[i] = ' ' then System.Delete(Str, i, 1);
  Str := UpperCase(Str);
  if Str[Length(Str)] <> ';' then
    Str := Str + ';';

  State := 0;
  SheetFlag := false;
  RangeType := rtUnknown;
  MapRange := nil;

  for i := 1 to Length(Str) do
  begin
    Ch := Str[i];

    SymbolType := GetSymbolType(Ch);
    if (SymbolType = stUnknown) and not (State in [11, 12]) then
      raise Exception.CreateFmt(sUnknownSymbol, [Ch, i]);

    Result := true;

    try
      case State of
        0: begin
          if not SheetFlag and Assigned(MapRow) and Assigned(MapRange) then
            MapRow.Add(MapRange);
          if not SheetFlag and Assigned(MapRow) then
            MapRange := TMapRange.Create(MapRow);
          Buf := EmptyStr;
          case SymbolType of
            stLetter: State := 1;
            stNumber: State := 8;
            stSheetStart: State := 11;
            else begin
              if IsRange(Ch)
                then raise Exception.Create(sUnexpectedEndOfRange)
                else raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
            end;
          end;
        end;
        1: case SymbolType of
             stLetter: State := 1;
             stNumber: begin // Buf contains col name
               T := Letter2Col(Buf);
               CheckColNumber(T);
               if Assigned(MapRange) then MapRange.Col1 := T;
               Buf := EmptyStr;
               State := 2;
             end;
             stArray: begin // Buf contains col name or keyword
               T := Letter2Col(Buf);
               if (T <= 0) or (T > MAX_COL) then
               begin
                 if IsKeyword(Buf) then // expects a cell
                 begin
                   if IsRowKeyword(Buf) then
                   begin
                     RangeType := rtRow;
                     if IsRowStart(Buf)
                       then MapRange.Direction := rdDown
                       else MapRange.Direction := rdUp;
                   end
                   else begin
                     if IsColStart(Buf)
                       then MapRange.Direction := rdDown
                       else MapRange.Direction := rdUp;
                     RangeType := rtCol;
                   end;
                   State := 3;
                 end
                 else raise Exception.CreateFmt(sUnexpectedKeyword,
                   [COLSTART + ', ' + COLFINISH + ', ' +
                    ROWSTART + ' or ' + ROWFINISH, Buf]);
               end
               else begin
                 if Assigned(MapRange) then
                   MapRange.Col1 := T;
                 State := 6;
               end;
             end;
             else
               raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
           end;
        2: case SymbolType of
             stNumber: State := 2;
             stRange: begin // Buf contains row number (single  cell)
               T := Number2Row(Buf);
               CheckRowNumber(T);
               SheetFlag := false;
               if Assigned(MapRange) then
               begin
                 MapRange.Row1 := T;
                 MapRange.Col2 := MapRange.Col1;
                 MapRange.Row2 := MapRange.Row1;
               end;
               State := 0;
             end;
             stArray: begin // Buf contains row number of the first cell of range.
               T := Number2Row(Buf);
               CheckRowNumber(T);
               if Assigned(MapRange) then
                 MapRange.Row1 := T;
               State := 3;
             end;
             else
               raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
           end;
        3: begin
          Buf := EmptyStr;
          case SymbolType of
            stLetter: State := 4;
            else begin
              if IsRange(Ch)
                then raise Exception.CreateFmt(sLetterExpected, ['end of range', i])
                else raise Exception.CreateFmt(sLetterExpected, [Ch, i]);
            end;
          end;
        end;
        4: case SymbolType of
             stLetter: State := 4;
             stNumber: begin // Buf contains col name of the second cell in the range
               T := Letter2Col(Buf);
               CheckColNumber(T);
               if Assigned(MapRange) then MapRange.Col2 := T;
               Buf := EmptyStr;
               State := 5;
             end;
             stRange: begin
               if IsKeyword(Buf) then // Buf contains keyword
                 begin
                 SheetFlag := false;
                 if Assigned(MapRange) then
                   if IsRowKeyword(Buf) then
                   begin
                     MapRange.Row2 := MapRange.Row1;
                     if IsRowStart(Buf)
                       then MapRange.Direction := rdUp
                       else MapRange.Direction := rdDown;
                   end
                   else begin
                     MapRange.Col2 := MapRange.Col1;
                     if IsColStart(Buf)
                       then MapRange.Direction := rdUp
                       else MapRange.Direction := rdDown;
                   end;
                 State := 0;
               end
               else raise Exception.CreateFmt(sUnexpectedKeyword,
                 [COLSTART + ', ' + COLFINISH + ', ' +
                  ROWSTART + ' or ' + ROWFINISH, Buf]);
             end;
             else raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
           end;
        5: case SymbolType of
             stNumber: State := 5;
             stRange: begin // Buf contains row number of the second cell in the range
               T := Number2Row(Buf);
               CheckRowNumber(T);
               if Assigned(MapRange) then
               begin
                 MapRange.Row2 := T;
                 if RangeType = rtCol then
                   MapRange.Col1 := MapRange.Col2
                 else if RangeType = rtRow then
                   MapRange.Row1 := MapRange.Row2
                 else if RangeType = rtUnknown then
                 begin
                   if (MapRange.Col1 = MapRange.Col2) or
                      (MapRange.Row1 = MapRange.Row2) then
                   begin
                      if (MapRange.Col1 < MapRange.Col2) or
                         (MapRange.Row1 < MapRange.Row2) then
                        MapRange.Direction := rdDown
                      else if (MapRange.Col1 > MapRange.Col2) or
                              (MapRange.Row1 > MapRange.Row2) then
                        MapRange.Direction := rdUp;
                   end
                   else
                     raise Exception.CreateFmt(sRangeFail,
                       [Col2Letter(MapRange.Col1) + Row2Number(MapRange.Row1) + '-' +
                       Col2Letter(MapRange.Col2) + Row2Number(MapRange.Row2)]);
                 end;
                 RangeType := rtUnknown;
               end;
               SheetFlag := false;
               State := 0;
             end;
             else raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
           end;
        6: case SymbolType of
             stLetter: begin
               Buf := EmptyStr;
               State := 7;
             end
             else begin
              if IsRange(Ch)
                then raise Exception.CreateFmt(sLetterExpected, ['end of range', i])
                else raise Exception.CreateFmt(sLetterExpected, [Ch, i]);
             end;
           end;
        7: case SymbolType of
             stLetter: State := 7;
             stRange:  begin
              if IsColKeyword(Buf) then // Buf contains col keyword
              begin
                if Assigned(MapRange) then
                begin
                  MapRange.Col2 := MapRange.Col1;
                  if IsColStart(Buf)
                    then MapRange.Direction := rdUp
                    else MapRange.Direction := rdDown;
                end;
                SheetFlag := false;
                State := 0;
              end
              else
                raise Exception.CreateFmt(sUnexpectedKeyword,
                  [COLSTART + ' or ' + COLFINISH, Buf]);
             end;
             else raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
           end;
        8: case SymbolType of
             stNumber: State := 8;
             stArray: begin // Buf contains row number
               T := Number2Row(Buf);
               CheckRowNumber(T);
               if Assigned(MapRange) then
                 MapRange.Row1 := T;
               State := 9;
             end;
             else
               raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
           end;
        9: begin
          Buf := EmptyStr;
          case SymbolType of
            stLetter: State := 10;
            else begin
              if IsRange(Ch)
                then raise Exception.CreateFmt(sLetterExpected, ['end of range', i])
                else raise Exception.CreateFmt(sLetterExpected, [Ch, i]);
            end;
          end;
        end;
        10: case SymbolType of
              stLetter: State := 10;
              stRange:  begin
                if IsRowKeyword(Buf) then // Buf contains row keyword
                  begin
                  if Assigned(MapRange) then
                  begin
                    MapRange.Row2 := MapRange.Row1;
                    if IsRowStart(Buf)
                      then MapRange.Direction := rdUp
                      else MapRange.Direction := rdDown;
                  end;
                  SheetFlag := false;
                  State := 0;
                end
                else
                  raise Exception.CreateFmt(sUnexpectedKeyword,
                    [ROWSTART + ' or ' + ROWFINISH, Buf]);
              end;
              else
                raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
            end;
        11: begin
          Buf := EmptyStr;
          if not IsIllegalInSheet(Ch) then
            State := 12
          else if IsSheetNumber(Ch) then
            State := 13
          else
            raise Exception.CreateFmt(sIllegalSheetChar, [Ch, i]);

          {case SymbolType of
            stLetter: State := 12;
            stNumber: State := 12;
            stSheetNumber: State := 13;
            else begin
              if IsRange(Ch)
                then raise Exception.CreateFmt(sLetterOrNumberExpected, ['end of range', i])
                else raise Exception.CreateFmt(sLetterOrNumberExpected, [Ch, i]);
            end;
          end;}
        end;
        12: begin
          if not IsIllegalInSheet(Ch) then
            State := 12
          else if IsSheetFinish(Ch) then
          begin
            if Assigned(MapRange) then
            begin
              MapRange.SheetIDType := sitName;
              MapRange.SheetName := Buf;
            end;
            SheetFlag := true;
            State := 0;
          end
          else
            raise Exception.CreateFmt(sIllegalSheetChar, [Ch, i]);

          {  case SymbolType of
              stLetter: State := 12;
              stNumber: State := 12;
              stSheetFinish: begin // buff contains sheet name
                if Assigned(MapRange) then begin
                  MapRange.SheetIDType := sitName;
                  MapRange.SheetName := Buf;
                end;
                SheetFlag := true;
                State := 0;
              end;
              else begin
                if IsIllegalInSheet(Ch)
                  then raise Exception.CreateFmt(sIllegalSheetChar, [Ch, i]);
              end;
            end;}
        end;
        13: case SymbolType of
              stNumber: begin
                Buf := EmptyStr;
                State := 14;
              end;
              else raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
            end;
        14: case SymbolType of
              stNumber: State := 14;
              stSheetFinish: begin // buff contains sheet name
                if Assigned(MapRange) then
                begin
                  MapRange.SheetIDType := sitNumber;
                  MapRange.SheetNumber := StrToInt(Buf);
                end;
                SheetFlag := true;
                State := 0;
              end;
              else
                raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
            end;
      end;
    except
      if Assigned(MapRange) then
        MapRange.Free;
      raise;
    end;
    Buf := Buf + Ch;
  end;
  if Assigned(MapRow) and Assigned(MapRange) then
    MapRow.Add(MapRange);
end;

procedure ParseCellString(const CellString: string; var Col, Row: integer);
type
  TState = 0..2;
var
  i: integer;
  Str: string;
  State: TState;
  Ch: char;
  SymbolType: TSymbolType;
  Buf: string;
  T: integer;
begin
  Str := CellString;
  if Str = EmptyStr then
    raise Exception.Create(sCellIsEmpty);
  Str := AnsiUpperCase(Trim(Str));
  if Str[Length(Str)] <> RANGE_DELIMITER then
    Str := Str + RANGE_DELIMITER;

  State := 0;
  Buf := EmptyStr;

  for i := 1 to Length(Str) do begin
    Ch := Str[i];
    SymbolType := GetSymbolType(Ch);
    if SymbolType = stUnknown then
      raise Exception.CreateFmt(sUnknownSymbol, [Ch, i]);
    case State of
      0: case SymbolType of
           stLetter: State := 1;
           else raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
         end;
      1: case SymbolType of
           stLetter: State := 1;
           stNumber: begin
             T := Letter2Col(Buf);
             CheckColNumber(T);
             Buf := EmptyStr;
             Col := T;
             State := 2;
           end;
           else begin
             if IsRange(Ch)
                then raise Exception.CreateFmt(sNumberExpected, ['end of range', i])
                else raise Exception.CreateFmt(sNumberExpected, [Ch, i]);
           end;
         end;
      2: case SymbolType of
           stNumber: State := 2;
           stRange: begin
             T := Number2Row(Buf);
             CheckRowNumber(T);
             if i <> Length(Str) then
               raise Exception.Create(sSoLongCellDefinition);
             Row := T;
           end;
           else raise Exception.CreateFmt(sUnexpectedSymbol, [Ch, i]);
         end;
    end;
    Buf := Buf + Ch;
  end;
end;

procedure ParseColString(const ColString: string; var Col: integer);
type
  TState = 0..1;
var
  i: integer;
  Str: string;
  State: TState;
  Ch: char;
  SymbolType: TSymbolType;
  Buf: string;
  T: integer;
begin
  Str := ColString;
  if Str = EmptyStr then
    raise Exception.Create(sColIsEmpty);
  Str := AnsiUpperCase(Trim(Str));
  if Str[Length(Str)] <> RANGE_DELIMITER then
    Str := Str + RANGE_DELIMITER;

  State := 0;
  Buf := EmptyStr;

  for i := 1 to Length(Str) do begin
    Ch := Str[i];
    SymbolType := GetSymbolType(Ch);
    if SymbolType = stUnknown then
      raise Exception.CreateFmt(sUnknownSymbol, [Ch, i]);
    case State of
      0: case SymbolType of
           stLetter: State := 1;
           else raise Exception.CreateFmt(sLetterExpected, [Ch, i]);
         end;
      1: case SymbolType of
           stLetter: State := 1;
           stRange: begin
             T := Letter2Col(Buf);
             CheckColNumber(T);
             if i <> Length(Str) then
               raise Exception.Create(sSoLongColDefinition);
             Col := T;
           end;
           else raise Exception.CreateFmt(sLetterExpected, [Ch, i]);
         end;
    end;
    Buf := Buf + Ch;
  end;
end;

procedure ParseRowString(const RowString: string; var Row: integer);
type
  TState = 0..1;
var
  i: integer;
  Str: string;
  State: TState;
  Ch: char;
  SymbolType: TSymbolType;
  Buf: string;
  T: integer;
begin
  Str := RowString;
  if Str = EmptyStr then
    raise Exception.Create(sRowIsEmpty);
  Str := AnsiUpperCase(Trim(Str));
  if Str[Length(Str)] <> RANGE_DELIMITER then
    Str := Str + RANGE_DELIMITER;

  State := 0;
  Buf := EmptyStr;

  for i := 1 to Length(Str) do begin
    Ch := Str[i];
    SymbolType := GetSymbolType(Ch);
    if SymbolType = stUnknown then
      raise Exception.CreateFmt(sUnknownSymbol, [Ch, i]);
    case State of
      0: case SymbolType of
           stNumber: State := 1;
           else raise Exception.CreateFmt(sNumberExpected, [Ch, i]);
         end;
      1: case SymbolType of
           stNumber: State := 1;
           stRange: begin
             T := Number2Row(Buf);
             CheckRowNumber(T);
             if i <> Length(Str) then
               raise Exception.Create(sSoLongRowDefinition);
             Row := T;
           end;
           else raise Exception.CreateFmt(sNumberExpected, [Ch, i]);
         end;
    end;
    Buf := Buf + Ch;
  end;
end;

procedure ParseSheetNumber(const SheetNumber: string; var Sheet: integer);
type
  TState = 0..1;
var
  i: integer;
  Str: string;
  State: TState;
  Ch: char;
  SymbolType: TSymbolType;
  Buf: string;
  T: integer;
begin
  Str := SheetNumber;
  if Str = EmptyStr then
    raise Exception.Create(sSheetIsEmpty);
  Str := AnsiUpperCase(Trim(Str));
  if Str[Length(Str)] <> RANGE_DELIMITER then
    Str := Str + RANGE_DELIMITER;

  State := 0;
  Buf := EmptyStr;

  for i := 1 to Length(Str) do begin
    Ch := Str[i];
    SymbolType := GetSymbolType(Ch);
    if SymbolType = stUnknown then
      raise Exception.CreateFmt(sUnknownSymbol, [Ch, i]);
    case State of
      0: case SymbolType of
           stNumber: State := 1;
           else raise Exception.CreateFmt(sNumberExpected, [Ch, i]);
         end;
      1: case SymbolType of
           stNumber: State := 1;
           stRange: begin
             T := StrToInt(Buf);
             if i <> Length(Str) then
               raise Exception.Create(sSoLongSheetDefinition);
             Sheet := T;
           end;
           else raise Exception.CreateFmt(sNumberExpected, [Ch, i]);
         end;
    end;
    Buf := Buf + Ch;
  end;
end;

procedure ParseSheetName(const SheetName: string; var Sheet: string);
type
  TState = 0..1;
var
  i: integer;
  Str, Str1: string;
  State: TState;
  Ch: char;
  SymbolType: TSymbolType;
  Buf: string;
begin
  Str1 := Trim(SheetName);
  if Str1 = EmptyStr then
    raise Exception.Create(sSheetIsEmpty);
  Str := AnsiUpperCase(Str1);
  if Str[Length(Str)] <> RANGE_DELIMITER then
  begin
    Str := Str + RANGE_DELIMITER;
    Str1 := Str1 + RANGE_DELIMITER;
  end;

  State := 0;
  Buf := EmptyStr;

  for i := 1 to Length(Str) do begin
    Ch := Str[i];
    SymbolType := GetSymbolType(Ch);
    {if SymbolType = stUnknown then
      raise Exception.CreateFmt(sUnknownSymbol, [Str1[i], i]);}
    case State of
      0: begin
        if IsIllegalInSheet(Ch) then
          raise Exception.CreateFmt(sIllegalSheetChar, [Ch, i]);
        State := 1;
      end;
      1: case SymbolType of
           stRange: begin
             if i <> Length(Str) then
               raise Exception.Create(sSoLongSheetDefinition);
             Sheet := Buf;
           end;
           else begin
             if IsIllegalInSheet(Ch) then
               raise Exception.CreateFmt(sIllegalSheetChar, [Ch, i]);
             State := 1;
           end;
         end;
    end;
    Buf := Buf + Str1[i];
  end;
end;

function CellInRange(Range: TMapRange; const SheetName: string;
  SheetNumber, Col, Row: integer): boolean;
begin
  Result := false;
  Range.Update;

  if ((Range.SheetIDType = sitNumber) and (Range.SheetNumber <> SheetNumber)) or
     ((Range.SheetIDType = sitName) and (AnsiCompareText(Range.SheetName, SheetName) <> 0)) then
    Exit;

  case Range.RangeType of
    rtCol:
      if Range.Col1 = Col then begin
        if (Range.Row1 > 0) and (Range.Row2 > 0) then begin
          if Range.Row1 < Range.Row2 then
            Result := (Row >= Range.Row1) and (Row <= Range.Row2)
          else if Range.Row1 > Range.Row2 then
            Result := (Row >= Range.Row2) and (Row <= Range.Row1)
        end
        else if (Range.Row1 > 0) and (Range.Row2 = 0) then begin
          Result := ((Range.Direction = rdDown) and (Row >= Range.Row1)) or
                    ((Range.Direction = rdUp) and (Row <= Range.Row1));
        end
        else if (Range.Row1 = 0) and (Range.Row2 > 0) then begin
          Result := ((Range.Direction = rdDown) and (Row <= Range.Row1)) or
                    ((Range.Direction = rdUp) and (Row >= Range.Row1));
        end
        else if (Range.Row1 = 0) and (Range.Row2 = 0) then begin
          Result := true;
        end;
      end;
    rtRow:
      if Range.Row1 = Row then begin
        if (Range.Col1 > 0) and (Range.Col2 > 0) then begin
          if Range.Col1 < Range.Col2 then
            Result := (Col >= Range.Col1) and (Col <= Range.Col2)
          else if Range.Col1 > Range.Col2 then
            Result := (Col >= Range.Col2) and (Col <= Range.Col1)
        end
        else if (Range.Col1 > 0) and (Range.Col2 = 0) then begin
          Result := ((Range.Direction = rdDown) and (Col >= Range.Col1)) or
                    ((Range.Direction = rdUp) and (Col <= Range.Col1));
        end
        else if (Range.Col1 = 0) and (Range.Col2 > 0) then begin
          Result := ((Range.Direction = rdDown) and (Col <= Range.Col1)) or
                    ((Range.Direction = rdUp) and (Col >= Range.Col1));
        end
        else if (Range.Col1 = 0) and (Range.Col2 = 0) then begin
          Result := true;
        end;
      end;
    rtCell: Result := (Range.Col1 = Col) and (Range.Row1 = Row);
  end;
end;

function CellInRow(MapRow: TMapRow; const SheetName: string;
  SheetNumber, Col, Row: integer): boolean;
var
  i: integer;
begin
  Result := false;
  for i := 0 to MapRow.Count - 1 do
    if CellInRange(MapRow[i], SheetName, SheetNumber, Col, Row) then begin
      Result := true;
      Break;
    end;
end;

function GetCellNeighbours(MapRow: TMapRow; const SheetName: string;
  SheetNumber, Col, Row: integer): TCellNeighbours;
var
  i: integer;
begin
  Result := [];
  for i := 0 to MapRow.Count - 1 do begin
    if not (cnLeft in Result) and
       CellInRange(MapRow[i], SheetName, SheetNumber, Col - 1, Row) then
      Result := Result + [cnLeft];
    if not (cnRight in Result) and
       CellInRange(MapRow[i], SheetName, SheetNumber, Col + 1, Row) then
      Result := Result + [cnRight];
    if not (cnTop in Result) and
       CellInRange(MapRow[i], SheetName, SheetNumber, Col, Row - 1) then
      Result := Result + [cnTop];
    if not (cnBottom in Result) and
       CellInRange(MapRow[i], SheetName, SheetNumber, Col, Row + 1) then
      Result := Result + [cnBottom];
    if (cnLeft in Result) and (cnRight in Result) and (cnTop in Result) and
       (cnBottom in Result) then Break;
  end;
end;

procedure RemoveCellFromRow(MapRow: TMapRow; const SheetName: string;
  SheetNumber, Col, Row: integer);
var
  i: integer;
  str: string;
  R: TMapRange;
begin
  str := EmptyStr;
  for i := 0 to MapRow.Count - 1 do begin
    if not CellInRange(MapRow[i], SheetName, SheetNumber, Col, Row) then
      str := str + MapRow[i].AsString
    else begin
      case MapRow[i].RangeType of
        rtCol: begin
          if (MapRow[i].Row1 = Row) and (MapRow[i].Direction = rdDown) then
            MapRow[i].Row1 := MapRow[i].Row1 + 1
          else if (MapRow[i].Row1 = Row) and (MapRow[i].Direction = rdUp) then
            MapRow[i].Row1 := MapRow[i].Row1 - 1
          else if (MapRow[i].Row2 = Row) and (MapRow[i].Direction = rdDown) then
            MapRow[i].Row2 := MapRow[i].Row2 - 1
          else if (MapRow[i].Row2 = Row) and (MapRow[i].Direction = rdUp) then
            MapRow[i].Row2 := MapRow[i].Row2 + 1
          else if (MapRow[i].Row1 < Row) and (MapRow[i].Row2 > Row) and
                  (MapRow[i].Direction = rdDown) then begin
            R := TmapRange.Create(nil);
            try
              R.Col1 := MapRow[i].Col1;
              R.Col2 := MapRow[i].Col1;
              R.Row1 := MapRow[i].Row1;
              R.Row2 := Row - 1;
              R.Update;
              R.UpdateDirection;
              str := str + R.AsString;
            finally
              R.Free;
            end;
            MapRow[i].Row1 := Row + 1;
          end
          else if (MapRow[i].Row1 > Row) and (MapRow[i].Row2 < Row) and
                  (MapRow[i].Direction = rdUp) then begin
            R := TmapRange.Create(nil);
            try
              R.Col1 := MapRow[i].Col1;
              R.Col2 := MapRow[i].Col1;
              R.Row1 := MapRow[i].Row1;
              R.Row2 := Row + 1;
              R.Update;
              R.UpdateDirection;
              str := str + R.AsString;
            finally
              R.Free;
            end;
            MapRow[i].Row1 := Row - 1;
          end;
          str := str + MapRow[i].AsString;
        end;
        rtRow: begin
          if (MapRow[i].Col1 = Col) and (MapRow[i].Direction = rdDown) then
            MapRow[i].Col1 := MapRow[i].Col1 + 1
          else if (MapRow[i].Col1 = Col) and (MapRow[i].Direction = rdUp) then
            MapRow[i].Col1 := MapRow[i].Col1 - 1
          else if (MapRow[i].Col2 = Col) and (MapRow[i].Direction = rdDown) then
            MapRow[i].Col2 := MapRow[i].Col2 - 1
          else if (MapRow[i].Col2 = Col) and (MapRow[i].Direction = rdUp) then
            MapRow[i].Col2 := MapRow[i].Col2 + 1
          else if (MapRow[i].Col1 < Col) and (MapRow[i].Col2 > Col) and
                  (MapRow[i].Direction = rdDown) then begin
            R := TmapRange.Create(nil);
            try
              R.Col1 := MapRow[i].Col1;
              R.Col2 := Col - 1;
              R.Row1 := MapRow[i].Row1;
              R.Row2 := MapRow[i].Row1;
              R.Update;
              R.UpdateDirection;
              str := str + R.AsString;
            finally
              R.Free;
            end;
            MapRow[i].Col1 := Col + 1;
          end
          else if (MapRow[i].Col1 > Col) and (MapRow[i].Col2 < Col) and
                  (MapRow[i].Direction = rdUp) then begin
            R := TmapRange.Create(nil);
            try
              R.Col1 := MapRow[i].Col1;
              R.Col2 := Col + 1;
              R.Row1 := MapRow[i].Row1;
              R.Row2 := MapRow[i].Row1;
              R.Update;
              R.UpdateDirection;
              str := str + R.AsString;
            finally
              R.Free;
            end;
            MapRow[i].Col1 := Col - 1;
          end;
          str := str + MapRow[i].AsString;
        end;
      end;
    end;
  end;
  MapRow.AsString := str;
end;

procedure Str2ColRow(const Str: string; var ACol, ARow: integer);
var
  c, r: string;
  i: integer;
begin
  c := EmptyStr;
  r := EmptyStr;
  i := 1;
  while (i <= Length(Str)) and (Pos(Str[i], LETTERS) > 0) do begin
    c := c + Str[i];
    Inc(i);
  end;
  while (i <= Length(Str)) and (Pos(Str[i], NUMBERS) > 0) do begin
    r := r + Str[i];
    Inc(i);
  end;
  ACol := Letter2Col(c);
  ARow := Number2Row(r);
end;

procedure Str2Range(const Str: string; var ACol1, ARow1, ACol2, ARow2: integer);
begin
  Str2ColRow(Copy(Str, 1, Pos('-', Str) - 1), ACol1, ARow1);
  Str2ColRow(Copy(Str, Pos('-', Str) + 1, Length(Str) - Pos('-', Str) + 1), ACol2, ARow2);
end;

function GetRangeType(const Str: string): TRangeType;
var
  c1, r1, c2, r2: integer;
begin
  Result := rtCell;
  if Pos('-', Str) = 0 then Exit
  else begin
    Str2Range(Str, c1, r1, c2, r2);
    if c1 = c2 then Result := rtCol
    else if r1 = r2 then Result := rtRow;
  end;
end;

function OptimizeString(const Str: string): string;

function CanBeUnion(const Str1, Str2: string; var Str3: string): boolean;
var
  rt1, rt2: TRangeType;
  c1, r1, c2, r2, c3, r3, c4, r4: integer;
begin
  Result := false;

  rt1 := GetRangeType(Str1);
  rt2 := GetRangeType(Str2);

  if rt1 = rtCell then Str2ColRow(Str1, c1, r1)
  else Str2Range(Str1, c1, r1, c2, r2);
  if rt2 = rtCell then Str2ColRow(Str2, c4, r4)
  else Str2Range(Str2, c3, r3, c4, r4);

  case rt1 of
    rtCell: case rt2 of
              rtCell: Result := ((c1 = c4) and (Abs(r4 - r1) = 1)) or
                                ((r1 = r4) and (Abs(c1 - c4) = 1));
              rtCol : Result := (c1 = c3) and (Abs(r3 - r1) = 1);
              rtRow : Result := (r1 = r3) and (Abs(c3 - c1) = 1);
            end;
    rtCol : case rt2 of
              rtCell: Result := (c2 = c4) and (Abs(r4 - r2) = 1);
              rtCol : Result := (c1 = c3) and (Abs(r3 - r2) = 1);
            end;
    rtRow : case rt2 of
              rtCell: Result := (r2 = r4) and (Abs(c4 - c2) = 1);
              rtRow : Result := (r1 = r3) and (Abs(c3 - c2) = 1);
            end;
  end;
  if Result then begin
    Str3 := Col2Letter(c1) + Row2Number(r1);;
    if ((c1 = c4) and (r1 <> r4)) or ((r1 = r4) or (c1 <> c4)) then
      Str3 := Str3 + '-' + Col2Letter(c4) + Row2Number(r4);
  end;
end;

var
  p, p1, p11, p2: string;
  i, j: integer;
begin
  Result := Trim(Str);
  if Result = EmptyStr then Exit;
//!!!  if not CheckStringOfCells(nil, Result) then Exit;
  if Str[Length(Result)] <> ';' then Result := Str + ';';
  Result := UpperCase(Result);

  i := 1;
  j := 1;
  while i <= Length(Result) do begin
    p1 := EmptyStr;
    p2 := EmptyStr;
    while (i <= Length(Result)) and (Result[i] <> ';') do begin
      p1 := p1 + Result[i];
      Inc(i);
    end;
    p11 := CheckRange(p1);
    if CompareText(p11, p1) <> 0 then begin
      Delete(Result, i - Length(p1), Length(p1));
      Insert(p11, Result, i - Length(p1));
      Inc(i, Length(p11) - Length(p1));
    end;
    Inc(i);
    while (i <= Length(Result)) and (Result[i] <> ';') do begin
      p2 := p2 + Result[i];
      Inc(i);
    end;
    p11 := CheckRange(p2);
    if CompareText(p11, p2) <> 0 then begin
      Delete(Result, i - Length(p2), Length(p2));
      Insert(p11, Result, i - Length(p2));
      Inc(i, Length(p11) - Length(p2));
    end;
    if (p1 = EmptyStr) or (p2 = EmptyStr) then Exit;
    if CanBeUnion(p1, p2, p) then begin
      p := p + ';';
      Delete(Result, j, i - j + 1);
      Insert(p, Result, j);
      i := j;
    end
    else begin
      i := i - Length(p2);
      j := i;
    end;
  end;
end;

function SkipFirstRows(const Str: string; Rows: integer): string;
var
  k, j, c1, r1, c2, r2: integer;
  ss: string;
begin
  Result := Str;
  if Rows <= 0 then Exit;
  k := 1;
  j := Pos(';', Copy(Result, k, Length(Str) - k + 1));
  if (j = 0) and (Length(Str) > 0) then
    j := Length(Str);
  ss := Copy(Result, k, j - k + 1);
  while ss <> EmptyStr do begin
    case GetRangeType(ss) of
      rtCol: begin
        Str2Range(ss, c1, r1, c2, r2);
        if Rows <= r2 - r1 then
          {$IFDEF VCL15}
          ss := Col2Letter(c1) + Row2Number(Trunc(MaxIntValue([1 + Rows, r1]))) + '-' +
                Col2Letter(c2) + Row2Number(Trunc(MaxIntValue([1 + Rows, r2]))) + ';'
          {$ELSE}
          ss := Col2Letter(c1) + Row2Number(Trunc(MaxValue([1 + Rows, r1]))) + '-' +
                Col2Letter(c2) + Row2Number(Trunc(MaxValue([1 + Rows, r2]))) + ';'
          {$ENDIF}
        else ss := EmptyStr;
        Delete(Result, k, j - k + 1);
        Insert(ss, Result, k);
        k := k + Length(ss);
      end;
      rtRow: begin
        Str2Range(ss, c1, r1, c2, r2);
        if r1 <= Rows then Delete(Result, k, j - k + 1)
        else k := k + Length(ss);
      end;
      rtCell: begin
        Str2ColRow(ss, c1, r1);
        if r1 <= Rows then Delete(Result, k, j - k + 1)
        else k := k + Length(ss);
      end;
    end;
    j := Pos(';', Copy(Result, k, Length(Result) - k + 1)) + k - 1;
    ss := Copy(Result, k, j - k + 1);
  end;
end;

function SkipFirstCols(const Str: string; Cols: integer): string;
var
  k, j, c1, r1, c2, r2: integer;
  ss: string;
begin
  Result := Str;
  if Cols <= 0 then Exit;
  k := 1;
  j := Pos(';', Copy(Result, k, Length(Result) - k + 1));
  if (j = 0) and (Length(Str) > 0) then
    j := Length(Str);
  ss := Copy(Result, k, j - k + 1);
  while ss <> EmptyStr do begin
    case GetRangeType(ss) of
      rtCol: begin
        Str2Range(ss, c1, r1, c2, r2);
        if c1 <= Cols then Delete(Result, k, j - k + 1)
        else k := k + Length(ss);
      end;
      rtRow: begin
        Str2Range(ss, c1, r1, c2, r2);
        if Cols <= c2 - c1 then
          {$IFDEF VCL15}
          ss := Col2Letter(Trunc(MaxIntValue([1 + Cols, c1]))) + Row2Number(r1) + '-' +
                Col2Letter(Trunc(MaxIntValue([1 + Cols, c2]))) + Row2Number(r2) + ';'
          {$ELSE}
          ss := Col2Letter(Trunc(MaxValue([1 + Cols, c1]))) + Row2Number(r1) + '-' +
                Col2Letter(Trunc(MaxValue([1 + Cols, c2]))) + Row2Number(r2) + ';'
          {$ENDIF}
        else ss := EmptyStr;
        Delete(Result, k, j - k + 1);
        Insert(ss, Result, k);
        k := k + Length(ss);
      end;
      rtCell: begin
        Str2ColRow(ss, c1, r1);
        if c1 <= Cols then Delete(Result, k, j - k + 1)
        else k := k + Length(ss);
      end;
    end;
    j := Pos(';', Copy(Result, k, Length(Result) - k + 1)) + k - 1;
    ss := Copy(Result, k, j - k + 1);
  end;
end;

function CheckRange(const Str: string): string;
var
  c1, r1, c2, r2: integer;
begin
  Result := Trim(Str);
  if Result = EmptyStr then Exit;
//!!!  if not CheckStringOfCells(nil, Result) then Exit;
  Result := UpperCase(Result);
  if Pos('-', Result) = 0 then Exit;
  Str2Range(Result, c1, r1, c2, r2);
  if (c1 = c2) and (r1 = r2) then
    Delete(Result, Pos('-', Result), Length(Result) - Pos('-', Result) + 1);
end;

{ TMapRange }

constructor TMapRange.Create(MapRow: TMapRow);
begin
  inherited Create;;
  FMapRow := MapRow;
  if Assigned(FMapRow) then
    FXLSFile := FMapRow.XLSFile;

  FDirection := rdUnknown;
  FSheetIDType := sitUnknown;
  FSheetNumber := 0;
  FSheetName := EmptyStr;
  FRow1 := 0;
  FCol1 := 0;
  FRow2 := 0;
  FCol2 := 0;
end;

procedure TMapRange.Arrange;
var
  SheetIndex, Index: integer;
begin
  Update;
  if not Assigned(FXLSFile) then Exit;

  if not FXLSFile.Loaded then
    FXLSFile.Load;

  if GetHasSheet then
  begin
    if FSheetNumber > 0 then
      SheetIndex := FSheetNumber - 1
    else
      SheetIndex := FXLSFile.Workbook.WorkSheets.IndexOfName(FSheetName);

    if SheetIndex = -1 then
      raise Exception.CreateFmt(sSheetNotFound, [FSheetName]);
  end
  else
    SheetIndex := 0;

  if FRangeType = rtRow then
  begin
    if not FXLSFile.Workbook.WorkSheets[SheetIndex].Rows.Find(FRow1 - 1, Index) then
      Exit;

    if FCol1 = 0 then
    begin
      if FDirection = rdDown then
        FCol1 := FXLSFile.Workbook.WorkSheets[SheetIndex].Rows[Index].MinCol + 1
      else if FDirection = rdUp then
        FCol1 := FXLSFile.Workbook.WorkSheets[SheetIndex].Rows[Index].MaxCol + 1;
    end;

    if FCol2 = 0 then
    begin
      if FDirection = rdDown then
        FCol2 := FXLSFile.Workbook.WorkSheets[SheetIndex].Rows[Index].MaxCol + 1
      else if FDirection = rdUp then
        FCol2 := FXLSFile.Workbook.WorkSheets[SheetIndex].Rows[Index].MinCol + 1;
    end;
  end
  else if FRangeType = rtCol then
  begin
    if not FXLSFile.Workbook.WorkSheets[SheetIndex].Cols.Find(FCol1 - 1, Index) then
      Exit;

    if FRow1 = 0 then
    begin
      if FDirection = rdDown then
        // pai 
        //FRow1 := FXLSFile.Workbook.WorkSheets[SheetIndex].Cols[Index].MinRow + 1
        FRow1 := 1
        // pai
      else if FDirection = rdUp then
        FRow1 := FXLSFile.Workbook.WorkSheets[SheetIndex].Cols[Index].MaxRow + 1;
    end;

    if FRow2 = 0 then
    begin
      if FDirection = rdDown then
        FRow2 := FXLSFile.Workbook.WorkSheets[SheetIndex].Cols[Index].MaxRow + 1
      else if FDirection = rdUp then
        FRow2 := FXLSFile.Workbook.WorkSheets[SheetIndex].Cols[Index].MinRow + 1;
    end;
  end;
  Update;
end;

function TMapRange.GetHasSheet: boolean;
begin
  Result := (FSheetName <> EmptyStr) or (FSheetNumber > 0)
end;

procedure TMapRange.Update;
var
  T1, T2, S: integer;
begin
  FRangeType := rtUnknown;
  if (FCol1 > 0) and (FRow1 > 0) and (FCol2 > 0) and (FRow2 > 0) and
     (FCol1 = FCol2) and (FRow1 = FRow2) then
     FRangeType := rtCell
  else if (FCol1 > 0) or (FRow1 > 0) or (FCol2 > 0) or (FRow2 > 0) then
  begin
    if (FCol1 = FCol2) and (FCol1 > 0) then
      FRangeType := rtCol
    else if (FRow1 = FRow2) and (FRow1 > 0) then
      FRangeType := rtRow
  end;

  FLength := 0;
  case FRangeType of
    rtCell:
      if (FRow1 <= SkipFirstRows) or (FCol1 <= SkipFirstCols)
        then FLength := 0
        else FLength := 1;
    rtRow :
      if FRow1 <= SkipFirstRows then
        FLength := 0
      else begin
        S := SkipFirstCols;

        if (FCol1 <= S) and (FCol2 <= S) then
          FLength := 0
        else begin
          if FCol1 <= S
            then T1 := S + 1
            else T1 := FCol1;
          if FCol2 <= S
            then T2 := S + 1
            else T2 := FCol2;
          FLength := Abs(T2 - T1) + 1;
        end;
        //FLength := Abs(FCol2 - FCol1) + 1;
      end;
    rtCol :
      if FCol1 <= SkipFirstCols then
        FLength := 0
      else begin
        S := SkipFirstRows;

        if (FRow1 <= S) and (FRow2 <= S) then
          FLength := 0
        else begin
          if FRow1 <= S
            then T1 := S + 1
            else T1 := FRow1;
          if FRow2 <= S
            then T2 := S + 1
            else T2 := FRow2;
          FLength := Abs(T2 - T1) + 1;
        end;
        //FLength := Abs(FRow2 - FRow1) + 1;
      end;
  end;
end;

procedure TMapRange.Assign(Range: TMapRange);
begin
  FCol1 := Range.Col1;
  FRow1 := Range.Row1;
  FCol2 := Range.Col2;
  FRow2 := Range.Row2;
  FDirection := Range.Direction;
  FSheetIDType := Range.SheetIDType;
  FSheetNumber := Range.SheetNumber;
  FSheetName := Range.SheetName;
  Update;
end;

procedure TMapRange.UpdateDirection;
begin
  case FRangeType of
    rtCol:
      if (Row1 > 0) and (Row2 > 0) then
        if Row1 < Row2 then FDirection := rdDown
        else if Row1 > Row2 then FDirection := rdUp;
    rtRow:
      if (Col1 > 0) and (Col2 > 0) then
        if Col1 < Col2 then FDirection := rdDown
        else if Col1 > Col2 then FDirection := rdUp;
    rtCell: FDirection := rdUnknown;
  end;
end;

function TMapRange.GetAsString: string;
begin
  Result := EmptyStr;
  if SheetIDType <> sitUnknown then begin
    Result := Result + SHEET_START;
    if SheetIDType = sitNumber
      then Result := Result + SHEET_NUMBER + IntToStr(FSheetNumber)
      else Result := Result + FSheetName;
    Result := Result + SHEET_FINISH;
  end;
  case FRangeType of
    rtCell: Result := Result + Col2Letter(FCol1) + Row2Number(FRow1);
    rtCol: begin
      if (FRow1 > 0) and (FRow2 > 0) then
        Result := Result + Col2Letter(FCol1) + Row2Number(FRow1) +
          ARRAY_DELIMITER + Col2Letter(FCol2) + Row2Number(FRow2)
      else
        case FDirection of
          rdDown: begin
            if (FRow1 = 0) and (FRow2 = 0) then
              Result := Result + Col2Letter(FCol1) + ARRAY_DELIMITER + COLFINISH
            else if (FRow1 = 0) and (FRow2 > 0) then
              Result := Result + COLSTART + ARRAY_DELIMITER +
                Col2Letter(FCol2) + Row2Number(FRow2)
            else if (FRow1 > 0) and (FRow2 = 0) then
              Result := Result + Col2Letter(FCol1) + Row2Number(FRow1) +
                ARRAY_DELIMITER + COLFINISH;
          end;
          rdUp: begin
            if (FRow1 = 0) and (FRow2 = 0) then
              Result := Result + Col2Letter(FCol1) + ARRAY_DELIMITER + COLSTART
            else if (FRow1 = 0) and (FRow2 > 0) then
              Result := Result + COLFINISH + ARRAY_DELIMITER +
                Col2Letter(FCol2) + Row2Number(FRow2)
            else if (FRow1 > 0) and (FRow2 = 0) then
              Result := Result + Col2Letter(FCol1) + Row2Number(FRow1) +
                ARRAY_DELIMITER + COLSTART;
          end;
        end;
    end;
    rtRow: begin
      if (FCol1 > 0) and (FCol2 > 0) then
        Result := Result + Col2Letter(FCol1) + Row2Number(FRow1) +
          ARRAY_DELIMITER + Col2Letter(FCol2) + Row2Number(FRow2)
      else
        case FDirection of
          rdDown: begin
            if (FCol1 = 0) and (FCol2 = 0) then
              Result := Result + Row2Number(FRow1) + ARRAY_DELIMITER + ROWFINISH
            else if (FCol1 = 0) and (FCol2 > 0) then
              Result := Result + ROWSTART + ARRAY_DELIMITER +
                Col2Letter(FCol2) + Row2Number(FRow2)
            else if (FCol1 > 0) and (FCol2 = 0) then
              Result := Result + Col2Letter(FCol1) + Row2Number(FRow1) +
                ARRAY_DELIMITER + ROWFINISH;
          end;
          rdUp: begin
            if (FCol1 = 0) and (FCol2 = 0) then
              Result := Result + Row2Number(FRow1) + ARRAY_DELIMITER + ROWSTART
            else if (FCol1 = 0) and (FCol2 > 0) then
              Result := Result + ROWFINISH + ARRAY_DELIMITER +
                Col2Letter(FCol2) + Row2Number(FRow2)
            else if (FCol1 > 0) and (FCol2 = 0) then
              Result := Result + Col2Letter(FCol1) + Row2Number(FRow1) +
                ARRAY_DELIMITER + ROWSTART;
          end;
        end;
    end;
  end;
  if Result <> EmptyStr then Result := Result + RANGE_DELIMITER;
end;

function TMapRange.GetSkipFirstRows: integer;
begin
  Result := 0;
  if Assigned(MapRow) and Assigned(MapRow.MapRowList) then
    Result := MapRow.MapRowList.SkipFirstRows;
end;

function TMapRange.GetSkipFirstCols: integer;
begin
  Result := 0;
  if Assigned(MapRow) and Assigned(MapRow.MapRowList) then
    Result := MapRow.MapRowList.SkipFirstCols;
end;

{ TMapRow }

constructor TMapRow.Create(MapRowList: TMapRowList);
begin
  inherited Create;
  FMapRowList := MapRowList;
  if Assigned(FMapRowList) then
    FXLSFile := FMapRowList.XLSFile;
end;

function TMapRow.GetItems(Index: integer): TMapRange;
begin
  Result := TMapRange(inherited Items[Index]);
end;

procedure TMapRow.SetItems(Index: integer; Value: TMapRange);
begin
  inherited Items[Index] := Value;
end;

function TMapRow.Add(Item: TMapRange): integer;
begin
  Result := inherited Add(Item);
  Item.Arrange;
end;

procedure TMapRow.Delete(Index: integer);
var
  O: TObject;
begin
  if Assigned(Items[Index]) then
  begin
    O := TObject(Items[Index]);
    ObjFreeAndNil(O);
  end;
  inherited;
end;

procedure TMapRow.Update;
var
  i: integer;
begin
  FLength := 0;
  for i := 0 to Count - 1 do
    Inc(FLength, Items[i].Length);
end;

function TMapRow.GetCellValue(AbsoluteIndex: integer): WideString;
var
  i: integer;
  RangeIndex: integer;
  S, C, R: integer;
  WorkCell: TbiffCell;
begin
  Result := EmptyStr;

  if not Assigned(FXLSFile) then Exit;
  if (AbsoluteIndex <= 0) or (AbsoluteIndex > FLength) then Exit;

  RangeIndex := -1;
  for i := 0 to Count - 1 do
  begin
    if AbsoluteIndex > Items[i].Length then
      Dec(AbsoluteIndex, Items[i].Length)
    else begin
      RangeIndex := i;
      Break;
    end;
  end;
  if RangeIndex = -1 then Exit;

  S := 0;
  if Items[RangeIndex].HasSheet then
  begin
    if Items[RangeIndex].SheetNumber > 0
      then S := Items[RangeIndex].SheetNumber - 1
      else S := FXLSFile.Workbook.WorkSheets.IndexOfName(Items[RangeIndex].SheetName);
  end;

  C := -1; R := -1;
  case Items[RangeIndex].FRangeType of
    rtCell: begin
      C := Items[RangeIndex].Col1 - 1;
      R := Items[RangeIndex].Row1 - 1;
    end;
    rtCol: begin
      C := Items[RangeIndex].Col1 - 1;
      case Items[RangeIndex].FDirection of
        rdDown:
          if SkipFirstRows > 0
            then R := SkipFirstRows + 1 + (AbsoluteIndex - 1) - 1
            else R := Items[RangeIndex].Row1 + (AbsoluteIndex - 1) - 1;
        rdUp:
          R := Items[RangeIndex].Row1 - (AbsoluteIndex - 1) - 1;
      end
    end;
    rtRow: begin
      R := Items[RangeIndex].Row1 - 1;
      case Items[RangeIndex].FDirection of
        rdDown:
          if SkipFirstCols > 0
            then C := (SkipFirstCols + 1) + (AbsoluteIndex - 1) - 1
            else C := Items[RangeIndex].Col1 + (AbsoluteIndex - 1) - 1;
        rdUp: C := Items[RangeIndex].Col1 - (AbsoluteIndex - 1) - 1;
      end
    end;
  end;

  WorkCell := FXLSFile.Workbook.WorkSheets[S].Cells[R, C];
  if Assigned(WorkCell) then
    Result := WorkCell.AsString;
end;

procedure TMapRow.Optimize;
var
  i: integer;
begin
  for i := Count - 1 downto 0 do
    if i > 0 then
    begin
      if not ((Items[i - 1].SheetNumber = Items[i].SheetNumber) or
              (Items[i - 1].SheetName = Items[i].SheetName)) then Continue;

      case Items[i - 1].RangeType of
        rtCol:
          if (((Items[i].RangeType = rtCol) and
               (Items[i - 1].Direction = Items[i].Direction)) or
              (Items[i].RangeType = rtCell)) and
             (Items[i - 1].Col1 = Items[i].Col1) and
             (((Items[i - 1].Direction = rdDown) and
               (Items[i - 1].Row2 = Items[i].Row1 - 1)) or
              ((Items[i - 1].Direction = rdUp) and
               (Items[i - 1].Row2 = Items[i].Row1 + 1))) then
          begin
            Items[i - 1].Row2 := Items[i].Row2;
            Delete(i);
          end;
        rtRow:
          if (((Items[i].RangeType = rtRow) and
               (Items[i - 1].Direction = Items[i].Direction)) or
              (Items[i].RangeType = rtCell)) and
             (Items[i - 1].Row1 = Items[i].Row1) and
             (((Items[i - 1].Direction = rdDown) and
               (Items[i - 1].Col2 = Items[i].Col1 - 1)) or
              ((Items[i - 1].Direction = rdUp) and
               (Items[i - 1].Col2 = Items[i].Col1 + 1))) then
          begin
            Items[i - 1].Col2 := Items[i].Col2;
            Delete(i);
          end;
        rtCell:
          case Items[i].RangeType of
            rtCol:
              if Items[i - 1].Col1 = Items[i].Col1 then
              begin
                if ((Items[i].Direction = rdDown) and
                    (Items[i - 1].Row1 = Items[i].Row2 + 1)) or
                   ((Items[i].Direction = rdUp) and
                    (Items[i - 1].Row1 = Items[i].Row2 - 1)) then
                begin
                  Items[i - 1].Row1 := Items[i].Row1;

                  Items[i - 1].Update;
                  Items[i - 1].UpdateDirection;

                  Delete(i);
                end
                else if ((Items[i].Direction = rdDown) and
                         (Items[i - 1].Row1 = Items[i].Row1 - 1)) or
                        ((Items[i].Direction = rdUp) and
                         (Items[i - 1].Row1 = Items[i].Row1 + 1)) then
                begin
                  Items[i - 1].Row2 := Items[i].Row2;

                  Items[i - 1].Update;
                  Items[i - 1].UpdateDirection;

                  Delete(i);
                end
              end;
            rtRow:
              if Items[i - 1].Row1 = Items[i].Row1 then
              begin
                if ((Items[i].Direction = rdDown) and
                    (Items[i - 1].Col1 = Items[i].Col2 + 1)) or
                   ((Items[i].Direction = rdUp) and
                    (Items[i - 1].Col1 = Items[i].Col2 - 1)) then
                begin
                  Items[i - 1].Col1 := Items[i].Col1;

                  Items[i - 1].Update;
                  Items[i - 1].UpdateDirection;

                  Delete(i);
                end
                else if ((Items[i].Direction = rdDown) and
                         (Items[i - 1].Col1 = Items[i].Col1 - 1)) or
                        ((Items[i].Direction = rdUp) and
                         (Items[i - 1].Col1 = Items[i].Col1 + 1)) then
                begin
                  Items[i - 1].Col2 := Items[i].Col2;

                  Items[i - 1].Update;
                  Items[i - 1].UpdateDirection;

                  Delete(i);
                end
              end;
            rtCell:
              if (Items[i - 1].Col1 = Items[i].Col1) and
                 ((Items[i - 1].Row1 = Items[i].Row1 + 1) or
                  (Items[i - 1].Row1 = Items[i].Row1 - 1)) then
              begin
                Items[i - 1].Row2 := Items[i].Row1;

                Items[i - 1].Update;
                Items[i - 1].UpdateDirection;

                Delete(i);
              end
              else if (Items[i - 1].Row1 = Items[i].Row1) and
                      ((Items[i - 1].Col1 = Items[i].Col1 + 1) or
                       (Items[i - 1].Col1 = Items[i].Col1 - 1)) then
              begin
                Items[i - 1].Col2 := Items[i].Col1;

                Items[i - 1].Update;
                Items[i - 1].UpdateDirection;

                Delete(i);
              end
          end
      end;
    end
end;

function TMapRow.IndexOfRange(const RangeStr: string): integer;
var
  i: integer;
begin
  Result := -1;
  for i := 0 to Count - 1 do
    if AnsiCompareText(RangeStr, Items[i].AsString) = 0 then
    begin
      Result := i;
      Break;
    end;
end;

function TMapRow.GetAsString: string;
var
  i: integer;
begin
  Result := EmptyStr;
  for i := 0 to Count - 1 do
    Result := Result + Items[i].AsString;
end;

procedure TMapRow.SetAsString(const Value: string);
begin
  Clear;
  ParseMapString(Value, Self);
end;

function TMapRow.GetSkipFirstRows: integer;
begin
 Result := 0;
 if Assigned(MapRowList) then
   Result := MapRowList.SkipFirstRows;
end;

function TMapRow.GetSkipFirstCols: integer;
begin
 Result := 0;
 if Assigned(MapRowList) then
   Result := MapRowList.SkipFirstCols;
end;

{ TMapRowList }

constructor TMapRowList.Create(XLSFile: TxlsFile);
begin
  inherited Create;
  FXLSFile := XLSFile;
  FMinRow := -1;
  FMaxRow := -1;
  FSkipFirstCols := 0;
  FSkipFirstRows := 0;
end;

destructor TMapRowList.Destroy;
var
  I, J: Integer;
begin
  for I := Count - 1 downto 0 do
  begin
    for J := Items[I].Count - 1 downto 0 do
    begin
      Items[i].Delete(j);
    end;
    Items[i].Clear;
    Items[i].Free;
  end;
  inherited Destroy;
end;

function TMapRowList.GetItems(Index: integer): TMapRow;
begin
  Result := TMapRow(inherited Items[Index]);
end;

procedure TMapRowList.SetItems(Index: integer; Value: TMapRow);
begin
  inherited Items[Index] := Value;
end;

function TMapRowList.Add(Item: TMapRow): integer;
begin
  Result := inherited Add(Item);
end;

procedure TMapRowList.Delete(Index: integer);
begin
  if Assigned(Items[Index]) then
    TMapRow(Items[Index]).Free;
  inherited;
end;

procedure TMapRowList.Update;
var
  i: integer;
begin
  FMinRow := -1;
  FMaxRow := -1;
  for i := 0 to Count - 1 do
  begin
    if (FMinRow = -1) or (Items[i].Length < Items[FMinRow].Length) then
      FMinRow := i;
    if (FMaxRow = -1) or (Items[i].Length > Items[FMaxRow].Length) then
      FMaxRow := i;
  end;
end;

end.
