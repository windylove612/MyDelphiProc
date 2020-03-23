unit QImport3XlsxMapParser;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF XLSX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Classes,
    System.Types,
  {$ELSE}
    Classes,
    Types,
  {$ENDIF}
  QImport3StrTypes;
const
  RangeDelim = ';';
  SheetPrefix = '[';
  SheetPostfix = ']';
  SheetIndexPrefix = ':';
  DefSheetPrefixPos = 1;
  DefSheetIndexPrefixPos = 2;
  RowStartFlag = 'ROWSTART';
  RowFinishFlag = 'ROWFINISH';
  ColStartFlag = 'COLSTART';
  ColFinishFlag = 'COLFINISH';
  LeftRightRangesDelim = '-';
  FieldNameDelim = '=';
  BeginColStr = 'A1';
  EndColStr = 'XFD1';
  BeginRowStr = 'A1';
  EndRowStr = 'A1048576';
  BeginRowInt = 1;
  BeginColInt = 1;
  EndColInt = 16384;
  EndRowInt = 1048576;

type
  TXlsxRangeDirection = (xdNone, xdCell, xdColDown, xdColUp, xdRowRight,
    xdRowLeft);

  TXlsxColumns = BeginColInt..EndColInt;
  TXlsxRows = BeginRowInt..EndRowInt;
  TXlsxCellPoint = packed record
    Col: TXlsxColumns;
    Row: TXlsxRows;
  end;

  TXlsxRangeItem = class
  private
    FSheetName: qiString;
    FSheetIndex,
    FFirstCol,
    FFirstRow,
    FLastCol,
    FLastRow: Integer;
    function GetDirection: TXlsxRangeDirection;
    function GetCellRange(const Col, Row: Integer): qiString;
    function GetAsString: qiString;
    function SheetInRange(const ASheetIndex: Integer;
      const ASheetName: qiString): Boolean;
  public
    property SheetName: qiString read FSheetName write FSheetName;
    property SheetIndex: Integer read FSheetIndex write FSheetIndex;
    property FirstCol: Integer read FFirstCol write FFirstCol;
    property FirstRow: Integer read FFirstRow write FFirstRow;
    property LastCol: Integer read FLastCol write FLastCol;
    property LastRow: Integer read FLastRow write FLastRow;
    property Direction: TXlsxRangeDirection read GetDirection;
    property AsString: qiString read GetAsString;
    function CellInRange(const CellPoint: TXlsxCellPoint;
      const SheetIndex: Integer; const SheetName: qiString): Boolean;
    procedure Reset;
  end;

  TXlsxRangeItems = class
  private
    FList: TList;
    function GetItems(Index: Integer): TXlsxRangeItem;
    function GetCount: Integer;
    function GetAsString: qiString;
  public
    constructor Create;
    destructor Destroy; override;

    property Items[Index: Integer]: TXlsxRangeItem read GetItems; default;
    property Count: Integer read GetCount;
    function Add: TXlsxRangeItem;
    procedure Delete(const Index: Integer);
    procedure Clear;
    property AsString: qiString read GetAsString;
    function CellInRanges(const CellPoint: TXlsxCellPoint;
      const SheetIndex: Integer; const SheetName: qiString): Boolean;
    procedure Reset;
  end;

  TXlsxMapParserItem = class
  private
    FRanges: TXlsxRangeItems;
    FCounter: Integer;
    procedure DeleteCharFromBegin(const AChar: qiChar; var AString: qiString);
    procedure Parse(const MapString: qiString; const SkipRows, SkipCols: Integer);
    function Col2Int(const ColData: qiString): Integer;
    function Row2Int(const RowData: qiString): Integer;
    function GetRangeString(const MapString: qiString;
      var RangeString: qiString): Boolean;
    procedure DeleteFieldDescribtion(var MapString: qiString);
    procedure DeleteSheetDescribtion(var RangeString: qiString);
    function GetLeftSide(const RangeString: qiString): qiString;
    function GetRightSide(const RangeString: qiString): qiString;
    function GetColData(const RangeSide: qiString): qiString;
    function GetRowData(const RangeSide: qiString): qiString;
    procedure TruncMapString(var MapString: qiString);
    function GetSheetName(const RangeString: qiString): qiString;
    function GetSheetIndex(const RangeString: qiString): Integer;
    function GetFirstCol(const RangeString: qiString): Integer;
    function GetFirstRow(const RangeString: qiString): Integer;
    function GetLastCol(const RangeString: qiString): Integer;
    function GetLastRow(const RangeString: qiString): Integer;
    procedure FillRange(const Range: TXlsxRangeItem;
      const RangeString: qiString);
    procedure SetSkipRange(const Range: TXlsxRangeItem;
      const SkipRows, SkipCols: Integer);
    procedure RepairOldRange(var RangeString: qiString);
  public
    constructor Create(const MapString: qiString;
      const SkipRows, SkipCols: Integer);
    destructor Destroy; override;
    property Ranges: TXlsxRangeItems read FRanges;
    function GetNextRange(var Row, Col, SheetIndex: Integer;
      var SheetName: qiString): Boolean;
    procedure PackRanges;
  end;

  TXlsxMapParserItems = class
  private
    FList: TList;
    function GetItems(Index: Integer): TXlsxMapParserItem;
    function GetCount: Integer;
  public
    constructor Create;
    destructor Destroy; override;

    property Items[Index: Integer]: TXlsxMapParserItem read GetItems; default;
    property Count: Integer read GetCount;
    function Add(const MapString: qiString;
      const SkipRows, SkipCols: Integer): TXlsxMapParserItem;
    procedure Delete(const Index: Integer);
    procedure Clear;
  end;

function GetColIdFromString(const Index: qiString): Integer;
function GetRowIdFromString(const Index: qiString): Integer;
function GetColNameFromIndex(const ColIndex: Integer): qiString;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF XLSX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.SysUtils,
  {$ELSE}
    SysUtils,
  {$ENDIF}
  QImport3Common,
  QImport3StrIDs,
  QImport3;

function GetColIdFromColName(const ColIndex: qiString): Integer;
begin
  Result := 0;
  if ColIndex <> '' then
    case Length(ColIndex) of
      1: Result := Ord(ColIndex[1]) - 64;
      2: Result := (Ord(ColIndex[1]) - 64)*26 + (Ord(ColIndex[2]) - 64);
      3: Result := (Ord(ColIndex[1]) - 64)*676 + ((Ord(ColIndex[2]) - 64)*26 + Ord(ColIndex[3]) - 64);
    end;
end;

function GetColIdFromString(const Index: qiString): Integer;
var
  ColValue: qiString;
  i: Integer;
begin
  Result := 0;
  for i := 1 to Length(Index) do
    if QImport3Common.CharInSet(Index[i],['0'..'9']) then
    begin
      ColValue := Copy(Index, 1, i - 1);
      Result := GetColIdFromColName(ColValue);
      Break;
    end;
end;

function GetRowIdFromString(const Index: qiString): Integer;
var
  i: Integer;
begin
  Result := 0;
  for i := 1 to Length(Index) do
    if QImport3Common.CharInSet(Index[i], ['0'..'9']) then
    begin
      Result := StrToIntDef(Copy(Index, i, Length(Index)), 0);
      Break;
    end;
end;

function Byte2Str(const AByte: Byte): qiString;
begin
  if AByte = 0 then
    Result := ''
  else
    Result := Chr(AByte + 64);
end;

function GetColNameFromIndex(const ColIndex: Integer): qiString;
var
  PartDiv,
  PartMod: Byte;
  ColIndexTmp: Integer;

begin
  Result := '';
  ColIndexTmp := ColIndex;

  PartDiv := ColIndexTmp div 676;
  ColIndexTmp := ColIndexTmp mod 676;
  if ColIndexTmp = 0 then
  begin
    Dec(PartDiv);
    ColIndexTmp := 676;
  end;

  Result := Byte2Str(PartDiv);

  PartDiv := ColIndexTmp div 26;
  PartMod := ColIndexTmp mod 26;
  if PartMod = 0 then
  begin
    Dec(PartDiv);
    PartMod := 26;
  end;

  Result := Result + Byte2Str(PartDiv) +
    Byte2Str(PartMod);
end;

{ TQImport3XlsxMapParser }

constructor TXlsxMapParserItem.Create(const MapString: qiString;
  const SkipRows, SkipCols: Integer);
begin
  FRanges := TXlsxRangeItems.Create;
  FCounter := 0;
  Parse(MapString, SkipRows, SkipCols);
end;

procedure TXlsxMapParserItem.DeleteFieldDescribtion(var MapString: qiString);
var
  FieldNameDelimPos: Integer;
begin
  FieldNameDelimPos := Pos(FieldNameDelim, MapString);
  if FieldNameDelimPos > 0 then
    Delete(MapString, 1, FieldNameDelimPos);
end;

procedure TXlsxMapParserItem.DeleteSheetDescribtion(var RangeString: qiString);
var
  SheetPostfixPos: Integer;
begin
  SheetPostfixPos := Pos(SheetPostfix, RangeString);

  if SheetPostfixPos > 0  then
    Delete(RangeString, 1, SheetPostfixPos);
end;

destructor TXlsxMapParserItem.Destroy;
begin
  FRanges.Free;
  inherited;
end;

function TXlsxMapParserItem.GetNextRange(var Row, Col, SheetIndex: Integer;
  var SheetName: qiString): Boolean;
var
  Range: TXlsxRangeItem;
begin
  Row := 0;
  Col := 0;
  SheetIndex := 0;
  SheetName := '';

  Result := Ranges.Count > 0;
  if not Result then
    Exit;

  Range := Ranges[0];
  
  SheetName := Range.SheetName;
  SheetIndex := Range.SheetIndex;

  case Range.Direction of
    xdCell:
    begin
      Col := Range.FirstCol;
      Row := Range.FirstRow;
      PackRanges;
    end;  
    xdColDown:
    begin
      Col := Range.FirstCol;
      Row := Range.FirstRow + FCounter;
      Inc(FCounter);

      if Row = Range.LastRow then
        PackRanges;

      if Row > Range.LastRow then
      begin
        PackRanges;
        Result := GetNextRange(Row, Col, SheetIndex, SheetName);
      end;
    end;
    xdColUp:
    begin
      Col := Range.FirstCol;
      Row := Range.FirstRow - FCounter;
      Inc(FCounter);

      if Row = Range.LastRow then
        PackRanges;

      if Row < Range.LastRow then
      begin
        PackRanges;
        Result := GetNextRange(Row, Col, SheetIndex, SheetName);
      end;
    end;
    xdRowRight:
    begin
      Row := Range.FirstRow;
      Col := Range.FirstCol + FCounter;
      Inc(FCounter);

      if Col = Range.LastCol then
        PackRanges;

      if Col > Range.LastCol then
      begin
        PackRanges;
        Result := GetNextRange(Row, Col, SheetIndex, SheetName);
      end;
    end;
    xdRowLeft:
    begin
      Row := Range.FirstRow;
      Col := Range.FirstCol - FCounter;
      Inc(FCounter);

      if Col = Range.LastCol then
        PackRanges;

      if Col < Range.LastCol then
      begin
        PackRanges;
        Result := GetNextRange(Row, Col, SheetIndex, SheetName);
      end;
    end;
  end;
end;

procedure TXlsxMapParserItem.PackRanges;
begin
  FCounter := 0;
  if Ranges.Count > 0 then
    Ranges.Delete(0);
end;

function TXlsxMapParserItem.GetColData(const RangeSide: qiString): qiString;
begin
  if CompareText(RangeSide, RowStartFlag) = 0 then
    Result := BeginColStr
  else
    if CompareText(RangeSide, RowFinishFlag) = 0 then
      Result := EndColStr
    else
      if (CompareText(RangeSide, ColStartFlag) = 0) or
        (CompareText(RangeSide, ColFinishFlag) = 0) then
        Result := ''
      else
        Result := RangeSide;
end;

function TXlsxMapParserItem.GetRowData(const RangeSide: qiString): qiString;
begin
  if CompareText(RangeSide, ColStartFlag) = 0 then
    Result := BeginRowStr
  else
    if CompareText(RangeSide, ColFinishFlag) = 0 then
      Result := EndRowStr
    else
      if (CompareText(RangeSide, RowStartFlag) = 0) or
        (CompareText(RangeSide, RowFinishFlag) = 0) then
        Result := ''
      else
        Result := RangeSide;
end;

function TXlsxMapParserItem.GetFirstCol(const RangeString: qiString): Integer;
var
  LeftSide,
  ColData: qiString;
begin
  LeftSide := GetLeftSide(RangeString);
  ColData := GetColData(LeftSide);
  Result := Col2Int(ColData);
end;

function TXlsxMapParserItem.GetFirstRow(const RangeString: qiString): Integer;
var
  LeftSide,
  RowData: qiString;
begin
  LeftSide := GetLeftSide(RangeString);
  RowData := GetRowData(LeftSide);
  Result := Row2Int(RowData);
end;

function TXlsxMapParserItem.GetLastCol(const RangeString: qiString): Integer;
var
  RightSide,
  ColData: qiString;
begin
  RightSide := GetRightSide(RangeString);
  ColData := GetColData(RightSide);
  Result := Col2Int(ColData);
end;

function TXlsxMapParserItem.GetLastRow(const RangeString: qiString): Integer;
var
  RightSide,
  RowData: qiString;
begin
  RightSide := GetRightSide(RangeString);
  RowData := GetRowData(RightSide);
  Result := Row2Int(RowData);
end;

function TXlsxMapParserItem.GetLeftSide(const RangeString: qiString): qiString;
var
  LeftRightSidesDelimPos: Integer;
begin
  Result := RangeString;

  DeleteSheetDescribtion(Result);

  LeftRightSidesDelimPos := Pos(LeftRightRangesDelim, Result);

  if LeftRightSidesDelimPos > 0  then
    Result := Copy(Result, 1, LeftRightSidesDelimPos - 1);
end;

function TXlsxMapParserItem.Col2Int(const ColData: qiString): Integer;
begin
  Result := GetColIdFromString(ColData);
  if Result = 0 then
    if not TryStrToInt(ColData, Result) then
      Result := 0;
end;

function TXlsxMapParserItem.Row2Int(const RowData: qiString): Integer;
begin
  Result := GetRowIdFromString(RowData);
end;

function TXlsxMapParserItem.GetRangeString(const MapString: qiString;
  var RangeString: qiString): Boolean;
var
  RangeDelimPos: Integer;
begin
  RangeString := '';
  RangeDelimPos := Pos(RangeDelim, MapString);
  if RangeDelimPos > 0 then
    RangeString := Copy(MapString, 1, RangeDelimPos - 1);

  Result := RangeString <> '';
end;

function TXlsxMapParserItem.GetRightSide(const RangeString: qiString): qiString;
var
  LeftRightSidesDelimPos: Integer;
begin
  Result := RangeString;

  DeleteSheetDescribtion(Result);

  LeftRightSidesDelimPos := Pos(LeftRightRangesDelim, Result);

  if LeftRightSidesDelimPos > 0  then
    Delete(Result, 1, LeftRightSidesDelimPos)
  else
    Result := '';
end;

function TXlsxMapParserItem.GetSheetIndex(const RangeString: qiString): Integer;
var
  SheetPrefixPos,
  SheetPostfixPos,
  SheetIndexPrefixPos: Integer;
  SheetIndexStr: qiString;
begin
  Result := -1;
  SheetPrefixPos := Pos(SheetPrefix, RangeString);
  SheetPostfixPos := Pos(SheetPostfix, RangeString);
  SheetIndexPrefixPos := Pos(SheetIndexPrefix, RangeString);
  
  if (SheetPrefixPos = DefSheetPrefixPos) and
    (SheetIndexPrefixPos = DefSheetIndexPrefixPos) and
    (SheetPostfixPos > DefSheetIndexPrefixPos) then
  begin
    SheetIndexStr := Copy(RangeString, SheetIndexPrefixPos + 1, SheetPostfixPos -
      SheetIndexPrefixPos - 1);
    if not TryStrToInt(SheetIndexStr, Result) then
      Result := -1;
  end;
end;

function TXlsxMapParserItem.GetSheetName(const RangeString: qiString): qiString;
var
  SheetPrefixPos,
  SheetPostfixPos,
  SheetIndexPrefixPos: Integer;
begin
  Result := '';
  SheetPrefixPos := Pos(SheetPrefix, RangeString);
  SheetPostfixPos := Pos(SheetPostfix, RangeString);
  SheetIndexPrefixPos := Pos(SheetIndexPrefix, RangeString);

  if (SheetPrefixPos = DefSheetPrefixPos) and
    (SheetIndexPrefixPos <> DefSheetIndexPrefixPos) and
    (SheetPostfixPos > DefSheetPrefixPos) then
    Result := Copy(RangeString, SheetPrefixPos + 1, SheetPostfixPos -
      SheetPrefixPos - 1); 
end;

procedure TXlsxMapParserItem.DeleteCharFromBegin(const AChar: qiChar; var AString: qiString);
var
  PosChar: Integer;
begin
  PosChar := Pos(AChar, AString);
  while PosChar = 1 do
  begin
    Delete(AString, 1, 1);
    PosChar := Pos(AChar, AString);
  end;  
end;  

procedure TXlsxMapParserItem.SetSkipRange(const Range: TXlsxRangeItem;
  const SkipRows, SkipCols: Integer);
    begin
      case Range.Direction of
        xdColDown:
        begin
          Range.FirstRow := Range.FirstRow + SkipRows;
          if Range.FirstRow > Range.LastRow then
            PackRanges;
        end;
        xdColUp:
        begin
          Range.FirstRow := Range.FirstRow - SkipRows;
          if Range.FirstRow < Range.LastRow then
            PackRanges;
        end;
        xdRowRight:
        begin
          Range.FirstCol := Range.FirstCol + SkipCols;
          if Range.FirstCol > Range.LastCol then
            PackRanges;
        end;
        xdRowLeft:
        begin
          Range.FirstCol := Range.FirstCol - SkipCols;
          if Range.FirstCol < Range.LastCol then
            PackRanges;
        end;
      end;
end;

procedure TXlsxMapParserItem.RepairOldRange(var RangeString: qiString);
var
  LeftSideRange,
  RightSideRange,
  ColData: qiString;
  RowInt: Integer;
begin
  LeftSideRange := GetLeftSide(RangeString);
  RightSideRange := GetRightSide(RangeString);
  ColData := GetColData(LeftSideRange);
  RowInt := GetFirstRow(RangeString);
  if (ColData <> '') and (RowInt = 0) and (RightSideRange = '') then
    RangeString := RangeString + IntToStr(BeginRowInt) + LeftRightRangesDelim +
      ColFinishFlag;
end;

procedure TXlsxMapParserItem.FillRange(const Range: TXlsxRangeItem;
  const RangeString: qiString);
var
  RangeStringTmp: qiString;
begin
  RangeStringTmp := RangeString;
  RepairOldRange(RangeStringTmp);

  Range.SheetName := GetSheetName(RangeStringTmp);
  Range.SheetIndex := GetSheetIndex(RangeStringTmp);
  Range.FirstCol := GetFirstCol(RangeStringTmp);
  Range.FirstRow := GetFirstRow(RangeStringTmp);
  Range.LastCol := GetLastCol(RangeStringTmp);
  Range.LastRow := GetLastRow(RangeStringTmp);

  if Range.FirstCol = 0 then
    Range.FirstCol := Range.LastCol;
  if Range.LastCol = 0 then
    Range.LastCol := Range.FirstCol;
  if Range.FirstRow = 0 then
    Range.FirstRow := Range.LastRow;
  if Range.LastRow = 0 then
    Range.LastRow := Range.FirstRow;

  if Range.AsString = '' then
    raise Exception.CreateFmt(QImportLoadStr(QIE_XlsxMapError), [RangeStringTmp]);
end;

procedure TXlsxMapParserItem.Parse(const MapString: qiString; const SkipRows,
  SkipCols: Integer);
var
  MapStringTmp,
  RangeString: qiString;
  Range: TXlsxRangeItem;
  NeedSkip: Boolean;
  LenMapString: Integer;
begin
  MapStringTmp := MapString;
  LenMapString := Length(MapStringTmp);
  if LenMapString > 0 then
    if MapStringTmp[LenMapString] <> RangeDelim then
      MapStringTmp := MapStringTmp + RangeDelim;
      
  DeleteFieldDescribtion(MapStringTmp);

  NeedSkip := (SkipRows > 0) or (SkipCols > 0);
  DeleteCharFromBegin(RangeDelim, MapStringTmp);
  while GetRangeString(MapStringTmp, RangeString) do
  begin
    TruncMapString(MapStringTmp);

    Range := FRanges.Add;
    FillRange(Range, RangeString);

    if NeedSkip then
    begin
      SetSkipRange(Range, SkipRows, SkipCols);
      NeedSkip := False;
    end;
  end;
end;

procedure TXlsxMapParserItem.TruncMapString(var MapString: qiString);
var
  RangeDelimPos: Integer;
begin
  RangeDelimPos := Pos(RangeDelim, MapString);

  if RangeDelimPos > 0 then
  begin
    Delete(MapString, 1, RangeDelimPos);
    
    DeleteCharFromBegin(RangeDelim, MapString);
  end;
end;

{ TXlsxRangeItems }

function TXlsxRangeItems.Add: TXlsxRangeItem;
begin
  Result := TXlsxRangeItem.Create;
  FList.Add(Result);
end;

function TXlsxRangeItems.CellInRanges(const CellPoint: TXlsxCellPoint;
  const SheetIndex: Integer; const SheetName: qiString): Boolean;
var
  I: Integer;
begin
  Result := False;
  for I := 0 to Count - 1 do
  begin
    Result := Items[I].CellInRange(CellPoint, SheetIndex, SheetName);
    if Result then
      Break;
  end;
end;

procedure TXlsxRangeItems.Reset;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Items[I].Reset;
end;

procedure TXlsxRangeItems.Clear;
var
  I: Integer;
begin
  for I := Count - 1 downto 0 do
    Self.Delete(I);
end;

constructor TXlsxRangeItems.Create;
begin
  FList := TList.Create;
end;

procedure TXlsxRangeItems.Delete(const Index: Integer);
begin
  TXlsxRangeItem(FList[Index]).Free;
  FList.Delete(Index);
end;

destructor TXlsxRangeItems.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TXlsxRangeItems.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TXlsxRangeItems.GetItems(Index: Integer): TXlsxRangeItem;
begin
  Result := TXlsxRangeItem(FList[Index]);
end;

function TXlsxRangeItems.GetAsString: qiString;
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    Result := Result + RangeDelim + Items[I].AsString;
end;

{ TXlsxMapParserItems }

function TXlsxMapParserItems.Add(const MapString: qiString;
  const SkipRows, SkipCols: Integer): TXlsxMapParserItem;
begin
  Result := TXlsxMapParserItem.Create(MapString, SkipRows, SkipCols);
  FList.Add(Result);
end;

procedure TXlsxMapParserItems.Clear;
var
  I: Integer;
begin
  for I := Count - 1 downto 0 do
    Self.Delete(I);
end;

constructor TXlsxMapParserItems.Create;
begin
  FList := TList.Create;
end;

procedure TXlsxMapParserItems.Delete(const Index: Integer);
begin
  TXlsxMapParserItem(FList[Index]).Free;
  FList.Delete(Index);
end;

destructor TXlsxMapParserItems.Destroy;
begin
  Clear;
  FList.Free;
  inherited;
end;

function TXlsxMapParserItems.GetCount: Integer;
begin
  Result := FList.Count;
end;

function TXlsxMapParserItems.GetItems(Index: Integer): TXlsxMapParserItem;
begin
  Result := TXlsxMapParserItem(FList[Index]);
end;

{ TXlsxRangeItem }

function TXlsxRangeItem.CellInRange(const CellPoint: TXlsxCellPoint;
  const SheetIndex: Integer; const SheetName: qiString): Boolean;
begin
  Result := SheetInRange(SheetIndex, SheetName);
  case Direction of
    xdCell:
      Result := Result and (CellPoint.Col = FirstCol) and (CellPoint.Row = FirstRow);
    xdColDown:
      Result := Result and (CellPoint.Col = FirstCol) and
        (CellPoint.Row >= FirstRow) and (CellPoint.Row <= LastRow);
    xdColUp:
      Result := Result and (CellPoint.Col = FirstCol) and
        (CellPoint.Row <= FirstRow) and (CellPoint.Row >= LastRow);
    xdRowRight:
      Result := Result and (CellPoint.Row = FirstRow) and
        (CellPoint.Col >= FirstCol) and (CellPoint.Col <= LastCol);
    xdRowLeft:
      Result := Result and (CellPoint.Row = FirstRow) and
        (CellPoint.Col <= FirstCol) and (CellPoint.Col >= LastCol);
  end;
end;

function TXlsxRangeItem.GetCellRange(const Col, Row: Integer): qiString;
begin
  Result := GetColNameFromIndex(Col) + IntToStr(Row);
end;

function TXlsxRangeItem.GetDirection: TXlsxRangeDirection;
begin
  if (FirstCol = LastCol) and (FirstRow = LastRow) then
    Result := xdCell
  else
    if FirstCol = LastCol then
      if FirstRow < LastRow then
        Result := xdColDown
      else
        Result := xdColUp
    else
      if FirstCol < LastCol then
        Result := xdRowRight
      else
        Result := xdRowLeft;
end;

procedure TXlsxRangeItem.Reset;
begin
  FSheetName := '';
  FSheetIndex := 0;
  FFirstCol := 0;
  FLastCol := 0;
  FFirstRow := 0;
  FLastRow := 0;
end;

function TXlsxRangeItem.SheetInRange(const ASheetIndex: Integer;
  const ASheetName: qiString): Boolean;
begin
  Result := (CompareText(SheetName, ASheetName) = 0) or
    (SheetIndex = ASheetIndex) or ((CompareText(SheetName, '') = 0) and
    (SheetIndex < 1));
end;

function TXlsxRangeItem.GetAsString: qiString;
var
  BeginRange,
  EndRange,
  FullRange,
  SheetIdent: qiString;
begin
  Result := '';

  if (FirstCol = 0) or (FirstRow = 0) then
    Exit;

  if SheetName <> '' then
    SheetIdent := SheetPrefix + SheetName + SheetPostfix
  else
    if SheetIndex > 0 then
      SheetIdent := SheetPrefix + SheetIndexPrefix + IntToStr(SheetIndex) + SheetPostfix;

  BeginRange := GetCellRange(FirstCol, FirstRow);

  if (LastCol > 0) and (LastRow > 0)  then
    EndRange := GetCellRange(LastCol, LastRow)
  else
  begin
    LastCol := FirstCol;
    LastRow := FirstRow;
  end;  

  case Direction of
    xdColDown,
    xdColUp:
      if LastRow = EndRowInt then
        FullRange := BeginRange + LeftRightRangesDelim + ColFinishFlag
      else
        if FirstRow = BeginRowInt then
          FullRange := ColStartFlag + LeftRightRangesDelim + EndRange
        else
          FullRange := BeginRange + LeftRightRangesDelim + EndRange;
    xdRowRight,
    xdRowLeft:
      if LastCol = EndColInt then
        FullRange := BeginRange + LeftRightRangesDelim + RowFinishFlag
      else
        if FirstCol = BeginColInt then
          FullRange := RowStartFlag + LeftRightRangesDelim + EndRange
        else
          FullRange := BeginRange + LeftRightRangesDelim + EndRange;
    xdCell:
      FullRange := BeginRange;
  end;

  Result := SheetIdent + FullRange;
end;

{$ENDIF}
{$ENDIF}

end.
