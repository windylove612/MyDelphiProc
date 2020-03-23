unit QImport3Xlsx;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF XLSX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Classes,
    System.SysUtils,
    System.IniFiles,
    Winapi.msxml,
    Xml.XMLIntf,
  {$ELSE}
    Classes,
    SysUtils,
    IniFiles,
    XMLIntf,
    msxml,
  {$ENDIF}
  QImport3StrTypes,
  QImport3,
  QImport3Common,
  QImport3BaseDocumentFile,
  QImport3XlsxMapParser,
  QImport3XlsxDom;
type
  TXlsxStyle = class(TCollectionItem)
  private
    FNumFmtId: Integer;
    FFormatString: String;
  public
    property NumFmtId: Integer read FNumFmtId write FNumFmtId;
    property FormatString: String read FFormatString write FFormatString;
  end;

  TXlsxStyleList = class(TCollection)
  private
    function GetItem(Index: Integer): TXlsxStyle;
    procedure SetItem(Index: Integer; const Value: TXlsxStyle);
  public
    property Items[Index: Integer]: TXlsxStyle read GetItem write SetItem; default;
    function Add: TXlsxStyle;
    procedure SetFormatStringByNumFmtId(const NumFmtId: Integer;
      const FormatString: String);
  end;

  TXlsxSharedStrings = class(TCollectionItem)
  private
    FText: qiString;
  public
    property Text: qiString read FText write FText;
  end;

  TXlsxSharedStringList = class(TCollection)
  private
    function GetItem(Index: Integer): TXlsxSharedStrings;
    procedure SetItem(Index: Integer; const Value: TXlsxSharedStrings);
  public
    function Add: TXlsxSharedStrings;
    property Items[Index: Integer]: TXlsxSharedStrings read GetItem write SetItem; default;
  end;

  TXlsxMerge = class(TCollectionItem)
  private
//    FRange: string;
    FRange: qiString;
    FValue: qiString;
    FBeginRow: Integer;
    FBeginCol: Integer;
    FEndRow: Integer;
    FEndCol: Integer;
//    FFirstCellName: string;
    FFirstCellName: qiString;
//    procedure SetRange(const Value: string);
    procedure SetRange(const Value: qiString);
  public
//    property Range: string read FRange write SetRange;
    property Range: qiString read FRange write SetRange;
    property Value: qiString read FValue write FValue;
//    property FirstCellName: AnsiString read FFirstCellName;
    property FirstCellName: qiString read FFirstCellName;
    property BeginRow: Integer read FBeginRow;
    property BeginCol: Integer read FBeginCol;
    property EndRow: Integer read FEndRow;
    property EndCol: Integer read FEndCol;
  end;

  TXlsxMergeList = class(TCollection)
  private
    function GetItem(Index: Integer): TXlsxMerge;
    procedure SetItem(Index: Integer; const Value: TXlsxMerge);
  public
    function Add: TXlsxMerge;
    property Items[Index: Integer]: TXlsxMerge read GetItem write SetItem; default;
  end;

  TXlsxCell = class(TCollectionItem)
  private
    FName: string;
    FRow: Integer;
    FCol: Integer;
    FValue: qiString;
    FFormula: string;
    FIsFormulaExist: Boolean;
    FIsMerge: Boolean;
    procedure SetFormula(const Value: string);
    function GetName: string;
    procedure SetName(const Value: string);
  public
    constructor Create(Collection: TCollection); override;

    property Name: string read GetName write SetName;
    property Value: qiString read FValue write FValue;
    property Row: Integer read FRow write FRow;
    property Col: Integer read FCol write FCol;
    property IsMerge: Boolean read FIsMerge write FIsMerge;
    property Formula: string read FFormula write SetFormula;
    property IsFormulaExist: Boolean read FIsFormulaExist;
  end;

  TXlsxCellList = class(TCollection)
  private
    function GetItem(Index: Integer): TXlsxCell;
    procedure SetItem(Index: Integer; const Value: TXlsxCell);
  public
    function Add: TXlsxCell;
    property Items[Index: Integer]: TXlsxCell read GetItem write SetItem; default;
  end;

  TXlsxWorkSheet = class(TCollectionItem)
  private
    FName: string;
    FSheetID: integer;
    FColCount: Integer;
    FRowCount: Integer;
    FCells: TXlsxCellList;
    FDataCells: TqiStringGrid;
    FMergeCells: TXlsxMergeList;
    FIsHidden: boolean;
    procedure SetSheetID(const Value: integer);
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;

    procedure FillMerge(Cell: TXlsxCell);
    procedure LoadDataCells;
    property Name: string read FName write FName;
    property SheetID: integer read FSheetID write SetSheetID;
    property ColCount: Integer read FColCount write FColCount;
    property RowCount: Integer read FRowCount write FRowCount;
    property Cells: TXlsxCellList read FCells;
    property DataCells: TqiStringGrid read FDataCells;
    property MergeCells: TXlsxMergeList read FMergeCells;
    property IsHidden: boolean read FIsHidden write FIsHidden;
  end;

  TXlsxWorkSheetList = class(TCollection)
  private
    function GetItems(Index: integer): TXlsxWorkSheet;
    procedure SetItems(Index: integer; Value: TXlsxWorkSheet);
  public
    function Add: TXlsxWorkSheet;
    property Items[Index: integer]: TXlsxWorkSheet read GetItems
      write SetItems; default;

    function GetFirstSheet: TXlsxWorkSheet;
    function GetSheetByName(Name: qiString): TXlsxWorkSheet;
    function GetSheetByID(id: integer): TXlsxWorkSheet;
  end;

  TXlsxWorkbook = class
  private
    FWorkDir: string;
    FXMLDoc: IXMLDOMDocument;
    FWorkSheets: TXlsxWorkSheetList;
    FSharedStrings: IXMLDOMDocument;
    FStyles: TXlsxStyleList;
    FLoadHiddenSheets: Boolean;
    FNeedFillMerge: Boolean;
    function DoCorrectChars(const Value: qiString): qiString;
    procedure LoadSharedStrings;
    procedure LoadStyles;
    procedure SetWorkSheets;
    procedure LoadWorkSheets;
    procedure SetDataCells;
    procedure LoadSheet(SheetFile: qiString; Id: integer);
    function FormatValueByStyle(const Style: TXlsxStyle;
      const Value: Variant): Variant;
    function GetCellValue(const XlsxCell: IXMLCType): Variant;
    function GetSharedString(Index: Integer): qiString;
  private
    property Styles: TXlsxStyleList read FStyles;
  public
    constructor Create;
    destructor Destroy; override;

    procedure Load;
    property SharedStrings[index: Integer]: qiString read GetSharedString;
    property CurrFolder: string read FWorkDir write FWorkDir;
    property WorkSheets: TXlsxWorkSheetList read FWorkSheets;
    property LoadHiddenSheets: Boolean read FLoadHiddenSheets write FLoadHiddenSheets;
    property NeedFillMerge: Boolean read FNeedFillMerge write FNeedFillMerge;
  end;

  TXlsxFile = class(TBaseDocumentFile)
  private
    FWorkbook: TXlsxWorkbook;
  protected
    procedure LoadXML(CurrFolder: qiString); override;
  public
    constructor Create; override;
    destructor Destroy; override;
    property Workbook: TXlsxWorkbook read FWorkbook;
  end;

  TQImport3Xlsx = class(TQImport3)
  private
    FXlsxFile: TXlsxFile;
    FCounter: Integer;
    FSheetName: string;
    FLoadHiddenSheet: Boolean;
    FNeedFillMerge: Boolean;
    FMapParserRows: TXlsxMapParserItems;
    FEndImportFlag: Boolean;
    procedure SetSheetName(const Value: string);
    procedure SetLoadHiddenSheet(const Value: Boolean);
    procedure SetNeedFillMerge(const Value: Boolean);
    procedure ParseMap;
  protected
    procedure BeforeImport; override;
    procedure StartImport; override;
    function CheckCondition: Boolean; override;
    function Skip: Boolean; override;
    procedure ChangeCondition; override;
    procedure FinishImport; override;
    procedure AfterImport; override;
    procedure FillImportRow; override;
    function ImportData: TQImportResult; override;
    procedure DoLoadConfiguration(IniFile: TIniFile); override;
    procedure DoSaveConfiguration(IniFile: TIniFile); override;
    procedure DoUserDefinedImport; override;
    function AllowImportRowComplete: Boolean; override;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  published
    property FileName;
    property SkipFirstRows default 0;
    property SkipFirstCols default 0;
    property SheetName: string read FSheetName write SetSheetName;
    property LoadHiddenSheet: boolean read FLoadHiddenSheet
      write SetLoadHiddenSheet default False;
    property NeedFillMerge: Boolean read FNeedFillMerge
      write SetNeedFillMerge default False;
  end;
{$ENDIF}
{$ENDIF}

implementation

{$IFDEF XLSX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.StrUtils,
    System.Variants;
  {$ELSE}
    StrUtils,
    Variants;
  {$ENDIF}

function XlsxFloatFormat(const StrValue: String; var ResultValue: Variant): Boolean;
var
  TempStrValue: String;
  TempResultValue: Extended;
begin
  TempStrValue := StrValue;
  {$IFDEF VCL6}
    TempStrValue := AnsiReplaceStr(TempStrValue, '.', {$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator);
  {$ELSE}
    TempStrValue := StringReplace(TempStrValue, '.', DecimalSeparator);
  {$ENDIF}

  {$IFDEF VCL7}
    Result := TryStrToFloat(TempStrValue, TempResultValue);
    if Result then
      ResultValue := TempResultValue;
  {$ELSE}
    try
      TempResultValue := StrToFloat(TempStrValue);
      ResultValue := TempResultValue;
      Result := True;
    except
      Result := False;
    end;
  {$ENDIF}
end;

function XlsxDateTimeFormat(const DateTime: string): TDateTime;
var
  TempStr: string;
const
  cXMLDelimiter = '.';
begin
  TempStr := DateTime;
  if {$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator <> cXMLDelimiter then
    if Pos(cXMLDelimiter, TempStr) > 0 then
      {$IFDEF VCL6}
      TempStr := AnsiReplaceStr(DateTime, cXMLDelimiter, {$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator);
      {$ELSE}
      TempStr := StringReplace(DateTime, cXMLDelimiter, DecimalSeparator, [rfReplaceAll, rfIgnoreCase]);
      {$ENDIF}
  Result := StrToFloat(TempStr);
end;

{$IFNDEF VCL6}
function BoolToStr(B: Boolean; UseBoolStrs: Boolean = False): AnsiString;
const
  cSimpleBoolStrs: array [boolean] of AnsiString = ('0', '-1');
begin
  if UseBoolStrs then
  begin
    if B then
      Result := 'True'
    else
      Result := 'False';
  end
  else
    Result := cSimpleBoolStrs[B];
end;
{$ENDIF}

function FormatData(s: string; i: Integer): Double;
begin
  Result := Round(StrToFloat(s)*exp(i*ln(10)))/(exp(i*ln(10)));
end;

function GetXlsxPrecentValue(const AValue: string): string;
var
  snum: string;
  dnum: Double;
  i, digits: Integer;
begin
  Result := AValue;
  {e.g. format 3.2142857142857099E-3}
  if Pos('E-', AValue) > 0 then
  begin
    snum := Copy(AValue, 1, Pos('E-', AValue) - 1);
    {$IFDEF VCL6}
    dnum := FormatData(AnsiReplaceStr(snum, '.', {$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator), 2);
    {$ELSE}
    dnum := FormatData(StringReplace(snum, '.', DecimalSeparator), 2);
    {$ENDIF}
    digits := StrToIntDef(Copy(AValue, Pos('E-', AValue) + 2, Length(AValue)), 0) - 2;
    for i := 0 to digits - 1 do
      dnum := dnum / 10;
    Result := FormatFloat('0.00%', dnum);
  end else
  {e.g. format 0.55200000000000005}
  begin
    {$IFDEF VCL6}
    dnum := FormatData(AnsiReplaceStr(AValue, '.', {$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator), 4) * 100;
    {$ELSE}
    dnum := FormatData(StringReplace(AValue, '.', {$IFDEF VCL17}FormatSettings.{$ENDIF}DecimalSeparator), 4) * 100;
    {$ENDIF}
    Result := FormatFloat('0.00%', dnum);
  end;
end;


{ TXlsxStyleList }

function TXlsxStyleList.GetItem(Index: Integer): TXlsxStyle;
begin
  Result := TXlsxStyle(inherited Items[Index]);
end;

procedure TXlsxStyleList.SetItem(Index: Integer; const Value: TXlsxStyle);
begin
  inherited Items[Index] := Value;
end;

function TXlsxStyleList.Add: TXlsxStyle;
begin
  Result := TXlsxStyle(inherited Add);
end;

procedure TXlsxStyleList.SetFormatStringByNumFmtId(const NumFmtId: Integer;
  const FormatString: String);
var
  I: Integer;
begin
  for I := 0 to Count - 1 do
    if Items[I].NumFmtId = NumFmtId then
      Items[I].FormatString := FormatString;
end;

{ TXlsxSharedStringList }

function TXlsxSharedStringList.GetItem(Index: Integer): TXlsxSharedStrings;
begin
  Result := TXlsxSharedStrings(inherited Items[Index]);
end;

procedure TXlsxSharedStringList.SetItem(Index: Integer; const Value: TXlsxSharedStrings);
begin
  inherited Items[Index] := Value;
end;

function TXlsxSharedStringList.Add: TXlsxSharedStrings;
begin
  Result := TXlsxSharedStrings(inherited Add);
end;

{ TXlsxMerge }

//procedure TXlsxMerge.SetRange(const Value: string);
procedure TXlsxMerge.SetRange(const Value: qiString);
begin
  if FRange <> Value then
  begin
    FRange := Value;
    FFirstCellName := Copy(Value, 1, Pos(':', Value) - 1);
    FBeginCol := GetColIdFromString(FFirstCellName);
    FBeginRow := GetRowIdFromString(FFirstCellName);
    FEndCol := GetColIdFromString(Copy(Value, Pos(':', Value)+ 1, Length(Value)));
    FEndRow := GetRowIdFromString(Copy(Value, Pos(':', Value)+ 1, Length(Value)));
  end;
end;

{ TXlsxMergeList }

function TXlsxMergeList.GetItem(Index: Integer): TXlsxMerge;
begin
  Result := TXlsxMerge(inherited Items[Index]);
end;

procedure TXlsxMergeList.SetItem(Index: Integer;
  const Value: TXlsxMerge);
begin
  inherited Items[Index] := Value;
end;

function TXlsxMergeList.Add: TXlsxMerge;
begin
  Result := TXlsxMerge(inherited Add);
end;

{ TXlsxCell }

constructor TXlsxCell.Create(Collection: TCollection);
begin
  inherited;
  FCol := 0;
  FRow := 0;
  FName := '';
  FValue := '';
  FIsMerge := False;
  FFormula := '';
  FIsFormulaExist := False;
end;

procedure TXlsxCell.SetFormula(const Value: string);
begin
  if FFormula <> Value then
  begin
    FFormula := Value;
    FIsFormulaExist := True;
  end;
end;

function TXlsxCell.GetName: string;
begin
  if not Self.IsMerge then
    Result := FName
  else
    Result := 'Merge';
end;

procedure TXlsxCell.SetName(const Value: string);
begin
  if not Self.IsMerge then
    FName := Value;
end;

{ TXlsxCellList }

function TXlsxCellList.Add: TXlsxCell;
begin
  Result := TXlsxCell(inherited Add);
end;

function TXlsxCellList.GetItem(Index: Integer): TXlsxCell;
begin
  Result := TXlsxCell(inherited Items[Index]);
end;

procedure TXlsxCellList.SetItem(Index: Integer; const Value: TXlsxCell);
begin
  inherited Items[Index] := Value;
end;

{ TXlsxWorkSheet }

procedure TXlsxWorkSheet.SetSheetID(const Value: integer);
const
  sSheetIDCheck = 'Sheet ID must be >= 0!';
begin
  if Value < 0 then
    raise Exception.Create(sSheetIDCheck);

  if FSheetID <> Value then
    FSheetID := Value;
end;

constructor TXlsxWorkSheet.Create(Collection: TCollection);
begin
  inherited;
  FColCount := 0;
  FRowCount := 0;
  FIsHidden := False;
  FCells := TXlsxCellList.Create(TXlsxCell);
  FMergeCells := TXlsxMergeList.Create(TXlsxMerge);
  FDataCells := TqiStringGrid.Create(nil);
end;

destructor TXlsxWorkSheet.Destroy;
begin
  FCells.Free;
  FMergeCells.Free;
  FDataCells.Free;
  inherited;
end;

procedure TXlsxWorkSheet.FillMerge(Cell: TXlsxCell);
var
  i: Integer;
begin
  for i := 0 to FMergeCells.Count - 1 do
    if (Cell.Row in [FMergeCells[i].BeginRow..FMergeCells[i].EndRow]) and
      (Cell.Col in [FMergeCells[i].BeginCol..FMergeCells[i].EndCol]) and
      (not ((Cell.Row = FMergeCells[i].BeginRow) and (Cell.Col = FMergeCells[i].BeginCol)))
    then Cell.Value := FMergeCells[i].Value;
end;

procedure TXlsxWorkSheet.LoadDataCells;
var
  i: Integer;
begin
  FDataCells.ColCount := FColCount;
  FDataCells.RowCount := FRowCount;
  for i := FCells.Count - 1 downto 0 do
  begin
    FDataCells.Cells[FCells[i].Col - 1, Cells[i].Row - 1] := Cells[i].Value;
    //Cells[i].Free;
  end;
end;

{ TXlsxWorkSheetList }

function TXlsxWorkSheetList.GetItems(Index: integer): TXlsxWorkSheet;
begin
  Result := TXlsxWorkSheet(inherited Items[Index]);
end;

procedure TXlsxWorkSheetList.SetItems(Index: integer;
  Value: TXlsxWorkSheet);
begin
  inherited Items[Index] := Value;
end;

function TXlsxWorkSheetList.Add: TXlsxWorkSheet;
begin
  Result := TXlsxWorkSheet(inherited Add);
end;

function TXlsxWorkSheetList.GetFirstSheet: TXlsxWorkSheet;
const
  sSheetCountMustBeMore0 = 'Sheets Count must be > 0!';
begin
  if Self.Count > 0 then
    Result := Items[0]
  else
    raise Exception.Create(sSheetCountMustBeMore0);
end;

function TXlsxWorkSheetList.GetSheetByName(
  Name: qiString): TXlsxWorkSheet;
var
  i: Integer;
begin
  Result := nil;
  if Name = qiString('') then
    Result := GetFirstSheet
  else
    for i := 0 to Self.Count - 1 do
      if UpperCase(Items[i].Name) = UpperCase(Name) then
      begin
        Result := Items[i];
        Break;
      end;
end;

function TXlsxWorkSheetList.GetSheetByID(id: integer): TXlsxWorkSheet;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Self.Count - 1 do
    if Items[i].SheetID = id then
    begin
      Result := Items[i];
      Break;
    end;
end;

{ TXlsxWorkbook }
function TXlsxWorkbook.DoCorrectChars(const Value: qiString): qiString;
const
  OldPatternMask: qiString = '_x%s_';
var
  I: Byte;
  OldPattern: qiString;
begin
  Result := Value;
  for I := 1 to Ord(#32) - 1 do
  begin
    OldPattern := Format(OldPatternMask, [IntToHex(I, 4)]);
    Result := StringReplace(Result, OldPattern, Chr(I), [rfReplaceAll]);
  end;
end;

procedure TXlsxWorkbook.LoadSharedStrings;
var
  ARec: TSearchRec;
begin
  if (FindFirst(FWorkDir + 'xl\sharedStrings.xml', faAnyFile, ARec) = 0) then
//  if FileExists(FWorkDir + 'xl\sharedStrings.xml') then
    FSharedStrings.load(FWorkDir + 'xl\sharedStrings.xml');
end;

procedure TXlsxWorkbook.LoadStyles;
var
  StyleNode: IXMLDOMNode;
  XmlFileRec: TSearchRec;
  i: Integer;
  Node: IXMLDOMNode;
  NodeAttributes: IXMLDOMNamedNodeMap;
  NumFmtId: Integer;
begin
  try
    if FindFirst(FWorkDir + 'xl\styles.xml', faAnyFile, XmlFileRec) = 0 then
    begin
      FXMLDoc.load(FWorkDir + 'xl\styles.xml');
      StyleNode := FXMLdoc.selectSingleNode('/styleSheet/cellXfs');
      for i := 0 to StyleNode.childNodes.length - 1 do
      begin
        FStyles.Add;
        NodeAttributes := StyleNode.childNodes[i].attributes;
        if Assigned(NodeAttributes) then
          begin
            Node := NodeAttributes.getNamedItem('numFmtId');
            if Assigned(Node) then
              FStyles[i].NumFmtId := Node.nodeValue;
          end;
      end;

      StyleNode := FXMLdoc.selectSingleNode('/styleSheet/numFmts');

      if Assigned(StyleNode) then
        if Assigned(StyleNode.childNodes) then
        for i := 0 to StyleNode.childNodes.length - 1 do
        begin
          NodeAttributes := StyleNode.childNodes[i].attributes;
          if Assigned(NodeAttributes) then
          begin
            Node := NodeAttributes.getNamedItem('numFmtId');
            if Assigned(Node) then
            begin
              NumFmtId := Node.nodeValue;
              Node := NodeAttributes.getNamedItem('formatCode');
              if Assigned(Node) then
                FStyles.SetFormatStringByNumFmtId(NumFmtId, Node.nodeValue);
            end;
          end;
        end;
    end;
  finally
    FindClose(XmlFileRec);
  end;
end;

function GetSheetID(const r_id: string): integer;
begin
  Result := StrToIntDef(Copy(r_id, 4, Length(r_id) - 3),{0} -1);
  //igorp ticket #30556 r_id может начинаться с 0, поэтому default=-1
end;

procedure TXlsxWorkbook.SetWorkSheets;
var
  SheetNodes: IXMLDOMNodeList;
  i: Integer;
  TempWorkSheet: TXlsxWorkSheet;
  StateNodeItem,
  NameNodeItem: IXMLDOMNode;
  ARec: TSearchRec;
begin
  if (FindFirst(FWorkDir + 'xl\workbook.xml', faAnyFile, ARec) = 0)  then
//  if FileExists(FWorkDir + 'xl\workbook.xml') then
  begin
    FXMLDoc.load(FWorkDir + 'xl\workbook.xml');
    SheetNodes := FXMLdoc.selectNodes('/workbook/sheets/sheet');
    for i := 0 to SheetNodes.length - 1 do
    begin
      NameNodeItem := SheetNodes[i].attributes.getNamedItem('name');
      if Assigned(NameNodeItem) then
      begin
        StateNodeItem := SheetNodes[i].attributes.getNamedItem('state');
        if not Assigned(StateNodeItem) or (StateNodeItem.nodeValue = 'visible') then
        begin
          TempWorkSheet := FWorkSheets.Add;
          TempWorkSheet.Name := NameNodeItem.nodeValue;
          TempWorkSheet.SheetID := GetSheetID(SheetNodes[i].attributes.getNamedItem('r:id').nodeValue);
        end
        else
          if Assigned(StateNodeItem) and (StateNodeItem.nodeValue = 'hidden') and
            FLoadHiddenSheets then
          begin
            TempWorkSheet := FWorkSheets.Add;
            TempWorkSheet.Name := NameNodeItem.nodeValue;
            TempWorkSheet.IsHidden := True;
          end;
      end;
    end;
  end;
end;

procedure TXlsxWorkbook.LoadWorkSheets;
var
  Nodes: IXMLDOMNodeList;

  function FindRel(ASheetID: integer): string;
  var
    j: Integer;
  begin
    Result := EmptyStr;
    for j := 0 to Nodes.length - 1 do
      if Assigned(Nodes[j].attributes.getNamedItem('Id'))
      and Assigned(Nodes[j].attributes.getNamedItem('Target')) then
      begin
        if ASheetID = GetSheetID(Nodes[j].attributes.getNamedItem('Id').nodeValue) then
        begin
          Result := StringReplace(
                    Nodes[j].attributes.getNamedItem('Target').nodeValue,
                    '/','\', [rfReplaceAll, rfIgnoreCase]);
          Break;
        end;
      end;
  end;

var
  SRec: TSearchRec;
  i: integer;
  filename: string;
begin
  try
    if FindFirst(FWorkDir + 'xl\_rels\workbook.xml.rels', faDirectory, SRec) = 0 then
      FXMLDoc.load(FWorkDir + 'xl\_rels\workbook.xml.rels');
    Nodes := FXMLdoc.selectNodes('/Relationships/Relationship');
    for i := 0 to FWorkSheets.Count - 1 do
      begin
        filename := FWorkDir + 'xl\' + FindRel(FWorkSheets[i].SheetID);
        //igorp ticket #30556 - в xml.rels могут подставляться пути уже с \xl\
        filename := StringReplace(filename, 'xl\\xl\','xl\',[]);
        //\igorp
        LoadSheet( filename, FWorkSheets[i].SheetID);
      end;
  finally
    FindClose(SRec);
  end;
end;

procedure TXlsxWorkbook.SetDataCells;
var
  i: Integer;
begin
  for i := 0 to FWorkSheets.Count - 1 do
    FWorkSheets[i].LoadDataCells;
end;

procedure TXlsxWorkbook.LoadSheet(SheetFile: qiString; Id: integer);
var
  CurrentSheet: TXlsxWorkSheet;
  Dimension, DimensionMerge: qiString;
  i, j: Integer;
  TempMerge: TXlsxMerge;
  TempCell: TXlsxCell;
  SheetXlsx: IXMLWorksheetType;

  NodeValueF,
  ValueOfNode: Variant;
  k: Integer;

begin
  CurrentSheet := FWorkSheets.GetSheetByID(id);
  if Assigned(CurrentSheet) then
  begin
    SheetXlsx := Loadworksheet(SheetFile);

    try
      Dimension := SheetXlsx.Dimension.Ref;

      for i := 0 to SheetXlsx.MergeCells.Count - 1 do
      begin
        DimensionMerge := SheetXlsx.MergeCells.MergeCell[i].Ref;
        TempMerge := CurrentSheet.MergeCells.Add;
        TempMerge.Range := DimensionMerge;
      end;

      for i := 0 to SheetXlsx.SheetData.Count - 1 do
      begin
        for j := 0 to SheetXlsx.SheetData.Row[i].Count - 1 do
        begin
          TempCell := CurrentSheet.Cells.Add;
          TempCell.Name := SheetXlsx.SheetData.Row[i].C[j].R;
          Dimension := SheetXlsx.SheetData.Row[i].C[j].R;

          TempCell.Col := GetColIdFromString(Dimension);
          if CurrentSheet.ColCount < TempCell.Col then
            CurrentSheet.ColCount := TempCell.Col;
          TempCell.Row := GetRowIdFromString(Dimension);
          if CurrentSheet.RowCount < TempCell.Row then
            CurrentSheet.RowCount := TempCell.Row;

          NodeValueF := SheetXlsx.SheetData.Row[i].C[j].F;
          if not VarIsNull(NodeValueF) and (NodeValueF <> '') then
            TempCell.Formula := NodeValueF;

          ValueOfNode := GetCellValue(SheetXlsx.SheetData.Row[i].C[j]);
          if VarIsNull(ValueOfNode) then
            TempCell.Value := ''
          else
            TempCell.Value := ValueOfNode;

          for k := 0 to CurrentSheet.MergeCells.Count - 1 do
            if CurrentSheet.MergeCells[k].FirstCellName = TempCell.Name then
              CurrentSheet.MergeCells[k].Value := TempCell.Value;

          if NeedFillMerge then
            CurrentSheet.FillMerge(TempCell);
        end;
      end;
    finally
      SheetXlsx := nil;
    end;
  end;
end;

function TXlsxWorkbook.FormatValueByStyle(const Style: TXlsxStyle;
  const Value: Variant): Variant;
type
  TXlsxFormatType = (xftUnknown, xftDate, xftTime, xftDateTime, xftPercent);
function GetFormatByMask(const Mask: String): TXlsxFormatType;
const
  DateFlags: array[0..3] of String[2] = ('/m', 'm/', '/M', 'M/');
  TimeFlags: array[0..5] of String[2] = (':m', 'm:', ':m', 'm:', ':n', 'n:');
var
  IsTime,
  IsDate: Boolean;
  I: Integer;
begin
  Result := xftUnknown;

  for I := 0 to Length(DateFlags) - 1 do
  begin
    IsDate := Pos(String(DateFlags[I]), Style.FormatString) > 0;
    if IsDate then
      Break;
  end;
  for I := 0 to Length(TimeFlags) - 1 do
  begin
    IsTime := Pos(String(TimeFlags[I]), Style.FormatString) > 0;
    if IsTime then
      Break;
  end;

  if IsDate and IsTime then
    Result := xftDateTime
  else
    if IsDate then
      Result := xftDate
    else
      if IsTime then
        Result := xftTime;
end;
var
  XlsxFormatType: TXlsxFormatType;
begin
  Result := Null;
  if VarIsNull(Value) then
    Exit;

  XlsxFormatType := xftUnknown;
  case Style.NumFmtId of
    14..17, 164..187:
      if Style.FormatString = '' then
        XlsxFormatType := xftDate
      else
        XlsxFormatType := GetFormatByMask(Style.FormatString);
    18..21, 45..47:
      XlsxFormatType := xftTime;
    22:
      XlsxFormatType := xftDateTime;
    9, 10:
      XlsxFormatType := xftPercent;
  end;

  case XlsxFormatType of
    xftDate:
      if XlsxFloatFormat(VarToStr(Value), Result) then
        Result := DateToStr(XlsxDateTimeFormat(
          VarToStr(Value)));
    xftTime:
      Result := TimeToStr(XlsxDateTimeFormat(
        VarToStr(Value)));
    xftDateTime:
      Result := DateTimeToStr(XlsxDateTimeFormat(
        VarToStr(Value)));
    xftPercent:
      Result := GetXlsxPrecentValue(VarToStr(Value));
  end;
end;

constructor TXlsxWorkbook.Create;
begin
  inherited;
  FWorkDir := '';
  FLoadHiddenSheets := False;
  FNeedFillMerge := False;
  FXMLDoc := CoDOMDocument.Create;
  FSharedStrings := CoDOMDocument.Create;
  FStyles := TXlsxStyleList.Create(TXlsxStyle);
  FWorkSheets := TXlsxWorkSheetList.Create(TXlsxWorkSheet);
end;

destructor TXlsxWorkbook.Destroy;
begin
  FXMLDoc := nil;
  FSharedStrings := nil;
  if Assigned(FWorkSheets) then
    FWorkSheets.Free;
  if Assigned(FStyles) then
    FStyles.Free;
  inherited;
end;

function TXlsxWorkbook.GetCellValue(const XlsxCell: IXMLCType): Variant;
var
  NodeValueT: Variant;
  NodeValueV: Variant;
  NodeValueS: Variant;
  StyleIndex: Integer;
  ValueIsT: WideString;
begin
  NodeValueT := XlsxCell.T;

  NodeValueS := XlsxCell.S;

  NodeValueV := XlsxCell.V;

  Result := null;
  if (NodeValueT = 's') and not VarIsNull(NodeValueV) then
    Result := SharedStrings[Round(NodeValueV)]
  else
    if (NodeValueT = 'str') and not VarIsNull(NodeValueV) then
      Result := NodeValueV
    else
      if NodeValueT = 'inlineStr' then
      begin
        if Assigned(XlsxCell.Is_) then
          ValueIsT := XlsxCell.Is_.T;
        if ValueIsT <> '' then
          Result := ValueIsT;
      end
      else
        if (NodeValueT = 'b') and not VarIsNull(NodeValueV) then
          Result :=
            BoolToStr(Boolean(
              StrToIntDef(NodeValueV, 0)), True)
        else
          if not VarIsNull(NodeValueS) then
            begin
              StyleIndex := NodeValueS;
              if (Styles.Count > StyleIndex) and (StyleIndex >= 0) then
                Result := FormatValueByStyle(Styles[StyleIndex],
                  NodeValueV);
            end;

  if VarIsNull(Result) then
    if not VarIsNull(NodeValueV) then
      if not XlsxFloatFormat(VarToStr(NodeValueV), Result) then
        Result := NodeValueV;
end;

function TXlsxWorkbook.GetSharedString(Index: Integer): qiString;
var
  SharedNode: IXMLDOMNode;
begin
  SharedNode := FSharedStrings.selectSingleNode('/sst');
  if Assigned(SharedNode) and SharedNode.hasChildNodes  and (SharedNode.childNodes.length > Index) then
    Result := DoCorrectChars(SharedNode.childNodes.item[Index].text)
  else
    Result := '';
end;

procedure TXlsxWorkbook.Load;
begin
  if FWorkDir <> '' then
  begin
    LoadSharedStrings;
    LoadStyles;
    SetWorkSheets;
    LoadWorkSheets;
    SetDataCells;
  end;
end;

{ TXlsxFile }

constructor TXlsxFile.Create;
begin
  inherited;
  FWorkbook := TXlsxWorkbook.Create;
end;

destructor TXlsxFile.Destroy;
begin
  if Assigned(FWorkbook) then
    FWorkbook.Free;
  inherited;
end;

procedure TXlsxFile.LoadXML(CurrFolder: qiString);
begin
  FWorkbook.CurrFolder := CurrFolder;
  FWorkbook.Load;
end;

{ TQImport3Xlsx }

procedure TQImport3Xlsx.SetSheetName(const Value: string);
begin
  if FSheetName <> Value then
    FSheetName := Value;
end;

procedure TQImport3Xlsx.SetLoadHiddenSheet(const Value: Boolean);
begin
  if FLoadHiddenSheet <> Value then
    FLoadHiddenSheet := Value;
end;

procedure TQImport3Xlsx.SetNeedFillMerge(const Value: Boolean);
begin
  if FNeedFillMerge <> Value then
    FNeedFillMerge := Value;
end;

procedure TQImport3Xlsx.ParseMap;
var
  I: Integer;
begin
  FMapParserRows.Clear;
  for I := 0 to Map.Count - 1 do
    FMapParserRows.Add(Map[I], SkipFirstRows, SkipFirstCols);
end;  

procedure TQImport3Xlsx.BeforeImport;
begin
  FEndImportFlag := False;
  ParseMap;
  FXlsxFile := TXlsxFile.Create;
  FXlsxFile.Workbook.FLoadHiddenSheets := FLoadHiddenSheet;
  FXlsxFile.Workbook.NeedFillMerge := FNeedFillMerge;
  FXlsxFile.FileName := FileName;
  FXlsxFile.Load;
  if Assigned(FXlsxFile.Workbook.WorkSheets.GetSheetByName(FSheetName))then
    FTotalRecCount := FXlsxFile.Workbook.WorkSheets.GetSheetByName(FSheetName).RowCount -
      SkipFirstRows;
  inherited;
//  Formats.RestoreFormats;
end;

procedure TQImport3Xlsx.StartImport;
begin
  inherited;
  FCounter := 0;
end;

function TQImport3Xlsx.CheckCondition: Boolean;
begin
  Result := not FEndImportFlag;
end;

function TQImport3Xlsx.Skip: Boolean;
begin
  Result := (SkipFirstRows > 0) and (FCounter < SkipFirstRows);
end;

procedure TQImport3Xlsx.ChangeCondition;
begin
  Inc(FCounter);
end;

procedure TQImport3Xlsx.FinishImport;
begin
  if not Canceled and not IsCSV then
  begin
    if CommitAfterDone then
      DoNeedCommit
    else if (CommitRecCount > 0) and ((ImportedRecs + ErrorRecs) mod CommitRecCount > 0) then
      DoNeedCommit;
  end;
end;

procedure TQImport3Xlsx.AfterImport;
begin
  FXlsxFile.Free;
  FXlsxFile := nil;
  inherited;
end;

procedure TQImport3Xlsx.FillImportRow;
var
  i, MapIndex: Integer;
  strValue: qiString;
  p: Pointer;
  SheetNameTmp: qiString;
  RowTmp,
  ColTmp,
  SheetIndexTmp: Integer;
  WorkSheetTmp: TXlsxWorkSheet;
begin
  FImportRow.ClearValues;
  FEndImportFlag := True;
  RowIsEmpty := True;
  for i := 0 to FImportRow.Count - 1 do
  begin
    strValue := EmptyStr;
    if FImportRow.MapNameIdxHash.Search(FImportRow[i].Name, p) then
    begin
      MapIndex := Integer(p);
      WorkSheetTmp := nil;
      if FMapParserRows[MapIndex].GetNextRange(RowTmp, ColTmp, SheetIndexTmp,
        SheetNameTmp) then
      begin
        if SheetNameTmp = '' then
          SheetNameTmp := FSheetName;

        if SheetIndexTmp > 0 then
        begin
          WorkSheetTmp := FXlsxFile.Workbook.WorkSheets.GetSheetByID(SheetIndexTmp);
          if WorkSheetTmp = nil then
            WorkSheetTmp := FXlsxFile.Workbook.WorkSheets.GetSheetByName(SheetNameTmp);
        end
        else
          WorkSheetTmp := FXlsxFile.Workbook.WorkSheets.GetSheetByName(SheetNameTmp);

        if Assigned(WorkSheetTmp) and (RowTmp <= WorkSheetTmp.DataCells.RowCount) and
          (ColTmp <= WorkSheetTmp.DataCells.ColCount) then
          strValue := WorkSheetTmp.DataCells.Cells[ColTmp - 1, RowTmp - 1];
      end;

      FEndImportFlag := FEndImportFlag and (strValue = '') and
        ((WorkSheetTmp = nil) or (RowTmp > WorkSheetTmp.DataCells.RowCount));
      if AutoTrimValue then
        strValue := Trim(strValue);
      RowIsEmpty := RowIsEmpty and (strValue = '');
      FImportRow.SetValue(Map.Names[MapIndex], strValue, False);
    end;
    DoUserDataFormat(FImportRow[i]);
  end;
end;

function TQImport3Xlsx.ImportData: TQImportResult;
begin
  Result := qirOk;

  if FEndImportFlag then
    Exit;
    
  try
    try
      if Canceled  and not CanContinue then
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

procedure TQImport3Xlsx.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    SkipFirstRows := ReadInteger(XLSX_OPTIONS, XLSX_SKIP_LINES, SkipFirstRows);
    SheetName := ReadString(XLSX_OPTIONS, XLSX_SHEET_NAME, SheetName);
    LoadHiddenSheet := ReadBool(XLSX_OPTIONS, XLSX_LOAD_HIDDEN_SHEET,
      LoadHiddenSheet);
    NeedFillMerge := ReadBool(XLSX_OPTIONS, XLSX_NEED_FILL_MERGE,
      NeedFillMerge);
  end;
end;

procedure TQImport3Xlsx.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    WriteInteger(XLSX_OPTIONS, XLSX_SKIP_LINES, SkipFirstRows);
    WriteString(XLSX_OPTIONS, XLSX_SHEET_NAME, SheetName);
    WriteBool(XLSX_OPTIONS, XLSX_LOAD_HIDDEN_SHEET, LoadHiddenSheet);
    WriteBool(XLSX_OPTIONS, XLSX_NEED_FILL_MERGE, NeedFillMerge);
  end;
end;

procedure TQImport3Xlsx.DoUserDefinedImport; 
begin
  if not FEndImportFlag then
    inherited;
end;  

function TQImport3Xlsx.AllowImportRowComplete: Boolean; 
begin
  Result := not FEndImportFlag;
end;  

constructor TQImport3Xlsx.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
  FLoadHiddenSheet := False;
  FMapParserRows := TXlsxMapParserItems.Create;
end;

destructor TQImport3Xlsx.Destroy;
begin
  FMapParserRows.Free;
  inherited;
end;  

{$ENDIF}
{$ENDIF}

end.
