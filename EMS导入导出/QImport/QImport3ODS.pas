unit QImport3ODS;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF ODS}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Classes,
    System.SysUtils,
    Winapi.MSXML,
    System.IniFiles,
  {$ELSE}
    Classes,
    SysUtils,
    MSXML,
    IniFiles,
  {$ENDIF}
  QImport3StrTypes,
  QImport3BaseDocumentFile,
  QImport3Common,
  QImport3;

type
  TODSCell = class(TCollectionItem)
  private
    FRow: Integer;
    FCol: Integer;
    FValue: qiString;
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    property Value: qiString read FValue write FValue;
    property Row: Integer read FRow write FRow;
    property Col: Integer read FCol write FCol;
  end;

  TODSCellList = class(TCollection)
  private
    function GetItem(Index: Integer): TODSCell;
    procedure SetItem(Index: Integer; const Value: TODSCell);
  public
    function Add: TODSCell;
    property Items[Index: Integer]: TODSCell read GetItem write SetItem; default;
  end;

  TODSSpreadSheet = class(TCollectionItem)
  private
    FName: AnsiString;
    FColCount: Integer;
    FRowCount: Integer;
    FCells: TODSCellList;
    FDataCells: TqiStringGrid;
    FX: Integer;
    FY: Integer;
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    procedure LoadDataCells;
    property Cells: TODSCellList read FCells;
    property ColCount: Integer read FColCount write FColCount;
    property RowCount: Integer read FRowCount write FRowCount;
    property X: Integer read FX write FX;
    property Y: Integer read FY write FY;
    property Name: AnsiString read FName write FName;
    property DataCells: TqiStringGrid read FDataCells;
  end;

  TODSSpreadSheetList = class(TCollection)
  private
    function GetItems(Index: integer): TODSSpreadSheet;
    procedure SetItems(Index: integer; Value: TODSSpreadSheet);
  public
    function Add: TODSSpreadSheet;
    function GetSheetByName(Name: AnsiString): TODSSpreadSheet;
    property Items[Index: integer]: TODSSpreadSheet read GetItems
      write SetItems; default;
  end;

  TODSWorkbook = class
  private
    FWorkDir: qiString;
    FileName: qiString;
    FSpreadSheets: TODSSpreadSheetList;
    FXMLDoc: IXMLDOMDocument;
    FIsNotExpanding: Boolean;

    procedure SetSpreadSheets;
    procedure FindTables(NameOfFile: qiString);
    procedure ParseTable(Table: TODSSpreadSheet; Nodes: IXMLDOMNodeList;
      NumbOfRepColumns, NumbOfRepRows: Integer; IsSpanning: Boolean);
    procedure ExpandRowsNCols(Table: TODSSpreadSheet;
      ExpandValue: qiString;
      NumbOfRepRows, NumbOfRepColumns: Integer; IsSpanning: Boolean);
    function GetNodeFullTextValue(const Node: IXMLDOMNode): qiString;
  public
    constructor Create;
    destructor Destroy; override;
    procedure Load;

    property WorkDir: qiString read FWorkDir write FWorkDir;
    property SpreadSheets: TODSSpreadSheetList read FSpreadSheets;
    property IsNotExpanding: Boolean read FIsNotExpanding write FIsNotExpanding;
  end;

  TODSFile = class(TBaseDocumentFile)
    private
      FWorkbook: TODSWorkbook;
    protected
      procedure LoadXML(WorkDir: qiString); override;
    public
      constructor Create; override;
      destructor Destroy; override;

      property Workbook: TODSWorkbook read FWorkbook;
    end;

  TQImport3ODS = class(TQImport3)
  private
    FODSFile: TODSFile;
    FCounter: Integer;
    FSheetName: AnsiString;
    FNotExpandMergedValue: Boolean;
    procedure SetSheetName(const Value: AnsiString);
    procedure SetExpandFlag(const Value: Boolean);
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
  public
    constructor Create(AOwner: TComponent); override;
  published
    property FileName;
    property SkipFirstRows default 0;
    property SheetName: AnsiString read FSheetName write SetSheetName;
    property NotExpandMergedValue: Boolean read FNotExpandMergedValue
      write SetExpandFlag default true;
  end;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF ODS}
{$IFDEF VCL6}

uses
  QImport3WideStrUtils;

{ TODSCell }

constructor TODSCell.Create(Collection: TCollection);
begin
  inherited;
  FValue := '';
  FCol := 0;
  FRow := 0;
end;

destructor TODSCell.Destroy;
begin
  inherited;
end;

{ TODSCellList }

function TODSCellList.Add: TODSCell;
begin
  Result := TODSCell(inherited Add);
end;

function TODSCellList.GetItem(Index: Integer): TODSCell;
begin
  Result := TODSCell(inherited Items[Index]);
end;

procedure TODSCellList.SetItem(Index: Integer; const Value: TODSCell);
begin
  inherited Items[Index] := Value;
end;

{ TODSSpreadSheet }

constructor TODSSpreadSheet.Create(Collection: TCollection);
begin
  inherited;
  FColCount := 0;
  FRowCount := 0;
  FX := -1;
  FY := -1;
  FCells := TODSCellList.Create(TODSCell);
  FDataCells := TqiStringGrid.Create(nil);
end;

destructor TODSSpreadSheet.Destroy;
begin
  FCells.Free;
  FDataCells.Free;
  inherited;
end;

procedure TODSSpreadSheet.LoadDataCells;
var
  i: Integer;
begin
  FDataCells.ColCount := FColCount;
  FDataCells.RowCount := FRowCount;
  for i := 0 to FCells.Count - 1 do
    FDataCells.Cells[FCells[i].Col, Cells[i].Row] := Cells[i].Value;
end;

{ TODSSpreadSheetList }

function TODSSpreadSheetList.GetItems(Index: integer): TODSSpreadSheet;
begin
  Result := TODSSpreadSheet(inherited Items[Index]);
end;

procedure TODSSpreadSheetList.SetItems(Index: integer;
  Value: TODSSpreadSheet);
begin
  inherited Items[Index] := Value;
end;

function TODSSpreadSheetList.Add: TODSSpreadSheet;
begin
  Result := TODSSpreadSheet(inherited Add);
end;

function TODSSpreadSheetList.GetSheetByName(Name: AnsiString): TODSSpreadSheet;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Self.Count - 1 do
    if (UpperCase(Items[i].Name) = UpperCase(Name)) then
    begin
      Result := Items[i];
      Break;
    end;
end;

{ TODSWorkbook }

procedure TODSWorkBook.FindTables(NameOfFile: qiString);
var
  TableNodes: IXMLDOMNodeList;
  I: Integer;
  NumberOfSprSh: Integer;
begin
  NumberOfSprSh := -1;
  FXMLDoc.load(NameOfFile);
  TableNodes := FXMLDoc.selectNodes('//table:table');
  for I := 0 to TableNodes.length - 1 do
  begin
   if not
      assigned(TableNodes[I].attributes.getNamedItem('table:is-sub-table')) then
    begin
      FSpreadSheets.Add;
      Inc(NumberOfSprSh);
      FSpreadSheets[NumberOfSprSh].FName := AnsiString(
        TableNodes[I].attributes.getNamedItem('table:name').nodeValue);
      ParseTable(FSpreadSheets[NumberOfSprSh], TableNodes[I].childNodes, 0, 0, false);
      FSpreadSheets[NumberOfSprSh].LoadDataCells;
    end;
  end;
  if TableNodes.length = 0 then
    raise Exception.Create('No spreadsheets were found');
end;

function TODSWorkBook.GetNodeFullTextValue(const Node: IXMLDOMNode): qiString;
var
  I: Integer;
  ChildNode,
  ParentNode: IXMLDOMNode;
begin
  ParentNode := Node.parentNode;
  if Assigned(ParentNode) and ParentNode.hasChildNodes then
    for I := 0 to ParentNode.childNodes.length - 1 do
    begin
      ChildNode := ParentNode.childNodes[I];
      if ChildNode.nodeName = '#text' then
        Result := Result + ChildNode.nodeValue;
    end
  else
    Result := '';
end;

procedure TODSWorkBook.ParseTable(Table: TODSSpreadSheet; Nodes: IXMLDOMNodeList;
  NumbOfRepColumns, NumbOfRepRows: Integer; IsSpanning: Boolean);
var
  I: Integer;
  NORR, NORC: Integer;
  IsSp: Boolean;
  TempCell: TODSCell;
  Node,
  NamedItem: IXMLDOMNode;

  function CheckNotEmptySequence(Nds: IXMLDOMNodeList; start: Integer): Boolean;
  var
    j: Integer;
  begin
    Result := false;
    for j := start to Nds.length - 1 do
    begin
      if Nodes[j].hasChildNodes then
        Result := (Nds[j].childNodes.length > 1)
          or Nds[j].childNodes[0].hasChildNodes;
      if Result then
        break;
    end;
  end;

begin
  NORC := NumbOfRepColumns;
  NORR := NumbOfRepRows;
  IsSp := IsSpanning;
  for I := 0 to Nodes.length - 1 do
  begin
    NumbOfRepColumns := NORC;
    NumbOfRepRows := NORR;
    IsSpanning := IsSp;
    Node := Nodes[I];

    if Node.nodeName = 'text:s' then
      Exit;
      
    if Node.NodeName <> '#text' then
    begin
      if Node.NodeName = 'table:table-row' then
      begin
        Table.Y := Table.Y + 1;
        Table.X := -1;
        if not CheckNotEmptySequence(Nodes, i) then
          break;

        NamedItem := Node.attributes.getNamedItem('table:number-rows-repeated');
        if Assigned(NamedItem) then
          NumbOfRepRows := NamedItem.nodeValue - 1;
        if Table.Y + 1 > Table.RowCount then
          Table.RowCount := Table.Y + 1;
      end;

      if (Node.NodeName = 'table:table-cell')
        or (Node.NodeName = 'table:covered-table-cell') then
      begin
        Table.X := Table.X + 1;
        if Table.X + 1 > Table.ColCount then
          Table.ColCount := Table.X + 1;

        NamedItem := Node.attributes.getNamedItem('table:number-columns-repeated');
        if Assigned(NamedItem) then
        NumbOfRepColumns := NamedItem.nodeValue - 1;

        NamedItem := Node.attributes.getNamedItem('table:number-rows-spanned');
        if Assigned(NamedItem) then
        begin
          NumbOfRepRows := NamedItem.nodeValue - 1;
          IsSpanning := true;
        end;

        NamedItem := Node.attributes.getNamedItem('table:number-columns-spanned');
        if Assigned(NamedItem) then
        begin
          NumbOfRepColumns := NamedItem.nodeValue - 1;
          IsSpanning := true;
        end;
       if not Node.HasChildNodes then
         if (NumbOfRepRows > 0) or (NumbOfRepColumns > 0) then
           if not IsSpanning then
           begin
             if Table.X + NumbOfRepColumns + 1 > Table.ColCount then
               Table.ColCount := Table.X + NumbOfRepColumns + 1;
             if Table.Y + NumbOfRepRows + 1 > Table.RowCount then
               Table.RowCount := Table.Y + NumbOfRepRows + 1;
             Table.X := Table.X + NumbOfRepColumns;
             Table.Y := Table.Y + NumbOfRepRows;
           end
      end;

      if Node.HasChildNodes then
        ParseTable(Table, Node.ChildNodes, NumbOfRepColumns,
           NumbOfRepRows, IsSpanning);
    end
    else
    begin
      if (NumbOfRepRows > 0) or (NumbOfRepColumns > 0) then
      begin
        ExpandRowsNCols(Table, GetNodeFullTextValue(Node),
          NumbOfRepRows, NumbOfRepColumns, IsSpanning);
        if not IsSpanning then
        begin
          Table.X := Table.X + NumbOfRepColumns;
          Table.Y := Table.Y + NumbOfRepRows;
        end;
      end
      else
      begin
        TempCell := Table.Cells.Add;
        TempCell.Row := Table.Y;
        TempCell.Col := Table.X;
        TempCell.Value := GetNodeFullTextValue(Node);
      end;
    end;
  end;
end;

procedure TODSWorkBook.ExpandRowsNCols(Table: TODSSpreadSheet;
  ExpandValue: qiString;
  NumbOfRepRows, NumbOfRepColumns: Integer; IsSpanning: Boolean);
var
  I: Integer;
  TempCell: TODSCell;
begin
  I := NumbOfRepRows;
  if Table.X + NumbOfRepColumns + 1 > Table.ColCount then
    Table.ColCount := Table.X + NumbOfRepColumns + 1;
  if Table.Y + NumbOfRepRows + 1 > Table.RowCount then
    Table.RowCount := Table.Y + NumbOfRepRows + 1;
  if NumbOfRepColumns >= 0 then
  begin
    repeat
      if not IsSpanning
      then
      begin
        TempCell := Table.Cells.Add;
        TempCell.Row := Table.Y + I;
        TempCell.Col := Table.X + NumbOfRepColumns;
        TempCell.Value := ExpandValue;
      end
      else
      begin
        if IsNotExpanding then
        begin
          if (NumbOfRepColumns = 0) and (I = 0) then
          begin
            TempCell := Table.Cells.Add;
            TempCell.Row := Table.Y;
            TempCell.Col := Table.X;
            TempCell.Value := ExpandValue;
          end;
        end
        else
        begin
          TempCell := Table.Cells.Add;
          TempCell.Row := Table.Y + I;
          TempCell.Col := Table.X + NumbOfRepColumns;
          TempCell.Value := ExpandValue;
        end;
      end;
      I := I - 1;
    until (I < 0);
    ExpandRowsNCols(Table, ExpandValue, NumbOfRepRows,
      NumbOfRepColumns - 1, IsSpanning);
  end;
end;

procedure TODSWorkbook.SetSpreadSheets;
var
  SRec: TSearchRec;
begin
  try
    if FindFirst(FWorkDir + 'content.xml', faDirectory, SRec) = 0 then
    begin
      FindTables(FWorkDir + 'content.xml');
    end
    else
      FindTables(FWorkDir + FileName);
  finally
    FindClose(SRec);
  end;
end;

constructor TODSWorkbook.Create;
begin
  FXMLDoc := CoDOMDocument.Create;
  FSpreadSheets := TODSSpreadSheetList.Create(TODSSpreadSheet);
  FIsNotExpanding := true;
  FWorkDir := '';
  FileName := '';
end;

destructor TODSWorkbook.Destroy;
begin
  FXMLDoc := nil;
  FSpreadSheets.Free;
  inherited;
end;

procedure TODSWorkbook.Load;
begin
  if FWorkDir <> '' then
  begin
    SetSpreadSheets;
  end;
end;

{ TODSFile }

procedure TODSFile.LoadXML(WorkDir: qiString);
begin
  FWorkbook.WorkDir := WorkDir;
  FWorkbook.Load;
end;

constructor TODSFile.Create;
begin
  inherited;
  FWorkbook := TODSWorkbook.Create;
end;

destructor TODSFile.Destroy;
begin
  FWorkbook.Free;
  inherited;
end;

{TQImport3ODS}

procedure TQImport3ODS.AfterImport;
begin
  FODSFile.Free;
  inherited;
end;

procedure TQImport3ODS.BeforeImport;
begin
  inherited;
  FODSFile := TODSFile.Create;
  FODSFile.FileName := FileName;
  FODSFile.Workbook.IsNotExpanding := NotExpandMergedValue;
  FODSFile.Load;
  if Assigned(FODSFile.Workbook.SpreadSheets.GetSheetByName(FSheetName)) then
    FTotalRecCount := FODSFile.Workbook.SpreadSheets.GetSheetByName(FSheetName).RowCount -
      SkipFirstRows;
end;

procedure TQImport3ODS.ChangeCondition;
begin
  Inc(FCounter);
end;

function TQImport3ODS.CheckCondition: Boolean;
begin
  Result := FCounter < (FTotalRecCount + SkipFirstRows);
end;

constructor TQImport3ODS.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
  NotExpandMergedValue := true;
end;

procedure TQImport3ODS.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    SkipFirstRows := ReadInteger(ODS_OPTIONS, ODS_SKIP_LINES, SkipFirstRows);
    SheetName := AnsiString( ReadString(ODS_OPTIONS, ODS_SHEET_NAME,
      string(SheetName)));
    NotExpandMergedValue :=
      ReadBool(ODS_OPTIONS, ODS_NOT_EXPAND_MERGED_VALUE, NotExpandMergedValue);
  end;
end;

procedure TQImport3ODS.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    WriteInteger(ODS_OPTIONS, ODS_SKIP_LINES, SkipFirstRows);
    WriteString(ODS_OPTIONS, ODS_SHEET_NAME, string(SheetName));
    WriteBool(ODS_OPTIONS, ODS_NOT_EXPAND_MERGED_VALUE, NotExpandMergedValue);
  end;
end;

procedure TQImport3ODS.FillImportRow;
var
  i, k: Integer;
  strValue: qiString;
  p: Pointer;
  mapValue: qiString;
begin
  FImportRow.ClearValues;
  RowIsEmpty := True;
  for i := 0 to FImportRow.Count - 1 do
  begin
    if FImportRow.MapNameIdxHash.Search(FImportRow[i].Name, p) then
    begin
      k := Integer(p);
{$IFDEF VCL7}
      mapValue := Map.ValueFromIndex[k];
{$ELSE}
      mapValue := Map.Values[FImportRow[i].Name];
{$ENDIF}
      strValue :=
        FODSFile.Workbook.SpreadSheets.GetSheetByName(FSheetName).DataCells.Cells[
          GetColIdFromColIndex(mapValue) - 1, FCounter];
      if AutoTrimValue then
        strValue := Trim(strValue);

      RowIsEmpty := RowIsEmpty and (strValue = '');
      FImportRow.SetValue(Map.Names[k], strValue, False);
    end;
    DoUserDataFormat(FImportRow[i]);
  end;
end;

procedure TQImport3ODS.FinishImport;
begin
  if not Canceled and not IsCSV then
  begin
    if CommitAfterDone then
      DoNeedCommit
    else if (CommitRecCount > 0) and ((ImportedRecs + ErrorRecs) mod CommitRecCount > 0) then
      DoNeedCommit;
  end;
end;

function TQImport3ODS.ImportData: TQImportResult;
begin
  Result := qirOk;
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

procedure TQImport3ODS.SetSheetName(const Value: AnsiString);
begin
  if (FSheetName <> Value) then
    FSheetName := Value;
end;


procedure TQImport3ODS.SetExpandFlag(const Value: Boolean);
begin
  if (Value <> FNotExpandMergedValue) then
    FNotExpandMergedValue := Value;
end;

function TQImport3ODS.Skip: Boolean;
begin
  Result := (SkipFirstRows > 0) and (FCounter < SkipFirstRows);
end;

procedure TQImport3ODS.StartImport;
begin
  inherited;
  FCounter := 0;
end;

{$ENDIF}
{$ENDIF}

end.

