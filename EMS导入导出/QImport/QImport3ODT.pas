unit QImport3ODT;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF ODT}
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
  QImport3,
  QImport3Common;

type
  TODTCell = class(TCollectionItem)
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

  TODTCellList = class(TCollection)
  private
    function GetItem(Index: Integer): TODTCell;
    procedure SetItem(Index: Integer; const Value: TODTCell);
  public
    function Add: TODTCell;
    property Items[Index: Integer]: TODTCell read GetItem write SetItem; default;
    function GetItemByCoords(X, Y: Integer): TODTCell;
  end;

  TODTSpreadSheet = class(TCollectionItem)
  private
    FName: qiString;
    FColCount: Integer;
    FRowCount: Integer;
    FCells: TODTCellList;
    FDataCells: TqiStringGrid;
    FX: Integer;
    FY: Integer;
    FIsAfterSubTable: Boolean;
    FIsComplexTable: Boolean;
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    procedure LoadDataCells;
    property Cells: TODTCellList read FCells;
    property ColCount: Integer read FColCount write FColCount;
    property RowCount: Integer read FRowCount write FRowCount;
    property X: Integer read FX write FX;
    property Y: Integer read FY write FY;
    property Name: qiString read FName write FName;
    property IsAfterSubTable: Boolean read FIsAfterSubTable
      write FIsAfterSubTable;
    property DataCells: TqiStringGrid read FDataCells;
    property IsComplexTable: Boolean read FIsComplexTable write FIsComplexTable;
  end;

  TODTSpreadSheetList = class(TCollection)
  private
    function GetItems(Index: integer): TODTSpreadSheet;
    procedure SetItems(Index: integer; Value: TODTSpreadSheet);
  public
    function Add: TODTSpreadSheet;
    property Items[Index: integer]: TODTSpreadSheet read GetItems
      write SetItems; default;
    function GetSheetByName(Name: qiString): TODTSpreadSheet;
  end;

  TODTWorkbook = class
  private
    FWorkDir: qiString;
    FileName: qiString;
    FSpreadSheets: TODTSpreadSheetList;
    FXMLDoc: IXMLDOMDocument;
    procedure SetSpreadSheets;
    procedure FindTables(NameOfFile: qiString);
    procedure ParseTable(Table: TODTSpreadSheet; Nodes: IXMLDOMNodeList;
      NumbOfRepColumns, NumbOfRepRows: Integer; IsSpanning: Boolean);
    procedure ExpandRowsNCols(Table: TODTSpreadSheet;
      ExpandValue: qiString; NumbOfRepRows, NumbOfRepColumns: Integer;
        IsSpanning: Boolean);
  public
    constructor Create;
    destructor Destroy; override;
    procedure Load;

    property WorkDir: qiString read FWorkDir write FWorkDir;
    property SpreadSheets: TODTSpreadSheetList read FSpreadSheets;
  end;

  TODTFile = class(TBaseDocumentFile)
    private
      FWorkbook: TODTWorkbook;
    protected
      procedure LoadXML(WorkDir: qiString); override;
    public
      constructor Create; override;
      destructor Destroy; override;

      property Workbook: TODTWorkbook read FWorkbook;
  end;

  TQImport3ODT = class(TQImport3)
  private
    FODTFile: TODTFile;
    FCounter: Integer;
    FSheetName: AnsiString;
    FUseHeader: Boolean;
    FUseComplexTables: Boolean;
    procedure SetSheetName(const Value: AnsiString);
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
    property UseHeader: Boolean read FUseHeader write FUseHeader default false;
    property UseComplexTables: Boolean read FUseComplexTables
      write FUseComplexTables default true;
  end;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF ODT}
{$IFDEF VCL6}

uses
  QImport3WideStrUtils;


{ TODTCell }

constructor TODTCell.Create(Collection: TCollection);
begin
  inherited;
  FValue := '';
  FCol := 0;
  FRow := 0;
end;

destructor TODTCell.Destroy;
begin
  inherited;
end;

{ TODTCellList }

function TODTCellList.Add: TODTCell;
begin
  Result := TODTCell(inherited Add);
end;

function TODTCellList.GetItem(Index: Integer): TODTCell;
begin
  Result := TODTCell(inherited Items[Index]);
end;

function TODTCellList.GetItemByCoords(X, Y: Integer): TODTCell;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Self.Count - 1 do
    if (Items[i].Row = Y) and (Items[i].Col = X) then
    begin
      Result := Items[i];
      Break;
    end;
end;

procedure TODTCellList.SetItem(Index: Integer; const Value: TODTCell);
begin
  inherited Items[Index] := Value;
end;

{ TODTSpreadSheet }

constructor TODTSpreadSheet.Create(Collection: TCollection);
begin
  inherited;
  FColCount := 0;
  FRowCount := 0;
  FX := -1;
  FY := -1;
  FIsAfterSubTable := false;
  FCells := TODTCellList.Create(TODTCell);
  FDataCells := TqiStringGrid.Create(nil);
  FIsComplexTable := false;
end;

destructor TODTSpreadSheet.Destroy;
begin
  FCells.Free;
  FDataCells.Free;
  inherited;
end;

procedure TODTSpreadSheet.LoadDataCells;
var
  i: Integer;
begin
  FDataCells.ColCount := FColCount;
  FDataCells.RowCount := FRowCount;
  for i := 0 to FCells.Count - 1 do
    FDataCells.Cells[FCells[i].Col, Cells[i].Row] := Cells[i].Value;
end;

{ TODTSpreadSheetList }

function TODTSpreadSheetList.GetItems(Index: integer): TODTSpreadSheet;
begin
  Result := TODTSpreadSheet(inherited Items[Index]);
end;

procedure TODTSpreadSheetList.SetItems(Index: integer;
  Value: TODTSpreadSheet);
begin
  inherited Items[Index] := Value;
end;

function TODTSpreadSheetList.Add: TODTSpreadSheet;
begin
  Result := TODTSpreadSheet(inherited Add);
end;

function TODTSpreadSheetList.GetSheetByName(Name: qiString): TODTSpreadSheet;
var
  i: Integer;
begin
  Result := nil;
  for i := 0 to Self.Count - 1 do
    if UpperCase(Items[i].Name) = UpperCase(Name) then
    begin
      Result := Items[i];
      Break;
    end;
end;

{ TODTWorkbook }

procedure TODTWorkBook.FindTables(NameOfFile: qiString);
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
      FSpreadSheets[NumberOfSprSh].Name :=
        TableNodes[I].attributes.getNamedItem('table:name').nodeValue;
      ParseTable(FSpreadSheets[NumberOfSprSh],TableNodes[I].childNodes, 0, 0, false);
      FSpreadSheets[NumberOfSprSh].LoadDataCells;
    end;
  end;
  if TableNodes.length = 0 then
    raise Exception.Create('No tables were found'); 
end;

procedure TODTWorkBook.ParseTable(Table: TODTSpreadSheet; Nodes: IXMLDOMNodeList;
  NumbOfRepColumns, NumbOfRepRows: Integer; IsSpanning: Boolean);
var
  I, J: Integer;
  NORR, NORC: Integer;
  IsSp: Boolean;
  TempSpreadSheet: TODTSpreadSheet;
  TempCell: TODTCell;
begin
  NORC := NumbOfRepColumns;
  NORR := NumbOfRepRows;
  IsSp := IsSpanning;
  for I := 0 to Nodes.length - 1 do
  begin
    NumbOfRepColumns := NORC;
    NumbOfRepRows := NORR;
    IsSpanning := IsSp;
    if Nodes[I].NodeName <> '#text' then
    begin
      if Nodes[I].NodeName = 'table:table-row' then
      begin
        Table.IsAfterSubTable := false;
        if Table.RowCount - 1 > Table.Y then
          Table.Y := Table.RowCount
        else
          Table.Y := Table.Y + 1;
        Table.X := -1;
        if assigned(Nodes[I].attributes.getNamedItem('table:number-rows-repeated')) then
          NumbOfRepRows :=
            Nodes[I].attributes.getNamedItem('table:number-rows-repeated').nodeValue - 1;
        if Table.Y + 1 > Table.RowCount then
          Table.RowCount := Table.Y + 1;
      end;
      if (Nodes[I].NodeName = 'table:covered-table-cell') then
        if not Table.IsAfterSubTable then
        begin
          Table.X := Table.X + 1;
          if Table.X + 1 > Table.ColCount then
            Table.ColCount := Table.X + 1;
        end;
      if (Nodes[I].NodeName = 'table:table-cell') then
      begin
        Table.IsAfterSubTable := false;
        Table.X := Table.X + 1;
        if Table.X + 1 > Table.ColCount then
          Table.ColCount := Table.X + 1;
        if Assigned(Nodes[I].attributes.getNamedItem('table:number-columns-repeated')) then
        NumbOfRepColumns :=
           Nodes[I].attributes.getNamedItem('table:number-columns-repeated').nodeValue - 1;
        if Assigned(Nodes[I].attributes.getNamedItem('table:number-rows-spanned')) then
        begin
          NumbOfRepRows :=
            Nodes[I].attributes.getNamedItem('table:number-rows-spanned').nodeValue - 1;
          IsSpanning := true;
        end;
        if Assigned(Nodes[I].attributes.getNamedItem('table:number-columns-spanned')) then
        begin
          NumbOfRepColumns :=
            Nodes[I].attributes.getNamedItem('table:number-columns-spanned').nodeValue - 1;
          IsSpanning := true;
        end;
        if not Nodes[I].HasChildNodes then
          if (NumbOfRepRows > 0) or (NumbOfRepColumns > 0) then
          begin
              Table.X := Table.X + NumbOfRepColumns;
              Table.Y := Table.Y + NumbOfRepRows;
              if Table.X + 1 > Table.ColCount then
                Table.ColCount := Table.X + 1;
              if Table.RowCount - 1 > Table.Y then
               Table.Y := Table.RowCount;
              NumbOfRepColumns := 0;
              NumbOfRepRows := 0;
          end;
      end;
      if (Nodes[I].NodeName = 'table:table') then
        if Assigned(Nodes[I].attributes.getNamedItem('table:is-sub-table')) then
        begin
          Table.IsComplexTable := true;
          Table.IsAfterSubTable := true;
          TempSpreadSheet := TODTSpreadSheet.Create(nil);
          ParseTable(TempSpreadSheet, Nodes[I].childNodes, 0, 0, false);
          for J := 0 to TempSpreadSheet.Cells.Count - 1 do
          begin
            TempCell := Table.Cells.Add;
            TempCell.Row  := Table.Y + TempSpreadSheet.Cells[J].Row;
            TempCell.Col := Table.X + TempSpreadSheet.Cells[J].Col;
            TempCell.Value := TempSpreadSheet.Cells[J].Value;
          end;
          if Table.Y + TempSpreadSheet.FRowCount > Table.RowCount then
            Table.RowCount := Table.Y + TempSpreadSheet.RowCount;
          if Table.X + TempSpreadSheet.ColCount > Table.ColCount then
            Table.ColCount := Table.X + TempSpreadSheet.ColCount;
          Table.X := Table.X + TempSpreadSheet.ColCount - 1;
          TempSpreadSheet.Free;
          NumbOfRepRows := 0;
          NumbOfRepColumns := 0;
        end
        else
          continue;
      if (Nodes[I].HasChildNodes) and not (Table.IsAfterSubTable) then
        ParseTable(Table, Nodes[I].ChildNodes, NumbOfRepColumns,
           NumbOfRepRows, IsSpanning);
    end
    else
    begin
      if (NumbOfRepRows > 0) or (NumbOfRepColumns > 0) then
      begin
        ExpandRowsNCols(Table, Nodes[I].NodeValue,
          NumbOfRepRows, NumbOfRepColumns, IsSpanning);
        if not IsSpanning then
        begin
          Table.X := Table.X + NumbOfRepColumns;
          Table.Y := Table.Y + NumbOfRepRows;
        end;
      end
      else
      begin
        if assigned(Table.Cells.GetItemByCoords(Table.X, Table.Y)) then
          Table.Cells.GetItemByCoords(Table.X, Table.Y).Value :=
            Table.Cells.GetItemByCoords(Table.X, Table.Y).Value + #13#10 + Nodes[I].nodeValue
        else
        begin
          TempCell := Table.Cells.Add;
          TempCell.Row := Table.Y;
          TempCell.Col := Table.X;
          TempCell.Value := Nodes[I].nodeValue;
        end;
      end;
    end;
  end;
end;

procedure TODTWorkBook.ExpandRowsNCols(Table: TODTSpreadSheet;
  ExpandValue: qiString;
  NumbOfRepRows, NumbOfRepColumns: Integer; IsSpanning: Boolean);
var
  I: Integer;
  TempCell: TODTCell;
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
        if assigned(Table.Cells.GetItemByCoords(Table.X + NumbOfRepColumns, Table.Y + I)) then
          Table.Cells.GetItemByCoords(Table.X + NumbOfRepColumns, Table.Y + I).Value :=
            Table.Cells.GetItemByCoords(Table.X + NumbOfRepColumns, Table.Y + I).Value + ' ' + ExpandValue
        else
        begin
          TempCell := Table.Cells.Add;
          TempCell.Row := Table.Y + I;
          TempCell.Col := Table.X + NumbOfRepColumns;
          TempCell.Value := ExpandValue;
        end
      else
        if (NumbOfRepColumns = 0) and (I = 0) then
        begin
          if assigned(Table.Cells.GetItemByCoords(Table.X, Table.Y)) then
            Table.Cells.GetItemByCoords(Table.X, Table.Y).Value :=
              Table.Cells.GetItemByCoords(Table.X, Table.Y).Value + ' ' + ExpandValue
          else
          begin
            TempCell := Table.Cells.Add;
            TempCell.Row := Table.Y;
            TempCell.Col := Table.X;
            TempCell.Value := ExpandValue;
          end;
        end;
      I := I - 1;
    until (I < 0);
    ExpandRowsNCols(Table, ExpandValue, NumbOfRepRows,
      NumbOfRepColumns - 1, IsSpanning);
  end;
end;

procedure TODTWorkbook.SetSpreadSheets;
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

constructor TODTWorkbook.Create;
begin
  FXMLDoc := CoDOMDocument.Create;
  FSpreadSheets := TODTSpreadSheetList.Create(TODTSpreadSheet);
  FWorkDir := '';
  FileName := '';
end;

destructor TODTWorkbook.Destroy;
begin
  FXMLDoc := nil;
  if Assigned(FSpreadSheets) then
    FSpreadSheets.Free;
  inherited;
end;

procedure TODTWorkbook.Load;
begin
  if FWorkDir <> '' then
  begin
    SetSpreadSheets;
  end;
end;

{ TODTFile }

procedure TODTFile.LoadXML(WorkDir: qiString);
begin
  FWorkbook.WorkDir := WorkDir;
  FWorkbook.Load;
end;

constructor TODTFile.Create;
begin
  inherited;
  FWorkbook := TODTWorkbook.Create;
end;

destructor TODTFile.Destroy;
begin
  if Assigned(FWorkbook) then
    FWorkbook.Free;
  inherited;
end;

{TQImport3ODT}

procedure TQImport3ODT.AfterImport;
begin
  FODTFile.Free;
  inherited;
end;

procedure TQImport3ODT.BeforeImport;
begin
  inherited;
  FODTFile := TODTFile.Create;
  FODTFile.FileName := FileName;
  FODTFile.Load;
  if Assigned(FODTFile.Workbook.SpreadSheets.GetSheetByName(qiString(FSheetName))) then
    FTotalRecCount := FODTFile.Workbook.SpreadSheets.GetSheetByName(qiString(FSheetName)).RowCount -
      SkipFirstRows;
end;

procedure TQImport3ODT.ChangeCondition;
begin
 inc(FCounter);
end;

function TQImport3ODT.CheckCondition: Boolean;
begin
  Result := FCounter < (FTotalRecCount + SkipFirstRows);
end;

constructor TQImport3ODT.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
  FUseHeader := false;
  FUseComplexTables := true;
end;

procedure TQImport3ODT.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    SkipFirstRows := ReadInteger(ODT_OPTIONS, ODT_SKIP_LINES, SkipFirstRows);
    SheetName := AnsiString(ReadString(ODT_OPTIONS, ODT_SHEET_NAME, string(SheetName)));
    UseHeader := ReadBool(ODT_OPTIONS, ODT_USE_HEADER, UseHeader);
    UseComplexTables := ReadBool(ODT_OPTIONS, ODT_COMPLEX_TABLE, UseComplexTables);
  end;
end;

procedure TQImport3ODT.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    WriteInteger(ODT_OPTIONS, ODT_SKIP_LINES, SkipFirstRows);
    WriteString(ODT_OPTIONS, ODT_SHEET_NAME, string(SheetName));
    WriteBool(ODT_OPTIONS, ODT_USE_HEADER, UseHeader);
    WriteBool(ODT_OPTIONS, ODT_COMPLEX_TABLE, UseComplexTables);
  end;
end;

procedure TQImport3ODT.FillImportRow;
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
      strValue := FODTFile.Workbook.SpreadSheets.GetSheetByName(
        qiString(FSheetName)).DataCells.Cells[GetColIdFromColIndex(mapValue) - 1,
        FCounter];
      if AutoTrimValue then
        strValue := Trim(strValue);
      RowIsEmpty := RowIsEmpty and (strValue = '');
      FImportRow.SetValue(Map.Names[k], strValue, False);
    end;
    DoUserDataFormat(FImportRow[i]);
  end;
end;

procedure TQImport3ODT.FinishImport;
begin
  if not Canceled and not IsCSV then
  begin
    if CommitAfterDone then
      DoNeedCommit
    else if (CommitRecCount > 0) and ((ImportedRecs + ErrorRecs) mod CommitRecCount > 0) then
      DoNeedCommit;
  end;
end;

function TQImport3ODT.ImportData: TQImportResult;
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

procedure TQImport3ODT.SetSheetName(const Value: AnsiString);
begin
  if (FSheetName <> Value) then
    FSheetName := Value;
end;


function TQImport3ODT.Skip: Boolean;
begin
  if not (UseHeader) then
    Result := (SkipFirstRows > 0) and (FCounter < SkipFirstRows)
  else
    Result := ((SkipFirstRows + 1) > 0) and (FCounter < (SkipFirstRows + 1));
end;

procedure TQImport3ODT.StartImport;
begin
  inherited;
  if Assigned(FODTFile.Workbook.SpreadSheets.GetSheetByName(qiString(FSheetName))) then
    if (FODTFile.Workbook.SpreadSheets.GetSheetByName(qiString(FSheetName)).IsComplexTable)
      and not (UseComplexTables) then
      raise EQImportError.Create('Trying to convert complex tables');
  FCounter := 0;
end;

{$ENDIF}
{$ENDIF}

end.






