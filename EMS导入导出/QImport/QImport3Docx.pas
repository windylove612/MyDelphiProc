unit QImport3Docx;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF DOCX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Classes,
    System.IniFiles,
    Winapi.msxml,
    Data.DB,
  {$ELSE}
    Classes,
    IniFiles,
    DB,
    msxml,
  {$ENDIF}
  QImport3,
  QImport3Common,
  QImport3StrTypes,
  QImport3BaseDocumentFile;

type
  TDocxCol = class(TCollectionItem)
  private
    FText: qiString;
    FDisplayAsIcon: Boolean;
  public
    property Text: qiString read FText write FText;
    property DisplayAsIcon: Boolean read FDisplayAsIcon write FDisplayAsIcon default False;
  end;

  TDocxColList = class(TCollection)
  private
    function GetItem(Index: Integer): TDocxCol;
    procedure SetItem(Index: Integer; const Value: TDocxCol);
  public
    function Add: TDocxCol;
    property Items[Index: Integer]: TDocxCol read GetItem write SetItem; default;
  end;

  TDocxRow = class(TCollectionItem)
  private
    FCols: TDocxColList;
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    property Cols: TDocxColList read FCols;
  end;

  TDocxRowList = class(TCollection)
  private
    function GetItem(Index: Integer): TDocxRow;
    procedure SetItem(Index: Integer; const Value: TDocxRow);
  public
    function Add: TDocxRow;
    property Items[Index: Integer]: TDocxRow read GetItem write SetItem; default;
  end;

  TDocxTable = class(TCollectionItem)
  private
    FRows: TDocxRowList;
  public
    constructor Create(Collection: TCollection); override;
    destructor Destroy; override;
    property Rows: TDocxRowList read FRows;
  end;

  TDocxTableList = class(TCollection)
  private
    function GetItem(Index: Integer): TDocxTable;
    procedure SetItem(Index: Integer; const Value: TDocxTable);
  public
    function Add: TDocxTable;
    property Items[Index: Integer]: TDocxTable read GetItem write SetItem; default;
  end;

  TDocxFile = class(TBaseDocumentFile)
  private
    FTables: TDocxTableList;
    FXMLDoc: IXMLDOMDocument;
    FNeedFillMerge: Boolean;
    procedure SetNeedFillMerge(const Value: Boolean);
  protected
    procedure LoadXML(CurrFolder: qiString); override;
  public
    constructor Create; override;
    destructor Destroy; override;
    property Tables: TDocxTableList read FTables;
    property NeedFillMerge: Boolean read FNeedFillMerge
      write SetNeedFillMerge;
  end;

  TQImport3Docx = class(TQImport3)
  private
    FDocxFile: TDocxFile;
    FCounter: Integer;
    FNeedFillMerge: Boolean;
    FTableNumber: integer;
    procedure SetTableNumber(const Value: integer);
    procedure SetNeedFillMerge(const Value: Boolean);
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
    function GetStringValue(const AValue: qiString; const AFieldType: TFieldType): Variant; override;
  public
    constructor Create(AOwner: TComponent); override;
  published
    property FileName;
    property SkipFirstRows default 0;
    property TableNumber: integer read FTableNumber
      write SetTableNumber default 0;
    property NeedFillMerge: Boolean read FNeedFillMerge
      write SetNeedFillMerge default False;
  end;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF DOCX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.SysUtils;
  {$ELSE}
    SysUtils;
  {$ENDIF}

{ TDocxColList }

function TDocxColList.Add: TDocxCol;
begin
  Result := TDocxCol(inherited Add)
end;

function TDocxColList.GetItem(Index: Integer): TDocxCol;
begin
  Result := TDocxCol(inherited Items[Index]);
end;

procedure TDocxColList.SetItem(Index: Integer; const Value: TDocxCol);
begin
  inherited Items[Index] := Value;
end;

{ TDocxRow }

constructor TDocxRow.Create(Collection: TCollection);
begin
  inherited;
  FCols := TDocxColList.Create(TDocxCol);
end;

destructor TDocxRow.Destroy;
begin
  if Assigned(FCols) then
    FCols.Free;
  inherited;
end;

{ TDocxRowList }

function TDocxRowList.Add: TDocxRow;
begin
  Result := TDocxRow(inherited Add)
end;

function TDocxRowList.GetItem(Index: Integer): TDocxRow;
begin
  Result := TDocxRow(inherited Items[Index]);
end;

procedure TDocxRowList.SetItem(Index: Integer; const Value: TDocxRow);
begin
  inherited Items[Index] := Value;
end;

{ TDocxTable }

constructor TDocxTable.Create(Collection: TCollection);
begin
  inherited;
  FRows := TDocxRowList.Create(TDocxRow);
end;

destructor TDocxTable.Destroy;
begin
  if Assigned(FRows) then
    FRows.Free;
  inherited;
end;

{ TDocxTableList }

function TDocxTableList.Add: TDocxTable;
begin
  Result := TDocxTable(inherited Add)
end;

function TDocxTableList.GetItem(Index: Integer): TDocxTable;
begin
  Result := TDocxTable(inherited Items[Index]);
end;

procedure TDocxTableList.SetItem(Index: Integer; const Value: TDocxTable);
begin
  inherited Items[Index] := Value;
end;

{ TDocxFile }

procedure TDocxFile.SetNeedFillMerge(const Value: Boolean);
begin
  if FNeedFillMerge <> Value then
    FNeedFillMerge := Value;
end;

procedure TDocxFile.LoadXML(CurrFolder: qiString);


 function TryImageStr(ANode: IXMLDOMNode; out AIsIcon: Boolean): qiString;
 var
   TempNode: IXMLDOMNode;
   PicName: string;
   F: TSearchRec;
 begin
   Result := '';
   TempNode := ANode.selectSingleNode('w:drawing/wp:inline/wp:docPr');
   if Assigned(TempNode) and Assigned(TempNode.attributes)
    and Assigned(TempNode.attributes.getNamedItem('name')) then
      PicName := TempNode.attributes.getNamedItem('name').nodeValue;
   if PicName <> '' then
     if FindFirst(CurrFolder + 'word\media\' + PicName + '.*', faAnyFile, F) = 0 then
     begin
       Result := FileToBase64(CurrFolder + 'word\media\' +F.Name);
       AIsIcon := True;
     end;
 end;


 function GetCellText(TransfNodes: IXMLDOMNodeList; out IsIcon: Boolean): qiString; //w:p
 var
   i, j: Integer;
   TextNodes: IXMLDOMNodeList;
   Node: IXMLDOMNode;
 begin
   Result := '';
   for i := 0 to TransfNodes.length - 1 do
   begin
     TextNodes := TransfNodes[i].selectNodes('w:r');
     for j := 0 to TextNodes.length - 1 do
     begin
       Node := TextNodes[j].selectSingleNode('w:t');
       if Assigned(Node) then
       begin
         if (i <> 0) and (Result <> '') then
           Result := Result + #13;
         Result := Result + Node.text
       end else
         Result := TryImageStr(TextNodes[j], IsIcon);
     end;
   end;

 end;

var
  TableNodes, RowNodes, ColNodes,
  ValueNodes: IXMLDOMNodeList;
  XmlRec: TSearchRec;
  gSpan: IXMLDOMNode;
  i, j, k, n: Integer;
  TempCol: TDocxCol;
  IsIcon: Boolean;
begin
  try
    if FindFirst(CurrFolder + 'word\document.xml', faAnyFile, XmlRec) = 0 then
//  if FileExists(CurrFolder + 'word\document.xml') then
    begin
      FXMLDoc.load(CurrFolder + 'word\document.xml');
      TableNodes := FXMLDoc.selectNodes('//w:tbl');
      for i := 0 to TableNodes.length - 1 do
      begin
        FTables.Add;
        RowNodes := TableNodes[i].selectNodes('w:tr');
        for j := 0 to RowNodes.length - 1 do
        begin
          FTables[i].Rows.Add;
          ColNodes := RowNodes[j].selectNodes('w:tc');
          for k := 0 to ColNodes.length - 1 do
          begin
            IsIcon := False;
            ValueNodes := ColNodes[k].selectNodes('w:p');
            TempCol := FTables[i].Rows[j].Cols.Add;
            TempCol.Text := GetCellText(ValueNodes, IsIcon);
            TempCol.DisplayAsIcon := IsIcon;

            if FNeedFillMerge then
              if Assigned(ColNodes[k].selectSingleNode('w:tcPr/w:vmerge')) then
                if not Assigned(ColNodes[k].selectSingleNode('w:tcPr/w:vmerge').attributes.getNamedItem('w:val')) then
                  TempCol.Text :=
                    FTables[i].Rows[j - 1].Cols[FTables[i].Rows[j].Cols.Count - 1].Text;

            gSpan := ColNodes[k].selectSingleNode('w:tcPr/w:gridSpan');
            if Assigned(gSpan) then
              if Assigned(gSpan.attributes.getNamedItem('w:val')) then
                for n := 0 to gSpan.attributes.getNamedItem('w:val').nodeValue - 2 do // -2 because first cell is already added
                begin
                  TempCol := FTables[i].Rows[j].Cols.Add;
                  if FNeedFillMerge then
                    TempCol.Text :=
                      FTables[i].Rows[j].Cols[FTables[i].Rows[j].Cols.Count - 2].Text;
                end;
          end;
        end;
      end;
    end;
  finally
    FindClose(XmlRec);
  end;
end;

constructor TDocxFile.Create;
begin
  inherited;
  FTables := TDocxTableList.Create(TDocxTable);
  FXMLDoc := CoDOMDocument.Create;
end;

destructor TDocxFile.Destroy;
begin
  FXMLDoc := nil;
  FTables.Free;
  inherited;
end;

{ TQImport3Docx }

procedure TQImport3Docx.SetTableNumber(const Value: integer);
begin
  FTableNumber := Value;
end;

procedure TQImport3Docx.SetNeedFillMerge(const Value: Boolean);
begin
  FNeedFillMerge := Value;
end;

procedure TQImport3Docx.BeforeImport;
begin
  FDocxFile := TDocxFile.Create;
  FDocxFile.FileName := FileName;
  FDocxFile.NeedFillMerge := FNeedFillMerge;
  FDocxFile.Load;
  if Assigned(FDocxFile) and (FTableNumber > 0) then
    FTotalRecCount := FDocxFile.FTables[Pred(FTableNumber)].Rows.Count - SkipFirstRows;
  inherited;
end;

procedure TQImport3Docx.StartImport;
begin
  FCounter := 0;
end;

function TQImport3Docx.CheckCondition: Boolean;
begin
  Result := FCounter < (FTotalRecCount + SkipFirstRows);
end;

function TQImport3Docx.Skip: Boolean;
begin
  Result := (SkipFirstRows > 0) and (FCounter < SkipFirstRows);
end;

procedure TQImport3Docx.ChangeCondition;
begin
  Inc(FCounter);
end;

procedure TQImport3Docx.FinishImport;
begin
  if not Canceled and not IsCSV then
  begin
    if CommitAfterDone then
      DoNeedCommit
    else if (CommitRecCount > 0) and ((ImportedRecs + ErrorRecs) mod CommitRecCount > 0) then
      DoNeedCommit;
  end;
end;

procedure TQImport3Docx.AfterImport;
begin
  FDocxFile.Free;
  FDocxFile := nil;
  inherited;
end;

procedure TQImport3Docx.FillImportRow;
var
  j, k: Integer;
  strValue: qiString;
  p: Pointer;
  mapValue: qiString;
begin
  FImportRow.ClearValues;
  RowIsEmpty := True;
  for j := 0 to FImportRow.Count - 1 do
  begin
    if FImportRow.MapNameIdxHash.Search(FImportRow[j].Name, p) then
    begin
      k := Integer(p);
      strValue := '';
{$IFDEF VCL7}
      mapValue := Map.ValueFromIndex[k];
{$ELSE}
      mapValue := Map.Values[FImportRow[j].Name];
{$ENDIF}
      if FDocxFile.Tables[Pred(FTableNumber)].Rows[FCounter].Cols.Count >= StrToInt(mapValue) then
        strValue := FDocxFile.Tables[Pred(FTableNumber)].Rows[FCounter].Cols[Pred(StrToInt(mapValue))].Text;
      if AutoTrimValue then
        strValue := Trim(strValue);
      RowIsEmpty := RowIsEmpty and (strValue = '');
      FImportRow.SetValue(Map.Names[k], strValue, False);
    end;
    DoUserDataFormat(FImportRow[j]);
  end;
end;

function TQImport3Docx.ImportData: TQImportResult;
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

procedure TQImport3Docx.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    SkipFirstRows := ReadInteger(DOCX_OPTIONS, DOCX_SKIP_LINES, SkipFirstRows);
    TableNumber := ReadInteger(DOCX_OPTIONS, DOCX_TABLE_NUMBER, TableNumber);
    NeedFillMerge := ReadBool(DOCX_OPTIONS, DOCX_NEED_FILLMERGE, NeedFillMerge);
  end;
end;

procedure TQImport3Docx.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    WriteInteger(DOCX_OPTIONS, DOCX_SKIP_LINES, SkipFirstRows);
    WriteInteger(DOCX_OPTIONS, DOCX_TABLE_NUMBER, TableNumber);
    WriteBool(DOCX_OPTIONS, DOCX_NEED_FILLMERGE, NeedFillMerge);
  end;
end;

function TQImport3Docx.GetStringValue(const AValue: qiString; const AFieldType: TFieldType): Variant;
begin
  case AFieldType of
      ftBlob,ftOraClob,ftOraBlob:  Result := QIDecodeBase64(AValue);
    else
      Result := inherited GetStringValue(AValue, AFieldType);
  end;
end;


constructor TQImport3Docx.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
  FTableNumber := 0;
end;

{$ENDIF}
{$ENDIF}

end.
