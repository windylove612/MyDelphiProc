unit QImport3XMLDoc;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF XMLDOC}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Classes,
    Vcl.ComCtrls,
    System.IniFiles,
    Winapi.MSXML,
  {$ELSE}
    Classes,
    ComCtrls,
    IniFiles,
    MSXML,
  {$ENDIF}
  QImport3,
  QImport3StrTypes;

type
  TXMLDataLocation = (tlAttributes, tlSubNodes);

  TXMLDocFile = class
  private
    FFileName: string;
    FXPath: qiString;
    FDataLocation: TXMLDataLocation;
    
    FCells: array of array of qiString;
    FColCount: Integer;
    FRowCount: Integer;
    
    function GetCells(ACol, ARow: Integer): qiString;
    procedure SetFileName(const Value: string);
    procedure SetXPath(const Value: qiString);
  private
    FLoaded: Boolean;
    FIXML: IXMLDOMDocument;
    procedure Prepare;
    procedure LoadCells;
  public
    constructor Create;
    destructor Destroy; override;
    
    procedure Load;
    
    property FileName: string read FFileName write SetFileName;
    property XPath: qiString read FXPath write SetXPath;
    property DataLocation: TXMLDataLocation read FDataLocation write FDataLocation;

    property Loaded: boolean read FLoaded;
    property Cells[ACol, ARow: Integer]: qiString read GetCells;
    property ColCount: Integer read FColCount;
    property RowCount: Integer read FRowCount;
  end;

  TQImport3XMLDoc = class(TQImport3)
  private
    FXMLFile: TXMLDocFile;
    FCounter: Integer;
    FXPath: qiString;
    FDataLocation: TXMLDataLocation;
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
    property XPath: qiString read FXPath write FXPath;
    property DataLocation: TXMLDataLocation read FDataLocation write FDataLocation;
    property SkipFirstRows default 0;
  end;

procedure FillStringGrid(DataGrid: TqiStringGrid; XMLFile: TXMLDocFile);
procedure XMLFile2TreeView(TreeView: TTreeView; XMLFile: TXMLDocFile);

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF XMLDOC}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Variants,
    System.Math,
    System.SysUtils,
    Soap.EncdDecd,
  {$ELSE}
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
    Math,
    SysUtils,
    EncdDecd,
  {$ENDIF}
  QImport3Common;

const
  LF = #13#10;
  sFileNameNotDefined = 'File name is not defined';
  sFileNotFound = 'File %s not found';
  //NodeTypes
  qntAttribute  = 2;
  qntText       = 3;
  qntComment    = 8;

procedure FillStringGrid(DataGrid: TqiStringGrid; XMLFile: TXMLDocFile);
var
  i, j: Integer;
begin
  if not Assigned(XMLFile) then Exit;

  DataGrid.ColCount := XMLFile.ColCount;
  DataGrid.RowCount := Min(XMLFile.RowCount, 30);;
  for i := 0 to XMLFile.ColCount - 1 do
    for j := 0 to DataGrid.RowCount - 1 do
      DataGrid.Cells[i, j] := XMLFile.Cells[i, j];

  if DataGrid.RowCount > 1 then
    DataGrid.FixedRows := 1;
  if DataGrid.ColCount > 1 then
    DataGrid.FixedCols := 1;
end;

procedure XMLFile2TreeView(TreeView: TTreeView; XMLFile: TXMLDocFile);

  procedure ProcessNode(XMLNode: IXMLDOMNode; TreeNode: TTreeNode);
  var
    ChildXMLNode, XMLAtt: IXMLDOMNode;
    TreeNodeAtt: TTreeNode;
    i: Integer;
  const
    ElementImage = 0;
    TextImage = 1;
    AttImage = 2;
  begin
    if not Assigned(XMLNode) then Exit;
    if XMLNode.nodeType = 7 then Exit; //processing instruction

    if XMLNode.nodeType  = 1 then //element
    begin
      TreeNode := TreeView.Items.AddChild(TreeNode, XMLNode.nodeName);
      TreeNode.ImageIndex := ElementImage;
      TreeNode.SelectedIndex := ElementImage;
    end{ else
    if XMLNode.nodeType = 3 then //text
    begin
      TreeNode := TreeView.Items.AddChild(TreeNode, Trim(XMLNode.nodeValue));
      TreeNode.ImageIndex := TextImage;
      TreeNode.SelectedIndex := TextImage;
    end};
    if Assigned(XMLNode.attributes) then
      for i := 0 to XMLNode.attributes.length - 1 do
      begin
        XMLAtt := XMLNode.attributes[i];
        TreeNodeAtt := TreeView.Items.AddChild(TreeNode, Concat('@', XMLAtt.nodeName, ' = "',
          Trim(XMLAtt.nodeValue), '"'));
        TreeNodeAtt.ImageIndex := AttImage;
        TreeNodeAtt.SelectedIndex := AttImage;
        TreeNodeAtt.Data := Pointer(2);
      end;

    if XMLNode.hasChildNodes then
      ChildXMLNode := XMLNode.childNodes.item[0];
      
    while ChildXMLNode <> nil do
    begin
      ProcessNode(ChildXMLNode, TreeNode);
      ChildXMLNode := ChildXMLNode.NextSibling;
    end;
  end;

var
  XMLNode : IXMLDOMNode;
begin
  if Assigned(XMLFile) then
  begin
    TreeView.Items.BeginUpdate;
    TreeView.Items.Clear;
    if XMLFile.FIXML.hasChildNodes then
      XMLNode := XMLFile.FIXML.childNodes[0];
    while XMLNode <> nil do
    begin
      ProcessNode(XMLNode, nil);
      XMLNode := XMLNode.NextSibling;
    end;
    TreeView.Items.EndUpdate;
  end;
end;

{ TXMLFileDoc }

constructor TXMLDocFile.Create;
begin
  FIXML := CoDOMDocument.Create;
  FLoaded := False;
  FDataLocation := tlAttributes;
  FXPath := '/';
end;

destructor TXMLDocFile.Destroy;
begin
  FIXML := nil;
  inherited;
end;

function TXMLDocFile.GetCells(ACol, ARow: Integer): qiString;
begin
  Result := '';
  if Length(FCells) > ACol then
    if Length(FCells[ACol]) > ARow then
      Result := FCells[ACol, ARow];
end;

procedure TXMLDocFile.SetFileName(const Value: string);
begin
  if FFileName <> Value then
    FFileName := Value;
end;                                        

procedure TXMLDocFile.SetXPath(const Value: qiString);
begin
  if FXPath <> Value then
    FXPath := Value;
  if FLoaded then
    LoadCells;
end;

procedure TXMLDocFile.Prepare;
begin
  FLoaded := False;
  if FFileName = EmptyStr then
    raise Exception.Create(sFileNameNotDefined);
  if not FileExists(FFileName) then
    raise Exception.CreateFmt(sFileNotFound, [FFileName]);
end;

procedure TXMLDocFile.LoadCells;
var
  NamesList: TqiStrings;
  RowNumber, i: Integer;
  XMLNodeList: IXMLDOMNodeList;
const
  BeginCols = 2;
  BeginRows = 1;
  
  procedure WriteAttributes(CurrentXMLNode: IXMLDOMNode; Row: Integer);
  var
    i, Index: Integer;
    Attribute: IXMLDOMNode;
  begin
    if CurrentXMLNode.Text <> '' then
    begin
      if Length(FCells[1]) < Succ(Row) then
        SetLength(FCells[1], Succ(Row));
      FCells[1, Row] := Trim(CurrentXMLNode.Text);
    end;
    if Assigned(CurrentXMLNode.attributes) then
      for i := 0 to CurrentXMLNode.attributes.length - 1 do
      begin
        Attribute := CurrentXMLNode.attributes[i];
        if Attribute.nodeName <> '' then
        begin
          Index := NamesList.IndexOf(Attribute.nodeName);
          if Index = -1 then
          begin
            NamesList.Insert(NamesList.Count, Attribute.nodeName);
            FColCount := NamesList.Count + BeginCols;
            SetLength(FCells, FColCount);
            SetLength(FCells[FColCount - 1], Succ(Row));
            FCells[FColCount - 1, 0] := NamesList[NamesList.Count - 1];
            if Attribute.nodeValue <> null then
              FCells[FColCount - 1, Row] := Attribute.nodeValue
            else
              FCells[FColCount - 1, Row] := '';
          end else
          begin
            if Length(FCells[Index + BeginCols]) < Succ(Row) then
              SetLength(FCells[Index + BeginCols], Succ(Row));
            FCells[Index + BeginCols, Row] := Attribute.nodeValue;
          end;
        end;
      end;
  end;

  procedure WriteSubNodesText(CurrentXMLNode: IXMLDOMNode; Row: Integer);
  const
    Base64AttributeFlag = 'binary.base64';
    DatatypeAttributeFlag = 'datatype';
  var
    i, Index,
    CurrentCol: Integer;
    ChildNode,
    AttributeNode: IXMLDOMNode;
  begin
    if CurrentXMLNode.Text <> '' then
    begin
      if Length(FCells[1]) < Succ(Row) then
        SetLength(FCells[1], Succ(Row));
      FCells[1, Row] := Trim(CurrentXMLNode.Text);
    end;
    if Assigned(CurrentXMLNode.childNodes) then
      for i := 0 to CurrentXMLNode.childNodes.length - 1 do
      begin
        ChildNode := CurrentXMLNode.childNodes[i];
        if ChildNode.nodeName <> '' then
        begin
          Index := NamesList.IndexOf(ChildNode.nodeName);
          if Index = -1 then
          begin
            NamesList.Insert(NamesList.Count, ChildNode.nodeName);
            FColCount := NamesList.Count + BeginCols;
            SetLength(FCells, FColCount);
            SetLength(FCells[FColCount - 1], Succ(Row));
            FCells[FColCount - 1, 0] := NamesList[NamesList.Count - 1];
            CurrentCol := FColCount - 1;
          end else
          begin
            if Length(FCells[Index + BeginCols]) < Succ(Row) then
              SetLength(FCells[Index + BeginCols], Succ(Row));
            CurrentCol := Index + BeginCols;
          end;

          {$IFDEF VCL6}
          if Assigned(ChildNode.attributes) then
            AttributeNode := ChildNode.attributes.getNamedItem(DatatypeAttributeFlag)
          else
            AttributeNode := nil;

          if Assigned(AttributeNode) and 
            (CompareText(AttributeNode.nodeName, Base64AttributeFlag) = 0) then
            FCells[CurrentCol, Row] := 
              {$IFDEF VCL16}Soap.{$ENDIF}EncdDecd.DecodeString(ChildNode.text)
          else  
          {$ENDIF}
            FCells[CurrentCol, Row] := ChildNode.text
        end;
      end;
  end;

begin
  FColCount := BeginCols;
  FRowCount := BeginRows;
  SetLength(FCells, 0);
  
  SetLength(FCells, BeginCols);
  for i := 0 to BeginCols - 1 do
    SetLength(FCells[i], BeginRows);
  if FLoaded and (FXPath <> '') then
  begin
    XMLNodeList := FIXML.selectNodes(FXPath);
    FCells[0, 0] := 'Node name';
    FCells[1, 0] := 'Text';
    NamesList := TqiStringList.Create;
    try
      RowNumber := 1;
      for i := 0 to XMLNodeList.length - 1 do
      begin
        if (XMLNodeList.item[i].nodeType <> qntText) and (XMLNodeList.item[i].nodeType <> qntComment) then 
        begin
          SetLength(FCells[0], Succ(RowNumber));
          FCells[0, RowNumber] := XMLNodeList.item[i].nodeName;
          case FDataLocation of
            tlAttributes:
              WriteAttributes(XMLNodeList.item[i], RowNumber);
            tlSubNodes:
              WriteSubNodesText(XMLNodeList.item[i], RowNumber);
          end;
          Inc(RowNumber);
        end;
      end;
      FRowCount := RowNumber;
    finally
      NamesList.Free;
    end;
  end;
end;

procedure TXMLDocFile.Load;
begin
  Prepare;
  FIXML.load(FFileName);
  FLoaded := True;
end;

{ TQImport3XMLDoc }

constructor TQImport3XMLDoc.Create(AOwner: TComponent);
begin
  inherited;
  SkipFirstRows := 0;
end;

procedure TQImport3XMLDoc.BeforeImport;
begin
  FXMLFile := TXMLDocFile.Create;
  FXMLFile.FileName := FileName;
  FXMLFile.Load;
  FXMLFile.DataLocation := FDataLocation;
  FXMLFile.XPath := FXPath;
  FTotalRecCount := 0;
  if FXMLFile.RowCount > 0 then
    FTotalRecCount := FXMLFile.RowCount - SkipFirstRows;
  inherited;
end;

procedure TQImport3XMLDoc.StartImport;
begin
  FCounter := 0;
end;

function TQImport3XMLDoc.CheckCondition: Boolean;
begin
  Result := FCounter < (FTotalRecCount + SkipFirstRows);
end;

function TQImport3XMLDoc.Skip: Boolean;
begin
  Result := (SkipFirstRows > 0) and (FCounter < SkipFirstRows);
end;

procedure TQImport3XMLDoc.ChangeCondition;
begin
  Inc(FCounter);
end;

procedure TQImport3XMLDoc.FinishImport;
begin
  if not Canceled and not IsCSV then
  begin
    if CommitAfterDone then
      DoNeedCommit
    else if (CommitRecCount > 0) and ((ImportedRecs + ErrorRecs) mod CommitRecCount > 0) then
      DoNeedCommit;
  end;
end;

procedure TQImport3XMLDoc.AfterImport;
begin
  FXMLFile.Free;
  FXMLFile := nil;
  inherited;
end;

procedure TQImport3XMLDoc.FillImportRow;
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
{$IFDEF VCL7}
      mapValue := Map.ValueFromIndex[k];
{$ELSE}
      mapValue := Map.Values[FImportRow[j].Name];
{$ENDIF}
      strValue := FXMLFile.Cells[Pred(StrToInt(mapValue)), FCounter];
      if AutoTrimValue then
        strValue := Trim(strValue);
      RowIsEmpty := RowIsEmpty and (strValue = '');
      FImportRow.SetValue(Map.Names[k], strValue, False);
    end;
    DoUserDataFormat(FImportRow[j]);
  end;
end;

function TQImport3XMLDoc.ImportData: TQImportResult;
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

procedure TQImport3XMLDoc.DoLoadConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    SkipFirstRows := ReadInteger(XML_OPTIONS, XML_SKIP_LINES, SkipFirstRows);
  end;
end;

procedure TQImport3XMLDoc.DoSaveConfiguration(IniFile: TIniFile);
begin
  inherited;
  with IniFile do
  begin
    WriteInteger(XML_OPTIONS, XML_SKIP_LINES, SkipFirstRows);
  end;
end;

{$ENDIF}
{$ENDIF}

end.
