unit fuQImport3XMLDocEditor;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF XMLDOC}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Variants,
    Winapi.Windows,
    Winapi.Messages,
    System.SysUtils,
    System.Classes,
    Vcl.Graphics,
    Vcl.Controls,
    Vcl.Forms,
    Vcl.Dialogs,
    Vcl.ComCtrls,
    Vcl.ImgList,
    Vcl.StdCtrls,
    Vcl.ToolWin,
    Vcl.Grids,
    Vcl.Buttons,
    Vcl.ExtCtrls,
    Vcl.Menus,
  {$ELSE}
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    ComCtrls,
    ImgList,
    StdCtrls,
    ToolWin,
    Grids,
    Buttons,
    ExtCtrls,
    Menus,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3WideStringGrid,
  {$ENDIF}
  QImport3XMLDoc,
  QImport3StrTypes;

type
  TfmQImport3XMLDocEditor = class(TForm)
    laSkip_01: TLabel;
    laSkip_02: TLabel;
    laFileName: TLabel;
    Bevel2: TBevel;
    bBrowse: TSpeedButton;
    laColumn: TLabel;
    lvFields: TListView;
    edSkip: TEdit;
    edFileName: TEdit;
    bOk: TButton;
    bCancel: TButton;
    ToolBar: TToolBar;
    tbtAutoFill: TToolButton;
    tbtClear: TToolButton;
    cbColumn: TComboBox;
    odFileName: TOpenDialog;
    imField: TImageList;
    tvXMLDoc: TTreeView;
    laXPath: TLabel;
    edXPath: TEdit;
    Bevel1: TBevel;
    bFillGrig: TSpeedButton;
    bBuildTree: TSpeedButton;
    bGetXPath: TSpeedButton;
    pmTreeView: TPopupMenu;
    miGetXPath: TMenuItem;
    laDataLocation: TLabel;
    cbDataLocation: TComboBox;
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure bBrowseClick(Sender: TObject);
    procedure sgrXMLDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure sgrXMLMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure lvFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure cbColumnChange(Sender: TObject);
    procedure tbtAutoFillClick(Sender: TObject);
    procedure tbtClearClick(Sender: TObject);
    procedure bFillGrigClick(Sender: TObject);
    procedure edSkipChange(Sender: TObject);
    procedure bBuildTreeClick(Sender: TObject);
    procedure bGetXPathClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FImport: TQImport3XMLDoc;
    FNeedLoadFile: Boolean;
    FFileName: string;
    FSkipLines: Integer;
    FXMLFile: TXMLDocFile;
    FXPath: qiString;
    FDataLocation: TXMLDataLocation;

    procedure SetFileName(const Value: string);
    procedure SetSkipLines(const Value: Integer);
    procedure SetXPath(const Value: qiString);
    procedure SetDataLocation(const Value: TXMLDataLocation);

    procedure FillFieldList;
    procedure FillGrid;
    procedure ClearGrid;
    procedure FillComboColumn;
    procedure FillMap;
    procedure TuneButtons;
    procedure ApplyChanges;
    function XMLCol: Integer;
    function GetXPath: qiString;
  public
    sgrXML: TqiStringGrid;
    property Import: TQImport3XMLDoc read FImport write FImport;
    property FileName: string read FFileName write SetFileName;
    property XPath: qiString read FXPath write SetXPath;
    property DataLocation: TXMLDataLocation read FDataLocation
      write SetDataLocation;
    property SkipLines: Integer read FSkipLines write SetSkipLines;
  end;

var
  fmQImport3XMLDocEditor: TfmQImport3XMLDocEditor;

function RunQImportXMLDocEditor(AImport: TQImport3XMLDoc): boolean;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF XMLDOC}
{$IFDEF VCL6}

uses QImport3Common, fuQImport3Loading;

{$R *.dfm}

function RunQImportXMLDocEditor(AImport: TQImport3XMLDoc): boolean;
begin
  with TfmQImport3XMLDocEditor.Create(nil) do
  try
    Import := AImport;
    FileName := AImport.FileName;
    DataLocation := AImport.DataLocation;
    XPath := AImport.XPath;
    SkipLines := AImport.SkipFirstRows;
    FillFieldList;
    FillMap;
    TuneButtons;

    Result := ShowModal = mrOk;
    if Result then ApplyChanges;
  finally
    Free;
  end;
end;

{ TfmQImport3XMLDocEditor }

procedure TfmQImport3XMLDocEditor.FormCreate(Sender: TObject);
begin
  sgrXML := TqiStringGrid.Create(Self);
  sgrXML.Parent := Self;
  sgrXML.Left := 202;
  sgrXML.Top := 120;
  sgrXML.Width := 439;
  sgrXML.Height := 285;
  sgrXML.ColCount := 2;
  sgrXML.DefaultColWidth := 82;
  sgrXML.DefaultRowHeight := 16;
  sgrXML.RowCount := 2;
  sgrXML.Font.Charset := DEFAULT_CHARSET;
  sgrXML.Font.Color := clWindowText;
  sgrXML.Font.Height := -11;
  sgrXML.Font.Name := 'Courier New';
  sgrXML.Font.Style := [];
  sgrXML.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine];
  sgrXML.ParentFont := False;
  sgrXML.TabOrder := 7;
  sgrXML.OnDrawCell := sgrXMLDrawCell;
  sgrXML.OnMouseDown := sgrXMLMouseDown;

  FNeedLoadFile := True;
  FXMLFile := TXMLDocFile.Create;
end;

procedure TfmQImport3XMLDocEditor.FormDestroy(Sender: TObject);
begin
  FXMLFile.Destroy;
end;

procedure TfmQImport3XMLDocEditor.bBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FileName;
  if odFileName.Execute then FileName := odFileName.FileName;
end;

procedure TfmQImport3XMLDocEditor.sgrXMLDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: integer;
begin
  X := Rect.Left + (Rect.Right - Rect.Left -
    sgrXML.Canvas.TextWidth(sgrXML.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    sgrXML.Canvas.TextHeight(sgrXML.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    if (ACol = XMLCol - 1) and (ARow = 0)
      then sgrXML.Canvas.Font.Style := sgrXML.Canvas.Font.Style + [fsBold]
      else sgrXML.Canvas.Font.Style := sgrXML.Canvas.Font.Style - [fsBold];
    sgrXML.Canvas.FillRect(Rect);
    sgrXML.Canvas.TextOut(X - 1, Y + 1, sgrXML.Cells[ACol, ARow]);
  end
  else begin
    sgrXML.DefaultDrawing := false;
    sgrXML.Canvas.Brush.Color := clWindow;
    sgrXML.Canvas.FillRect(Rect);
    sgrXML.Canvas.Font.Color := clWindowText;
    sgrXML.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      sgrXML.Cells[ACol, ARow]);

    if (ACol = XMLCol - 1) and (ARow > 0) then
    begin
      sgrXML.Canvas.Font.Color := clHighLightText;
      sgrXML.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      sgrXML.Canvas.FillRect(Rect);
      sgrXML.Canvas.TextOut(Rect.Left + 2, Y, sgrXML.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State
    then DrawFocusRect(sgrXML.Canvas.Handle, Rect);
  sgrXML.DefaultDrawing := true;
end;

procedure TfmQImport3XMLDocEditor.sgrXMLMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: integer;
begin
  sgrXML.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvFields.Selected) then Exit;
  if XMLCol = ACol + 1
    then lvFields.Selected.SubItems[0] := EmptyStr
    else lvFields.Selected.SubItems[0] := IntToStr(ACol + 1);
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3XMLDocEditor.lvFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;          
  cbColumn.OnChange := nil;
  try
    if Item.SubItems.Count > 0 then
      cbColumn.ItemIndex := cbColumn.Items.IndexOf(Item.SubItems[0]);
    sgrXML.Repaint;
  finally
    cbColumn.OnChange := cbColumnChange;
  end;
end;

procedure TfmQImport3XMLDocEditor.cbColumnChange(Sender: TObject);
begin
  if not Assigned(lvFields.Selected) then Exit;
  lvFields.Selected.SubItems[0] := cbColumn.Text;
  sgrXML.Repaint;
end;

procedure TfmQImport3XMLDocEditor.tbtAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
    if (i <= sgrXML.ColCount - 2) then
      lvFields.Items[i].SubItems[0] := IntToStr(i + 2)
    else
      lvFields.Items[i].SubItems[0] := EmptyStr;
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3XMLDocEditor.tbtClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
    lvFields.Items[i].SubItems[0] := EmptyStr;
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3XMLDocEditor.bFillGrigClick(Sender: TObject);
begin
  DataLocation := TXMLDataLocation(cbDataLocation.ItemIndex);
  XPath := edXPath.Text;
end;

procedure TfmQImport3XMLDocEditor.edSkipChange(Sender: TObject);
begin
  SkipLines := StrToIntDef(edSkip.Text, 0);
end;

procedure TfmQImport3XMLDocEditor.SetFileName(const Value: String);
var
  F: TForm;
  Start, Finish: TDateTime;
begin
  if Value <> EmptyStr then
  begin
    FNeedLoadFile := True;
    FFileName := Value;
    edFileName.Text := FFileName;
    FXMLFile.FileName := FFileName;
    if FNeedLoadFile then
    begin
      Start := Now;
      F := ShowLoading(Self, FFileName);
      try
        Application.ProcessMessages;
        FXMLFile.Load;
        FNeedLoadFile := False;
        ClearGrid;
      finally
        Finish := Now;
        while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
          Finish := Now;
        F.Free;
      end;
    end;
  end;
  TuneButtons;
end;

procedure TfmQImport3XMLDocEditor.SetSkipLines(const Value: Integer);
begin
  if FSkipLines <> Value then
  begin
    FSkipLines := Value;
    edSkip.Text := IntToStr(FSkipLines);
  end;
end;

procedure TfmQImport3XMLDocEditor.SetXPath(const Value: qiString);
var
  F: TForm;
  Start, Finish: TDateTime;
begin
  if Value <> EmptyStr then
  begin
    FXPath := Value;
    edXPath.Text := Value;
    Start := Now;
    F := ShowLoading(Self, Value);
    try
      Application.ProcessMessages;
      FXMLFile.XPath := Value;
     finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      F.Free;
    end;
    FillGrid;
  end;
end;

procedure TfmQImport3XMLDocEditor.SetDataLocation(
  const Value: TXMLDataLocation);
begin
  if Value <> FDataLocation then
  begin
    FDataLocation := Value;
    cbDataLocation.ItemIndex := Integer(Value);
    FXMLFile.DataLocation := Value;
  end;
end;

procedure TfmQImport3XMLDocEditor.FillFieldList;

  procedure ClearFieldList;
  begin
     lvFields.Items.Clear;
  end;

var
  i: Integer;
begin
  if not QImportDestinationAssigned(False, FImport.ImportDestination,
      FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid)
    then Exit;
  ClearFieldList;
  for i := 0 to QImportDestinationColCount(False, FImport.ImportDestination,
    FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid) - 1 do
    with lvFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, FImport.ImportDestination,
        FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid,
        FImport.GridCaptionRow, i);
      SubItems.Add(EmptyStr);
      ImageIndex := 0;
    end;

    if lvFields.Items.Count > 0 then
    begin
      lvFields.Items[0].Focused := True;
      lvFields.Items[0].Selected := True;
    end;
end;

procedure TfmQImport3XMLDocEditor.FillGrid;
begin
  ClearGrid;
  FillStringGrid(sgrXML, FXMLFile);
  FillComboColumn;
end;

procedure TfmQImport3XMLDocEditor.ClearGrid;
var
  i: Integer;
begin
  for i := 0 to sgrXML.ColCount - 1 do
    sgrXML.Cols[i].Clear;
  sgrXML.ColCount := 2;
  sgrXML.FixedCols := 1;
  sgrXML.RowCount := 2;
  sgrXML.FixedRows := 1;
end;

procedure TfmQImport3XMLDocEditor.FillComboColumn;
var
  i: Integer;
begin
  cbColumn.Clear;
  cbColumn.Items.Add('');
  cbColumn.ItemIndex := 0;
  for i := 0 to sgrXML.ColCount - 1 do
    cbColumn.Items.Add(IntToStr(Succ(i)));
end;

procedure TfmQImport3XMLDocEditor.FillMap;
var
  i, j: Integer;
begin
  for i := 0 to Import.Map.Count - 1 do
    for j := 0 to lvFields.Items.Count - 1 do
      if AnsiCompareText(lvFields.Items[j].Caption, Import.Map.Names[i]) = 0 then
      begin
        lvFields.Items[j].SubItems[0] := Import.Map.Values[Import.Map.Names[i]];
        Break;
      end;
end;

procedure TfmQImport3XMLDocEditor.TuneButtons;
var
  Condition: boolean;
begin
  Condition := (lvFields.Items.Count > 0) and FileExists(FileName);
  tbtAutoFill.Enabled := Condition;
  tbtClear.Enabled := Condition;
  sgrXML.Enabled := Condition;
  laColumn.Enabled := Condition;
  cbColumn.Enabled := Condition;
  laSkip_01.Enabled := Condition;
  edSkip.Enabled := Condition;
  laSkip_02.Enabled := Condition;
end;

procedure TfmQImport3XMLDocEditor.ApplyChanges;
var
  i: Integer;
begin
  Import.Map.BeginUpdate;
  try
    Import.Map.Clear;
    for i := 0 to lvFields.Items.Count - 1 do
      if lvFields.Items[i].SubItems[0] <> EmptyStr then
        Import.Map.Values[lvFields.Items[i].Caption] :=
          lvFields.Items[i].SubItems[0];
  finally
    Import.Map.EndUpdate;
  end;
  Import.FileName := FileName;
  Import.DataLocation := DataLocation;
  Import.XPath := XPath;
  Import.SkipFirstRows := SkipLines;
end;

function TfmQImport3XMLDocEditor.XMLCol: Integer;
begin
  Result := 0;
  if Assigned(lvFields.Selected) then
    Result := StrToIntDef(lvFields.Selected.SubItems[0], 0);
end;

procedure TfmQImport3XMLDocEditor.bBuildTreeClick(Sender: TObject);
var
  F: TForm;
  Start, Finish: TDateTime;
begin
  if not FNeedLoadFile then
  begin
    Start := Now;
    F := ShowLoading(Self, 'Build TreeView');
    try
      Application.ProcessMessages;
      XMLFile2TreeView(tvXMLDoc, FXMLFile);
     finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      F.Free;
    end;
  end;
end;

function TfmQImport3XMLDocEditor.GetXPath: qiString;
var
  CurrentNode: TTreeNode;
begin
  Result := '/';                                                                                    
  if Assigned(tvXMLDoc.Selected) then
  begin
    CurrentNode := tvXMLDoc.Selected;
    if Assigned(CurrentNode) then
      if not Assigned(CurrentNode.Data) then
        Result := Concat('/', CurrentNode.Text)
      else
      if Integer(CurrentNode.Data) = 2 then
        Result := Concat('[', CurrentNode.Text, ']');

    while CurrentNode.Parent <> nil do
    begin
      CurrentNode := CurrentNode.Parent;
      Result := Concat('/', CurrentNode.Text, Result);
    end;
  end;  
end;

procedure TfmQImport3XMLDocEditor.bGetXPathClick(Sender: TObject);
begin
  edXPath.Text := GetXPath;
end;

procedure TfmQImport3XMLDocEditor.FormShow(Sender: TObject);
begin
  Caption := Import.Name + ' - Component Editor';
end;

{$ENDIF}
{$ENDIF}

end.
