unit fuQImport3XlsxEditor;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF XLSX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    Winapi.Windows,
    Winapi.Messages,
    System.SysUtils,
    System.Variants,
    System.Classes,
    Vcl.Graphics,
    Vcl.Controls,
    Vcl.Forms,
    Vcl.Dialogs,
    Vcl.Buttons,
    Vcl.ExtCtrls,
    Vcl.StdCtrls,
    Vcl.ComCtrls,
    Vcl.ToolWin,
    Vcl.ImgList,
    Vcl.Grids,
    System.Math,
  {$ELSE}
    Windows,
    Messages,
    SysUtils,
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    Buttons,
    ExtCtrls,
    StdCtrls,
    ComCtrls,
    ToolWin,
    ImgList,
    Grids,
    Math,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3WideStringGrid,
  {$ENDIF}
  QImport3StrTypes,
  QImport3Xlsx;

type
  TfmQImport3XlsxEditor = class(TForm)
    laXlsxSkipRows_01: TLabel;
    laXlsxSkipRows_02: TLabel;
    edXlsxSkipRows: TEdit;
    pnlFileName: TPanel;
    bvlBrowse: TBevel;
    laFileName: TLabel;
    btnBrowse: TSpeedButton;
    edFileName: TEdit;
    bvl1: TBevel;
    pnlButtons: TPanel;
    buOk: TButton;
    buCancel: TButton;
    paXlsxFields: TPanel;
    lvXlsxFields: TListView;
    pcXlsxFile: TPageControl;
    tlbXlsxUtils: TToolBar;
    tlbXlsxAutoFill: TToolButton;
    tlbXlsxClear: TToolButton;
    ilXlsx: TImageList;
    odFileName: TOpenDialog;
    procedure FormCreate(Sender: TObject);
    procedure btnBrowseClick(Sender: TObject);
    procedure sgrXlsxDrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure sgrXlsxMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure lvXlsxFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure pcXlsxFileChange(Sender: TObject);
    procedure tlbXlsxAutoFillClick(Sender: TObject);
    procedure tlbXlsxClearClick(Sender: TObject);
    procedure edXlsxSkipRowsChange(Sender: TObject);
    procedure edFileNameChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FNeedLoadFile: Boolean;
    FSkipLines: Integer;
    FFileName: string;
    FImport: TQImport3Xlsx;
    FSheetName: string;
    FCurrentStringGrid: TqiStringGrid;
    procedure SetFileName(const Value: string);
    procedure SetSkipLines(const Value: Integer);
    procedure SetSheetName(const Value: string);

    procedure ClearAll;
    procedure FillFieldList;
    procedure FillGrids;
    procedure FillMap;
    procedure ClearMap;
    procedure TuneButtons;
    procedure ApplyChanges;

    function XlsxCol: Integer;
  public
    property Import: TQImport3Xlsx read FImport write FImport;
    property FileName: string read FFileName write SetFileName;
    property SkipLines: Integer read FSkipLines write SetSkipLines;
    property SheetName: string read FSheetName write SetSheetName;
  end;

var
  fmQImport3XlsxEditor: TfmQImport3XlsxEditor;

function RunQImportXlsxEditor(AImport: TQImport3Xlsx): boolean;

{$ENDIF}
{$ENDIF}


implementation

{$IFDEF XLSX}
{$IFDEF VCL6}

uses
  QImport3Common, fuQImport3Loading, QImport3XLSUtils;

{$R *.dfm}

function RunQImportXlsxEditor(AImport: TQImport3Xlsx): boolean;
begin
  with TfmQImport3XlsxEditor.Create(nil) do
  try
    Import := AImport;
    FileName := AImport.FileName;
    SkipLines := AImport.SkipFirstRows;
    FillFieldList;
    SheetName := string(AImport.SheetName);
    FillMap;
    TuneButtons;

    Result := ShowModal = mrOk;
    if Result then ApplyChanges;
  finally
    Free;
  end;
end;

procedure TfmQImport3XlsxEditor.FormCreate(Sender: TObject);
begin
  FNeedLoadFile := True;
end;

procedure TfmQImport3XlsxEditor.SetFileName(const Value: string);
begin
  if FFileName <> Value then
  begin
    FFileName := Value;
    edFileName.Text := FFileName;
    FNeedLoadFile := True;
    FillGrids;
  end;
  TuneButtons;
end;

procedure TfmQImport3XlsxEditor.SetSkipLines(const Value: Integer);
begin
  if FSkipLines <> Value then
  begin
    FSkipLines := Value;
    edXlsxSkipRows.Text := IntToStr(FSkipLines);
  end;
end;

procedure TfmQImport3XlsxEditor.SetSheetName(const Value: string);
var
  i: Integer;
begin
  for i := 0 to pcXlsxFile.PageCount - 1 do
    if pcXlsxFile.Pages[i].Caption = Value then
    begin
      pcXlsxFile.ActivePage := pcXlsxFile.Pages[i];
      Break;
    end;
  if Assigned(pcXlsxFile.ActivePage) then
  begin
    FSheetName := pcXlsxFile.ActivePage.Caption;
    if pcXlsxFile.ActivePage.Components[0] is TqiStringGrid then
      FCurrentStringGrid := TqiStringGrid(pcXlsxFile.ActivePage.Components[0]);
  end;
end;

procedure TfmQImport3XlsxEditor.ClearAll;
begin
  SheetName := '';
  while pcXlsxFile.PageCount > 0 do
    pcXlsxFile.Pages[0].Free;
end;

procedure TfmQImport3XlsxEditor.FillFieldList;
var
  i: Integer;

  procedure ClearFieldList;
  begin
     lvXlsxFields.Items.Clear;
  end;

begin
  if not QImportDestinationAssigned(False, FImport.ImportDestination,
      FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid)
    then Exit;
  ClearFieldList;
  for i := 0 to QImportDestinationColCount(False, FImport.ImportDestination,
    FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid) - 1 do
    with lvXlsxFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, FImport.ImportDestination,
        FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid,
        FImport.GridCaptionRow, i);
      SubItems.Add(EmptyStr);
      ImageIndex := 0;
    end;

    if lvXlsxFields.Items.Count > 0 then
    begin
      lvXlsxFields.Items[0].Focused := True;
      lvXlsxFields.Items[0].Selected := True;
    end;
end;

procedure TfmQImport3XlsxEditor.FillGrids;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  Start, Finish: TDateTime;
  XlsxFile: TXlsxFile;
  F: TForm;
  i, j: Integer;
begin
  ClearAll;
  if not FileExists(FileName) then Exit;
  XlsxFile := TXlsxFile.Create;
  try
    XlsxFile.FileName := FFileName;
    XlsxFile.Workbook.NeedFillMerge := Import.NeedFillMerge;
    XlsxFile.Workbook.LoadHiddenSheets := Import.LoadHiddenSheet;
    Start := Now;
    F := ShowLoading(Self, FFileName);
    try
      Application.ProcessMessages;
      XlsxFile.Load;
      for i := 0 to XlsxFile.Workbook.WorkSheets.Count - 1 do
      begin
        TabSheet := TTabSheet.Create(pcXlsxFile);
        TabSheet.PageControl := pcXlsxFile;
        TabSheet.Caption := XlsxFile.Workbook.WorkSheets[i].Name;

        StringGrid := TqiStringGrid.Create(TabSheet);
        StringGrid.Parent := TabSheet;
        StringGrid.Align := alClient;
        StringGrid.ColCount := 257;
        StringGrid.RowCount := 257;
        StringGrid.FixedCols := 1;
        StringGrid.FixedRows := 1;
        StringGrid.DefaultColWidth := 64;
        StringGrid.DefaultRowHeight := 16;
        StringGrid.ColWidths[0] := 30;
        StringGrid.Tag := i;
        StringGrid.Options := StringGrid.Options - [goRangeSelect];
        GridFillFixedCells(StringGrid);
        StringGrid.OnDrawCell := sgrXlsxDrawCell;
        StringGrid.OnMouseDown := sgrXlsxMouseDown;
        for j := 0 to XlsxFile.Workbook.WorkSheets[i].Cells.Count - 1 do
          StringGrid.Cells[XlsxFile.Workbook.WorkSheets[i].Cells[j].Col,
            XlsxFile.Workbook.WorkSheets[i].Cells[j].Row] := XlsxFile.Workbook.WorkSheets[i].Cells[j].Value;
        if (i = 0) then
          SheetName := TabSheet.Caption;
      end;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  finally
    XlsxFile.Free;
  end;
end;

procedure TfmQImport3XlsxEditor.btnBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FFileName;
  if odFileName.Execute then
  begin
    FileName := odFileName.FileName;
    ClearMap;
  end;
end;

procedure TfmQImport3XlsxEditor.sgrXlsxDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  X := Rect.Left + (Rect.Right - Rect.Left -
   grid.Canvas.TextWidth(grid.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    grid.Canvas.TextHeight(grid.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    if (ACol = XlsxCol) and (ARow = 0) then
      grid.Canvas.Font.Style := grid.Canvas.Font.Style + [fsBold]
    else
      grid.Canvas.Font.Style := grid.Canvas.Font.Style - [fsBold];
    grid.Canvas.FillRect(Rect);
    grid.Canvas.TextOut(X - 1, Y + 1, grid.Cells[ACol, ARow]);
  end
  else begin
    grid.DefaultDrawing := false;
    grid.Canvas.Brush.Color := clWindow;
    grid.Canvas.FillRect(Rect);
    grid.Canvas.Font.Color := clWindowText;
    grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      grid.Cells[ACol, ARow]);

    if (ACol = XlsxCol) and (ARow > 0) then
    begin
      grid.Canvas.Font.Color := clHighLightText;
      grid.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      grid.Canvas.FillRect(Rect);
      grid.Canvas.TextOut(Rect.Left + 2, Y, grid.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State then
    DrawFocusRect(grid.Canvas.Handle, Rect);
  grid.DefaultDrawing := true;
end;

procedure TfmQImport3XlsxEditor.sgrXlsxMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  grid.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvXlsxFields.Selected) then Exit;
  if XlsxCol = ACol
    then lvXlsxFields.Selected.SubItems[0] := EmptyStr
//    else lvXlsxFields.Selected.SubItems[0] := IntToStr(ACol + 1);
    else lvXlsxFields.Selected.SubItems[0] := Col2Letter(ACol);
  lvXlsxFieldsChange(lvXlsxFields, lvXlsxFields.Selected, ctState);
end;

procedure TfmQImport3XlsxEditor.lvXlsxFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  if Assigned(FCurrentStringGrid) then
    FCurrentStringGrid.Repaint;
end;

procedure TfmQImport3XlsxEditor.pcXlsxFileChange(Sender: TObject);
begin
  SheetName := pcXlsxFile.ActivePage.Caption;
end;

procedure TfmQImport3XlsxEditor.tlbXlsxAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  if Assigned(FCurrentStringGrid) then
  begin
    for i := 0 to lvXlsxFields.Items.Count - 1 do
      if (i <= FCurrentStringGrid.ColCount - 2) then
        lvXlsxFields.Items[i].SubItems[0] := Col2Letter(i + 1)
      else
        lvXlsxFields.Items[i].SubItems[0] := EmptyStr;
    lvXlsxFieldsChange(lvXlsxFields, lvXlsxFields.Selected, ctState);
  end;
end;

procedure TfmQImport3XlsxEditor.tlbXlsxClearClick(Sender: TObject);
begin
  ClearMap;
end;

procedure TfmQImport3XlsxEditor.edXlsxSkipRowsChange(Sender: TObject);
begin
  SkipLines := StrToIntDef(edXlsxSkipRows.Text, 0);
end;

procedure TfmQImport3XlsxEditor.FillMap;
var
  i, j: Integer;
begin
  for i := 0 to Import.Map.Count - 1 do
    for j := 0 to lvXlsxFields.Items.Count - 1 do
      if AnsiCompareText(lvXlsxFields.Items[j].Caption, Import.Map.Names[i]) = 0 then
      begin
        lvXlsxFields.Items[j].SubItems[0] := Import.Map.Values[Import.Map.Names[i]];
        Break;
      end;
end;

procedure TfmQImport3XlsxEditor.ClearMap;
var
  i: Integer;
begin
  for i := 0 to lvXlsxFields.Items.Count - 1 do
    lvXlsxFields.Items[i].SubItems[0] := EmptyStr;
  lvXlsxFieldsChange(lvXlsxFields, lvXlsxFields.Selected, ctState);
end;

procedure TfmQImport3XlsxEditor.TuneButtons;
var
  Condition: boolean;
begin
  Condition := (lvXlsxFields.Items.Count > 0) and FileExists(FileName) and
    (pcXlsxFile.PageCount > 0);
  tlbXlsxAutoFill.Enabled := Condition;
  tlbXlsxClear.Enabled := Condition;

  laXlsxSkipRows_01.Enabled := Condition;
  edXlsxSkipRows.Enabled := Condition;
  laXlsxSkipRows_02.Enabled := Condition;
end;

procedure TfmQImport3XlsxEditor.ApplyChanges;
var
  i: Integer;
begin
  Import.Map.BeginUpdate;
  try
    Import.Map.Clear;
    for i := 0 to lvXlsxFields.Items.Count - 1 do
      if lvXlsxFields.Items[i].SubItems[0] <> EmptyStr then
        Import.Map.Values[lvXlsxFields.Items[i].Caption] :=
          lvXlsxFields.Items[i].SubItems[0];
  finally
    Import.Map.EndUpdate;
  end;
  Import.FileName := FileName;
  Import.SheetName := SheetName;
  Import.SkipFirstRows := SkipLines;
end;

function TfmQImport3XlsxEditor.XlsxCol: Integer;
begin
  Result := 0;
  if Assigned(lvXlsxFields.Selected) then
    if lvXlsxFields.Selected.SubItems[0] <> EmptyStr then
      Result := GetColIdFromColIndex(lvXlsxFields.Selected.SubItems[0]);
end;

procedure TfmQImport3XlsxEditor.edFileNameChange(Sender: TObject);
begin
  if FFileName <> edFileName.Text then
  begin
    FFileName := edFileName.Text;
    FNeedLoadFile := True;
    FillGrids;
  end;
  TuneButtons;
end;

procedure TfmQImport3XlsxEditor.FormShow(Sender: TObject);
begin
  Caption := Import.Name + ' - Component Editor';
end;

{$ENDIF}
{$ENDIF}

end.
