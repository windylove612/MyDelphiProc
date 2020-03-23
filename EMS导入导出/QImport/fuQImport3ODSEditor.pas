unit fuQImport3ODSEditor;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF ODS}
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
    Vcl.ComCtrls,
    Vcl.Buttons,
    Vcl.StdCtrls,
    Vcl.ImgList,
    Vcl.ToolWin,
    Vcl.Grids,
    Vcl.ExtCtrls,
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
    ComCtrls,
    Buttons,
    StdCtrls,
    ImgList,
    ToolWin,
    Grids,
    ExtCtrls,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3WideStringGrid,
  {$ENDIF}
  QImport3ODS,
  QImport3Common,
  QImport3StrTypes,
  QImport3BaseDocumentFile,
  QImport3XLSUtils;

type
  TfmQImport3ODSEditor = class(TForm)
    pcODSFile: TPageControl;
    lvFields: TListView;
    odFileName: TOpenDialog;
    ToolBar: TToolBar;
    tbtAutoFill: TToolButton;
    tbtClear: TToolButton;
    ilODS: TImageList;
    edODSSkipRows: TEdit;
    laSkipRows: TLabel;
    Panel1: TPanel;
    laFileName: TLabel;
    edFileName: TEdit;
    bBrowse: TSpeedButton;
    Bevel1: TBevel;
    Panel2: TPanel;
    bOk: TButton;
    bCancel: TButton;
    Bevel2: TBevel;
    procedure bBrowseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tbtAutoFillClick(Sender: TObject);
    procedure pcODSFileChange(Sender: TObject);
    procedure lvFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure tbtClearClick(Sender: TObject);
    procedure edODSSkipRowsChange(Sender: TObject);
    procedure edFileNameChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FSkipLines: Integer;
    FFileName: string;
    FImport: TQImport3ODS;
    FODSFile: TODSFile;
    FSheetName: AnsiString;
    FCurrentStringGrid: TqiStringGrid;

    procedure ApplyChanges;
    procedure ClearDataSheets;
    procedure FillFieldList;
    procedure FillGrid;
    procedure FillMap;
    procedure SetFileName(const Value: string);
    procedure SetSheetName(const Value: AnsiString);
    procedure SetSkipLines(const Value: Integer);
    procedure sgrODSMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure sgrODSDrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure TuneButtons;

    function ODSCol: Integer;
  public
    property Import: TQImport3ODS read FImport write FImport;
    property FileName: String read FFileName write SetFileName;
    property SkipLines: Integer read FSkipLines write SetSkipLines;
    property SheetName: AnsiString read FSheetName write SetSheetName;
  end;

var
  fmQImport3ODSEditor: TfmQImport3ODSEditor;

function RunQImportODSEditor(AImport: TQImport3ODS): boolean;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF ODS}
{$IFDEF VCL6}

uses
  fuQImport3Loading;

{$R *.dfm}

function RunQImportODSEditor(AImport: TQImport3ODS): boolean;
begin
  with TfmQImport3ODSEditor.Create(nil) do
  try
    Import := AImport;
    FileName := AImport.FileName;
    SkipLines := AImport.SkipFirstRows;
    FillFieldList;
    FillMap;
    SheetName := AImport.SheetName;
    TuneButtons;
    Result := ShowModal = mrOk;
    if Result then ApplyChanges;
  finally
    Free;
  end;
end;

{ TfmQImport3ODSEditor }

procedure TfmQImport3ODSEditor.FillFieldList;
var
  i: Integer;

  procedure ClearFieldList;
  begin
     lvFields.Items.Clear;
  end;

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

procedure TfmQImport3ODSEditor.SetFileName(const Value: String);
begin
  if FFileName <> Value then
  begin
    FFileName := Value;
    edFileName.Text := FFileName;
    FillFieldList;
    FillGrid;
    FillMap;
   end;
  TuneButtons;
end;

procedure TfmQImport3ODSEditor.SetSkipLines(const Value: Integer);
begin
  if FSkipLines <> Value then
  begin
    FSkipLines := Value;
    edODSSkipRows.Text := IntToStr(FSkipLines);
  end;
  if (edODSSkipRows.Text = EmptyStr) then
    edODSSkipRows.Text := IntToStr(FSkipLines);
end;

procedure TfmQImport3ODSEditor.bBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FileName;
  if odFileName.Execute then
    FileName := odFileName.FileName;
end;

procedure TfmQImport3ODSEditor.ClearDataSheets;
var
  i: integer;
begin
  for i := pcODSFile.PageCount - 1 downto 0 do
    pcODSFile.Pages[i].Free;
end;

procedure TfmQImport3ODSEditor.FillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k, i, j: integer;
  F: TForm;
  Start, Finish: TDateTime;
  W: integer;
begin
  FODSFile.Free;
  FODSFile := TODSFile.Create;
  if assigned(FCurrentStringGrid) then
    FCurrentStringGrid := nil;
  ClearDataSheets;
  SheetName := '';
  if not FileExists(FFileName) then Exit;
  FODSFile.FileName := FFileName;
  Start := Now;
  F := ShowLoading(Self, FFileName);
  FODSFile.Workbook.IsNotExpanding := Import.NotExpandMergedValue;
  try
    Application.ProcessMessages;
    FODSFile.Load;
    for k := 0 to FODSFile.Workbook.SpreadSheets.Count - 1 do
    begin
      TabSheet := TTabSheet.Create(pcODSFile);
      TabSheet.PageControl := pcODSFile;
      TabSheet.Caption := string(FODSFile.Workbook.SpreadSheets[k].Name);
      TabSheet.Tag := k;

      StringGrid := TqiStringGrid.Create(TabSheet);
      StringGrid.Parent := TabSheet;
      StringGrid.Align := alClient;
      StringGrid.ColCount := 257;
      StringGrid.RowCount := 257;
      StringGrid.FixedCols := 1;
      StringGrid.FixedRows := 1;
      StringGrid.Tag := k;
      StringGrid.DefaultColWidth := 64;
      StringGrid.DefaultRowHeight := 16;
      StringGrid.ColWidths[0] := 30;
      StringGrid.Options := StringGrid.Options - [goRangeSelect];
      GridFillFixedCells(StringGrid);
      StringGrid.OnDrawCell := sgrODSDrawCell;
      StringGrid.OnMouseDown := sgrODSMouseDown;

      for i := 0 to FODSFile.Workbook.SpreadSheets[k].DataCells.RowCount - 1 do
        for j := 0 to FODSFile.Workbook.SpreadSheets[k].DataCells.ColCount - 1 do
        begin
        StringGrid.Cells[j + 1, i + 1] :=
          FODSFile.Workbook.SpreadSheets[k].DataCells.Cells[j, i];
        W := StringGrid.Canvas.TextWidth(StringGrid.Cells[j + 1, i + 1]);
        if W + 10 > StringGrid.ColWidths[j + 1] then
          if W + 10 < 130 then
            StringGrid.ColWidths[j + 1] := W + 10
          else
            StringGrid.ColWidths[j + 1] := 130;
        end;
       if (k = 0) and (SheetName = '') then
         SheetName := AnsiString(TabSheet.Caption);
    end;
  finally
    Finish := Now;
    while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
      Finish := Now;
    if Assigned(F) then
      F.Free;
    TuneButtons;      
  end;
end;

procedure TfmQImport3ODSEditor.FormCreate(Sender: TObject);
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k: Integer;
begin
  FODSFile := TODSFile.Create;
  for k := 0 to 2 do
  begin
    TabSheet := TTabSheet.Create(pcODSFile);
    TabSheet.PageControl := pcODSFile;
    TabSheet.Caption := 'Table' + IntToStr(k + 1);
    TabSheet.Tag := k;

    StringGrid := TqiStringGrid.Create(TabSheet);
    StringGrid.Parent := TabSheet;
    StringGrid.Align := alClient;
    StringGrid.ColCount := 257;
    StringGrid.RowCount := 257;
    StringGrid.FixedCols := 1;
    StringGrid.FixedRows := 1;
    StringGrid.Tag := k;
    StringGrid.DefaultColWidth := 64;
    StringGrid.DefaultRowHeight := 16;
    StringGrid.ColWidths[0] := 30;
    StringGrid.Options := StringGrid.Options - [goRangeSelect];
    GridFillFixedCells(StringGrid);
  end;
end;

procedure TfmQImport3ODSEditor.FormDestroy(Sender: TObject);
begin
  if assigned(FODSFile) then
    FODSFile.Free;
  SheetName := '';
  if assigned(FCurrentStringGrid) then
    FCurrentStringGrid := nil;
end;

procedure TfmQImport3ODSEditor.SetSheetName(const Value: AnsiString);
var
  i: Integer;
begin
  if FSheetName <> Value then
  begin
    FSheetName := Value;
    for i := 0 to pcODSFile.PageCount - 1 do
      if pcODSFile.Pages[i].Caption = string(Value) then
      begin
        pcODSFile.ActivePage := pcODSFile.Pages[i];
        Break;
      end;
    if assigned(pcODSFile.ActivePage) then
      for i := 0 to pcODSFile.ActivePage.ComponentCount - 1 do
        if pcODSFile.ActivePage.Components[i].Tag = pcODSFile.ActivePage.Tag then
          FCurrentStringGrid := TqiStringGrid(pcODSFile.ActivePage.Components[i]);
   end;
end;

procedure TfmQImport3ODSEditor.sgrODSMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  grid.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvFields.Selected) then Exit;
  if ACol = 0 then Exit;
  if ODSCol = ACol
    then lvFields.Selected.SubItems[0] := EmptyStr
    else lvFields.Selected.SubItems[0] := Col2Letter(ACol);
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

function TfmQImport3ODSEditor.ODSCol: Integer;
begin
  Result := 0;
  if Assigned(lvFields.Selected) then
    if lvFields.Selected.SubItems[0] <> EmptyStr then
      Result := GetColIdFromColIndex(lvFields.Selected.SubItems[0]);
end;

procedure TfmQImport3ODSEditor.sgrODSDrawCell(Sender: TObject; ACol,
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
    if (ACol = ODSCol) and (ARow = 0)then
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
    grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, grid.Cells[ACol, ARow]);

    if (ACol = ODSCol) and (ARow > 0) then
    begin
      grid.Canvas.Font.Color := clHighLightText;
      grid.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      grid.Canvas.FillRect(Rect);
      grid.Canvas.TextOut(Rect.Left + 2, Y, grid.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State
    then DrawFocusRect(grid.Canvas.Handle, Rect);
  grid.DefaultDrawing := true;
end;

procedure TfmQImport3ODSEditor.FillMap;
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

procedure TfmQImport3ODSEditor.ApplyChanges;
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
  Import.SheetName := AnsiString(SheetName);
  Import.SkipFirstRows := SkipLines;
end;

procedure TfmQImport3ODSEditor.tbtAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  if Assigned(FCurrentStringGrid) then
  begin
    for i := 0 to lvFields.Items.Count - 1 do
      if (i <= FCurrentStringGrid.ColCount - 2) then
        lvFields.Items[i].SubItems[0] := Col2Letter(i + 1)
      else
        lvFields.Items[i].SubItems[0] := EmptyStr;
    lvFieldsChange(lvFields, lvFields.Selected, ctState);
  end;
end;

procedure TfmQImport3ODSEditor.pcODSFileChange(Sender: TObject);
begin
  SheetName := AnsiString(pcODSFile.ActivePage.Caption);
end;

procedure TfmQImport3ODSEditor.lvFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  if Assigned(FCurrentStringGrid) then
    FCurrentStringGrid.Repaint;
end;

procedure TfmQImport3ODSEditor.tbtClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
    lvFields.Items[i].SubItems[0] := EmptyStr;
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3ODSEditor.TuneButtons;
var
  Condition: boolean;
begin
  Condition := (lvFields.Items.Count > 0) and FileExists(FileName)
    and (pcODSFile.PageCount > 0);
  tbtAutoFill.Enabled := Condition;
  tbtClear.Enabled := Condition;
  laSkipRows.Enabled := FileExists(FileName) and (pcODSFile.PageCount > 0);
  edODSSkipRows.Enabled := FileExists(FileName) and (pcODSFile.PageCount > 0);
  pcODSFile.Enabled := FileExists(FileName) and (pcODSFile.PageCount > 0);
end;

procedure TfmQImport3ODSEditor.edODSSkipRowsChange(Sender: TObject);
begin
  SkipLines := StrToIntDef(edODSSkipRows.Text, 0);
end;

procedure TfmQImport3ODSEditor.edFileNameChange(Sender: TObject);
begin
  if FFileName <> edFileName.Text then
  begin
    FFileName := edFileName.Text;
    FillGrid;
  end;
  TuneButtons;
end;

procedure TfmQImport3ODSEditor.FormShow(Sender: TObject);
begin
  Caption := Import.Name + ' - Component Editor';
end;

{$ENDIF}
{$ENDIF}

end.
