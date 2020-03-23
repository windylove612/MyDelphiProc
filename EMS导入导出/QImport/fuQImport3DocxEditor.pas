unit fuQImport3DocxEditor;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF DOCX}
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
    Vcl.ImgList,
    Vcl.ComCtrls,
    Vcl.ToolWin,
    Vcl.StdCtrls,
    Vcl.Buttons,
    Vcl.ExtCtrls,
    Vcl.Grids,
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
    ImgList,
    ComCtrls,
    ToolWin,
    StdCtrls,
    Buttons,
    ExtCtrls,
    Grids,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3WideStringGrid,
  {$ENDIF}
  QImport3StrTypes,
  QImport3Docx;

type
  TfmQImport3DocxEditor = class(TForm)
    bvl1: TBevel;
    pnlFileName: TPanel;
    bvlBrowse: TBevel;
    laFileName: TLabel;
    btnBrowse: TSpeedButton;
    edFileName: TEdit;
    pnlButtons: TPanel;
    buOk: TButton;
    buCancel: TButton;
    paDocxFields: TPanel;
    lvDocxFields: TListView;
    tlbDocxUtils: TToolBar;
    tlbDocxAutoFill: TToolButton;
    tlbDocxClear: TToolButton;
    ilDocx: TImageList;
    odFileName: TOpenDialog;
    laSkip_01: TLabel;
    laSkip_02: TLabel;
    edDocxSkip: TEdit;
    pcDocxFile: TPageControl;
    procedure btnBrowseClick(Sender: TObject);
    procedure lvDocxFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure tlbDocxAutoFillClick(Sender: TObject);
    procedure tlbDocxClearClick(Sender: TObject);
    procedure edDocxSkipChange(Sender: TObject);
    procedure sgrDocxDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure sgrDocxMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pcDocxFileChange(Sender: TObject);
    procedure edFileNameChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FImport: TQImport3Docx;
    FNeedLoadFile: Boolean;
    FFileName: String;
    FSkipLines: Integer;
    FCurrentStringGrid: TqiStringGrid;
    FTableNumber: Integer;

    procedure SetFileName(const Value: String);
    procedure SetSkipLines(const Value: Integer);
    procedure SetTableNumber(const Value: Integer);

    procedure ClearAll;
    procedure FillFieldList;
    procedure FillGrid;
    procedure FillMap;
    procedure ClearMap;
    procedure TuneButtons;
    procedure ApplyChanges;

    function DocxCol: Integer;
  public
    property Import: TQImport3Docx read FImport write FImport;
    property FileName: String read FFileName write SetFileName;
    property SkipLines: Integer read FSkipLines write SetSkipLines;
    property TableNumber: Integer read FTableNumber write SetTableNumber;
  end;

var
  fmQImport3DocxEditor: TfmQImport3DocxEditor;

function RunQImportDocxEditor(AImport: TQImport3Docx): boolean;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF DOCX}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Math,
  {$ELSE}
    Math,
  {$ENDIF}
  QImport3Common, fuQImport3Loading;

function RunQImportDocxEditor(AImport: TQImport3Docx): boolean;
begin
  with TfmQImport3DocxEditor.Create(nil) do
  try
    Import := AImport;
    FileName := AImport.FileName;
    SkipLines := AImport.SkipFirstRows;
    TableNumber := AImport.TableNumber;
    FillFieldList;
    FillMap;
    TuneButtons;

    Result := ShowModal = mrOk;
    if Result then ApplyChanges;
  finally
    Free;
  end;
end;

{$R *.dfm}

procedure TfmQImport3DocxEditor.SetFileName(const Value: String);
begin
  if FFileName <> Value then
  begin
    FFileName := Value;
    edFileName.Text := FFileName;
    FNeedLoadFile := True;
    FillGrid;
  end;
  TuneButtons;
end;

procedure TfmQImport3DocxEditor.SetSkipLines(const Value: Integer);
begin
  if FSkipLines <> Value then
  begin
    FSkipLines := Value;
    edDocxSkip.Text := IntToStr(FSkipLines);
  end;
end;

procedure TfmQImport3DocxEditor.SetTableNumber(const Value: Integer);
var
  i: Integer;
begin
  for i := 0 to pcDocxFile.PageCount - 1 do
    if Value = pcDocxFile.Pages[i].Tag then
    begin
      pcDocxFile.ActivePage := pcDocxFile.Pages[i];
      Break;
    end;
  if Assigned(pcDocxFile.ActivePage) then
  begin
    FTableNumber := pcDocxFile.ActivePage.Tag;
    if pcDocxFile.ActivePage.Components[0] is TqiStringGrid then
      FCurrentStringGrid := TqiStringGrid(pcDocxFile.ActivePage.Components[0]);
  end;
end;

procedure TfmQImport3DocxEditor.ClearAll;
begin
  if Assigned(FCurrentStringGrid) then
    FCurrentStringGrid := nil;
  while pcDocxFile.PageCount > 0 do
    pcDocxFile.Pages[0].Free;
  TableNumber := 0;
end;

procedure TfmQImport3DocxEditor.FillFieldList;
var
  i: Integer;

  procedure ClearFieldList;
  begin
     lvDocxFields.Items.Clear;
  end;

begin
  if not QImportDestinationAssigned(False, FImport.ImportDestination,
      FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid)
    then Exit;
  ClearFieldList;
  for i := 0 to QImportDestinationColCount(False, FImport.ImportDestination,
    FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid) - 1 do
    with lvDocxFields.Items.Add do
    begin
      Caption := QImportDestinationColName(False, FImport.ImportDestination,
        FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid,
        FImport.GridCaptionRow, i);
      SubItems.Add(EmptyStr);
      ImageIndex := 0;
    end;

    if lvDocxFields.Items.Count > 0 then
    begin
      lvDocxFields.Items[0].Focused := True;
      lvDocxFields.Items[0].Selected := True;
    end;
end;

procedure TfmQImport3DocxEditor.FillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  DocxFile: TDocxFile;
  F: TForm;
  Start, Finish: TDateTime;
  i, j, k:  Integer;
begin
  ClearAll;
  if not FileExists(FFileName) then Exit;
  DocxFile := TDocxFile.Create;
  try
    DocxFile.FileName := FFileName;
    DocxFile.NeedFillMerge := Import.NeedFillMerge;
    Start := Now;
    F := ShowLoading(Self, FFileName);
    try
      Application.ProcessMessages;
      DocxFile.Load;
      ClearAll;
      for i := 0 to DocxFile.Tables.Count - 1 do
      begin
        TabSheet := TTabSheet.Create(pcDocxFile);
        TabSheet.PageControl := pcDocxFile;
        TabSheet.Caption := 'Table ' + IntToStr(Succ(i));
        TabSheet.Tag := Succ(i);

        StringGrid := TqiStringGrid.Create(TabSheet);
        StringGrid.Parent := TabSheet;
        StringGrid.Align := alClient;
        StringGrid.DefaultRowHeight := 16;
        StringGrid.ColCount := 1;
        StringGrid.RowCount := 1;
        StringGrid.FixedCols := 0;
        StringGrid.FixedRows := 0;
        StringGrid.Options := StringGrid.Options - [goRangeSelect];
        StringGrid.OnDrawCell := sgrDocxDrawCell;
        StringGrid.OnMouseDown := sgrDocxMouseDown;
        
        StringGrid.ColCount := 1;
        StringGrid.RowCount := Min(DocxFile.Tables[i].Rows.Count, 30);
        for j := 0 to StringGrid.RowCount - 1 do
        begin
          if StringGrid.ColCount < DocxFile.Tables[i].Rows[j].Cols.Count then
            StringGrid.ColCount := DocxFile.Tables[i].Rows[j].Cols.Count;
          for k := 0 to DocxFile.Tables[i].Rows[j].Cols.Count - 1 do
            StringGrid.Cells[k, j] := DocxFile.Tables[i].Rows[j].Cols[k].Text;
        end;
        if (i = 0) then
          TableNumber := TabSheet.Tag;
      end;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  finally
    DocxFile.Free;
  end;
end;

procedure TfmQImport3DocxEditor.FillMap;
var
  i, j: Integer;
begin
  for i := 0 to Import.Map.Count - 1 do
    for j := 0 to lvDocxFields.Items.Count - 1 do
      if AnsiCompareText(lvDocxFields.Items[j].Caption, Import.Map.Names[i]) = 0 then
      begin
        lvDocxFields.Items[j].SubItems[0] := Import.Map.Values[Import.Map.Names[i]];
        Break;
      end;
end;

procedure TfmQImport3DocxEditor.ClearMap;
var
  i: Integer;
begin
  for i := 0 to lvDocxFields.Items.Count - 1 do
    lvDocxFields.Items[i].SubItems[0] := EmptyStr;
  lvDocxFieldsChange(lvDocxFields, lvDocxFields.Selected, ctState);
end;

procedure TfmQImport3DocxEditor.TuneButtons;
var
  Condition: boolean;
begin
  Condition := (lvDocxFields.Items.Count > 0) and FileExists(FileName);
//  sgrDocx.Enabled := Condition;
  tlbDocxAutoFill.Enabled := Condition;
  tlbDocxClear.Enabled := Condition;

  laSkip_01.Enabled := Condition;
  edDocxSkip.Enabled := Condition;
  laSkip_02.Enabled := Condition;

//  Condition := (not FNeedLoadFile) and FileExists(FileName);
end;

procedure TfmQImport3DocxEditor.ApplyChanges;
var
  i: Integer;
begin
  Import.Map.BeginUpdate;
  try
    Import.Map.Clear;
    for i := 0 to lvDocxFields.Items.Count - 1 do
      if lvDocxFields.Items[i].SubItems[0] <> EmptyStr then
        Import.Map.Values[lvDocxFields.Items[i].Caption] :=
          lvDocxFields.Items[i].SubItems[0];
  finally
    Import.Map.EndUpdate;
  end;
  Import.TableNumber := TableNumber;
  Import.FileName := FileName;
  Import.SkipFirstRows := SkipLines;
end;

function TfmQImport3DocxEditor.DocxCol: Integer;
begin
  Result := 0;
  if Assigned(lvDocxFields.Selected) then
    Result := StrToIntDef(lvDocxFields.Selected.SubItems[0], 0);
end;

procedure TfmQImport3DocxEditor.btnBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FileName;
  if odFileName.Execute then
  begin
    FileName := odFileName.FileName;
    ClearMap;
  end;
end;

procedure TfmQImport3DocxEditor.lvDocxFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  if Assigned(FCurrentStringGrid) then
    FCurrentStringGrid.Repaint;
end;

procedure TfmQImport3DocxEditor.tlbDocxAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  if Assigned(FCurrentStringGrid) then
  begin
    for i := 0 to lvDocxFields.Items.Count - 1 do
      if (i <= FCurrentStringGrid.ColCount - 1) then
        lvDocxFields.Items[i].SubItems[0] := IntToStr(i + 1)
      else
        lvDocxFields.Items[i].SubItems[0] := EmptyStr;
    lvDocxFieldsChange(lvDocxFields, lvDocxFields.Selected, ctState);
  end;
end;

procedure TfmQImport3DocxEditor.tlbDocxClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvDocxFields.Items.Count - 1 do
    lvDocxFields.Items[i].SubItems[0] := EmptyStr;
  lvDocxFieldsChange(lvDocxFields, lvDocxFields.Selected, ctState);
end;

procedure TfmQImport3DocxEditor.edDocxSkipChange(Sender: TObject);
begin
  SkipLines := StrToIntDef(edDocxSkip.Text, 0);
end;

procedure TfmQImport3DocxEditor.sgrDocxDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: Integer;
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
    grid.Canvas.FillRect(Rect);
    grid.Canvas.TextOut(X - 1, Y + 1, grid.Cells[ACol, ARow]);
  end
  else begin
    grid.DefaultDrawing := False;
    grid.Canvas.Brush.Color := clWindow;
    grid.Canvas.FillRect(Rect);
    grid.Canvas.Font.Color := clWindowText;
    grid.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, grid.Cells[ACol, ARow]);

    if (ACol = DocxCol - 1) then
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
  grid.DefaultDrawing := True;
end;

procedure TfmQImport3DocxEditor.sgrDocxMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
  grid: TqiStringGrid;
begin
  if not (Sender is TqiStringGrid) then Exit;
  grid := Sender as TqiStringGrid;

  grid.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvDocxFields.Selected) then Exit;
  if DocxCol = ACol + 1 then
    lvDocxFields.Selected.SubItems[0] := EmptyStr
  else
    if ACol > -1 then
      lvDocxFields.Selected.SubItems[0] := IntToStr(ACol + 1);

  lvDocxFieldsChange(lvDocxFields, lvDocxFields.Selected, ctState);
end;

procedure TfmQImport3DocxEditor.pcDocxFileChange(Sender: TObject);
begin
  TableNumber := pcDocxFile.ActivePage.Tag;
end;

procedure TfmQImport3DocxEditor.edFileNameChange(Sender: TObject);
begin
  if FFileName <> edFileName.Text then
  begin
    FFileName := edFileName.Text;
    FNeedLoadFile := True;
    FillGrid;
  end;
  TuneButtons;
end;

procedure TfmQImport3DocxEditor.FormShow(Sender: TObject);
begin
  Caption := Import.Name + ' - Component Editor';
end;

{$ENDIF}
{$ENDIF}

end.
