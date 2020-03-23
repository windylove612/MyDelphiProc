unit fuQImport3ODTEditor;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF ODT}
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
  QImport3ODT,
  QImport3Common,
  QImport3StrTypes,
  QImport3BaseDocumentFile,
  QImport3XLSUtils;

type
  TfmQImport3ODTEditor = class(TForm)
    pcODTFile: TPageControl;
    lvFields: TListView;
    ToolBar: TToolBar;
    tbtAutoFill: TToolButton;
    tbtClear: TToolButton;
    odFileName: TOpenDialog;
    ilODT: TImageList;
    edODTSkipRows: TEdit;
    laSkipRows: TLabel;
    cbHeaderRow: TCheckBox;
    Panel1: TPanel;
    Bevel1: TBevel;
    laFileName: TLabel;
    edFileName: TEdit;
    bBrowse: TSpeedButton;
    Panel2: TPanel;
    Bevel2: TBevel;
    bCancel: TButton;
    bOk: TButton;
    procedure bBrowseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure tbtAutoFillClick(Sender: TObject);
    procedure pcODTFileChange(Sender: TObject);
    procedure lvFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure tbtClearClick(Sender: TObject);
    procedure edODTSkipRowsChange(Sender: TObject);
    procedure cbHeaderRowClick(Sender: TObject);
    procedure bOkClick(Sender: TObject);
    procedure edFileNameChange(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FSkipLines: Integer;
    FFileName: String;
    FImport: TQImport3ODT;
    FODTFile: TODTFile;
    FSheetName: AnsiString;
    FCurrentStringGrid: TqiStringGrid;
    FIsHeaderUsed: Integer;

    procedure ApplyChanges;
    procedure ClearDataSheets;
    procedure FillFieldList;
    procedure FillGrid;
    procedure FillMap;
    procedure sgrODTMouseDown(Sender: TObject;
      Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
    procedure sgrODTDrawCell(Sender: TObject; ACol,
      ARow: Integer; Rect: TRect; State: TGridDrawState);
    procedure SetFileName(const Value: string);
    procedure SetSheetName(const Value: AnsiString);
    procedure SetSkipLines(const Value: Integer);
    procedure TuneButtons;

    function ODTCol: Integer;
  public
    property Import: TQImport3ODT read FImport write FImport;
    property FileName: String read FFileName write SetFileName;
    property SkipLines: Integer read FSkipLines write SetSkipLines;
    property SheetName: AnsiString read FSheetName write SetSheetName;
    property IsHeaderUsed: Integer read FIsHeaderUsed write FIsHeaderUsed;
  end;

var
  fmQImport3ODTEditor: TfmQImport3ODTEditor;

function RunQImportODTEditor(AImport: TQImport3ODT): boolean;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF ODT}
{$IFDEF VCL6}

uses
  fuQImport3Loading;

{$R *.dfm}

function RunQImportODTEditor(AImport: TQImport3ODT): boolean;
begin
  with TfmQImport3ODTEditor.Create(nil) do
  try
    Import := AImport;
    SkipLines := AImport.SkipFirstRows;
    if AImport.UseHeader then
      IsHeaderUsed := 0
    else
      IsHeaderUsed := 1;
    if (FileName = AImport.FileName) then
    begin
     FillFieldList;
     FillMap;
    end;
    FileName := AImport.FileName;
    SheetName := AImport.SheetName;
    
    Result := ShowModal = mrOk;
  finally
    Free;
  end;
end;

{ TfmQImport3ODTEditor }

procedure TfmQImport3ODTEditor.FillFieldList;
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

procedure TfmQImport3ODTEditor.SetFileName(const Value: string);
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

procedure TfmQImport3ODTEditor.SetSheetName(const Value: AnsiString);
var
  i: Integer;
begin
  if FSheetName <> Value then
  begin
    FSheetName := Value;
     for i := 0 to pcODTFile.PageCount - 1 do
      if pcODTFile.Pages[i].Caption = string(Value) then
      begin
        pcODTFile.ActivePage := pcODTFile.Pages[i];
        Break;
      end;
    if assigned(pcODTFile.ActivePage) then
      for i := 0 to pcODTFile.ActivePage.ComponentCount - 1 do
        if pcODTFile.ActivePage.Components[i].Tag = pcODTFile.ActivePage.Tag then
          FCurrentStringGrid := TqiStringGrid(pcODTFile.ActivePage.Components[i]);
  end;
end;

procedure TfmQImport3ODTEditor.SetSkipLines(const Value: Integer);
begin
  if FSkipLines <> Value then
  begin
    FSkipLines := Value;
    edODTSkipRows.Text := IntToStr(FSkipLines);
  end;
  if (edODTSkipRows.Text = EmptyStr) then
    edODTSkipRows.Text := IntToStr(FSkipLines);
end;

procedure TfmQImport3ODTEditor.bBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FileName;
  if odFileName.Execute then
    FileName := odFileName.FileName;
end;

procedure TfmQImport3ODTEditor.ClearDataSheets;
var
  i: integer;
begin
  for i := pcODTFile.PageCount - 1 downto 0 do
    pcODTFile.Pages[i].Free;
end;

procedure TfmQImport3ODTEditor.FillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k, i, j: integer;
  F: TForm;
  Start, Finish: TDateTime;
  W: integer;
begin
  if assigned(FODTFile) then
    FODTFile.Free;
  FODTFile := TODTFile.Create;
  if assigned(FCurrentStringGrid) then
    FCurrentStringGrid := nil;
  ClearDataSheets;
  SheetName := '';
  if not FileExists(FFileName) then Exit;
  FODTFile.FileName := FFileName;
  Start := Now;
  F := ShowLoading(Self, FFileName);
  try
    Application.ProcessMessages;
    FODTFile.Load;

    for k := 0 to FODTFile.Workbook.SpreadSheets.Count - 1 do
    begin
      TabSheet := TTabSheet.Create(pcODTFile);
      TabSheet.PageControl := pcODTFile;
      TabSheet.Caption := string(FODTFile.Workbook.SpreadSheets[k].Name);
      TabSheet.Tag := k;

      StringGrid :=
        TqiStringGrid.Create(TabSheet);
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
      if IsHeaderUsed = 0 then
        for i := 0 to StringGrid.ColCount - 1 do
          StringGrid.Cells[i, 0] := EmptyStr;
      StringGrid.OnDrawCell := sgrODTDrawCell;
      StringGrid.OnMouseDown := sgrODTMouseDown;
      for i := 0 to FODTFile.Workbook.SpreadSheets[k].DataCells.RowCount - 1 do
        for j := 0 to FODTFile.Workbook.SpreadSheets[k].DataCells.ColCount - 1 do
        begin
        StringGrid.Cells[j + 1 , i + 1 * IsHeaderUsed] :=
          FODTFile.Workbook.SpreadSheets[k].DataCells.Cells[j, i];
        W := StringGrid.Canvas.TextWidth(StringGrid.Cells[j + 1, i + 1 * IsHeaderUsed]);
        if W + 10 > StringGrid.ColWidths[j + 1] then
          if W + 10 < 130 then
            StringGrid.ColWidths[j + 1] := W + 10
          else
            StringGrid.ColWidths[j + 1] := 130;
        end;

       if (k = 0) and (SheetName = '') then
       begin
         SheetName := AnsiString(TabSheet.Caption);
       end;
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


procedure TfmQImport3ODTEditor.FormCreate(Sender: TObject);
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k: Integer;
begin
  FODTFile := TODTFile.Create;
  for k := 0 to 2 do
  begin
    TabSheet := TTabSheet.Create(pcODTFile);
    TabSheet.PageControl := pcODTFile;
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
  TuneButtons;
end;

procedure TfmQImport3ODTEditor.FormDestroy(Sender: TObject);
begin
  if assigned(FODTFile) then
    FODTFile.Free;
  if assigned(FCurrentStringGrid) then
    FCurrentStringGrid := nil;
  ClearDataSheets;    
end;

procedure TfmQImport3ODTEditor.sgrODTMouseDown(Sender: TObject;
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
  if ODTCol = ACol
    then lvFields.Selected.SubItems[0] := EmptyStr
    else lvFields.Selected.SubItems[0] := Col2Letter(ACol);
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

function TfmQImport3ODTEditor.ODTCol: Integer;
begin
  Result := 0;
  if Assigned(lvFields.Selected) then
    if lvFields.Selected.SubItems[0] <> EmptyStr then
      Result := GetColIdFromColIndex(lvFields.Selected.SubItems[0]);
end;

procedure TfmQImport3ODTEditor.sgrODTDrawCell(Sender: TObject; ACol,
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
    if (ACol = ODTCol) and (ARow = 0)then
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

    if (ACol = ODTCol) and (ARow > 0) then
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

procedure TfmQImport3ODTEditor.FillMap;
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

procedure TfmQImport3ODTEditor.ApplyChanges;
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
  if IsHeaderUsed = 0 then
    Import.UseHeader := true
  else
    Import.UseHeader := false;
end;

procedure TfmQImport3ODTEditor.tbtAutoFillClick(Sender: TObject);
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

procedure TfmQImport3ODTEditor.pcODTFileChange(Sender: TObject);
begin
    SheetName := AnsiString(pcODTFile.ActivePage.Caption);
end;

procedure TfmQImport3ODTEditor.lvFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  if Assigned(FCurrentStringGrid) then
    FCurrentStringGrid.Repaint;
end;

procedure TfmQImport3ODTEditor.tbtClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
    lvFields.Items[i].SubItems[0] := EmptyStr;
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3ODTEditor.TuneButtons;
var
  Condition: boolean;
begin
  Condition := (lvFields.Items.Count > 0) and FileExists(FileName)
    and (pcODTFile.PageCount > 0);
  tbtAutoFill.Enabled := Condition;
  tbtClear.Enabled := Condition;
  laSkipRows.Enabled := FileExists(FileName) and (pcODTFile.PageCount > 0);
  edODTSkipRows.Enabled := FileExists(FileName) and (pcODTFile.PageCount > 0);
  cbHeaderRow.Enabled := FileExists(FileName) and (pcODTFile.PageCount > 0);
  pcODTFile.Enabled := FileExists(FileName) and (pcODTFile.PageCount > 0);
  if cbHeaderRow.Enabled then
    if IsHeaderUsed = 0 then
      cbHeaderRow.Checked := true
    else
      cbHeaderRow.Checked := false;
end;

procedure TfmQImport3ODTEditor.edODTSkipRowsChange(Sender: TObject);
begin
  SkipLines := StrToIntDef(edODTSkipRows.Text, 0);
end;

procedure TfmQImport3ODTEditor.cbHeaderRowClick(Sender: TObject);
var
  i, k, j, W, border: Integer;
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  TempName: AnsiString;
  F: TForm;
  Start, Finish: TDateTime;
begin
  if not FileExists(FFileName) or (FODTFile.FileName = '') then Exit;
  if cbHeaderRow.Checked then
    IsHeaderUsed := 0
  else
    IsHeaderUsed := 1;
  border := pcODTFile.PageCount - 1;
  TempName := SheetName;
  SheetName := '';
  ClearDataSheets;
  Start := Now;
  F := ShowLoading(Self, 'Applying settings');
  try
    Application.ProcessMessages;
    for k := 0 to border do
    begin
        TabSheet := TTabSheet.Create(pcODTFile);
        TabSheet.PageControl := pcODTFile;
        TabSheet.Caption := string(FODTFile.Workbook.SpreadSheets[k].Name);
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
        if IsHeaderUsed = 0 then
          for i := 0 to StringGrid.ColCount - 1 do
            StringGrid.Cells[i, 0] := EmptyStr;
        StringGrid.OnDrawCell := sgrODTDrawCell;
        StringGrid.OnMouseDown := sgrODTMouseDown;

        for i := 0 to FODTFile.Workbook.SpreadSheets[k].DataCells.RowCount - 1 do
          for j := 0 to FODTFile.Workbook.SpreadSheets[k].DataCells.ColCount - 1 do
          begin
          StringGrid.Cells[j + 1 , i + 1 * IsHeaderUsed] :=
            FODTFile.Workbook.SpreadSheets[k].DataCells.Cells[j, i];
          W := StringGrid.Canvas.TextWidth(StringGrid.Cells[j + 1, i + 1 * IsHeaderUsed]);
          if W + 10 > StringGrid.ColWidths[j + 1] then
            if W + 10 < 130 then
              StringGrid.ColWidths[j + 1] := W + 10
            else
              StringGrid.ColWidths[j + 1] := 130;
          end;
      end;
    SheetName := TempName;
  finally
    Finish := Now;
    while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
      Finish := Now;
    if Assigned(F) then
      F.Free;
  end;
end;

procedure TfmQImport3ODTEditor.bOkClick(Sender: TObject);
begin
  if assigned(FODTFile) then
    if assigned(FODTFile.Workbook.SpreadSheets.GetSheetByName(qiString(SheetName))) then
      if FODTFile.Workbook.SpreadSheets.GetSheetByName(qiString(SheetName)).IsComplexTable then
      begin
        if MessageDlg('Selected table has a complex structure and could be improperly converted' + #10 +
          '(it contains vertically merged cells and/or subtables).'
            + #10 + 'Do you want to convert this table anyway?' + #10 +
            'You could examine internal structure of selected table by pressing No.',
              mtInformation, [mbYes, mbNo], 0) = mrYes then
        begin
          ApplyChanges;
          Self.ModalResult := mrOk;
          exit;
        end;
        exit;
      end;
  ApplyChanges;
  Self.ModalResult := mrOk;
end;

procedure TfmQImport3ODTEditor.edFileNameChange(Sender: TObject);
begin
  if FFileName <> edFileName.Text then
  begin
    FFileName := edFileName.Text;
    FillGrid;
  end;
  TuneButtons;
end;

procedure TfmQImport3ODTEditor.FormShow(Sender: TObject);
begin
  Caption := Import.Name + ' - Component Editor';
end;

{$ENDIF}
{$ENDIF}

end.
