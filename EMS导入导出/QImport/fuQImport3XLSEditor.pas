unit fuQImport3XLSEditor;

{$I QImport3VerCtrl.Inc}

{$DEFINE XLSEDITOR_MAX_ROW}

interface

uses
  {$IFDEF VCL16}
    Vcl.Grids,
    Vcl.Forms,
    Data.Db,
    Vcl.Dialogs,
    Vcl.StdCtrls,
    Vcl.Controls,
    Vcl.ExtCtrls,
    System.Classes,
    Winapi.Windows,
    Vcl.ComCtrls,
    Vcl.ToolWin,
    Vcl.Buttons,
    Vcl.ImgList,
  {$ELSE}
    Grids,
    Forms,
    Db,
    Dialogs,
    StdCtrls,
    Controls,
    ExtCtrls,
    Classes,
    Windows,
    ComCtrls,
    ToolWin,
    Buttons,
    ImgList,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3WideStringGrid,
  {$ENDIF}
  QImport3XLS,
  QImport3StrTypes,
  QImport3XLSFile,
  QImport3XLSMapParser;

type
  TfmQImport3XLSEditor = class(TForm)
    paFileName: TPanel;
    bvlBrowse: TBevel;
    laFileName: TLabel;
    edFileName: TEdit;
    bBrowse: TSpeedButton;
    Bevel1: TBevel;
    odFileName: TOpenDialog;
    paButtons: TPanel;
    buOk: TButton;
    buCancel: TButton;
    Bevel2: TBevel;
    paXLSFieldsAndRanges: TPanel;
    lvXLSSelection: TListView;
    lvXLSFields: TListView;
    lvXLSRanges: TListView;
    tbXLSRanges: TToolBar;
    tbtXLSAddRange: TToolButton;
    tbtXLSEditRange: TToolButton;
    tbtXLSDelRange: TToolButton;
    tbtSeparator_01: TToolButton;
    tbtXLSMoveRangeUp: TToolButton;
    tbtXLSMoveRangeDown: TToolButton;
    tbXLSUtils: TToolBar;
    tbtXLSAutoFillCols: TToolButton;
    tbtXLSAutoFillRows: TToolButton;
    tbtXLSClearFieldRanges: TToolButton;
    tbtXLSClearAllRanges: TToolButton;
    laXLSSkipCols_01: TLabel;
    edXLSSkipCols: TEdit;
    laXLSSkipCols_02: TLabel;
    laXLSSkipRows_01: TLabel;
    edXLSSkipRows: TEdit;
    laXLSSkipRows_02: TLabel;
    pcXLSFile: TPageControl;
    ilWizard: TImageList;
    procedure FormDestroy(Sender: TObject);
    procedure ClearData;
    procedure bBrowseClick(Sender: TObject);
    procedure edFileNameChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure InitObjs;
    procedure FormShow(Sender: TObject);
    procedure tbtXLSAutoFillColsClick(Sender: TObject);
    procedure tbtXLSAutoFillRowsClick(Sender: TObject);
{    procedure lvXLSFieldsSelectItem(Sender: TObject; Item: TListItem;
      Selected: Boolean);}
    procedure tbtXLSAddRangeClick(Sender: TObject);
    procedure tbtXLSEditRangeClick(Sender: TObject);
    procedure lvXLSRangesChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure tbtXLSMoveRangeUpClick(Sender: TObject);
    procedure tbtXLSMoveRangeDownClick(Sender: TObject);
    procedure tbtXLSDelRangeClick(Sender: TObject);
    procedure tbtXLSClearFieldRangesClick(Sender: TObject);
    procedure lvXLSRangesDblClick(Sender: TObject);
    procedure tbtXLSClearAllRangesClick(Sender: TObject);
    procedure edXLSSkipColsChange(Sender: TObject);
    procedure edXLSSkipRowsChange(Sender: TObject);
    procedure lvXLSFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure lvXLSFieldsEnter(Sender: TObject);
    procedure lvXLSFieldsExit(Sender: TObject);
  private
    FImport: TQImport3XLS;
    FFileName: string;
    FXLSFile: TXLSFile;
    FXLSIsEditingGrid: boolean;
    FXLSGridSelection: TMapRow;
    FXLSDefinedRanges: TMapRow;

    FSkipFirstRows: integer;
    FSkipFirstCols: integer;

    procedure FillFieldList;
    procedure ClearFieldList;
    procedure ClearDataSheets;
    procedure FillGrid;

    procedure XLSDrawCell(Sender: TObject; ACol, ARow: Integer; Rect: TRect;
      State: TGridDrawState);
    procedure XLSMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure XLSSelectCell(Sender: TObject; ACol, ARow: Integer;
      var CanSelect: Boolean);
    procedure XLSGridExit(Sender: TObject);
    procedure XLSGridKeyDown(Sender: TObject; var Key: Word;
      Shift: TShiftState);

    procedure XLSStartEditing;
    procedure XLSFinishEditing;
    procedure XLSApplyEditing;
    procedure XLSDeleteSelectedRanges;
    procedure XLSFillSelection;
    function XLSGetCurrentGrid: TqiStringGrid;
    procedure XLSRepaintCurrentGrid;

    procedure TuneButtons;

    procedure SetFileName(const Value: string);
    procedure SetCaption;

    procedure ApplyChanges;

    procedure SetSkipFirstRows(const Value: integer);
    procedure SetSkipFirstCols(const Value: integer);

    procedure SetEnabledControls;
  public
    property Import: TQImport3XLS read FImport write FImport;
    property FileName: string read FFileName write SetFileName;

    property SkipFirstRows: integer read FSkipFirstRows write SetSkipFirstRows;
    property SkipFirstCols: integer read FSkipFirstCols write SetSkipFirstCols;
  end;

function RunQImportXLSEditor(AImport: TQImport3XLS): boolean;

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
    Winapi.Messages,
    System.Variants,
  {$ELSE}
    SysUtils,
    Messages,
    {$IFDEF VCL6}
      Variants,
    {$ENDIF}
  {$ENDIF}
  fuQImport3Loading,
  QImport3Common,
  QImport3XLSCalculate,
  QImport3XLSUtils,
  fuQImport3XLSRangeEdit,
  QImport3XLSCommon;

{$R *.DFM}

function RunQImportXLSEditor(AImport: TQImport3XLS): boolean;
var
  Editor: TfmQImport3XLSEditor;
begin
  Editor := TfmQImport3XLSEditor.Create(nil);
  try
    Editor.Import := AImport;
    Editor.FileName := AImport.FileName;
    Editor.SkipFirstRows := AImport.SkipFirstRows;
    Editor.SkipFirstCols := AImport.SkipFirstCols;

    Editor.FillFieldList;
    Editor.SetEnabledControls;

    Result := (Editor.ShowModal = mrOk);
    if Result then Editor.ApplyChanges;
  finally
    Editor.Free;
  end;
end;

{ TfmQImport3XLSEditor }

procedure TfmQImport3XLSEditor.FillFieldList;
var
  i, j: integer;
  WasActive: boolean;
begin
  if not QImportDestinationAssigned(false, FImport.ImportDestination,
    FImport.DataSet, FImport.DBGrid, FImport.ListView,
    FImport.StringGrid) then Exit;

  WasActive := false;
  lvXLSFields.Items.BeginUpdate;
  try
    WasActive := QImportIsDestinationActive(false, Import.ImportDestination,
      Import.DataSet, Import.DBGrid, Import.ListView, Import.StringGrid);

    ClearFieldList;

    if not WasActive and
       (QImportDestinationColCount(false, Import.ImportDestination,
          Import.DataSet, Import.DBGrid, Import.ListView,
          Import.StringGrid) = 0) then
    try
      QImportIsDestinationOpen(false, Import.ImportDestination,
        Import.DataSet, Import.DBGrid, Import.ListView, Import.StringGrid);
    except
      Exit;
    end;

    for i := 0 to QImportDestinationColCount(false, FImport.ImportDestination,
                    FImport.DataSet, FImport.DBGrid, FImport.ListView,
                    FImport.StringGrid) - 1 do
    begin
      with lvXLSFields.Items.Add do
      begin
        Caption := QImportDestinationColName(false, FImport.ImportDestination,
                     FImport.DataSet, FImport.DBGrid, FImport.ListView,
                     FImport.StringGrid, FImport.GridCaptionRow, i);
        ImageIndex := 0;
        Data := TMapRow.Create(nil);

        j := FImport.Map.IndexOfName(Caption);
        if j > -1 then
          TMapRow(Data).AsString := FImport.Map.Values[FImport.Map.Names[j]];
      end;
    end;

    if lvXLSFields.Items.Count > 0 then
    begin
      lvXLSFields.Items[0].Focused := true;
      lvXLSFields.Items[0].Selected := true;
    end;
  finally
    if not WasActive and
       QImportIsDestinationActive(false, Import.ImportDestination,
         Import.DataSet, Import.DBGrid, Import.ListView,
         Import.StringGrid) then
    try
      QImportIsDestinationClose(false, Import.ImportDestination,
        Import.DataSet, Import.DBGrid, Import.ListView, Import.StringGrid);
    except
    end;

    lvXLSFields.Items.EndUpdate;
  end;
end;

procedure TfmQImport3XLSEditor.ClearFieldList;
var
  i, j: integer;
begin
  for i := lvXLSFields.Items.Count - 1 downto 0 do
  begin
    if Assigned(lvXLSFields.Items[i].Data) then
    begin
      for j := 0 to TMapRow(lvXLSFields.Items[i].Data).Count - 1 do
        TMapRow(lvXLSFields.Items[i].Data).Delete(j);
      TMapRow(lvXLSFields.Items[i].Data).Free;
    end;
    lvXLSFields.Items.Delete(i);
  end;
end;

procedure TfmQImport3XLSEditor.ClearDataSheets;
var
  i: integer;
begin
  for i := pcXLSFile.PageCount - 1 downto 0 do
    pcXLSFile.Pages[i].Free;
end;

procedure TfmQImport3XLSEditor.FillGrid;
var
  TabSheet: TTabSheet;
  StringGrid: TqiStringGrid;
  k, i, j, n: integer;
  Cell: TbiffCell;
//  V: Variant;
  F: TForm;
  Start, Finish: TDateTime;
  W: integer;
//  ExprLen: word;
//  Expr: PByteArray;
begin
  ClearDataSheets;

  if not FileExists(FFileName) then Exit;

  FXLSFile.FileName := FFileName;
  Start := Now;
  F := ShowLoading(Self, FFileName);
  try
    Application.ProcessMessages;
    FXLSFile.Clear;
{$IFDEF XLSEDITOR_MAX_ROW}
    FXLSFile.LoadRows(XLSEDITOR_MAX_ROW_COUNT);
{$ELSE}
    FXLSFile.Load;
{$ENDIF}

    for k := 0 to FXLSFile.Workbook.WorkSheets.Count - 1 do
    begin
      TabSheet := TTabSheet.Create(pcXLSFile);
      TabSheet.PageControl := pcXLSFile;
      TabSheet.Caption := FXLSFile.Workbook.WorkSheets[k].Name;

      StringGrid := TqiStringGrid.Create(TabSheet);
      StringGrid.Parent := TabSheet;
      StringGrid.Align := alClient;
      StringGrid.ColCount := 257;
{$IFDEF XLSEDITOR_MAX_ROW}
      n := XLSEDITOR_MAX_ROW_COUNT;
{$ELSE}
      n := 256;
{$ENDIF}
      {if (Wizard.ExcelViewerRows > 0) and (Wizard.ExcelViewerRows <= 65536) then
        n := Wizard.ExcelViewerRows;}
      StringGrid.RowCount := n + 1;
      StringGrid.FixedCols := 1;
      StringGrid.FixedRows := 1;
      StringGrid.DefaultColWidth := 64;
      StringGrid.DefaultRowHeight := 16;
      StringGrid.ColWidths[0] := 30;
      StringGrid.Options := StringGrid.Options - [goRangeSelect];
      StringGrid.OnDrawCell := XLSDrawCell;
      StringGrid.OnMouseDown := XLSMouseDown;
      StringGrid.OnSelectCell := XLSSelectCell;
      StringGrid.OnExit := XLSGridExit;
      StringGrid.OnKeyDown := XLSGridKeyDown;
      StringGrid.Tag := 1000;

      GridFillFixedCells(StringGrid);

      for i := 0 to FXLSFile.Workbook.WorkSheets[k].Rows.Count - 1 do
        for j := 0 to FXLSFile.Workbook.WorkSheets[k].Rows[i].Count - 1 do
        begin
          Cell := FXLSFile.Workbook.WorkSheets[k].Rows[i][j];
          if (Cell.Col < StringGrid.ColCount - 1) and
             (Cell.Row < StringGrid.RowCount - 1) then
          begin
            case Cell.CellType of
              bctString  :
                StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] := Cell.AsString;
              bctBoolean :
                if Cell.AsBoolean
                  then StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] := 'true'
                  else StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] := 'false';
              bctNumeric :
                StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                  FloatToStr(Cell.AsFloat);
              bctDateTime:
                if Cell.IsDateOnly then
                  StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                    FormatDateTime(FImport.Formats.ShortDateFormat, Cell.AsDateTime)
                else if Cell.IsTimeOnly then
                  StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                    FormatDateTime(FImport.Formats.ShortTimeFormat, Cell.AsDateTime)
                else StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                  FormatDateTime(FImport.Formats.ShortDateFormat + ' ' +
                    FImport.Formats.ShortTimeFormat, Cell.AsDateTime);
              bctUnknown :
(*  dee test
                if Cell.IsFormula then
                begin
                  ExprLen := GetWord(Cell.Data, 20);
                  if ExprLen > 0 then
                  begin
                    GetMem(Expr, ExprLen);
                    try
                      Move(Cell.Data[22], Expr^, ExprLen);
                      V := CalculateFormula(Cell, Expr, ExprLen);
                    finally
                      FreeMem(Expr);
                    end;
                  end
                  {V := CalculateFormula(Cell, (Cell as TbiffFormula).Expression,
                    (Cell as TbiffFormula).ExprLen);
                  StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] := VarToStr(V);}
                end
                else
*)
                  StringGrid.Cells[Cell.Col + 1, Cell.Row + 1] :=
                    VarToStr(Cell.AsVariant);
            end;
            W := StringGrid.Canvas.TextWidth(StringGrid.Cells[Cell.Col + 1, Cell.Row + 1]);
            if W + 10 > StringGrid.ColWidths[Cell.Col + 1] then
              if W + 10 < 130
                then StringGrid.ColWidths[Cell.Col + 1] := W + 10
                else StringGrid.ColWidths[Cell.Col + 1] := 130;
          end;
        end;
    end;
  finally
    Finish := Now;
    while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
      Finish := Now;

    if Assigned(F) then
      F.Free;
  end;
end;

procedure TfmQImport3XLSEditor.XLSDrawCell(Sender: TObject; ACol, ARow: Integer;
  Rect: TRect; State: TGridDrawState);
var
  i: integer;
begin
  FXLSDefinedRanges.Clear;

  if lvXLSFields.Focused then
  begin
    if Assigned(lvXLSFields.ItemFocused) and
       Assigned(lvXLSFields.ItemFocused.Data) then
      for i := 0 to TMapRow(lvXLSFields.ItemFocused.Data).Count - 1 do
        FXLSDefinedRanges.Add(TMapRow(lvXLSFields.ItemFocused.Data)[i]);
  end
  else begin
    if Assigned(lvXLSFields.ItemFocused) and
       Assigned(lvXLSFields.ItemFocused.Data) and
       Assigned(lvXLSRanges.ItemFocused) then
      FXLSDefinedRanges.Add(TMapRow(lvXLSFields.ItemFocused.Data)[lvXLSRanges.ItemFocused.Index]);
  end;

  if Sender is TqiStringGrid then
    GridDrawCell(Sender as TqiStringGrid, pcXLSFile.ActivePage.Caption,
      pcXLSFile.ActivePage.PageIndex + 1, ACol, ARow, Rect, State,
      FXLSDefinedRanges, SkipFirstCols, SkipFirstRows, FXLSIsEditingGrid,
      FXLSGridSelection);

      end;

procedure TfmQImport3XLSEditor.XLSMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);

  procedure AddColRowToSelection(IsCol, IsCtrl: boolean; Number: integer);
  var
    str, str1: string;
    N: integer;
  begin
    if IsCol
      then str := Format('[%s]%s-%s;', [pcXLSFile.ActivePage.Caption, Col2Letter(Number), COLFINISH])
      else str := Format('[%s]%s-%s;', [pcXLSFile.ActivePage.Caption, Row2Number(Number), ROWFINISH]);

      N := FXLSGridSelection.IndexOfRange(str);
      if N > -1 then FXLSGridSelection.Delete(N);

      if (not IsCtrl) or (N = -1) then
      begin
        str1 := FXLSGridSelection.AsString;
        str1 := str1 + str;
        FXLSGridSelection.AsString := str1;
      end;
  end;

var
  Grid: TqiStringGrid;
  ACol, ARow, SCol, SRow, N, i: integer;
  IsCtrl, IsShift: boolean;

  procedure ChangeCurrentCell(Col, Row: integer);
  var
    Event: TSelectCellEvent;
  begin
    Event := Grid.OnSelectCell;
    Grid.OnSelectCell := nil;
    Grid.Col := Col;
    Grid.Row := Row;
    Grid.OnSelectCell := Event;
  end;

begin
  if not (Sender is TqiStringGrid) then Exit;
  Grid := Sender as TqiStringGrid;

  IsShift := GetKeyState(VK_SHIFT) < 0;
  IsCtrl := GetKeyState(VK_CONTROL) < 0;

  if not (IsShift or IsCtrl) then begin
    Grid.Repaint;
    Exit;
  end;

  Grid.MouseToCell(X, Y, ACol, ARow);

  if not ((ACol = 0) xor (ARow = 0)) then begin
    Grid.Repaint;
    Exit;
  end;

  if not FXLSIsEditingGrid then
    XLSStartEditing;

  if IsCtrl then begin
    if ACol = 0
      then N := ARow
      else N := ACol;

    AddColRowToSelection(ARow = 0, true, N);

    if ACol = 0
      then ChangeCurrentCell(Grid.Col, ARow)
      else ChangeCurrentCell(ACol, Grid.Row);
  end
  else if IsShift then begin
    SCol := Grid.Col;
    SRow := Grid.Row;

    if ACol = 0 then begin
      if SRow <= ARow then
        for i := SRow to ARow do
          AddColRowToSelection(false, false, i)
      else
        for i := SRow downto ARow do
          AddColRowToSelection(false, false, i);
      ChangeCurrentCell(Grid.Col, ARow);
    end
    else begin
      if SCol <= ACol then
        for i := SCol to ACol do
          AddColRowToSelection(true, false, i)
      else
        for i := SCol downto ACol do
          AddColRowToSelection(true, false, i);
      ChangeCurrentCell(ACol, Grid.Row);
    end;
  end;

  XLSFillSelection;
  Grid.Repaint;
end;

procedure TfmQImport3XLSEditor.XLSSelectCell(Sender: TObject; ACol, ARow: Integer;
  var CanSelect: Boolean);
var
  Grid: TqiStringGrid;
  SCol, SRow, i: integer;
  Str: string;
  IsShift, IsCtrl, Cut: boolean;
begin
  if not (Sender is TqiStringGrid) then Exit;
  Grid := Sender as TqiStringGrid;

  IsShift := GetKeyState(VK_SHIFT) < 0;
  IsCtrl := GetKeyState(VK_CONTROL) < 0;

  if not (IsShift or IsCtrl) then begin
    XLSFinishEditing;
    Exit;
  end;

  SCol := Grid.Col;
  SRow := Grid.Row;

  if IsShift and not ((SCol = ACol) or (SRow = ARow)) then begin
    XLSFinishEditing;
    Exit;
  end;

  if not FXLSIsEditingGrid then
    XLSStartEditing;

  Cut := false;
  if (SCol = ACol) and (SRow = ARow) then begin
  end
  else begin
    if IsShift then begin
      if FXLSGridSelection.Count > 0 then begin

        if SCol <> ACol then begin
          if FXLSGridSelection[FXLSGridSelection.Count - 1].RangeType = rtRow then begin
            if SCol > ACol then begin
              for i := SCol downto ACol + 1 do
                if CellInRange(FXLSGridSelection[FXLSGridSelection.Count - 1],
                     EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, i, ARow) then begin
                  Cut := true;
                  FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 :=
                    FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 - 1;
                end;
            end
            else begin
              for i := SCol to ACol - 1 do
                if (FXLSGridSelection[FXLSGridSelection.Count - 1].RangeType = rtRow) and
                   CellInRange(FXLSGridSelection[FXLSGridSelection.Count - 1],
                     EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, i, ARow) then begin
                  Cut := true;
                  FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 :=
                    FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 + 1;
                end;
            end;
          end
        end
        else if SRow <> ARow then begin
          if (FXLSGridSelection[FXLSGridSelection.Count - 1].RangeType = rtCol) then begin
            if SRow > ARow then begin
              for i := SRow downto ARow + 1 do
                if CellInRange(FXLSGridSelection[FXLSGridSelection.Count - 1],
                     EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, ACol, i) then begin
                  Cut := true;
                  FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 :=
                    FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 - 1;
                end;
            end
            else begin
              for i := SRow to ARow - 1 do
                if CellInRange(FXLSGridSelection[FXLSGridSelection.Count - 1],
                     EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, ACol, i) then begin
                  Cut := true;
                  FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 :=
                    FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 + 1;
                end;
            end;
          end;
        end;

        if (FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 = SCol) and
           (FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 = SRow) then begin
          if FXLSGridSelection[FXLSGridSelection.Count - 1].Col2 = ACol then begin
            if ARow > SRow
              then SRow := SRow + 1
              else SRow := SRow - 1;
          end
          else if FXLSGridSelection[FXLSGridSelection.Count - 1].Row2 = ARow then begin
            if ACol > SCol
              then SCol := SCol + 1
              else SCol := SCol - 1;
          end;
        end;

      end;
      if not Cut then begin
        str := FXLSGridSelection.AsString;
        if SCol = ACol then begin
          if SRow > ARow then begin
            for i := SRow downto ARow do
              if not CellInRow(FXLSGridSelection, EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, ACol, i) then
                str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(ACol), i]);
          end
          else begin
            for i := SRow to ARow do
              if not CellInRow(FXLSGridSelection, EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, ACol, i) then
                str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(ACol), i]);
          end
        end
        else if SRow = ARow then begin
          if SCol > ACol then begin
            for i := SCol downto ACol do
              if not CellInRow(FXLSGridSelection, EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, i, ARow) then
                str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(i), ARow]);
          end
          else begin
            for i := SCol to ACol do
              if not CellInRow(FXLSGridSelection, EmptyStr, pcXLSFile.ActivePage.PageIndex + 1, i, ARow) then
                str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(i), ARow]);
          end
        end;
        FXLSGridSelection.AsString := str;
      end;
      XLSFillSelection;
    end
    else if IsCtrl then begin
      if not CellInRow(FXLSGridSelection, EmptyStr,
           pcXLSFile.ActivePage.PageIndex + 1, ACol, ARow) then begin
        str := FXLSGridSelection.AsString;
        str := str + Format('[%s]%s%d;', [pcXLSFile.ActivePage.Caption, Col2Letter(ACol), ARow]);
        FXLSGridSelection.AsString := str;
      end
      else begin
        RemoveCellFromRow(FXLSGridSelection, EmptyStr,
          pcXLSFile.ActivePage.PageIndex + 1, ACol, ARow);
      end;
      XLSFillSelection;
    end;
  end;
  Grid.Repaint;
end;

procedure TfmQImport3XLSEditor.XLSGridExit(Sender: TObject);
begin
  XLSFinishEditing;
end;

procedure TfmQImport3XLSEditor.XLSGridKeyDown(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin
  if Shift = [] then
    case Key of
      VK_RETURN: XLSApplyEditing;
      VK_ESCAPE: XLSFinishEditing;
    end
end;

procedure TfmQImport3XLSEditor.XLSStartEditing;
begin
  FXLSIsEditingGrid := true;
  lvXLSRanges.Visible := false;
  lvXLSSelection.Visible := true;
  tbXLSRanges.Visible := false;
end;

procedure TfmQImport3XLSEditor.XLSFinishEditing;
begin
  FXLSIsEditingGrid := false;
  FXLSGridSelection.Clear;
  XLSRepaintCurrentGrid;

  lvXLSSelection.Visible:= false;

  lvXLSRanges.Visible := true;
  tbXLSRanges.Visible := true;;
end;

procedure TfmQImport3XLSEditor.XLSApplyEditing;
var
  i: integer;
begin
  if FXLSGridSelection.Count = 0 then Exit;
    if not Assigned(lvXLSFields.Selected) then Exit;

  XLSDeleteSelectedRanges;

  for i := 0 to FXLSGridSelection.Count - 1 do begin
    FXLSGridSelection[i].MapRow := TMapRow(lvXLSFields.Selected.Data);
    TMapRow(lvXLSFields.ItemFocused.Data).Add(FXLSGridSelection[i]);
    with lvXLSRanges.Items.Add do begin
      Caption := FXLSGridSelection[i].AsString;
      Data := FXLSGridSelection[i];
      ImageIndex := 3;
    end;
  end;
  if (lvXLSRanges.Items.Count > 0) and
     not Assigned(lvXLSRanges.ItemFocused) then begin
    lvXLSRanges.Items[0].Focused := true;
    lvXLSRanges.Items[0].Selected := true;
  end;

  XLSFinishEditing;
end;

procedure TfmQImport3XLSEditor.XLSDeleteSelectedRanges;
var
  List: TList;
  i: integer;
begin
  if lvXLSRanges.SelCount = 0 then Exit;

  lvXLSRanges.OnChange := nil;
  try
    List := TList.Create;
    try
      for i := 0 to lvXLSRanges.Items.Count - 1 do
        if lvXLSRanges.Items[i].Selected then
          List.Add(Pointer(i));
      for i := List.Count - 1 downto 0 do begin
        TMapRow(lvXLSFields.ItemFocused.Data).Delete(Integer(List[i]));
        lvXLSRanges.Items[Integer(List[i])].Delete;
        List.Delete(i);
      end;

      if (lvXLSRanges.Items.Count > 0) and Assigned(lvXLSRanges.ItemFocused) and
         not lvXLSRanges.ItemFocused.Selected then
        lvXLSRanges.ItemFocused.Selected := true;

    finally
      List.Free;
    end;
    TuneButtons;
  finally
    lvXLSRanges.OnChange := lvXLSRangesChange;
  end;
end;

procedure TfmQImport3XLSEditor.XLSFillSelection;
var
  i: integer;
begin
//  lvXLSSelection.Items.BeginUpdate;
  try
    lvXLSSelection.Items.Clear;
    FXLSGridSelection.Optimize;
    for i := 0 to FXLSGridSelection.Count - 1 do
      with lvXLSSelection.Items.Add do begin
        Caption := FXLSGridSelection[i].AsString;
        ImageIndex := 3;
      end;
    if lvXLSSelection.Items.Count > 0 then begin
      lvXLSSelection.Items[0].Focused := true;
      lvXLSSelection.Items[0].Selected := true;
    end
  finally
  //  lvXLSSelection.Items.EndUpdate;
  end;
end;

function TfmQImport3XLSEditor.XLSGetCurrentGrid: TqiStringGrid;
var
  i: integer;
begin
  Result := nil;
  if not Assigned(pcXLSFile.ActivePage) then Exit;
  for i := 0 to pcXLSFile.ActivePage.ComponentCount - 1 do
    if (pcXLSFile.ActivePage.Components[i] is TqiStringGrid) and
       ((pcXLSFile.ActivePage.Components[i] as TqiStringGrid).Tag  = 1000) then
    begin
      Result := pcXLSFile.ActivePage.Components[i] as TqiStringGrid;
      Break;
    end;
end;

procedure TfmQImport3XLSEditor.XLSRepaintCurrentGrid;
var
  Grid: TqiStringGrid;
begin
  Grid := XLSGetCurrentGrid;
  if Assigned(Grid) then Grid.Repaint;
end;

procedure TfmQImport3XLSEditor.TuneButtons;
begin
  tbtXLSAddRange.Enabled := Assigned(lvXLSFields.ItemFocused) and
                            (FileName <> EmptyStr);
  tbtXLSEditRange.Enabled := Assigned(lvXLSFields.ItemFocused) and
                            Assigned(lvXLSRanges.ItemFocused);
  tbtXLSDelRange.Enabled := Assigned(lvXLSFields.ItemFocused) and
                           (lvXLSRanges.SelCount > 0);
  tbtXLSMoveRangeUp.Enabled := Assigned(lvXLSFields.ItemFocused) and
                              Assigned(lvXLSRanges.ItemFocused) and
                              (lvXLSRanges.ItemFocused.Index > 0);
  tbtXLSMoveRangeDown.Enabled := Assigned(lvXLSFields.ItemFocused) and
                                Assigned(lvXLSRanges.ItemFocused) and
                                (lvXLSRanges.ItemFocused.Index < lvXLSRanges.Items.Count - 1);
end;

procedure TfmQImport3XLSEditor.SetCaption;
begin
end;

procedure TfmQImport3XLSEditor.SetFileName(const Value: string);
begin
  if AnsiCompareText(FFileName, Trim(Value)) <> 0 then
  begin
    FFileName := Trim(Value);
    edFileName.Text := FFileName;
    ClearDataSheets;
    if FileExists(FFileName) then FillGrid;
    SetCaption;
  end;
  SetEnabledControls;
end;

procedure TfmQImport3XLSEditor.edFileNameChange(Sender: TObject);
begin
  FileName := edFileName.Text;
end;

procedure TfmQImport3XLSEditor.ClearData;
begin
  ClearFieldList;
  FXLSDefinedRanges.Free;
  FXLSGridSelection.Free;
  FXLSFile.Free;
end;

procedure TfmQImport3XLSEditor.FormDestroy(Sender: TObject);
begin
  ClearData;
end;

procedure TfmQImport3XLSEditor.ApplyChanges;
var
  i: integer;
  str: string;
begin
  FImport.Map.BeginUpdate;
  try
    FImport.Map.Clear;
    for i := 0 to lvXLSFields.Items.Count - 1 do begin
      str := TMapRow(lvXLSFields.Items[i].Data).AsString;
      if str <> EmptyStr then
        FImport.Map.Values[lvXLSFields.Items[i].Caption] := str;
    end;
  finally
    FImport.Map.EndUpdate;
  end;
  FImport.FileName := edFileName.Text;
  FImport.SkipFirstRows := SkipFirstRows;
  FImport.SkipFirstCols := SkipFirstCols;
end;

procedure TfmQImport3XLSEditor.bBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FFileName;
  if odFileName.Execute then FileName := odFileName.FileName;
end;

procedure TfmQImport3XLSEditor.SetSkipFirstRows(const Value: integer);
begin
  if FSkipFirstRows <> Value then begin
    FSkipFirstRows := Value;
    edXLSSkipRows.Text := IntToStr(FSkipFirstRows);
    XLSRepaintCurrentGrid;
  end;
end;

procedure TfmQImport3XLSEditor.SetSkipFirstCols(const Value: integer);
begin
  if FSkipFirstCols <> Value then begin
    FSkipFirstCols := Value;
    edXLSSkipCols.Text := IntToStr(FSkipFirstCols);
    XLSRepaintCurrentGrid;
  end;
end;

procedure TfmQImport3XLSEditor.SetEnabledControls;
var
  Condition: boolean;
begin
  Condition := (lvXLSFields.Items.Count > 0) and (FileName <> EmptyStr);
  edXLSSkipRows.Enabled := Condition;
  laXLSSkipRows_01.Enabled := Condition;
  laXLSSkipRows_02.Enabled := Condition;
  edXLSSkipCols.Enabled := Condition;
  laXLSSkipCols_01.Enabled := Condition;
  laXLSSkipCols_02.Enabled := Condition;
  tbtXLSAutoFillCols.Enabled := Condition;
  tbtXLSAutoFillRows.Enabled := Condition;
  tbtXLSClearFieldRanges.Enabled := Condition;
  tbtXLSClearAllRanges.Enabled := Condition;
  tbtXLSAddRange.Enabled := Condition;
  tbtXLSEditRange.Enabled := Assigned(lvXLSFields.ItemFocused) and
                             Assigned(lvXLSRanges.ItemFocused) and
                             (FileName <> EmptyStr);
  tbtXLSDelRange.Enabled := Assigned(lvXLSFields.ItemFocused) and
                            (lvXLSRanges.SelCount > 0) and
                            (FileName <> EmptyStr);
  tbtXLSMoveRangeUp.Enabled := Assigned(lvXLSFields.ItemFocused) and
                               Assigned(lvXLSRanges.ItemFocused) and
                               (lvXLSRanges.ItemFocused.Index > 0) and
                               (FileName <> EmptyStr);
  tbtXLSMoveRangeDown.Enabled := Assigned(lvXLSFields.ItemFocused) and
                                 Assigned(lvXLSRanges.ItemFocused) and
                                 (lvXLSRanges.ItemFocused.Index < lvXLSRanges.Items.Count - 1) and
                                 (FileName <> EmptyStr);
  pcXLSFile.Enabled := Condition;
end;

procedure TfmQImport3XLSEditor.InitObjs;
begin
  FXLSFile := TXLSFile.Create;
  FXLSIsEditingGrid := false;
  FXLSGridSelection := TMapRow.Create(nil);
  FXLSDefinedRanges := TMapRow.Create(nil);
end;

procedure TfmQImport3XLSEditor.FormCreate(Sender: TObject);
begin
  InitObjs;
end;

procedure TfmQImport3XLSEditor.FormShow(Sender: TObject);
begin
  if Assigned(Import) then
    Caption := Import.Name + ' - Component Editor';
end;

procedure TfmQImport3XLSEditor.tbtXLSAutoFillColsClick(Sender: TObject);
var
  i, j: integer;
  MapRow: TMapRow;
  MR: TMapRange;
begin
  j := pcXLSFile.ActivePage.TabIndex;

  for i := 0 to lvXLSFields.Items.Count - 1 do begin
    MapRow := TMapRow(lvXLSFields.Items[i].Data);
    MapRow.Clear;
    if i <= FXLSFile.Workbook.WorkSheets[j].ColCount - 1 then begin
      MR := TMapRange.Create(MapRow);
      MR.Col1 := FXLSFile.Workbook.WorkSheets[j].Cols[i].ColNumber + 1;
      MR.Col2 := MR.Col1;
      MR.Row1 := 0;
      MR.Row2 := 0;
      MR.SheetIDType := sitName;
      MR.SheetName := FXLSFile.Workbook.WorkSheets[j].Name;
      MR.SheetNumber := 0;
      MR.Direction := rdDown;
      MapRow.Add(MR);
    end;
  end;
  if Assigned(lvXLSFields.ItemFocused) then
    lvXLSFieldsChange(lvXLSFields, lvXLSFields.ItemFocused, ctState);
  TuneButtons;
end;

procedure TfmQImport3XLSEditor.tbtXLSAutoFillRowsClick(Sender: TObject);
var
  i, j: integer;
  MapRow: TMapRow;
  MR: TMapRange;
begin
  j := pcXLSFile.ActivePage.TabIndex;

  for i := 0 to lvXLSFields.Items.Count - 1 do begin
    MapRow := TMapRow(lvXLSFields.Items[i].Data);
    MapRow.Clear;
    if i <= FXLSFile.Workbook.WorkSheets[j].RowCount - 1 then begin
      MR := TMapRange.Create(MapRow);
      MR.Row1 := FXLSFile.Workbook.WorkSheets[j].Rows[i].RowNumber + 1;
      MR.Row2 := MR.Row1;
      MR.Col1 := 0;
      MR.Col2 := 0;
      MR.SheetIDType := sitName;
      MR.SheetName := FXLSFile.Workbook.WorkSheets[j].Name;
      MR.SheetNumber := 0;
      MR.Direction := rdDown;
      MapRow.Add(MR);
    end;
  end;

  if Assigned(lvXLSFields.ItemFocused) then
    lvXLSFieldsChange(lvXLSFields, lvXLSFields.ItemFocused, ctState);
  TuneButtons;
end;

procedure TfmQImport3XLSEditor.lvXLSFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
var
  i: integer;
  Row: TMapRow;
begin
  if (csDestroying in ComponentState) then Exit;
  if not Assigned(Item) or not Assigned(Item.Data) then Exit;
  lvXLSRanges.Items.BeginUpdate;
  try
    lvXLSRanges.Items.Clear;
    Row := TMapRow(Item.Data);
    for i := 0 to Row.Count - 1 do
      with lvXLSRanges.Items.Add do begin
        Caption := Row[i].AsString;
        ImageIndex := 3;
        Data := Row[i];
      end;
    if lvXLSRanges.Items.Count > 0 then begin
      lvXLSRanges.Items[0].Focused := true;
      lvXLSRanges.Items[0].Selected := true;
    end
    else lvXLSRangesChange(lvXLSRanges, nil, ctState);

    XLSRepaintCurrentGrid;
  finally
    lvXLSRanges.Items.EndUpdate;
  end;
  TuneButtons;
end;

{procedure TfmQImportXLSEditor.lvXLSFieldsSelectItem(Sender: TObject;
  Item: TListItem; Selected: Boolean);
var
  i: integer;
  Row: TMapRow;
begin
  if not Assigned(Item) or not Assigned(Item.Data) then Exit;
//  lvXLSRanges.Items.BeginUpdate;
  try
    lvXLSRanges.Items.Clear;
    Row := TMapRow(Item.Data);
    for i := 0 to Row.Count - 1 do
      with lvXLSRanges.Items.Add do begin
        Caption := Row[i].AsString;
        ImageIndex := 3;
        Data := Row[i];
      end;
    if lvXLSRanges.Items.Count > 0 then begin
      lvXLSRanges.Items[0].Focused := true;
      lvXLSRanges.Items[0].Selected := true;
    end
    else lvXLSRangesSelectItem(lvXLSRanges, nil, false);

    XLSRepaintCurrentGrid;
  finally
  //  lvXLSRanges.Items.EndUpdate;
  end;
  TuneButtons;
end;}

procedure TfmQImport3XLSEditor.tbtXLSAddRangeClick(Sender: TObject);
var
  Range: TMapRange;
  MapRow: TMapRow;
  i: integer;
  Item: TListItem;
begin
  MapRow := TMapRow(lvXLSFields.ItemFocused.Data);
  Range := TMapRange.Create(MapRow);
  Range.Col1 := 1;
  Range.Row1 := 0;
  Range.Col2 := 1;
  Range.Row2 := 0;
  Range.Direction := rdDown;
  Range.SheetIDType := sitName;
  Range.SheetName := pcXLSFile.ActivePage.Caption;
  Range.SheetNumber := pcXLSFile.ActivePage.PageIndex + 1;
  Range.Update;

  if EditRange(Range, FXLSFile) then begin
    MapRow.Add(Range);
    lvXLSRanges.Items.BeginUpdate;
    try
      Item := lvXLSRanges.Items.Add;
      with Item do begin
        Caption := Range.AsString;
        ImageIndex := 3;
        Data := Range;
      end;
      for i := 0 to lvXLSRanges.Items.Count - 1 do begin
        lvXLSRanges.Items[i].Focused := lvXLSRanges.Items[i] = Item;
        lvXLSRanges.Items[i].Selected := lvXLSRanges.Items[i] = Item;
      end;
    finally
      lvXLSRanges.Items.EndUpdate;
    end;
  end
  else Range.Free;
  TuneButtons;
end;

procedure TfmQImport3XLSEditor.tbtXLSEditRangeClick(Sender: TObject);
begin
  if not Assigned(lvXLSFields.ItemFocused) and
     not Assigned(lvXLSRanges.ItemFocused) then Exit;

  if EditRange(TMapRange(lvXLSRanges.ItemFocused.Data), FXLSFile) then
    lvXLSRanges.ItemFocused.Caption := TMapRange(lvXLSRanges.ItemFocused.Data).AsString;
  XLSRepaintCurrentGrid;
  TuneButtons;
end;

procedure TfmQImport3XLSEditor.lvXLSRangesChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if (csDestroying in ComponentState) then Exit;
  XLSRepaintCurrentGrid;
  TuneButtons;
end;

procedure TfmQImport3XLSEditor.tbtXLSMoveRangeUpClick(Sender: TObject);
var
  Index, i: integer;
begin
  Index := lvXLSRanges.ItemFocused.Index;
  TMapRow(lvXLSFields.ItemFocused.Data).Exchange(Index, Index - 1);
  lvXLSRanges.Items.BeginUpdate;
  try
    lvXLSRanges.Items[Index].Data := TMapRow(lvXLSFields.ItemFocused.Data)[Index];
    lvXLSRanges.Items[Index].Caption := TMapRange(lvXLSRanges.Items[Index].Data).AsString;
    lvXLSRanges.Items[Index - 1].Data := TMapRow(lvXLSFields.ItemFocused.Data)[Index - 1];
    lvXLSRanges.Items[Index - 1].Caption := TMapRange(lvXLSRanges.Items[Index - 1].Data).AsString;
    for i := 0 to lvXLSRanges.Items.Count - 1 do begin
      lvXLSRanges.Items[i].Focused := lvXLSRanges.Items[i] = lvXLSRanges.Items[Index - 1];
      lvXLSRanges.Items[i].Selected := lvXLSRanges.Items[i] = lvXLSRanges.Items[Index - 1];
    end;
  finally
    lvXLSRanges.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TfmQImport3XLSEditor.tbtXLSMoveRangeDownClick(Sender: TObject);
var
  Index, i: integer;
begin
  lvXLSRanges.Visible := false;
  lvXLSRanges.Visible := true;
  //Exit;
  
  Index := lvXLSRanges.ItemFocused.Index;
  TMapRow(lvXLSFields.ItemFocused.Data).Exchange(Index, Index + 1);
  lvXLSRanges.Items.BeginUpdate;
  try
    lvXLSRanges.Items[Index].Data := TMapRow(lvXLSFields.ItemFocused.Data)[Index];
    lvXLSRanges.Items[Index].Caption := TMapRange(lvXLSRanges.Items[Index].Data).AsString;
    lvXLSRanges.Items[Index + 1].Data := TMapRow(lvXLSFields.ItemFocused.Data)[Index + 1];
    lvXLSRanges.Items[Index + 1].Caption := TMapRange(lvXLSRanges.Items[Index + 1].Data).AsString;
    for i := 0 to lvXLSRanges.Items.Count - 1 do begin
      lvXLSRanges.Items[i].Focused := lvXLSRanges.Items[i] = lvXLSRanges.Items[Index + 1];
      lvXLSRanges.Items[i].Selected := lvXLSRanges.Items[i] = lvXLSRanges.Items[Index + 1];
    end;
  finally
    lvXLSRanges.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TfmQImport3XLSEditor.tbtXLSDelRangeClick(Sender: TObject);
begin
  XLSDeleteSelectedRanges;
  XLSRepaintCurrentGrid;
end;

procedure TfmQImport3XLSEditor.tbtXLSClearFieldRangesClick(Sender: TObject);
var
  k: Integer;
begin
  if not (Assigned(lvXLSFields.ItemFocused) and
          Assigned(lvXLSFields.ItemFocused.Data)) then Exit;

  for k := 0 to TMapRow(lvXLSFields.ItemFocused.Data).Count - 1 do
    TMapRow(lvXLSFields.ItemFocused.Data).Delete(k);

  TMapRow(lvXLSFields.ItemFocused.Data).Clear;
  lvXLSFieldsChange(lvXLSFields, lvXLSFields.ItemFocused, ctState);
end;

procedure TfmQImport3XLSEditor.lvXLSRangesDblClick(Sender: TObject);
begin
  if tbtXLSEditRange.Enabled then
    tbtXLSEditRange.Click;
end;

procedure TfmQImport3XLSEditor.tbtXLSClearAllRangesClick(Sender: TObject);
var
  i, k: integer;
begin
  for i := 0 to lvXLSFields.Items.Count - 1 do
  begin

    for k := 0 to TMapRow(lvXLSFields.Items[i].Data).Count - 1 do
      TMapRow(lvXLSFields.Items[i].Data).Delete(k);

    TMapRow(lvXLSFields.Items[i].Data).Clear;
  end;
  lvXLSFieldsChange(lvXLSFields, lvXLSFields.ItemFocused, ctState);
  TuneButtons;
end;

procedure TfmQImport3XLSEditor.edXLSSkipColsChange(Sender: TObject);
begin
  SkipFirstCols := StrToIntDef(edXLSSkipCols.Text, 0);
end;

procedure TfmQImport3XLSEditor.edXLSSkipRowsChange(Sender: TObject);
begin
  SkipFirstRows := StrToIntDef(edXLSSkipRows.Text, 0);
end;

procedure TfmQImport3XLSEditor.lvXLSFieldsEnter(Sender: TObject);
begin
  XLSRepaintCurrentGrid;
end;

procedure TfmQImport3XLSEditor.lvXLSFieldsExit(Sender: TObject);
begin
  XLSRepaintCurrentGrid;
end;

end.
