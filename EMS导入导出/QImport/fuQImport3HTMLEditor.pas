unit fuQImport3HTMLEditor;

{$I QImport3VerCtrl.Inc}

interface

{$IFDEF HTML}
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
    Vcl.StdCtrls,
    Vcl.ToolWin,
    Vcl.Grids,
    Vcl.ExtCtrls,
    Vcl.ImgList,
    Vcl.Buttons,
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
    StdCtrls,
    ToolWin,
    Grids,
    ExtCtrls,
    ImgList,
    Buttons,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3WideStringGrid,
  {$ENDIF}
  QImport3StrTypes,
  QImport3HTML;

type
  TfmQImport3HTMLEditor = class(TForm)
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
    ilHTML: TImageList;
    laTable: TLabel;
    cbTable: TComboBox;
    procedure bBrowseClick(Sender: TObject);
    procedure cbTableChange(Sender: TObject);
    procedure cbColumnChange(Sender: TObject);
    procedure edSkipChange(Sender: TObject);
    procedure lvFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure sgrHTMLDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure sgrHTMLMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure tbtAutoFillClick(Sender: TObject);
    procedure tbtClearClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormShow(Sender: TObject);
  private
    FImport: TQImport3HTML;
    FNeedLoadFile: Boolean;
    FFileName: String;
    FSkipLines: Integer;
    FHTML: THTMLFile;

    procedure SetFileName(const Value: String);
    procedure SetSkipLines(const Value: Integer);

    procedure FillFieldList;
    procedure FillGrid;
    procedure FillComboColumn;
    procedure FillMap;
    procedure TuneButtons;
    procedure ApplyChanges;

    function HTMLCol: Integer;
  public
    sgrHTML: TqiStringGrid;
    property Import: TQImport3HTML read FImport write FImport;
    property FileName: String read FFileName write SetFileName;
    property SkipLines: Integer read FSkipLines write SetSkipLines;
  end;

var
  fmQImport3HTMLEditor: TfmQImport3HTMLEditor;

function RunQImportHTMLEditor(AImport: TQImport3HTML): boolean;

{$ENDIF}
{$ENDIF}

implementation

{$IFDEF HTML}
{$IFDEF VCL6}

uses
  {$IFDEF VCL16}
    System.Math,
  {$ELSE}
    Math,
  {$ENDIF}
  QImport3Common,
  fuQImport3Loading;

{$R *.dfm}

function RunQImportHTMLEditor(AImport: TQImport3HTML): boolean;
begin
  with TfmQImport3HTMLEditor.Create(nil) do
  try
    Import := AImport;
    FileName := AImport.FileName;
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

{ TfmQImport3HTMLEditor }

procedure TfmQImport3HTMLEditor.FormCreate(Sender: TObject);
begin
  sgrHTML := TqiStringGrid.Create(Self);
  sgrHTML.Parent := Self;
  sgrHTML.Left := 202;
  sgrHTML.Top := 60;
  sgrHTML.Width := 405;
  sgrHTML.Height := 293;
  sgrHTML.ColCount := 1;
  sgrHTML.DefaultRowHeight := 16;
  sgrHTML.FixedCols := 0;
  sgrHTML.RowCount := 1;
  sgrHTML.FixedRows := 0;
  sgrHTML.Font.Charset := DEFAULT_CHARSET;
  sgrHTML.Font.Color := clWindowText;
  sgrHTML.Font.Height := -11;
  sgrHTML.Font.Name := 'Courier New';
  sgrHTML.Font.Style := [];
  sgrHTML.Options := [goFixedVertLine, goFixedHorzLine, goVertLine, goHorzLine];
  sgrHTML.ParentFont := False;
  sgrHTML.TabOrder := 4;
  sgrHTML.OnDrawCell := sgrHTMLDrawCell;
  sgrHTML.OnMouseDown := sgrHTMLMouseDown;

  FNeedLoadFile := True;
end;

procedure TfmQImport3HTMLEditor.bBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FileName;
  if odFileName.Execute then FileName := odFileName.FileName;
end;

procedure TfmQImport3HTMLEditor.cbTableChange(Sender: TObject);
begin
  FillGrid;
end;

procedure TfmQImport3HTMLEditor.cbColumnChange(Sender: TObject);
begin
  if not Assigned(lvFields.Selected) then Exit;
  lvFields.Selected.SubItems[0] := cbColumn.Text;
  sgrHTML.Repaint;
end;

procedure TfmQImport3HTMLEditor.edSkipChange(Sender: TObject);
begin
  SkipLines := StrToIntDef(edSkip.Text, 0);
end;

procedure TfmQImport3HTMLEditor.lvFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  cbColumn.OnChange := nil;
  try
    if Item.SubItems.Count > 0 then
      cbColumn.ItemIndex := cbColumn.Items.IndexOf(Item.SubItems[0]);
    sgrHTML.Repaint;
  finally
    cbColumn.OnChange := cbColumnChange;
  end;
end;

procedure TfmQImport3HTMLEditor.sgrHTMLDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: Integer;
begin
  X := Rect.Left + (Rect.Right - Rect.Left -
    sgrHTML.Canvas.TextWidth(sgrHTML.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    sgrHTML.Canvas.TextHeight(sgrHTML.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    sgrHTML.Canvas.FillRect(Rect);
    sgrHTML.Canvas.TextOut(X - 1, Y + 1, sgrHTML.Cells[ACol, ARow]);
  end
  else begin
    sgrHTML.DefaultDrawing := False;
    sgrHTML.Canvas.Brush.Color := clWindow;
    sgrHTML.Canvas.FillRect(Rect);
    sgrHTML.Canvas.Font.Color := clWindowText;
    sgrHTML.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2, sgrHTML.Cells[ACol, ARow]);

    if (ACol = HTMLCol - 1) then
    begin
      sgrHTML.Canvas.Font.Color := clHighLightText;
      sgrHTML.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      sgrHTML.Canvas.FillRect(Rect);
      sgrHTML.Canvas.TextOut(Rect.Left + 2, Y, sgrHTML.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State then
    DrawFocusRect(sgrHTML.Canvas.Handle, Rect);
  sgrHTML.DefaultDrawing := True;
end;

procedure TfmQImport3HTMLEditor.sgrHTMLMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: Integer;
begin
  sgrHTML.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvFields.Selected) then Exit;
  if HTMLCol = ACol + 1 then
    lvFields.Selected.SubItems[0] := EmptyStr
  else
    if ACol > -1 then
      lvFields.Selected.SubItems[0] := IntToStr(ACol + 1);
      
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3HTMLEditor.tbtAutoFillClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
    if (i <= sgrHTML.ColCount - 1) then
      lvFields.Items[i].SubItems[0] := IntToStr(i + 1)
    else
      lvFields.Items[i].SubItems[0] := EmptyStr;
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3HTMLEditor.tbtClearClick(Sender: TObject);
var
  i: Integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
    lvFields.Items[i].SubItems[0] := EmptyStr;
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3HTMLEditor.SetFileName(const Value: String);
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

procedure TfmQImport3HTMLEditor.SetSkipLines(const Value: Integer);
begin
  if FSkipLines <> Value then
  begin
    FSkipLines := Value;
    edSkip.Text := IntToStr(FSkipLines);
  end;
end;

procedure TfmQImport3HTMLEditor.FillFieldList;
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

procedure TfmQImport3HTMLEditor.FillGrid;
var
  F: TForm;
  Start, Finish: TDateTime;
  i, j:  Integer;
begin
  if not FileExists(FileName) then Exit;

  if FNeedLoadFile then  
  begin
    Start := Now;
    F := ShowLoading(Self, FFileName);
    try
      Application.ProcessMessages;
      if Assigned(FHTML) then FHTML.Free;
      FHTML := THTMLFile.Create;
      FHTML.FileName := FileName;
      FHTML.Load(0);
      cbTable.Items.Clear;
      if FHTML.TableList.Count >= 0 then
      begin
        for i := 0 to FHTML.TableList.Count - 1 do
          cbTable.Items.Add(IntToStr(Succ(i)));
        cbTable.ItemIndex := 0;
      end;
      FNeedLoadFile := False;
    finally
      Finish := Now;
      while (Finish - Start) < EncodeTime(0, 0, 0, 500) do
        Finish := Now;
      if Assigned(F) then
        F.Free;
    end;
  end;

  if FHTML.TableList.Count >= 0 then
  begin
    sgrHTML.ColCount := 1;
    sgrHTML.RowCount := Min(FHTML.TableList[cbTable.ItemIndex].Rows.Count, 30);
    for i := 0 to sgrHTML.RowCount - 1 do
    begin
      if sgrHTML.ColCount < FHTML.TableList[cbTable.ItemIndex].Rows[i].Cells.Count then
        sgrHTML.ColCount := FHTML.TableList[cbTable.ItemIndex].Rows[i].Cells.Count;
      for j := 0 to FHTML.TableList[cbTable.ItemIndex].Rows[i].Cells.Count - 1 do
        sgrHTML.Cells[j, i] := FHTML.TableList[cbTable.ItemIndex].Rows[i].Cells[j].Text;
    end;
    FillComboColumn;
  end;
end;

procedure TfmQImport3HTMLEditor.FillComboColumn;
var
  i: Integer;
begin
  cbColumn.Clear;
  cbColumn.Items.Add('');
  cbColumn.ItemIndex := 0;
  for i := 0 to sgrHTML.ColCount - 1 do
    cbColumn.Items.Add(IntToStr(Succ(i)));
end;

procedure TfmQImport3HTMLEditor.FillMap;
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

procedure TfmQImport3HTMLEditor.TuneButtons;
var
  Condition: boolean;
begin
  Condition := (lvFields.Items.Count > 0) and FileExists(FileName);
  sgrHTML.Enabled := Condition;
  tbtAutoFill.Enabled := Condition;
  tbtClear.Enabled := Condition;

  laSkip_01.Enabled := Condition;
  edSkip.Enabled := Condition;
  laSkip_02.Enabled := Condition;
  laColumn.Enabled := Condition;
  cbColumn.Enabled := Condition;

  Condition := (not FNeedLoadFile) and FileExists(FileName);
  laTable.Enabled := Condition;
  cbTable.Enabled := Condition;
end;

procedure TfmQImport3HTMLEditor.ApplyChanges;
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
  Import.TableNumber := StrToIntDef(cbTable.Text, 0);
  Import.FileName := FileName;
  Import.SkipFirstRows := SkipLines;
end;

function TfmQImport3HTMLEditor.HTMLCol: Integer;
begin
  Result := 0;
  if Assigned(lvFields.Selected) then
    Result := StrToIntDef(lvFields.Selected.SubItems[0], 0);
end;

procedure TfmQImport3HTMLEditor.FormShow(Sender: TObject);
begin
  Caption := Import.Name + ' - Component Editor';
end;

{$ENDIF}
{$ENDIF}

end.
