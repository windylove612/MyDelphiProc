unit fuQImport3CSVEditor;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    Vcl.Forms,
    Vcl.Dialogs,
    Vcl.StdCtrls,
    Vcl.Controls,
    Vcl.ExtCtrls,
    System.Classes,
    Data.Db,
    Vcl.Grids,
    Vcl.ComCtrls,
    Winapi.Windows,
    Vcl.Buttons,
    Vcl.ImgList,
    Vcl.ToolWin,
  {$ELSE}
    Forms,
    Dialogs,
    StdCtrls,
    Controls,
    ExtCtrls,
    Classes,
    Db,
    Grids,
    ComCtrls,
    Windows,
    Buttons,
    ImgList,
    ToolWin,
  {$ENDIF}
  {$IFDEF QI_UNICODE}
    QImport3WideStringGrid,
  {$ENDIF}
  QImport3StrTypes,
  QImport3ASCII;

type
  TfmQImport3CSVEditor = class(TForm)
    odFileName: TOpenDialog;
    laSkip_01: TLabel;
    edSkip: TEdit;
    laSkip_02: TLabel;
    laFileName: TLabel;
    edFileName: TEdit;
    Bevel2: TBevel;
    bBrowse: TSpeedButton;
    bOk: TButton;
    bCancel: TButton;
    ToolBar: TToolBar;
    tbtAutoFill: TToolButton;
    tbtClear: TToolButton;
    ImageList1: TImageList;
    lvFields: TListView;
    cbColumn: TComboBox;
    laColumn: TLabel;
    procedure bBrowseClick(Sender: TObject);
    procedure edSkipChange(Sender: TObject);
    procedure sgrCSVDrawCell(Sender: TObject; ACol, ARow: Integer;
      Rect: TRect; State: TGridDrawState);
    procedure sgrCSVMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure bCSVAutoFillClick(Sender: TObject);
    procedure bClearClick(Sender: TObject);
    procedure lvFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
    procedure cbColumnChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    FImport: TQImport3ASCII;
    FFileName: string;
    FComma: AnsiChar;
    FQuote: AnsiChar;
    FSkipLines: integer;
    FNeedLoad: boolean;

    procedure SetFileName(const Value: string);
    procedure SetSkipLines(const Value: integer);

    procedure ClearFieldList;
    procedure FillFieldList;
    procedure LoadFile;
    procedure FillGrid;
    procedure ApplyChanges;

    function CSVCol: integer;

    procedure SetEnabledControls;
  public
    sgrCSV: TqiStringGrid;

    property Import: TQImport3ASCII read FImport write FImport;
    property FileName: string read FFileName write SetFileName;
    property Comma: AnsiChar read FComma write FComma;
    property Quote: AnsiChar read FQuote write FQuote;
    property SkipLines: integer read FSkipLines write SetSkipLines;
  end;

function RunQImportCSVEditor(AImport: TQImport3ASCII): boolean;

implementation

uses
  {$IFDEF VCL16}
    {$IFDEF QI_UNICODE}
      System.WideStrings,
      QImport3GpTextFile,
    {$ENDIF}
    System.SysUtils,
    Vcl.Graphics,
  {$ELSE}
    {$IFDEF QI_UNICODE}
      {$IFDEF VCL10}
        WideStrings,
      {$ELSE}
        QImport3WideStrings,
      {$ENDIF}
      QImport3GpTextFile,
    {$ENDIF}
    SysUtils,
    Graphics,
  {$ENDIF}
  QImport3StrIDs,
  QImport3,
  QImport3Common;

{$R *.DFM}

function RunQImportCSVEditor(AImport: TQImport3ASCII): boolean;
begin
  with TfmQImport3CSVEditor.Create(Application) do
  try
    Import := AImport;
    Comma := AImport.Comma;
    FQuote := AImport.Quote;
    FillFieldList;
    LoadFile;
    FileName := AImport.FileName;
    SkipLines := AImport.SkipFirstRows;

    FNeedLoad := true;
    SetEnabledControls;

    Result := ShowModal = mrOk;
    if Result then ApplyChanges;
  finally
    Free;
  end;
end;

{ TfmQImport3CSVEditor }

procedure TfmQImport3CSVEditor.SetFileName(const Value: string);
begin
  if FFileName <> Value then begin
    FFileName := Value;
    edFileName.Text := FFileName;
    FNeedLoad := true;
    LoadFile;
  end;
  SetEnabledControls;
end;

procedure TfmQImport3CSVEditor.SetSkipLines(const Value: integer);
begin
  if FSkipLines <> Value then begin
    FSkipLines := Value;
    edSkip.Text := IntToStr(FSkipLines);
  end;
end;

procedure TfmQImport3CSVEditor.ClearFieldList;
begin
   lvFields.Items.Clear;
end;

procedure TfmQImport3CSVEditor.FillFieldList;
var
  i: integer;
begin
  if not QImportDestinationAssigned(false, FImport.ImportDestination,
      FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid)
    then Exit;
  ClearFieldList;
  for i := 0 to QImportDestinationColCount(false, FImport.ImportDestination,
    FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid) - 1 do
    with lvFields.Items.Add do begin
      Caption := QImportDestinationColName(false, FImport.ImportDestination,
        FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid,
        FImport.GridCaptionRow, i);
      SubItems.Add(EmptyStr);
      ImageIndex := 0;
    end;
end;

procedure TfmQImport3CSVEditor.LoadFile;
begin
  if not FileExists(FileName) then Exit;
  FillGrid;
  FNeedLoad := false;
end;

procedure TfmQImport3CSVEditor.bBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FFileName;
  if odFileName.Execute then FileName := odFileName.FileName;
end;

procedure TfmQImport3CSVEditor.ApplyChanges;
var
  i: integer;
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
  Import.SkipFirstRows := SkipLines;
end;

function TfmQImport3CSVEditor.CSVCol: integer;
begin
  Result := 0;
  if Assigned(lvFields.Selected) then
    Result := StrToIntDef(lvFields.Selected.SubItems[0], 0);
end;

procedure TfmQImport3CSVEditor.SetEnabledControls;
var
  Condition: boolean;
begin
  Condition := (lvFields.Items.Count > 0) and FileExists(FileName);

  tbtClear.Enabled := Condition;
  laSkip_01.Enabled := Condition;
  edSkip.Enabled := Condition;
  laSkip_02.Enabled := Condition;

  laColumn.Enabled := Condition;
  cbColumn.Enabled := Condition;
  tbtAutoFill.Enabled := Condition;
  sgrCSV.Enabled := Condition;
end;

procedure TfmQImport3CSVEditor.edSkipChange(Sender: TObject);
begin
  SkipLines := StrToIntDef(edSkip.Text, 0);
end;

procedure TfmQImport3CSVEditor.sgrCSVDrawCell(Sender: TObject; ACol,
  ARow: Integer; Rect: TRect; State: TGridDrawState);
var
  X, Y: integer;
begin
  X := Rect.Left + (Rect.Right - Rect.Left -
   sgrCSV.Canvas.TextWidth(sgrCSV.Cells[ACol, ARow])) div 2;
  Y := Rect.Top + (Rect.Bottom - Rect.Top -
    sgrCSV.Canvas.TextHeight(sgrCSV.Cells[ACol, ARow])) div 2;
  if gdFixed in State then
  begin
    if (ACol = CSVCol - 1) and (ARow = 0) then
      sgrCSV.Canvas.Font.Style := sgrCSV.Canvas.Font.Style + [fsBold]
    else
      sgrCSV.Canvas.Font.Style := sgrCSV.Canvas.Font.Style - [fsBold];
    sgrCSV.Canvas.FillRect(Rect);
    sgrCSV.Canvas.TextOut(X - 1, Y + 1, sgrCSV.Cells[ACol, ARow]);
  end
  else begin
    sgrCSV.DefaultDrawing := False;
    sgrCSV.Canvas.Brush.Color := clWindow;
    sgrCSV.Canvas.FillRect(Rect);
    sgrCSV.Canvas.Font.Color := clWindowText;
    sgrCSV.Canvas.TextRect(Rect, Rect.Left + 2, Rect.Top + 2,
      sgrCSV.Cells[ACol, ARow]);

    if (ACol = CSVCol - 1) and (ARow > 0) then
    begin
      sgrCSV.Canvas.Font.Color := clHighLightText;
      sgrCSV.Canvas.Brush.Color := clHighLight;
      Rect.Bottom := Rect.Bottom + 1;
      sgrCSV.Canvas.FillRect(Rect);
      sgrCSV.Canvas.TextOut(Rect.Left + 2, Y, sgrCSV.Cells[ACol, ARow]);
      Rect.Bottom := Rect.Bottom - 1;
    end;
  end;
  if gdFocused in State
    then DrawFocusRect(sgrCSV.Canvas.Handle, Rect);
  sgrCSV.DefaultDrawing := true;
end;

procedure TfmQImport3CSVEditor.sgrCSVMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
var
  ACol, ARow: integer;
begin
  sgrCSV.MouseToCell(X, Y, ACol, ARow);
  if not Assigned(lvFields.Selected) then Exit;
  if CSVCol = ACol + 1
    then lvFields.Selected.SubItems[0] := EmptyStr
    else lvFields.Selected.SubItems[0] := IntToStr(ACol + 1);
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3CSVEditor.FillGrid;
var
{$IFDEF QI_UNICODE}
  F: TGpTextFile;
  Str: qiString;
  Strings: TqiStrings;
{$ELSE}
  F: TextFile;
  Str: string;
  Strings: TStrings;
{$ENDIF}
  n, i: integer;
  List: TList;
begin
  for i := 0 to sgrCSV.RowCount - 1 do
    for n := 0 to sgrCSV.ColCount - 1 do
      sgrCSV.Cells[n, i] := '';

  if not FileExists(FileName) then Exit;
{$IFDEF QI_UNICODE}
  F := TGpTextFile.CreateEx(FileName, FILE_ATTRIBUTE_NORMAL, GENERIC_READ, FILE_SHARE_READ);
  F.Reset;
  try
    Strings := TWideStringList.Create;
    try
      List := TList.Create;
      try
        n := 0;
        while not F.Eof and (n <= 20) do
        begin
          Str := F.Readln;
          CSVStringToStrings(Str, Quote, Comma, Strings);
          for i := 0 to Strings.Count - 1 do
          begin
            if List.Count < i + 1 then
              List.Add(Pointer(Length(QImportLoadStr(QIW_CSV_GridCol))));
            if Integer(List[i]) < Length(Strings[i]) then
              List[i] := Pointer(Length(Strings[i]));
            if sgrCSV.ColCount < List.Count then
              sgrCSV.ColCount := List.Count;
            sgrCSV.Cells[i, n + 1] := Strings[i];
          end;
          Inc(n);
        end;
        sgrCSV.ColCount := List.Count;
        cbColumn.Items.Clear;
        cbColumn.Items.Add(EmptyStr);
        for i := 0 to List.Count - 1 do
        begin
          sgrCSV.ColWidths[i] := sgrCSV.Canvas.TextWidthW('X') * (Integer(List[i]) + 1);
          sgrCSV.Cells[i, 0] := Format(QImportLoadStr(QIW_CSV_GridCol), [i + 1]);
          cbColumn.Items.Add(IntToStr(i + 1));
        end;
        sgrCSV.RowHeights[0] := 18;
      finally
        List.Free;
      end;
    finally
      Strings.Free;
    end;
  finally
    F.Free;
  end;
{$ELSE}
  AssignFile(F, FileName);
  Reset(F);
  try
    Strings := TStringList.Create;
    try
      List := TList.Create;
      try
        n := 0;
        while not Eof(F) and (n <= 20) do
        begin
          Readln(F, Str);
          CSVStringToStrings(Str, Quote, Comma, Strings);
          for i := 0 to Strings.Count - 1 do
          begin
            if List.Count < i + 1 then
              List.Add(Pointer(Length(QImportLoadStr(QIW_CSV_GridCol))));
            if Integer(List[i]) < Length(Strings[i]) then
              List[i] := Pointer(Length(Strings[i]));
            if sgrCSV.ColCount < List.Count then
              sgrCSV.ColCount := List.Count;
            sgrCSV.Cells[i, n + 1] := Strings[i];
          end;
          Inc(n);
        end;
        sgrCSV.ColCount := List.Count;
        cbColumn.Items.Clear;
        cbColumn.Items.Add(EmptyStr);
        for i := 0 to List.Count - 1 do
        begin
          sgrCSV.ColWidths[i] := sgrCSV.Canvas.TextWidth('X') * (Integer(List[i]) + 1);
          sgrCSV.Cells[i, 0] := Format(QImportLoadStr(QIW_CSV_GridCol), [i + 1]);
          cbColumn.Items.Add(IntToStr(i + 1));
        end;
        sgrCSV.RowHeights[0] := 18;
      finally
        List.Free;
      end;
    finally
      Strings.Free;
    end;
  finally
    CloseFile(F);
  end;
{$ENDIF}  
  FNeedLoad := false;
end;

procedure TfmQImport3CSVEditor.bCSVAutoFillClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
    if (i <= sgrCSV.ColCount - 1)
      then lvFields.Items[i].SubItems[0] := IntToStr(i + 1)
      else lvFields.Items[i].SubItems[0] := EmptyStr;
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3CSVEditor.bClearClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
    lvFields.Items[i].SubItems[0] := EmptyStr;
  lvFieldsChange(lvFields, lvFields.Selected, ctState);
end;

procedure TfmQImport3CSVEditor.lvFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
begin
  if not Assigned(Item) then Exit;
  cbColumn.OnChange := nil;
  try
    if Item.SubItems.Count > 0 then
      cbColumn.ItemIndex := cbColumn.Items.IndexOf(Item.SubItems[0]);
    sgrCSV.Repaint;
  finally
    cbColumn.OnChange := cbColumnChange;;
  end;
end;

procedure TfmQImport3CSVEditor.cbColumnChange(Sender: TObject);
begin
  if not Assigned(lvFields.Selected) then Exit;
  lvFields.Selected.SubItems[0] := cbColumn.Text;
  sgrCSV.Repaint;
end;

procedure TfmQImport3CSVEditor.FormCreate(Sender: TObject);
begin
  sgrCSV := TqiStringGrid.Create(Self);
  sgrCSV.Parent := Self;
  sgrCSV.Left := 202;
  sgrCSV.Top := 63;
  sgrCSV.Width := 405;
  sgrCSV.Height := 226;
  sgrCSV.ColCount := 1;
  sgrCSV.DefaultRowHeight := 16;
  sgrCSV.FixedCols := 0;
  sgrCSV.RowCount := 21;
  sgrCSV.Font.Charset := DEFAULT_CHARSET;
  sgrCSV.Font.Color := clWindowText;
  sgrCSV.Font.Height := -11;
  sgrCSV.Font.Name := 'Courier New';
  sgrCSV.Font.Style := [];
  sgrCSV.Options := [goFixedVertLine, goFixedHorzLine, goVertLine];
  sgrCSV.ParentFont := False;
  sgrCSV.TabOrder := 1;
  sgrCSV.OnDrawCell := sgrCSVDrawCell;
  sgrCSV.OnMouseDown := sgrCSVMouseDown;

  odFileName.Filter := QImportLoadStr(QIF_CSV);
end;

end.
