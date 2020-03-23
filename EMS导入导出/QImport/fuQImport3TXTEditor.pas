unit fuQImport3TXTEditor;

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
    Vcl.ComCtrls,
    Winapi.Windows,
    Vcl.ToolWin,
    Vcl.ImgList,
    Vcl.Buttons,
  {$ELSE}
    Forms,
    Dialogs,
    StdCtrls,
    Controls,
    ExtCtrls,
    Classes,
    Db,
    ComCtrls,
    Windows,
    ToolWin,
    ImgList,
    Buttons,
  {$ENDIF}
  QImport3TXTView,
  QImport3ASCII;

type
  TfmQImport3TXTEditor = class(TForm)
    odFileName: TOpenDialog;
    lvFields: TListView;
    laSkip_01: TLabel;
    edSkip: TEdit;
    laSkip_02: TLabel;
    laFileName: TLabel;
    edFileName: TEdit;
    bBrowse: TSpeedButton;
    Bevel2: TBevel;
    bOk: TButton;
    bCancel: TButton;
    ilTXT: TImageList;
    ToolBar: TToolBar;
    tbtClear: TToolButton;
    procedure bBrowseClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure edSkipChange(Sender: TObject);
    procedure bClearClick(Sender: TObject);
    procedure lvFieldsChange(Sender: TObject; Item: TListItem;
      Change: TItemChange);
  private
    FImport: TQImport3ASCII;
    FFileName: string;
    FSkipLines: integer;

    FItemIndex: integer;
    FNeedLoad: boolean;
    FClearing: boolean;

    procedure SetFileName(const Value: string);
    procedure SetSkipLines(const Value: integer);

    procedure ClearFieldList;
    procedure FillFieldList;
    procedure LoadFile;
    procedure ApplyChanges;
    procedure ViewerChangeSelection(Sender: TObject);

    procedure SetEnabledControls;
  public
    vwTXT: TQImport3TXTViewer;

    property Import: TQImport3ASCII read FImport write FImport;
    property FileName: string read FFileName write SetFileName;
    property SkipLines: integer read FSkipLines write SetSkipLines;
  end;

function RunQImportTXTEditor(AImport: TQImport3ASCII): boolean;

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
    Vcl.Graphics,
  {$ELSE}
    SysUtils,
    Graphics,
  {$ENDIF}
  QImport3StrIDs,
  QImport3Common;

{$R *.DFM}

function RunQImportTXTEditor(AImport: TQImport3ASCII): boolean;
begin
  with TfmQImport3TXTEditor.Create(nil) do
  try
    Import := AImport;
    vwTXT.Import := AImport;
    FileName := AImport.FileName;
    FillFieldList;
    SkipLines := AImport.SkipFirstRows;

    FItemIndex := -1;

    FNeedLoad := true;

    SetEnabledControls;
    FClearing := false;

    Result := ShowModal = mrOk;
    if Result then ApplyChanges;
  finally
    Free;
  end;
end;

{ TfmQImport3TXTEditor }

procedure TfmQImport3TXTEditor.SetFileName(const Value: string);
begin
  if FFileName <> Value then
  begin
    FFileName := Value;
    edFileName.Text := FFileName;
    FNeedLoad := true;
    LoadFile;
  end;
  SetEnabledControls;
end;

procedure TfmQImport3TXTEditor.SetSkipLines(const Value: integer);
begin
  if FSkipLines <> Value then
  begin
    FSkipLines := Value;
    edSkip.Text := IntToStr(FSkipLines);
  end;
end;

procedure TfmQImport3TXTEditor.ClearFieldList;
begin
  lvFields.Items.Clear;
end;

procedure TfmQImport3TXTEditor.FillFieldList;
var
  i, j, P, L: integer;
  SS: string;
  WasActive: boolean;
  Item: TListItem;
begin
  if not QImportDestinationAssigned(false, FImport.ImportDestination,
       FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid)
    then Exit;

  WasActive := false;
  lvFields.Items.BeginUpdate;
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
      FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid) - 1 do
    begin
      Item := lvFields.Items.Add;
      Item.Caption := QImportDestinationColName(false, FImport.ImportDestination,
        FImport.DataSet, FImport.DBGrid, FImport.ListView, FImport.StringGrid,
        FImport.GridCaptionRow, i);
      Item.ImageIndex := 0;

      j := FImport.Map.IndexOfName(Item.Caption);
      if j > -1 then
      begin
        SS := FImport.Map.Values[FImport.Map.Names[j]];

        P := StrToIntDef(Copy(SS, 1, Pos(';', SS) - 1), 0);
        L := StrToIntDef(Copy(SS, Pos(';', SS) + 1, Length(SS)), 0);

        if L > 0 then
        begin
          vwTXT.AddArrow(P);
          vwTXT.AddArrow(P + L);
        end;

        Item.SubItems.Add(IntToStr(P));
        Item.SubItems.Add(IntToStr(L));
      end
      else begin
        Item.SubItems.Add(EmptyStr);
        Item.SubItems.Add(EmptyStr);
      end;
    end;
    if lvFields.Items.Count > 0 then
    begin
      lvFields.Items[0].Focused := true;
      lvFields.Items[0].Selected := true;
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

    lvFields.Items.EndUpdate;
  end;
end;

procedure TfmQImport3TXTEditor.LoadFile;
begin
  if not FileExists(FileName) then Exit;
  vwTXT.LoadFromFile(FFileName);
  FNeedLoad := false;
end;

procedure TfmQImport3TXTEditor.bBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FFileName;
  if odFileName.Execute then
    FileName := odFileName.FileName;
end;

procedure TfmQImport3TXTEditor.ApplyChanges;
var
  i: integer;
begin
  Import.Map.BeginUpdate;
  try
    Import.Map.Clear;
    for i := 0 to lvFields.Items.Count - 1 do
      if (lvFields.Items[i].SubItems[0] <> EmptyStr) and
         (lvFields.Items[i].SubItems[1] <> EmptyStr) then
        Import.Map.Values[lvFields.Items[i].Caption] :=
          Format('%s;%s', [lvFields.Items[i].SubItems[0],
                           lvFields.Items[i].SubItems[1]]);
  finally
    Import.Map.EndUpdate;
  end;
  Import.FileName := FileName;
  Import.SkipFirstRows := SkipLines;
end;

procedure TfmQImport3TXTEditor.FormCreate(Sender: TObject);
begin
  vwTXT := TQImport3TXTViewer.Create(Self);
  vwTXT.Parent := Self;
  vwTXT.Height := 226;
  vwTXT.Left := 202;
  vwTXT.Top := 63;
  vwTXT.Width := 405;
  vwTXT.OnChangeSelection := ViewerChangeSelection;
end;

procedure TfmQImport3TXTEditor.SetEnabledControls;
var
  Condition: boolean;
begin
  Condition := (lvFields.Items.Count > 0) and FileExists(FileName);

  tbtClear.Enabled := Condition;
  laSkip_01.Enabled := Condition;
  edSkip.Enabled := Condition;
  laSkip_02.Enabled := Condition;

  if Assigned(vwTXT) then
    vwTXT.Enabled := Condition;
end;

procedure TfmQImport3TXTEditor.edSkipChange(Sender: TObject);
begin
  SkipLines := StrToIntDef(edSkip.Text, 0);
end;

procedure TfmQImport3TXTEditor.bClearClick(Sender: TObject);
var
  i: integer;
begin
  for i := 0 to lvFields.Items.Count - 1 do
  begin
    lvFields.Items[i].SubItems[0] := EmptyStr;
    lvFields.Items[i].SubItems[1] := EmptyStr;
  end;
  lvFieldsChange(lvFields, lvFields.ItemFocused, ctState);
end;

procedure TfmQImport3TXTEditor.ViewerChangeSelection(Sender: TObject);
var
  P, S: integer;
begin
  if not Assigned(lvFields.ItemFocused) then Exit;
  vwTXT.GetSelection(P, S);
  if (P > -1) and (S > -1) then
  begin
    lvFields.ItemFocused.SubItems[0] := IntToStr(P);
    lvFields.ItemFocused.SubItems[1] := IntToStr(S);
  end
  else begin
    lvFields.ItemFocused.SubItems[0] := EmptyStr;
    lvFields.ItemFocused.SubItems[1] := EmptyStr;
  end;
end;

procedure TfmQImport3TXTEditor.lvFieldsChange(Sender: TObject;
  Item: TListItem; Change: TItemChange);
var
  P, S: integer;
begin
  if not Assigned(Item) then Exit;
  if Item.SubItems.Count < 2 then Exit;
  P := StrToIntDef(Item.SubItems[0], -1);
  S := StrToIntDef(Item.SubItems[1], -1);
  vwTXT.SetSelection(P, S);
end;

end.
