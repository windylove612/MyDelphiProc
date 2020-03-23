unit fuQImport3DBFEditor;

{$I QImport3VerCtrl.Inc}

interface

uses
  {$IFDEF VCL16}
    Vcl.Forms,
    Vcl.ExtCtrls,
    Vcl.StdCtrls,
    Vcl.ComCtrls,
    Vcl.Dialogs,
    Vcl.Controls,
    System.Classes,
    Vcl.ImgList,
    Vcl.Buttons,
  {$ELSE}
    Forms,
    ExtCtrls,
    StdCtrls,
    ComCtrls,
    Dialogs,
    Controls,
    Classes,
    ImgList,
    Buttons,
  {$ENDIF}
  QImport3DBF;

type
  TfmQImport3DBFEditor = class(TForm)
    paFileName: TPanel;
    paButtons: TPanel;
    bOk: TButton;
    bCancel: TButton;
    paWork: TPanel;
    edFileName: TEdit;
    bBrowse: TSpeedButton;
    bvlBrowse: TBevel;
    laFileName: TLabel;
    lstDataSet: TListView;
    lstDBF: TListView;
    lstMap: TListView;
    Bevel1: TBevel;
    Bevel2: TBevel;
    odFileName: TOpenDialog;
    buAdd: TSpeedButton;
    pbAdd: TPaintBox;
    buAutoFill: TSpeedButton;
    pbAutoFill: TPaintBox;
    buRemove: TSpeedButton;
    pbRemove: TPaintBox;
    buClear: TSpeedButton;
    pbClear: TPaintBox;
    ilDBF: TImageList;
    procedure buAddClick(Sender: TObject);
    procedure buRemoveClick(Sender: TObject);
    procedure buClearClick(Sender: TObject);
    procedure bBrowseClick(Sender: TObject);
    procedure buAutoFillClick(Sender: TObject);
    procedure edFileNameChange(Sender: TObject);
    procedure pbAddPaint(Sender: TObject);
    procedure buAddMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure buAddMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbAutoFillPaint(Sender: TObject);
    procedure buAutoFillMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure buAutoFillMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbRemovePaint(Sender: TObject);
    procedure buRemoveMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure buRemoveMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure pbClearPaint(Sender: TObject);
    procedure buClearMouseDown(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
    procedure buClearMouseUp(Sender: TObject; Button: TMouseButton;
      Shift: TShiftState; X, Y: Integer);
  private
    FImport: TQImport3DBF;
    FFileName: string;

    procedure SetFileName(const Value: string);

    procedure FillDataSetList;
    procedure FillDBFList;
    procedure FillMapList;

    procedure ApplyChanges;
    procedure TuneButtons;
  public
    property Import: TQImport3DBF read FImport write FImport;
    property FileName: string read FFileName write SetFileName;
  end;

function RunQImportDBFEditor(AImport: TQImport3DBF): boolean;

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
    Winapi.Windows,
    Vcl.Graphics,
  {$ELSE}
    SysUtils,
    Windows,
    Graphics,
  {$ENDIF}
  QImport3DBFFile;

{$R *.DFM}

function RunQImportDBFEditor(AImport: TQImport3DBF): boolean;
begin
  with TfmQImport3DBFEditor.Create(nil) do
  try
    Import := AImport;
    FileName := AImport.FileName;
    FillDataSetList;
    FillMapList;
    if Assigned(Import)
      then Caption := Import.Name + ' -- Component editor';
    TuneButtons;

    Result := ShowModal = mrOk;
    if Result then ApplyChanges;
  finally
    Free;
  end;
end;

{ TfmQImport3DBFEditor }

procedure TfmQImport3DBFEditor.SetFileName(const Value: string);
begin
  if FFileName <> Trim(Value) then
  begin
    if not FileExists(Value) then
      if Application.MessageBox(PChar('File ' + Value + ' doesn''t exist. Continue ?'),
        'Warning', MB_YESNO + MB_ICONWARNING + MB_DEFBUTTON2) = IDNO
        then begin
          edFileName.Text := FFileName;
          Abort;
        end;
    if lstMap.Items.Count > 0 then
      if Application.MessageBox('File name was changed. Want you clear map list ?',
        'Question', MB_YESNO + MB_ICONQUESTION + MB_DEFBUTTON1) = IDYES
        then lstMap.Items.Clear;
    FFileName := Trim(Value);
    edFileName.Text := FFileName;
    FillDBFList;
  end;
  TuneButtons;
end;

procedure TfmQImport3DBFEditor.FillDataSetList;
var
  i: integer;
  WasActive: boolean;
begin
  if not Assigned(Import.DataSet) then Exit;
  lstDataSet.Items.BeginUpdate;
  WasActive := Import.DataSet.Active;
  try
    if not WasActive and (Import.DataSet.FieldCount = 0) then
    try
      Import.DataSet.Open;
    except
      Exit;
    end;
    lstDataSet.Items.Clear;
    for i := 0 to Import.DataSet.FieldCount - 1 do
      with lstDataSet.Items.Add do begin
        Caption := (Import.DataSet.Fields[i].FieldName);
        ImageIndex := 8;
      end;
    if lstDataSet.Items.Count > 0 then begin
      lstDataSet.Items[0].Focused := true;
      lstDataSet.Items[0].Selected := true;
    end;
  finally
    if not WasActive and Import.DataSet.Active then Import.DataSet.Close;
    lstDataSet.Items.endUpdate;
  end;
end;

procedure TfmQImport3DBFEditor.FillDBFList;
var
  DBF: TDBFRead;
  i: integer;
begin
  lstDBF.Items.BeginUpdate;
  try
    lstDBF.Items.Clear;
    if not FileExists(FileName) then Exit;
    DBF := TDBFRead.Create(FileName);
    try
      for i := 0 to DBF.FieldCount - 1 do
        with lstDBF.Items.Add do begin
          Caption := string(DBF.FieldName[i]);
          ImageIndex := 8;
        end;
      if lstDBF.Items.Count > 0 then begin
        lstDBF.Items[0].Focused := true;
        lstDBF.Items[0].Selected := true;
      end;
    finally
      DBF.Free;
    end;
  finally
    lstDBF.Items.EndUpdate;
  end;
end;

procedure TfmQImport3DBFEditor.FillMapList;
var
  i, j: integer;
  b: boolean;
begin
  lstMap.Items.BeginUpdate;
  try
    lstDataSet.Items.BeginUpdate;
    try
      lstMap.Items.Clear;
      for i := 0 to Import.Map.Count - 1 do begin
        b := false;
        for j := 0 to lstDBF.Items.Count - 1 do begin
          b := b or (AnsiCompareText(lstDBF.Items[j].Caption,
             Import.Map.Values[Import.Map.Names[i]]) = 0);
          if b then Break;
        end;
        if not b then Continue;
        with lstMap.Items.Add do begin
          Caption := Import.Map.Names[i];
          SubItems.Add('=');
          SubItems.Add(Import.Map.Values[Import.Map.Names[i]]);
          ImageIndex := 8;
        end;
        for j := 0 to lstDataSet.Items.Count - 1 do
          if AnsiCompareText(lstDataSet.Items[j].Caption,
            Import.Map.Names[i]) = 0 then begin
            lstDataSet.Items[j].Delete;
            Break;
          end;
      end;
      if lstMap.Items.Count > 0 then begin
        lstMap.Items[0].Focused := true;
        lstMap.Items[0].Selected := true;
      end;
    finally
      lstDataSet.Items.EndUpdate;
    end;
  finally
    lstMap.Items.EndUpdate;
  end;
end;

procedure TfmQImport3DBFEditor.ApplyChanges;
var
  i: integer;
begin
  Import.Map.BeginUpdate;
  try
    Import.Map.Clear;
    for i := 0 to lstMap.Items.Count - 1 do
      Import.Map.Values[lstMap.Items[i].Caption] := lstMap.Items[i].SubItems[1];
  finally
    Import.Map.EndUpdate;
  end;
  Import.FileName := FileName;
end;

procedure TfmQImport3DBFEditor.TuneButtons;
begin
  buAdd.Enabled := Assigned(lstDataSet.Selected) and Assigned(lstDBF.Selected);
  buRemove.Enabled := Assigned(lstMap.Selected);
  buClear.Enabled := Assigned(lstMap.Selected);
  buAutoFill.Enabled := (lstDBF.Items.Count > 0) and
    ((lstDataSet.Items.Count > 0) or (lstMap.Items.Count > 0));
end;

procedure TfmQImport3DBFEditor.bBrowseClick(Sender: TObject);
begin
  odFileName.FileName := FileName;
  if odFileName.Execute then FileName := odFileName.FileName;
end;

procedure TfmQImport3DBFEditor.buAddClick(Sender: TObject);
begin
  with lstMap.Items.Add do begin
    Caption := lstDataSet.Selected.Caption;
    SubItems.Add('=');
    SubItems.Add(lstDBF.Selected.Caption);
    ListView.Selected := lstMap.Items[Index];
    ImageIndex := 8;
  end;
  lstDataSet.Items.Delete(lstDataSet.Selected.Index);
  if lstDataSet.Items.Count > 0 then begin
    lstDataSet.Items[0].Focused := true;
    lstDataSet.Items[0].Selected := true;
  end;
  if lstMap.Items.Count > 0 then begin
    lstMap.Items[0].Focused := true;
    lstMap.Items[0].Selected := true;
  end;
  TuneButtons;
end;

procedure TfmQImport3DBFEditor.buAutoFillClick(Sender: TObject);
var
  i, N: integer;
begin
  lstDataSet.Items.BeginUpdate;
  try
    lstDBF.Items.BeginUpdate;
    try
      lstMap.Items.BeginUpdate;
      try
        lstMap.Items.Clear;
        FillDataSetList;
        N := lstDBF.Items.Count;
        if N > lstDataSet.Items.Count
          then N := lstDataSet.Items.Count;
        for i := N - 1 downto 0 do begin
          with lstMap.Items.Insert(0) do begin
            Caption := lstDataSet.Items[i].Caption;
            SubItems.Add('=');
            SubItems.Add(lstDBF.Items[i].Caption);
            ImageIndex := 8;
          end;
          lstDataSet.Items[i].Delete;
        end;
        if lstMap.Items.Count > 0 then begin
          lstMap.Items[0].Focused := true;
          lstMap.Items[0].Selected := true;
        end;
      finally
        lstMap.Items.EndUpdate;
      end;
    finally
      lstDBF.Items.EndUpdate;
    end;
  finally
    lstDataSet.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TfmQImport3DBFEditor.buRemoveClick(Sender: TObject);
begin
  with lstDataSet.Items.Add do begin
    Caption := lstMap.Selected.Caption;
    ImageIndex := 8;
  end;
  lstMap.Items.Delete(lstMap.Selected.Index);
  if lstMap.Items.Count > 0 then begin
    lstMap.Items[0].Focused := true;
    lstMap.Items[0].Selected := true;
  end;
  if lstDataSet.Items.Count > 0 then begin
    lstDataSet.Items[0].Focused := true;
    lstDataSet.Items[0].Selected := true;
  end;
  TuneButtons;
end;

procedure TfmQImport3DBFEditor.buClearClick(Sender: TObject);
begin
  lstDataSet.Items.BeginUpdate;
  try
    lstMap.Items.BeginUpdate;
    try
      lstMap.Items.Clear;
      FillDataSetList;
    finally
      lstMap.Items.EndUpdate;
    end;
  finally
    lstDataSet.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TfmQImport3DBFEditor.edFileNameChange(Sender: TObject);
begin
  FileName := edFileName.Text;
end;

procedure TfmQImport3DBFEditor.pbAddPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := true;
    if buAdd.Enabled
      then i := 0
      else i := 4;
    ilDBF.GetBitmap(i, Bmp);
    pbAdd.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TfmQImport3DBFEditor.buAddMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAdd.Left := pbAdd.Left + 1;
  pbAdd.Top := pbAdd.Top + 1;
end;

procedure TfmQImport3DBFEditor.buAddMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAdd.Left := pbAdd.Left - 1;
  pbAdd.Top := pbAdd.Top - 1;
end;

procedure TfmQImport3DBFEditor.pbAutoFillPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := true;
    if buAutoFill.Enabled
      then i := 1
      else i := 5;
    ilDBF.GetBitmap(i, Bmp);
    pbAutoFill.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TfmQImport3DBFEditor.buAutoFillMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAutoFill.Left := pbAutoFill.Left + 1;
  pbAutoFill.Top := pbAutoFill.Top + 1;
end;

procedure TfmQImport3DBFEditor.buAutoFillMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAutoFill.Left := pbAutoFill.Left - 1;
  pbAutoFill.Top := pbAutoFill.Top - 1;
end;

procedure TfmQImport3DBFEditor.pbRemovePaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := true;
    if buRemove.Enabled
      then i := 2
      else i := 6;
    ilDBF.GetBitmap(i, Bmp);
    pbRemove.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TfmQImport3DBFEditor.buRemoveMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbRemove.Left := pbRemove.Left + 1;
  pbRemove.Top := pbRemove.Top + 1;
end;

procedure TfmQImport3DBFEditor.buRemoveMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbRemove.Left := pbRemove.Left - 1;
  pbRemove.Top := pbRemove.Top - 1;
end;

procedure TfmQImport3DBFEditor.pbClearPaint(Sender: TObject);
var
  Bmp: TBitmap;
  i: integer;
begin
  Bmp := TBitmap.Create;
  try
    Bmp.Transparent := true;
    if buClear.Enabled
      then i := 2
      else i := 6;
    ilDBF.GetBitmap(i, Bmp);
    pbClear.Canvas.Draw(0, 0, Bmp)
  finally
    Bmp.Free;
  end;
end;

procedure TfmQImport3DBFEditor.buClearMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbClear.Left := pbClear.Left + 1;
  pbClear.Top := pbClear.Top + 1;
end;

procedure TfmQImport3DBFEditor.buClearMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbClear.Left := pbClear.Left - 1;
  pbClear.Top := pbClear.Top - 1;
end;

end.
