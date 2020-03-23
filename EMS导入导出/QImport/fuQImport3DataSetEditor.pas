unit fuQImport3DataSetEditor;

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
    Data.Db,
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
    Db,
    ImgList,
    Buttons,
  {$ENDIF}
  QImport3DataSet;

type
  TfmQImport3DataSetEditor = class(TForm)
    paButtons: TPanel;
    bOk: TButton;
    bCancel: TButton;
    paWork: TPanel;
    lstDestination: TListView;
    lstSource: TListView;
    lstMap: TListView;
    Bevel2: TBevel;
    odFileName: TOpenDialog;
    ilDBF: TImageList;
    buAdd: TSpeedButton;
    pbAdd: TPaintBox;
    buAutoFill: TSpeedButton;
    pbAutoFill: TPaintBox;
    buRemove: TSpeedButton;
    pbRemove: TPaintBox;
    buClear: TSpeedButton;
    pbClear: TPaintBox;
    procedure buAddClick(Sender: TObject);
    procedure buRemoveClick(Sender: TObject);
    procedure buClearClick(Sender: TObject);
    procedure buAutoFillClick(Sender: TObject);
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
    FImport: TQImport3DataSet;

    procedure FillListView(ListView: TListView; DataSet: TDataSet);
    procedure FillMapList;

    procedure ApplyChanges;
    procedure TuneButtons;
  public
    property Import: TQImport3DataSet read FImport write FImport;
  end;

function RunQImportDataSetEditor(AImport: TQImport3DataSet): boolean;

implementation

uses
  {$IFDEF VCL16}
    System.SysUtils,
    Vcl.Graphics;
  {$ELSE}
    SysUtils,
    Graphics;
  {$ENDIF}

{$R *.DFM}

function RunQImportDataSetEditor(AImport: TQImport3DataSet): boolean;
begin
  with TfmQImport3DataSetEditor.Create(nil) do
  try
    Import := AImport;
    FillListView(lstDestination, Import.DataSet);
    FillListView(lstSource, Import.Source);
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

{ TfmQImport3DataSetEditor }

procedure TfmQImport3DataSetEditor.FillListView(ListView: TListView; DataSet: TDataSet);
var
  i: integer;
  WasActive: boolean;
begin
  if not Assigned(DataSet) then Exit;
  ListView.Items.BeginUpdate;
  WasActive := DataSet.Active;
  try
    if not WasActive and (DataSet.FieldCount = 0) then
    try
      DataSet.Open;
    except
      Exit;
    end;
    ListView.Items.Clear;
    for i := 0 to DataSet.FieldCount - 1 do
      with ListView.Items.Add do begin
        Caption := (DataSet.Fields[i].FieldName);
        ImageIndex := 8;
      end;
    if ListView.Items.Count > 0 then begin
      ListView.Items[0].Focused := true;
      ListView.Items[0].Selected := true;
    end;
  finally
    if not WasActive and DataSet.Active then DataSet.Close;
    ListView.Items.EndUpdate;
  end;
end;

procedure TfmQImport3DataSetEditor.FillMapList;
var
  i, j: integer;
  b: boolean;
begin
  lstMap.Items.BeginUpdate;
  try
    lstDestination.Items.BeginUpdate;
    try
      lstMap.Items.Clear;
      for i := 0 to Import.Map.Count - 1 do begin
        b := false;
        for j := 0 to lstSource.Items.Count - 1 do begin
          b := b or (AnsiCompareText(lstSource.Items[j].Caption,
             Import.Map.Values[Import.Map.Names[i]]) = 0);
          if b then Break;
        end;
        if not b then Continue;
        with lstMap.Items.Add do begin
          Caption := Import.Map.Names[i];
          ImageIndex := 8;
          SubItems.Add('=');
          SubItems.Add(Import.Map.Values[Import.Map.Names[i]]);
        end;
        for j := 0 to lstDestination.Items.Count - 1 do
          if AnsiCompareText(lstDestination.Items[j].Caption,
            Import.Map.Names[i]) = 0 then begin
            lstDestination.Items[j].Delete;
            Break;
          end;
      end;
      if lstMap.Items.Count > 0 then begin
        lstMap.Items[0].Focused := true;
        lstMap.Items[0].Selected := true;
      end;
    finally
      lstDestination.Items.EndUpdate;
    end;
  finally
    lstMap.Items.EndUpdate;
  end;
end;

procedure TfmQImport3DataSetEditor.ApplyChanges;
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
end;

procedure TfmQImport3DataSetEditor.TuneButtons;
begin
  buAdd.Enabled := Assigned(lstDestination.Selected) and Assigned(lstSource.Selected);
  buRemove.Enabled := Assigned(lstMap.Selected);
  buClear.Enabled := Assigned(lstMap.Selected);
  buAutoFill.Enabled := (lstSource.Items.Count > 0) and
    ((lstDestination.Items.Count > 0) or (lstMap.Items.Count > 0));
end;

procedure TfmQImport3DataSetEditor.buAddClick(Sender: TObject);
begin
  with lstMap.Items.Add do begin
    Caption := lstDestination.Selected.Caption;
    ImageIndex := 8;
    SubItems.Add('=');
    SubItems.Add(lstSource.Selected.Caption);
    ListView.Selected := lstMap.Items[Index];
  end;
  lstDestination.Items.Delete(lstDestination.Selected.Index);
  if lstDestination.Items.Count > 0 then begin
    lstDestination.Items[0].Focused := true;
    lstDestination.Items[0].Selected := true;
  end;
  if lstMap.Items.Count > 0 then begin
    lstMap.Items[0].Focused := true;
    lstMap.Items[0].Selected := true;
  end;
  TuneButtons;
end;

procedure TfmQImport3DataSetEditor.buAutoFillClick(Sender: TObject);
var
  i, N: integer;
begin
  lstDestination.Items.BeginUpdate;
  try
    lstSource.Items.BeginUpdate;
    try
      lstMap.Items.BeginUpdate;
      try
        lstMap.Items.Clear;
        FillListView(lstDestination, Import.DataSet);
        N := lstSource.Items.Count;
        if N > lstDestination.Items.Count
          then N := lstDestination.Items.Count;
        for i := N - 1 downto 0 do begin
          with lstMap.Items.Insert(0) do begin
            Caption := lstDestination.Items[i].Caption;
            ImageIndex := 8;
            SubItems.Add('=');
            SubItems.Add(lstSource.Items[i].Caption);
          end;
          lstDestination.Items[i].Delete;
        end;
        if lstMap.Items.Count > 0 then begin
          lstMap.Items[0].Focused := true;
          lstMap.Items[0].Selected := true;
        end;
      finally
        lstMap.Items.EndUpdate;
      end;
    finally
      lstSource.Items.EndUpdate;
    end;
  finally
    lstDestination.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TfmQImport3DataSetEditor.buRemoveClick(Sender: TObject);
begin
  with lstDestination.Items.Add do begin
    Caption := lstMap.Selected.Caption;
    ImageIndex := 8;
  end;
  lstMap.Items.Delete(lstMap.Selected.Index);
  if lstMap.Items.Count > 0 then begin
    lstMap.Items[0].Focused := true;
    lstMap.Items[0].Selected := true;
  end;
  if lstDestination.Items.Count > 0 then begin
    lstDestination.Items[0].Focused := true;
    lstDestination.Items[0].Selected := true;
  end;
  TuneButtons;
end;

procedure TfmQImport3DataSetEditor.buClearClick(Sender: TObject);
begin
  lstDestination.Items.BeginUpdate;
  try
    lstMap.Items.BeginUpdate;
    try
      lstMap.Items.Clear;
      FillListView(lstDestination, Import.DataSet);
    finally
      lstMap.Items.EndUpdate;
    end;
  finally
    lstDestination.Items.EndUpdate;
  end;
  TuneButtons;
end;

procedure TfmQImport3DataSetEditor.pbAddPaint(Sender: TObject);
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

procedure TfmQImport3DataSetEditor.buAddMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAdd.Left := pbAdd.Left + 1;
  pbAdd.Top := pbAdd.Top + 1;
end;

procedure TfmQImport3DataSetEditor.buAddMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAdd.Left := pbAdd.Left - 1;
  pbAdd.Top := pbAdd.Top - 1;
end;

procedure TfmQImport3DataSetEditor.pbAutoFillPaint(Sender: TObject);
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

procedure TfmQImport3DataSetEditor.buAutoFillMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAutoFill.Left := pbAutoFill.Left + 1;
  pbAutoFill.Top := pbAutoFill.Top + 1;
end;

procedure TfmQImport3DataSetEditor.buAutoFillMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbAutoFill.Left := pbAutoFill.Left - 1;
  pbAutoFill.Top := pbAutoFill.Top - 1;
end;

procedure TfmQImport3DataSetEditor.pbRemovePaint(Sender: TObject);
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

procedure TfmQImport3DataSetEditor.buRemoveMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbRemove.Left := pbRemove.Left + 1;
  pbRemove.Top := pbRemove.Top + 1;
end;

procedure TfmQImport3DataSetEditor.buRemoveMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbRemove.Left := pbRemove.Left - 1;
  pbRemove.Top := pbRemove.Top - 1;
end;

procedure TfmQImport3DataSetEditor.pbClearPaint(Sender: TObject);
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

procedure TfmQImport3DataSetEditor.buClearMouseDown(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbClear.Left := pbClear.Left + 1;
  pbClear.Top := pbClear.Top + 1;
end;

procedure TfmQImport3DataSetEditor.buClearMouseUp(Sender: TObject;
  Button: TMouseButton; Shift: TShiftState; X, Y: Integer);
begin
  pbClear.Left := pbClear.Left - 1;
  pbClear.Top := pbClear.Top - 1;
end;

end.
