unit fuQImport3XLSRangeEdit;
{$I QImport3VerCtrl.Inc}
interface

uses
  {$IFDEF VCL16}
    Winapi.Windows,
    Winapi.Messages,
    System.SysUtils,
    System.Classes,
    Vcl.Graphics,
    Vcl.Controls,
    Vcl.Forms,
    Vcl.Dialogs,
    Vcl.StdCtrls,
    Vcl.ExtCtrls,
  {$ELSE}
    Windows,
    Messages,
    SysUtils,
    Classes,
    Graphics,
    Controls,
    Forms,
    Dialogs,
    StdCtrls,
    ExtCtrls,
  {$ENDIF}
  QImport3XLSMapParser,
  QImport3XLSUtils,
  QImport3XLSFile;

type
  TfmQImport3XLSRangeEdit = class(TForm)
    gbRangeType: TGroupBox;
    laColRowCell: TLabel;
    edColRowCell: TEdit;
    gbFinishCondition: TGroupBox;
    rbFinishColRow: TRadioButton;
    rbFinishData: TRadioButton;
    edFinishColRow: TEdit;
    gbSheetType: TGroupBox;
    rbDefaultSheet: TRadioButton;
    rbCustomSheet: TRadioButton;
    bCancel: TButton;
    bOk: TButton;
    paCustomSheet: TPanel;
    rbSheetNumber: TRadioButton;
    edSheetNumber: TComboBox;
    rbSheetName: TRadioButton;
    edSheetName: TComboBox;
    cbRangeType: TComboBox;
    gbStartCondition: TGroupBox;
    rbStartData: TRadioButton;
    rbStartColRow: TRadioButton;
    edStartColRow: TEdit;
    rgDirection: TRadioGroup;
    procedure bOkClick(Sender: TObject);
    procedure UpdateControls(Sender: TObject);
    procedure FormCreate(Sender: TObject);
  private
    FRange: TMapRange;
    procedure FillData;
    procedure LoadData;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

function EditRange(Range: TMapRange; XLSFile: TXLSFile): boolean;

implementation

uses QImport3StrIDs, QImport3;

{$R *.DFM}

function EditRange(Range: TMapRange; XLSFile: TXLSFile): boolean;
var
  i: integer;
begin
  with TfmQImport3XLSRangeEdit.Create(nil) do
  try
    FRange.Assign(Range);
    for i := 0 to XLSFile.Workbook.WorkSheets.Count - 1 do
    begin
      edSheetNumber.Items.Add(IntToStr(i + 1));
      edSheetName.Items.Add(XLSFile.Workbook.WorkSheets[i].Name);
    end;
    FillData;
    Result := ShowModal = mrOk;
    if Result then Range.Assign(FRange);
  finally
    Free;
  end;
end;

{ TfmQImport3XLSRangeEdit }

constructor TfmQImport3XLSRangeEdit.Create(AOwner: TComponent);
begin
  inherited;
  FRange := TMapRange.Create(nil);
end;

destructor TfmQImport3XLSRangeEdit.Destroy;
begin
  FRange.Free;
  inherited;
end;

procedure TfmQImport3XLSRangeEdit.FillData;
begin
  FRange.Update;
  cbRangeType.ItemIndex := Integer(FRange.RangeType);

  case FRange.RangeType of
    rtCell:
      edColRowCell.Text := Col2Letter(FRange.Col1) + Row2Number(FRange.Row1);
    rtCol: begin
      edColRowCell.Text := Col2Letter(FRange.Col1);
      rbStartData.Checked := FRange.Row1 = 0;
      rbStartColRow.Checked := FRange.Row1 > 0;
      if rbStartColRow.Checked
        then edStartColRow.Text := Row2Number(FRange.Row1)
        else edStartColRow.Text := EmptyStr;
      rbFinishData.Checked := FRange.Row2 = 0;
      rbFinishColRow.Checked := FRange.Row2 > 0;
      if rbFinishColRow.Checked
        then edFinishColRow.Text := Row2Number(FRange.Row2)
        else edFinishColRow.Text := EmptyStr;
      if (FRange.Row1 = 0) or (FRange.Row2 = 0) then
        rgDirection.ItemIndex := Integer(FRange.Direction);
    end;
    rtRow: begin
      edColRowCell.Text := Row2Number(FRange.Row1);
      rbStartData.Checked := FRange.Col1 = 0;
      rbStartColRow.Checked := FRange.Col1 > 0;
      if rbStartColRow.Checked
        then edStartColRow.Text := Col2Letter(FRange.Col1)
        else edStartColRow.Text := EmptyStr;
      rbFinishData.Checked := FRange.Col2 = 0;
      rbFinishColRow.Checked := FRange.Col2 > 0;
      if rbFinishColRow.Checked
        then edFinishColRow.Text := Col2Letter(FRange.Col2)
        else edFinishColRow.Text := EmptyStr;
      if (FRange.Col1 = 0) or (FRange.Col2 = 0) then
        rgDirection.ItemIndex := Integer(FRange.Direction);
    end;
  end;

  rbDefaultSheet.Checked := FRange.SheetIDType = sitUnknown;
  rbCustomSheet.Checked := FRange.SheetIDType <> sitUnknown;
  rbSheetNumber.Checked := FRange.SheetIDType = sitNumber;
  edSheetNumber.ItemIndex := edSheetNumber.Items.IndexOf(IntToStr(FRange.SheetNumber));
  rbSheetName.Checked := FRange.SheetIDType = sitName;
  edSheetName.ItemIndex := edSheetName.Items.IndexOf(FRange.SheetName);
end;

procedure TfmQImport3XLSRangeEdit.LoadData;
var
  C, R: integer;
  N: string;
begin
  case TRangeType(cbRangeType.ItemIndex) of
    rtCol: begin
      try
        ParseColString(edColRowCell.Text, C);
      except
        edColRowCell.SetFocus;
        raise;
      end;
      FRange.Col1 := C;
      FRange.Col2 := C;

      if rbStartColRow.Checked then begin
        try
          ParseRowString(edStartColRow.Text, R);
        except
          edStartColRow.SetFocus;
          raise;
        end;
        FRange.Row1 := R;
      end
      else FRange.Row1 := 0;

      if rbFinishColRow.Checked then begin
        try
          ParseRowString(edFinishColRow.Text, R);
        except
          edFinishColRow.SetFocus;
          raise;
        end;
        FRange.Row2 := R;
      end
      else FRange.Row2 := 0;

      if (FRange.Row1 = 0) or (FRange.Row2 = 0) then
        FRange.Direction := TRangeDirection(rgDirection.ItemIndex)
      else if FRange.Row1 < FRange.Row2 then
        FRange.Direction := rdDown
      else if FRange.Row1 > FRange.Row2 then
        FRange.Direction := rdUp
      else if (FRange.Row1 = FRange.Row2) and (FRange.Row1 > 1) then begin
        FRange.Direction := rdUnknown;
        FRange.Update;
      end;
    end;
    rtRow: begin
      try
        ParseRowString(edColRowCell.Text, R);
      except
        edColRowCell.SetFocus;
        raise;
      end;
      FRange.Row1 := R;
      FRange.Row2 := R;

      if rbStartColRow.Checked then begin
        try
          ParseColString(edStartColRow.Text, C);
        except
          edStartColRow.SetFocus;
          raise;
        end;
        FRange.Col1 := C;
      end
      else FRange.Col1 := 0;

      if rbFinishColRow.Checked then begin
        try
          ParseColString(edFinishColRow.Text, C);
        except
          edFinishColRow.SetFocus;
          raise;
        end;
        FRange.Col2 := C;
      end
      else FRange.Col2 := 0;

      if (FRange.Col1 = 0) or (FRange.Col2 = 0) then
        FRange.Direction := TRangeDirection(rgDirection.ItemIndex)
      else if FRange.Col1 < FRange.Col2 then
        FRange.Direction := rdDown
      else if FRange.Col1 > FRange.Col2 then
        FRange.Direction := rdUp
      else if (FRange.Col1 = FRange.Col2) and (FRange.Col1 > 1) then begin
        FRange.Direction := rdUnknown;
        FRange.Update;
      end;
    end;
    rtCell: begin
      try
        ParseCellString(edColRowCell.Text, C, R);
      except
        edColRowCell.SetFocus;
        raise;
      end;
      FRange.Col1 := C;
      FRange.Row1 := R;
      FRange.Col2 := C;
      Frange.Row2 := R;
      FRange.Direction := rdUnknown;
    end;
  end;

  if rbDefaultSheet.Checked then begin
    FRange.SheetIDType := sitUnknown;
    FRange.SheetNumber := 0;
    FRange.SheetName := EmptyStr;
  end
  else if rbCustomSheet.Checked then begin
    if rbSheetNumber.Checked then begin
      FRange.SheetIDType := sitNumber;
      ParseSheetNumber(edSheetNumber.Text, C);
      FRange.SheetNumber := C;
      FRange.SheetName := EmptyStr;
    end
    else if rbSheetName.Checked then begin
      FRange.SheetIDType := sitName;
      ParseSheetName(edSheetName.Text, N);
      FRange.SheetNumber := 0;
      FRange.SheetName := N;
    end;
  end;
end;

procedure TfmQImport3XLSRangeEdit.UpdateControls(Sender: TObject);
var
  IsColRow: boolean;
begin
  case TRangeType(cbRangeType.ItemIndex) of
    rtCol: begin
      laColRowCell.Caption := QImportLoadStr(QIRE_RangeType_Col);
      rbStartColRow.Caption := QImportLoadStr(QIRE_Start_Row);
      rbFinishColRow.Caption := QImportLoadStr(QIRE_Finish_Row);
      rgDirection.Items[0] := QImportLoadStr(QIRE_Direction_Down);
      rgDirection.Items[1] := QImportLoadStr(QIRE_Direction_Up);
    end;
    rtRow: begin
      laColRowCell.Caption := QImportLoadStr(QIRE_RangeType_Row);
      rbStartColRow.Caption := QImportLoadStr(QIRE_Start_Col);
      rbFinishColRow.Caption := QImportLoadStr(QIRE_Finish_Col);
      rgDirection.Items[0] := QImportLoadStr(QIRE_Direction_Right);
      rgDirection.Items[1] := QImportLoadStr(QIRE_Direction_Left);
    end;
    rtCell: laColRowCell.Caption := QImportLoadStr(QIRE_RangeType_Cell);
  end;
  bCancel.Caption := QImportLoadStr(QIRE_Cancel);

  IsColRow := TRangeType(cbRangeType.ItemIndex) in [rtCol, rtRow];

  rbStartData.Enabled := IsColRow;
  rbStartColRow.Enabled := IsColRow;
  edStartColRow.Enabled := IsColRow and rbStartColRow.Checked;

  rbFinishData.Enabled := IsColRow;
  rbFinishColRow.Enabled := IsColRow;
  edFinishColRow.Enabled := IsColRow and rbFinishColRow.Checked;

  rgDirection.Enabled := IsColRow and
    (rbStartData.Checked or rbFinishData.Checked);

  rbDefaultSheet.Enabled := IsColRow;
  rbCustomSheet.Enabled := IsColRow;
  rbSheetNumber.Enabled := IsColRow and rbCustomSheet.Checked;
  edSheetNumber.Enabled := IsColRow and rbCustomSheet.Checked and
    rbSheetNumber.Checked;
  rbSheetName.Enabled := IsColRow and rbCustomSheet.Checked;
  edSheetName.Enabled := IsColRow and rbCustomSheet.Checked and
    rbSheetName.Checked;
end;

procedure TfmQImport3XLSRangeEdit.bOkClick(Sender: TObject);
begin
  LoadData;
  ModalResult := mrOk;
end;

procedure TfmQImport3XLSRangeEdit.FormCreate(Sender: TObject);
begin
  Caption := QImportLoadStr(QIRE_Caption);
  gbRangeType.Caption := QImportLoadStr(QIRE_RangeType);
  cbRangeType.Items[0] := QImportLoadStr(QIRE_RangeType_Col);
  cbRangeType.Items[1] := QImportLoadStr(QIRE_RangeType_Row);
  cbRangeType.Items[2] := QImportLoadStr(QIRE_RangeType_Cell);
  laColRowCell.Caption := QImportLoadStr(QIRE_RangeType_Col);
  gbStartCondition.Caption := QImportLoadStr(QIRE_Start);
  rbStartData.Caption := QImportLoadStr(QIRE_Start_DataStarted);
  rbStartColRow.Caption := QImportLoadStr(QIRE_Start_Row);
  gbFinishCondition.Caption := QImportLoadStr(QIRE_Finish);
  rbFinishData.Caption := QImportLoadStr(QIRE_Finish_DataFinished);
  rbFinishColRow.Caption := QImportLoadStr(QIRE_Finish_Row);
  rgDirection.Caption := QImportLoadStr(QIRE_Direction);
  rgDirection.Items[0] := QImportLoadStr(QIRE_Direction_Down);
  rgDirection.Items[1] := QImportLoadStr(QIRE_Direction_Up);
  gbSheetType.Caption := QImportLoadStr(QIRE_Sheet);
  rbDefaultSheet.Caption := QImportLoadStr(QIRE_Sheet_Default);
  rbCustomSheet.Caption := QImportLoadStr(QIRE_Sheet_Custom);
  rbSheetNumber.Caption := QImportLoadStr(QIRE_Sheet_Custom_Number);
  rbSheetName.Caption := QImportLoadStr(QIRE_Sheet_Custom_Name);
end;

end.
